#!/usr/bin/env python3
"""ChemScript bridge for the DSH ChemDraw plugin.

Runs on the Python 3.10 interpreter that matches ChemDraw's bundled
ChemScript extension module (D:\\CD\\ChemScript\\Lib).

Usage:  python chem.py <op> <args-json-file>
Ops:    inspect | export | convert
Prints a single line  @@JSON@@{...}  on stdout.
"""
import sys, os, json, io, contextlib, traceback

CHEMSCRIPT_LIB = os.environ.get("CHEMSCRIPT_LIB") or r"D:\CD\ChemScript\Lib"
sys.path.insert(0, CHEMSCRIPT_LIB)

_banner = io.StringIO()
with contextlib.redirect_stdout(_banner):
    import ChemScript  # noqa: E402  (prints its version banner)

MIMES = {
    "smiles": "chemical/x-daylight-smiles",
    "inchi": "chemical/x-inchi",
    "name": "chemical/x-name",
    "mol": "chemical/x-mdl-molfile",
    "sdf": "chemical/x-mdl-sdfile",
    "cdx": "chemical/x-cdx",
    "cdxml": "text/xml",
    "c3xml": "chemical/x-c3xml",
}

FILE_MIME = {
    ".cdx": "cdx",
    ".cdxml": "cdxml",
    ".mol": "mol",
    ".sdf": "sdf",
    ".smi": "smiles",
    ".smiles": "smiles",
}


class ChemError(Exception):
    pass


def emit(payload):
    sys.stdout.write("@@JSON@@" + json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def guess_from(inp):
    if inp.get("from"):
        return inp["from"]
    text = (inp.get("text") or "").strip()
    if text.startswith("InChI="):
        return "inchi"
    if "/" in text or "\\" in text:
        return "smiles"
    if " " in text and any(c.isalpha() for c in text):
        return "name"
    return "smiles"


def load_structure(inp):
    """Build a ChemScript StructureData from {text|path, from}."""
    path = inp.get("path")
    if path:
        if not os.path.isfile(path):
            raise ChemError("file not found: %s" % path)
        return ChemScript.StructureData.LoadFile(path)
    text = (inp.get("text") or "").strip()
    if not text:
        raise ChemError("no structure input given (text or path required)")
    frm = guess_from(inp)
    if frm not in MIMES:
        raise ChemError("unsupported input format: %s" % frm)
    try:
        return ChemScript.StructureData.LoadData(text, MIMES[frm])
    except Exception as exc:  # noqa: BLE001 - surface ChemScript message
        raise ChemError("could not parse %s input: %s" % (frm, exc))


def maybe_number(value, digits=4):
    try:
        return round(float(value), digits)
    except Exception:  # noqa: BLE001
        return None


def atom_count(struct):
    try:
        return len(struct.Atoms)
    except Exception:  # noqa: BLE001
        return None


def count_atoms_of(struct, attr):
    """ChemScript exposes some atom groups as list attributes, some as methods."""
    try:
        value = getattr(struct, attr)
    except Exception:  # noqa: BLE001
        return None
    try:
        if callable(value):
            value = value()
        return len(value)
    except Exception:  # noqa: BLE001
        return None


def describe(struct, inp=None):
    props = {}
    try:
        props["formula"] = struct.Formula()
    except Exception:  # noqa: BLE001
        props["formula"] = None
    for key, attr, digits in (
        ("molecular_weight", "Weight", 3),
        ("exact_mass", "ExactMass", 5),
        ("logp", "PartitionCoefficient", 3),
        ("refractivity", "Refractivity", 3),
    ):
        try:
            props[key] = maybe_number(getattr(struct, attr), digits)
        except Exception:  # noqa: BLE001
            props[key] = None
    for key, attr in (("smiles", "Smiles"), ("inchi", "Inchi")):
        try:
            props[key] = getattr(struct, attr)
        except Exception:  # noqa: BLE001
            props[key] = None
    try:
        props["charge"] = int(struct.TotalCharge)
    except Exception:  # noqa: BLE001
        props["charge"] = None
    try:
        props["has_coordinates"] = bool(struct.HasCoordinates)
    except Exception:  # noqa: BLE001
        props["has_coordinates"] = None
    try:
        props["atom_count"] = atom_count(struct)
    except Exception:  # noqa: BLE001
        props["atom_count"] = None
    try:
        topo = struct.Topology()
        props["tpsa"] = maybe_number(topo.PolarSurfaceArea, 2)
        props["rotatable_bonds"] = int(topo.NumOfRotatableBonds)
        props["ring_count"] = int(topo.CyclomaticNumber)
        props["num_hydrogens"] = int(topo.NumOfHydrogens)
    except Exception:  # noqa: BLE001
        pass
    props["h_bond_donors"] = count_atoms_of(struct, "HydrogenDonors")
    props["h_bond_acceptors"] = count_atoms_of(struct, "HydrogenAcceptors")
    try:
        props["iupac_name"] = struct.ChemicalName()
    except Exception:  # noqa: BLE001
        props["iupac_name"] = None
    return props


def maybe_neutralize(struct, args):
    """OCR often drops the acidic proton; only fix net-anionic structures."""
    if not args.get("neutralize"):
        return struct
    try:
        if struct.TotalCharge < 0:
            struct.Neutralize()
    except Exception:  # noqa: BLE001
        pass
    return struct


def op_inspect(args):
    struct = load_structure(args)
    struct = maybe_neutralize(struct, args)
    data = describe(struct)
    data["ok"] = True
    return data


def ensure_coordinates(struct, kekule=True):
    """Give the structure 2D coordinates and a clean, conventional bond rendering."""
    try:
        struct.NormalizeStructure()
    except Exception:  # noqa: BLE001
        pass
    try:
        if not struct.HasCoordinates:
            struct.CleanupStructure()
    except Exception:  # noqa: BLE001
        try:
            struct.CleanupStructure()
        except Exception:  # noqa: BLE001
            pass
    try:
        struct.CleanupStructure()
    except Exception:  # noqa: BLE001
        pass
    if kekule:
        try:
            struct.ConvertToFixedBondRep()
        except Exception:  # noqa: BLE001
            pass
    return struct


def write_cdxml(struct, out_path):
    data = struct.WriteData("text/xml")
    if isinstance(data, str):
        text = data
    else:
        text = bytes(data).decode("utf-8", "replace")
    if not text.lstrip().startswith("<"):
        raise ChemError("ChemScript did not return CDXML data")
    with open(out_path, "w", encoding="utf-8", newline="") as fh:
        fh.write(text)
    return out_path


def write_serialized(struct, fmt, out_path):
    """Write cdx / mol / sdf through ChemScript's own writers."""
    struct.WriteFile(out_path)
    if not os.path.isfile(out_path) or os.path.getsize(out_path) == 0:
        raise ChemError("ChemScript could not write %s output" % fmt)
    return out_path


def op_export(args):
    struct = load_structure(args)
    struct = maybe_neutralize(struct, args)
    struct = ensure_coordinates(struct)
    out_dir = args.get("out_dir") or os.getcwd()
    os.makedirs(out_dir, exist_ok=True)
    base = args.get("basename") or "structure"
    formats = args.get("formats") or ["cdxml"]
    written = {}
    for fmt in formats:
        fmt = fmt.lower().lstrip(".")
        path = os.path.join(out_dir, base + "." + fmt)
        if fmt == "cdxml":
            written[fmt] = write_cdxml(struct, path)
        elif fmt in ("cdx", "mol", "sdf"):
            written[fmt] = write_serialized(struct, fmt, path)
        elif fmt in ("smiles", "inchi"):
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(describe(struct).get(fmt) or "")
            written[fmt] = path
        else:
            raise ChemError("unsupported export format: %s" % fmt)
    return {"ok": True, "files": written, "properties": describe(struct)}


def op_convert(args):
    """Convert structure text/file into a normalized CDXML (+ optional serialized files)."""
    struct = load_structure(args)
    struct = ensure_coordinates(struct)
    out_path = args["out"]
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    ext = os.path.splitext(out_path)[1].lower()
    if ext == ".cdxml":
        write_cdxml(struct, out_path)
    elif ext in (".cdx", ".mol", ".sdf"):
        write_serialized(struct, ext[1:], out_path)
    else:
        raise ChemError("unsupported output extension: %s" % ext)
    props = describe(struct)
    props["ok"] = True
    props["file"] = out_path
    return props


OPS = {"inspect": op_inspect, "export": op_export, "convert": op_convert}


def main(argv):
    if len(argv) < 3:
        emit({"ok": False, "error": "usage: chem.py <op> <args-json-file>"})
        return 2
    op = argv[1]
    with open(argv[2], "r", encoding="utf-8-sig") as fh:
        args = json.load(fh)
    if op not in OPS:
        emit({"ok": False, "error": "unknown op: %s" % op})
        return 2
    try:
        result = OPS[op](args)
    except ChemError as exc:
        emit({"ok": False, "error": str(exc)})
        return 1
    except Exception as exc:  # noqa: BLE001
        emit({
            "ok": False,
            "error": "%s: %s" % (type(exc).__name__, exc),
            "traceback": traceback.format_exc(limit=4),
        })
        return 1
    emit(result)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
