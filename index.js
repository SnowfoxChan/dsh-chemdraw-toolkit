/**
 * DSH ChemDraw toolkit - host plugin.
 *
 * Bridges DeepSeek Harness tools to the locally installed ChemDraw suite:
 *   - ChemScript (Python 3.10 extension in D:\CD\ChemScript\Lib) does structure
 *     parsing, property calculation, name <-> structure and file conversion.
 *   - ChemDraw COM automation renders publication quality images, re-lays out
 *     (beautifies) structures and runs ChemNMR shift predictions.
 *   - CACTUS OSRA (optional, network) turns pictures of structures into MOL.
 */
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BACKEND = path.join(HERE, 'backend');

const DEFAULTS = {
  pythonPath: 'C:\\Users\\MSI\\AppData\\Local\\Programs\\Python\\Python310\\python.exe',
  chemscriptLib: 'D:\\CD\\ChemScript\\Lib',
  outputDir: path.join(os.homedir(), 'Documents', 'DSH-ChemDraw'),
  allowNetwork: true,
  timeoutMs: 240000,
};

const OSRA_BASE = 'https://cactus.nci.nih.gov';

/* ------------------------------------------------------------------ utils */

function log(...parts) {
  if (process.env.DSH_CHEMDRAW_DEBUG) console.error('[chemdraw]', ...parts);
}

function execFile(file, args, { timeoutMs = 120000, env } = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(file, args, {
      windowsHide: true,
      env: { ...process.env, PYTHONIOENCODING: 'utf-8', ...env },
    });
    let stdout = '';
    let stderr = '';
    let settled = false;
    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      child.kill();
      reject(new Error(`command timed out after ${Math.round(timeoutMs / 1000)}s: ${path.basename(file)}`));
    }, timeoutMs);
    child.stdout.on('data', (chunk) => { stdout += chunk.toString('utf8'); });
    child.stderr.on('data', (chunk) => { stderr += chunk.toString('utf8'); });
    child.on('error', (err) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      reject(err);
    });
    child.on('close', (code) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve({ code, stdout, stderr });
    });
  });
}

function parseEnvelope(stdout, stderr, code, what) {
  const line = stdout.split(/\r?\n/).find((l) => l.startsWith('@@JSON@@'));
  if (!line) {
    const tail = (stderr || stdout || '').trim().split(/\r?\n/).slice(-6).join('\n');
    throw new Error(`${what} failed (exit ${code}).\n${tail || 'no output'}`);
  }
  let payload;
  try {
    payload = JSON.parse(line.slice('@@JSON@@'.length));
  } catch (err) {
    throw new Error(`${what} returned malformed output: ${err.message}`);
  }
  if (payload && payload.ok === false) {
    const stage = payload.stage ? ` [stage: ${payload.stage}]` : '';
    throw new Error(`${what} failed${stage}: ${payload.error || 'unknown error'}`);
  }
  return payload;
}

function tmpDir() {
  return fsp.mkdtemp(path.join(os.tmpdir(), 'dsh-chemdraw-'));
}

function writeJson(file, value) {
  fs.writeFileSync(file, JSON.stringify(value), { encoding: 'utf8' });
  return file;
}

async function runPython(cfg, op, args) {
  const dir = await tmpDir();
  const argFile = writeJson(path.join(dir, 'args.json'), args);
  const script = path.join(BACKEND, 'chem.py');
  const { code, stdout, stderr } = await execFile(cfg.pythonPath, [script, op, argFile], {
    timeoutMs: 120000,
    env: { CHEMSCRIPT_LIB: cfg.chemscriptLib },
  });
  return parseEnvelope(stdout, stderr, code, `ChemScript ${op}`);
}

async function runChemDraw(cfg, args) {
  const dir = await tmpDir();
  const argFile = writeJson(path.join(dir, 'args.json'), args);
  const script = path.join(BACKEND, 'chemdraw.ps1');
  const { code, stdout, stderr } = await execFile(
    'powershell.exe',
    ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script, '-ArgsJson', argFile],
    { timeoutMs: cfg.timeoutMs },
  );
  return parseEnvelope(stdout, stderr, code, 'ChemDraw automation');
}

async function runChem3D(cfg, args) {
  const dir = await tmpDir();
  const argFile = writeJson(path.join(dir, 'args.json'), args);
  const script = path.join(BACKEND, 'chem3d.ps1');
  const { code, stdout, stderr } = await execFile(
    'powershell.exe',
    ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script, '-ArgsJson', argFile],
    { timeoutMs: cfg.timeoutMs },
  );
  return parseEnvelope(stdout, stderr, code, 'Chem3D automation');
}

function resolveOutDir(cfg, requested) {
  const dir = path.resolve(requested || cfg.outputDir);
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function safeBase(name, fallback = 'structure') {
  const base = String(name || '').replace(/[\\/:*?"<>|]+/g, '_').trim();
  return base || fallback;
}

function isExistingFile(value) {
  try {
    return Boolean(value) && fs.statSync(String(value)).isFile();
  } catch {
    return false;
  }
}

/** Turn a tool `input` (text or a path) into the backend's structure argument. */
function structureArg(input, from) {
  const raw = String(input ?? '').trim();
  if (!raw) throw new Error('structure input is required (SMILES, InChI, name, or a .cdx/.cdxml/.mol file path)');
  if (isExistingFile(raw)) return { path: path.resolve(raw) };
  const fmt = from && from !== 'auto' ? from : undefined;
  return fmt ? { text: raw, from: fmt } : { text: raw };
}

function structureHint(input, from) {
  const raw = String(input ?? '').trim();
  if (isExistingFile(raw)) return path.basename(raw);
  const fmt = from && from !== 'auto' ? from : 'auto';
  return `${fmt}: ${raw.length > 60 ? raw.slice(0, 57) + '...' : raw}`;
}

function pct(style) {
  if (!style) return undefined;
  const out = {};
  const map = {
    bond_length: 'bond_length',
    line_width: 'line_width',
    label_size: 'label_size',
    label_font: 'label_font',
    caption_size: 'caption_size',
    hide_implicit_hydrogens: 'hide_implicit_hydrogens',
  };
  for (const [k, v] of Object.entries(style)) {
    if (map[k] !== undefined && v !== undefined && v !== null) out[k] = v;
  }
  return Object.keys(out).length ? out : undefined;
}

/* ------------------------------------------------------------- formatting */

function fmtNum(value, digits = 3) {
  return typeof value === 'number' && Number.isFinite(value) ? Number(value.toFixed(digits)).toString() : '-';
}

function formatProperties(p) {
  if (!p) return 'no properties';
  const lines = [];
  if (p.formula) lines.push(`Formula: ${p.formula}`);
  if (p.molecular_weight) lines.push(`Molecular weight: ${fmtNum(p.molecular_weight, 3)}`);
  if (p.exact_mass) lines.push(`Exact mass: ${fmtNum(p.exact_mass, 4)}`);
  if (p.iupac_name) lines.push(`IUPAC name: ${p.iupac_name}`);
  if (p.smiles) lines.push(`SMILES: ${p.smiles}`);
  if (p.inchi) lines.push(`InChI: ${p.inchi}`);
  if (p.logp !== undefined && p.logp !== null) lines.push(`cLogP: ${fmtNum(p.logp, 3)}`);
  if (p.tpsa !== undefined && p.tpsa !== null) lines.push(`TPSA: ${fmtNum(p.tpsa, 2)} A^2`);
  const extras = [];
  if (p.charge !== undefined && p.charge !== null) extras.push(`charge ${p.charge}`);
  if (p.h_bond_donors !== undefined && p.h_bond_donors !== null) extras.push(`HBD ${p.h_bond_donors}`);
  if (p.h_bond_acceptors !== undefined && p.h_bond_acceptors !== null) extras.push(`HBA ${p.h_bond_acceptors}`);
  if (p.rotatable_bonds !== undefined && p.rotatable_bonds !== null) extras.push(`rotatable bonds ${p.rotatable_bonds}`);
  if (p.ring_count !== undefined && p.ring_count !== null) extras.push(`rings ${p.ring_count}`);
  if (extras.length) lines.push(extras.join(' | '));
  return lines.join('\n');
}

function formatNmr(result) {
  const lines = [`ChemNMR ${result.nucleus} shift prediction`];
  if (result.solvent || result.frequency) {
    lines.push([result.solvent && `solvent ${result.solvent}`, result.frequency && `${result.frequency} MHz`]
      .filter(Boolean).join(' | '));
  }
  lines.push('');
  if (result.shifts?.length) {
    lines.push('Atom       Shift (ppm)   H    Note');
    for (const s of result.shifts) {
      lines.push(
        `${String(s.atom ?? '?').padEnd(10)} ${String(fmtNum(s.shift, 2)).padEnd(13)} ` +
        `${String(s.protons ?? (result.nucleus === '13C' ? '' : 1)).padEnd(4)} ${s.comment || ''}`,
      );
    }
  } else {
    lines.push('No per-atom shifts were parsed from the ChemNMR report.');
  }
  if (result.couplings?.length) {
    lines.push('', '1H coupling constants:');
    for (const c of result.couplings) {
      const partners = (c.partners || [])
        .map((p) => `${p.atom}${p.constant ? ` (${fmtNum(p.constant, 1)} Hz, ${p.vector || ''})` : ''}`)
        .join(', ');
      lines.push(`  H at ${fmtNum(c.shift, 2)} ppm (atom ${c.atom}) -> ${partners || 'no partners'}`);
    }
  }
  if (result.image) lines.push('', `Spectrum + assignment sheet image: ${result.image}`);
  if (result.protocol_file) lines.push(`Full ChemNMR protocol: ${result.protocol_file}`);
  return lines.join('\n');
}

function formatImageFile(file, size) {
  if (!file) return 'no image';
  const dim = size && size.width ? ` (${size.width}x${size.height} px)` : '';
  return `${file}${dim}`;
}

/* ------------------------------------------------------------------- OSRA */

function normalizeOsraMol(raw) {
  // OSRA's CTAB comes back with "|" as the line separator and no MOL header.
  const parts = raw.replace(/\r/g, '').split('|');
  const cleaned = parts.slice(1).filter((seg) => seg.trim() !== '');
  if (!cleaned.length) throw new Error('OSRA returned an empty MOL file');
  const program = cleaned.shift();
  return ['OSRA', program, '', ...cleaned].join('\n') + '\n';
}

async function fetchText(url, timeoutMs = 30000) {
  const res = await fetch(url, { signal: AbortSignal.timeout(timeoutMs) });
  if (!res.ok) return null;
  return res.text();
}

async function osraRecognize(cfg, imagePath) {
  if (!cfg.allowNetwork) throw new Error('image -> structure needs network access (CACTUS OSRA); enable "allowNetwork" in the plugin config');
  const buffer = await fsp.readFile(imagePath);
  const fid = String(Math.floor(Math.random() * 1e9));
  const form = new FormData();
  form.append('fid', fid);
  form.append('file', new Blob([buffer], { type: 'image/png' }), path.basename(imagePath));
  log('osra upload', imagePath, buffer.length, 'bytes, fid', fid);
  const res = await fetch(`${OSRA_BASE}/cgi-bin/osra/index.cgi`, {
    method: 'POST',
    body: form,
    signal: AbortSignal.timeout(120000),
  });
  await res.text();
  const doneUrl = `${OSRA_BASE}/tmp/osra-${fid}/done`;
  let count = 0;
  for (let attempt = 0; attempt < 90; attempt += 1) {
    const done = await fetchText(doneUrl, 15000);
    if (done !== null) {
      count = Number.parseInt(done.trim(), 10) || 0;
      break;
    }
    await new Promise((resolve) => setTimeout(resolve, 1000));
  }
  const mols = [];
  for (let i = 0; i < Math.min(count, 8); i += 1) {
    const raw = await fetchText(`${OSRA_BASE}/tmp/osra-${fid}/ctab.${i}`, 30000);
    if (!raw) continue;
    try {
      mols.push(normalizeOsraMol(raw));
    } catch (err) {
      log('osra mol parse failed', i, err.message);
    }
  }
  return { count, mols };
}

/* ------------------------------------------------------------- NMR parsing */

function parseNmrReport(captions, nucleus) {
  const protocol = captions.find((c) => /protocol of the/i.test(c)) || captions.join('\n');
  const lines = protocol.split(/\r\n|\r|\n/);
  const solvent = /solvent\s*=\s*([A-Za-z0-9-]+)/i.exec(protocol)?.[1];
  const frequency = /(\d{2,3})\s*MHz/i.exec(protocol)?.[1];
  const shifts = [];
  const shiftRe = /^\s*([A-Za-z][A-Za-z0-9]*)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s*(.*)$/;
  const stopRe = /coupling constant prediction/i;
  const couplingSection = [];
  let inCoupling = false;
  for (const line of lines) {
    if (stopRe.test(line)) { inCoupling = true; continue; }
    if (inCoupling) { couplingSection.push(line); continue; }
    if (/^\s*(node|shift)\b/i.test(line) || /base\s*\+\s*inc/i.test(line)) continue;
    const m = shiftRe.exec(line);
    if (!m) continue;
    if (/^shifts?$/i.test(m[1])) continue;
    const label = m[1];
    const protons = /^(?:C)?H(\d*)$/.test(label) ? (Number(label.replace(/^(?:C)?H/, '')) || 1) : undefined;
    shifts.push({ atom: label, shift: Number(m[2]), base: Number(m[3]), comment: (m[4] || '').trim(), protons });
  }
  const couplings = [];
  let current = null;
  for (const line of couplingSection) {
    const entry = /^\s*(\d+(?:\.\d+)?)\s+(\d+)\s*$/.exec(line);
    if (entry) {
      current = { shift: Number(entry[1]), atom: entry[2], partners: [] };
      couplings.push(current);
      continue;
    }
    const partner = /^\s+(\d+)\s+(\d+(?:\.\d+)?)\s+(\S+)/.exec(line);
    if (partner && current) {
      current.partners.push({ atom: partner[1], constant: Number(partner[2]), vector: partner[3] });
    }
  }
  return { solvent, frequency: frequency ? Number(frequency) : undefined, shifts, couplings, protocol };
}

/* ------------------------------------------------------------------- tools */

// C3DDisplayType values understood by Chem3D (k3DWireFrame=0 ... k3DRibbons=5).
const DISPLAY_TYPES = {
  wireframe: 0,
  sticks: 1,
  ball_and_stick: 2,
  cylindrical: 3,
  space_filling: 4,
};

function toolDefinitions(cfg) {
  const structureInput = {
    type: 'string',
    description: 'Structure to work on: SMILES, InChI, a chemical name, or the path of a .cdx / .cdxml / .mol / .sdf / .smi file.',
  };
  const fromParam = {
    type: 'string',
    enum: ['auto', 'smiles', 'inchi', 'name', 'mol', 'sdf', 'cdx', 'cdxml'],
    description: 'How to read `input`. Defaults to auto-detection.',
  };
  const outDirParam = {
    type: 'string',
    description: `Directory for produced files (created if missing). Defaults to ${cfg.outputDir}.`,
  };
  const basenameParam = {
    type: 'string',
    description: 'Base file name (without extension) for produced files.',
  };
  const styleParam = {
    type: 'object',
    description: 'Optional ChemDraw styling: bond_length (pt), line_width (pt), label_size (pt), label_font, caption_size (pt).',
    additionalProperties: true,
  };

  const objSchema = { type: 'object', additionalProperties: true };

  return [
    {
      name: 'chem_analyze',
      description: 'Identify a chemical structure and report its properties: formula, molecular weight, exact mass, IUPAC name, SMILES, InChI, cLogP, TPSA, H-bond donors/acceptors, rotatable bonds and ring count. Accepts SMILES, InChI, a common name, or a ChemDraw/MOL file path.',
      parameters: {
        type: 'object',
        properties: { input: structureInput, from: fromParam },
        required: ['input'],
      },
      output: {
        schema: objSchema,
        render: (_args, value) => [{ type: 'text', text: formatProperties(value.properties) }],
      },
      async execute(args) {
        const payload = await runPython(cfg, 'inspect', structureArg(args.input, args.from));
        return { ok: true, properties: payload };
      },
    },

    {
      name: 'chem_draw',
      description: 'Draw a structure with ChemDraw and save it as an image or a ChemDraw file. Produces a crisp, tightly cropped PNG (or SVG) plus optional .cdx/.cdxml/.mol copies. Use this whenever a structure illustration is needed for an answer, report or slide.',
      parameters: {
        type: 'object',
        properties: {
          input: structureInput,
          from: fromParam,
          format: { type: 'string', enum: ['png', 'svg', 'cdx', 'cdxml', 'mol'], description: 'Primary output format. Defaults to png.' },
          out_dir: outDirParam,
          basename: basenameParam,
          style: styleParam,
        },
        required: ['input'],
      },
      output: {
        schema: objSchema,
        render: (args, value) => {
          const lines = [`Drew ${structureHint(args.input, args.from)} with ChemDraw.`];
          if (value.properties) lines.push('', formatProperties(value.properties));
          lines.push('', 'Files:');
          for (const [fmt, file] of Object.entries(value.files || {})) lines.push(`  ${fmt}: ${file}`);
          return [{ type: 'text', text: lines.filter((l) => l !== undefined).join('\n') }];
        },
      },
      async execute(args) {
        const outDir = resolveOutDir(cfg, args.out_dir);
        const base = safeBase(args.basename, 'structure');
        const format = (args.format || 'png').toLowerCase();
        const structure = structureArg(args.input, args.from);
        const cdxmlPath = path.join(outDir, `${base}.cdxml`);
        const props = await runPython(cfg, 'export', { ...structure, out_dir: outDir, basename: base, formats: ['cdxml'] });
        const files = { cdxml: cdxmlPath };
        if (['png', 'svg'].includes(format)) {
          const imagePath = path.join(outDir, `${base}.${format}`);
          const rendered = await runChemDraw(cfg, { op: 'render', input: cdxmlPath, output: imagePath, clean: true, style: pct(args.style) });
          files[format] = rendered.file;
          return { ok: true, properties: props.properties, files, image: rendered.image };
        }
        const extraFormats = format === 'cdxml' ? [] : [format];
        if (extraFormats.length) {
          const exported = await runPython(cfg, 'export', { ...structure, out_dir: outDir, basename: base, formats: extraFormats });
          Object.assign(files, exported.files);
        }
        return { ok: true, properties: props.properties, files };
      },
    },

    {
      name: 'chem_beautify',
      description: 'Clean up and beautify an existing structure: ChemDraw recomputes a tidy 2D layout, normalises bond lengths, line widths and atom label font, and re-exports it. Use on messy .cdx/.cdxml/.mol files (or any structure input) to get presentation-ready artwork.',
      parameters: {
        type: 'object',
        properties: {
          input: structureInput,
          from: fromParam,
          formats: {
            type: 'array',
            items: { type: 'string', enum: ['png', 'svg', 'cdx', 'cdxml'] },
            description: 'Output formats to write. Defaults to png + cdxml.',
          },
          out_dir: outDirParam,
          basename: basenameParam,
          style: styleParam,
        },
        required: ['input'],
      },
      output: {
        schema: objSchema,
        render: (args, value) => {
          const lines = [`Beautified ${structureHint(args.input, args.from)}.`];
          lines.push('', 'Files:');
          for (const [fmt, file] of Object.entries(value.files || {})) lines.push(`  ${fmt}: ${file}`);
          return [{ type: 'text', text: lines.join('\n') }];
        },
      },
      async execute(args) {
        const outDir = resolveOutDir(cfg, args.out_dir);
        const base = safeBase(args.basename, 'beautified');
        const formats = (args.formats?.length ? args.formats : ['png', 'cdxml']).map((f) => f.toLowerCase());
        const structure = structureArg(args.input, args.from);
        const cdxmlPath = path.join(outDir, `${base}_source.cdxml`);
        const props = await runPython(cfg, 'export', { ...structure, out_dir: outDir, basename: `${base}_source`, formats: ['cdxml'] });
        const beautified = await runChemDraw(cfg, {
          op: 'beautify',
          input: cdxmlPath,
          output_base: path.join(outDir, base),
          formats,
          style: { line_width: 0.75, label_size: 11, caption_size: 10, ...(pct(args.style) || {}) },
        });
        return { ok: true, properties: props.properties, files: beautified.files, image: beautified.image };
      },
    },

    {
      name: 'chem_nmr',
      description: 'Predict a 1H or 13C NMR spectrum with ChemDraw ChemNMR. Returns per-atom predicted shifts (ppm), 1H coupling constants, the full ChemNMR protocol and an image of the predicted spectrum with the shift-assignment table.',
      parameters: {
        type: 'object',
        properties: {
          input: structureInput,
          from: fromParam,
          nucleus: { type: 'string', enum: ['1H', '13C'], description: 'Nucleus to predict. Defaults to 1H.' },
          out_dir: outDirParam,
          basename: basenameParam,
        },
        required: ['input'],
      },
      output: {
        schema: objSchema,
        render: (_args, value) => [{ type: 'text', text: formatNmr(value) }],
      },
      timeoutMs: 300000,
      async execute(args) {
        const outDir = resolveOutDir(cfg, args.out_dir);
        const base = safeBase(args.basename, nucleusSlug(args.nucleus));
        const nucleus = args.nucleus === '13C' ? '13C' : '1H';
        const structure = structureArg(args.input, args.from);
        const cdxmlPath = path.join(outDir, `${base}_source.cdxml`);
        await runPython(cfg, 'export', { ...structure, out_dir: outDir, basename: `${base}_source`, formats: ['cdxml'] });
        const predicted = await runChemDraw(cfg, {
          op: 'nmr',
          input: cdxmlPath,
          output_base: path.join(outDir, base),
          nucleus,
          visible: true,
        });
        const report = parseNmrReport(predicted.captions || [], nucleus);
        const protocolFile = path.join(outDir, `${base}_protocol.txt`);
        fs.writeFileSync(protocolFile, report.protocol, 'utf8');
        return {
          ok: true,
          nucleus,
          solvent: report.solvent,
          frequency: report.frequency,
          shifts: report.shifts,
          couplings: report.couplings,
          protocol_file: protocolFile,
          image: predicted.png,
          image_size: predicted.image,
        };
      },
    },

    {
      name: 'chem_convert',
      description: 'Convert a structure between formats using ChemScript/ChemDraw: cdx, cdxml, mol, sdf, smiles, inchi, svg and png. Useful for turning drawn .cdx files into SMILES/MOL for other software, or for exporting clean images.',
      parameters: {
        type: 'object',
        properties: {
          input: structureInput,
          from: fromParam,
          formats: {
            type: 'array',
            items: { type: 'string', enum: ['cdx', 'cdxml', 'mol', 'sdf', 'smiles', 'inchi', 'png', 'svg'] },
            description: 'Formats to write. Defaults to cdxml + mol.',
          },
          out_dir: outDirParam,
          basename: basenameParam,
        },
        required: ['input'],
      },
      output: {
        schema: objSchema,
        render: (args, value) => {
          const lines = [`Converted ${structureHint(args.input, args.from)}.`];
          lines.push('', 'Files:');
          for (const [fmt, file] of Object.entries(value.files || {})) lines.push(`  ${fmt}: ${file}`);
          if (value.properties?.smiles) lines.push('', `SMILES: ${value.properties.smiles}`);
          return [{ type: 'text', text: lines.join('\n') }];
        },
      },
      async execute(args) {
        const outDir = resolveOutDir(cfg, args.out_dir);
        const base = safeBase(args.basename, 'converted');
        const formats = (args.formats?.length ? args.formats : ['cdxml', 'mol']).map((f) => f.toLowerCase());
        const structure = structureArg(args.input, args.from);
        const structureFormats = formats.filter((f) => !['png', 'svg'].includes(f));
        const imageFormats = formats.filter((f) => ['png', 'svg'].includes(f));
        const files = {};
        let properties = null;
        if (structureFormats.length) {
          const exported = await runPython(cfg, 'export', { ...structure, out_dir: outDir, basename: base, formats: structureFormats });
          Object.assign(files, exported.files);
          properties = exported.properties;
        }
        if (imageFormats.length) {
          const cdxmlPath = files.cdxml || path.join(outDir, `${base}_render.cdxml`);
          if (!files.cdxml) {
            const exported = await runPython(cfg, 'export', { ...structure, out_dir: outDir, basename: `${base}_render`, formats: ['cdxml'] });
            files.cdxml = exported.files.cdxml;
          }
          for (const fmt of imageFormats) {
            const rendered = await runChemDraw(cfg, { op: 'render', input: cdxmlPath, output: path.join(outDir, `${base}.${fmt}`), clean: true });
            files[fmt] = rendered.file;
          }
        }
        return { ok: true, files, properties };
      },
    },

    {
      name: 'chem_image_to_structure',
      description: 'Recognise the chemical structure drawn in an image (screenshot, scanned page, photo) and convert it into machine-readable structure data. Uses CACTUS OSRA optical structure recognition, then validates and redraws the result with ChemDraw. Returns SMILES, formula, IUPAC name and a clean redrawn image.',
      parameters: {
        type: 'object',
        properties: {
          image: { type: 'string', description: 'Path to an image file (PNG/JPEG/BMP/GIF) containing a drawn chemical structure.' },
          neutralize: { type: 'boolean', description: 'Add the missing acidic proton when OCR returns a net-anionic structure (typical for carboxylic acids). Defaults to true; set false for genuine salts/anions.' },
          out_dir: outDirParam,
          basename: basenameParam,
        },
        required: ['image'],
      },
      output: {
        schema: objSchema,
        render: (args, value) => {
          const lines = [`Recognised structure in ${path.basename(String(args.image))}.`];
          if (value.candidates > 1) lines.push(`OSRA returned ${value.candidates} candidate structures; showing the first one.`);
          if (value.properties) lines.push('', formatProperties(value.properties));
          if (value.image) lines.push('', `Redrawn structure: ${formatImageFile(value.image, value.image_size)}`);
          lines.push('', 'Verify the result against the source image before relying on it.');
          return [{ type: 'text', text: lines.join('\n') }];
        },
      },
      timeoutMs: 300000,
      async execute(args) {
        if (!isExistingFile(args.image)) throw new Error(`image not found: ${args.image}`);
        const outDir = resolveOutDir(cfg, args.out_dir);
        const base = safeBase(args.basename, 'recognized');
        const prepared = path.join(outDir, `${base}_input.png`);
        await runChemDraw(cfg, { op: 'image', input: path.resolve(args.image), output: prepared, min_width: 900 });
        const { count, mols } = await osraRecognize(cfg, prepared);
        if (!mols.length) {
          throw new Error(
            `OSRA could not recognise a structure (candidates: ${count}). ` +
            'Tips: crop to a single structure, use a higher resolution image, and prefer plain solid bonds (alternating double bonds) over dashed aromatic rings.',
          );
        }
        const firstMolPath = path.join(outDir, `${base}.mol`);
        fs.writeFileSync(firstMolPath, mols[0], 'utf8');
        const neutralize = args.neutralize !== false;
        const props = await runPython(cfg, 'inspect', { path: firstMolPath, neutralize });
        let redrawn = null;
        let redrawnSize = null;
        if (props.smiles) {
          const cdxmlPath = path.join(outDir, `${base}_clean.cdxml`);
          await runPython(cfg, 'export', { text: props.smiles, from: 'smiles', neutralize, out_dir: outDir, basename: `${base}_clean`, formats: ['cdxml'] });
          const rendered = await runChemDraw(cfg, { op: 'render', input: cdxmlPath, output: path.join(outDir, `${base}_clean.png`), clean: true });
          redrawn = rendered.file;
          redrawnSize = rendered.image;
        }
        return {
          ok: true,
          candidates: mols.length,
          properties: props,
          mol_file: firstMolPath,
          image: redrawn,
          image_size: redrawnSize,
        };
      },
    },

    {
      name: 'chem_3d',
      description: 'Turn a 2D structure into a 3D conformer with Chem3D: hydrogens are added, the geometry is minimised with the MM2 force field, and a ball-and-stick (or other style) rendering is saved. Returns the 3D .mol (opens in Chem3D, PyMOL, Avogadro, Jmol) plus an optional Chem3D .c3xml and PNG.',
      parameters: {
        type: 'object',
        properties: {
          input: structureInput,
          from: fromParam,
          display: {
            type: 'string',
            enum: ['ball_and_stick', 'sticks', 'space_filling', 'wireframe', 'cylindrical'],
            description: '3D rendering style. Defaults to ball_and_stick.',
          },
          iterations: { type: 'number', description: 'Maximum MM2 minimisation iterations. Defaults to 2000.' },
          open: { type: 'boolean', description: 'Keep the Chem3D window open so the user can rotate or edit the model by hand. Defaults to false.' },
          out_dir: outDirParam,
          basename: basenameParam,
        },
        required: ['input'],
      },
      output: {
        schema: objSchema,
        render: (args, value) => {
          const lines = [`3D model of ${structureHint(args.input, args.from)} (Chem3D, MM2 minimisation).`];
          if (value.properties) lines.push('', formatProperties(value.properties));
          const metrics = [];
          if (value.atoms) metrics.push(`${value.atoms} atoms`);
          if (value.bonds) metrics.push(`${value.bonds} bonds`);
          if (value.steric_energy !== undefined && value.steric_energy !== null) metrics.push(`MM2 steric energy ${fmtNum(value.steric_energy, 3)} kcal/mol`);
          if (value.rms_gradient !== undefined && value.rms_gradient !== null) metrics.push(`RMS gradient ${fmtNum(value.rms_gradient, 3)}`);
          if (metrics.length) lines.push('', metrics.join(' | '));
          lines.push('', 'Files:');
          if (value.mol) lines.push(`  mol: ${value.mol}`);
          if (value.png) lines.push(`  png: ${value.png}${value.png_width ? ` (${value.png_width}x${value.png_height})` : ''}`);
          if (value.c3xml) lines.push(`  c3xml: ${value.c3xml}`);
          return [{ type: 'text', text: lines.join('\n') }];
        },
      },
      timeoutMs: 300000,
      async execute(args) {
        const outDir = resolveOutDir(cfg, args.out_dir);
        const base = safeBase(args.basename, 'structure');
        const structure = structureArg(args.input, args.from);
        const props = await runPython(cfg, 'inspect', structure);
        const exported = await runPython(cfg, 'export', { ...structure, out_dir: outDir, basename: `${base}_2d`, formats: ['mol'] });
        const mol2d = exported.files && exported.files.mol;
        if (!mol2d) throw new Error('could not export a 2D MOL file for Chem3D');
        const displayType = DISPLAY_TYPES[String(args.display || 'ball_and_stick').toLowerCase()];
        const iterations = Number.isFinite(args.iterations) ? Math.max(50, Math.min(20000, Math.round(args.iterations))) : 2000;
        const keepOpen = args.open === true;
        const result = await runChem3D(cfg, {
          input: mol2d,
          out_dir: outDir,
          basename: base,
          iterations,
          open: keepOpen,
          visible: keepOpen,
          style: displayType === undefined ? undefined : { display_type: displayType },
          max_edge: 1100,
        });
        if (!result.ok) throw new Error(result.error || 'Chem3D could not build the 3D model');
        return {
          ok: true,
          properties: props,
          mol: result.mol,
          png: result.png,
          png_width: result.png_width,
          png_height: result.png_height,
          c3xml: result.c3xml,
          atoms: result.atoms,
          bonds: result.bonds,
          steric_energy: result.steric_energy,
          rms_gradient: result.rms_gradient,
          model: result.model,
        };
      },
    },

    {
      name: 'chem_open',
      description: 'Open a structure file in the interactive ChemDraw window so the user can inspect or edit it by hand. Creates a .cdx file first when the input is text (SMILES/name).',
      parameters: {
        type: 'object',
        properties: { input: structureInput, from: fromParam, out_dir: outDirParam, basename: basenameParam },
        required: ['input'],
      },
      output: {
        schema: objSchema,
        render: (_args, value) => [{ type: 'text', text: `Opened "${value.name}" in ChemDraw.` }],
      },
      async execute(args) {
        const outDir = resolveOutDir(cfg, args.out_dir);
        const base = safeBase(args.basename, 'structure');
        const structure = structureArg(args.input, args.from);
        let openedPath;
        if (structure.path) {
          openedPath = structure.path;
        } else {
          openedPath = path.join(outDir, `${base}.cdx`);
          await runPython(cfg, 'export', { ...structure, out_dir: outDir, basename: base, formats: ['cdx'] });
        }
        const opened = await runChemDraw(cfg, { op: 'open', input: openedPath, keep: true });
        return { ok: true, name: opened.name, file: openedPath };
      },
    },
  ];
}

function nucleusSlug(nucleus) {
  return nucleus === '13C' ? 'nmr_13c' : 'nmr_1h';
}

/* -------------------------------------------------------------- registration */

export const inject = ['tools'];

/** Build the tool definitions for a given configuration (also used by the test harness). */
export function createTools(config) {
  return sanitizeTools(toolDefinitions({ ...DEFAULTS, ...(config || {}) }));
}

/**
 * The harness rejects tool results that cannot survive a JSON round-trip (for
 * example properties whose value is undefined), so normalise every payload.
 */
function lossless(value) {
  if (value === undefined) return value;
  return JSON.parse(JSON.stringify(value));
}

function sanitizeTools(tools) {
  return tools.map((definition) => ({
    ...definition,
    execute: async (args, exec) => lossless(await definition.execute(args, exec)),
  }));
}

export function apply(ctx, config) {
  const tools = createTools(config);
  const register = () => {
    const disposers = tools.map((definition) => ctx.tools.register(definition));
    return () => {
      for (const dispose of disposers) {
        try { dispose(); } catch (err) { log('dispose failed', err.message); }
      }
    };
  };
  if (typeof ctx.effect === 'function') {
    ctx.effect(register);
  } else {
    register();
  }
}
