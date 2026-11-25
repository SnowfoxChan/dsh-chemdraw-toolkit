param(
  [Parameter(Mandatory = $true)][string]$ArgsJson
)

$ErrorActionPreference = 'Stop'

$opt = Get-Content -Raw -Encoding UTF8 $ArgsJson | ConvertFrom-Json

function Emit([string]$json) {
  Write-Output ("@@JSON@@" + $json)
}

$code = @'
using System;
using System.Text;
using System.Drawing;
using System.Runtime.InteropServices;

public static class Chem3DRun
{
    [DllImport("user32.dll")] static extern bool MoveWindow(IntPtr h, int x, int y, int w, int ht, bool repaint);

    static string Esc(string s)
    {
        if (s == null) return "";
        return s.Replace("\\", "/").Replace("\"", "\\\"");
    }

    static string Num(double v)
    {
        if (double.IsNaN(v) || double.IsInfinity(v)) return "null";
        return Math.Round(v, 4).ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    // Content touching the image border means the view is cropped.
    static bool TouchesBorder(string path)
    {
        using (Bitmap bmp = new Bitmap(path))
        {
            int w = bmp.Width, h = bmp.Height;
            Color bg = bmp.GetPixel(0, 0);
            int band = 5;
            for (int x = 0; x < w; x++)
            {
                for (int d = 0; d < band; d++)
                {
                    if (!Same(bmp.GetPixel(x, d), bg)) return true;
                    if (!Same(bmp.GetPixel(x, h - 1 - d), bg)) return true;
                }
            }
            for (int y = 0; y < h; y++)
            {
                for (int d = 0; d < band; d++)
                {
                    if (!Same(bmp.GetPixel(d, y), bg)) return true;
                    if (!Same(bmp.GetPixel(w - 1 - d, y), bg)) return true;
                }
            }
        }
        return false;
    }

    static bool Same(Color a, Color b)
    {
        return Math.Abs(a.R - b.R) < 24 && Math.Abs(a.G - b.G) < 24 && Math.Abs(a.B - b.B) < 24;
    }

    public static string Execute(
        string inPath, string outDir, string baseName,
        int iterations, double rmsg, bool visible, bool keepOpen,
        int atomSize, int bondWidth, int displayType, int bgColor, int maxEdge)
    {
        var sb = new StringBuilder();
        dynamic app = null;
        dynamic doc = null;
        bool finished = false;
        try
        {
            app = Activator.CreateInstance(Type.GetTypeFromProgID("Chem3D_x64.Application"));
            app.Visible = visible;
            try
            {
                int hwnd = (int)app.MainWindow;
                if (hwnd != 0) MoveWindow(new IntPtr(hwnd), 0, 0, 2000, 1600, true);
            }
            catch { }
            System.Threading.Thread.Sleep(1200);

            doc = app.Documents.Open(inPath, null);
            doc.Activate();

            dynamic job = Activator.CreateInstance(Type.GetTypeFromProgID("Chem3D9.MM2Job"));
            job.Type = 1;
            job.MaximumNumberOfIterations = iterations;
            job.MinimumRMSGradient = rmsg;
            job.DisplayEveryIteration = false;
            job.RecordEveryIteration = false;
            job.MoveOnlySelectedAtoms = false;
            doc.MM2Compute(job);

            try { if (atomSize > 0) doc.AtomSize = atomSize; } catch { }
            try { if (bondWidth > 0) doc.BondWidth = bondWidth; } catch { }
            try { if (displayType >= 0) doc.DisplayType = displayType; } catch { }
            try { if (bgColor >= 0) doc.BackgroundColor = bgColor; } catch { }

            double baseScale = 102.0;
            try { baseScale = Convert.ToDouble(doc.ScaleFactor); } catch { }
            if (baseScale < 10.0) baseScale = 102.0;

            string probe = outDir + "\\" + baseName + "_probe.png";
            string finalPng = outDir + "\\" + baseName + "_3d.png";
            string molPath = outDir + "\\" + baseName + "_3d.mol";
            string c3xmlPath = outDir + "\\" + baseName + "_3d.c3xml";

            Func<double, int[]> shot = (scale) =>
            {
                doc.ScaleFactor = scale;
                System.Threading.Thread.Sleep(200);
                doc.SaveAs(probe);
                int w = 0, h = 0;
                using (Image img = Image.FromFile(probe)) { w = img.Width; h = img.Height; }
                return new int[] { w, h };
            };

            // A zoom that is too high renders the model cropped; widen the window
            // once more when the model is large, then scan upward for the last
            // zoom whose render does not touch the image border.
            int[] d0 = shot(baseScale);
            int bestW = d0[0], bestH = d0[1];
            double bestScale = baseScale;
            bool baseClipped = TouchesBorder(probe);
            if (baseClipped)
            {
                try
                {
                    int hwnd2 = (int)app.MainWindow;
                    if (hwnd2 != 0) MoveWindow(new IntPtr(hwnd2), 0, 0, 2600, 2000, true);
                }
                catch { }
                System.Threading.Thread.Sleep(900);
                d0 = shot(baseScale);
                bestW = d0[0]; bestH = d0[1];
            }

            double good = baseScale;
            double candidate = baseScale;
            for (int i = 0; i < 8; i++)
            {
                candidate = good * 1.15;
                if (maxEdge > 0 && Math.Max(candidate * 1.0, 0) > 0 && bestW > 0)
                {
                    double predicted = bestW * (candidate / bestScale);
                    if (predicted > maxEdge) break;
                }
                int[] d = shot(candidate);
                if (TouchesBorder(probe)) break;
                good = candidate;
                bestW = d[0]; bestH = d[1]; bestScale = candidate;
            }

            doc.ScaleFactor = bestScale;
            System.Threading.Thread.Sleep(250);
            doc.SaveAs(finalPng);

            doc.SaveAs(molPath);
            bool gotC3xml = false;
            try { doc.SaveAs(c3xmlPath); gotC3xml = System.IO.File.Exists(c3xmlPath); } catch { gotC3xml = false; }

            double steric = double.NaN, rms = double.NaN;
            try { steric = Convert.ToDouble(doc.StericEnergy); } catch { }
            try { rms = Convert.ToDouble(doc.RMSGradient); } catch { }

            sb.Append("{\"ok\":true");
            sb.Append(",\"png\":\"" + Esc(finalPng) + "\"");
            sb.Append(",\"png_width\":" + bestW + ",\"png_height\":" + bestH);
            sb.Append(",\"mol\":\"" + Esc(molPath) + "\"");
            if (gotC3xml) sb.Append(",\"c3xml\":\"" + Esc(c3xmlPath) + "\"");
            sb.Append(",\"steric_energy\":" + Num(steric));
            sb.Append(",\"rms_gradient\":" + Num(rms));
            sb.Append(",\"scale_factor\":" + Num(bestScale));
            sb.Append(",\"model\":\"MM2\"}");
            finished = true;

            if (!keepOpen)
            {
                try { doc.Close(0, null); } catch { }
                doc = null;
            }
        }
        catch (Exception e)
        {
            sb.Length = 0;
            sb.Append("{\"ok\":false,\"error\":\"" + Esc(e.GetType().Name + ": " + e.Message) + "\"}");
        }
        finally
        {
            if (finished && !keepOpen)
            {
                try { app.Quit(); } catch { }
            }
            if (finished && keepOpen)
            {
                try { doc.Activate(); app.Visible = true; } catch { }
            }
            try { System.IO.File.Delete(outDir + "\\" + baseName + "_probe.png"); } catch { }
        }
        return sb.ToString();
    }
}
'@

$results = ''
try {
  Add-Type -TypeDefinition $code -ReferencedAssemblies 'System.Core', 'Microsoft.CSharp', 'System.Drawing' -ErrorAction Stop 3>$null
  $inPath = [string]$opt.input
  $outDir = [string]$opt.out_dir
  $base = [string]$opt.basename
  if (-not (Test-Path -LiteralPath $inPath)) { throw "input not found: $inPath" }
  if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
  $iterations = 2000; if ($opt.iterations) { $iterations = [int]$opt.iterations }
  $rmsg = 0.01; if ($opt.rmsg) { $rmsg = [double]$opt.rmsg }
  $visible = $false; if ($null -ne $opt.visible) { $visible = [bool]$opt.visible }
  $keep = $false; if ($null -ne $opt.open) { $keep = [bool]$opt.open }
  $atomSize = 0; if ($opt.style -and $opt.style.atom_size) { $atomSize = [int]$opt.style.atom_size }
  $bondWidth = 0; if ($opt.style -and $opt.style.bond_width) { $bondWidth = [int]$opt.style.bond_width }
  $displayType = -1; if ($opt.style -and $null -ne $opt.style.display_type) { $displayType = [int]$opt.style.display_type }
  $bgColor = -1; if ($opt.style -and $null -ne $opt.style.background_color) { $bgColor = [int]$opt.style.background_color }
  $maxEdge = 1100; if ($opt.max_edge) { $maxEdge = [int]$opt.max_edge }

  $before = @(Get-Process Chem3D -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
  # Chem3D's COM activation occasionally fails right after a previous instance
  # shuts down (0x800711C7 / 0x80040154); retry a few times before giving up.
  for ($attempt = 0; $attempt -lt 3; $attempt++) {
    $results = [Chem3DRun]::Execute($inPath, $outDir, $base, $iterations, $rmsg, $visible, $keep, $atomSize, $bondWidth, $displayType, $bgColor, $maxEdge)
    if ($results -notmatch '"ok":false') { break }
    if ($results -notmatch '800711C7|80040154|80080005|COMException') { break }
    Start-Sleep -Seconds 2
    Get-Process Chem3D -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | Stop-Process -Force -ErrorAction SilentlyContinue
  }
  if (-not $keep) {
    Start-Sleep -Milliseconds 400
    Get-Process Chem3D -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | Stop-Process -Force -ErrorAction SilentlyContinue
  }
}
catch {
  $msg = $_.Exception.Message.Replace('\', '/').Replace('"', '\"')
  $results = "{`"ok`":false,`"error`":`"$msg`"}"
}

$obj = $null
try { $obj = $results | ConvertFrom-Json } catch { $obj = $null }
if ($obj -and $obj.ok) {
  try {
    $lines = [System.IO.File]::ReadAllLines([string]$obj.mol)
    if ($lines.Count -ge 4) {
      $counts = $lines[3].Trim().Split(' ')
      $obj | Add-Member -NotePropertyName atoms -NotePropertyValue ([int]$counts[0]) -Force
      $obj | Add-Member -NotePropertyName bonds -NotePropertyValue ([int]$counts[1]) -Force
    }
  } catch { }
  # Chem3D exports the model view at roughly 550 px; scale it up once so the
  # picture is presentation friendly instead of tiny.
  try {
    $w = [int]$obj.png_width; $h = [int]$obj.png_height
    if ($w -gt 0 -and $h -gt 0 -and [Math]::Max($w, $h) -lt 900) {
      $factor = 1.9
      $newW = [int]($w * $factor); $newH = [int]($h * $factor)
      $src = [System.Drawing.Image]::FromFile([string]$obj.png)
      $bmp = New-Object System.Drawing.Bitmap($newW, $newH)
      $g = [System.Drawing.Graphics]::FromImage($bmp)
      $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
      $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
      $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
      $g.DrawImage($src, 0, 0, $newW, $newH)
      $g.Dispose(); $src.Dispose()
      $bmp.Save([string]$obj.png, [System.Drawing.Imaging.ImageFormat]::Png)
      $bmp.Dispose()
      $obj | Add-Member -NotePropertyName png_width -NotePropertyValue $newW -Force
      $obj | Add-Member -NotePropertyName png_height -NotePropertyValue $newH -Force
    }
  } catch { }
  Emit ($obj | ConvertTo-Json -Compress -Depth 6)
} else {
  Emit $results
}
