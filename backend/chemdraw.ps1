# ChemDraw COM bridge for the DSH ChemDraw plugin.
# Invoked as:  powershell.exe -NoProfile -ExecutionPolicy Bypass -File chemdraw.ps1 -ArgsJson <file.json>
# Prints one line  @@JSON@@{...}  on stdout.
param(
    [Parameter(Mandatory = $true)][string]$ArgsJson
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$opt = Get-Content -LiteralPath $ArgsJson -Raw -Encoding UTF8 | ConvertFrom-Json

function Emit($obj) {
    Write-Output ('@@JSON@@' + ($obj | ConvertTo-Json -Compress -Depth 12))
}

function Get-ChemDraw {
    param([bool]$Visible = $false, [bool]$WeStarted = $false)
    $app = New-Object -ComObject 'ChemDraw_x64.Application'
    try { $app.Visible = $Visible } catch { }
    Start-Sleep -Milliseconds 700
    return $app
}

function Open-Doc($app, [string]$path) {
    for ($i = 0; $i -lt 5; $i++) {
        try {
            $doc = $app.Documents.Open($path, $null, $null)
            if ($doc) { return $doc }
        } catch {
            Start-Sleep -Milliseconds 700
        }
    }
    throw "ChemDraw could not open: $path"
}

function Find-Command($app, [int]$cmdId) {
    $mb = $app.MenuBars.Item([int16]1)
    $menus = $mb.Menus
    for ($j = 1; $j -le $menus.Count; $j++) {
        $items = $menus.Item($j).MenuItems
        for ($k = 1; $k -le $items.Count; $k++) {
            $it = $items.Item($k)
            $cid = 0
            try { $cid = $it.CommandID } catch { }
            if ($cid -eq $cmdId) { return $it }
        }
    }
    return $null
}

function Close-Doc($doc) {
    try { $doc.Close([ref]0, [ref]$null) } catch { }
}

function Get-ImageSize([string]$path) {
    try {
        Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
        $img = [System.Drawing.Image]::FromFile($path)
        $size = @{ width = $img.Width; height = $img.Height }
        $img.Dispose()
        return $size
    } catch {
        return $null
    }
}

function Apply-Style($doc, $style) {
    if (-not $style) { return }
    $s = $null
    try { $s = $doc.Settings } catch { return }
    if (-not $s) { return }
    foreach ($pair in @(
            @('bond_length', 'BondLength'),
            @('line_width', 'LineWidth'),
            @('label_size', 'LabelSize'),
            @('label_font', 'LabelFont'),
            @('caption_size', 'CaptionSize')
        )) {
        $key = $pair[0]; $prop = $pair[1]
        $value = $style.$key
        if ($null -ne $value) {
            try { $s.$prop = $value } catch { }
        }
    }
    if ($style.hide_implicit_hydrogens -eq $true) {
        try { $s.HideImplicitHydrogens = $true } catch { }
    }
}

$preExisting = @(Get-Process ChemDraw -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

try {
    switch ($opt.op) {

        'render' {
            $app = Get-ChemDraw -Visible $false
            $doc = Open-Doc $app $opt.input
            try { $doc.Activate() } catch { }
            Apply-Style $doc $opt.style
            if ($opt.clean -ne $false) {
                try { $doc.Objects.Clean($true) } catch { }
            }
            $doc.SaveAs($opt.output, $null, $null, $null, $null)
            $size = Get-ImageSize $opt.output
            $bytes = (Get-Item -LiteralPath $opt.output).Length
            Close-Doc $doc
            Emit @{ ok = $true; file = $opt.output; bytes = $bytes; image = $size }
        }

        'beautify' {
            $app = Get-ChemDraw -Visible $false
            $doc = Open-Doc $app $opt.input
            try { $doc.Activate() } catch { }
            Apply-Style $doc $opt.style
            try { $doc.Objects.Clean($true) } catch { }
            $files = @{}
            # Save raster/vector images first: saving to .cdx can reflow the page
            # and would then squash a later PNG export.
            $imageFormats = @('png', 'svg', 'gif', 'wmf', 'emf')
            $ordered = @()
            foreach ($fmt in $opt.formats) { if ($imageFormats -contains $fmt) { $ordered += $fmt } }
            foreach ($fmt in $opt.formats) { if ($imageFormats -notcontains $fmt) { $ordered += $fmt } }
            foreach ($fmt in $ordered) {
                $path = $opt.output_base + '.' + $fmt
                $doc.SaveAs($path, $null, $null, $null, $null)
                $files[$fmt] = $path
            }
            $size = $null
            if ($files.ContainsKey('png')) { $size = Get-ImageSize $files['png'] }
            Close-Doc $doc
            Emit @{ ok = $true; files = $files; image = $size }
        }

        'nmr' {
            $cmdId = if ($opt.nucleus -eq '13C') { 371 } else { 370 }
            $app = Get-ChemDraw -Visible ($opt.visible -eq $true)
            $doc = Open-Doc $app $opt.input
            $script:stage = 'activate'
            try { $doc.Activate() } catch { }
            $script:stage = 'select'
            try { $doc.Objects.Select() } catch { }
            Start-Sleep -Milliseconds 400
            $before = $app.Documents.Count
            $item = Find-Command $app $cmdId
            if ($null -eq $item) { throw 'ChemNMR menu command not found' }
            $script:stage = 'execute'
            if (-not $item.Enabled) {
                # second attempt after activating and selecting again
                try { $doc.Activate(); $doc.Objects.Select(); Start-Sleep -Milliseconds 300 } catch { }
            }
            $item.Execute()
            Start-Sleep -Seconds 2
            $after = $app.Documents.Count
            $script:stage = 'find-spectrum-doc'
            $specDoc = $null
            if ($after -gt $before) {
                $specDoc = $app.Documents.Item([object]$after)
            } else {
                for ($i = 1; $i -le $after; $i++) {
                    $d = $app.Documents.Item([object]$i)
                    if ($d.Name -ne $doc.Name) { $specDoc = $d }
                }
            }
            if ($null -eq $specDoc) { throw 'ChemNMR did not create a prediction document' }
            $captions = @()
            try {
                $caps = $specDoc.Captions
                for ($c = 1; $c -le $caps.Count; $c++) {
                    try { $captions += [string]$caps.Item($c).Text } catch { }
                }
            } catch { }
            $png = $opt.output_base + '.png'
            try { $specDoc.Activate() } catch { }
            Start-Sleep -Milliseconds 300
            $specDoc.SaveAs($png, $null, $null, $null, $null)
            $size = Get-ImageSize $png
            Close-Doc $specDoc
            Close-Doc $doc
            Emit @{
                ok       = $true
                png      = $png
                image    = $size
                captions = $captions
            }
        }

        'image' {
            Add-Type -AssemblyName System.Drawing
            $src = [System.Drawing.Image]::FromFile($opt.input)
            try {
                $scale = 1.0
                $minWidth = 900
                if ($opt.min_width) { $minWidth = [int]$opt.min_width }
                if ($src.Width -lt $minWidth) {
                    $scale = [Math]::Min(3.0, [Math]::Ceiling($minWidth / $src.Width))
                }
                $width = [int][Math]::Round($src.Width * $scale)
                $height = [int][Math]::Round($src.Height * $scale)
                $bmp = New-Object System.Drawing.Bitmap($width, $height)
                $bmp.SetResolution(96, 96)
                $g = [System.Drawing.Graphics]::FromImage($bmp)
                $g.Clear([System.Drawing.Color]::White)
                $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
                $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
                $g.DrawImage($src, 0, 0, $width, $height)
                $g.Dispose()
                $bmp.Save($opt.output, [System.Drawing.Imaging.ImageFormat]::Png)
                $bmp.Dispose()
                $size = Get-ImageSize $opt.output
                Emit @{ ok = $true; file = $opt.output; image = $size; scale = $scale }
            } finally {
                $src.Dispose()
            }
        }

        'open' {
            $app = Get-ChemDraw -Visible $true -WeStarted (-not $preExisting)
            $doc = Open-Doc $app $opt.input
            try { $doc.Activate() } catch { }
            if ($opt.keep -eq $true) { $script:keepAlive = $true }
            Emit @{ ok = $true; name = $doc.Name; started = (-not $preExisting) }
        }

        default {
            throw "unknown op: $($opt.op)"
        }
    }
} catch {
    Emit @{
        ok        = $false
        stage     = $script:stage
        error     = $_.Exception.Message
        exception = $_.Exception.GetType().FullName
    }
} finally {
    if ($script:keepAlive -ne $true) {
        # Only stop the ChemDraw instances this call created; leave any window the
        # user already had open untouched.
        $stillRunning = @(Get-Process ChemDraw -ErrorAction SilentlyContinue)
        foreach ($proc in $stillRunning) {
            if ($preExisting -notcontains $proc.Id) {
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            }
        }
    }
}
