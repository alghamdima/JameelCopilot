<#
.SYNOPSIS
  Renders the Jameel Copilot mark (VSTO2\JameelCopilot\WebUI\jc-mark.svg) to the image files
  the add-in ships: Resources\jc-256.png and Resources\jc.ico.

.DESCRIPTION
  Run it only when the mark changes, then commit the two files. The build does
  not run it. It needs Microsoft Edge (to draw the SVG) and writes nothing
  outside the repository and a temporary folder.

  The page shows the mark at eight sizes side by side. The mark is the brand
  kit's emblem PNG (105 px) inside the SVG, so the larger sizes are scaled up
  from it; a vector emblem would make them sharper. Edge cannot take a
  screenshot with a transparent background, so the page is captured twice, on
  white and on black, and each pixel's transparency is worked out from the
  difference between the two.

.EXAMPLE
  .\tools\Render-Logo.ps1
#>
param(
    [string]$Edge
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$svg = Join-Path $repo 'VSTO2\JameelCopilot\WebUI\jc-mark.svg'
$outDir = Join-Path $repo 'VSTO2\JameelCopilot\Resources'
$sizes = 16, 20, 24, 32, 40, 48, 64, 256

if (-not $Edge) {
    $Edge = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
    ) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
}
if (-not $Edge) { throw 'Microsoft Edge was not found. Pass its path with -Edge.' }
if (-not (Test-Path -LiteralPath $svg)) { throw "The mark was not found: $svg" }

$work = Join-Path ([IO.Path]::GetTempPath()) ('jc-logo-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $work | Out-Null
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

# One row of marks, each at its exact size.
$svgUrl = 'file:///' + ($svg -replace '\\', '/')
$x = 0
$tags = foreach ($size in $sizes) {
    '<img src="{0}" width="{1}" height="{1}" style="position:absolute;left:{2}px;top:0">' -f $svgUrl, $size, $x
    $x += $size
}
$width = $x
$height = ($sizes | Measure-Object -Maximum).Maximum

function Save-Sheet([string]$Background) {
    $page = Join-Path $work "sheet-$Background.html"
    $html = '<!doctype html><html><head><meta charset="utf-8"><style>html,body{margin:0;background:' + $Background +
        '}</style></head><body>' + ($tags -join '') + '</body></html>'
    Set-Content -LiteralPath $page -Value $html -Encoding ascii
    $picture = Join-Path $work "sheet-$Background.png"
    $edgeArgs = '--headless=new --disable-gpu --no-first-run --hide-scrollbars --force-device-scale-factor=1 ' +
        ('--user-data-dir="{0}" --window-size={1},{2} --screenshot="{3}" "file:///{4}"' -f
            (Join-Path $work 'edge'), $width, $height, $picture, ($page -replace '\\', '/'))
    $run = Start-Process -FilePath $Edge -ArgumentList $edgeArgs -Wait -PassThru
    if ($run.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $picture)) {
        throw "Edge did not produce the picture (exit code $($run.ExitCode))."
    }
    return $picture
}

Add-Type -AssemblyName System.Drawing
Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @'
using System;
using System.Drawing;

public static class JcLogoAlpha
{
    // A pixel seen on white (w) and on black (b): its opacity is 255 - (w - b),
    // and its own colour is what showed on black, divided by that opacity.
    public static Bitmap Combine(Bitmap onWhite, Bitmap onBlack)
    {
        var result = new Bitmap(onWhite.Width, onWhite.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        for (var y = 0; y < onWhite.Height; y++)
        {
            for (var x = 0; x < onWhite.Width; x++)
            {
                var w = onWhite.GetPixel(x, y);
                var b = onBlack.GetPixel(x, y);
                var alpha = 255 - (int)Math.Round(((w.R - b.R) + (w.G - b.G) + (w.B - b.B)) / 3.0);
                alpha = Math.Max(0, Math.Min(255, alpha));
                if (alpha == 0) { result.SetPixel(x, y, Color.FromArgb(0, 0, 0, 0)); continue; }
                result.SetPixel(x, y, Color.FromArgb(alpha, Own(b.R, alpha), Own(b.G, alpha), Own(b.B, alpha)));
            }
        }
        return result;
    }

    private static int Own(int onBlack, int alpha)
    {
        return Math.Max(0, Math.Min(255, (int)Math.Round(onBlack * 255.0 / alpha)));
    }
}
'@

$white = New-Object System.Drawing.Bitmap((Save-Sheet 'white'))
$black = New-Object System.Drawing.Bitmap((Save-Sheet 'black'))
try {
    if ($white.Width -lt $width -or $white.Height -lt $height -or $black.Width -ne $white.Width -or $black.Height -ne $white.Height) {
        throw "The pictures are $($white.Width)x$($white.Height) and $($black.Width)x$($black.Height); expected at least ${width}x${height}."
    }
    $bitmap = [JcLogoAlpha]::Combine($white, $black)
}
finally {
    $white.Dispose()
    $black.Dispose()
}

try {
    $images = @()
    $x = 0
    foreach ($size in $sizes) {
        $rect = New-Object System.Drawing.Rectangle($x, 0, $size, $size)
        $cut = $bitmap.Clone($rect, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        # The mark has no background, so the very first pixel is clear, or nearly so at the small sizes.
        if ($cut.GetPixel(0, 0).A -ge 128) { throw "The $size px mark has no transparent corner." }
        $stream = New-Object System.IO.MemoryStream
        $cut.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
        $cut.Dispose()
        $images += , @{ Size = $size; Bytes = $stream.ToArray() }
        $stream.Dispose()
        $x += $size
    }
}
finally {
    $bitmap.Dispose()
}

$master = $images | Where-Object { $_.Size -eq 256 }
[IO.File]::WriteAllBytes((Join-Path $outDir 'jc-256.png'), $master.Bytes)

# An .ico file whose images are stored as PNG: a 6-byte header, one 16-byte
# entry per image, then the images.
$ico = New-Object System.IO.MemoryStream
$writer = New-Object System.IO.BinaryWriter($ico)
$writer.Write([uint16]0); $writer.Write([uint16]1); $writer.Write([uint16]$images.Count)
$offset = 6 + 16 * $images.Count
foreach ($image in $images) {
    $dimension = [byte]($image.Size % 256)   # 256 is written as 0
    $writer.Write($dimension); $writer.Write($dimension)
    $writer.Write([byte]0); $writer.Write([byte]0)
    $writer.Write([uint16]1); $writer.Write([uint16]32)
    $writer.Write([uint32]$image.Bytes.Length); $writer.Write([uint32]$offset)
    $offset += $image.Bytes.Length
}
foreach ($image in $images) { $writer.Write([byte[]]$image.Bytes) }
$writer.Flush()
[IO.File]::WriteAllBytes((Join-Path $outDir 'jc.ico'), $ico.ToArray())
$writer.Dispose()

Write-Host ("Wrote {0} and {1}" -f (Join-Path $outDir 'jc-256.png'), (Join-Path $outDir 'jc.ico'))
