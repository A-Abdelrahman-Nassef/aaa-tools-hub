param([string]$pptx, [string]$outDir, [int]$width = 1280)
# Export every slide of a .pptx to PNG with PowerPoint (for checking a deck). Needs Microsoft PowerPoint.
# Usage: powershell -File tools/pptx_to_png.ps1 "<deck>.pptx" "<out folder>"
New-Item -ItemType Directory -Force $outDir | Out-Null
$pp = New-Object -ComObject PowerPoint.Application
try {
  $pres = $pp.Presentations.Open((Resolve-Path -LiteralPath $pptx).Path, $true, $false, $false)
  $h = [int]($width * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth)
  foreach ($s in $pres.Slides) { $s.Export((Join-Path (Resolve-Path $outDir).Path ("slide{0:D2}.png" -f $s.SlideIndex)), 'PNG', $width, $h) }
  "exported $($pres.Slides.Count) slides to $outDir"
  $pres.Close()
} finally { $pp.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
