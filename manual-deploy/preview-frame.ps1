# preview-frame.ps1
# 使い方: .\preview-frame.ps1 -img "screenshots/35-1-before.png" -top 93 -left 11 -w 5 -h 7
param(
  [string]$img,
  [double]$top,
  [double]$left,
  [double]$w,
  [double]$h
)
Add-Type -AssemblyName System.Drawing
$base = "C:\gas\unko2-kanri\manual-deploy\deploy-assets"
$src = Join-Path $base $img
$out = Join-Path $base "frame-preview.png"
$bmp = [System.Drawing.Bitmap]::new($src)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Yellow, 3)
$x = [int]($bmp.Width  * $left / 100)
$y = [int]($bmp.Height * $top  / 100)
$rw = [int]($bmp.Width  * $w / 100)
$rh = [int]($bmp.Height * $h / 100)
$g.DrawRectangle($pen, $x, $y, $rw, $rh)
$g.Dispose()
$bmp.Save($out)
$bmp.Dispose()
Write-Host "Saved: $out  (rect: x=${x} y=${y} w=${rw} h=${rh} on $($bmp.Width)x$($bmp.Height))"
