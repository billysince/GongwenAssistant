$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
$out = $args[0]
if (-not $out) { $out = 'D:\นคื๗\20260520\GongwenAssistant\dist\snap.png' }
New-Item -ItemType Directory -Force (Split-Path $out -Parent) | Out-Null
$b = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$bm = New-Object System.Drawing.Bitmap $b.Width, $b.Height
$g = [System.Drawing.Graphics]::FromImage($bm)
$g.CopyFromScreen($b.X, $b.Y, 0, 0, $b.Size)
$g.Dispose()
$bm.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$bm.Dispose()
Write-Host "saved: $out  $((Get-Item $out).Length)b"
