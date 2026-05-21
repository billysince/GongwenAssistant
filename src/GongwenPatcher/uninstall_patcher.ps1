$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$installDir = 'C:\ProgramData\GongwenAssistant\Patcher'
$progId = 'A_GongwenPatcher.Connect'
$regasm = 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe'
$targetDll = Join-Path $installDir 'GongwenPatcher.dll'

if (Test-Path $targetDll) {
    Write-Host "[1] regasm /u"
    & $regasm $targetDll /u /nologo 2>&1 | ForEach-Object { Write-Host "    $_" }
}

$clsid = '{9F1A2B3C-4D5E-6F70-8190-A1B2C3D4E5F6}'

foreach ($p in @(
    "HKCU:\SOFTWARE\Microsoft\Office\Word\Addins\$progId",
    "HKCU:\Software\Classes\$progId",
    "HKCU:\Software\Classes\Wow6432Node\$progId",
    "HKCU:\Software\Classes\CLSID\$clsid",
    "HKCU:\Software\Classes\Wow6432Node\CLSID\$clsid"
)) {
    if (Test-Path $p) { Remove-Item -Path $p -Recurse -Force; Write-Host "  removed $p" }
}

$wl = 'HKCU:\Software\Kingsoft\Office\WPS\AddinsWL'
if (Test-Path $wl) {
    Remove-ItemProperty -Path $wl -Name $progId -ErrorAction SilentlyContinue
    Write-Host "  cleaned $wl::$progId"
}

if (Test-Path $installDir) {
    Remove-Item $installDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "  removed $installDir"
}
Write-Host 'UNINSTALL OK'
