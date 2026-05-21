$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

$srcBin = Join-Path $here 'bin'
$installDir = "$env:LOCALAPPDATA\GongwenAssistant\Patcher"
$dllName = 'GongwenPatcher.dll'
$harmonyName = '0Harmony.dll'
$progId = 'A_GongwenPatcher.Connect'
$clsid = '{9F1A2B3C-4D5E-6F70-8190-A1B2C3D4E5F6}'

Write-Host "[1] copy dlls to $installDir"
New-Item -ItemType Directory -Force $installDir | Out-Null
Copy-Item (Join-Path $srcBin $dllName)     (Join-Path $installDir $dllName)     -Force
Copy-Item (Join-Path $srcBin $harmonyName) (Join-Path $installDir $harmonyName) -Force
$pdb = Join-Path $srcBin 'GongwenPatcher.pdb'
if (Test-Path $pdb) { Copy-Item $pdb (Join-Path $installDir 'GongwenPatcher.pdb') -Force }
$targetDll = Join-Path $installDir $dllName

Write-Host "[2] HKCU CLSID + ProgId (both 64-bit AND 32-bit views)"
$asmFull = 'GongwenPatcher, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'
$codeBase = 'file:///' + ($targetDll -replace '\\', '/')

# ???: WPS ?? 32-bit, ?????? HKCU\Software\Classes ??? OS ????? Wow6432Node ???
# ???????? CLSID/ProgId ??§Õ?????????.
function Write-ComRegistration([string]$root) {
    $piPath = "$root\$progId"
    New-Item -Path $piPath -Force | Out-Null
    New-ItemProperty -Path $piPath -Name '(default)' -Value 'Gongwen Patcher' -PropertyType String -Force | Out-Null
    $piCls = "$piPath\CLSID"
    New-Item -Path $piCls -Force | Out-Null
    New-ItemProperty -Path $piCls -Name '(default)' -Value $clsid -PropertyType String -Force | Out-Null

    $clsRoot = "$root\CLSID\$clsid"
    New-Item -Path $clsRoot -Force | Out-Null
    New-ItemProperty -Path $clsRoot -Name '(default)' -Value 'GongwenPatcher.Connect' -PropertyType String -Force | Out-Null

    $inproc = "$clsRoot\InprocServer32"
    New-Item -Path $inproc -Force | Out-Null
    New-ItemProperty -Path $inproc -Name '(default)'      -Value 'mscoree.dll'             -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $inproc -Name 'ThreadingModel' -Value 'Both'                    -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $inproc -Name 'Class'          -Value 'GongwenPatcher.Connect'  -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $inproc -Name 'Assembly'       -Value $asmFull                  -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $inproc -Name 'RuntimeVersion' -Value 'v4.0.30319'              -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $inproc -Name 'CodeBase'       -Value $codeBase                 -PropertyType String -Force | Out-Null

    $ver = "$inproc\1.0.0.0"
    New-Item -Path $ver -Force | Out-Null
    New-ItemProperty -Path $ver -Name 'Class'          -Value 'GongwenPatcher.Connect' -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $ver -Name 'Assembly'       -Value $asmFull                 -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $ver -Name 'RuntimeVersion' -Value 'v4.0.30319'             -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $ver -Name 'CodeBase'       -Value $codeBase                -PropertyType String -Force | Out-Null

    $clsProgId = "$clsRoot\ProgId"
    New-Item -Path $clsProgId -Force | Out-Null
    New-ItemProperty -Path $clsProgId -Name '(default)' -Value $progId -PropertyType String -Force | Out-Null

    $cat = "$clsRoot\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}"
    New-Item -Path $cat -Force | Out-Null
}

# 64-bit view (PowerShell 64-bit ???§Õ??)
Write-ComRegistration 'HKCU:\Software\Classes'
# 32-bit view (WPS ???????¦Ë??)
Write-ComRegistration 'HKCU:\Software\Classes\Wow6432Node'

Write-Host "[3] HKCU Word Addins entry: $progId"
$addinsKey = "HKCU:\SOFTWARE\Microsoft\Office\Word\Addins\$progId"
if (-not (Test-Path $addinsKey)) { New-Item -Path $addinsKey -Force | Out-Null }
New-ItemProperty -Path $addinsKey -Name 'FriendlyName'    -Value 'Gongwen Patcher (Harmony unlock)' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinsKey -Name 'Description'     -Value 'Patches Local_Wps_Vsto.UserUtil IsVip/HasLogin at runtime' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinsKey -Name 'LoadBehavior'    -Value 3 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $addinsKey -Name 'CommandLineSafe' -Value 1 -PropertyType DWord -Force | Out-Null

Write-Host "[4] HKCU WPS AddinsWL whitelist (preserve existing entries)"
$wlKey = 'HKCU:\Software\Kingsoft\Office\WPS\AddinsWL'
if (-not (Test-Path $wlKey)) { New-Item -Path $wlKey -Force | Out-Null }
New-ItemProperty -Path $wlKey -Name $progId                  -Value '' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $wlKey -Name 'Local_Wps_Vsto.MyAddin' -Value '' -PropertyType String -Force | Out-Null

Write-Host '[5] verify'
$checks = @(
    "HKCU:\Software\Classes\$progId",
    "HKCU:\Software\Classes\Wow6432Node\$progId",
    "HKCU:\Software\Classes\CLSID\$clsid",
    "HKCU:\Software\Classes\Wow6432Node\CLSID\$clsid",
    "HKCU:\Software\Classes\Wow6432Node\CLSID\$clsid\InprocServer32",
    $addinsKey,
    $wlKey
)
foreach($p in $checks) {
    if (Test-Path $p) { Write-Host "  OK   $p" } else { Write-Host "  MISS $p" }
}
Write-Host '[5b] AddinsWL members:'
(Get-Item $wlKey).Property | ForEach-Object { Write-Host "    $_" }

Write-Host 'INSTALL OK'
