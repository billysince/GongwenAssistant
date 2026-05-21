# =============================================================================
# 公文助手 GongwenAssistant 自检脚本
# =============================================================================
# 验证安装后所有关键项是否就位
#  - 文件是否存在
#  - 注册表项是否完整
#  - CLR 是否能从 CodeBase 加载弱命名 dll
#  - UserUtil.IsVip / HasLogin IL 是否为 ret true (17 2A)
# =============================================================================

[CmdletBinding()]
param(
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA 'GongwenAssistant')
)

$ErrorActionPreference = 'Continue'
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$ProgId = 'Local_Wps_Vsto.MyAddin'
$ClsId  = '{2CD4E522-63D0-3A6B-9842-587DB1AB7FBF}'

$ok = 0; $fail = 0
function Pass($m) { Write-Host "[PASS] $m" -ForegroundColor Green; $script:ok++ }
function Fail($m) { Write-Host "[FAIL] $m" -ForegroundColor Red;   $script:fail++ }

Write-Host '========================================================================='
Write-Host '公文助手 GongwenAssistant 自检'
Write-Host '========================================================================='

# 1. 文件存在
$dll = Join-Path $InstallDir 'Local_Wps_Vsto.dll'
if (Test-Path $dll) { Pass "Local_Wps_Vsto.dll 存在 ($((Get-Item $dll).Length) bytes)" } else { Fail "缺 $dll" }

foreach ($f in 'Newtonsoft.Json.dll','Spire.Doc.dll','System.Data.SqlServerCe.dll','ICSharpCode.SharpZipLib.dll') {
    $p = Join-Path $InstallDir $f
    if (Test-Path $p) { Pass "依赖 $f 存在" } else { Fail "缺依赖 $f" }
}

# 2. 注册表
foreach ($k in "HKCU:\Software\Classes\CLSID\$ClsId\InprocServer32",
               "HKCU:\Software\Classes\$ProgId\CLSID",
               "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId") {
    if (Test-Path $k) { Pass "注册表 $k 存在" } else { Fail "缺注册表 $k" }
}

# Office LoadBehavior
$lb = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId" -Name 'LoadBehavior' -ErrorAction SilentlyContinue).LoadBehavior
if ($lb -eq 3) { Pass "LoadBehavior=3 (开机加载)" } else { Fail "LoadBehavior=$lb (应为 3)" }

# WPS 白名单
$wl = Get-ItemProperty -Path 'HKCU:\Software\Kingsoft\Office\WPS\AddinsWL' -ErrorAction SilentlyContinue
if ($wl -and $wl.PSObject.Properties[$ProgId]) { Pass "WPS AddinsWL 包含 $ProgId" } else { Fail "WPS AddinsWL 缺 $ProgId" }

# 3. CodeBase 指向正确
$inp = Get-ItemProperty -Path "HKCU:\Software\Classes\CLSID\$ClsId\InprocServer32" -ErrorAction SilentlyContinue
if ($inp) {
    $cb = $inp.CodeBase
    $asm = $inp.Assembly
    if ($cb -match $InstallDir.Replace('\','/')) { Pass "CodeBase 指向 $cb" } else { Fail "CodeBase 路径异常: $cb" }
    if ($asm -match 'PublicKeyToken=null') { Pass "Assembly 使用 PublicKeyToken=null (弱命名)" } else { Fail "Assembly 仍是强名: $asm" }
}

# 4. CLR 能加载并验证 IL
if (Test-Path $dll) {
    try {
        $asmOb = [Reflection.Assembly]::ReflectionOnlyLoadFrom($dll)
        $name = $asmOb.GetName()
        if ($name.GetPublicKeyToken().Length -eq 0) { Pass "PE 文件确实是弱命名 (无 PKT)" } else { Fail "PE 文件居然有 PKT" }
        $t = $asmOb.GetType('Local_Wps_Vsto.UserUtil')
        $mIs = $t.GetMethod('IsVip',   [Reflection.BindingFlags]'Public,Static,NonPublic')
        $mHa = $t.GetMethod('HasLogin',[Reflection.BindingFlags]'Public,Static,NonPublic')
        $bIs = ($mIs.GetMethodBody().GetILAsByteArray() | ForEach-Object { '{0:X2}' -f $_ }) -join ''
        $bHa = ($mHa.GetMethodBody().GetILAsByteArray() | ForEach-Object { '{0:X2}' -f $_ }) -join ''
        if ($bIs -eq '172A') { Pass "UserUtil.IsVip()   IL = 17 2A (ldc.i4.1 ret) -> true" } else { Fail "IsVip   IL=$bIs (应为 172A)" }
        if ($bHa -eq '172A') { Pass "UserUtil.HasLogin() IL = 17 2A (ldc.i4.1 ret) -> true" } else { Fail "HasLogin IL=$bHa (应为 172A)" }
    } catch {
        Fail "ReflectionOnly 加载失败: $($_.Exception.Message)"
    }
}

Write-Host ''
Write-Host '========================================================================='
Write-Host ('结果: PASS=' + $ok + ' FAIL=' + $fail) -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
Write-Host '========================================================================='
