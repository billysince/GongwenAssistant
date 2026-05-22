# =============================================================================
# 公文助手 GongwenAssistant 自检脚本 (v1.0.1, 双路线版)
# =============================================================================
#  支持两条互斥路线的安装自检：
#    路线 B · 自家重编译版    (默认)
#      - %LOCALAPPDATA%\GongwenAssistant\Local_Wps_Vsto.dll  (弱命名 dll, IL 已 patch)
#      - HKCU\...\Word\Addins\Local_Wps_Vsto.MyAddin
#      - HKCU\...\Classes\CLSID\{2CD4E522-...}
#    路线 A · Patcher 运行时注入
#      - %LOCALAPPDATA%\GongwenAssistant\Patcher\GongwenPatcher.dll
#      - %LOCALAPPDATA%\GongwenAssistant\Patcher\0Harmony.dll
#      - HKCU\...\Word\Addins\A_GongwenPatcher.Connect
#      - HKCU\...\Classes\Wow6432Node\CLSID\{9F1A2B3C-...}     ← 32-bit 视图必备
#
#  默认 Auto 模式: 同时跑 A + B 两条路线的检查，并按"未装/已装"区分输出.
#  显式选择: -Mode A 或 -Mode B 只检查一条.
# =============================================================================

[CmdletBinding()]
param(
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA 'GongwenAssistant'),
    [ValidateSet('Auto','A','B')]
    [string]$Mode = 'Auto'
)

$ErrorActionPreference = 'Continue'
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ---------- 路线 B (重编译) 常量 ----------
$ProgIdB = 'Local_Wps_Vsto.MyAddin'
$ClsIdB  = '{2CD4E522-63D0-3A6B-9842-587DB1AB7FBF}'

# ---------- 路线 A (Patcher) 常量 ----------
$ProgIdA = 'A_GongwenPatcher.Connect'
$ClsIdA  = '{9F1A2B3C-4D5E-6F70-8190-A1B2C3D4E5F6}'
$PatcherDir = Join-Path $InstallDir 'Patcher'

$ok = 0; $fail = 0; $skip = 0
function Pass($m) { Write-Host "[PASS] $m" -ForegroundColor Green;  $script:ok++ }
function Fail($m) { Write-Host "[FAIL] $m" -ForegroundColor Red;    $script:fail++ }
function Skip($m) { Write-Host "[SKIP] $m" -ForegroundColor DarkGray; $script:skip++ }
function Info($m) { Write-Host "[INFO] $m" -ForegroundColor Cyan }

Write-Host '========================================================================='
Write-Host '公文助手 GongwenAssistant 自检 (v1.0.1)'
Write-Host '========================================================================='
Write-Host ('  InstallDir = ' + $InstallDir)
Write-Host ('  PatcherDir = ' + $PatcherDir)
Write-Host ('  Mode       = ' + $Mode)
Write-Host ''

# ============================================================================
# 检测已装路线
# ============================================================================
$bDetected = (Test-Path (Join-Path $InstallDir 'Local_Wps_Vsto.dll')) -or
             (Test-Path "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgIdB")
$aDetected = (Test-Path (Join-Path $PatcherDir 'GongwenPatcher.dll')) -or
             (Test-Path "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgIdA")

if ($Mode -eq 'Auto') {
    if (-not $bDetected -and -not $aDetected) {
        Write-Host '[WARN] 两条路线均未检测到, 默认按 B 路线检查 (会大量 FAIL).' -ForegroundColor Yellow
    } elseif ($bDetected -and $aDetected) {
        Write-Host '[WARN] A + B 两条路线同时检测到 - 它们用的 ProgId 不冲突, 可共存,' -ForegroundColor Yellow
        Write-Host '       但路线 A 会在 WPS 进程内对 B 路线的 dll 也做一次 IL hook,' -ForegroundColor Yellow
        Write-Host '       行为上仍然安全. 如果想二选一, 卸掉其中一条即可.' -ForegroundColor Yellow
    }
}

# ============================================================================
# 路线 B 检查 (16 项)
# ============================================================================
if ($Mode -eq 'Auto' -or $Mode -eq 'B') {
    Write-Host ''
    Write-Host '------------------------------------------------------------'
    Write-Host '路线 B · 自家重编译版'
    Write-Host '------------------------------------------------------------'

    if (-not $bDetected -and $Mode -eq 'Auto') {
        Skip 'B 路线未安装, 跳过 16 项检查'
    } else {
        # B-1. 文件存在
        $dll = Join-Path $InstallDir 'Local_Wps_Vsto.dll'
        if (Test-Path $dll) { Pass "B-Local_Wps_Vsto.dll 存在 ($((Get-Item $dll).Length) bytes)" } else { Fail "B-缺 $dll" }

        foreach ($f in 'Newtonsoft.Json.dll','Spire.Doc.dll','System.Data.SqlServerCe.dll','ICSharpCode.SharpZipLib.dll') {
            $p = Join-Path $InstallDir $f
            if (Test-Path $p) { Pass "B-依赖 $f 存在" } else { Fail "B-缺依赖 $f" }
        }

        # B-2. 注册表
        foreach ($k in "HKCU:\Software\Classes\CLSID\$ClsIdB\InprocServer32",
                       "HKCU:\Software\Classes\$ProgIdB\CLSID",
                       "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgIdB") {
            if (Test-Path $k) { Pass "B-注册表 $k 存在" } else { Fail "B-缺注册表 $k" }
        }

        $lb = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgIdB" -Name 'LoadBehavior' -ErrorAction SilentlyContinue).LoadBehavior
        if ($lb -eq 3) { Pass "B-LoadBehavior=3" } else { Fail "B-LoadBehavior=$lb (应为 3)" }

        $wl = Get-ItemProperty -Path 'HKCU:\Software\Kingsoft\Office\WPS\AddinsWL' -ErrorAction SilentlyContinue
        if ($wl -and $wl.PSObject.Properties[$ProgIdB]) { Pass "B-WPS AddinsWL 包含 $ProgIdB" } else { Fail "B-WPS AddinsWL 缺 $ProgIdB" }

        # B-3. CodeBase 指向正确
        $inp = Get-ItemProperty -Path "HKCU:\Software\Classes\CLSID\$ClsIdB\InprocServer32" -ErrorAction SilentlyContinue
        if ($inp) {
            $cb = $inp.CodeBase
            $asm = $inp.Assembly
            if ($cb -match $InstallDir.Replace('\','/')) { Pass "B-CodeBase 指向 $cb" } else { Fail "B-CodeBase 路径异常: $cb" }
            if ($asm -match 'PublicKeyToken=null') { Pass "B-Assembly 弱命名 (PKT=null)" } else { Fail "B-Assembly 仍是强名: $asm" }
        }

        # B-4. CLR 能加载并验证 IL
        if (Test-Path $dll) {
            try {
                $asmOb = [Reflection.Assembly]::ReflectionOnlyLoadFrom($dll)
                $name = $asmOb.GetName()
                if ($name.GetPublicKeyToken().Length -eq 0) { Pass "B-PE 文件确实弱命名" } else { Fail "B-PE 文件居然有 PKT" }
                $t = $asmOb.GetType('Local_Wps_Vsto.UserUtil')
                $mIs = $t.GetMethod('IsVip',   [Reflection.BindingFlags]'Public,Static,NonPublic')
                $mHa = $t.GetMethod('HasLogin',[Reflection.BindingFlags]'Public,Static,NonPublic')
                $bIs = ($mIs.GetMethodBody().GetILAsByteArray() | ForEach-Object { '{0:X2}' -f $_ }) -join ''
                $bHa = ($mHa.GetMethodBody().GetILAsByteArray() | ForEach-Object { '{0:X2}' -f $_ }) -join ''
                if ($bIs -eq '172A') { Pass "B-UserUtil.IsVip()   IL = 17 2A -> true" } else { Fail "B-IsVip   IL=$bIs" }
                if ($bHa -eq '172A') { Pass "B-UserUtil.HasLogin() IL = 17 2A -> true" } else { Fail "B-HasLogin IL=$bHa" }
            } catch {
                Fail "B-ReflectionOnly 加载失败: $($_.Exception.Message)"
            }
        }
    }
}

# ============================================================================
# 路线 A 检查 (8 项)
# ============================================================================
if ($Mode -eq 'Auto' -or $Mode -eq 'A') {
    Write-Host ''
    Write-Host '------------------------------------------------------------'
    Write-Host '路线 A · Patcher 运行时注入'
    Write-Host '------------------------------------------------------------'

    if (-not $aDetected -and $Mode -eq 'Auto') {
        Skip 'A 路线未安装, 跳过 8 项检查'
    } else {
        # A-1. 文件存在
        $patcherDll = Join-Path $PatcherDir 'GongwenPatcher.dll'
        $harmonyDll = Join-Path $PatcherDir '0Harmony.dll'
        if (Test-Path $patcherDll) { Pass "A-GongwenPatcher.dll 存在 ($((Get-Item $patcherDll).Length) bytes)" } else { Fail "A-缺 $patcherDll" }
        if (Test-Path $harmonyDll) { Pass "A-0Harmony.dll 存在 ($((Get-Item $harmonyDll).Length) bytes, 应 ~2.2 MB Fat 单文件)" } else { Fail "A-缺 $harmonyDll" }

        # A-2. 32-bit 视图 CLSID 注册 (WPS 32-bit 进程实际查的位置)
        $clsId32 = "HKCU:\Software\Classes\Wow6432Node\CLSID\$ClsIdA\InprocServer32"
        if (Test-Path $clsId32) { Pass "A-32位视图 CLSID 注册存在 (WPS 实际查询路径)" } else { Fail "A-缺 32位视图 CLSID 注册 (致命: 32-bit WPS 看不到)" }

        # A-3. ProgId 32-bit 视图
        $pi32 = "HKCU:\Software\Classes\Wow6432Node\$ProgIdA"
        if (Test-Path $pi32) { Pass "A-32位视图 ProgId 注册存在" } else { Fail "A-缺 32位视图 ProgId 注册" }

        # A-4. Word Addins LoadBehavior
        $aLb = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgIdA" -Name 'LoadBehavior' -ErrorAction SilentlyContinue).LoadBehavior
        if ($aLb -eq 3) { Pass "A-LoadBehavior=3" } else { Fail "A-LoadBehavior=$aLb (应为 3)" }

        # A-5. AddinsWL 含 Patcher
        $wlA = Get-ItemProperty -Path 'HKCU:\Software\Kingsoft\Office\WPS\AddinsWL' -ErrorAction SilentlyContinue
        if ($wlA -and $wlA.PSObject.Properties[$ProgIdA]) { Pass "A-WPS AddinsWL 包含 $ProgIdA" } else { Fail "A-WPS AddinsWL 缺 $ProgIdA" }

        # A-6. CodeBase 指向 Patcher 目录
        $aInp = Get-ItemProperty -Path $clsId32 -ErrorAction SilentlyContinue
        if ($aInp) {
            $aCb = $aInp.CodeBase
            if ($aCb -match $PatcherDir.Replace('\','/')) { Pass "A-CodeBase 指向 $aCb" } else { Fail "A-CodeBase 路径异常: $aCb" }
        }

        # A-7. Patcher dll 是合法 .NET assembly
        if (Test-Path $patcherDll) {
            try {
                $aAsm = [Reflection.Assembly]::ReflectionOnlyLoadFrom($patcherDll)
                $aTy = $aAsm.GetType('GongwenPatcher.Connect')
                if ($aTy) { Pass "A-GongwenPatcher.Connect 类型可解析" } else { Fail "A-Connect 类型缺失" }
            } catch {
                Fail "A-ReflectionOnly 加载失败: $($_.Exception.Message)"
            }
        }

        # A-8. AddinsCL 应不含 ProgIdA (避免被误标 crash)
        $cl = Get-ItemProperty -Path 'HKCU:\Software\Kingsoft\Office\WPS\AddinsCL' -ErrorAction SilentlyContinue
        if ($cl -and $cl.PSObject.Properties[$ProgIdA]) {
            Fail "A-AddinsCL 标记 $ProgIdA 为 crash, 启动会弹禁用对话框 (重启 WPS 一次让 OnStartupComplete 自动清零)"
        } else {
            Pass "A-AddinsCL 不含 Patcher 条目 (健康)"
        }
    }
}

# ============================================================================
# 最终汇总
# ============================================================================
Write-Host ''
Write-Host '========================================================================='
$summaryColor = if ($fail -eq 0) { 'Green' } else { 'Red' }
Write-Host ("结果: PASS=" + $ok + " FAIL=" + $fail + " SKIP=" + $skip) -ForegroundColor $summaryColor
Write-Host '========================================================================='
if ($fail -gt 0) { exit 1 }
