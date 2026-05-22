# =============================================================================
# 公文助手 GongwenAssistant 安装脚本（零障碍 / 仅 HKCU / 无需管理员）
# =============================================================================
# 设计目标:
#   1. 全部走 HKEY_CURRENT_USER 注册表分支, 不动 HKLM / GAC / Wow6432Node
#   2. dll 通过 CodeBase 指向本地路径加载, 不依赖强名 / GAC
#   3. 完整登记 COM CLSID / ProgId / Office Word Addin / WPS AddinsWL
#   4. 自动检测已有冲突的旧版「公文助手」注册并提示用户清理
#
# 用法:
#   PowerShell -ExecutionPolicy Bypass -File install.ps1 [-InstallDir <path>]
# =============================================================================

[CmdletBinding()]
param(
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA 'GongwenAssistant')
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# -----------------------------------------------------------------------------
# 常量
# -----------------------------------------------------------------------------
$ProgId    = 'Local_Wps_Vsto.MyAddin'
$ClsId     = '{2CD4E522-63D0-3A6B-9842-587DB1AB7FBF}'   # MyAddin 类的自动生成 CLSID, 与原版兼容
$AsmName   = 'Local_Wps_Vsto'
$AsmVer    = '1.0.0.0'
$FriendlyName = '公文助手 GongwenAssistant'

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RuntimeSrc = Join-Path (Split-Path -Parent $ScriptRoot) 'runtime'

function Info($m)   { Write-Host "[INFO ] $m" -ForegroundColor Cyan }
function Ok($m)     { Write-Host "[ OK  ] $m" -ForegroundColor Green }
function Warn($m)   { Write-Host "[WARN ] $m" -ForegroundColor Yellow }
function Err($m)    { Write-Host "[ERROR] $m" -ForegroundColor Red }

Info "公文助手 GongwenAssistant 安装程序启动"
Info "安装目录: $InstallDir"

# -----------------------------------------------------------------------------
# Step 0: 预检
# -----------------------------------------------------------------------------
if (-not (Test-Path $RuntimeSrc)) {
    Err "找不到 runtime 目录: $RuntimeSrc"
    Err "请确认本脚本与 runtime 目录是同一父目录下的 installer 子目录"
    exit 1
}
if (-not (Test-Path (Join-Path $RuntimeSrc 'Local_Wps_Vsto.dll'))) {
    Err "runtime 目录中缺少 Local_Wps_Vsto.dll"
    exit 1
}

# 关闭可能锁文件的 WPS 进程
$wps = Get-Process wps,wpsoffice,wpscloudsvr,wpscloudlaunch -ErrorAction SilentlyContinue
if ($wps) {
    Warn "检测到 WPS 正在运行, 将自动关闭以便完成注册"
    $wps | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

# -----------------------------------------------------------------------------
# Step 1: 复制 runtime 文件
# -----------------------------------------------------------------------------
Info "[1/5] 复制 runtime 文件到 $InstallDir"
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
Copy-Item -Path (Join-Path $RuntimeSrc '*') -Destination $InstallDir -Recurse -Force
Ok "runtime 文件已部署"

$DllPath  = Join-Path $InstallDir 'Local_Wps_Vsto.dll'
$TlbPath  = Join-Path $InstallDir 'Local_Wps_Vsto.tlb'
$DllUri   = 'file:///' + ($DllPath -replace '\\','/')

# -----------------------------------------------------------------------------
# Step 2: HKCU COM CLSID 注册（64 位 + 32 位 Wow6432Node 双视图）
# -----------------------------------------------------------------------------
# 关键: WPS 是 32 位进程, 必须在 Wow6432Node 视图写注册才能被 32 位进程看到!
# 历史 bug: 早期版本只写 64 位视图, 导致 v2 弱命名版从未真的被加载
# 正确做法: 同时写 HKCU\Software\Classes\CLSID\... 和 HKCU\Software\Classes\Wow6432Node\CLSID\...
# COM lookup 优先级: HKCU\Wow6432Node > HKLM\Wow6432Node, 这样 v2 才能 win 过 GAC 强名原版
function Write-CLSID-Registration {
    param([string]$ClsidRoot)
    New-Item -Path $ClsidRoot -Force | Out-Null
    Set-ItemProperty -Path $ClsidRoot -Name '(Default)' -Value $ProgId

    $Inproc = Join-Path $ClsidRoot 'InprocServer32'
    New-Item -Path $Inproc -Force | Out-Null
    Set-ItemProperty -Path $Inproc -Name '(Default)'      -Value 'mscoree.dll'
    Set-ItemProperty -Path $Inproc -Name 'ThreadingModel' -Value 'Both'
    Set-ItemProperty -Path $Inproc -Name 'Class'          -Value $ProgId
    Set-ItemProperty -Path $Inproc -Name 'Assembly'       -Value "$AsmName, Version=$AsmVer, Culture=neutral, PublicKeyToken=null"
    Set-ItemProperty -Path $Inproc -Name 'CodeBase'       -Value $DllUri
    Set-ItemProperty -Path $Inproc -Name 'RuntimeVersion' -Value 'v4.0.30319'

    $InprocVer = Join-Path $Inproc $AsmVer
    New-Item -Path $InprocVer -Force | Out-Null
    Set-ItemProperty -Path $InprocVer -Name 'Class'          -Value $ProgId
    Set-ItemProperty -Path $InprocVer -Name 'Assembly'       -Value "$AsmName, Version=$AsmVer, Culture=neutral, PublicKeyToken=null"
    Set-ItemProperty -Path $InprocVer -Name 'CodeBase'       -Value $DllUri
    Set-ItemProperty -Path $InprocVer -Name 'RuntimeVersion' -Value 'v4.0.30319'

    $ProgIdKey = Join-Path $ClsidRoot 'ProgId'
    New-Item -Path $ProgIdKey -Force | Out-Null
    Set-ItemProperty -Path $ProgIdKey -Name '(Default)' -Value $ProgId

    $ImplCat = Join-Path $ClsidRoot 'Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}'
    New-Item -Path $ImplCat -Force | Out-Null
}

Info "[2/5] 注册 COM CLSID (64 位视图 HKCU\Software\Classes\CLSID\$ClsId)"
Write-CLSID-Registration -ClsidRoot "HKCU:\Software\Classes\CLSID\$ClsId"
Ok "  64 位视图注册完成"

Info "[2/5] 注册 COM CLSID (32 位视图 HKCU\Software\Classes\Wow6432Node\CLSID\$ClsId)"
Write-CLSID-Registration -ClsidRoot "HKCU:\Software\Classes\Wow6432Node\CLSID\$ClsId"
Ok "  32 位视图注册完成 (WPS 32 位进程必需)"

# 检测 HKLM Wow6432Node 是否有原版强名注册压在下面
$HklmConflict = "HKLM:\Software\Wow6432Node\Classes\CLSID\$ClsId\InprocServer32"
if (Test-Path $HklmConflict) {
    $hklmAsm = (Get-ItemProperty $HklmConflict -ErrorAction SilentlyContinue).Assembly
    if ($hklmAsm -and $hklmAsm -match 'PublicKeyToken=([0-9a-f]{16})') {
        Warn "检测到 HKLM\Wow6432Node 存在原版强名注册:"
        Warn "    $hklmAsm"
        Warn "我们 HKCU 优先级理论上 win, 但 .NET CLR Fusion 在 PKT 不一致时可能 fallback 到 GAC."
        Warn "如果安装后 patcher.log 显示 PKT=$($matches[1]) (而不是 null),"
        Warn "说明 GAC 强名原版仍 dominant, 需进一步 卸 GAC + 移 HKLM 注册."
    }
}

# -----------------------------------------------------------------------------
# Step 3: HKCU ProgId 注册（64 位 + 32 位双视图）
# -----------------------------------------------------------------------------
function Write-ProgId-Registration {
    param([string]$ProgIdRoot)
    New-Item -Path $ProgIdRoot -Force | Out-Null
    Set-ItemProperty -Path $ProgIdRoot -Name '(Default)' -Value $FriendlyName

    $ProgClsid = Join-Path $ProgIdRoot 'CLSID'
    New-Item -Path $ProgClsid -Force | Out-Null
    Set-ItemProperty -Path $ProgClsid -Name '(Default)' -Value $ClsId
}

Info "[3/5] 注册 ProgId: $ProgId (64 位视图)"
Write-ProgId-Registration -ProgIdRoot "HKCU:\Software\Classes\$ProgId"
Ok "  64 位视图 ProgId 注册完成"

Info "[3/5] 注册 ProgId: $ProgId (32 位 Wow6432Node 视图)"
Write-ProgId-Registration -ProgIdRoot "HKCU:\Software\Classes\Wow6432Node\$ProgId"
Ok "  32 位视图 ProgId 注册完成"

# -----------------------------------------------------------------------------
# Step 4: Office Word Addins 注册
# -----------------------------------------------------------------------------
Info "[4/5] 注册 Office Word Add-in: $ProgId"
$WordAddin = "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"
New-Item -Path $WordAddin -Force | Out-Null
Set-ItemProperty -Path $WordAddin -Name 'FriendlyName'    -Value $FriendlyName
Set-ItemProperty -Path $WordAddin -Name 'Description'     -Value '公文助手 - WPS 公文写作辅助插件'
Set-ItemProperty -Path $WordAddin -Name 'LoadBehavior'    -Value 3 -Type DWord
Set-ItemProperty -Path $WordAddin -Name 'CommandLineSafe' -Value 1 -Type DWord
Ok "Office Word Add-in 注册完成"

# -----------------------------------------------------------------------------
# Step 5: WPS AddinsWL 白名单
# -----------------------------------------------------------------------------
Info "[5/5] 注册 WPS Add-in 白名单"
$WpsWL = 'HKCU:\Software\Kingsoft\Office\WPS\AddinsWL'
New-Item -Path $WpsWL -Force | Out-Null
Set-ItemProperty -Path $WpsWL -Name $ProgId -Value ''

# 同时尝试登记 KsoCorpAddin
$WpsCorp = 'HKCU:\Software\Kingsoft\Office\Common\AddinsWL'
New-Item -Path $WpsCorp -Force | Out-Null
Set-ItemProperty -Path $WpsCorp -Name $ProgId -Value '' -ErrorAction SilentlyContinue
Ok "WPS 白名单注册完成"

# -----------------------------------------------------------------------------
# 完成
# -----------------------------------------------------------------------------
Write-Host ''
Ok '========================================================================='
Ok '安装完成! 启动 WPS, 顶部菜单栏会出现「公文助手 1.0.0」选项卡'
Ok '========================================================================='
Write-Host ''
Info "安装目录:   $InstallDir"
Info "DLL 路径:   $DllPath"
Info "CodeBase:   $DllUri"
Info "卸载命令:   PowerShell -ExecutionPolicy Bypass -File uninstall.ps1"
