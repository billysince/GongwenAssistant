# =============================================================================
# 公文助手 GongwenAssistant 卸载脚本
# =============================================================================
# 清理由 install.ps1 写入的所有 HKCU 注册表项 + 安装目录
# 不动 HKLM / GAC, 因此完全无需管理员权限
#
# 用法:
#   PowerShell -ExecutionPolicy Bypass -File uninstall.ps1 [-KeepUserData]
# =============================================================================

[CmdletBinding()]
param(
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA 'GongwenAssistant'),
    [switch]$KeepUserData
)

$ErrorActionPreference = 'Continue'
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$ProgId = 'Local_Wps_Vsto.MyAddin'
$ClsId  = '{2CD4E522-63D0-3A6B-9842-587DB1AB7FBF}'

function Info($m)  { Write-Host "[INFO ] $m" -ForegroundColor Cyan }
function Ok($m)    { Write-Host "[ OK  ] $m" -ForegroundColor Green }
function Warn($m)  { Write-Host "[WARN ] $m" -ForegroundColor Yellow }

Info "公文助手 GongwenAssistant 卸载程序启动"

# 关 WPS 防止文件锁定
$wps = Get-Process wps,wpsoffice,wpscloudsvr,wpscloudlaunch -ErrorAction SilentlyContinue
if ($wps) {
    Warn "检测到 WPS 正在运行, 将自动关闭"
    $wps | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

# -----------------------------------------------------------------------------
# Step 1: 移除 HKCU 注册表项
# -----------------------------------------------------------------------------
$keys = @(
    "HKCU:\Software\Classes\CLSID\$ClsId",
    "HKCU:\Software\Classes\$ProgId",
    "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"
)
foreach ($k in $keys) {
    if (Test-Path $k) {
        Remove-Item -Path $k -Recurse -Force -ErrorAction SilentlyContinue
        Ok "移除 $k"
    } else {
        Info "已不存在 $k"
    }
}

# AddinsWL 是值, 不是子键
foreach ($wl in 'HKCU:\Software\Kingsoft\Office\WPS\AddinsWL', 'HKCU:\Software\Kingsoft\Office\Common\AddinsWL') {
    if (Test-Path $wl) {
        Remove-ItemProperty -Path $wl -Name $ProgId -ErrorAction SilentlyContinue
        Ok "移除白名单条目 $wl\$ProgId"
    }
}

# -----------------------------------------------------------------------------
# Step 2: 删除安装目录
# -----------------------------------------------------------------------------
if (Test-Path $InstallDir) {
    if ($KeepUserData) {
        Warn "保留用户数据 (-KeepUserData), 仅删除程序文件"
        Get-ChildItem $InstallDir -Filter '*.dll' | Remove-Item -Force -ErrorAction SilentlyContinue
        Get-ChildItem $InstallDir -Filter '*.tlb' | Remove-Item -Force -ErrorAction SilentlyContinue
        Get-ChildItem $InstallDir -Filter '*.pdb' | Remove-Item -Force -ErrorAction SilentlyContinue
    } else {
        Remove-Item -Path $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
        Ok "移除安装目录 $InstallDir"
    }
}

Write-Host ''
Ok '卸载完成! 启动 WPS 不再加载公文助手插件'
