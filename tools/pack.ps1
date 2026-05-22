# ============================================================================
#  pack.ps1  ——  打包发布 zip（新机器零依赖一键安装）
# ----------------------------------------------------------------------------
#  从 GongwenAssistant 仓库根目录打包出一个 zip：
#      GongwenAssistant-<version>-win.zip
#  内含：
#      runtime/        编译产物 + 第三方依赖 + 模板 + 配置
#      installer/      install.ps1 / uninstall.ps1 / verify.ps1
#      README.md       使用说明
#      LICENSE         开源协议
#      docs/           完整文档
#
#  不包含：
#      src/            源码（仓库 clone 才看得到）
#      tools/          开发用脚本
#      bin/            构建日志
#
#  用法:
#      pwsh tools\pack.ps1
#      pwsh tools\pack.ps1 -Version 1.0.1   # 自定义版本号
#      pwsh tools\pack.ps1 -OutDir D:\drop\ # 自定义输出目录
# ============================================================================

param(
    [string]$Version = '1.0.1',
    [string]$OutDir  = ''
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $OutDir) { $OutDir = Join-Path $repoRoot 'dist' }
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }

$staging = Join-Path $env:TEMP ("gw_pack_" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $staging | Out-Null

Write-Host '========== GongwenAssistant Pack =========='
Write-Host "  repo      : $repoRoot"
Write-Host "  version   : $Version"
Write-Host "  out       : $OutDir"
Write-Host "  staging   : $staging"
Write-Host ''

# 复制入包内容
$pkgRoot = Join-Path $staging ("GongwenAssistant-" + $Version)
New-Item -ItemType Directory -Force -Path $pkgRoot | Out-Null

# runtime 全量 (路线 B / 自家重编译版资产)
Copy-Item -Recurse -Force (Join-Path $repoRoot 'runtime')   (Join-Path $pkgRoot 'runtime')
# installer 全量 (路线 B 安装链)
Copy-Item -Recurse -Force (Join-Path $repoRoot 'installer') (Join-Path $pkgRoot 'installer')
# docs 全量
Copy-Item -Recurse -Force (Join-Path $repoRoot 'docs')      (Join-Path $pkgRoot 'docs')
# 顶层文件
Copy-Item -Force (Join-Path $repoRoot 'README.md') (Join-Path $pkgRoot 'README.md')
Copy-Item -Force (Join-Path $repoRoot 'LICENSE')   (Join-Path $pkgRoot 'LICENSE')

# 路线 A · Patcher 资产 (运行时 IL 注入)
$patcherDst = Join-Path $pkgRoot 'patcher'
New-Item -ItemType Directory -Force -Path $patcherDst | Out-Null
$patcherSrc = Join-Path $repoRoot 'src\GongwenPatcher'
foreach ($f in @(
    'bin\GongwenPatcher.dll',
    'bin\0Harmony.dll',
    'install_patcher.ps1',
    'uninstall_patcher.ps1'
)) {
    $src = Join-Path $patcherSrc $f
    if (Test-Path $src) {
        Copy-Item -Force $src (Join-Path $patcherDst (Split-Path $f -Leaf))
    } else {
        Write-Warning ("missing patcher asset: " + $src)
    }
}

# 列入包文件统计
$pkgFiles = [System.IO.Directory]::GetFiles($pkgRoot, '*', [System.IO.SearchOption]::AllDirectories)
$pkgSize  = ($pkgFiles | ForEach-Object { (Get-Item $_).Length } | Measure-Object -Sum).Sum
Write-Host ("  files     : {0}" -f $pkgFiles.Count)
Write-Host ("  size raw  : {0:N0} bytes ({1:N2} MB)" -f $pkgSize, ($pkgSize / 1MB))

# 出 zip
$zipName = "GongwenAssistant-$Version-win.zip"
$zipPath = Join-Path $OutDir $zipName
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

Write-Host ''
Write-Host '  compressing...'
Compress-Archive -Path (Join-Path $pkgRoot '*') -DestinationPath $zipPath -CompressionLevel Optimal

$zipSize = (Get-Item $zipPath).Length
Write-Host ("  zip out   : {0}" -f $zipPath)
Write-Host ("  zip size  : {0:N0} bytes ({1:N2} MB)" -f $zipSize, ($zipSize / 1MB))

# 清理
Remove-Item -Recurse -Force $staging

# 算 sha256 给 release 校验用
$hash = (Get-FileHash $zipPath -Algorithm SHA256).Hash.ToLower()
Write-Host ("  sha256    : {0}" -f $hash)

# 写 .sha256 sidecar
[System.IO.File]::WriteAllText($zipPath + '.sha256', ($hash + '  ' + $zipName + "`n"), (New-Object System.Text.UTF8Encoding $false))
Write-Host ('  sidecar   : ' + $zipPath + '.sha256')

Write-Host ''
Write-Host '[OK] pack done.'
