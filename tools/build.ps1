# ============================================================================
#  build.ps1  ——  从 src 编译并部署到 runtime
# ----------------------------------------------------------------------------
#  零外部依赖：只用 Windows 自带的 .NET Framework 4.x msbuild。
#  不需要 Visual Studio / dotnet SDK / Nuget。
#
#  用法：
#      pwsh -File tools\build.ps1
#      pwsh -File tools\build.ps1 -Configuration Debug
#      pwsh -File tools\build.ps1 -SkipDeploy   # 只编译，不拷贝到 runtime
#
#  退出码：
#      0 = 成功（且部署完成）
#      1 = 找不到 msbuild
#      2 = 编译失败
#      3 = 部署失败
# ============================================================================

param(
    [ValidateSet('Release','Debug')]
    [string]$Configuration = 'Release',
    [switch]$SkipDeploy
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$repoRoot = Split-Path -Parent $PSScriptRoot
$csproj   = Join-Path $repoRoot 'src\Local_Wps_Vsto_v2\Local_Wps_Vsto.csproj'
$binDir   = Join-Path $repoRoot ('src\Local_Wps_Vsto_v2\bin\' + $Configuration)
$rtDir    = Join-Path $repoRoot 'runtime'

Write-Host '========== GongwenAssistant Build =========='
Write-Host "  csproj    : $csproj"
Write-Host "  config    : $Configuration"
Write-Host "  outBinDir : $binDir"
Write-Host "  runtime   : $rtDir"

# 1. 定位 msbuild — 优先 .NET Framework 64-bit
$candidates = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe"
)
$msbuild = $null
foreach ($p in $candidates) {
    if (Test-Path $p) { $msbuild = $p; break }
}
if (-not $msbuild) {
    Write-Host '[FATAL] msbuild not found in .NET Framework 4.x directory.' -ForegroundColor Red
    exit 1
}
Write-Host "  msbuild   : $msbuild"

# 2. 编译
Write-Host ''
Write-Host '----- Compiling -----'
$args = @(
    $csproj,
    '/p:Configuration=' + $Configuration,
    '/p:Platform=AnyCPU',
    '/nologo',
    '/v:minimal'
)
& $msbuild @args
if ($LASTEXITCODE -ne 0) {
    Write-Host "[FATAL] msbuild failed with exit code $LASTEXITCODE" -ForegroundColor Red
    exit 2
}

$outDll = Join-Path $binDir 'Local_Wps_Vsto.dll'
if (-not (Test-Path $outDll)) {
    Write-Host "[FATAL] expected output not found: $outDll" -ForegroundColor Red
    exit 2
}

$outSize = (Get-Item $outDll).Length
Write-Host ''
Write-Host ("  Output    : $outDll ({0} bytes)" -f $outSize)

# 3. 部署到 runtime
if ($SkipDeploy) {
    Write-Host '  SkipDeploy=true, runtime not updated.'
    Write-Host ''
    Write-Host '[OK] build done.'
    exit 0
}

Write-Host ''
Write-Host '----- Deploy to runtime -----'
try {
    Copy-Item -Force $outDll                       (Join-Path $rtDir 'Local_Wps_Vsto.dll')
    $pdb = Join-Path $binDir 'Local_Wps_Vsto.pdb'
    if (Test-Path $pdb) { Copy-Item -Force $pdb (Join-Path $rtDir 'Local_Wps_Vsto.pdb') }
    $cfg = Join-Path $binDir 'Local_Wps_Vsto.dll.config'
    if (Test-Path $cfg) { Copy-Item -Force $cfg (Join-Path $rtDir 'Local_Wps_Vsto.dll.config') }
    Write-Host '  copied DLL/PDB/.config -> runtime\'
} catch {
    Write-Host "[FATAL] deploy failed: $_" -ForegroundColor Red
    exit 3
}

# 4. 健康检查
Write-Host ''
Write-Host '----- Post-build check -----'
$asm = [Reflection.Assembly]::ReflectionOnlyLoadFrom((Join-Path $rtDir 'Local_Wps_Vsto.dll'))
$nm  = $asm.GetName()
Write-Host ("  Assembly  : {0}" -f $nm.Name)
Write-Host ("  Version   : {0}" -f $nm.Version)
$pkt = $nm.GetPublicKeyToken()
if ($pkt -and $pkt.Length -gt 0) {
    $hex = ($pkt | ForEach-Object { '{0:x2}' -f $_ }) -join ''
    Write-Host ("  PublicKey : {0}  (WARNING: should be empty for weak naming)" -f $hex) -ForegroundColor Yellow
} else {
    Write-Host '  PublicKey : (empty, weak-named, GOOD)'
}

$userUtil = $asm.GetTypes() | Where-Object { $_.Name -eq 'UserUtil' } | Select-Object -First 1
if ($userUtil) {
    $isVip = $userUtil.GetMethod('IsVip', [Reflection.BindingFlags]'Public,Static,NonPublic')
    if ($isVip) {
        $il = $isVip.GetMethodBody().GetILAsByteArray()
        $hex = ($il | ForEach-Object { '{0:X2}' -f $_ }) -join ' '
        Write-Host ("  IsVip IL  : {0}" -f $hex)
        if ($hex -eq '17 2A') {
            Write-Host '  ★ IsVip = return true (patched, OK)'
        } else {
            Write-Host '  ! IsVip IL unexpected, check source' -ForegroundColor Yellow
        }
    }
}

Write-Host ''
Write-Host '[OK] build + deploy + verify done.'
exit 0
