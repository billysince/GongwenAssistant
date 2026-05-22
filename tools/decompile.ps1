$ErrorActionPreference = 'Stop'

$ilspyDir = 'D:\工作\20260520\GongwenAssistant\tools\ilspy72'
$inputDll = 'D:\工作\20260520\公文助手Wps插件单机版2.4.1\免安装版本\公文助手Wps插件单机版\Local_Wps_Vsto.dll'
$searchDir = 'D:\工作\20260520\公文助手Wps插件单机版2.4.1\免安装版本\公文助手Wps插件单机版'
$outDir   = 'D:\工作\20260520\GongwenAssistant\src\Local_Wps_Vsto_v2'

if (Test-Path $outDir) { Remove-Item $outDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$global:ILSPY_DIR = $ilspyDir
$resolveHandler = [ResolveEventHandler]{
  param($s, $e)
  $name = ($e.Name -split ',')[0].Trim()
  $candidate = Join-Path $global:ILSPY_DIR ($name + '.dll')
  if (Test-Path $candidate) { return [Reflection.Assembly]::LoadFrom($candidate) }
  return $null
}
[AppDomain]::CurrentDomain.add_AssemblyResolve($resolveHandler)

$asmCore = [Reflection.Assembly]::LoadFrom((Join-Path $ilspyDir 'ICSharpCode.Decompiler.dll'))

$tPEFile   = $asmCore.GetType('ICSharpCode.Decompiler.Metadata.PEFile')
$tUAR      = $asmCore.GetType('ICSharpCode.Decompiler.Metadata.UniversalAssemblyResolver')
$tWPD      = $asmCore.GetType('ICSharpCode.Decompiler.CSharp.ProjectDecompiler.WholeProjectDecompiler')
$tSettings = $asmCore.GetType('ICSharpCode.Decompiler.DecompilerSettings')
$tLangVer  = $asmCore.GetType('ICSharpCode.Decompiler.CSharp.LanguageVersion')

$peStreamOptions = [Enum]::Parse([Reflection.PortableExecutable.PEStreamOptions], 'PrefetchEntireImage')
$mdOptions = [Enum]::Parse([Reflection.Metadata.MetadataReaderOptions], 'Default')
$pefile = $tPEFile.GetConstructor([Type[]]@([string],[Reflection.PortableExecutable.PEStreamOptions],[Reflection.Metadata.MetadataReaderOptions])).Invoke(@($inputDll, $peStreamOptions, $mdOptions))

$uar = $tUAR.GetConstructors()[0].Invoke(@($inputDll, $false, '.NETFramework,Version=v4.5', $null, $peStreamOptions, $mdOptions))
$addSearchDir = $tUAR.GetMethod('AddSearchDirectory')
$addSearchDir.Invoke($uar, @($searchDir))
$addSearchDir.Invoke($uar, @($ilspyDir))

# DecompilerSettings(CSharp5) -> 限制反编译为 C# 5 语言特性
$langVerCSharp5 = [Enum]::Parse($tLangVer, 'CSharp5')
$settings = $tSettings.GetConstructor([Type[]]@($tLangVer)).Invoke(@($langVerCSharp5))
Write-Host ('Settings LangVer ctor used: CSharp5')

# 确保 PIA 类型不要嵌入（这样 EmbedInteropTypes 不会生成 dnSpy 那种烂代码）
# 试着关掉某些 propery
foreach ($p in 'AggressiveScalarReplacementOfAggregates','LoadInMemory') {
  try {
    $prop = $tSettings.GetProperty($p)
    if ($prop -and $prop.CanWrite) { $prop.SetValue($settings, $false) }
  } catch {}
}

$decompiler = $tWPD.GetConstructors() | Where-Object { $_.GetParameters().Count -eq 4 } | Select-Object -First 1
$wpd = $decompiler.Invoke(@($settings, $uar, $null, $null))

$decompileMethod = $tWPD.GetMethods() | Where-Object { $_.Name -eq 'DecompileProject' -and $_.GetParameters().Count -eq 3 } | Select-Object -First 1
$ct = [Threading.CancellationToken]::None
Write-Host 'Decompiling (CSharp5)...'
$start = Get-Date
$decompileMethod.Invoke($wpd, @($pefile, $outDir, $ct))
$elapsed = ((Get-Date) - $start).TotalSeconds
Write-Host ('Done ' + $elapsed + 's, files=' + (Get-ChildItem $outDir -Recurse -File | Measure-Object).Count)
