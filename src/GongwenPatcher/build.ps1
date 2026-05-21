$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

$csc = (Get-ChildItem 'C:\Windows\Microsoft.NET\Framework\v4.0*\csc.exe' -ErrorAction SilentlyContinue |
        Sort-Object FullName -Descending | Select-Object -First 1).FullName
if (-not $csc) { throw 'csc.exe not found' }
Write-Host "csc: $csc"

$src = @('AssemblyInfo.cs','Extensibility.cs','Connect.cs')

$cscArgs = @(
    '/nologo',
    '/target:library',
    '/platform:x86',
    '/out:bin\GongwenPatcher.dll',
    '/reference:lib\0Harmony.dll',
    '/optimize+',
    '/debug:pdbonly'
)

& $csc @cscArgs @src
if ($LASTEXITCODE -ne 0) { throw "csc failed exit=$LASTEXITCODE" }

Copy-Item 'lib\0Harmony.dll' 'bin\0Harmony.dll' -Force
$out = Get-Item 'bin\GongwenPatcher.dll'
Write-Host ("OK size={0} mtime={1}" -f $out.Length, $out.LastWriteTime)

$asm = [System.Reflection.Assembly]::ReflectionOnlyLoadFrom($out.FullName)
foreach($t in $asm.GetExportedTypes()){ Write-Host ('  type: ' + $t.FullName) }
