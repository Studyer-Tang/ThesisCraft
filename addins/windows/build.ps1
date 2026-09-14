param([string]$OutputDirectory = (Join-Path $PSScriptRoot 'dist'))
$ErrorActionPreference = 'Stop'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $compiler)) {
    $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$targetDll = Join-Path $OutputDirectory 'StudyTang.WordFormatter.dll'
& $compiler /nologo /target:library /platform:anycpu /optimize+ "/out:$targetDll" /r:System.Windows.Forms.dll /r:Microsoft.CSharp.dll /r:System.Web.Extensions.dll (Join-Path $PSScriptRoot 'FormatterAddin.cs')
if ($LASTEXITCODE -ne 0) { throw '插件编译失败。' }
Write-Output $targetDll
