param(
    [Parameter(Mandatory=$true)][string]$ExecutablePath,
    [switch]$Python,
    [string]$ProjectPath = '',
    [string]$BinaryDirectory = (Join-Path $PSScriptRoot 'bin'),
    [ValidateSet('Both','Word','WPS')][string]$Target = 'Both'
)
$ErrorActionPreference = 'Stop'
$executable = (Resolve-Path -LiteralPath $ExecutablePath).Path
if ($Python -and -not $ProjectPath) { throw '源码运行必须提供 ProjectPath。' }
if ($ProjectPath) { $ProjectPath = (Resolve-Path -LiteralPath $ProjectPath).Path }
else { $ProjectPath = Split-Path -Parent $executable }
# Validate every binary before changing registration. Keep loaded older DLLs intact.
$binaries = @{}
foreach ($arch in @('x86_64','i686')) {
    $binary = (Resolve-Path -LiteralPath (Join-Path $BinaryDirectory "$arch/ThesisCraft.Addin.dll")).Path
    $bytes = [IO.File]::ReadAllBytes($binary)
    $pe = [BitConverter]::ToInt32($bytes,60)
    $expected = if ($arch -eq 'x86_64') { 0x8664 } else { 0x14c }
    if ([BitConverter]::ToUInt16($bytes,$pe+4) -ne $expected) { throw "DLL 位数不匹配：$arch" }
    $binaries[$arch] = $binary
}
$classId = '{F93F581A-74C6-4728-9406-392B6873CA1C}'
$progId = 'StudyTang.ThesisCraft'
foreach ($arch in @('x86_64','i686')) {
    $binary = $binaries[$arch]
    $hash = (Get-FileHash -LiteralPath $binary -Algorithm SHA256).Hash.Substring(0,16)
    $folder = Join-Path $env:LOCALAPPDATA "Study-Tang/ThesisCraft/NativeAddin/$hash/$arch"
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    $installed = Join-Path $folder 'ThesisCraft.Addin.dll'
    if (-not (Test-Path -LiteralPath $installed)) { Copy-Item -LiteralPath $binary -Destination $installed }
    $view = if ($arch -eq 'x86_64') { [Microsoft.Win32.RegistryView]::Registry64 } else { [Microsoft.Win32.RegistryView]::Registry32 }
    $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,$view)
    try {
        $server = $root.CreateSubKey("Software\Classes\CLSID\$classId\InprocServer32")
        $server.SetValue('',$installed)
        $server.SetValue('ThreadingModel','Apartment'); $server.Dispose()
        $prog = $root.CreateSubKey("Software\Classes\$progId\CLSID")
        $prog.SetValue('',$classId); $prog.Dispose()
        $prog = $root.CreateSubKey("Software\Classes\CLSID\$classId\ProgID")
        $prog.SetValue('',$progId); $prog.Dispose()
        $settings = $root.CreateSubKey('Software\Study-Tang\ThesisCraft')
        $settings.SetValue('NativeExecutable',$executable)
        $settings.SetValue('NativePython',$(if ($Python) { '1' } else { '0' }))
        $settings.SetValue('ProjectPath',$ProjectPath); $settings.Dispose()
        $locations = @('Software\Microsoft\Office\Word\Addins')
        if ($Target -in @('Both','WPS')) {
            $locations += 'Software\Kingsoft\Office\WPS\Addins'
            $allow = $root.CreateSubKey('Software\Kingsoft\Office\WPS\AddinsWL')
            $allow.SetValue($progId,''); $allow.Dispose()
        }
        foreach ($location in $locations) {
            $addin = $root.CreateSubKey("$location\$progId")
            $addin.SetValue('FriendlyName','学研排版')
            $addin.SetValue('Description','ThesisCraft · Study-Tang')
            $addin.SetValue('LoadBehavior',3,[Microsoft.Win32.RegistryValueKind]::DWord)
            $addin.Dispose()
        }
    } finally { $root.Dispose() }
}
Write-Output '原生功能区入口已安装。正常重启 Word/WPS 后查看“学研排版”选项卡。无需开启宏。'
Write-Output 'WPS 使用兼容的 Word 发现键，因此 WPS 安装也会提供 Word 入口。其他版本仍需验证。'
