param(
    [Parameter(Mandatory=$true)][string]$PythonPath,
    [ValidateSet('Both','Word','WPS')][string]$Target = 'Both',
    [switch]$Experimental
)
$ErrorActionPreference = 'Stop'
if (-not $Experimental) { throw '旧 .NET DLL 已停止推荐：普通启动可触发 CLR 崩溃。请使用 addins/native/install.ps1。仅隔离开发机研究可传 -Experimental。' }
$projectPath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$pythonExe = (Resolve-Path -LiteralPath $PythonPath).Path
& $pythonExe -c 'import docx, tkinter'
if ($LASTEXITCODE -ne 0) { throw '此 Python 缺少运行依赖，请先安装项目。' }
$dll = & (Join-Path $PSScriptRoot 'build.ps1')
$dll = @($dll)[-1]
$installPath = Join-Path $env:LOCALAPPDATA 'Study-Tang\ThesisCraft\Addin'
New-Item -ItemType Directory -Path $installPath -Force | Out-Null
$installedDll = Join-Path $installPath 'StudyTang.WordFormatter.dll'
Copy-Item -LiteralPath $dll -Destination $installedDll -Force
$classId = '{6B7621A4-AC12-47DE-AF27-F7AEB5A34B7C}'
$progId = 'StudyTang.WordFormatter'
$assemblyName = [Reflection.AssemblyName]::GetAssemblyName($installedDll).FullName
$codeBase = ([Uri]$installedDll).AbsoluteUri
# Register only this add-in under the current user, for either Office bitness.
foreach ($view in @([Microsoft.Win32.RegistryView]::Registry32, [Microsoft.Win32.RegistryView]::Registry64)) {
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
    try {
        $settings = $base.CreateSubKey('Software\Study-Tang\ThesisCraft')
        $settings.SetValue('PythonPath', $pythonExe)
        $settings.SetValue('ProjectPath', $projectPath)
        $settings.Dispose()
        $clsid = $base.CreateSubKey("Software\Classes\CLSID\$classId")
        $clsid.SetValue('', 'ThesisCraft')
        $clsid.Dispose()
        $server = $base.CreateSubKey("Software\Classes\CLSID\$classId\InprocServer32")
        $server.SetValue('', 'mscoree.dll')
        $server.SetValue('ThreadingModel', 'Both')
        $server.SetValue('Class', 'StudyTang.WordFormatter.FormatterAddin')
        $server.SetValue('Assembly', $assemblyName)
        $server.SetValue('RuntimeVersion', 'v4.0.30319')
        $server.SetValue('CodeBase', $codeBase)
        $versionKey = $server.CreateSubKey('3.0.0.0')
        foreach ($name in @('Class','Assembly','RuntimeVersion','CodeBase')) {
            $versionKey.SetValue($name, $server.GetValue($name))
        }
        $versionKey.Dispose()
        $server.Dispose()
        $category = $base.CreateSubKey("Software\Classes\CLSID\$classId\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}")
        $category.Dispose()
        $prog = $base.CreateSubKey("Software\Classes\$progId\CLSID")
        $prog.SetValue('', $classId); $prog.Dispose()
        $clsProg = $base.CreateSubKey("Software\Classes\CLSID\$classId\ProgId")
        $clsProg.SetValue('', $progId); $clsProg.Dispose()
        $addinRoots = @()
        if ($Target -in @('Both','Word')) { $addinRoots += 'Software\Microsoft\Office\Word\Addins' }
        if ($Target -in @('Both','WPS')) {
            $addinRoots += 'Software\Kingsoft\Office\WPS\Addins'
            # WPS consumes compatible Word COM registration with a per-add-in allowlist.
            $addinRoots += 'Software\Microsoft\Office\Word\Addins'
            $allow = $base.CreateSubKey('Software\Kingsoft\Office\WPS\AddinsWL')
            $allow.SetValue($progId, '', [Microsoft.Win32.RegistryValueKind]::String); $allow.Dispose()
        }
        foreach ($location in ($addinRoots | Select-Object -Unique)) {
            $addin = $base.CreateSubKey("$location\$progId")
            $addin.SetValue('FriendlyName', '学研排版')
            $addin.SetValue('Description', 'ThesisCraft 4.0.0 · Study-Tang')
            $addin.SetValue('LoadBehavior', 3, [Microsoft.Win32.RegistryValueKind]::DWord)
            $addin.SetValue('CommandLineSafe', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
            $addin.Dispose()
        }
    } finally { $base.Dispose() }
}
Write-Output '安装完成。保存工作并重新打开 Word / WPS，查找“学研排版”选项卡。'
Write-Output "插件文件：$installedDll"
