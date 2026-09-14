$ErrorActionPreference = 'Stop'
$classId = '{6B7621A4-AC12-47DE-AF27-F7AEB5A34B7C}'
$progId = 'StudyTang.WordFormatter'
foreach ($view in @([Microsoft.Win32.RegistryView]::Registry32, [Microsoft.Win32.RegistryView]::Registry64)) {
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
    try {
        foreach ($path in @("Software\Classes\CLSID\$classId", "Software\Classes\$progId",
            "Software\Microsoft\Office\Word\Addins\$progId", "Software\Kingsoft\Office\WPS\Addins\$progId",
            "Software\Microsoft\Office\16.0\Word\Addins\$progId")) {
            $base.DeleteSubKeyTree($path, $false)
        }
        $allow = $base.OpenSubKey('Software\Kingsoft\Office\WPS\AddinsWL', $true)
        if ($allow) { $allow.DeleteValue($progId, $false); $allow.Dispose() }
        $settings = $base.OpenSubKey('Software\Study-Tang\ThesisCraft', $true)
        if ($settings) { $settings.DeleteValue('OfficeProbe', $false); $settings.Dispose() }
    } finally { $base.Dispose() }
}
Write-Output '已移除此插件的注册项。保留配置与 DLL，避免影响仍在运行的 Office；重启 Office 后不再加载。'
