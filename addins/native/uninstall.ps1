$ErrorActionPreference = 'Stop'
$classId = '{F93F581A-74C6-4728-9406-392B6873CA1C}'
$progId = 'StudyTang.ThesisCraft'
foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64,[Microsoft.Win32.RegistryView]::Registry32)) {
    $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,$view)
    try {
        foreach ($key in @("Software\Classes\CLSID\$classId","Software\Classes\$progId",
            "Software\Microsoft\Office\Word\Addins\$progId","Software\Kingsoft\Office\WPS\Addins\$progId")) {
            $root.DeleteSubKeyTree($key,$false)
        }
        $allow = $root.OpenSubKey('Software\Kingsoft\Office\WPS\AddinsWL',$true)
        if ($allow) { $allow.DeleteValue($progId,$false); $allow.Dispose() }
    } finally { $root.Dispose() }
}
Write-Output '已移除原生入口注册；重启 Word/WPS 后生效。保留排版程序和个人模板。'
