$ErrorActionPreference = 'Stop'
$installer = Get-ChildItem release/*Windows-x64.Setup.exe | Select-Object -First 1
$destination = Join-Path $env:RUNNER_TEMP 'ThesisCraft install test 中文'
$process = Start-Process $installer.FullName -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP- /DIR=`"$destination`"" -Wait -PassThru
if ($process.ExitCode -ne 0) { throw "Installer failed: $($process.ExitCode)" }
$report = Join-Path $env:RUNNER_TEMP 'thesiscraft-installed-check.json'
$process = Start-Process (Join-Path $destination 'ThesisCraft.exe') -ArgumentList "--self-check `"$report`"" -Wait -PassThru
if ($process.ExitCode -ne 0 -or !(Test-Path $report)) { throw 'Installed app self-check failed' }
if (!(Get-Content $report -Raw | ConvertFrom-Json).success) { throw 'Installed app reported a failure' }
$process = Start-Process (Join-Path $destination 'unins000.exe') -ArgumentList '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -Wait -PassThru
if ($process.ExitCode -ne 0) { throw 'Uninstaller failed' }
if (Test-Path (Join-Path $destination 'ThesisCraft.exe')) { throw 'Uninstall left executable behind' }
