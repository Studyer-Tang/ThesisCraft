; User-scope install; stable identity and executable path preserve upgrade behavior.
#ifndef Version
  #error Version is required
#endif
[Setup]
AppId=StudyTang.ThesisCraft
AppName=ThesisCraft
AppVersion={#Version}
AppPublisher=Study-Tang
AppPublisherURL=https://github.com/Studyer-Tang/ThesisCraft
DefaultDirName={localappdata}\Programs\ThesisCraft
DefaultGroupName=ThesisCraft
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\release
OutputBaseFilename=ThesisCraft.v{#Version}.Windows-x64.Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\ThesisCraft.exe
CloseApplications=yes
RestartApplications=no
LicenseFile=..\LICENSE
#ifdef SignedBuild
SignTool=publisher
SignedUninstaller=yes
#endif
[Files]
Source: "..\build\release\windows\dist\ThesisCraft\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
[Icons]
Name: "{group}\ThesisCraft"; Filename: "{app}\ThesisCraft.exe"
[Run]
Filename: "{app}\ThesisCraft.exe"; Description: "Launch ThesisCraft"; Flags: nowait postinstall skipifsilent
