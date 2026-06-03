[Setup]
AppName=Phonodex
AppVersion=1.0
AppPublisher=Your Name
DefaultDirName={autopf}\Phonodex
DefaultGroupName=Phonodex
UninstallDisplayIcon={app}\Phonodex.exe
OutputDir=installer
OutputBaseFilename=Phonodex-Setup-v1
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
SetupIconFile=assets\phonodex.ico
PrivilegesRequired=lowest

[Files]
Source: "dist\Phonodex.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Phonodex"; Filename: "{app}\Phonodex.exe"
Name: "{group}\Uninstall Phonodex"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Phonodex"; Filename: "{app}\Phonodex.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Run]
Filename: "{app}\Phonodex.exe"; Description: "Launch Phonodex"; Flags: nowait postinstall skipifsilent