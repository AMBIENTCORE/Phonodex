[Setup]
; AppId uniquely identifies THIS application to Windows. Never change it
; between versions -- it's what tells Inno Setup "this new installer is
; an upgrade of the existing install" rather than a separate application.
; Changing this GUID after v1.0 ships would leave users with two
; parallel installs and no upgrade path. Generate a fresh GUID only if
; you're forking to create a genuinely different application.
AppId={{A8F3E2D1-6B4C-4A9E-8F7D-3C5B9E2A1F6D}
AppName=Phonodex
AppVersion=1.1
AppPublisher=Your Name
DefaultDirName={autopf}\Phonodex
DefaultGroupName=Phonodex
UninstallDisplayIcon={app}\Phonodex.exe
OutputDir=installer
OutputBaseFilename=Phonodex-Setup-v1.1
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