[Setup]
AppName=Preventiv - Auto Servis Aldo
AppVersion=1.0
AppPublisher=Auto Servis Aldo
DefaultDirName={autopf}\PreventivAldo
DefaultGroupName=Preventiv Aldo
OutputDir=C:\Users\enkla\OneDrive\Desktop\PreventivaApp\Output
OutputBaseFilename=Preventiv_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "C:\Users\enkla\OneDrive\Desktop\PreventivaApp\dist\Preventiv.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Preventiv"; Filename: "{app}\Preventiv.exe"
Name: "{commondesktop}\Preventiv"; Filename: "{app}\Preventiv.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Krijo shortcut ne Desktop"

[Run]
Filename: "{app}\Preventiv.exe"; Description: "Hap Preventiv tani"; Flags: nowait postinstall skipifsilent
