[Setup]
AppName=ZiTax Automator
AppVersion=1.3
DefaultDirName={autopf}\ZiTaxAutomator
DefaultGroupName=ZiTax Automator
UninstallDisplayIcon={app}\ZiTax_Automator.exe
SetupIconFile=zitax_icon.ico
Compression=lzma2
SolidCompression=yes
OutputDir=Output
OutputBaseFilename=ZiTax_Automator_Setup_v1.3
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64
ArchitecturesAllowed=x64

[Tasks]
Name: "desktopicon"; Description: "Buat shortcut di Desktop"; GroupDescription: "Additional icons:"

[Files]
; Salin semua isi folder hasil build PyInstaller
Source: "dist\ZiTax_Automator\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\ZiTax Automator"; Filename: "{app}\ZiTax_Automator.exe"
Name: "{autodesktop}\ZiTax Automator"; Filename: "{app}\ZiTax_Automator.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\ZiTax_Automator.exe"; Description: "Jalankan ZiTax Automator Sekarang"; Flags: nowait postinstall skipifsilent
