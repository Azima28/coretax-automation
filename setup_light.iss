[Setup]
AppName=ZiTax Automator (Light)
AppVersion=1.4
DefaultDirName={autopf}\ZiTaxAutomatorLight
DefaultGroupName=ZiTax Automator
UninstallDisplayIcon={app}\ZiTax_Automator.exe
SetupIconFile=zitax_icon.ico
Compression=lzma2
SolidCompression=yes
OutputDir=Output
OutputBaseFilename=ZiTax_Automator_Setup_Light_v1.6
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64
ArchitecturesAllowed=x64

[Tasks]
Name: "desktopicon"; Description: "Buat shortcut di Desktop"; GroupDescription: "Additional icons:"

[Files]
; Salin semua isi folder hasil build PyInstaller KECUALI browser Chromium bawaan (hemat ~150MB!)
Source: "dist\ZiTax_Automator\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs; Excludes: "_internal\playwright\driver\package\.local-browsers\*"

[Icons]
Name: "{group}\ZiTax Automator Light"; Filename: "{app}\ZiTax_Automator.exe"
Name: "{autodesktop}\ZiTax Automator Light"; Filename: "{app}\ZiTax_Automator.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\ZiTax_Automator.exe"; Description: "Jalankan ZiTax Automator Light Sekarang"; Flags: nowait postinstall skipifsilent
