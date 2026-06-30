[Setup]
AppName=Finance Collection Manager
AppVersion=1.0
AppPublisher=KATASoftware
AppPublisherURL=https://example.com
DefaultDirName={pf}\Finance Collection Manager
DefaultGroupName=Finance Collection Manager
OutputDir=Output
OutputBaseFilename=FinanceManager_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile=finance.ico

[Files]
Source: "dist\FinanceManager.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Finance Collection Manager"; Filename: "{app}\FinanceManager.exe"
Name: "{commondesktop}\Finance Collection Manager"; Filename: "{app}\FinanceManager.exe"

[Run]
Filename: "{app}\FinanceManager.exe"; Description: "Launch Finance Collection Manager"; Flags: nowait postinstall skipifsilent

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"
