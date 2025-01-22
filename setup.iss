[Setup]
AppName=RDW Kenteken Checker
AppVersion=1.0.0
DefaultDirName={pf}\RDW Kenteken Checker
DefaultGroupName=RDW Kenteken Checker
OutputBaseFilename=RDW_Kenteken_Checker_Setup
Compression=lzma
SolidCompression=yes
SetupIconFile=icon.ico

[Files]
Source: "dist\RDW_Kenteken_Checker.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "config.ini"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\RDW Kenteken Checker"; Filename: "{app}\RDW_Kenteken_Checker.exe"
Name: "{commondesktop}\RDW Kenteken Checker"; Filename: "{app}\RDW_Kenteken_Checker.exe"

[Run]
Filename: "{app}\RDW_Kenteken_Checker.exe"; Description: "Start RDW Kenteken Checker"; Flags: nowait postinstall skipifsilent