[Setup]
AppName=PDF Tool
AppVersion=1.0
AppPublisher=PDF Tool
DefaultDirName={autopf}\PDF Tool
DefaultGroupName=PDF Tool
UninstallDisplayIcon={app}\PDF Tool.exe
OutputDir=installer_output
OutputBaseFilename=Install PDF Tool
SetupIconFile=app_icon.ico
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
WizardStyle=modern
DisableProgramGroupPage=yes
DisableDirPage=yes
DisableReadyPage=yes

[Files]
Source: "dist\PDF Tool.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "User Guide - PDF Tool.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autodesktop}\PDF Tool"; Filename: "{app}\PDF Tool.exe"; IconFilename: "{app}\PDF Tool.exe"
Name: "{group}\PDF Tool"; Filename: "{app}\PDF Tool.exe"
Name: "{group}\User Guide"; Filename: "{app}\User Guide - PDF Tool.txt"
Name: "{group}\Uninstall PDF Tool"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\PDF Tool.exe"; Description: "Open PDF Tool now"; Flags: nowait postinstall
