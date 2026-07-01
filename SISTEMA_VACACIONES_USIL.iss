#define MyAppName "Sistema de Vacaciones USIL"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "People Analytics USIL"
#define MyAppExeName "Sistema Vacaciones USIL.exe"

[Setup]
AppId={{8A8F9B18-9246-4A90-B99C-4F9C0CCB7A25}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Sistema de Vacaciones USIL
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=dist
OutputBaseFilename=Instalador_Sistema_Vacaciones_USIL
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear acceso directo en el escritorio"; GroupDescription: "Accesos directos:"; Flags: unchecked

[Files]
Source: "dist_electron\win-unpacked\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir Sistema de Vacaciones USIL"; Flags: postinstall runascurrentuser nowait skipifsilent