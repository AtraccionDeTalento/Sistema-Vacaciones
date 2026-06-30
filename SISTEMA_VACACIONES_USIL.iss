#define MyAppName "Sistema de Vacaciones USIL"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "People Analytics USIL"
#define MyAppExeName "ABRIR_SISTEMA_VACACIONES.vbs"

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
Source: "assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "DATAS\*"; DestDir: "{app}\DATAS"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "DATA SENSIBLE\*"; DestDir: "{app}\DATA SENSIBLE"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "static\*"; DestDir: "{app}\static"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "VACACIONES\*"; DestDir: "{app}\VACACIONES"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "requirements.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "servidor.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "index_vacaciones.html"; DestDir: "{app}"; Flags: ignoreversion
Source: "run_sistema.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "INSTALAR_SISTEMA_VACACIONES.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "INSTALAR_SISTEMA_VACACIONES.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "ABRIR_SISTEMA_VACACIONES.vbs"; DestDir: "{app}"; Flags: ignoreversion
Source: "GUIA_INSTALACION.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\INSTALAR_SISTEMA_VACACIONES.bat"; Description: "Instalar dependencias y preparar la aplicacion"; Flags: postinstall runascurrentuser waituntilterminated
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir Sistema de Vacaciones USIL"; Flags: postinstall runascurrentuser nowait skipifsilent