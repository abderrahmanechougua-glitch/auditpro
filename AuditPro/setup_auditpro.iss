; ═══════════════════════════════════════════════════════════════
;  AuditPro — Script Inno Setup
;  Prérequis : Inno Setup 6+ (https://jrsoftware.org/isinfo.php)
;  Usage     : Ouvrir ce fichier dans Inno Setup → Compiler
;  Résultat  : Output\AuditPro_Setup.exe
; ═══════════════════════════════════════════════════════════════

#define AppName       "AuditPro"
#define AppVersion    "1.0.0"
#define AppPublisher  "FIDAROC GRANT THORNTON"
#define AppExeName    "AuditPro.exe"
#define AppId         "{{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}"

[Setup]
AppId={#AppId}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} v{#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL=https://www.fidaroc-gt.com
AppSupportURL=https://www.fidaroc-gt.com
AppUpdatesURL=https://www.fidaroc-gt.com

; Dossier d'installation par défaut
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes

; Compression maximale
Compression=lzma2/ultra64
SolidCompression=yes
LZMANumBlockThreads=4

; Icône et fichiers
SetupIconFile=resources\AuditPro.ico
UninstallDisplayIcon={app}\{#AppExeName}

; Sortie
OutputDir=Output
OutputBaseFilename=AuditPro_Setup_v{#AppVersion}

; Compatibilité Windows
MinVersion=10.0
ArchitecturesInstallIn64BitMode=x64compatible

; Pas besoin d'admin (installation utilisateur)
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=commandline dialog

; Licence (optionnel — commenter si pas de licence)
; LicenseFile=LICENSE.txt

[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

[Tasks]
Name: "desktopicon"; Description: "Créer un raccourci sur le Bureau"; GroupDescription: "Raccourcis additionnels:"; Flags: unchecked
Name: "quicklaunchicon"; Description: "Créer un raccourci dans la barre des tâches"; GroupDescription: "Raccourcis additionnels:"; OnlyBelowVersion: 6.1; Flags: unchecked

[Files]
; Exécutable principal et toutes ses dépendances (depuis PyInstaller)
Source: "dist\AuditPro\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Menu Démarrer
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\{#AppExeName}"
Name: "{group}\Désinstaller {#AppName}"; Filename: "{uninstallexe}"

; Bureau (optionnel)
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
; Lancer l'app après installation (optionnel)
Filename: "{app}\{#AppExeName}"; Description: "Lancer {#AppName} maintenant"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Nettoyer les données utilisateur à la désinstallation (optionnel)
Type: filesandordirs; Name: "{app}\data"

[Code]
// Vérification que PyInstaller a bien été lancé avant
function InitializeSetup(): Boolean;
var
  ExePath: String;
begin
  ExePath := ExpandConstant('{src}\dist\AuditPro\AuditPro.exe');
  if not FileExists(ExePath) then
  begin
    MsgBox(
      'Le fichier dist\AuditPro\AuditPro.exe est introuvable.' + #13#10 +
      'Veuillez d''abord lancer build.bat pour compiler l''application,' + #13#10 +
      'puis relancer ce programme d''installation.',
      mbError, MB_OK
    );
    Result := False;
  end else
    Result := True;
end;
