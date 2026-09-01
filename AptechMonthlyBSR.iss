; =============================================================================
;  Aptech Monthly BSR (Batch Status Report) - Inno Setup Script
;  Lab Timetable & Faculty Workload Manager
; =============================================================================
;  Prerequisites:
;    1. Build the app first with PyInstaller (recommended onedir mode):
;
;         pyinstaller --noconfirm --windowed --name AptechTimetable ^
;             --add-data "templates;templates" ^
;             --add-data "database;database" ^
;             main.py
;
;       Or use the simpler command from README and rely on this script
;       to copy templates / database / Output folders next to the exe.
;
;    2. Place this .iss file in the project root (same folder as main.py).
;    3. Open this file in Inno Setup Compiler and press Compile (or run
;       iscc AptechMonthlyBSR.iss from the command line).
;
;  Output installer will appear in:  Output\AptechMonthlyBSR_Setup.exe
; =============================================================================

#define MyAppName       "Aptech Monthly BSR"
#define MyAppVersion    "1.0.0"
#define MyAppPublisher  "Aptech Computer Education / GLS"
#define MyAppURL        ""
#define MyAppExeName    "AptechMonthlyBSR.exe"
#define MyAppId         "{{A3F8C2E1-9B4D-4E7A-8F1C-2D6B9A0E5F3C}"

[Setup]
; Unique AppId – never change after first release or upgrades will break
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
; Output
OutputDir=Output
OutputBaseFilename=AptechMonthlyBSR_Setup
; Compression
Compression=lzma2/ultra64
SolidCompression=yes
LZMAUseSeparateProcess=yes
; Appearance
WizardStyle=modern
SetupIconFile=
UninstallDisplayIcon={app}\{#MyAppExeName}
; Privileges – install for all users by default
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog
; Architecture
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; Other
DisableProgramGroupPage=yes
DisableReadyPage=no
DisableFinishedPage=no
ShowLanguageDialog=no
VersionInfoVersion={#MyAppVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName} Installer
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; ---------------------------------------------------------------------------
;  Main application (PyInstaller onedir output)
;  Expected layout after build:
;      dist\AptechMonthlyBSR\AptechMonthlyBSR.exe
;      dist\AptechMonthlyBSR\_internal\...
; ---------------------------------------------------------------------------
Source: "dist\AptechMonthlyBSR\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\AptechMonthlyBSR\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; ---------------------------------------------------------------------------
;  Data folders required by the application at runtime
;  (templates is mandatory; database and Output are created/used by the app)
; ---------------------------------------------------------------------------
; Master Excel template(s)
Source: "templates\*"; DestDir: "{app}\templates"; Flags: ignoreversion recursesubdirs createallsubdirs

; Optional: ship a starter / empty database so first launch is instant.
; Comment the next line out if you prefer the app to create a fresh DB itself.
; Source: "database\monthly_bsr.db"; DestDir: "{app}\database"; Flags: ignoreversion onlyifdoesntexist

; Ensure the Output folder exists (empty is fine)
Source: "Output\*"; DestDir: "{app}\Output"; Flags: ignoreversion recursesubdirs createallsubdirs onlyifdoesntexist skipifsourcedoesntexist

; Optional historical source workbook (used by import_historical on first run)
Source: "data\21-August-2026 GLS Labstatus.xlsx"; DestDir: "{app}\data"; Flags: ignoreversion skipifsourcedoesntExist

[Dirs]
; Guarantee these folders exist even if source was empty / missing
Name: "{app}\database"
Name: "{app}\Output"
Name: "{app}\templates"
Name: "{app}\data"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Launch the application after successful install
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Clean up generated reports and the live database on full uninstall
; (users who want to keep data should back up database\monthly_bsr.db first)
Type: filesandordirs; Name: "{app}\Output"
Type: filesandordirs; Name: "{app}\database"
Type: filesandordirs; Name: "{app}\data"

[Code]
// Optional: warn if an older version is already installed
function InitializeSetup(): Boolean;
begin
  Result := True;
end;
