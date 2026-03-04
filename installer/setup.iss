; ============================================================================
;  setup.iss  –  Inno Setup 6  –  iBiznes Bot v3.0 Installer
;  Kompilacja: iscc setup.iss  (lub Inno Setup Compiler GUI)
;  Wynik: dist\installer\iBiznesBot-Setup-v3.0.0.exe
;
;  Wymaga: folderu app\ obok tego pliku (wypełniony przez build.bat)
;  Pobierz Inno Setup: https://jrsoftware.org/isinfo.php
; ============================================================================

#define AppName      "iBiznes Bot"
#define AppVersion   "3.0.0"
#define AppPublisher "SanTobinoOfficial"
#define AppURL       "https://github.com/SanTobinoOfficial/iBiznesPythonBot"
#define AppExeName   "iBiznesBot.exe"
#define AhkInstUrl   "https://www.autohotkey.com/download/ahk-v2.exe"
#define AhkExePath   "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe"

[Setup]
AppId={{B7A3F2D1-4E9C-4B28-A6F5-3D1E8C2A7B4F}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}/issues
AppUpdatesURL={#AppURL}/releases
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes
; Wymagane uprawnienia admina (instalacja do Program Files)
PrivilegesRequired=admin
OutputDir=..\dist\installer
OutputBaseFilename=iBiznesBot-Setup-v{#AppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
; Minimalna wersja Windows (10)
MinVersion=10.0
; Nie pokazuj ReadMe – w razie potrzeby odkomentuj:
; InfoAfterFile=..\README.md
DisableProgramGroupPage=yes
UninstallDisplayIcon={app}\{#AppExeName}

[Languages]
Name: "polish"; MessagesFile: "compiler:Languages\Polish.isl"

[Tasks]
Name: "desktopicon"; Description: "Utwórz skrót na pulpicie"; GroupDescription: "Ikony:"; Flags: checkedonce

[Files]
; Całe dist\iBiznesBot\ → Program Files\iBiznes Bot\
Source: "app\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Start Menu
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\Odinstaluj {#AppName}"; Filename: "{uninstallexe}"
; Pulpit (opcjonalny – zaznaczony domyślnie)
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
; Zainstaluj AutoHotkey v2 jeśli nie ma
Filename: "{tmp}\ahk-v2.exe"; \
  Parameters: "/S"; \
  StatusMsg: "Instalowanie AutoHotkey v2..."; \
  Flags: waituntilterminated skipifsilent; \
  Check: not FileExists('{#AhkExePath}')

; Uruchom program po instalacji
Filename: "{app}\{#AppExeName}"; \
  Description: "Uruchom {#AppName}"; \
  Flags: nowait postinstall skipifsilent

[UninstallRun]
; Zatrzymaj bot przed deinstalacją
Filename: "taskkill"; Parameters: "/f /im {#AppExeName}"; Flags: runhidden; RunOnceId: "KillBot"

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Code]
// Pobierz instalator AutoHotkey v2 przed instalacją (jeśli brak)
procedure InitializeWizard();
var
  AhkExists: Boolean;
begin
  AhkExists := FileExists(ExpandConstant('{#AhkExePath}'));
  if not AhkExists then begin
    // Pobierz ahk-v2.exe do folderu temp – zostanie zainstalowany w [Run]
    idpAddFile('{#AhkInstUrl}', ExpandConstant('{tmp}\ahk-v2.exe'));
    idpDownloadAfter(wpReady);
  end;
end;

// Informacja o lokalizacji danych użytkownika
procedure CurStepChanged(CurStep: TSetupStep);
var
  AppDataPath: String;
begin
  if CurStep = ssDone then begin
    AppDataPath := ExpandConstant('{userappdata}\iBiznesBot');
    MsgBox(
      'Instalacja zakończona!' + #13#10 + #13#10 +
      'Dane użytkownika (koordynaty, konfiguracja, historia) są przechowywane w:' + #13#10 +
      AppDataPath + #13#10 + #13#10 +
      'Przy odinstalowaniu te dane NIE są usuwane.',
      mbInformation, MB_OK
    );
  end;
end;
