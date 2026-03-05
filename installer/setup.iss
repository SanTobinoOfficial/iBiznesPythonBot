; ============================================================================
;  setup.iss  –  Inno Setup 6.1+  –  iBiznes Bot v3.2 Installer
;  Kompilacja: iscc setup.iss  (lub Inno Setup Compiler GUI)
;  Wynik: dist\installer\iBiznesBot-Setup-v3.2.1.exe
;
;  Wymaga: folderu app\ obok tego pliku (wypełniony przez build.bat)
;  Pobierz Inno Setup 6.1+: https://jrsoftware.org/isinfo.php
;
;  UWAGA: NIE wymaga zewnętrznych pluginów (IDP itp.)
;  Pobieranie AHK działa przez wbudowany CreateDownloadPage (Inno Setup 6.1+)
; ============================================================================

#define AppName      "iBiznes Bot"
#define AppVersion   "3.2.1"
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
; Zainstaluj AutoHotkey v2 jeśli został pobrany (plik w {tmp}\ahk-v2.exe)
Filename: "{tmp}\ahk-v2.exe"; \
  Parameters: "/S"; \
  StatusMsg: "Instalowanie AutoHotkey v2..."; \
  Flags: waituntilterminated skipifsilent; \
  Check: FileExists(ExpandConstant('{tmp}\ahk-v2.exe'))

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
// ─────────────────────────────────────────────────────────────────────────────
// Pobieranie AutoHotkey v2 przez wbudowany Inno Setup Download Page
// Nie wymaga zewnętrznych pluginów (IDP itp.)
// Wymaga Inno Setup 6.1+
// ─────────────────────────────────────────────────────────────────────────────

var
  DownloadPage: TDownloadWizardPage;

procedure InitializeWizard();
begin
  DownloadPage := CreateDownloadPage(
    'Pobieranie AutoHotkey v2',
    'Proszę czekać podczas pobierania AutoHotkey v2...',
    nil
  );
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if CurPageID = wpReady then begin
    if not FileExists(ExpandConstant('{#AhkExePath}')) then begin
      DownloadPage.Clear;
      DownloadPage.Add('{#AhkInstUrl}', 'ahk-v2.exe', '');
      DownloadPage.Show;
      try
        try
          DownloadPage.Download;
        except
          if DownloadPage.AbortedByUser then
            MsgBox(
              'Pobieranie AutoHotkey v2 zostało anulowane.' + #13#10 +
              'Zainstaluj ręcznie po instalacji programu:' + #13#10 +
              'https://www.autohotkey.com/',
              mbInformation, MB_OK
            )
          else
            MsgBox(
              'Błąd pobierania AutoHotkey v2:' + #13#10 +
              GetExceptionMessage + #13#10 + #13#10 +
              'Zainstaluj ręcznie po instalacji programu:' + #13#10 +
              'https://www.autohotkey.com/',
              mbError, MB_OK
            );
          // Kontynuuj instalację nawet jeśli AHK się nie pobrało
        end;
      finally
        DownloadPage.Hide;
      end;
    end;
  end;
end;

// ─────────────────────────────────────────────────────────────────────────────
// Informacja o lokalizacji danych użytkownika po instalacji
// ─────────────────────────────────────────────────────────────────────────────

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
