@echo off
setlocal enabledelayedexpansion

:: ============================================================
::  Tryb: normalny (instalacja) lub FORCE/UPDATE (aktualizacja)
:: ============================================================
set "FORCE=0"
if /i "%1"=="FORCE"  set "FORCE=1"
if /i "%1"=="UPDATE" set "FORCE=1"

if "!FORCE!"=="1" (
    title iBiznes Bot - Aktualizacja
    color 0E
    echo.
    echo  iBiznes Bot - Aktualizacja
    echo  ==============================
    echo.
) else (
    title iBiznes Bot - Pelna Instalacja
    color 0B
    echo.
    echo  iBiznes Bot - Pelna Instalacja
    echo  ===================================
    echo.
)

:: ============================================================
::  KONFIGURACJA
::  Zmien REPO_ZIP na adres ZIP swojego repozytorium GitHub
::  Format: https://github.com/UZYTKOWNIK/REPO/archive/refs/heads/BRANCH.zip
:: ============================================================
set "REPO_ZIP=https://github.com/SanTobinoOfficial/iBiznesPythonBot/archive/refs/heads/main.zip"
set "REPO_FOLDER=iBiznesPythonBot-main"

:: Plik konfiguracyjny NIE nadpisywany przy aktualizacji
set "PLIK_CONFIG=coords.json"

:: ============================================================
::  [1/4] SPRAWDZ PYTHON
:: ============================================================
echo  [1/4] Sprawdzanie Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    color 0C
    echo.
    echo  BLAD: Python nie znaleziony!
    echo  Pobierz Python ze strony: https://www.python.org/downloads/
    echo  Pamietaj zaznaczyc "Add Python to PATH" podczas instalacji.
    echo.
    if "!FORCE!"=="0" pause
    exit /b 1
)
echo  [OK] Python znaleziony:
python --version
echo.

:: ============================================================
::  [2/4] POBIERZ I ROZPAKUJ REPO (caly ZIP)
:: ============================================================
if "!FORCE!"=="1" (
    echo  [2/4] Pobieranie aktualnej wersji z GitHub...
) else (
    echo  [2/4] Pobieranie plikow programu z GitHub...
)
echo.

:: Usun stary ZIP i folder tymczasowy jesli istnieja
if exist "_repo.zip"  del /f /q "_repo.zip"
if exist "_repo_tmp"  rd /s /q "_repo_tmp"

echo  Pobieranie repozytorium...
powershell -NoProfile -Command ^
    "try { Invoke-WebRequest -Uri '%REPO_ZIP%' -OutFile '_repo.zip' -UseBasicParsing -ErrorAction Stop; Write-Host '  [OK] Repo pobrane.' } catch { Write-Host '  [BLAD] ' $_.Exception.Message; exit 1 }"

if not exist "_repo.zip" (
    color 0C
    echo.
    echo  BLAD: Nie udalo sie pobrac repozytorium.
    echo  Sprawdz polaczenie z internetem lub adres REPO_ZIP.
    if "!FORCE!"=="0" pause
    exit /b 1
)

echo  Rozpakowywanie...
powershell -NoProfile -Command ^
    "Expand-Archive -Path '_repo.zip' -DestinationPath '_repo_tmp' -Force"

if not exist "_repo_tmp\%REPO_FOLDER%" (
    color 0C
    echo.
    echo  BLAD: Nie mozna rozpakowac. Oczekiwano folderu: _repo_tmp\%REPO_FOLDER%
    del /f /q "_repo.zip"
    if "!FORCE!"=="0" pause
    exit /b 1
)

:: Kopia zapasowa coords.json jesli istnieje
if exist "%PLIK_CONFIG%" (
    echo  [Kopia zapasowa] %PLIK_CONFIG% zachowany...
    copy /y "%PLIK_CONFIG%" "%PLIK_CONFIG%.bak" >nul 2>&1
)

:: Skopiuj wszystkie pliki z repo (bez podfolderow .github)
echo  Kopiowanie plikow programu...
for %%F in ("_repo_tmp\%REPO_FOLDER%\*.py" "_repo_tmp\%REPO_FOLDER%\*.ahk" "_repo_tmp\%REPO_FOLDER%\*.html" "_repo_tmp\%REPO_FOLDER%\*.bat" "_repo_tmp\%REPO_FOLDER%\*.txt" "_repo_tmp\%REPO_FOLDER%\*.json") do (
    for %%G in (%%F) do (
        set "FNAME=%%~nxG"
        :: Nie nadpisuj coords.json (ustawienia uzytkownika)
        if /i "!FNAME!"=="coords.json" (
            if not exist "coords.json" (
                copy /y "%%G" "." >nul
                echo    [NOWY] coords.json
            ) else (
                echo    [ZACHOWANY] coords.json
            )
        ) else (
            copy /y "%%G" "." >nul
            echo    [OK] !FNAME!
        )
    )
)

:: Sprzatanie
del /f /q "_repo.zip"
rd /s /q "_repo_tmp"
echo.
echo  [OK] Pliki skopiowane.
echo.

:: ============================================================
::  [3/4] SPRAWDZ / POBIERZ AUTOHOTKEY v2
:: ============================================================
echo  [3/4] Sprawdzanie AutoHotkey v2...

set "AHK_PATH=C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe"
if not exist "!AHK_PATH!" set "AHK_PATH=C:\Program Files (x86)\AutoHotkey\v2\AutoHotkey64.exe"

if not exist "!AHK_PATH!" (
    echo.
    echo  AutoHotkey v2 nie znaleziony. Pobieranie instalatora...
    powershell -NoProfile -Command ^
        "try { Invoke-WebRequest -Uri 'https://www.autohotkey.com/download/ahk-v2.exe' -OutFile 'ahk_setup.exe' -UseBasicParsing -ErrorAction Stop; Write-Host '  [OK] Instalator AHK v2 pobrany.' } catch { Write-Host '  [UWAGA] Nie udalo sie pobrac AutoHotkey v2: ' $_.Exception.Message }"
    if exist "ahk_setup.exe" (
        echo.
        echo  Uruchamianie instalatora AutoHotkey v2...
        echo  Postepuj zgodnie z instrukcjami, nastepnie program wroci tutaj.
        start /wait ahk_setup.exe
        del /f /q ahk_setup.exe 2>nul
        echo  [OK] AutoHotkey v2 zainstalowany.
    ) else (
        echo  [UWAGA] Pobierz AutoHotkey v2 recznie ze: https://www.autohotkey.com/
    )
) else (
    echo  [OK] AutoHotkey v2: !AHK_PATH!
)
echo.

:: ============================================================
::  [4/4] INSTALACJA / AKTUALIZACJA BIBLIOTEK PYTHON
:: ============================================================
echo  [4/4] Instalacja bibliotek Python...
echo.

python -m pip install --upgrade pip --quiet
python -m pip install flask flask-cors requests pandas pywinauto Pillow pdfplumber openpyxl xlwt

if %errorlevel% neq 0 (
    color 0C
    echo.
    echo  BLAD: Nie udalo sie zainstalowac bibliotek.
    echo  Sprobuj uruchomic jako Administrator (prawy przycisk -> Uruchom jako administrator).
    if "!FORCE!"=="0" pause
    exit /b 1
)

:: Znacznik instalacji
echo. > _installed.flag

if "!FORCE!"=="1" (
    color 0A
    echo.
    echo  ============================
    echo  Aktualizacja zakonczona!
    echo  ============================
    echo.
) else (
    color 0A
    echo.
    echo  =====================================
    echo  Instalacja zakonczona pomyslnie!
    echo  =====================================
    echo.
    echo  Nastepny krok: uruchom START.bat
    echo.
    echo  Jesli bot nie klika poprawnie - skonfiguruj
    echo  koordynaty w: Ustawienia -> Koordynaty iBiznes
    echo.
    pause
)
