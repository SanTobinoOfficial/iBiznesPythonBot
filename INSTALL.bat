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
::  Zmien BASE_URL na adres raw swojego repozytorium GitHub
::  Przyklad: https://raw.githubusercontent.com/mojNick/iBiznesPythonBot/main
:: ============================================================
set "BASE_URL=https://raw.githubusercontent.com/SanTobinoOfficial/iBiznesPythonBot/main"

:: Pliki projektu (zawsze pobierane/aktualizowane)
set "PLIKI_UPDATE=server.py pdf_to_csv.py ibiznes.ahk ui.html START.bat version.txt"

:: Pliki konfiguracyjne (pobierane TYLKO jesli nie istnieja - nie nadpisujemy ustawien)
set "PLIKI_CONFIG=coords.json"

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
::  [2/4] POBIERZ PLIKI PROJEKTU
:: ============================================================
if "!FORCE!"=="1" (
    echo  [2/4] Aktualizowanie plikow projektu z: %BASE_URL%
) else (
    echo  [2/4] Pobieranie brakujacych plikow z: %BASE_URL%
)
echo.

:: --- Pliki projektu (w trybie FORCE zawsze pobierane, normalnie tylko brakujace) ---
for %%F in (%PLIKI_UPDATE%) do (
    if "!FORCE!"=="1" (
        echo  [UPDATE] %%F ...
        powershell -NoProfile -Command ^
            "try { Invoke-WebRequest -Uri '%BASE_URL%/%%F' -OutFile '%%F' -UseBasicParsing -ErrorAction Stop; Write-Host '  [OK] Zaktualizowano: %%F' } catch { Write-Host '  [UWAGA] Nie udalo sie: %%F - ' $_.Exception.Message }"
    ) else (
        if not exist "%%F" (
            echo  [NOWY] Pobieranie: %%F ...
            powershell -NoProfile -Command ^
                "try { Invoke-WebRequest -Uri '%BASE_URL%/%%F' -OutFile '%%F' -UseBasicParsing -ErrorAction Stop; Write-Host '  [OK] Pobrano: %%F' } catch { Write-Host '  [UWAGA] Nie udalo sie: %%F - ' $_.Exception.Message }"
        ) else (
            echo  [OK - juz istnieje] %%F
        )
    )
)

:: --- Pliki konfiguracyjne (NIGDY nie nadpisujemy - zawieraja ustawienia uzytkownika) ---
for %%F in (%PLIKI_CONFIG%) do (
    if not exist "%%F" (
        echo  [NOWY - konfiguracja] %%F ...
        powershell -NoProfile -Command ^
            "try { Invoke-WebRequest -Uri '%BASE_URL%/%%F' -OutFile '%%F' -UseBasicParsing -ErrorAction Stop; Write-Host '  [OK] Pobrano domyslna konfiguracje: %%F' } catch { Write-Host '  [UWAGA] Nie udalo sie: %%F' }"
    ) else (
        echo  [OK - zachowuje ustawienia] %%F
    )
)
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
        echo  Postepuj zgodnie z instrukcjami, nastepnie program wróci tutaj.
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
