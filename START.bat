@echo off
setlocal enabledelayedexpansion
title iBiznes Bot
color 0A
echo.
echo  iBiznes Bot - Start
echo  ====================
echo.

:: ============================================================
::  KONFIGURACJA – taki sam BASE_URL jak w INSTALL.bat!
::  Zmien na adres swojego repozytorium GitHub.
:: ============================================================
set "BASE_URL=https://raw.githubusercontent.com/TWOJ_NICK/iBiznesPythonBot/main"

:: ============================================================
::  [1] SPRAWDZ INSTALACJE
:: ============================================================
if not exist "_installed.flag" (
    color 0C
    echo  BLAD: Program nie jest zainstalowany.
    echo  Uruchom najpierw INSTALL.bat
    echo.
    pause
    exit /b 1
)

:: ============================================================
::  [2] SPRAWDZ PYTHON
:: ============================================================
python --version >nul 2>&1
if %errorlevel% neq 0 (
    color 0C
    echo  BLAD: Python nie znaleziony.
    echo  Uruchom INSTALL.bat aby ponownie skonfigurowac program.
    pause
    exit /b 1
)

:: ============================================================
::  [3] SPRAWDZ AKTUALIZACJE
:: ============================================================
echo  Sprawdzanie aktualizacji...

del /f /q "_remote_ver.tmp" 2>nul
powershell -NoProfile -Command ^
    "try { Invoke-WebRequest -Uri '%BASE_URL%/version.txt' -OutFile '_remote_ver.tmp' -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop } catch { }" 2>nul

if exist "_remote_ver.tmp" (
    set /p REMOTE_VER=<_remote_ver.tmp
    del /f /q "_remote_ver.tmp" 2>nul

    :: Trim whitespace/newline from REMOTE_VER
    for /f "tokens=* delims= " %%V in ("!REMOTE_VER!") do set "REMOTE_VER=%%V"

    set "LOCAL_VER=brak"
    if exist "version.txt" (
        set /p LOCAL_VER=<version.txt
        for /f "tokens=* delims= " %%V in ("!LOCAL_VER!") do set "LOCAL_VER=%%V"
    )

    if not "!REMOTE_VER!"=="!LOCAL_VER!" (
        color 0E
        echo.
        echo  ==========================================
        echo   DOSTEPNA AKTUALIZACJA  v!REMOTE_VER!
        echo   Twoja wersja: !LOCAL_VER!
        echo  ==========================================
        echo.
        echo  Aktualizowanie plikow programu...
        echo  (coords.json z ustawieniami NIE jest nadpisywany)
        echo.
        call INSTALL.bat FORCE
        color 0A
        echo.
        echo  Aktualizacja zakonczona. Uruchamiam program...
        echo.
    ) else (
        echo  [OK] Program jest aktualny (v!LOCAL_VER!)
        echo.
    )
) else (
    echo  [Brak polaczenia z internenem - pomijam sprawdzanie aktualizacji]
    echo.
)

:: ============================================================
::  [4] SPRAWDZ PLIKI
:: ============================================================
if not exist "server.py" (
    color 0C
    echo  BLAD: Brak pliku server.py
    echo  Uruchom INSTALL.bat aby ponownie zainstalowac program.
    pause
    exit /b 1
)

:: ============================================================
::  [5] URUCHOM SERWER FLASK
:: ============================================================
echo  Uruchamiam serwer Flask...
start "iBiznes Bot Serwer" cmd /k "color 0A && echo. && echo  Serwer: http://localhost:5000 && echo  Nie zamykaj tego okna! && echo. && python server.py"
timeout /t 3 /nobreak >nul

:: ============================================================
::  [6] OTWORZ PANEL W PRZEGLADARCE
:: ============================================================
if exist "_firstrun.flag" (
    echo  Pierwsze uruchomienie - wlaczam tryb diagnostyczny...
    del "_firstrun.flag" >nul 2>&1
    start "" "http://localhost:5000?diag=1"
    echo.
    echo  TRYB DIAGNOSTYCZNY wlaczony.
    echo  Otwórz iBiznes z widocznym oknem Zakup, nastepnie uruchom bota.
) else (
    start "" "http://localhost:5000"
    echo  Panel: http://localhost:5000
)

echo.
echo  Nie zamykaj okna "iBiznes Bot Serwer"!
echo.
timeout /t 5 /nobreak >nul
exit
