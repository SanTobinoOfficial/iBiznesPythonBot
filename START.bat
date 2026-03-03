@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"
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
set "BASE_URL=https://raw.githubusercontent.com/SanTobinoOfficial/iBiznesPythonBot/main"

:: ============================================================
::  [1] SPRAWDZ INSTALACJE – jesli brak, uruchom automatycznie
:: ============================================================
if not exist "_installed.flag" (
    color 0E
    echo  Program nie jest zainstalowany. Uruchamiam instalacje...
    echo.

    :: Pobierz INSTALL.bat jesli rowniez go brakuje
    if not exist "INSTALL.bat" (
        echo  Pobieranie INSTALL.bat z GitHub...
        powershell -NoProfile -Command ^
            "try { Invoke-WebRequest -Uri '%BASE_URL%/INSTALL.bat' -OutFile 'INSTALL.bat' -UseBasicParsing -ErrorAction Stop; Write-Host '  [OK] Pobrano INSTALL.bat' } catch { Write-Host '  [BLAD] Nie mozna pobrac INSTALL.bat: ' $_.Exception.Message }"
        if not exist "INSTALL.bat" (
            color 0C
            echo.
            echo  BLAD: Nie mozna pobrac INSTALL.bat
            echo  Pobierz recznie ze: https://github.com/SanTobinoOfficial/iBiznesPythonBot
            echo.
            pause
            exit /b 1
        )
    )

    echo.
    :: UWAGA: start /wait zamiast call – INSTALL.bat moze nadpisac START.bat,
    ::         wiec uruchamiamy w NOWYM procesie i po skonczeniu restartujemy sie
    start /wait "" "%~dp0INSTALL.bat"

    :: Sprawdz czy instalacja sie powiodla
    if not exist "_installed.flag" (
        color 0C
        echo.
        echo  BLAD: Instalacja nie powiodla sie.
        echo  Sprawdz bledy powyzej i spróbuj ponownie.
        pause
        exit /b 1
    )
    :: Restart – wczytaj swiezo zainstalowane pliki (w tym ew. nowy START.bat)
    start "" "%~f0" SKIP_UPDATE
    exit /b 0
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
::  [3] SPRAWDZ AKTUALIZACJE  (pomijamy jesli SKIP_UPDATE przekazany)
:: ============================================================
if /i "%1"=="SKIP_UPDATE" (
    echo  [Pomijam sprawdzanie aktualizacji po swiezej instalacji]
    echo.
    goto :RUN_SERVER
)

echo  Sprawdzanie aktualizacji...

del /f /q "_remote_ver.tmp" 2>nul
:: Pobierz wersje zdalnie przez PowerShell – zapisz jako UTF-8 bez BOM, LF
powershell -NoProfile -Command ^
    "try { $v = (Invoke-WebRequest -Uri '%BASE_URL%/version.txt' -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop).Content.Trim(); [IO.File]::WriteAllText('_remote_ver.tmp', $v) } catch { }" 2>nul

if exist "_remote_ver.tmp" (
    :: set /p czyta do LF – PowerShell zapisal bez CR wiec trim jest zbedny, ale zostawiamy
    set /p REMOTE_VER=<_remote_ver.tmp
    del /f /q "_remote_ver.tmp" 2>nul
    for /f "tokens=* delims= " %%V in ("!REMOTE_VER!") do set "REMOTE_VER=%%V"

    set "LOCAL_VER=brak"
    if exist "version.txt" (
        :: Czytaj lokalny version.txt tak samo – przez PowerShell aby uniknac CR
        for /f "usebackq tokens=* delims=" %%V in (`powershell -NoProfile -Command "(Get-Content version.txt -Raw).Trim()"`) do set "LOCAL_VER=%%V"
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
        :: UWAGA: start /wait zamiast call – INSTALL.bat nadpisuje ten plik!
        ::         Po aktualizacji restartujemy START.bat z SKIP_UPDATE
        start /wait "" "%~dp0INSTALL.bat" FORCE
        echo  Aktualizacja zakonczona. Restartuje program...
        start "" "%~f0" SKIP_UPDATE
        exit /b 0
    ) else (
        echo  [OK] Program jest aktualny (v!LOCAL_VER!)
        echo.
    )
) else (
    echo  [Brak polaczenia z internetem - pomijam sprawdzanie aktualizacji]
    echo.
)

:RUN_SERVER

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
start /d "%~dp0" "iBiznes Bot Serwer" cmd /k "color 0A && echo. && echo  iBiznes Bot Serwer && echo  Adres: http://localhost:5000 && echo  Nie zamykaj tego okna! && echo. && python server.py || (color 0C && echo. && echo  BLAD: Serwer nie uruchomil sie! && echo  Sprawdz czy Python i biblioteki sa zainstalowane. && echo  Uruchom INSTALL.bat aby naprawic. && pause)"

:: Czekaj az serwer wystartuje (max 15 sekund)
echo  Czekam na uruchomienie serwera...
set SERVER_READY=0
for /l %%i in (1,1,15) do (
    if "!SERVER_READY!"=="0" (
        powershell -NoProfile -Command "try { $r = Invoke-WebRequest http://localhost:5000 -UseBasicParsing -TimeoutSec 1 -ErrorAction Stop; exit 0 } catch { exit 1 }" >nul 2>&1
        if !errorlevel! equ 0 (
            set SERVER_READY=1
            echo  [OK] Serwer uruchomiony!
        ) else (
            timeout /t 1 /nobreak >nul
        )
    )
)

if "!SERVER_READY!"=="0" (
    color 0C
    echo.
    echo  UWAGA: Serwer nie odpowiada po 15 sekundach.
    echo  Sprawdz okno "iBiznes Bot Serwer" czy sa bledy.
    echo.
    pause
    exit /b 1
)

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
echo  Program uruchomiony! Nie zamykaj okna "iBiznes Bot Serwer".
echo.
timeout /t 3 /nobreak >nul
exit
