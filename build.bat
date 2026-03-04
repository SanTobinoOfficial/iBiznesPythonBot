@echo off
setlocal
cd /d "%~dp0"
title iBiznes Bot - Budowanie v3.0

echo.
echo  iBiznes Bot v3.0 - Budowanie .exe
echo  ====================================
echo.

:: Sprawdz czy Python jest dostepny
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  BLAD: Python nie jest zainstalowany lub nie jest w PATH.
    echo  Pobierz Python 3.9+ ze strony: https://www.python.org/
    echo  Upewnij sie ze zaznaczyles "Add Python to PATH" podczas instalacji.
    pause & exit /b 1
)

python --version
echo.

:: Sprawdz czy wymagane pliki istnieja
if not exist "main.py" (
    echo  BLAD: Nie znaleziono main.py – uruchom z folderu projektu.
    pause & exit /b 1
)
if not exist "coords.json" (
    echo  BLAD: Nie znaleziono coords.json – wymagany do bundlowania.
    pause & exit /b 1
)

:: ── [1] Zainstaluj zaleznosci ──────────────────────────────────────────────
echo  [1/3] Instalacja zaleznosci Python...
echo  (flaskwebgui zamiast pywebview – nie wymaga .NET ani pythonnet)
echo.
python -m pip install --upgrade pip --quiet
python -m pip install ^
    pyinstaller ^
    flaskwebgui ^
    flask ^
    flask-cors ^
    requests ^
    pandas ^
    numpy ^
    pdfplumber ^
    openpyxl ^
    xlwt ^
    pywin32

if %errorlevel% neq 0 (
    echo  BLAD: Instalacja zaleznosci nie powiodla sie.
    pause & exit /b 1
)
echo.

:: ── [2] Buduj .exe przez PyInstaller ───────────────────────────────────────
echo  [2/3] Budowanie iBiznesBot.exe (PyInstaller)...
echo  To moze zajac kilka minut...
python -m PyInstaller iBiznesBot.spec --clean --noconfirm

if %errorlevel% neq 0 (
    echo  BLAD: PyInstaller zakonczyl sie bledem.
    pause & exit /b 1
)
echo.

:: ── [3] Skopiuj do folderu instalatora ─────────────────────────────────────
echo  [3/3] Kopiowanie do installer\app\ ...
if not exist "installer" mkdir "installer"
if not exist "installer\app" mkdir "installer\app"
xcopy /s /y /q "dist\iBiznesBot\*" "installer\app\" >nul

echo.
echo  ============================================
echo  Gotowe! dist\iBiznesBot\iBiznesBot.exe
echo  ============================================
echo.
echo  Nastepny krok – zbuduj instalator:
echo.
echo    Opcja A (linia polecen):
echo      iscc installer\setup.iss
echo.
echo    Opcja B (GUI):
echo      Otworz installer\setup.iss w Inno Setup Compiler
echo      i kliknij Build ^> Compile
echo.
echo  Wynik: dist\installer\iBiznesBot-Setup-v3.0.0.exe
echo.
pause
