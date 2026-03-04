@echo off
setlocal
cd /d "%~dp0"
title iBiznes Bot - Budowanie v3.0

echo.
echo  iBiznes Bot v3.0 - Budowanie .exe
echo  ====================================
echo.

:: ── [1] Zainstaluj zaleznosci ──────────────────────────────────────────────
echo  [1/3] Instalacja zaleznosci Python...
pip install --upgrade pip --quiet
pip install pyinstaller pywebview flask flask-cors requests pandas pdfplumber openpyxl xlwt Pillow pywin32

if %errorlevel% neq 0 (
    echo  BLAD: Instalacja zaleznosci nie powiodla sie.
    pause & exit /b 1
)
echo.

:: ── [2] Buduj .exe przez PyInstaller ───────────────────────────────────────
echo  [2/3] Budowanie iBiznesBot.exe (PyInstaller)...
echo  To moze zajac kilka minut...
pyinstaller iBiznesBot.spec --clean --noconfirm

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
echo  Nastepny krok (opcjonalny):
echo  Zainstaluj Inno Setup i uruchom:
echo    iscc installer\setup.iss
echo  lub otworz installer\setup.iss w Inno Setup Compiler
echo.
pause
