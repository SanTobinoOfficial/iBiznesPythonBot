# -*- mode: python ; coding: utf-8 -*-
# iBiznesBot.spec – PyInstaller spec dla iBiznes Bot v3.1
# Budowanie: python -m PyInstaller iBiznesBot.spec --clean --noconfirm
#
# Używa flaskwebgui (Edge/Chrome w trybie app) zamiast pywebview.
# Nie wymaga pythonnet ani .NET – działa na Pythonie 3.14+.

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        # Bundlowane zasoby – kopiowane do katalogu _MEIPASS
        ('ibiznes.ahk', '.'),     # AHK script → kopiowany do APPDATA przy starcie
        ('ui.html',     '.'),     # Frontend → serwowany przez Flask
        ('coords.json', '.'),     # Domyślne koordynaty → kopiowane do APPDATA jeśli brak
        ('version.txt', '.'),     # Wersja
    ],
    hiddenimports=[
        # Flask + CORS
        'flask',
        'flask_cors',
        'werkzeug',
        'werkzeug.serving',
        'werkzeug.routing',
        'werkzeug.middleware.proxy_fix',
        # Requests / networking
        'requests',
        'urllib3',
        'charset_normalizer',
        'certifi',
        # PDF parsing
        'pdfplumber',
        'pdfminer',
        'pdfminer.high_level',
        'pdfminer.layout',
        'pdfminer.pdfpage',
        'pdfminer.pdfinterp',
        # Data – numpy MUSI być, bo pandas go wymaga
        'pandas',
        'pandas.core.frame',
        'pandas.core.series',
        'numpy',
        'openpyxl',
        'xlwt',
        # Windows (pywin32)
        'winreg',
        'win32api',
        'win32con',
        'win32gui',
        'win32com',
        'pythoncom',
        'pywintypes',
        # flaskwebgui – okno Edge/Chrome w trybie app
        'flaskwebgui',
        # Database
        'pyodbc',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # NIE wykluczaj numpy – pandas go potrzebuje!
        'tkinter',
        'matplotlib',
        'scipy',
        'IPython',
        'jupyter',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='iBiznesBot',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,          # Brak czarnego okna CMD w tle
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='installer/icon.ico',  # Odkomentuj jeśli masz ikonę .ico
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='iBiznesBot',       # → dist\iBiznesBot\
)
