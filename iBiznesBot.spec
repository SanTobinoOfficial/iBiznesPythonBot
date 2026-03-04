# -*- mode: python ; coding: utf-8 -*-
# iBiznesBot.spec – PyInstaller spec dla iBiznes Bot v3.0
# Budowanie: pyinstaller iBiznesBot.spec --clean --noconfirm

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
        # Requests / networking
        'requests',
        'urllib3',
        'charset_normalizer',
        # PDF parsing
        'pdfplumber',
        'pdfminer',
        'pdfminer.high_level',
        'pdfminer.layout',
        # Data
        'pandas',
        'openpyxl',
        'xlwt',
        # PyWebView
        'webview',
        'webview.platforms.winforms',
        'clr_loader',
        # Windows
        'winreg',
        # Misc
        'Pillow',
        'PIL',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'scipy',
        'numpy',
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
