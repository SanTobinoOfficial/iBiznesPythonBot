"""
================================================================================
  main.py  –  iBiznes Bot v3.2.2  –  Entry point
  Uruchamia Flask + otwiera okno aplikacji przez flaskwebgui
  (Edge/Chrome w trybie --app: brak paska adresu, zakładek i menu przeglądarki,
  własna ikona w pasku zadań – wygląda i działa jak natywna aplikacja Windows).
  Bundlowany przez PyInstaller do iBiznesBot.exe.

  Dlaczego flaskwebgui, nie pywebview?
  pywebview wymaga pythonnet (clr) do załadowania backendu WinForms/WebView2.
  Kombinacja pythonnet + PyInstaller powoduje błędy DLL i wymaga konkretnej
  wersji .NET na maszynie użytkownika. flaskwebgui działa na każdym Windows 10/11
  bez żadnych dodatkowych zależności (Edge/Chrome są zawsze dostępne).
================================================================================
"""

import os
import shutil
import sys


def resource_path(rel: str) -> str:
    """Ścieżka do zasobów bundlowanych przez PyInstaller lub katalog dev."""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


# ── KATALOG DANYCH UŻYTKOWNIKA ─────────────────────────────────────────────
DATA_DIR = os.path.join(os.environ.get('APPDATA', '.'), 'iBiznesBot')
os.makedirs(os.path.join(DATA_DIR, 'uploads'), exist_ok=True)


def setup_user_data() -> None:
    """
    Kopiuje bundlowane zasoby do %APPDATA%\\iBiznesBot\\ przy uruchomieniu.
    - ibiznes.ahk   – zawsze aktualizowany (nie jest plikiem użytkownika)
    - coords.json   – kopiowany tylko jeśli NIE istnieje (zachowuje kalibracje)
    """
    ahk_src = resource_path("ibiznes.ahk")
    ahk_dst = os.path.join(DATA_DIR, "ibiznes.ahk")
    if os.path.isfile(ahk_src):
        shutil.copy2(ahk_src, ahk_dst)

    coords_src = resource_path("coords.json")
    coords_dst = os.path.join(DATA_DIR, "coords.json")
    if os.path.isfile(coords_src) and not os.path.isfile(coords_dst):
        shutil.copy2(coords_src, coords_dst)


def main() -> None:
    # 1. Przygotuj dane użytkownika
    setup_user_data()

    # 2. Załaduj Flask app
    from server import app, VERSION

    # 3. Uruchom okno – flaskwebgui otwiera Edge/Chrome w trybie --app
    #    Brak paska adresu, zakładek i menu przeglądarki.
    #    Własna ikona w pasku zadań. Wygląda jak natywna aplikacja.
    try:
        from flaskwebgui import FlaskUI
        ui = FlaskUI(
            app=app,
            server="flask",
            width=1280,
            height=820,
            port=5000,
        )
        ui.run()
    except ImportError:
        # Fallback – flaskwebgui nie zainstalowane, otwórz w przeglądarce
        import threading
        import time
        import webbrowser

        def _start_flask():
            app.run(host="127.0.0.1", port=5000,
                    debug=False, threaded=True, use_reloader=False)

        t = threading.Thread(target=_start_flask, daemon=True)
        t.start()
        time.sleep(1.5)
        webbrowser.open("http://127.0.0.1:5000")
        print(f"iBiznes Bot v{VERSION} – Panel: http://127.0.0.1:5000")
        print("Zamknij terminal aby zatrzymać serwer.")
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass


if __name__ == "__main__":
    main()
