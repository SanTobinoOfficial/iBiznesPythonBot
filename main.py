"""
================================================================================
  main.py  –  iBiznes Bot v3.2.1  –  Entry point
  Uruchamia Flask w wątku daemon + otwiera natywne okno desktopowe (pywebview).
  pywebview używa WebView2 (Windows 10/11) lub MSHTML – prawdziwe okno Win32,
  nie przeglądarka. Bundlowany przez PyInstaller do iBiznesBot.exe.
================================================================================
"""

import os
import shutil
import sys
import threading
import time


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


def _run_flask(app) -> None:
    """Uruchamia Flask w wątku daemon (nie blokuje pętli zdarzeń pywebview)."""
    app.run(
        host="127.0.0.1",
        port=5000,
        debug=False,
        threaded=True,
        use_reloader=False,
    )


def main() -> None:
    # 1. Przygotuj dane użytkownika
    setup_user_data()

    # 2. Załaduj Flask app
    from server import app, VERSION

    # 3. Uruchom Flask w tle
    flask_thread = threading.Thread(target=_run_flask, args=(app,), daemon=True)
    flask_thread.start()

    # Chwila na start Flask (zwykle <0.5s, 1.5s to zapas)
    time.sleep(1.5)

    # 4. Otwórz natywne okno desktopowe przez pywebview
    #    WebView2 (Windows 10/11) – prawdziwe okno Win32, nie przeglądarka.
    #    Nie ma paska adresu, zakładek ani żadnego chrome'u przeglądarki.
    try:
        import webview
        window = webview.create_window(
            title="iBiznes Bot",
            url="http://127.0.0.1:5000",
            width=1280,
            height=820,
            min_size=(900, 600),
            resizable=True,
        )
        webview.start()
    except Exception as e:
        # Fallback – pywebview niedostępne (dev bez GUI / headless CI)
        print(f"[WARN] pywebview niedostępne ({e}). Otwieram w przeglądarce.")
        import webbrowser
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
