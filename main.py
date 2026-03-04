"""
================================================================================
  main.py  –  iBiznes Bot v3.0  –  Entry point
  Uruchamia Flask w wątku, otwiera okno PyWebView (Windows WebView2).
  Bundlowany przez PyInstaller do iBiznesBot.exe.
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
    Kopiuje bundlowane zasoby do %APPDATA%\\iBiznesBot\\ przy pierwszym uruchomieniu.
    - ibiznes.ahk   – zawsze aktualizowany (nie jest plikiem użytkownika)
    - coords.json   – kopiowany tylko jeśli NIE istnieje (zachowuje kalibracje)
    """
    # ibiznes.ahk – zawsze aktualizuj z bundle (nowa wersja może mieć poprawki)
    ahk_src = resource_path("ibiznes.ahk")
    ahk_dst = os.path.join(DATA_DIR, "ibiznes.ahk")
    if os.path.isfile(ahk_src):
        shutil.copy2(ahk_src, ahk_dst)

    # coords.json – kopiuj tylko jeśli użytkownik nie ma własnego
    coords_src = resource_path("coords.json")
    coords_dst = os.path.join(DATA_DIR, "coords.json")
    if os.path.isfile(coords_src) and not os.path.isfile(coords_dst):
        shutil.copy2(coords_src, coords_dst)


def start_flask() -> None:
    """Uruchamia Flask server w wątku daemonicznym."""
    from server import app
    app.run(host="127.0.0.1", port=5000, debug=False,
            threaded=True, use_reloader=False)


def main() -> None:
    # 1. Przygotuj dane użytkownika
    setup_user_data()

    # 2. Uruchom Flask w tle
    flask_thread = threading.Thread(target=start_flask, daemon=True)
    flask_thread.start()

    # 3. Poczekaj na start serwera (max 10 sekund)
    import urllib.request
    for _ in range(20):
        try:
            urllib.request.urlopen("http://127.0.0.1:5000/api/ping", timeout=1)
            break
        except Exception:
            time.sleep(0.5)

    # 4. Otwórz okno PyWebView
    try:
        import webview
        webview.create_window(
            title="iBiznes Bot v3.0",
            url="http://127.0.0.1:5000",
            width=1280,
            height=820,
            min_size=(900, 600),
            confirm_close=False,
            text_select=True,
        )
        webview.start(debug=False)
    except ImportError:
        # Fallback – jeśli pywebview nie jest zainstalowany, otwórz w przeglądarce
        import webbrowser
        webbrowser.open("http://127.0.0.1:5000")
        # Utrzymaj proces przy życiu
        print("iBiznes Bot v3.0 – Panel: http://127.0.0.1:5000")
        print("Zamknij to okno aby zatrzymac serwer.")
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass


if __name__ == "__main__":
    main()
