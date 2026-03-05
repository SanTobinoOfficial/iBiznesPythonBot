# iBiznes Bot v3.0 – Panel automatyzacji faktur zakupowych

Zautomatyzowany panel do wprowadzania faktur zakupowych w programie **iBiznes**.
Odczytuje dane z pliku **PDF lub CSV** i za pomocą **AutoHotkey v2** klika w odpowiednie
elementy iBiznes, wpisując kody produktów i ilości. Posiada też **TRYB BEZPIECZNY** –
konwersję PDF/CSV do pliku Excel 2003 (.xls) gotowego do ręcznego importu.

Od **v3.0** program dystrybuowany jest jako **plik wykonywalny .exe** z własnym
oknem aplikacji (HTML UI wbudowany w Edge/Chrome, bez osobnej przeglądarki).

---

## Spis treści

1. [Wymagania](#wymagania)
2. [Instalacja](#instalacja)
3. [Jak uruchomić](#jak-uruchomić)
4. [Jak działa program](#jak-działa-program)
   - [Tryb Normalny (AHK)](#tryb-normalny-ahk)
   - [TRYB BEZPIECZNY → XLS](#tryb-bezpieczny--xls)
5. [Konfiguracja koordynatów](#konfiguracja-koordynatów)
6. [Dane użytkownika](#dane-użytkownika)
7. [Struktura plików repo](#struktura-plików-repo)
8. [Budowanie .exe (dla deweloperów)](#budowanie-exe-dla-deweloperów)
9. [Rozwiązywanie problemów](#rozwiązywanie-problemów)
10. [FAQ](#faq)
11. [Changelog](#changelog)

---

## Wymagania

| Oprogramowanie | Minimalna wersja | Uwagi |
|---|---|---|
| System | Windows 10 / 11 | — |
| Microsoft Edge | dowolna | Pre-instalowany na Win10/11 |
| AutoHotkey | v2.0 | Instalowany automatycznie przez instalator |
| iBiznes | dowolna | — |
| Python | — | **Nie wymagany** – zawarty w .exe |

> Microsoft Edge jest pre-instalowany na Windows 10 i 11.
> Program otwiera się jako okno Edge w trybie aplikacji (brak paska adresu).

---

## Instalacja

### Opcja A – Instalator .exe (zalecane)

1. Pobierz **`iBiznesBot-Setup-v3.1.0.exe`** z [Releases](https://github.com/SanTobinoOfficial/iBiznesPythonBot/releases)
2. Uruchom instalator jako **Administrator** (prawy przycisk → Uruchom jako administrator)
3. Postępuj zgodnie z kreatorem instalacji
4. Program instaluje się do `C:\Program Files\iBiznes Bot\`
5. Skrót **iBiznes Bot** pojawi się na pulpicie i w menu Start
6. Instalator automatycznie zainstaluje **AutoHotkey v2** jeśli nie ma

### Opcja B – Pojedynczy .exe (portable)

Pobierz `iBiznesBot.exe` z [Releases](https://github.com/SanTobinoOfficial/iBiznesPythonBot/releases),
uruchom bezpośrednio. Żadnej instalacji, żadnych uprawnień administratora.

### Po instalacji

Kliknij dwukrotnie skrót **iBiznes Bot** na pulpicie lub uruchom `iBiznesBot.exe`.

> Przy pierwszym uruchomieniu program może potrzebować kilku sekund na inicjalizację.

---

## Jak uruchomić

Kliknij dwukrotnie **skrót iBiznes Bot** na pulpicie lub `iBiznesBot.exe`.

Otworzy się okno aplikacji z wbudowanym panelem UI (Edge w trybie app).

> Jeśli okno się nie otworzy – sprawdź Windows Defender / antywirus (patrz FAQ).

---

## Jak działa program

### Tryb Normalny (AHK)

Bot automatycznie wprowadza fakturę zakupową do iBiznes klikając w odpowiednie piksele ekranu.

**Przepływ (8 kroków):**
1. Otwiera iBiznes (jeśli nie jest uruchomiony – ścieżka z Ustawień)
2. Klika **"Zakup (...)"** – otwiera moduł zakupów
3. Klika **"Nowy dokument"** – tworzy nową fakturę w lewym panelu
4. Klika **pole dostawcy** → wpisuje nazwę → Enter (iBiznes ładuje dane dostawcy)
5. Klika **zakładkę "Pozycje"**
6. Klika **F7** – otwiera okno "Dodaj z Kartoteki"
7. **Dla każdej pozycji z CSV/PDF:** F3 → kod produktu → Enter → ilość → Enter
8. **Ctrl+S** – zapisuje dokument

**Jak użyć:**
1. Wgraj plik **PDF faktury** (auto-wypełnia pola formularza) lub **CSV**
2. Uzupełnij: NIP dostawcy, Nazwa dostawcy, Numer faktury, Data
3. Wybierz walutę (kurs pobierany automatycznie z NBP)
4. Kliknij **"▶ Uruchom bota"**

> **Uwaga:** iBiznes musi być otwarty i widoczny na ekranie. Nie używaj komputera podczas
> działania bota – przejmuje on sterowanie myszą i klawiaturą.

---

### TRYB BEZPIECZNY → XLS

Konwertuje PDF lub CSV do pliku **Excel 2003 (.xls)** w formacie importu iBiznes –
bez uruchamiania AHK i bez automatycznych kliknięć.

**Jak użyć:**
1. Kliknij **"🔒 TRYB BEZPIECZNY (→ XLS)"**
2. Wgraj plik PDF lub CSV
3. Wybierz walutę faktury
4. Kliknij **"🔄 Konwertuj do XLS"**
5. Pobierz wygenerowany plik `.xls`
6. Zaimportuj ręcznie do iBiznes: **Dokumenty → Import z pliku EXCEL'a**

---

## Konfiguracja koordynatów

Bot klika na **bezwzględnych współrzędnych ekranu** (Screen X, Y). Koordynaty
są zapisane w `%APPDATA%\iBiznesBot\coords.json` i edytowalne w UI.

> **Ważne:** Jeśli zmienisz rozdzielczość, przesuniesz okno iBiznes lub podłączysz
> inny monitor – **zaktualizuj koordynaty** w panelu.

### Jak znaleźć koordynaty (WindowSpy)

1. Upewnij się że **AutoHotkey v2** jest zainstalowany
2. Uruchom **iBiznes** i ustaw okno w normalnej pozycji
3. Kliknij prawym na ikonę AutoHotkey w zasobniku → **"WindowSpy"**
4. Najedź kursorem na element w iBiznes (np. przycisk "Zakup")
5. Odczytaj wartości **"Screen"**: `X: 256  Y: 77`
6. Wpisz w panelu: **⚙ Ustawienia → Koordynaty iBiznes**

---

## Dane użytkownika

Wszystkie dane użytkownika przechowywane są w:

```
%APPDATA%\iBiznesBot\
├── coords.json       ← Twoje koordynaty kliknięć
├── config.json       ← Konfiguracja (ścieżki, domyślne wartości)
├── history.json      ← Historia przetworzonych faktur
├── uploads\          ← Przesłane PDF/CSV i wygenerowane XLS
└── *.log             ← Logi (server.log, ahk.log, pdf_converter.log)
```

> Dane użytkownika **NIE są usuwane** przy odinstalowaniu programu.

---

## Struktura plików repo

```
iBiznesPythonBot/
│
├── main.py            # Entry point – flaskwebgui + Flask (okno Edge app)
├── server.py          # Flask backend API (wszystkie endpointy)
├── pdf_to_csv.py      # Parser PDF faktur + eksporter CSV/XLS
├── ibiznes.ahk        # AutoHotkey v2 – automatyzacja GUI iBiznes
│
├── ui.html            # Panel UI (bundlowany w .exe)
├── coords.json        # Domyślne koordynaty (kopiowane do %APPDATA%)
├── version.txt        # Wersja programu
│
├── iBiznesBot.spec    # PyInstaller spec – budowanie .exe
├── build.bat          # Skrypt budowania (PyInstaller)
├── installer/
│   └── setup.iss      # Inno Setup – budowanie instalatora .exe (opcjonalne)
│
└── .github/
    └── workflows/
        └── ci.yml     # GitHub Actions – syntax check + auto-merge
```

---

## Budowanie .exe (dla deweloperów)

### Wymagania

- Python 3.9+ (testowane na 3.14)
- [Inno Setup 6.1+](https://jrsoftware.org/isinfo.php) – **tylko** jeśli chcesz budować instalator

### Krok 1 – Zbuduj iBiznesBot.exe

```bat
build.bat
```

`build.bat` automatycznie instaluje wszystkie zależności (`flask`, `flaskwebgui`, `pdfplumber`,
`pandas`, `pyinstaller` itp.) i uruchamia PyInstaller.

**Wynik:** `dist\iBiznesBot\iBiznesBot.exe` (folder z .exe – gotowy do użycia)

### Krok 2 – Zbuduj instalator (opcjonalnie)

```bat
iscc installer\setup.iss
```

Lub otwórz `installer/setup.iss` w **Inno Setup Compiler** GUI → Build → Compile.

**Wynik:** `dist\installer\iBiznesBot-Setup-v3.1.0.exe`

> **Bez Inno Setup:** Możesz rozdystrybuować folder `dist\iBiznesBot\` lub sam plik
> `dist\iBiznesBot\iBiznesBot.exe` (portable, nie wymaga instalacji).

---

## Rozwiązywanie problemów

### Okno programu się nie otwiera

**Przyczyna:** Microsoft Edge nie jest zainstalowany (bardzo rzadkie na Win10/11).

**Rozwiązanie:** Zainstaluj [Microsoft Edge](https://www.microsoft.com/edge).
Na Windows 10/11 Edge jest pre-instalowany – sprawdź czy nie został ręcznie usunięty.

---

### Bot klika w złe miejsca

**Przyczyna:** Niepoprawne koordynaty w `coords.json`.

**Rozwiązanie:**
1. Uruchom iBiznes i ustaw okno w normalnej pozycji
2. Użyj **WindowSpy** (AutoHotkey)
3. Zaktualizuj: **⚙ Ustawienia → Koordynaty → Zapisz**

---

### "AHK nie znaleziony"

**Rozwiązanie:**
1. Zainstaluj **AutoHotkey v2**: https://www.autohotkey.com/
2. Panel: **⚙ Ustawienia → Ścieżka do AutoHotkey64.exe → 🔍 Wykryj**

---

### "iBiznes nie znaleziony"

**Rozwiązanie:**
1. Panel: **⚙ Ustawienia → Ścieżka do iBiznes.exe → 🔍 Wykryj**
2. Lub wpisz ręcznie: `C:\Program Files\iBiznes\iBiznes.exe`

---

### Błąd parsowania PDF

Program obsługuje faktury **LEVIOR** i **FESTA**. Inne formaty mogą wymagać dostosowania
kolumn w `pdf_to_csv.py` (stałe `COL_*`).

**Rozwiązanie:** Użyj pliku **CSV** zamiast PDF (przycisk "⬇ Przykładowy CSV" w panelu).

---

### Antywirus blokuje iBiznesBot.exe

PyInstaller bundluje Python interpreter + biblioteki w jeden .exe, co może wywołać
fałszywy alarm (false positive). Dodaj `iBiznesBot.exe` do wyjątków antywirusa.

---

## FAQ

**P: Czy program wymaga Pythona?**
O: Nie. Od v3.0 Python jest zawarty w pliku .exe (bundlowany przez PyInstaller).

**P: Czy program wymaga WebView2 lub .NET?**
O: Nie. Od v3.0.0 używamy flaskwebgui (Edge w trybie app) – brak zależności od .NET lub WebView2 Runtime.

**P: Czy dane z v2.x zostaną zachowane?**
O: Tak – jeśli masz skonfigurowane `coords.json` i `config.json`, skopiuj je do
`%APPDATA%\iBiznesBot\` po instalacji v3.0.

**P: Skąd pobierać aktualizacje?**
O: Program automatycznie sprawdza nowe wersje (banner w górnej części okna).
Kliknij "Pobierz" aby przejść do strony Releases na GitHubie.

**P: Gdzie są moje koordynaty i konfiguracja?**
O: W `%APPDATA%\iBiznesBot\` (wpisz w pasku Explorer: `%APPDATA%\iBiznesBot`).

**P: Co jeśli mam dwa monitory?**
O: Koordynaty muszą odpowiadać pozycji okna iBiznes na konkretnym monitorze.
WindowSpy wyświetla absolutne Screen X/Y z uwzględnieniem układu wielomonitorowego.

**P: Czy mogę używać komputera podczas działania bota?**
O: Nie – bot przejmuje sterowanie myszą i klawiaturą.

---

## Changelog

### v3.1.0 (2026-03)
- **Auto-wykrywanie aktualizacji** – program sprawdza GitHub Releases przy starcie i co 6 godzin; zielony banner gdy dostępna jest nowsza wersja
- **Poprawne porównanie wersji semver** – endpoint `/api/check-update` używa porównania krotek zamiast string compare
- **Polskie nazwy produktów z bazy MDB** – `pdf_to_csv.py` wczytuje nazwy z bazy iBiznes (Access `.mdb`) przez pyodbc; fallback na nazwę z PDF gdy brak sterownika/bazy
- **Kod produktu = 5 cyfr** – Excel/XLS skraca `10048.01` → `10048`
- **bazaMdbPath** – nowe pole konfiguracji wskazujące na plik `.mdb`
- **build.bat** – dodano `pyodbc` do listy zależności
- **iBiznesBot.spec** – `pyodbc` w `hiddenimports`
- **installer/setup.iss** – wersja 3.1.0

### v3.0.0 (2026-03)
**Pełny rewrite projektu:**
- **Program .exe** – flaskwebgui + Flask bundlowany przez PyInstaller; okno Edge w trybie app (brak przeglądarki, brak Pythona na systemie, działa na Python 3.14+)
- **Instalator .exe** – Inno Setup 6.1+; instalacja do Program Files, skrót pulpit/Start Menu, deinstalator, AutoHotkey v2 bundlowany
- **Dane użytkownika** przeniesione do `%APPDATA%\iBiznesBot\` (coords.json, config.json, logi, uploads)
- **ibiznes.ahk** – ścieżki plików zaktualizowane do APPDATA
- **server.py** – DATA_DIR refactor; nowy endpoint `/api/check-update` (GitHub Releases API)
- **pdf_to_csv.py** – log w APPDATA
- **main.py** – nowy entry point (flaskwebgui + Flask + setup APPDATA)
- **ui.html** – banner aktualizacji, branding v3.0
- **build.bat** – skrypt automatycznego budowania (python -m pip, python -m PyInstaller)
- **installer/setup.iss** – Inno Setup script (wbudowany download AHK, bez zewnętrznych pluginów)
- **CI** – zaktualizowany na branch `v3.0`
- Usunięto: `INSTALL.bat`, `START.bat`

### v2.2.12 (2026-03)
- **START.bat**: naprawiono `Nie można odnaleźć dysku.` – przywrócono `start /wait "" "%~dp0INSTALL.bat"`

### v2.x (2026-03)
- CI GitHub Actions, ZIP installer, parser PDF, TRYB BEZPIECZNY, historia faktur, auto-update

### v2.0 (2026-03)
- Nowy przepływ AHK (absolutne koordynaty pikseli), TRYB BEZPIECZNY → XLS,
  koordynaty edytowalne w UI, auto-update, INSTALL.bat + START.bat

### v1.x
- Podstawowa automatyzacja AHK, parser PDF LEVIOR/FESTA, historia, alerty cenowe, kurs NBP
