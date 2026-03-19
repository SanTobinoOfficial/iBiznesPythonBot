# iBiznes Bot v3.2.2 – Panel automatyzacji faktur zakupowych

Zautomatyzowany panel do wprowadzania faktur zakupowych w programie **iBiznes**.
Odczytuje dane z pliku **PDF lub CSV** i za pomocą **AutoHotkey v2** klika w odpowiednie
elementy iBiznes, wpisując kody produktów i ilości. Posiada też **TRYB BEZPIECZNY** –
konwersję PDF/CSV do pliku Excel 2003 (.xls) gotowego do ręcznego importu.

Od **v3.0** program dystrybuowany jest jako **plik wykonywalny .exe** z własnym
oknem aplikacji (Edge/Chrome w trybie `--app` – brak paska adresu i zakładek,
własna ikona w pasku zadań, wygląda jak natywna aplikacja Windows).

---

## Spis treści

1. [Wymagania](#wymagania)
2. [Instalacja](#instalacja)
3. [Jak uruchomić](#jak-uruchomić)
4. [Wszystkie funkcje programu](#wszystkie-funkcje-programu)
   - [1. Tryb Normalny – automatyczne wprowadzanie (AHK)](#1-tryb-normalny--automatyczne-wprowadzanie-ahk)
   - [2. TRYB BEZPIECZNY – konwersja do XLS](#2-tryb-bezpieczny--konwersja-do-xls)
   - [3. Parser faktur PDF](#3-parser-faktur-pdf)
   - [4. Kurs walut z NBP](#4-kurs-walut-z-nbp)
   - [5. Historia faktur](#5-historia-faktur)
   - [6. Alerty cenowe](#6-alerty-cenowe)
   - [7. Auto-wykrywanie ścieżek](#7-auto-wykrywanie-ścieżek)
   - [8. Auto-aktualizacje](#8-auto-aktualizacje)
   - [9. Polskie nazwy produktów z bazy MDB](#9-polskie-nazwy-produktów-z-bazy-mdb)
   - [10. Konfiguracja koordynatów](#10-konfiguracja-koordynatów)
   - [11. Panel logów](#11-panel-logów)
   - [12. Podgląd i eksport danych](#12-podgląd-i-eksport-danych)
5. [Konfiguracja koordynatów – szczegóły](#konfiguracja-koordynatów--szczegóły)
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

1. Pobierz **`iBiznesBot-Setup-v3.2.2.exe`** z [Releases](https://github.com/SanTobinoOfficial/iBiznesPythonBot/releases)
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

## Wszystkie funkcje programu

### 1. Tryb Normalny – automatyczne wprowadzanie (AHK)

Bot automatycznie wprowadza fakturę zakupową do iBiznes, klikając w odpowiednie piksele ekranu przy pomocy **AutoHotkey v2**.

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

Postęp wyświetlany jest na żywo w panelu (SSE streaming – bez odświeżania strony).

> **Uwaga:** iBiznes musi być otwarty i widoczny na ekranie. Nie używaj komputera podczas
> działania bota – przejmuje on sterowanie myszą i klawiaturą.

---

### 2. TRYB BEZPIECZNY – konwersja do XLS

Konwertuje PDF lub CSV do pliku **Excel 2003 (.xls)** w 24-kolumnowym formacie importu iBiznes – bez uruchamiania AHK i bez automatycznych kliknięć.

Plik `.xls` zawiera wszystkie wymagane kolumny: kod towaru, nazwa, ilość, cena netto PLN, cena brutto PLN, cena dewizowa, VAT, JM, dostawca i inne.

**Jak użyć:**
1. Kliknij zakładkę **"🔒 TRYB BEZPIECZNY"**
2. Wgraj plik PDF lub CSV
3. Wybierz walutę faktury (kurs pobierany z NBP lub wpisz ręcznie)
4. Kliknij **"🔄 Konwertuj do XLS"**
5. Pobierz wygenerowany plik `.xls`
6. Zaimportuj ręcznie do iBiznes: **Dokumenty → Import z pliku EXCEL'a**

---

### 3. Parser faktur PDF

Automatycznie odczytuje dane z pliku PDF faktury zakupowej:

- **Pozycje produktów**: kod (5 cyfr), nazwa, ilość, cena netto w walucie, EAN
- **Nagłówek faktury**: numer faktury, nazwa dostawcy, data, waluta
- **Deduplikacja**: jeśli ten sam kod pojawi się na kilku stronach PDF, ilości są sumowane
- **Obsługiwane formaty**: LEVIOR s.r.o., FESTA Professional Tools (i inne faktury o podobnej strukturze kolumnowej)

Po wgraniu PDF pola formularza (NIP, dostawca, numer faktury, data, waluta) są wypełniane automatycznie.

Generowane pliki wynikowe:
- `.csv` – dane surowe dla bota
- `_ibiznes.xlsx` – podgląd w Excelu (z nagłówkami)
- `_raport.html` – czytelny raport HTML z podsumowaniem wartości faktury

---

### 4. Kurs walut z NBP

Kurs walut (USD, EUR i inne) pobierany jest automatycznie z **API Narodowego Banku Polskiego** (`api.nbp.pl`).

- Kurs odświeżany przy każdym uruchomieniu bota
- Możliwość wpisania kursu ręcznie (pole w formularzu)
- Wyświetlany w nagłówku panelu (badge z aktualnym kursem i datą)
- Fallback: `4.05 PLN` gdy NBP niedostępny
- Obsługiwane waluty: USD, EUR, GBP, CHF, CZK i inne kody obsługiwane przez NBP

---

### 5. Historia faktur

Panel **Historia** przechowuje ostatnie 50 przetworzonych faktur.

Dla każdej faktury zapisane są:
- Numer faktury, NIP dostawcy, data
- Waluta, liczba pozycji
- Status (`running` / `done` / `error`)
- Liczba pozycji dodanych, pominiętych (alert cenowy), błędów
- Czas rozpoczęcia i zakończenia

Historia przechowywana jest w `%APPDATA%\iBiznesBot\history.json` i nie jest usuwana przy aktualizacji programu.

---

### 6. Alerty cenowe

Gdy cena produktu z faktury różni się od ceny w systemie iBiznes o więcej niż **tolerancja** (domyślnie 0.05 PLN), pozycja jest oznaczana jako alert cenowy zamiast dodania do faktury.

- Alerty widoczne w zakładce **"Alerty"** w panelu
- Każdy alert zapisywany jest do pliku `%APPDATA%\iBiznesBot\price_alerts.txt` z timestampem
- Format: `KOD | NAZWA | FAKTURA: X PLN | SYSTEM: Y PLN | RÓŻNICA: Z PLN | KURS: K`
- Tolerancja konfigurowalna w Ustawieniach
- Możliwość wyczyszczenia historii alertów jednym kliknięciem

---

### 7. Auto-wykrywanie ścieżek

Program automatycznie szuka plików wykonywalnych na dysku:

**iBiznes.exe** – sprawdzane lokalizacje:
- `C:\Program Files\iBiznes\`
- `C:\Program Files (x86)\iBiznes\`
- `C:\iBiznes\`, `D:\iBiznes\`
- Rekurencyjne przeszukanie `Program Files` na dyskach C i D
- Rejestr Windows (`HKLM\SOFTWARE\iBiznes`)

**AutoHotkey64.exe** – sprawdzane lokalizacje:
- `C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe`
- `C:\Program Files\AutoHotkey\v2\AutoHotkey.exe`
- `C:\Program Files (x86)\AutoHotkey\v2\`

Przycisk **"🔍 Wykryj"** w Ustawieniach uruchamia wykrywanie ręcznie.

---

### 8. Auto-aktualizacje

Program automatycznie sprawdza dostępność nowej wersji na **GitHub Releases**:

- Pierwsze sprawdzenie: **2 sekundy** po uruchomieniu
- Kolejne sprawdzenia: co **6 godzin** (w tle, bez restartu)
- Porównanie wersji: semver (`3.2.0 > 3.1.1`) – bez błędów przy porównaniu `3.1.9` vs `3.2.0`
- Gdy dostępna nowa wersja: zielony banner na górze okna z przyciskiem **"Pobierz"**
- Kliknięcie otwiera stronę Release na GitHubie

---

### 9. Polskie nazwy produktów z bazy MDB

Jeśli posiadasz bazę danych iBiznes (plik `.mdb`), program może odczytywać **polskie nazwy produktów** zamiast anglojęzycznych nazw z faktury PDF.

- Połączenie przez **pyodbc + Microsoft Access Driver**
- Wyszukiwanie po kodzie 5-cyfrowym (np. `10048`)
- Automatyczny fallback na nazwę z PDF gdy:
  - Brak ścieżki do `.mdb` w konfiguracji
  - Brak sterownika Microsoft Access Database Engine
  - Produkt nie znaleziony w bazie
- Ścieżka do `.mdb` ustawiana w: **⚙ Ustawienia → Ścieżka do bazy MDB**

> Wymaga: [Microsoft Access Database Engine 2016 x64](https://www.microsoft.com/en-us/download/details.aspx?id=54920)

---

### 10. Konfiguracja koordynatów

Bot klika na **bezwzględnych współrzędnych ekranu** (Screen X, Y). Koordynaty
są zapisane w `%APPDATA%\iBiznesBot\coords.json` i edytowalne w UI bez restartowania programu.

Konfigurowane punkty:
| Klucz | Co klika |
|---|---|
| `btnZakup` | Przycisk "Zakup (...)" w menu |
| `btnNewDoc` | Przycisk nowego dokumentu w lewym panelu |
| `supplierField` | Pole nazwy/NIP dostawcy |
| `tabPositions` | Zakładka "Pozycje" |
| `btnF7` | (rezerwowy, używany przez AHK przez Send F7) |

> **Ważne:** Jeśli zmienisz rozdzielczość, przesuniesz okno iBiznes lub podłączysz
> inny monitor – **zaktualizuj koordynaty** w panelu.

---

### 11. Panel logów

Zakładka **"Logi"** wyświetla ostatnie 100 linii z pliku `server.log`.

- Logi z serwera Flask (operacje PDF, starty bota, błędy)
- Odświeżane ręcznie przyciskiem
- Pełne logi dostępne w `%APPDATA%\iBiznesBot\`:
  - `server.log` – backend Python
  - `ahk.log` – skrypt AutoHotkey (każde kliknięcie, każdy krok)
  - `pdf_converter.log` – parsowanie PDF

---

### 12. Podgląd i eksport danych

Po wgraniu pliku PDF lub CSV, dane wyświetlane są w tabeli podglądu przed uruchomieniem bota:

- Kod produktu, nazwa (polska jeśli MDB skonfigurowane), ilość, cena netto w walucie
- Możliwość pobrania wygenerowanego pliku **`.xlsx`** (podgląd z nagłówkami)
- Możliwość pobrania pliku **`.xls`** (24 kolumny, gotowy do importu iBiznes)
- Raport HTML (`_raport.html`) generowany automatycznie przy parsowaniu PDF – zawiera podsumowanie wartości faktury

---

## Konfiguracja koordynatów – szczegóły

### Jak znaleźć koordynaty (WindowSpy)

1. Upewnij się że **AutoHotkey v2** jest zainstalowany
2. Uruchom **iBiznes** i ustaw okno w normalnej pozycji
3. Kliknij prawym na ikonę AutoHotkey w zasobniku → **"WindowSpy"**
4. Najedź kursorem na element w iBiznes (np. przycisk "Zakup")
5. Odczytaj wartości **"Screen"**: `X: 256  Y: 77`
6. Wpisz w panelu: **⚙ Ustawienia → Koordynaty iBiznes → Zapisz**

---

## Dane użytkownika

Wszystkie dane użytkownika przechowywane są w:

```
%APPDATA%\iBiznesBot\
├── coords.json         ← Twoje koordynaty kliknięć (NIE nadpisywane przy update)
├── config.json         ← Konfiguracja (ścieżki, domyślne wartości, bazaMdbPath)
├── history.json        ← Historia przetworzonych faktur (ostatnie 50)
├── price_alerts.txt    ← Log alertów cenowych
├── uploads\            ← Przesłane PDF/CSV i wygenerowane XLS/XLSX
├── ibiznes.ahk         ← Skrypt AHK (aktualizowany przy każdym uruchomieniu)
└── *.log               ← Logi (server.log, ahk.log, pdf_converter.log)
```

> Dane użytkownika **NIE są usuwane** przy odinstalowaniu ani aktualizacji programu.

---

## Struktura plików repo

```
iBiznesPythonBot/
│
├── main.py            # Entry point – flaskwebgui + Flask (okno Edge w trybie app)
├── server.py          # Flask backend API (wszystkie endpointy)
├── pdf_to_csv.py      # Parser PDF faktur + eksporter CSV/XLS/XLSX
├── ibiznes.ahk        # AutoHotkey v2 – automatyzacja GUI iBiznes
│
├── ui.html            # Panel UI (bundlowany w .exe)
├── coords.json        # Domyślne koordynaty (kopiowane do %APPDATA% jeśli brak)
├── version.txt        # Wersja programu
│
├── iBiznesBot.spec    # PyInstaller spec – budowanie .exe
├── build.bat          # Skrypt budowania (PyInstaller)
├── installer/
│   └── setup.iss      # Inno Setup – budowanie instalatora .exe
│
└── .github/
    └── workflows/
        ├── ci.yml     # GitHub Actions – syntax check + auto-merge PR
        └── build.yml  # GitHub Actions – auto-build .exe + instalator przy tagu
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
`pandas`, `pyinstaller`, `pyodbc` itp.) i uruchamia PyInstaller.

**Wynik:** `dist\iBiznesBot\iBiznesBot.exe` (folder z .exe – gotowy do użycia)

### Krok 2 – Zbuduj instalator (opcjonalnie)

```bat
iscc installer\setup.iss
```

Lub otwórz `installer/setup.iss` w **Inno Setup Compiler** GUI → Build → Compile.

**Wynik:** `dist\installer\iBiznesBot-Setup-v3.2.2.exe`

> **Bez Inno Setup:** Możesz rozdystrybuować folder `dist\iBiznesBot\` lub sam plik
> `dist\iBiznesBot\iBiznesBot.exe` (portable, nie wymaga instalacji).

### Auto-build (GitHub Actions)

Każdy nowy tag `v*.*.*` na `main` automatycznie uruchamia workflow `.github/workflows/build.yml`,
który buduje `.exe` + instalator i wgrywa `iBiznesBot-Setup-vX.Y.Z.exe` do GitHub Release.

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

### Polskie nazwy nie działają (MDB)

**Przyczyna:** Brak sterownika Microsoft Access Database Engine lub błędna ścieżka do `.mdb`.

**Rozwiązanie:**
1. Pobierz i zainstaluj [Microsoft Access Database Engine 2016 x64](https://www.microsoft.com/en-us/download/details.aspx?id=54920)
2. Panel: **⚙ Ustawienia → Ścieżka do bazy MDB** → wpisz ścieżkę do pliku `.mdb`

> Bez sterownika program działa normalnie – używa nazw z PDF.

---

### Antywirus blokuje iBiznesBot.exe

PyInstaller bundluje Python interpreter + biblioteki w jeden .exe, co może wywołać
fałszywy alarm (false positive). Dodaj `iBiznesBot.exe` do wyjątków antywirusa.

---

## FAQ

**P: Czy program wymaga Pythona?**
O: Nie. Python jest zawarty w pliku .exe (bundlowany przez PyInstaller).

**P: Czy program wymaga WebView2 lub .NET?**
O: Na Windows 10/11 WebView2 Runtime jest wbudowany w system – nie wymaga dodatkowej instalacji.
Na starszych systemach WebView2 zostanie zainstalowany automatycznie przez Microsoft.

**P: Czy dane z v2.x zostaną zachowane?**
O: Tak – jeśli masz skonfigurowane `coords.json` i `config.json`, skopiuj je do
`%APPDATA%\iBiznesBot\` po instalacji.

**P: Skąd pobierać aktualizacje?**
O: Program automatycznie sprawdza nowe wersje co 6 godzin (banner w górnej części okna).
Kliknij "Pobierz" aby przejść do strony Releases na GitHubie.

**P: Gdzie są moje koordynaty i konfiguracja?**
O: W `%APPDATA%\iBiznesBot\` (wpisz w pasku Explorer: `%APPDATA%\iBiznesBot`).

**P: Co jeśli mam dwa monitory?**
O: Koordynaty muszą odpowiadać pozycji okna iBiznes na konkretnym monitorze.
WindowSpy wyświetla absolutne Screen X/Y z uwzględnieniem układu wielomonitorowego.

**P: Czy mogę używać komputera podczas działania bota?**
O: Nie – bot przejmuje sterowanie myszą i klawiaturą.

**P: Co robi bot gdy AHK nie jest zainstalowany?**
O: Uruchamia tryb symulacji – pokazuje pozycje w panelu bez wprowadzania ich do iBiznes.
Służy do weryfikacji danych z PDF/CSV przed faktycznym uruchomieniem.

**P: Co to jest tolerancja cenowa?**
O: Gdy cena z faktury różni się od ceny w systemie iBiznes o więcej niż tolerancja (domyślnie 0.05 PLN), pozycja trafia do alertów zamiast do faktury. Ustawiana w **⚙ Ustawienia**.

---

## Changelog

### v3.2.2 (2026-03) – Hotfix: przywrócono flaskwebgui (błędy instalatora v3.2.1)
- **Przywrócono `flaskwebgui`** – pywebview + pythonnet powodowały błędy DLL i wymagały specyficznej wersji .NET na maszynie użytkownika; flaskwebgui działa na każdym Windows 10/11 bez dodatkowych zależności
- **v3.2.1 oznaczona jako ZEPSUTA** – nie używaj tej wersji

### v3.2.1 (2026-03) – ZEPSUTA (nie używaj)
- Próba zastąpienia flaskwebgui przez pywebview – spowodowała błędy instalatora (konflikty DLL pythonnet)

### v3.2.0 (2026-03) – Duży bugfix
- **Naprawiono krytyczny błąd podwójnego `_finish()`** – gdy AutoHotkey nie był zainstalowany, symulacja woła `_finish(True)`, ale `server.py` woła następnie `_finish(False)` drugi raz → UI pokazywał błąd mimo poprawnego działania symulacji. Naprawione przez zwrócenie `True` po symulacji
- **Naprawiono `.replace(".pdf", ...)` w `api_pdf_upload()`** – użycie `str.replace` mogło nadpisać wiele wystąpień `.pdf` w ścieżce pliku; zastąpione przez `os.path.splitext()` (robustność)
- **Naprawiono `.replace('.csv', ...)` w `pdf_to_csv.py convert()`** – ta sama klasa błędów; zastąpione przez `os.path.splitext()`
- **Naprawiono `workflow_dispatch` w `build.yml`** – ręczne uruchomienie CI kończyło się błędem `gh release upload main ...` (nazwa brancha zamiast tagu); dodano warunek `if: startsWith(github.ref, 'refs/tags/')`
- **Naprawiono Inno Setup version pin** – usunięto `--version 6.2.2` z `choco install innosetup` (wersja mogła nie istnieć w Chocolatey)
- **Dodano pywin32 post-install step w build.yml** – rejestracja DLL pywin32 na Windows CI (`continue-on-error: true`)
- **Zaktualizowano stale docstringi** – `server.py`, `main.py`, `pdf_to_csv.py` wskazywały na v3.0 / PyWebView / pywinauto
- **Zaktualizowano UI** – tytuł okna i logo-version wyświetlały `v3.0.0` zamiast bieżącej wersji

### v3.1.1 (2026-03)
- **Naprawiono błąd krytyczny ibiznes.ahk** – parser JSON zawierał `;` jako separator instrukcji (niedozwolony w AHK v2 – traktowany jako komentarz); wszystkie `{ stmt1; stmt2 }` przepisane na poprawny styl wieloliniowy
- **GitHub Actions auto-build** – nowy workflow `.github/workflows/build.yml`; przy każdym nowym release tagu automatycznie buduje `.exe` (PyInstaller + Python 3.11) i instalator (Inno Setup 6) oraz wgrywa `iBiznesBot-Setup-vX.Y.Z.exe` do GitHub Release

### v3.1.0 (2026-03)
- **Auto-wykrywanie aktualizacji** – program sprawdza GitHub Releases przy starcie i co 6 godzin; zielony banner gdy dostępna jest nowsza wersja
- **Poprawne porównanie wersji semver** – endpoint `/api/check-update` używa porównania krotek zamiast string compare
- **Polskie nazwy produktów z bazy MDB** – `pdf_to_csv.py` wczytuje nazwy z bazy iBiznes (Access `.mdb`) przez pyodbc; fallback na nazwę z PDF gdy brak sterownika/bazy
- **Kod produktu = 5 cyfr** – Excel/XLS skraca `10048.01` → `10048`
- **bazaMdbPath** – nowe pole konfiguracji wskazujące na plik `.mdb`
- **build.bat** – dodano `pyodbc` do listy zależności
- **iBiznesBot.spec** – `pyodbc` w `hiddenimports`

### v3.0.0 (2026-03)
**Pełny rewrite projektu:**
- **Program .exe** – flaskwebgui + Flask bundlowany przez PyInstaller; okno Edge w trybie app (brak Pythona na systemie)
- **Instalator .exe** – Inno Setup 6.1+; instalacja do Program Files, skrót pulpit/Start Menu, deinstalator, AutoHotkey v2 bundlowany
- **Dane użytkownika** przeniesione do `%APPDATA%\iBiznesBot\` (coords.json, config.json, logi, uploads)
- **ibiznes.ahk** – ścieżki plików zaktualizowane do APPDATA
- **server.py** – DATA_DIR refactor; nowy endpoint `/api/check-update` (GitHub Releases API)
- **main.py** – nowy entry point (flaskwebgui + Flask + setup APPDATA)
- **ui.html** – banner aktualizacji, branding v3.0
- **build.bat** – skrypt automatycznego budowania (python -m pip, python -m PyInstaller)
- **installer/setup.iss** – Inno Setup script (wbudowany download AHK, bez zewnętrznych pluginów)
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
