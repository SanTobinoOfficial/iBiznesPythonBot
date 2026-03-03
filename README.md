# iBiznes Bot v2.0 – Panel automatyzacji faktur zakupowych

Zautomatyzowany panel do wprowadzania faktur zakupowych w programie **iBiznes**. Odczytuje dane z pliku PDF lub CSV i za pomocą **AutoHotkey v2** klika w odpowiednie elementy iBiznes, wpisując kody produktów i ilości. Posiada też **TRYB BEZPIECZNY** – konwersję PDF/CSV do pliku Excel 2003 (.xls) gotowego do ręcznego importu.

---

## Spis treści

1. [Wymagania](#wymagania)
2. [Instalacja](#instalacja)
3. [Uruchomienie](#uruchomienie)
4. [Jak działa program](#jak-działa-program)
   - [Tryb Normalny (AHK)](#tryb-normalny-ahk)
   - [TRYB BEZPIECZNY → XLS](#tryb-bezpieczny--xls)
5. [Konfiguracja koordynatów](#konfiguracja-koordynatów)
6. [Struktura plików](#struktura-plików)
7. [Rozwiązywanie problemów](#rozwiązywanie-problemów)
8. [FAQ](#faq)
9. [Changelog](#changelog)

---

## Wymagania

| Oprogramowanie | Minimalna wersja | Link |
|---|---|---|
| Python | 3.9+ | https://www.python.org/downloads/ |
| AutoHotkey | v2.0 | https://www.autohotkey.com/ |
| iBiznes | dowolna | — |
| System | Windows 10/11 | — |

> **Ważne:** Podczas instalacji Pythona zaznacz opcję **"Add Python to PATH"**.

---

## Instalacja

### Szybki start (masz już pliki)

1. Uruchom **`INSTALL.bat`** jako Administrator (prawy przycisk → Uruchom jako administrator)
2. Poczekaj na komunikat `Instalacja zakonczona pomyslnie!`
3. Uruchom program przez **`START.bat`**

### Instalacja z jednego pliku (tylko INSTALL.bat)

Jeśli masz tylko plik `INSTALL.bat`, wystarczy go uruchomić jako Administrator.

Instalator automatycznie (4 kroki):
1. Sprawdza Python
2. **Pobiera całe repozytorium jako ZIP** z GitHuba i rozpakowuje wszystkie pliki projektu (`coords.json` **nigdy** nie jest nadpisywany)
3. Pobiera i instaluje **AutoHotkey v2** (jeśli nie jest zainstalowany)
4. Instaluje biblioteki Python

**Pobierane pliki projektu (z repo ZIP):**
```
server.py, pdf_to_csv.py, ibiznes.ahk, ui.html, START.bat, version.txt
coords.json  ← tylko jeśli NIE istnieje (zawiera Twoje ustawienia koordynatów)
```

**Instalowane biblioteki Python:**
```
flask, flask-cors, requests, pandas, pywinauto,
Pillow, pdfplumber, openpyxl, xlwt
```

---

### Automatyczne aktualizacje

**`START.bat` automatycznie sprawdza aktualizacje** przy każdym uruchomieniu:

1. Pobiera `version.txt` z GitHuba (timeout 5 sekund)
2. Porównuje z lokalną wersją
3. Jeśli wersja się różni → uruchamia `INSTALL.bat FORCE` (w nowym procesie), który pobiera ZIP z repo i rozpakowuje pliki
4. Po aktualizacji program restartuje się automatycznie z flagą `SKIP_UPDATE`
5. `coords.json` z Twoimi koordynatami **nigdy** nie jest nadpisywany

> Jeśli nie masz internetu – sprawdzanie jest pomijane i program działa normalnie.

---

## Uruchomienie

Kliknij dwukrotnie **`START.bat`**.

Otworzy się:
- Okno terminala **"iBiznes Bot Serwer"** (nie zamykaj go!)
- Przeglądarka z panelem UI na `http://localhost:5000`

> Jeśli przeglądarka się nie otworzy automatycznie – wejdź ręcznie na `http://localhost:5000`.

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

> **Uwaga:** iBiznes musi być otwarty i widoczny na ekranie. Nie używaj komputera podczas działania bota – przejmuje on sterowanie myszą i klawiaturą.

---

### TRYB BEZPIECZNY → XLS

Konwertuje PDF lub CSV do pliku **Excel 2003 (.xls)** w formacie importu iBiznes – bez uruchamiania AHK i bez automatycznych kliknięć.

**Jak użyć:**
1. Kliknij **"🔒 TRYB BEZPIECZNY (→ XLS)"** na górze panelu Bot
2. Wgraj plik PDF lub CSV
3. Wybierz walutę faktury
4. Kliknij **"🔄 Konwertuj do XLS"**
5. Pobierz wygenerowany plik `.xls`
6. Zaimportuj ręcznie do iBiznes: **Dokumenty → Import z pliku EXCEL'a**

**Format pliku XLS (24 kolumny, bez nagłówków):**

| Kol. | Pole iBiznes | Źródło |
|------|-------------|--------|
| A | Kod towaru | kod_produktu |
| B | Nr katalogowy | kod_produktu |
| C | Nazwa towaru | nazwa |
| D | Magazyn | (puste) |
| E | Rodzaj (T/U) | T |
| F | Dodać do kartoteki | N |
| G | Ilość | ilosc |
| H | Cena zakupu NETTO PLN | cena × kurs NBP |
| I | Cena zakupu BRUTTO PLN | H × 1.23 (VAT 23%) |
| J | EAN | ean |
| K | Zmienić cenę sprzedaży | N |
| L | VAT % | 23 |
| M | Nazwa VAT | 23% |
| N | Jednostka miary | jm / szt |
| O | PKWiU/CN | (puste) |
| P | Cena sprzedaży NETTO 1 | (puste) |
| Q | Cena sprzedaży BRUTTO 1 | (puste) |
| R | Cena zakupu DEWIZOWA | cena oryginalna w walucie |
| S | Nazwa dostawcy | z nagłówka faktury |
| T | Producent | (puste) |
| U | Grupa | (puste) |
| V | Waga netto | (puste) |
| W | Waga brutto | (puste) |
| X | Kraj pochodzenia | (puste) |

---

## Konfiguracja koordynatów

### Co to są koordynaty?

Bot klika na **bezwzględnych współrzędnych ekranu** (Screen X, Y) – piksel liczony od lewego górnego rogu monitora (0,0). Koordynaty dla każdego z 5 kroków są zapisane w pliku `coords.json`.

> **Ważne:** Jeśli zmienisz rozdzielczość monitora, przesuniesz lub zmienisz rozmiar okna iBiznes albo podłączysz inny monitor – **musisz zaktualizować koordynaty**.

### Jak znaleźć koordynaty (WindowSpy)

1. Upewnij się że **AutoHotkey v2** jest zainstalowany
2. Uruchom **iBiznes** i ustaw okno w normalnej pozycji
3. Kliknij **prawym przyciskiem** na ikonę AutoHotkey w zasobniku systemowym (prawy dolny róg ekranu → pasek zadań → pokaż ukryte ikony)
4. Wybierz **"WindowSpy"**
5. Najedź kursorem myszy na element w iBiznes (np. przycisk "Zakup")
6. Odczytaj wartości **"Screen"** w sekcji "Mouse Position": `X: 256  Y: 77`
7. Wpisz te wartości w panelu

### Gdzie ustawić koordynaty

**Opcja 1 – W panelu UI (zalecane):**
1. Otwórz zakładkę **⚙ Ustawienia**
2. Zejdź do sekcji **"Koordynaty iBiznes · coords.json"**
3. Wpisz X i Y dla każdego z 5 kroków bota
4. Kliknij **"💾 Zapisz koordynaty"**

**Opcja 2 – Ręcznie w pliku `coords.json`:**
```json
{
  "_comment": "Absolutne koordynaty ekranu (Screen X,Y).",
  "btnZakup":      {"x": 256, "y":  77},
  "btnNewDoc":     {"x":  72, "y": 172},
  "supplierField": {"x": 256, "y": 157},
  "tabPositions":  {"x":  67, "y": 313},
  "btnF7":         {"x": 420, "y": 117}
}
```

**Opis kluczy:**

| Klucz | Krok | Opis elementu w iBiznes |
|---|---|---|
| `btnZakup` | 1 | Przycisk "Zakup (...)" – otwiera moduł zakupów |
| `btnNewDoc` | 2 | Nowy dokument w lewym panelu (lista faktur) |
| `supplierField` | 3 | Pole dostawcy (wpisz nazwę i Enter) |
| `tabPositions` | 4 | Zakładka "Pozycje" w dokumencie |
| `btnF7` | 5 | Przycisk "F7 – Dodaj z Kartoteki" |

---

## Struktura plików

```
iBiznesPythonBot/
│
├── server.py          # Serwer Flask (backend API, porty, endpointy)
├── ui.html            # Panel UI (frontend, otwierany w przeglądarce)
├── ibiznes.ahk        # Skrypt AutoHotkey v2 (klika piksele, wpisuje tekst)
├── pdf_to_csv.py      # Parser PDF faktur + eksporter do CSV/XLS
│
├── coords.json        # Koordynaty pikseli dla AHK (edytuj w Ustawieniach!)
├── version.txt        # Numer wersji – porównywany z GitHubem przy starcie
├── config.json        # Konfiguracja programu (auto-generowany)
├── history.json       # Historia faktur (auto-generowany)
│
├── task.json          # Tymczasowy plik zadania dla AHK (auto, nie edytuj)
├── result.json        # Wyniki z AHK (auto, nie edytuj)
├── ahk.log            # Log działania AHK – sprawdź przy błędach!
├── server.log         # Log serwera Flask
├── price_alerts.txt   # Alerty cenowe (auto)
│
├── uploads/           # Wgrane PDF/CSV i wygenerowane XLS (auto)
│
├── INSTALL.bat        # Pełny instalator: pobiera pliki + AHK v2 + biblioteki Python
└── START.bat          # Uruchamianie programu (codzienne użytkowanie)
```

---

## Rozwiązywanie problemów

### Aktualizacja nie uruchamia się / START.bat nie wykrywa nowej wersji

**Sprawdź:**
1. **BASE_URL** – czy jest poprawnie ustawiony w `START.bat` (musi być taki sam jak w `INSTALL.bat`)
2. **version.txt na GitHubie** – sprawdź czy plik istnieje w repozytorium i ma inną wersję niż lokalny `version.txt`
3. **Brak internetu** – przy braku połączenia sprawdzanie jest pomijane automatycznie
4. **Timeout** – domyślnie 5 sekund. Jeśli masz wolne łącze, możesz zwiększyć go w `START.bat` (zmień `-TimeoutSec 5` na `-TimeoutSec 15`)

**Ręczna aktualizacja:**
```cmd
INSTALL.bat FORCE
```
Lub kliknij dwukrotnie `INSTALL.bat` – pobierze wszystkie brakujące pliki.

---

### INSTALL.bat nie pobiera plików / "Nie udalo sie pobrac"

**Przyczyny i rozwiązania:**

1. **Zły REPO_ZIP** – otwórz `INSTALL.bat` w Notatniku i sprawdź linię:
   ```
   set "REPO_ZIP=https://github.com/SanTobinoOfficial/iBiznesPythonBot/archive/refs/heads/main.zip"
   ```
   Musi wskazywać na właściwy URL ZIP repozytorium GitHub.

2. **Brak internetu** – sprawdź połączenie sieciowe i spróbuj ponownie.

3. **Firewall / antywirus blokuje PowerShell** – uruchom `INSTALL.bat` jako Administrator.

4. **Brak miejsca na dysku** – instalator pobiera ZIP (~kilka MB). Upewnij się że masz wolne miejsce.

---

### Bot klika w złe miejsca / nie klika wcale

**Przyczyna:** Niepoprawne koordynaty w `coords.json`.

**Rozwiązanie:**
1. Uruchom iBiznes i ustaw okno w normalnej pozycji
2. Użyj **WindowSpy** (patrz: [Jak znaleźć koordynaty](#jak-znaleźć-koordynaty))
3. Zaktualizuj w panelu: **⚙ Ustawienia → Koordynaty → Zapisz**

**Diagnostyka:** Sprawdź plik `ahk.log` – każde kliknięcie jest logowane:
```
2026-03-03 12:00:01 [AHK] ClickAbs(btnZakup) → (256, 77)
```
Jeśli koordynaty w logu są błędne – zaktualizuj `coords.json`.

---

### "iBiznes nie znaleziony" / iBiznes się nie otwiera

**Rozwiązanie:**
1. Otwórz **⚙ Ustawienia → Ścieżki → Ścieżka do iBiznes.exe**
2. Kliknij **"🔍 Wykryj"** – auto-wykrywa w typowych lokalizacjach
3. Lub wpisz ścieżkę ręcznie, np.: `C:\Program Files\iBiznes\iBiznes.exe`
4. Kliknij "Zapisz ustawienia"

---

### "AHK nie znaleziony" / AutoHotkey nie uruchamia się

**Rozwiązanie:**
1. Zainstaluj **AutoHotkey v2** (nie v1!): https://www.autohotkey.com/
2. Otwórz **⚙ Ustawienia → Ścieżka do AutoHotkey64.exe**
3. Kliknij **"🔍 Wykryj"** lub wpisz ręcznie:
   `C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe`

> **Ważne:** Wymagana jest wersja **v2**. Wersja v1 nie obsługuje składni używanej w skrypcie.

---

### "❌ Błąd: Brak biblioteki: xlwt"

**Rozwiązanie:**
```cmd
python -m pip install xlwt
```
Lub uruchom ponownie `INSTALL.bat`.

---

### "❌ Błąd parsowania PDF" / nie znaleziono pozycji w PDF

Program obsługuje PDF-y faktur **LEVIOR** i **FESTA**. Inne formaty mogą nie być obsługiwane.

**Rozwiązanie:**
- Użyj pliku **CSV** zamiast PDF
- Pobierz przykładowy CSV z przycisku **"⬇ Przykładowy CSV"** w panelu
- Format CSV (pierwsza linia to nagłówek):
  ```
  kod_produktu,nazwa,ilosc,cena_netto_usd
  10048.01,Tape measure KOMELON 8mx25mm,3,4.135
  11105.01,Tape measure FESTA Magnetic 5mx19mm,2,1.888
  ```

---

### Serwer Flask niedostępny / "Brak połączenia"

**Objawy:** Panel UI nie odpowiada, kropka statusu jest szara.

**Rozwiązanie:**
1. Uruchom **`START.bat`** – automatycznie uruchamia serwer
2. Upewnij się że okno **"iBiznes Bot Serwer"** jest otwarte (nie zamykaj!)
3. Ręczne uruchomienie serwera (w CMD w folderze programu):
   ```cmd
   python server.py
   ```
4. Sprawdź czy port 5000 nie jest zajęty przez inny program

---

### Plik coords.json nie istnieje lub jest pusty

**Rozwiązanie:**
1. Otwórz panel → **⚙ Ustawienia → Koordynaty**
2. Wpisz wartości (nawet tymczasowe 0,0)
3. Kliknij **"💾 Zapisz koordynaty"** – plik zostanie utworzony
4. Następnie zaktualizuj prawdziwe koordynaty przez WindowSpy

---

### Bot wprowadza złe ceny / nieprawidłowy kurs walutowy

**Sprawdź:**
- Połączenie z internetem (kurs pobierany z NBP API)
- Wybraną walutę w formularzu (USD/EUR/PLN/CZK)
- Kliknij **↻** obok kursu aby go odświeżyć

---

### Timeout – "AHK nie odpowiedział" (po 600 sekundach)

**Przyczyny i rozwiązania:**
- **Dialog w iBiznes czeka na kliknięcie** – zamknij go ręcznie, powtórz próbę
- **iBiznes reaguje wolno** – zwiększ opóźnienia: **⚙ Ustawienia → Opóźnienie między krokami** (np. 1000 ms)
- **Bot utknął na złym koordynacie** – zaktualizuj coords.json

---

### Kod produktu nie jest znajdowany (F3 nie działa)

**Przyczyny:**
- Produkt nie istnieje w kartotece iBiznes
- Pole wyszukiwania nie jest aktywne po F7

**Rozwiązanie:**
- Sprawdź czy produkty z CSV istnieją w iBiznes (kartoteka towarów)
- Zwiększ opóźnienie po F7: w `ibiznes.ahk` zmień `Sleep(1000)` na `Sleep(2000)` po `Send("{F7}")`

---

## FAQ

**P: Czy bot działa z innymi programami księgowymi?**
O: Nie. Bot jest zaprojektowany wyłącznie pod interfejs iBiznes. Koordynaty i klawisze skrótów (F3, F7, Ctrl+S) są specyficzne dla tego programu.

**P: Co jeśli mam dwa monitory?**
O: Koordynaty muszą odpowiadać pozycji okna iBiznes na konkretnym monitorze. WindowSpy wyświetla absolutne Screen X/Y z uwzględnieniem układu wielomonitorowego.

**P: Czy mogę używać komputera podczas działania bota?**
O: Nie – bot przejmuje sterowanie myszą i klawiaturą. Podczas działania nie używaj komputera.

**P: Co jeśli produkt nie istnieje w kartotece iBiznes?**
O: Bot wpisze kod i naciśnie Enter – jeśli iBiznes nie znajdzie produktu, kolejne kroki mogą trafić w złe miejsca. Upewnij się że wszystkie produkty z CSV/PDF są w kartotece.

**P: Jak zmienić walutę?**
O: W formularzu (tryb normalny) wybierz walutę z listy – kurs pobierze się automatycznie. W TRYBIE BEZPIECZNYM wybierz walutę przed konwersją.

**P: Jak dodać własny format CSV?**
O: CSV musi mieć nagłówek z kolumnami (kolejność dowolna):
```
kod_produktu,nazwa,ilosc,cena_netto_usd
```
Kolumny opcjonalne: `ean`, `jednostka` (domyślnie "szt").

**P: Jak zresetować program do ustawień domyślnych?**
O: Usuń pliki: `config.json`, `history.json`, `coords.json`, `_installed.flag`.
Następnie uruchom `INSTALL.bat` i `START.bat`.

**P: Jak zmienić domyślny NIP lub walutę?**
O: Otwórz **⚙ Ustawienia → Wartości domyślne** i ustaw Domyślny NIP dostawcy oraz Domyślną walutę.

**P: Bot działa ale nie zapisuje dokumentu (Ctrl+S nie działa)?**
O: Upewnij się że fokus jest na oknie iBiznes. Zwiększ opóźnienie po ostatnim F3 → Enter w `ibiznes.ahk` (zmień `Sleep(800)` na `Sleep(1500)`).

---

## Changelog

### v2.2.9 (2026-03)
- **START.bat**: Naprawiono krytyczny błąd – `start /wait` zamiast `call` dla INSTALL.bat (zapobiega nadpisaniem pliku w trakcie jego działania przez samego siebie)
- **START.bat**: Naprawiono porównywanie wersji – PowerShell `.Trim()` usuwa znaki `\r`/CRLF (błąd powodował nieskończoną pętlę aktualizacji)
- **START.bat**: Dodano flagę `SKIP_UPDATE` – po aktualizacji/instalacji program restartuje się bez ponownego sprawdzania wersji

### v2.2.8 (2026-03)
- **START.bat**: Dodano `cd /d "%~dp0"` – poprawne uruchamianie z dowolnej lokalizacji
- **START.bat**: Pętla oczekiwania na serwer Flask (maks. 15 sekund, co sekundę pinguje `localhost:5000`)
- **CI**: Naprawiono auto-merge – dynamiczne wyszukiwanie otwartego PR zamiast hardcoded `#1`

### v2.2.7 (2026-03)
- **server.py**: Naprawiono błąd numeru faktury – zwracano NIP zamiast `invoiceNr`
- **server.py**: Naprawiono pobieranie XLS dla ścieżek ze spacjami i polskimi znakami (URL encode)
- **pdf_to_csv.py**: Naprawiono operator precedence w funkcji parsowania ilości (dodano nawiasy)
- **ibiznes.ahk**: Dodano `WinActivate` po F7 zapobiegające kradzieży focusu przez dialog
- **ibiznes.ahk**: Naprawiono detekcję dialogu iBiznes (`WinExist` z `ahk_exe iBiznes.exe`)
- **INSTALL.bat**: Przepisano – pobiera pełne repozytorium jako ZIP (zamiast plików po kolei)
- **CI**: Dodano GitHub Actions (syntax check Python + auto-merge v2.0 → main)

### v2.0 (2026-03)
- Nowy przepływ AHK: absolutne koordynaty pikseli (`ClickAbs`) – usunięto `ControlSetText`, `DllCall`, `VirtualAllocEx`
- TRYB BEZPIECZNY: konwersja PDF/CSV → Excel 2003 (.xls) w formacie importu iBiznes (24 kolumny, bez nagłówków)
- Koordynaty edytowalne w panelu UI (⚙ Ustawienia → Koordynaty iBiznes)
- Nowe endpointy API: `/api/coords` (GET/POST), `/api/safe-convert`
- Pole "Nazwa dostawcy" – auto-wypełniane z nagłówka PDF, wysyłane do AHK
- Usunięty stary TRYB BEZPIECZNY (podgląd tabeli przed startem, checkbox)
- Dodano `xlwt` do instalatora i zależności
- **INSTALL.bat** rozbudowany: pobiera pliki projektu z GitHub + AutoHotkey v2 + biblioteki Python (wystarczy mieć tylko INSTALL.bat)
- **INSTALL.bat FORCE** – tryb aktualizacji: pobiera wszystkie pliki projektu (nie nadpisuje `coords.json`)
- **START.bat** – automatyczne sprawdzanie aktualizacji przy starcie (porównuje `version.txt` z GitHubem)
- Nowy plik `version.txt` – kontrola wersji dla auto-update

### v1.x
- Podstawowa automatyzacja AHK (ControlSetText, UIA)
- Parser PDF faktur LEVIOR/FESTA (pdfplumber)
- Historia faktur, alerty cenowe, kurs NBP
- Panel Flask + SSE streaming logów
