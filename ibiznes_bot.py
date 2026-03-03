"""
================================================================================
  iBiznes / Zakup – Bot automatyzacji faktur zakupowych
  Oparty na rzeczywistym UI aplikacji (screenshoty 26.02.2026)
================================================================================

INSTALACJA:
    pip install pywinauto pandas Pillow requests

STRUKTURA CSV (kodowanie UTF-8):
    kod_produktu,nazwa,ilosc,cena_netto_usd
    10048.01,Tape measure KOMELON,3,4.135
    17005.01,Side splitting pliers,4,3.031

DZIAŁANIE:
    1. Podłącza się do okna "Zakup" (aplikacja iBiznes)
    2. Krok 1. Faktura  → wpisuje NIP dostawcy → auto-uzupełnienie → nr faktury
    3. Pobiera aktualny kurs USD/PLN z API Narodowego Banku Polskiego
    4. Krok 2. Pozycje  → dla każdego wiersza CSV:
         - wyszukuje produkt po kodzie w polu "Szukaj (F3)"
         - porównuje cenę z systemu z ceną z faktury (przeliczoną USD→PLN)
         - jeśli różnica > próg → alert w price_alerts.txt, produkt pomijany
         - jeśli OK → klika Dodaj → wpisuje ilość → zatwierdza
    5. Klika "Zakończ" aby zapisać fakturę

UWAGA PRZY PIERWSZYM UŻYCIU:
    Ustaw DIAGNOSTIC_MODE = True – bot wydrukuje pełne drzewo UIA
    i zatrzyma się. Sprawdź rzeczywiste auto_id / title kontrolek
    i zaktualizuj stałe w sekcji "MAPOWANIE KONTROLEK" poniżej.
================================================================================
"""

# ── importy ───────────────────────────────────────────────────────────────────
import logging
import os
import sys
import time
import re
from datetime import datetime
from typing import Optional, Any

import pandas as pd
import requests
from pywinauto import Application
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError
from pywinauto.timings import TimeoutError as PWTimeout

# ── konfiguracja globalna ─────────────────────────────────────────────────────

LOG_FILE           = "bot.log"
PRICE_ALERTS_FILE  = "price_alerts.txt"
PRICE_TOLERANCE    = 0.05    # Dopuszczalna różnica ceny faktury vs system [PLN]
MAX_RETRIES        = 3       # Liczba ponownych prób przy błędzie
RETRY_DELAY        = 2.0     # Przerwa między próbami [s]
DEFAULT_TIMEOUT    = 20      # Timeout czekania na kontrolkę [s]
SHORT_TIMEOUT      = 8       # Krótki timeout [s]
APP_LAUNCH_WAIT    = 6       # Oczekiwanie po uruchomieniu EXE [s]

# NBP API – kurs średni USD/PLN (tabela A)
NBP_API_URL = "https://api.nbp.pl/api/exchangerates/rates/a/usd/?format=json"

# ══════════════════════════════════════════════════════════════════════════════
# MAPOWANIE KONTROLEK
# Stałe oparte na screenach. Dostosuj jeśli Twoja wersja programu różni się.
# ══════════════════════════════════════════════════════════════════════════════

# Tytuł głównego okna (pasek tytułu widoczny na screenie: "Zakup")
WINDOW_TITLE_RE    = ".*Zakup.*"

# Zakładki nawigacji (screeny pokazują polskie nazwy)
TAB_STEP1_RE       = "Krok 1. Faktura"
TAB_STEP2_RE       = "Krok 2. Pozycje"

# Przycisk przejścia z Kroku 1 do Kroku 2 (pomarańczowy na screenie)
BTN_WPROWADZ_RE    = ".*[Ww]prowad[źz].*[Pp]ozycj.*"

# Przycisk "Dodaj" w Kroku 2 (między F6-Nowy a F7-Ukryj, screen 2)
BTN_DODAJ_RE       = "Dodaj"

# Przycisk "Zakończ" (prawy górny róg obu screenów)
BTN_ZAKONCZY_RE    = ".*[Zz]ako[ńn]cz.*"

# Indeks kolumny "Cena" w tabeli produktów (0-based)
# Screen 2: Kod=0 | Nazwa=1 | NrKat=2 | Cena=3 | STAN=4 | CenaSprzedazy=5
PRICE_COLUMN_INDEX = 3

# ── logger ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)-8s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("ZakupBot")


# ═════════════════════════════════════════════════════════════════════════════
class ZakupBot:
    """
    Bot automatyzujący wprowadzanie faktur zakupowych w aplikacji iBiznes
    przez natywne kontrolki Windows UIA (pywinauto backend='uia').

    Brak jakichkolwiek kliknięć po współrzędnych / pikselach.
    """

    def __init__(
        self,
        exe_path:       str,
        csv_path:       str,
        supplier_nip:   str,
        invoice_number: str,
        invoice_date:   str = "",
    ) -> None:
        self.exe_path       = exe_path
        self.csv_path       = csv_path
        self.supplier_nip   = supplier_nip
        self.invoice_number = invoice_number
        self.invoice_date   = invoice_date

        self.app:      Optional[Application] = None
        self.main_win: Optional[UIAWrapper]  = None

        self.usd_rate: float = 0.0
        self.data:     Optional[pd.DataFrame] = None

        self.added   = 0
        self.skipped = 0
        self.errors  = 0

        log.info("=" * 64)
        log.info("ZakupBot zainicjalizowany")
        log.info(f"  CSV     : {csv_path}")
        log.info(f"  NIP     : {supplier_nip}")
        log.info(f"  Faktura : {invoice_number}  ({invoice_date})")
        log.info("=" * 64)

    # ─────────────────────────────────────────────────────────────────────────
    # NARZĘDZIA WEWNĘTRZNE
    # ─────────────────────────────────────────────────────────────────────────

    def _retry(self, fn, *a, label: str = "op", **kw) -> Any:
        """Wykonuje fn maksymalnie MAX_RETRIES razy z opóźnieniem."""
        last = None
        for i in range(1, MAX_RETRIES + 1):
            try:
                log.debug(f"[{i}/{MAX_RETRIES}] {label}")
                return fn(*a, **kw)
            except (ElementNotFoundError, PWTimeout,
                    ElementAmbiguousError, Exception) as e:
                last = e
                log.warning(f"[{i}/{MAX_RETRIES}] {label} – {e}")
                time.sleep(RETRY_DELAY)
        log.error(f"Wszystkie próby nieudane: {label}")
        raise last

    def _screenshot(self, ctx: str = "error") -> None:
        """Zapisuje zrzut ekranu z timestampem."""
        try:
            from PIL import ImageGrab
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            fn = f"error_{ctx[:20]}_{ts}.png"
            ImageGrab.grab().save(fn)
            log.info(f"Screenshot: {fn}")
        except Exception as e:
            log.warning(f"Screenshot niemożliwy: {e}")

    def _price_alert(self, row: pd.Series,
                     sys_pln: float, inv_pln: float) -> None:
        """Zapisuje alert cenowy do pliku i loguje ostrzeżenie."""
        ts   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        diff = abs(inv_pln - sys_pln)
        line = (
            f"[{ts}] ALERT CENOWY | "
            f"KOD: {row['kod_produktu']} | NAZWA: {row['nazwa']} | "
            f"CENA_FAKTURA: {inv_pln:.4f} PLN | "
            f"CENA_SYSTEM: {sys_pln:.4f} PLN | "
            f"RÓŻNICA: {diff:.4f} PLN | KURS_USD: {self.usd_rate:.4f}\n"
        )
        with open(PRICE_ALERTS_FILE, "a", encoding="utf-8") as f:
            f.write(line)
        log.warning(
            f"⚠  ALERT CENOWY – {row['kod_produktu']} | "
            f"faktura={inv_pln:.4f} | system={sys_pln:.4f} | Δ={diff:.4f} PLN"
        )

    def _find(self, parent: UIAWrapper, timeout: int = DEFAULT_TIMEOUT,
              **criteria) -> UIAWrapper:
        """Czeka na kontrolkę UIA i zwraca ją."""
        ctrl = parent.child_window(**criteria)
        ctrl.wait("visible", timeout=timeout)
        return ctrl

    def _set_text(self, ctrl: UIAWrapper, value: str) -> None:
        """Bezpieczne wpisanie tekstu: czyści pole i wpisuje wartość."""
        ctrl.set_edit_text("")
        time.sleep(0.1)
        ctrl.set_edit_text(value)
        time.sleep(0.15)

    # ─────────────────────────────────────────────────────────────────────────
    # DIAGNOSTYKA
    # ─────────────────────────────────────────────────────────────────────────

    def print_controls(self, depth: int = 5,
                       window: Optional[UIAWrapper] = None) -> None:
        """
        Drukuje drzewo kontrolek UIA aktywnego okna.

        Uruchom przy PIERWSZYM starcie:
            bot = ZakupBot(...)
            bot.connect_to_app()
            bot.print_controls(depth=5)

        Na podstawie wydruku zaktualizuj stałe w sekcji MAPOWANIE KONTROLEK.
        """
        w = window or self.main_win
        if w is None:
            log.warning("Brak okna – najpierw connect_to_app()")
            return
        print("\n" + "─" * 60)
        print("DRZEWO KONTROLEK UIA:")
        print("─" * 60)
        w.print_control_identifiers(depth=depth)
        print("─" * 60 + "\n")

    # ─────────────────────────────────────────────────────────────────────────
    # 0A. POŁĄCZENIE Z APLIKACJĄ
    # ─────────────────────────────────────────────────────────────────────────

    def connect_to_app(self) -> None:
        """
        Podłącza się do okna 'Zakup' w iBiznes.
        Jeśli aplikacja nie działa i podano exe_path – uruchamia ją.
        """
        log.info("Łączenie z aplikacją...")

        def _connect():
            # Próba podłączenia do działającej instancji
            try:
                self.app = Application(backend="uia").connect(
                    title_re=WINDOW_TITLE_RE,
                    timeout=4
                )
                log.info("Podłączono do działającej instancji.")
            except Exception:
                if not self.exe_path or not os.path.isfile(self.exe_path):
                    raise FileNotFoundError(
                        f"Aplikacja nie działa i EXE niedostępne: '{self.exe_path}'"
                    )
                log.info(f"Uruchamianie: {self.exe_path}")
                self.app = Application(backend="uia").start(
                    self.exe_path, wait_for_idle=False
                )
                time.sleep(APP_LAUNCH_WAIT)

            self.main_win = self.app.window(title_re=WINDOW_TITLE_RE)
            self.main_win.wait("visible", timeout=DEFAULT_TIMEOUT)
            self.main_win.set_focus()
            log.info(f"Okno gotowe: '{self.main_win.window_text()}'")

        self._retry(_connect, label="connect_to_app")

    # ─────────────────────────────────────────────────────────────────────────
    # 0B. KURS USD Z NBP
    # ─────────────────────────────────────────────────────────────────────────

    def fetch_usd_rate(self) -> float:
        """
        Pobiera aktualny kurs USD/PLN z Narodowego Banku Polskiego.
        Używa tabeli kursów średnich (tabela A).

        Endpoint: https://api.nbp.pl/api/exchangerates/rates/a/usd/?format=json

        Returns:
            Kurs USD/PLN jako float (np. 4.0231)
        """
        log.info("Pobieranie kursu USD/PLN z NBP API...")
        try:
            resp = requests.get(NBP_API_URL, timeout=10)
            resp.raise_for_status()
            rate = float(resp.json()["rates"][0]["mid"])
            self.usd_rate = rate
            log.info(f"Kurs USD/PLN (NBP, tabela A): {self.usd_rate:.4f}")
        except requests.RequestException as e:
            log.error(f"Błąd NBP API: {e}")
            self.usd_rate = 4.05  # kurs awaryjny
            log.warning(f"Używam kursu awaryjnego: {self.usd_rate}")
        return self.usd_rate

    # ─────────────────────────────────────────────────────────────────────────
    # 0C. WCZYTANIE CSV
    # ─────────────────────────────────────────────────────────────────────────

    def load_csv(self) -> None:
        """
        Wczytuje plik CSV z pozycjami faktury.

        Wymagane kolumny:
            kod_produktu   – kod towaru (np. "10048.01")
            nazwa          – nazwa towaru
            ilosc          – ilość
            cena_netto_usd – cena netto z faktury w USD
        """
        log.info(f"Wczytywanie CSV: {self.csv_path}")
        df = pd.read_csv(self.csv_path, dtype={"kod_produktu": str,
                                                "nazwa": str}, encoding="utf-8")
        wymagane = {"kod_produktu", "nazwa", "ilosc", "cena_netto_usd"}
        brakujace = wymagane - set(df.columns)
        if brakujace:
            raise ValueError(f"Brakujące kolumny w CSV: {brakujace}")

        df["kod_produktu"]   = df["kod_produktu"].str.strip()
        df["ilosc"]          = pd.to_numeric(df["ilosc"], errors="coerce")
        df["cena_netto_usd"] = pd.to_numeric(df["cena_netto_usd"], errors="coerce")
        df.dropna(subset=["kod_produktu", "ilosc", "cena_netto_usd"], inplace=True)

        self.data = df
        log.info(f"Wczytano {len(self.data)} pozycji z CSV.")
        log.debug("\n" + df.to_string())

    # ─────────────────────────────────────────────────────────────────────────
    # 1. KROK 1. FAKTURA
    # ─────────────────────────────────────────────────────────────────────────

    def fill_step1_faktura(self) -> None:
        """
        Wypełnia zakładkę 'Krok 1. Faktura':
          1. Wpisuje NIP dostawcy → Enter → auto-uzupełnienie nazwy/adresu
          2. Wpisuje numer faktury
          3. Opcjonalnie: ustawia datę faktury

        Oparte na screenshocie #1 aplikacji Zakup.
        """
        log.info("=== KROK 1. FAKTURA ===")

        def _krok1():
            w = self.main_win

            # ── 1a. Aktywuj zakładkę "Krok 1. Faktura" ──────────────────────
            try:
                tab = self._find(w, timeout=SHORT_TIMEOUT,
                                 title=TAB_STEP1_RE, control_type="TabItem")
                tab.click_input()
                time.sleep(0.4)
                log.debug("Aktywowano zakładkę 'Krok 1. Faktura'.")
            except ElementNotFoundError:
                log.debug("Zakładka Krok 1 niedostępna – może już jest aktywna.")

            # ── 1b. Pole NIP ─────────────────────────────────────────────────
            # Screen: sekcja "Dostawca" → wiersz "NIP" → pole Edit
            # Szukamy pola po auto_id lub po kolejności EditBox w oknie
            nip_edit = self._find_nip_field(w)
            self._set_text(nip_edit, self.supplier_nip)
            log.info(f"Wpisano NIP: {self.supplier_nip}")

            # Enter lub Tab → wyzwala auto-fill dostawcy
            nip_edit.type_keys("{TAB}")
            time.sleep(2.0)   # ← poczekaj aż system pobierze dane dostawcy

            # Sprawdź auto-fill (opcjonalnie – tylko logujemy)
            self._log_supplier_name(w)

            # ── 1c. Numer faktury ────────────────────────────────────────────
            nr_edit = self._find_invoice_nr_field(w)
            if nr_edit:
                self._set_text(nr_edit, self.invoice_number)
                log.info(f"Wpisano nr faktury: {self.invoice_number}")
            else:
                log.warning("Pole nr faktury niedostępne – pomijam.")

            # ── 1d. Data faktury (opcjonalnie) ────────────────────────────────
            if self.invoice_date:
                self._set_invoice_date(w)

        self._retry(_krok1, label="fill_step1_faktura")
        log.info("Krok 1 wypełniony.")

    def _find_nip_field(self, w: UIAWrapper) -> UIAWrapper:
        """Szuka pola NIP – 3 strategie fallback."""
        # Strategia 1 – auto_id zawierający "nip"
        for aid in ["nip", "NIP", "txtNip", "editNip"]:
            try:
                f = w.child_window(auto_id=aid, control_type="Edit")
                f.wait("visible", timeout=3)
                log.debug(f"NIP field via auto_id='{aid}'")
                return f
            except (ElementNotFoundError, PWTimeout):
                pass

        # Strategia 2 – Edit po etykiecie "NIP"
        try:
            lbl = w.child_window(title="NIP", control_type="Text")
            lbl.wait("visible", timeout=SHORT_TIMEOUT)
            # Pierwsze Edit za etykietą (w tej samej grupie)
            grp = lbl.parent()
            edits = grp.children(control_type="Edit")
            if edits:
                log.debug("NIP field via parent group Edit[0]")
                return edits[0]
        except Exception:
            pass

        # Strategia 3 – pierwsze Edit w oknie (zakładamy że to NIP)
        log.debug("NIP field via fallback Edit[0]")
        return self._find(w, control_type="Edit", found_index=0)

    def _find_invoice_nr_field(self, w: UIAWrapper) -> Optional[UIAWrapper]:
        """Szuka pola numeru faktury."""
        # Strategia 1 – auto_id
        for aid in ["nrFaktury", "nr", "txtNr", "invoiceNr", "numer"]:
            try:
                f = w.child_window(auto_id=aid, control_type="Edit")
                f.wait("visible", timeout=3)
                return f
            except (ElementNotFoundError, PWTimeout):
                pass

        # Strategia 2 – Edit w sekcji "Faktura"
        try:
            grp = w.child_window(title_re=".*[Ff]aktura.*", control_type="Group")
            edits = grp.children(control_type="Edit")
            if edits:
                return edits[0]
        except ElementNotFoundError:
            pass

        # Strategia 3 – Edit[1] (drugie Edit w oknie po NIP)
        try:
            f = w.child_window(control_type="Edit", found_index=1)
            f.wait("visible", timeout=SHORT_TIMEOUT)
            return f
        except (ElementNotFoundError, PWTimeout):
            return None

    def _log_supplier_name(self, w: UIAWrapper) -> None:
        """Odczytuje i loguje auto-uzupełnioną nazwę dostawcy."""
        try:
            for aid in ["nazwaPelna", "nazwa", "txtNazwa"]:
                try:
                    f = w.child_window(auto_id=aid, control_type="Edit")
                    val = f.get_value()
                    if val:
                        log.info(f"Auto-uzupełniono dostawcę: {val[:60]}")
                        return
                except Exception:
                    pass
        except Exception:
            pass

    def _set_invoice_date(self, w: UIAWrapper) -> None:
        """Ustawia datę faktury jeśli podano."""
        try:
            for aid in ["zDnia", "dataFaktury", "data", "txtData"]:
                try:
                    f = w.child_window(auto_id=aid, control_type="Edit")
                    f.wait("visible", timeout=3)
                    self._set_text(f, self.invoice_date)
                    log.info(f"Wpisano datę: {self.invoice_date}")
                    return
                except (ElementNotFoundError, PWTimeout):
                    pass
        except Exception:
            log.debug("Data faktury – pole niedostępne.")

    # ─────────────────────────────────────────────────────────────────────────
    # PRZEJŚCIE DO KROKU 2
    # ─────────────────────────────────────────────────────────────────────────

    def go_to_step2(self) -> None:
        """
        Klika pomarańczowy przycisk 'Wprowadź pozycje -->'
        lub zakładkę 'Krok 2. Pozycje'.
        """
        log.info("Przejście do Kroku 2. Pozycje...")

        def _go():
            w = self.main_win

            # PRÓBA A – pomarańczowy przycisk "Wprowadź pozycje -->"
            try:
                btn = self._find(w, timeout=SHORT_TIMEOUT,
                                 title_re=BTN_WPROWADZ_RE, control_type="Button")
                btn.click_input()
                log.info("Kliknięto 'Wprowadź pozycje -->'.")
                time.sleep(1.0)
                return
            except (ElementNotFoundError, PWTimeout):
                pass

            # PRÓBA B – zakładka "Krok 2. Pozycje"
            tab = self._find(w, timeout=DEFAULT_TIMEOUT,
                             title_re=TAB_STEP2_RE, control_type="TabItem")
            tab.click_input()
            time.sleep(0.8)
            log.info("Kliknięto zakładkę 'Krok 2. Pozycje'.")

        self._retry(_go, label="go_to_step2")

    # ─────────────────────────────────────────────────────────────────────────
    # 2. PRZETWARZANIE PRODUKTÓW
    # ─────────────────────────────────────────────────────────────────────────

    def process_products(self) -> None:
        """
        Główna pętla – iteruje po wierszach CSV:
          dla każdej pozycji:
            1. Wyszukaj produkt w kartotece po kodzie
            2. Odczytaj cenę z systemu (PLN)
            3. Przelicz cenę z faktury USD → PLN kursem NBP
            4. Porównaj – jeśli różnica > PRICE_TOLERANCE → alert + pomiń
            5. Jeśli OK → kliknij Dodaj → wpisz ilość → zatwierdź
        """
        if self.data is None:
            raise RuntimeError("Brak danych CSV. Wywołaj load_csv().")
        if self.usd_rate == 0.0:
            raise RuntimeError("Brak kursu USD. Wywołaj fetch_usd_rate().")

        log.info(f"=== KROK 2. POZYCJE – {len(self.data)} wierszy ===")

        for idx, row in self.data.iterrows():
            kod = str(row["kod_produktu"])
            qty = float(row["ilosc"])
            usd = float(row["cena_netto_usd"])
            pln = round(usd * self.usd_rate, 4)

            log.info(
                f"\n  [{idx+1}/{len(self.data)}] KOD={kod} | "
                f"qty={qty} | {usd} USD × {self.usd_rate:.4f} = {pln:.4f} PLN"
            )

            try:
                sys_pln = self._search_and_get_price(kod)

                if sys_pln is None:
                    log.warning(f"  → Produkt '{kod}' nie znaleziony – pomijam.")
                    self._price_alert(row, 0.0, pln)
                    self.skipped += 1
                    self._escape_search()
                    continue

                diff = abs(pln - sys_pln)
                log.info(
                    f"  system={sys_pln:.4f} PLN | faktura={pln:.4f} PLN | "
                    f"Δ={diff:.4f} | próg={PRICE_TOLERANCE}"
                )

                if diff > PRICE_TOLERANCE:
                    self._price_alert(row, sys_pln, pln)
                    self.skipped += 1
                    self._escape_search()
                    log.info("  → POMINIĘTO (różnica cen > próg).")
                    continue

                # Cena OK – dodaj produkt
                self._click_dodaj_and_set_qty(qty)
                self.added += 1
                log.info(f"  → DODANO  qty={qty}  ✓")

            except (ElementNotFoundError, PWTimeout) as e:
                self._screenshot(f"UIA_{kod}")
                log.error(f"  Błąd UIA przy {kod}: {e}")
                self.errors += 1
                self._escape_search()
            except Exception as e:
                self._screenshot(f"ERR_{kod}")
                log.error(f"  Błąd przy {kod}: {e}", exc_info=True)
                self.errors += 1
                self._escape_search()

            time.sleep(0.4)

        log.info("\n" + "=" * 64)
        log.info(f"PODSUMOWANIE: dodano={self.added} | "
                 f"pominięto={self.skipped} | błędy={self.errors}")
        log.info("=" * 64)

    # ── wewnętrzne: wyszukiwanie produktu ────────────────────────────────────

    def _search_and_get_price(self, kod: str) -> Optional[float]:
        """
        Wyszukuje produkt po kodzie w tabeli Kroku 2 i odczytuje jego cenę.

        Widok z screenshota #2:
          - Zakładki: Kod | Nazwa | Nr katalogu
          - Pole "Szukaj  (F3)"
          - Tabela: Kod towaru (0) | Nazwa towaru (1) | Nr Kat (2) | Cena (3) | ...
          - Przycisk "Dodaj"

        Returns:
            Cena netto w PLN lub None jeśli nie znaleziono
        """
        w = self.main_win

        # ── Aktywuj zakładkę "Kod" ───────────────────────────────────────────
        try:
            tab_kod = self._find(w, timeout=SHORT_TIMEOUT,
                                 title="Kod", control_type="TabItem")
            tab_kod.click_input()
            time.sleep(0.25)
        except (ElementNotFoundError, PWTimeout):
            pass  # Zakładka może już być aktywna lub mieć inną nazwę

        # ── Wpisz kod w pole "Szukaj  (F3)" ─────────────────────────────────
        search_edit = self._find_search_field(w)
        self._set_text(search_edit, kod)
        time.sleep(0.4)
        search_edit.type_keys("{ENTER}")
        time.sleep(0.9)

        # ── Odczytaj cenę z pierwszego wiersza tabeli ─────────────────────────
        return self._read_price_from_grid(w, kod)

    def _find_search_field(self, w: UIAWrapper) -> UIAWrapper:
        """Lokalizuje pole wyszukiwania produktu – wielostopniowy fallback."""
        # Strategia 1 – placeholder/title zawierający "Szukaj"
        for kw in [{"title_re": ".*[Ss]zukaj.*", "control_type": "Edit"},
                   {"auto_id_re": "(?i).*szukaj.*", "control_type": "Edit"},
                   {"auto_id_re": "(?i).*search.*", "control_type": "Edit"}]:
            try:
                f = w.child_window(**kw)
                f.wait("visible", timeout=SHORT_TIMEOUT)
                log.debug(f"Pole szukaj: {kw}")
                return f
            except (ElementNotFoundError, PWTimeout):
                pass

        # Strategia 2 – pierwsze Edit na zakładce Krok 2
        log.debug("Pole szukaj: fallback Edit[0]")
        return self._find(w, control_type="Edit", found_index=0)

    def _read_price_from_grid(self, w: UIAWrapper, kod: str) -> Optional[float]:
        """
        Odczytuje cenę z kolumny PRICE_COLUMN_INDEX tabeli produktów.
        Zaznacza pierwszy wiersz wyników.
        """
        # Znajdź tabelę (DataGrid lub Table)
        grid = None
        for ct in ("DataGrid", "Table", "List"):
            try:
                grid = w.child_window(control_type=ct)
                grid.wait("visible", timeout=SHORT_TIMEOUT)
                break
            except (ElementNotFoundError, PWTimeout):
                pass

        if grid is None:
            log.debug(f"Tabela produktów niedostępna dla '{kod}'.")
            return None

        # Pobierz wiersze danych
        rows_ctrl = (
            grid.children(control_type="DataItem") or
            grid.children(control_type="ListItem") or
            grid.children(control_type="Custom")
        )
        if not rows_ctrl:
            log.debug(f"Tabela pusta dla '{kod}'.")
            return None

        first = rows_ctrl[0]
        first.click_input()   # zaznacz
        time.sleep(0.25)

        cells = first.children()
        if not cells:
            return None

        # Log pierwszej komórki (kod) dla diagnostyki
        code_cell_text = cells[0].window_text().strip() if cells else "?"
        log.debug(f"Pierwszy wiersz tabeli: kod='{code_cell_text}' "
                  f"(szukano: '{kod}')")

        # Odczytaj cenę z kolumny PRICE_COLUMN_INDEX
        if len(cells) > PRICE_COLUMN_INDEX:
            raw = cells[PRICE_COLUMN_INDEX].window_text().strip()
            try:
                price = self._parse_price(raw)
                log.debug(f"Cena odczytana z kol.{PRICE_COLUMN_INDEX}: '{raw}' → {price}")
                return price
            except ValueError:
                log.debug(f"Nie można sparsować '{raw}' jako ceny.")

        # Fallback – szukamy pierwszej komórki wyglądającej jak cena
        for i, cell in enumerate(cells):
            txt = cell.window_text().strip()
            if self._looks_like_price(txt):
                try:
                    p = self._parse_price(txt)
                    log.debug(f"Cena znaleziona heurystycznie w kol.{i}: {p}")
                    return p
                except ValueError:
                    continue

        log.debug(f"Brak ceny w wierszu dla '{kod}'.")
        return None

    # ── wewnętrzne: dodawanie pozycji ─────────────────────────────────────────

    def _click_dodaj_and_set_qty(self, qty: float) -> None:
        """
        Klika przycisk 'Dodaj' (screen 2) i wpisuje ilość w odpowiednim polu.

        Dwa warianty:
          A) Pojawia się dialog z polem Ilość → wpisz → OK
          B) Ilość wpisuje się bezpośrednio w komórce dolnej tabeli
        """
        w = self.main_win
        qty_str = str(int(qty)) if qty == int(qty) else f"{qty:.3f}".replace(".", ",")

        # ── Kliknij "Dodaj" ───────────────────────────────────────────────────
        dodaj = self._find(w, timeout=DEFAULT_TIMEOUT,
                           title=BTN_DODAJ_RE, control_type="Button")
        dodaj.click_input()
        log.debug("Kliknięto 'Dodaj'.")
        time.sleep(0.7)

        # ── Wariant A: osobny dialog ──────────────────────────────────────────
        try:
            dlg = self.app.window(
                title_re=".*[Ii]lo[śs][śs][ć].*|.*[Pp]ozycj.*|.*[Dd]odaj.*"
            )
            dlg.wait("visible", timeout=4)
            log.debug(f"Dialog: '{dlg.window_text()}'")

            # Pole ilości w dialogu
            qty_edit = self._find_qty_edit(dlg)
            self._set_text(qty_edit, qty_str)
            log.debug(f"Wpisano ilość w dialogu: {qty_str}")

            # Zatwierdź
            for title in ("OK", "Zatwierdź", "Dodaj", "Akceptuj"):
                try:
                    btn = dlg.child_window(title=title, control_type="Button")
                    btn.wait("visible", timeout=3)
                    btn.click_input()
                    time.sleep(0.7)
                    return
                except (ElementNotFoundError, PWTimeout):
                    pass
            # Fallback – Enter
            qty_edit.type_keys("{ENTER}")
            time.sleep(0.7)

        except (ElementNotFoundError, PWTimeout):
            # ── Wariant B: pole ilości w dolnej tabeli ────────────────────────
            log.debug("Brak dialogu – szukam pola Ilość w dolnej tabeli.")
            qty_edit = self._find_qty_edit(w)
            self._set_text(qty_edit, qty_str)
            qty_edit.type_keys("{TAB}")
            log.debug(f"Wpisano ilość w tabeli: {qty_str}")
            time.sleep(0.5)

    def _find_qty_edit(self, parent: UIAWrapper) -> UIAWrapper:
        """Lokalizuje pole Ilość – wiele strategii."""
        for kw in [
            {"auto_id_re": "(?i).*ilo[sś].*", "control_type": "Edit"},
            {"auto_id_re": "(?i).*qty.*|.*quantity.*", "control_type": "Edit"},
            {"title_re": ".*[Ii]lo[śs][śs][ćc].*", "control_type": "Edit"},
        ]:
            try:
                f = parent.child_window(**kw)
                f.wait("visible", timeout=SHORT_TIMEOUT)
                return f
            except (ElementNotFoundError, PWTimeout):
                pass
        # Ostateczny fallback – pierwsze Edit
        return self._find(parent, control_type="Edit", found_index=0)

    def _escape_search(self) -> None:
        """Wychodzi z aktywnego dialogu wyszukiwania przez Escape lub przycisk X."""
        try:
            x_btn = self.main_win.child_window(title="X", control_type="Button")
            x_btn.wait("visible", timeout=3)
            x_btn.click_input()
            time.sleep(0.3)
        except (ElementNotFoundError, PWTimeout):
            try:
                self.main_win.type_keys("{ESCAPE}")
                time.sleep(0.3)
            except Exception:
                pass

    # ─────────────────────────────────────────────────────────────────────────
    # 3. ZAPIS I ZAKOŃCZENIE
    # ─────────────────────────────────────────────────────────────────────────

    def save_and_close(self) -> None:
        """
        Klika przycisk 'Zakończ' (widoczny w prawym górnym rogu na obu screenach)
        aby zapisać i zamknąć fakturę.
        """
        log.info("Zapisywanie – klikam 'Zakończ'...")

        def _save():
            w = self.main_win

            # PRÓBA A – przycisk "Zakończ" (górny prawy, pomarańczowy)
            try:
                btn = self._find(w, timeout=DEFAULT_TIMEOUT,
                                 title_re=BTN_ZAKONCZY_RE, control_type="Button")
                btn.click_input()
                log.info("Kliknięto 'Zakończ'.")
                time.sleep(1.2)
                self._confirm_dialog()
                return
            except (ElementNotFoundError, PWTimeout):
                pass

            # PRÓBA B – menu Plik → Zapisz
            try:
                mbar = self._find(w, timeout=SHORT_TIMEOUT,
                                  control_type="MenuBar")
                plik = mbar.child_window(title_re=".*[Pp]lik.*",
                                         control_type="MenuItem")
                plik.click_input()
                time.sleep(0.4)
                zap = w.child_window(title_re=".*[Zz]apisz.*",
                                     control_type="MenuItem")
                zap.click_input()
                time.sleep(1.0)
                self._confirm_dialog()
                return
            except (ElementNotFoundError, PWTimeout):
                pass

            # PRÓBA C – Ctrl+S
            w.set_focus()
            w.type_keys("^s")
            time.sleep(1.0)
            self._confirm_dialog()
            log.info("Zapisano przez Ctrl+S.")

        self._retry(_save, label="save_and_close")

    def _confirm_dialog(self) -> None:
        """Klika 'Tak' / 'OK' w dialogach potwierdzających."""
        for title_re in [".*[Tt]ak.*", ".*OK.*", ".*[Yy]es.*"]:
            try:
                dlg = self.app.top_window()
                btn = dlg.child_window(title_re=title_re, control_type="Button")
                btn.wait("visible", timeout=4)
                btn.click_input()
                log.info(f"Dialog potwierdzony: '{btn.window_text()}'.")
                time.sleep(0.6)
                return
            except (ElementNotFoundError, PWTimeout):
                pass

    # ─────────────────────────────────────────────────────────────────────────
    # PARSOWANIE CEN
    # ─────────────────────────────────────────────────────────────────────────

    @staticmethod
    def _parse_price(raw: str) -> float:
        """
        Konwertuje tekst ceny na float.
        Obsługuje formaty: "6,67"  "6.67"  "1.234,56"  "29,29 zł"  "  17,34  "
        """
        if not raw or not raw.strip():
            raise ValueError("Puste pole ceny")
        s = re.sub(r"[^\d,.\-]", "", raw.strip())
        if not s:
            raise ValueError(f"Nieparsowalna cena: '{raw}'")
        if "," in s and "." in s:
            # Format "1.234,56" → odrzuć separator tysięcy, zamień przecinek
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", ".")
        return float(s)

    @staticmethod
    def _looks_like_price(text: str) -> bool:
        """True jeśli text wygląda jak cena (np. '6,67' '17.34')."""
        return bool(re.match(r"^\d{1,10}[.,]\d{2,4}$", text.strip()))

    # ─────────────────────────────────────────────────────────────────────────
    # MAIN RUN
    # ─────────────────────────────────────────────────────────────────────────

    def run(self) -> None:
        """Uruchamia pełny proces automatyzacji."""
        log.info("╔═══════════════════════════════════════╗")
        log.info("║    ZakupBot – START PROCESU            ║")
        log.info("╚═══════════════════════════════════════╝")
        try:
            self.connect_to_app()
            self.fetch_usd_rate()
            self.load_csv()
            self.fill_step1_faktura()
            self.go_to_step2()
            self.process_products()
            self.save_and_close()

            print("\n✅  ZAKOŃCZONO POMYŚLNIE")
            print(f"   Dodano pozycji : {self.added}")
            print(f"   Pominięto      : {self.skipped}  (→ {PRICE_ALERTS_FILE})")
            print(f"   Błędy          : {self.errors}")
            print(f"   Kurs USD/PLN   : {self.usd_rate:.4f}  (NBP)")

        except FileNotFoundError as e:
            self._screenshot("FileNotFound")
            log.critical(str(e))
            print(f"\n❌ Nie znaleziono pliku: {e}")
            sys.exit(1)
        except ElementNotFoundError as e:
            self._screenshot("ElementNotFound")
            log.critical(str(e))
            print(f"\n❌ Kontrolka UIA niedostępna: {e}")
            print("   Ustaw DIAGNOSTIC_MODE=True aby wydrukować drzewo kontrolek.")
            sys.exit(1)
        except Exception as e:
            self._screenshot("CriticalError")
            log.critical(str(e), exc_info=True)
            print(f"\n❌ Błąd krytyczny: {e}  (log: {LOG_FILE})")
            sys.exit(1)


# ═════════════════════════════════════════════════════════════════════════════
# GENERATOR PRZYKŁADOWEGO CSV
# ═════════════════════════════════════════════════════════════════════════════

def generate_sample_csv(path: str = "faktura_202600961.csv") -> str:
    """Tworzy przykładowy CSV z pozycjami faktury LEVIOR 202600961."""
    rows = [
        ("10048.01", "Tape measure KOMELON 8mx25mm ECO",        3,  4.135),
        ("11105.01", "Tape measure FESTA Magnetic 5mx19mm",      2,  1.888),
        ("11325.01", "Tape measure FESTA Magnet 5mx25mm",        2,  2.841),
        ("11425.01", "Pendant 1 meter L 175S",                   5,  0.423),
        ("11445.01", "Spirit level pendant",                     5,  0.284),
        ("12250.01", "Steel band FESTA 50mx10mm",                1,  9.397),
        ("13212.01", "Marker permanent WESTBERG double-sided",  12,  0.226),
        ("17005.01", "Side splitting pliers FESTA 180mm",        4,  3.031),
        ("17062.01", "Pliers SIKO FESTA CrV 300mm 0-40mm",       7,  7.241),
        ("19025.01", "Hammer FESTA 500g FIBERGLASS",              2,  3.290),
        ("20493.01", "Step drill WESTBERG HSS 7-40.5mm",         3, 10.569),
        ("38017.01", "Gun for PUR foam FESTA Teflon",            14,  8.102),
        ("42172.01", "Mesh square 16/1.2x1000mmx25m ZN",         2, 58.358),
    ]
    df = pd.DataFrame(rows, columns=["kod_produktu", "nazwa", "ilosc", "cena_netto_usd"])
    df.to_csv(path, index=False, encoding="utf-8")
    print(f"Wygenerowano CSV: {path}")
    return path


# ═════════════════════════════════════════════════════════════════════════════
# PUNKT WEJŚCIA
# ═════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":

    # ──────────────────────────────────────────────────────────────────────────
    # KONFIGURACJA  ← EDYTUJ PRZED URUCHOMIENIEM
    # ──────────────────────────────────────────────────────────────────────────

    # Ścieżka do EXE – zostaw "" jeśli aplikacja jest już otwarta
    EXE_PATH = r""   # np. r"C:\Program Files\iBiznes\iBiznes.exe"

    # Plik CSV z pozycjami faktury
    CSV_PATH = "faktura_202600961.csv"

    # Dane nagłówka faktury (z dokumentu LEVIOR 202600961)
    SUPPLIER_NIP    = "6197393900"   # NIP LEVIOR s.r.o. bez myślników
    INVOICE_NUMBER  = "202600961"
    INVOICE_DATE    = "23.02.2026"

    # ── TRYB DIAGNOSTYCZNY ────────────────────────────────────────────────────
    # Ustaw True przy pierwszym uruchomieniu. Bot wydrukuje drzewo UIA
    # i zatrzyma się – użyj tych danych do weryfikacji nazw kontrolek.
    DIAGNOSTIC_MODE = False
    # ─────────────────────────────────────────────────────────────────────────

    # Generuj CSV jeśli nie istnieje
    if not os.path.isfile(CSV_PATH):
        CSV_PATH = generate_sample_csv(CSV_PATH)

    print("=" * 62)
    print("  ZakupBot – automatyzacja faktur zakupowych (iBiznes UIA)")
    print("=" * 62)
    print(f"  EXE     : {EXE_PATH or '(podłącz do działającej aplikacji)'}")
    print(f"  CSV     : {CSV_PATH}")
    print(f"  NIP     : {SUPPLIER_NIP}")
    print(f"  Faktura : {INVOICE_NUMBER}  ({INVOICE_DATE})")
    print(f"  LOG     : {LOG_FILE}")
    print(f"  ALERTY  : {PRICE_ALERTS_FILE}")
    print("=" * 62 + "\n")

    bot = ZakupBot(
        exe_path       = EXE_PATH,
        csv_path       = CSV_PATH,
        supplier_nip   = SUPPLIER_NIP,
        invoice_number = INVOICE_NUMBER,
        invoice_date   = INVOICE_DATE,
    )

    if DIAGNOSTIC_MODE:
        print("\n=== TRYB DIAGNOSTYCZNY ===")
        print("Łączę z aplikacją i drukuję drzewo kontrolek UIA...\n")
        bot.connect_to_app()
        print("── Krok 1. Faktura ──────────────────────────────────")
        bot.print_controls(depth=5)
        input("\nPrzejdź ręcznie do 'Krok 2. Pozycje' i naciśnij Enter...")
        print("── Krok 2. Pozycje ──────────────────────────────────")
        bot.print_controls(depth=5)
        print("\nUżyj powyższych danych do zaktualizowania stałych")
        print("w sekcji MAPOWANIE KONTROLEK, następnie ustaw DIAGNOSTIC_MODE=False.")
        sys.exit(0)

    bot.run()
