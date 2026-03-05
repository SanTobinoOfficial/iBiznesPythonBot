"""
================================================================================
  server.py  –  Flask backend dla panelu UI bota iBiznes
  Wersja 3.2.2 – dane w %APPDATA%\\iBiznesBot\\, flaskwebgui-ready
================================================================================
  pip install flask flask-cors requests pandas pdfplumber openpyxl xlwt pyodbc flaskwebgui
  Uruchomienie standalone: python server.py
  Uruchomienie w .exe:     import server; server.app.run(...)
================================================================================
"""

import glob
import json
import logging
import os
import queue
import subprocess
import sys
import threading
import time
import urllib.parse
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests
from flask import Flask, Response, jsonify, request, send_file, send_from_directory
from flask_cors import CORS

from pdf_to_csv import InvoicePDFParser, CSVExporter

# ── ŚCIEŻKI ────────────────────────────────────────────────────────────────────

def resource_path(rel: str) -> str:
    """Ścieżka do zasobów bundlowanych przez PyInstaller lub katalog dev."""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


# Folder danych użytkownika – %APPDATA%\iBiznesBot\
DATA_DIR    = os.path.join(os.environ.get('APPDATA', '.'), 'iBiznesBot')
UPLOADS_DIR = os.path.join(DATA_DIR, 'uploads')
os.makedirs(UPLOADS_DIR, exist_ok=True)

LOG_FILE          = os.path.join(DATA_DIR, "server.log")
CONFIG_FILE       = os.path.join(DATA_DIR, "config.json")
HISTORY_FILE      = os.path.join(DATA_DIR, "history.json")
PRICE_ALERTS_FILE = os.path.join(DATA_DIR, "price_alerts.txt")
AHK_SCRIPT        = os.path.join(DATA_DIR, "ibiznes.ahk")
TASK_FILE         = os.path.join(DATA_DIR, "task.json")
RESULT_FILE       = os.path.join(DATA_DIR, "result.json")
COORDS_FILE       = os.path.join(DATA_DIR, "coords.json")

NBP_API           = "https://api.nbp.pl/api/exchangerates/rates/a/{}//?format=json"
GITHUB_RELEASES   = "https://api.github.com/repos/SanTobinoOfficial/iBiznesPythonBot/releases/latest"
VERSION           = "3.2.2"

DEFAULT_COORDS = {
    "_comment":     "Absolutne koordynaty ekranu (Screen X,Y). Zaktualizuj jesli przesuniesz okno iBiznes.",
    "btnZakup":     {"x": 256, "y":  77},
    "btnNewDoc":    {"x":  72, "y": 172},
    "supplierField":{"x": 256, "y": 157},
    "tabPositions": {"x":  67, "y": 313},
    "btnF7":        {"x": 420, "y": 117},
}

HISTORY_MAX = 50

# ── DOMYŚLNA KONFIGURACJA ──────────────────────────────────────────────────────
DEFAULT_CONFIG = {
    "exePath":             "",
    "ahkExePath":          r"C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe",
    "defaultNIP":          "",
    "defaultCurrency":     "USD",
    "tolerance":           0.05,
    "runMode":             "attach",
    "autoOpenIBiznes":     True,
    "autoNavigateZakup":   True,
    "lastInvoiceDate":     "",
    "lastInvoiceNr":       "",
    "lastNIP":             "",
    "ibiznesMenuPath":     "auto",
    "customMenuPath":      "",
    "priceAlertEmail":     "",
    "maxRetries":          3,
    "stepDelay":           500,
    "bazaMdbPath":         "",
}

# ── LOGGING ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)-8s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("Server")

app = Flask(__name__)
CORS(app)

active_jobs: Dict[str, Any] = {}


# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────

def load_config() -> dict:
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, encoding="utf-8") as f:
                saved = json.load(f)
            return {**DEFAULT_CONFIG, **saved}
        except Exception as e:
            log.warning(f"Blad odczytu config.json: {e}")
    return dict(DEFAULT_CONFIG)


def save_config(cfg: dict) -> None:
    merged = {**DEFAULT_CONFIG, **cfg}
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(merged, f, ensure_ascii=False, indent=2)
    log.info("Config zapisany.")


# ─────────────────────────────────────────────────────────────────────────────
# HISTORIA FAKTUR
# ─────────────────────────────────────────────────────────────────────────────

def load_history() -> List[dict]:
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return []


def save_history(entry: dict) -> None:
    history = load_history()
    history = [h for h in history if h.get("invoiceNr") != entry.get("invoiceNr")]
    history.insert(0, entry)
    history = history[:HISTORY_MAX]
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)


# ─────────────────────────────────────────────────────────────────────────────
# AUTO-WYKRYWANIE
# ─────────────────────────────────────────────────────────────────────────────

IBIZNES_SEARCH_PATHS = [
    r"C:\Program Files\iBiznes\iBiznes.exe",
    r"C:\Program Files (x86)\iBiznes\iBiznes.exe",
    r"C:\iBiznes\iBiznes.exe",
    r"C:\Program Files\Comarch ERP\iBiznes\iBiznes.exe",
    r"D:\iBiznes\iBiznes.exe",
    r"D:\Program Files\iBiznes\iBiznes.exe",
]


def autodetect_ibiznes() -> Optional[str]:
    for path in IBIZNES_SEARCH_PATHS:
        if os.path.isfile(path):
            return path
    for base in [r"C:\Program Files", r"C:\Program Files (x86)", r"D:\Program Files"]:
        matches = glob.glob(os.path.join(base, "**", "iBiznes.exe"), recursive=True)
        if matches:
            return matches[0]
    try:
        import winreg
        for root in [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]:
            try:
                key  = winreg.OpenKey(root, r"SOFTWARE\iBiznes")
                path, _ = winreg.QueryValueEx(key, "InstallDir")
                exe  = os.path.join(path, "iBiznes.exe")
                if os.path.isfile(exe):
                    return exe
            except Exception:
                pass
    except ImportError:
        pass
    return None


def autodetect_ahk() -> Optional[str]:
    paths = [
        r"C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe",
        r"C:\Program Files\AutoHotkey\v2\AutoHotkey.exe",
        r"C:\Program Files (x86)\AutoHotkey\v2\AutoHotkey64.exe",
        r"C:\Program Files\AutoHotkey\AutoHotkey64.exe",
    ]
    for p in paths:
        if os.path.isfile(p):
            return p
    return None


# ─────────────────────────────────────────────────────────────────────────────
# JOB RUNNER
# ─────────────────────────────────────────────────────────────────────────────

class JobRunner:
    def __init__(self, job_id: str, payload: dict) -> None:
        self.job_id   = job_id
        self.payload  = payload
        self.config   = load_config()
        self.queue:   queue.Queue = queue.Queue()
        self.running  = True
        self.added    = 0
        self.skipped  = 0
        self.errors   = 0
        self.usd_rate: float = 0.0
        self.thread   = threading.Thread(target=self._run, daemon=True)
        self.thread.start()

    def _emit(self, type_: str, **kwargs) -> None:
        self.queue.put({**kwargs, "type": type_})

    def _log(self, msg: str, level: str = "info") -> None:
        log.info(f"[{self.job_id[:8]}] {msg}")
        self._emit("log", msg=msg, level=level)

    def _stats(self, progress: float, label: str = "") -> None:
        self._emit("stats",
                   added=self.added, skipped=self.skipped,
                   errors=self.errors, progress=progress, label=label)

    def _alert(self, kod: str, nazwa: str, inv_pln: float, sys_pln: float) -> None:
        diff = abs(inv_pln - sys_pln)
        self._emit("alert", kod=kod, nazwa=nazwa,
                   invPLN=round(inv_pln, 4), sysPLN=round(sys_pln, 4))
        ts   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = (f"[{ts}] ALERT | KOD: {kod} | NAZWA: {nazwa} | "
                f"FAKTURA: {inv_pln:.4f} PLN | SYSTEM: {sys_pln:.4f} PLN | "
                f"ROZNICA: {diff:.4f} | KURS: {self.usd_rate:.4f}\n")
        with open(PRICE_ALERTS_FILE, "a", encoding="utf-8") as f:
            f.write(line)

    def _run(self) -> None:
        try:
            self._emit("status", status="running")

            currency = self.payload.get("currency", "USD")
            rate     = self.payload.get("usdRate")
            if not rate:
                rate = self._fetch_rate(currency)
            self.usd_rate = rate
            self._log(f"Kurs {currency}/PLN: {rate:.4f}", "ok")

            data     = self.payload.get("data", [])
            tol      = float(self.payload.get("tolerance", 0.05))
            run_mode = self.payload.get("runMode", "attach")
            diag     = self.payload.get("diagMode", False)

            self._log(f"Pozycji: {len(data)} | Tolerancja: {tol:.2f} PLN | Tryb: {run_mode}")

            if not data:
                self._log("Brak pozycji w CSV.", "error")
                self._finish(False)
                return

            exe_path = self.payload.get("exePath", "") or self.config.get("exePath", "")
            if not exe_path:
                exe_path = autodetect_ibiznes() or ""

            task = {
                "jobId":             self.job_id,
                "nip":               self.payload.get("nip", ""),
                "invoiceNr":         self.payload.get("invoiceNr", ""),
                "invoiceDate":       self.payload.get("invoiceDate", ""),
                "supplier":          self.payload.get("supplier", "") or self.payload.get("nip", ""),
                "exePath":           exe_path,
                "tolerance":         tol,
                "usdRate":           rate,
                "currency":          currency,
                "runMode":           run_mode,
                "diagMode":          diag,
                "autoOpenIBiznes":   self.config.get("autoOpenIBiznes", True),
                "autoNavigateZakup": self.config.get("autoNavigateZakup", True),
                "ibiznesMenuPath":   self.config.get("ibiznesMenuPath", "auto"),
                "customMenuPath":    self.config.get("customMenuPath", ""),
                "stepDelay":         int(self.config.get("stepDelay", 500)),
                "maxRetries":        int(self.config.get("maxRetries", 3)),
                "items": [
                    {
                        "kod":      row.get("kod_produktu", ""),
                        "nazwa":    row.get("nazwa", ""),
                        "ilosc":    float(row.get("ilosc", 0)),
                        "priceUSD": float(row.get("cena_netto_usd", 0)
                                         or row.get("cena_netto", 0)),
                        "pricePLN": round(
                            float(row.get("cena_netto_usd", 0)
                                  or row.get("cena_netto", 0)) * rate, 4
                        ) if currency != "PLN" else float(
                            row.get("cena_netto_usd", 0) or row.get("cena_netto", 0)
                        ),
                    }
                    for row in data
                ],
            }

            with open(TASK_FILE, "w", encoding="utf-8") as f:
                json.dump(task, f, ensure_ascii=False, indent=2)
            self._log(f"task.json zapisany ({len(task['items'])} pozycji).")

            if os.path.exists(RESULT_FILE):
                os.remove(RESULT_FILE)

            save_history({
                "invoiceNr":   task["invoiceNr"],
                "nip":         task["nip"],
                "invoiceDate": task["invoiceDate"],
                "currency":    currency,
                "items":       len(data),
                "startedAt":   datetime.now().isoformat(),
                "status":      "running",
            })

            cfg = load_config()
            cfg["lastNIP"]         = task["nip"]
            cfg["lastInvoiceNr"]   = task["invoiceNr"]
            cfg["lastInvoiceDate"] = task["invoiceDate"]
            if exe_path:
                cfg["exePath"]     = exe_path
            save_config(cfg)

            if not self._launch_ahk():
                self._finish(False)
                return

            self._wait_for_result(timeout=600)

        except Exception as e:
            self._log(f"Blad krytyczny: {e}", "error")
            log.exception("Blad w JobRunner")
            self._finish(False)

    def _fetch_rate(self, currency: str) -> float:
        if currency.upper() == "PLN":
            return 1.0
        try:
            r = requests.get(NBP_API.format(currency.lower()), timeout=8)
            r.raise_for_status()
            return float(r.json()["rates"][0]["mid"])
        except Exception as e:
            self._log(f"Blad NBP API: {e} – kurs 4.05", "warn")
            return 4.05

    def _launch_ahk(self) -> bool:
        cfg     = load_config()
        ahk_exe = cfg.get("ahkExePath", "") or autodetect_ahk() or ""

        if not ahk_exe or not os.path.isfile(ahk_exe):
            self._log(f"AutoHotkey nie znaleziony: {ahk_exe}", "warn")
            self._log("Kontynuuje bez AHK (symulacja).", "warn")
            self._simulate_without_ahk()
            # _simulate_without_ahk() woła _finish(True) – zwracamy True
            # żeby _run() NIE wołał _finish() drugi raz (double-finish bug)
            return True

        if not os.path.isfile(AHK_SCRIPT):
            self._log(f"Skrypt AHK nie znaleziony: {AHK_SCRIPT}", "error")
            return False

        try:
            subprocess.Popen(
                [ahk_exe, AHK_SCRIPT, TASK_FILE],
                creationflags=subprocess.CREATE_NEW_CONSOLE
                if sys.platform == "win32" else 0,
            )
            self._log(f"Uruchomiono AHK: {ahk_exe}", "ok")
            return True
        except Exception as e:
            self._log(f"Blad uruchomienia AHK: {e}", "error")
            return False

    def _wait_for_result(self, timeout: int = 600) -> None:
        self._log("Oczekiwanie na wynik z AHK...")
        waited = 0
        while waited < timeout and self.running:
            time.sleep(1)
            waited += 1
            if os.path.exists(RESULT_FILE):
                try:
                    with open(RESULT_FILE, encoding="utf-8") as f:
                        result = json.load(f)
                    self._process_result(result)
                    return
                except json.JSONDecodeError:
                    time.sleep(0.5)
                    continue
            if waited % 15 == 0:
                self._log(f"  Czekam... {waited}s / {timeout}s", "debug")
        if waited >= timeout:
            self._log("Timeout – AHK nie odpowiedzial.", "error")
            self._finish(False)

    def _process_result(self, result: dict) -> None:
        self._log("Wynik odebrany z AHK.", "ok")
        items = result.get("items", [])
        total = len(items)
        for i, item in enumerate(items):
            kod    = item.get("kod", "?")
            nazwa  = item.get("nazwa", "")
            ok     = item.get("success", False)
            sysPLN = float(item.get("sysPLN", 0))
            invPLN = float(item.get("invPLN", 0))
            msg    = item.get("message", "")
            if ok:
                self.added += 1
                self._log(f"+ {kod} | qty={item.get('ilosc',0)} | {invPLN:.4f} PLN", "ok")
            elif "cen" in msg.lower() or "roznic" in msg.lower():
                self.skipped += 1
                self._alert(kod, nazwa, invPLN, sysPLN)
            else:
                self.errors += 1
                self._log(f"x {kod}: {msg}", "error")
            pct = ((i + 1) / max(total, 1)) * 100
            self._stats(pct, f"[{i+1}/{total}] {kod}")

        history = load_history()
        if history:
            history[0]["status"]     = "done" if result.get("success") else "error"
            history[0]["added"]      = self.added
            history[0]["skipped"]    = self.skipped
            history[0]["errors"]     = self.errors
            history[0]["finishedAt"] = datetime.now().isoformat()
            with open(HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump(history, f, ensure_ascii=False, indent=2)

        self._finish(result.get("success", True))

    def _simulate_without_ahk(self) -> None:
        self._log("-- Symulacja bez AHK (walidacja CSV) --", "warn")
        items = []
        try:
            with open(TASK_FILE, encoding="utf-8") as f:
                task = json.load(f)
            items = task.get("items", [])
        except Exception:
            pass
        for i, item in enumerate(items):
            self._log(f"[SYM] {item['kod']} | qty={item['ilosc']} | "
                      f"{item['pricePLN']:.4f} PLN", "debug")
            self.skipped += 1
            self._stats(((i + 1) / max(len(items), 1)) * 100,
                        f"[{i+1}/{len(items)}] {item['kod']}")
            time.sleep(0.05)
        self._log("Symulacja zakonczona. Zainstaluj AutoHotkey v2.", "warn")
        self._finish(True)

    def _finish(self, success: bool) -> None:
        self.running = False
        self._stats(100, "Zakonczone")
        self._emit("done", success=success,
                   added=self.added, skipped=self.skipped, errors=self.errors)
        self._emit("status", status="done" if success else "error")

    def stop(self) -> None:
        self.running = False
        self._log("Zatrzymane przez uzytkownika.", "warn")
        self._finish(False)


# ─────────────────────────────────────────────────────────────────────────────
# API ENDPOINTS
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/api/ping")
def api_ping():
    return jsonify({"ok": True, "time": datetime.now().isoformat(), "version": VERSION})


def _version_gt(remote: str, current: str) -> bool:
    """Porownuje wersje semver. Zwraca True jesli remote > current."""
    try:
        r_parts = tuple(int(x) for x in remote.split("."))
        c_parts = tuple(int(x) for x in current.split("."))
        return r_parts > c_parts
    except Exception:
        return remote != current


@app.route("/api/check-update")
def api_check_update():
    """Sprawdza najnowszą wersję na GitHub Releases (semver)."""
    try:
        r = requests.get(GITHUB_RELEASES, timeout=5,
                         headers={"Accept": "application/vnd.github+json"})
        r.raise_for_status()
        data        = r.json()
        remote_tag  = data.get("tag_name", "").lstrip("v")
        html_url    = data.get("html_url", "")
        has_update  = bool(remote_tag) and _version_gt(remote_tag, VERSION)
        return jsonify({
            "current":   VERSION,
            "latest":    remote_tag,
            "hasUpdate": has_update,
            "url":       html_url,
        })
    except Exception as e:
        return jsonify({"current": VERSION, "latest": None,
                        "hasUpdate": False, "error": str(e)})


@app.route("/api/rate")
def api_rate():
    currency = request.args.get("currency", "usd").lower()
    if currency == "pln":
        return jsonify({"rate": 1.0, "date": datetime.now().date().isoformat(), "currency": "PLN"})
    try:
        r = requests.get(NBP_API.format(currency), timeout=8)
        r.raise_for_status()
        data = r.json()
        return jsonify({
            "rate":     data["rates"][0]["mid"],
            "date":     data["rates"][0]["effectiveDate"],
            "currency": currency.upper(),
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/config", methods=["GET"])
def api_config_get():
    cfg = load_config()
    if not cfg.get("exePath"):
        found = autodetect_ibiznes()
        if found:
            cfg["exePath"] = found
    if not cfg.get("ahkExePath") or not os.path.isfile(cfg.get("ahkExePath", "")):
        found = autodetect_ahk()
        if found:
            cfg["ahkExePath"] = found
    return jsonify(cfg)


@app.route("/api/config", methods=["POST"])
def api_config_post():
    data = request.get_json(force=True) or {}
    cfg  = load_config()
    cfg.update(data)
    save_config(cfg)
    return jsonify({"ok": True})


@app.route("/api/autodetect")
def api_autodetect():
    ibiznes = autodetect_ibiznes()
    ahk     = autodetect_ahk()
    return jsonify({
        "ibiznes":      ibiznes,
        "ahk":          ahk,
        "ibiznesFound": bool(ibiznes),
        "ahkFound":     bool(ahk),
    })


@app.route("/api/history")
def api_history():
    return jsonify(load_history())


@app.route("/api/history/clear", methods=["POST"])
def api_history_clear():
    if os.path.exists(HISTORY_FILE):
        os.remove(HISTORY_FILE)
    return jsonify({"ok": True})


@app.route("/api/start", methods=["POST"])
def api_start():
    payload = request.get_json(force=True)
    job_id  = str(uuid.uuid4())
    runner  = JobRunner(job_id, payload)
    active_jobs[job_id] = runner
    return jsonify({"jobId": job_id})


@app.route("/api/stop", methods=["POST"])
def api_stop():
    data   = request.get_json(force=True) or {}
    job_id = data.get("jobId")
    if job_id and job_id in active_jobs:
        active_jobs[job_id].stop()
        return jsonify({"ok": True})
    for job in active_jobs.values():
        if job.running:
            job.stop()
    return jsonify({"ok": True})


@app.route("/api/stream/<job_id>")
def api_stream(job_id: str):
    job = active_jobs.get(job_id)
    if not job:
        return Response('data: {"type":"error","msg":"Job not found"}\n\n',
                        mimetype="text/event-stream")
    def generate():
        while True:
            try:
                event = job.queue.get(timeout=30)
                yield f"data: {json.dumps(event, ensure_ascii=False)}\n\n"
                if event.get("type") == "done":
                    break
            except queue.Empty:
                yield 'data: {"type":"ping"}\n\n'
    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache",
                             "X-Accel-Buffering": "no",
                             "Connection": "keep-alive"})


@app.route("/api/alerts")
def api_alerts():
    if not os.path.exists(PRICE_ALERTS_FILE):
        return jsonify({"alerts": []})
    with open(PRICE_ALERTS_FILE, encoding="utf-8") as f:
        lines = [l.strip() for l in f if l.strip()]
    return jsonify({"alerts": lines})


@app.route("/api/alerts/clear", methods=["POST"])
def api_alerts_clear():
    if os.path.exists(PRICE_ALERTS_FILE):
        os.remove(PRICE_ALERTS_FILE)
    return jsonify({"ok": True})


@app.route("/api/logs")
def api_logs():
    lines_n = int(request.args.get("n", 100))
    if not os.path.exists(LOG_FILE):
        return jsonify({"lines": []})
    with open(LOG_FILE, encoding="utf-8", errors="replace") as f:
        all_lines = f.readlines()
    return jsonify({"lines": [l.rstrip() for l in all_lines[-lines_n:]]})


@app.route("/api/pdf-upload", methods=["POST"])
def api_pdf_upload():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "Brak pola 'file' w zadaniu."}), 400
    f = request.files["file"]
    if not f.filename or not f.filename.lower().endswith(".pdf"):
        return jsonify({"ok": False, "error": "Plik musi byc w formacie PDF."}), 400

    safe_name  = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{f.filename}"
    pdf_path   = os.path.join(UPLOADS_DIR, safe_name)
    base_path  = os.path.splitext(pdf_path)[0]
    csv_path   = base_path + ".csv"
    excel_path = base_path + "_ibiznes.xlsx"

    f.save(pdf_path)
    log.info(f"PDF upload: {pdf_path}")

    try:
        parser = InvoicePDFParser(pdf_path)
        items  = parser.parse()

        if not items:
            return jsonify({
                "ok":    False,
                "error": "Nie znaleziono pozycji produktow w PDF.",
            }), 422

        mdb_path = load_config().get("bazaMdbPath", "")
        exporter = CSVExporter(items, parser.header, mdb_path=mdb_path)
        exporter.to_csv(csv_path)

        actual_excel = None
        try:
            exporter.to_excel(excel_path)
            actual_excel = excel_path
        except ImportError:
            log.warning("openpyxl nie zainstalowany – pomijam Excel.")

        with open(csv_path, encoding="utf-8") as fcsv:
            csv_data = fcsv.read()

        log.info(f"PDF sparsowany: {len(items)} pozycji, faktura={parser.header.get('invoice_nr','?')}")

        return jsonify({
            "ok":        True,
            "items":     items,
            "header":    parser.header,
            "csvData":   csv_data,
            "csvPath":   csv_path,
            "excelPath": actual_excel,
            "errors":    parser.errors,
        })

    except Exception as e:
        log.exception("Blad parsowania PDF")
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/safe-convert", methods=["POST"])
def api_safe_convert():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "Brak pola 'file' w zadaniu."}), 400
    f = request.files["file"]
    if not f.filename:
        return jsonify({"ok": False, "error": "Brak nazwy pliku."}), 400

    currency  = request.form.get("currency", "USD").upper()
    rate_str  = request.form.get("rate", "")

    safe_name   = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{f.filename}"
    upload_path = os.path.join(UPLOADS_DIR, safe_name)
    f.save(upload_path)
    log.info(f"Safe-convert upload: {upload_path}")

    try:
        if rate_str:
            try:
                rate = float(rate_str)
            except ValueError:
                rate = 1.0
        else:
            if currency == "PLN":
                rate = 1.0
            else:
                try:
                    r = requests.get(NBP_API.format(currency.lower()), timeout=8)
                    r.raise_for_status()
                    rate = float(r.json()["rates"][0]["mid"])
                except Exception as e:
                    log.warning(f"Blad NBP API w safe-convert: {e} – kurs 4.05")
                    rate = 4.05

        fname_lower = f.filename.lower()

        if fname_lower.endswith(".pdf"):
            parser = InvoicePDFParser(upload_path)
            items  = parser.parse()
            header = parser.header
            if not items:
                return jsonify({"ok": False, "error": "Nie znaleziono pozycji w PDF."}), 422

        elif fname_lower.endswith(".csv"):
            import pandas as pd
            try:
                df = pd.read_csv(upload_path, encoding="utf-8")
            except UnicodeDecodeError:
                df = pd.read_csv(upload_path, encoding="cp1250")
            items = df.to_dict("records")
            normalized = []
            for row in items:
                normalized.append({
                    "kod_produktu":   str(row.get("kod_produktu", row.get("kod", ""))),
                    "nazwa":          str(row.get("nazwa", "")),
                    "ilosc":          float(row.get("ilosc", 1)),
                    "cena_netto_usd": float(row.get("cena_netto_usd",
                                            row.get("cena_netto", 0))),
                    "ean":            str(row.get("ean", "")),
                    "jednostka":      str(row.get("jednostka", row.get("jm", "szt"))),
                })
            items = normalized
            header = {}
            if not items:
                return jsonify({"ok": False, "error": "Brak pozycji w CSV."}), 422
        else:
            return jsonify({"ok": False, "error": "Plik musi być PDF lub CSV."}), 400

        mdb_path   = load_config().get("bazaMdbPath", "")
        exporter   = CSVExporter(items, header, mdb_path=mdb_path)
        xls_name   = safe_name.rsplit(".", 1)[0] + "_ibiznes.xls"
        xls_path   = os.path.join(UPLOADS_DIR, xls_name)
        exporter.to_ibiznes_xls(xls_path, currency=currency, rate=rate)

        download_url = f"/api/download?path={urllib.parse.quote(xls_path)}"
        log.info(f"Safe-convert OK: {xls_path} ({len(items)} pozycji, kurs={rate})")

        return jsonify({
            "ok":          True,
            "downloadUrl": download_url,
            "filename":    xls_name,
            "items":       len(items),
            "currency":    currency,
            "rate":        rate,
        })

    except ImportError as e:
        return jsonify({"ok": False, "error": f"Brak biblioteki: {e}. Zainstaluj: pip install xlwt"}), 500
    except Exception as e:
        log.exception("Blad safe-convert")
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/coords", methods=["GET"])
def api_coords_get():
    if os.path.exists(COORDS_FILE):
        try:
            with open(COORDS_FILE, encoding="utf-8") as f:
                return jsonify(json.load(f))
        except Exception as e:
            return jsonify({"error": str(e)}), 500
    return jsonify(DEFAULT_COORDS)


@app.route("/api/coords", methods=["POST"])
def api_coords_post():
    data = request.get_json(force=True) or {}
    try:
        existing = dict(DEFAULT_COORDS)
        if os.path.exists(COORDS_FILE):
            with open(COORDS_FILE, encoding="utf-8") as f:
                existing = json.load(f)
        existing.update(data)
        with open(COORDS_FILE, "w", encoding="utf-8") as f:
            json.dump(existing, f, ensure_ascii=False, indent=2)
        log.info("coords.json zaktualizowany.")
        return jsonify({"ok": True})
    except Exception as e:
        log.error(f"Blad zapisu coords.json: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/download")
def api_download():
    file_path = request.args.get("path", "")
    if not file_path:
        return jsonify({"error": "Brak parametru path"}), 400
    abs_path    = os.path.abspath(file_path)
    abs_uploads = os.path.abspath(UPLOADS_DIR)
    if not abs_path.startswith(abs_uploads):
        return jsonify({"error": "Niedozwolona sciezka"}), 403
    if not os.path.isfile(abs_path):
        return jsonify({"error": "Plik nie istnieje"}), 404
    directory = os.path.dirname(abs_path)
    filename  = os.path.basename(abs_path)
    return send_from_directory(directory, filename, as_attachment=True)


@app.route("/")
def serve_ui():
    """Serwuje ui.html z bundlowanych zasobów (PyInstaller) lub folderu dev."""
    return send_file(resource_path("ui.html"))


# ─────────────────────────────────────────────────────────────────────────────
# MAIN – uruchomienie standalone (python server.py)
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    cfg     = load_config()
    changed = False
    if not cfg.get("exePath"):
        found = autodetect_ibiznes()
        if found:
            cfg["exePath"] = found
            changed = True
    if not cfg.get("ahkExePath") or not os.path.isfile(cfg.get("ahkExePath", "")):
        found = autodetect_ahk()
        if found:
            cfg["ahkExePath"] = found
            changed = True
    if changed:
        save_config(cfg)

    print("=" * 60)
    print(f"  iBiznes Bot v{VERSION} - Serwer Flask")
    print("=" * 60)
    print(f"  Panel UI  : http://localhost:5000")
    print(f"  iBiznes   : {cfg.get('exePath') or '(nie znaleziony)'}")
    print(f"  AHK       : {cfg.get('ahkExePath') or '(nie znaleziony)'}")
    print(f"  DataDir   : {DATA_DIR}")
    print("=" * 60)

    app.run(host="127.0.0.1", port=5000, debug=False,
            threaded=True, use_reloader=False)
