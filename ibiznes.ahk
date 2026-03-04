; ============================================================================
;  ibiznes.ahk  –  AutoHotkey v2  –  Automatyzacja faktur zakupowych
;  Wywoływany przez server.py:  AutoHotkey64.exe ibiznes.ahk task.json
;  Tryb: klikanie absolutnych pikseli (koordynaty z coords.json)
;  v3.0 – dane w %APPDATA%\iBiznesBot\
; ============================================================================

#Requires AutoHotkey v2.0
#SingleInstance Force

; ── ŚCIEŻKI DANYCH (APPDATA) ──────────────────────────────────────────────
global DataDir    := A_AppData . "\iBiznesBot"
global TaskFile   := DataDir . "\task.json"
global ResultFile := DataDir . "\result.json"
global CoordsFile := DataDir . "\coords.json"
global LogFile    := DataDir . "\ahk.log"

; ── INICJALIZACJA ──────────────────────────────────────────────────────────
global ResultItems := []
global AhkLog      := FileOpen(LogFile, "a", "UTF-8")
global Coords      := Map()

LogMsg("AutoHotkey bot uruchomiony (v3.0)")

; Plik zadania – z argumentu lub domyślny w APPDATA
taskFilePath := (A_Args.Length >= 1) ? A_Args[1] : TaskFile

if !FileExist(taskFilePath) {
    LogMsg("BŁĄD: Nie znaleziono pliku: " taskFilePath)
    WriteResult(false, "task.json nie znaleziony")
    ExitApp(1)
}

taskJson := FileRead(taskFilePath, "UTF-8")
task     := JSON.parse(taskJson)

nip       := task["nip"]
invoiceNr := task["invoiceNr"]
supplier  := task.Has("supplier") ? task["supplier"] : nip
exePath   := task["exePath"]
items     := task["items"]

LogMsg("Task wczytany: nip=" nip " faktura=" invoiceNr " dostawca=" supplier " pozycji=" items.Length)

; Wczytaj koordynaty
LoadCoords()

try {
    ; 1. Otwórz iBiznes jeśli nie działa
    if !WinExist("ahk_exe iBiznes.exe") {
        if (exePath != "" && FileExist(exePath)) {
            LogMsg("Uruchamianie: " exePath)
            Run(exePath)
            if !WinWait("ahk_exe iBiznes.exe",, 30) {
                WriteResult(false, "iBiznes nie uruchomiony w 30s")
                ExitApp(1)
            }
            Sleep(5000)
        } else {
            LogMsg("BŁĄD: iBiznes nie uruchomiony i brak ścieżki EXE")
            WriteResult(false, "iBiznes nie znaleziony")
            ExitApp(1)
        }
    }

    WinActivate("ahk_exe iBiznes.exe")
    WinWaitActive("ahk_exe iBiznes.exe",, 15)
    LogMsg("Okno iBiznes aktywne.")
    Sleep(500)

    ; 2. Krok 1 – kliknij przycisk "Zakup (...)"
    LogMsg("=== Krok 1: btnZakup ===")
    ClickAbs("btnZakup")
    Sleep(1500)

    ; 3. Krok 2 – kliknij nowy dokument w lewym panelu
    LogMsg("=== Krok 2: btnNewDoc ===")
    ClickAbs("btnNewDoc")
    Sleep(1000)

    ; 4. Krok 3 – pole dostawcy: wpisz i zatwierdź Enterem
    LogMsg("=== Krok 3: supplierField – wpisuję: " supplier)
    ClickAbs("supplierField")
    Sleep(200)
    Send("^a")
    Send(supplier)
    Send("{Enter}")
    Sleep(2500)   ; iBiznes ładuje dane dostawcy

    ; 5. Krok 4 – kliknij zakładkę Pozycje
    LogMsg("=== Krok 4: tabPositions ===")
    ClickAbs("tabPositions")
    Sleep(800)

    ; 6. Krok 5 – F7 (Dodaj z Kartoteki)
    LogMsg("=== Krok 5: F7 – Dodaj z Kartoteki ===")
    Send("{F7}")
    Sleep(1000)
    WinActivate("ahk_exe iBiznes.exe")
    Sleep(200)

    ; 7. Pętla F3 – dodaj każdą pozycję
    LogMsg("=== Krok 6: Pętla F3 – " items.Length " pozycji ===")
    Loop items.Length {
        item   := items[A_Index]
        kod    := item["kod"]
        nazwa  := item["nazwa"]
        ilosc  := item["ilosc"]
        qtyStr := FormatQty(ilosc)

        LogMsg("─── [" A_Index "/" items.Length "] " kod " qty=" ilosc " ───")

        Send("{F3}")
        Sleep(600)
        Send(kod)
        Sleep(200)
        Send("{Enter}")
        Sleep(1200)   ; iBiznes szuka produktu
        Send(qtyStr)
        Sleep(200)
        Send("{Enter}")
        Sleep(800)

        AddResult(kod, nazwa, ilosc, true, "")
        LogMsg("  Dodano: " kod " x" ilosc)
    }

    ; 8. Zapisz dokument (Ctrl+S)
    LogMsg("=== Krok 7: Zapis (Ctrl+S) ===")
    Send("^s")
    Sleep(1000)

    ; Obsługa ewentualnego dialogu potwierdzenia
    Sleep(500)
    if WinExist("ahk_class #32770 ahk_exe iBiznes.exe") {
        Send("{Enter}")
        LogMsg("Dialog potwierdzony.")
        Sleep(500)
    }

    LogMsg("=== BOT ZAKOŃCZONY SUKCESEM ===")
    WriteResult(true, "")

} catch Error as e {
    LogMsg("BŁĄD KRYTYCZNY: " e.Message)
    WriteResult(false, e.Message)
    ExitApp(1)
}

ExitApp(0)

; ============================================================================
; FUNKCJE
; ============================================================================

; ─── WCZYTAJ KOORDYNATY ───────────────────────────────────────────────────
LoadCoords() {
    global Coords, CoordsFile
    if !FileExist(CoordsFile) {
        LogMsg("OSTRZEŻENIE: coords.json nie znaleziony – klikanie niemożliwe")
        return
    }
    try {
        raw    := FileRead(CoordsFile, "UTF-8")
        parsed := JSON.parse(raw)
        for key, val in parsed {
            if (key != "_comment")
                Coords[key] := val
        }
        LogMsg("Koordynaty wczytane: " Coords.Count " punktów")
    } catch Error as e {
        LogMsg("BŁĄD wczytywania coords.json: " e.Message)
    }
}

; ─── KLIKNIJ NA ABSOLUTNYCH KOORDYNATACH ─────────────────────────────────
ClickAbs(name) {
    global Coords
    if !Coords.Has(name) {
        LogMsg("OSTRZEŻENIE: brak koordynatu '" name "' w coords.json")
        return
    }
    x := Coords[name]["x"]
    y := Coords[name]["y"]
    Click(x, y)
    LogMsg("  ClickAbs(" name ") → (" x "," y ")")
    Sleep(150)
}

; ─── FORMATOWANIE ILOŚCI ──────────────────────────────────────────────────
FormatQty(ilosc) {
    if (Mod(ilosc, 1) = 0)
        return Integer(ilosc)
    return StrReplace(Format("{:.3f}", ilosc), ".", ",")
}

; ─── WYNIKI ───────────────────────────────────────────────────────────────
AddResult(kod, nazwa, ilosc, success, msg) {
    global ResultItems
    ResultItems.Push(Map(
        "kod",     kod,
        "nazwa",   nazwa,
        "ilosc",   ilosc,
        "invPLN",  0,
        "sysPLN",  0,
        "success", JsonBool(success),
        "message", msg,
    ))
}

WriteResult(success, errorMsg) {
    global ResultItems, ResultFile

    items := []
    for item in ResultItems {
        items.Push(Map(
            "kod",     item["kod"],
            "nazwa",   item["nazwa"],
            "ilosc",   item["ilosc"],
            "invPLN",  item["invPLN"],
            "sysPLN",  item["sysPLN"],
            "success", item["success"],
            "message", item["message"],
        ))
    }

    result := Map(
        "success", JsonBool(success),
        "error",   errorMsg,
        "items",   items,
    )

    try FileDelete(ResultFile)
    FileAppend(JSON.stringify(result), ResultFile, "UTF-8-RAW")
    LogMsg("Zapisano result.json (" items.Length " pozycji).")
}

; ─── LOGOWANIE ────────────────────────────────────────────────────────────
LogMsg(msg) {
    global AhkLog
    ts := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    line := ts " [AHK] " msg
    AhkLog.WriteLine(line)
    OutputDebug(line)
}

; ============================================================================
; Opakowanie dla boolean
; ============================================================================
class JsonBool {
    __New(v) => this.val := (v ? true : false)
}

; ============================================================================
; JSON – minimalna implementacja dla AHK v2
; ============================================================================
class JSON {
    static parse(str) {
        pos := 1
        return JSON._parseValue(str, &pos)
    }

    static stringify(val) {
        return JSON._serializeValue(val)
    }

    static _parseValue(s, &p) {
        JSON._skipWS(s, &p)
        c := SubStr(s, p, 1)
        if (c = '"')  return JSON._parseString(s, &p)
        if (c = '{')  return JSON._parseObject(s, &p)
        if (c = '[')  return JSON._parseArray(s, &p)
        if (c = 't')  { p += 4; return true }
        if (c = 'f')  { p += 5; return false }
        if (c = 'n')  { p += 4; return "" }
        return JSON._parseNumber(s, &p)
    }

    static _skipWS(s, &p) {
        while (p <= StrLen(s) && InStr(" `t`r`n", SubStr(s, p, 1)))
            p++
    }

    static _parseString(s, &p) {
        p++
        result := ""
        while (p <= StrLen(s)) {
            c := SubStr(s, p, 1)
            if (c = '"') { p++; return result }
            if (c = '\') {
                p++
                ec := SubStr(s, p, 1)
                result .= (ec = 'n') ? '`n'
                        : (ec = 't') ? '`t'
                        : (ec = 'r') ? '`r'
                        : ec
            } else {
                result .= c
            }
            p++
        }
        return result
    }

    static _parseNumber(s, &p) {
        start := p
        while (p <= StrLen(s) && InStr("0123456789.-eE+", SubStr(s, p, 1)))
            p++
        n := SubStr(s, start, p - start)
        return InStr(n, '.') ? Float(n) : Integer(n)
    }

    static _parseObject(s, &p) {
        p++
        obj := Map()
        JSON._skipWS(s, &p)
        if (SubStr(s, p, 1) = '}') { p++; return obj }
        loop {
            JSON._skipWS(s, &p)
            key := JSON._parseString(s, &p)
            JSON._skipWS(s, &p)
            p++
            val := JSON._parseValue(s, &p)
            obj[key] := val
            JSON._skipWS(s, &p)
            c := SubStr(s, p, 1)
            p++
            if (c = '}') return obj
        }
    }

    static _parseArray(s, &p) {
        p++
        arr := []
        JSON._skipWS(s, &p)
        if (SubStr(s, p, 1) = ']') { p++; return arr }
        loop {
            arr.Push(JSON._parseValue(s, &p))
            JSON._skipWS(s, &p)
            c := SubStr(s, p, 1)
            p++
            if (c = ']') return arr
        }
    }

    static _serializeValue(val) {
        if (Type(val) = "JsonBool")
            return (val.val ? "true" : "false")
        t := Type(val)
        if (t = "String")
            return '"' StrReplace(StrReplace(val, '\', '\\'), '"', '\"') '"'
        if (t = "Integer" || t = "Float")
            return val
        if (t = "Map")
            return JSON._serializeObject(val)
        if (t = "Array")
            return JSON._serializeArray(val)
        return "null"
    }

    static _serializeObject(obj) {
        if (obj.Count = 0)
            return "{}"
        parts := ""
        for k, v in obj {
            if (parts != "")
                parts .= ","
            parts .= '"' k '":' JSON._serializeValue(v)
        }
        return "{" parts "}"
    }

    static _serializeArray(arr) {
        if (arr.Length = 0)
            return "[]"
        parts := ""
        for v in arr {
            if (parts != "")
                parts .= ","
            parts .= JSON._serializeValue(v)
        }
        return "[" parts "]"
    }
}
