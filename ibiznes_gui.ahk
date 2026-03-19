; ============================================================================
;  ibiznes_gui.ahk  –  AutoHotkey v2  –  Tryb AHK GUI (standalone)
;  v3.3.0
;
;  Użycie:
;    1. Uruchom serwer iBiznesBot (iBiznesBot.exe lub python server.py)
;    2. Otwórz FakturaF.exe ręcznie
;    3. Uruchom ten skrypt dwuklikiem
;    4. Wybierz plik PDF, ustaw parametry, kliknij START
;
;  Skrypt parsuje PDF przez API serwera, tworzy task.json z flagą
;  skipLaunch=true, po czym uruchamia ibiznes.ahk (bez auto-uruchamiania FakturaF).
; ============================================================================

#Requires AutoHotkey v2.0
#SingleInstance Force

; ── KONFIGURACJA ─────────────────────────────────────────────────────────────
global SERVER_URL    := "http://127.0.0.1:5000"
global DataDir       := A_AppData . "\iBiznesBot"
global AhkScriptPath := DataDir . "\ibiznes.ahk"
global TaskFilePath  := DataDir . "\task.json"
global LogFilePath   := DataDir . "\ahk.log"

; ── STAN ─────────────────────────────────────────────────────────────────────
global g_JobPID      := 0
global g_LastLogSize := 0
global g_Running     := false

; ── GUI ───────────────────────────────────────────────────────────────────────
MyGui := Gui("+MinSize560x480 -MaximizeBox", "iBiznesBot — Tryb AHK GUI  v3.3.0")
MyGui.SetFont("s10", "Segoe UI")

; Plik PDF
MyGui.Add("GroupBox", "xm y8 w600 h58", "Plik PDF faktury")
global ePdf := MyGui.Add("Edit", "x20 y28 w520 h22")
MyGui.Add("Button", "x548 y26 w60 h26", "...").OnEvent("Click", BrowsePDF)

; Parametry
MyGui.Add("GroupBox", "xm y74 w600 h70", "Parametry faktury")
MyGui.Add("Text", "x20 y94", "Waluta:")
global eCurrency := MyGui.Add("ComboBox", "x70 y92 w70", ["USD", "EUR", "GBP", "PLN"])
eCurrency.Value := 1
MyGui.Add("Text", "x155 y94", "Kurs (0=NBP):")
global eRate := MyGui.Add("Edit", "x240 y92 w80 h22", "0")
MyGui.Add("Text", "x335 y94", "Rabat (%):")
global eDiscount := MyGui.Add("Edit", "x400 y92 w60 h22", "8")
MyGui.Add("Text", "x475 y94", "NIP:")
global eNip := MyGui.Add("Edit", "x500 y92 w110 h22", "")

; Przyciski + status serwera
global btnStart := MyGui.Add("Button", "xm y152 w150 h36", "▶  START")
btnStart.SetFont("s10 Bold")
btnStart.OnEvent("Click", StartBot)
global btnStop := MyGui.Add("Button", "x+8 yp w130 h36 Disabled", "■  STOP")
btnStop.OnEvent("Click", StopBot)
MyGui.Add("Button", "x+8 yp w120 h36", "Wyczyść log").OnEvent("Click",
    (*) => (eLog.Value := "", global g_LastLogSize := 0))
global lblSrv := MyGui.Add("Text", "x+16 yp+10 w120 cGray", "● sprawdzam...")

; Log
MyGui.Add("Text", "xm y196", "Log automacji:")
global eLog := MyGui.Add("Edit", "xm y214 w600 h290 ReadOnly VScroll -E0x200", "")
eLog.SetFont("s9", "Consolas")

MyGui.Add("Text", "xm y+6 w600 cGray",
    "Uwaga: otwórz FakturaF.exe zanim klikniesz START.  Serwer musi być uruchomiony.")

MyGui.OnEvent("Close", (*) => ExitApp())
MyGui.Show("w630 h538")

; Sprawdzaj serwer co 4s
SetTimer(CheckServer, 4000)
CheckServer()
return

; ============================================================================
; FUNKCJE
; ============================================================================

BrowsePDF(*) {
    f := FileSelect(1,, "Wybierz plik PDF faktury", "PDF (*.pdf)")
    if f
        ePdf.Value := f
}

; ─────────────────────────────────────────────────────────────────────────────
StartBot(*) {
    global g_JobPID, g_LastLogSize, g_Running

    pdf := Trim(ePdf.Value)
    if (!pdf || !FileExist(pdf)) {
        MsgBox("Wybierz istniejący plik PDF faktury!", "Błąd", 16)
        return
    }
    if !WinExist("ahk_exe FakturaF.exe") {
        if MsgBox("FakturaF.exe nie jest uruchomiony.`n`nCzy na pewno kontynuować?`n"
                . "(Skrypt będzie czekał max 60s na uruchomienie FakturaF.)",
                "Brak FakturaF.exe", 0x24) = "No"
            return
    }

    ; ── Wywołaj API serwera: parsuje PDF → task.json z skipLaunch=true ───────
    currency := eCurrency.Text
    rate     := (Trim(eRate.Value) = "" || Trim(eRate.Value) = "0") ? 0 : Float(eRate.Value)
    disc     := (Trim(eDiscount.Value) = "") ? 8 : Integer(eDiscount.Value)
    nip      := Trim(eNip.Value)

    body := '{"pdfPath":' _JsonStr(pdf)
          . ',"currency":"' currency '"'
          . ',"usdRate":' rate
          . ',"discount":' disc
          . ',"nip":"' nip '"'
          . ',"skipLaunch":true}'

    AddLog("=== Wysyłam zadanie do serwera... ===`n")
    AddLog("PDF: " pdf "`n")

    try {
        http := ComObject("WinHttp.WinHttpRequest.5.1")
        http.Open("POST", SERVER_URL "/api/run-ahk-gui", false)
        http.SetRequestHeader("Content-Type", "application/json")
        http.SetTimeouts(5000, 5000, 30000, 30000)
        http.Send(body)
        if (http.Status != 200) {
            AddLog("BŁĄD serwera (" http.Status "): " SubStr(http.ResponseText, 1, 300) "`n")
            MsgBox("Błąd API serwera (HTTP " http.Status "):`n"
                 . SubStr(http.ResponseText, 1, 300), "Błąd API", 16)
            return
        }
        AddLog("Serwer OK – task.json wygenerowany`n")
    } catch Error as e {
        AddLog("BŁĄD połączenia z serwerem: " e.Message "`n")
        MsgBox("Nie można połączyć się z serwerem (127.0.0.1:5000).`n`n"
             . "Uruchom iBiznesBot.exe i spróbuj ponownie.", "Brak serwera", 16)
        return
    }

    ; ── Znajdź AutoHotkey exe ────────────────────────────────────────────────
    ahkExe := _FindAhkExe()
    if !ahkExe {
        MsgBox("Nie znaleziono AutoHotkey v2.`n"
             . "Zainstaluj AHK v2 lub ustaw ścieżkę w ustawieniach serwera.", "Brak AHK", 16)
        return
    }
    if !FileExist(AhkScriptPath) {
        MsgBox("Nie znaleziono ibiznes.ahk w:`n" AhkScriptPath
             . "`n`nSprawdź czy serwer iBiznesBot jest uruchomiony (kopiuje skrypt do AppData).",
               "Brak skryptu", 16)
        return
    }

    ; Wyczyść stary log
    try FileDelete(LogFilePath)
    g_LastLogSize := 0

    ; ── Uruchom ibiznes.ahk ──────────────────────────────────────────────────
    Run('"' ahkExe '" "' AhkScriptPath '" "' TaskFilePath '"',, "Hide", &g_JobPID)
    AddLog("Uruchomiono ibiznes.ahk  (PID=" g_JobPID ")`n")
    AddLog("Log: " LogFilePath "`n")
    AddLog("──────────────────────────────────────────────────`n")

    g_Running        := true
    btnStart.Enabled := false
    btnStop.Enabled  := true
    SetTimer(MonitorLog, 500)
    SetTimer(CheckFinished, 2000)
}

; ─────────────────────────────────────────────────────────────────────────────
StopBot(*) {
    global g_JobPID, g_Running

    if g_JobPID
        try ProcessClose(g_JobPID)
    SetTimer(MonitorLog, 0)
    SetTimer(CheckFinished, 0)
    g_Running        := false
    g_JobPID         := 0
    btnStart.Enabled := true
    btnStop.Enabled  := false
    AddLog("`n=== Zatrzymano przez użytkownika ===`n")
}

; ─────────────────────────────────────────────────────────────────────────────
MonitorLog() {
    global g_LastLogSize

    if !FileExist(LogFilePath)
        return
    try {
        f := FileOpen(LogFilePath, "r", "UTF-8")
        if !IsObject(f)
            return
        f.Seek(0, 2)
        newSize := f.Tell()
        if (newSize > g_LastLogSize) {
            f.Seek(g_LastLogSize)
            chunk := f.Read(newSize - g_LastLogSize)
            f.Close()
            g_LastLogSize := newSize
            AddLog(chunk)
        } else {
            f.Close()
        }
    } catch {
        ; plik chwilowo zablokowany przez AHK – czekamy
    }
}

; ─────────────────────────────────────────────────────────────────────────────
CheckFinished() {
    global g_JobPID, g_Running

    if (!g_Running || !g_JobPID)
        return
    if !ProcessExist(g_JobPID) {
        MonitorLog()                 ; ostatni odczyt logu
        SetTimer(MonitorLog, 0)
        SetTimer(CheckFinished, 0)
        g_Running        := false
        g_JobPID         := 0
        btnStart.Enabled := true
        btnStop.Enabled  := false
        AddLog("`n=== Bot zakończył pracę ===`n")
    }
}

; ─────────────────────────────────────────────────────────────────────────────
CheckServer() {
    try {
        http := ComObject("WinHttp.WinHttpRequest.5.1")
        http.Open("GET", SERVER_URL "/api/ping", false)
        http.SetTimeouts(2000, 2000, 3000, 3000)
        http.Send()
        if (http.Status = 200) {
            lblSrv.Text := "● Serwer: online"
            lblSrv.Opt("cGreen")
        } else {
            lblSrv.Text := "● Serwer: błąd"
            lblSrv.Opt("cRed")
        }
    } catch {
        lblSrv.Text := "● Serwer: offline"
        lblSrv.Opt("cRed")
    }
}

; ─────────────────────────────────────────────────────────────────────────────
AddLog(text) {
    eLog.Value := eLog.Value . text
    SendMessage(0x115, 7, 0,, eLog.Hwnd)   ; WM_VSCROLL SB_BOTTOM
}

; ─────────────────────────────────────────────────────────────────────────────
_FindAhkExe() {
    ; 1) Spróbuj odczytać z config.json serwera
    cfgFile := DataDir . "\config.json"
    if FileExist(cfgFile) {
        try {
            raw := FileRead(cfgFile, "UTF-8")
            cfg := JSON.parse(raw)
            if (cfg.Has("ahkExePath") && FileExist(cfg["ahkExePath"]))
                return cfg["ahkExePath"]
        }
    }
    ; 2) Standardowe lokalizacje AHK v2
    for p in [
        "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe",
        "C:\Program Files\AutoHotkey\AutoHotkey64.exe",
        A_ProgramFiles "\AutoHotkey\v2\AutoHotkey64.exe",
        A_ProgramFiles "\AutoHotkey\AutoHotkey64.exe",
    ] {
        if FileExist(p)
            return p
    }
    return ""
}

; ─────────────────────────────────────────────────────────────────────────────
; Escape string do JSON
_JsonStr(s) {
    s := StrReplace(s, "\",  "\\")
    s := StrReplace(s, '"',  '\"')
    s := StrReplace(s, "`n", "\n")
    s := StrReplace(s, "`r", "\r")
    s := StrReplace(s, "`t", "\t")
    return '"' s '"'
}

; ============================================================================
; JSON – minimalna implementacja (taka sama jak w ibiznes.ahk)
; ============================================================================

class JsonBool {
    __New(v) => this.val := (v ? true : false)
}

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
        if (c = 't') { p += 4 ; return true }
        if (c = 'f') { p += 5 ; return false }
        if (c = 'n') { p += 4 ; return "" }
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
            if (c = '"')  { p++ ; return result }
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
        if (SubStr(s, p, 1) = '}') { p++ ; return obj }
        loop {
            JSON._skipWS(s, &p)
            key := JSON._parseString(s, &p)
            JSON._skipWS(s, &p)
            p++   ; ':'
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
        if (SubStr(s, p, 1) = ']') { p++ ; return arr }
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
            return '"' StrReplace(StrReplace(StrReplace(val, '\', '\\'), '"', '\"'), '`n', '\n') '"'
        if (t = "Integer" || t = "Float")
            return val
        if (t = "Map")   return JSON._serializeObject(val)
        if (t = "Array") return JSON._serializeArray(val)
        return "null"
    }

    static _serializeObject(obj) {
        if (obj.Count = 0) return "{}"
        parts := ""
        for k, v in obj {
            if (parts != "") parts .= ","
            parts .= '"' k '":' JSON._serializeValue(v)
        }
        return "{" parts "}"
    }

    static _serializeArray(arr) {
        if (arr.Length = 0) return "[]"
        parts := ""
        for v in arr {
            if (parts != "") parts .= ","
            parts .= JSON._serializeValue(v)
        }
        return "[" parts "]"
    }
}
