; ============================================================================
;  ibiznes.ahk  –  AutoHotkey v2  –  Automatyzacja faktur zakupowych
;  v4.1 – nowy workflow, FakturaF.exe, detekcja WinGetText, bardzo szcz. logi
;
;  Wywołanie: AutoHotkey64.exe ibiznes.ahk task.json
;
;  DETEKCJA PRODUKTU W KATALOGU:
;    Po F3 + kod + Enter (1. Enter = szukaj):
;      - WinGetText liczy wystąpienia kodu 5-cyfrowego
;      - 1 wystąpienie  = tylko pole szukania → NIE MA w katalogu
;      - 2+ wystąpienia = pole szukania + wiersz listy  → JEST w katalogu
;      - Fallback: ControlGetText na typowych kontrolkach siatki
; ============================================================================

#Requires AutoHotkey v2.0
#SingleInstance Force

; ── ŚCIEŻKI ─────────────────────────────────────────────────────────────────
global DataDir    := A_AppData . "\iBiznesBot"
global TaskFile   := DataDir . "\task.json"
global ResultFile := DataDir . "\result.json"
global LogFile    := DataDir . "\ahk.log"

; ── ZMIENNE GLOBALNE ─────────────────────────────────────────────────────────
global ResultItems := []
global AhkLog      := FileOpen(LogFile, "a", "UTF-8")
global Task        := Map()
global D           := Map()
global g_Added     := 0
global g_New       := 0
global g_Errors    := 0
global g_StartTime := A_TickCount

LogSep()
LogMsg("=== iBiznes Bot v4.1 URUCHOMIONY ===")
LogMsg("Data/czas: " FormatTime(, "yyyy-MM-dd HH:mm:ss"))

; ── WCZYTAJ TASK.JSON ────────────────────────────────────────────────────────
taskFilePath := (A_Args.Length >= 1) ? A_Args[1] : TaskFile
LogMsg("Plik zadania: " taskFilePath)

if !FileExist(taskFilePath) {
    LogMsg("BŁĄD KRYTYCZNY: Nie znaleziono pliku task.json: " taskFilePath)
    WriteResult(false, "task.json nie znaleziony: " taskFilePath)
    ExitApp(1)
}

taskJson := FileRead(taskFilePath, "UTF-8")
LogMsg("task.json wczytany (" StrLen(taskJson) " bajtów)")
Task := JSON.parse(taskJson)

; ── OPÓŹNIENIA (ms) – konfigurowalne z config.json ───────────────────────────
D["afterLaunch"]   := Task.Has("delayAfterLaunch")   ? Task["delayAfterLaunch"]   : 5000
D["afterClick"]    := Task.Has("delayAfterClick")    ? Task["delayAfterClick"]    : 300
D["afterType"]     := Task.Has("delayAfterType")     ? Task["delayAfterType"]     : 200
D["afterEnter"]    := Task.Has("delayAfterEnter")    ? Task["delayAfterEnter"]    : 1000
D["afterSupplier"] := Task.Has("delayAfterSupplier") ? Task["delayAfterSupplier"] : 2500
D["afterF3"]       := Task.Has("delayAfterF3")       ? Task["delayAfterF3"]       : 1000
D["afterF7"]       := Task.Has("delayAfterF7")       ? Task["delayAfterF7"]       : 1000
D["afterF12"]      := Task.Has("delayAfterF12")      ? Task["delayAfterF12"]      : 500
D["afterSave"]     := Task.Has("delayAfterSave")     ? Task["delayAfterSave"]     : 1500
D["searchWait"]    := Task.Has("delaySearchWait")    ? Task["delaySearchWait"]    : 800

LogMsg("Opóźnienia (ms):")
for k, v in D
    LogMsg("  " k " = " v)

; ── DANE ZADANIA ─────────────────────────────────────────────────────────────
invoiceNr      := Task["invoiceNr"]
usdRate        := Task["usdRate"]
discount       := Task.Has("discount") ? Task["discount"] : 8
items          := Task["items"]
exePath        := Task.Has("exePath") ? Task["exePath"] : ""
supplierSearch := Task.Has("supplierSearch") ? Task["supplierSearch"] : "levior"

LogSep()
LogMsg("ZADANIE:")
LogMsg("  Faktura nr    : " invoiceNr)
LogMsg("  Dostawca      : " supplierSearch)
LogMsg("  Kurs USD/PLN  : " usdRate)
LogMsg("  Rabat         : " discount "%")
LogMsg("  Pozycji       : " items.Length)
LogMsg("  Exe           : " exePath)
LogSep()

; ── GŁÓWNA PROCEDURA ─────────────────────────────────────────────────────────
skipLaunch := Task.Has("skipLaunch") && Task["skipLaunch"]
LogMsg("Tryb: " (skipLaunch ? "AHK GUI (skipLaunch=true)" : "normalny (auto-launch)"))

try {

    ; =========================================================================
    ; KROK 1 – Uruchom / aktywuj FakturaF.exe
    ; =========================================================================
    LogMsg("[KROK 1] Sprawdzam FakturaF.exe...")
    if skipLaunch {
        ; Tryb AHK GUI – użytkownik sam otworzył FakturaF, czekamy max 60s
        LogMsg("  [skipLaunch] Czekam na FakturaF.exe otwartego ręcznie (max 60s)...")
        if !WinWait("ahk_exe FakturaF.exe",, 60) {
            LogMsg("  BŁĄD: FakturaF.exe nie pojawił się w 60s – otwórz program ręcznie!")
            WriteResult(false, "FakturaF.exe nie uruchomiony – otwórz ręcznie i uruchom skrypt ponownie")
            ExitApp(1)
        }
        LogMsg("  [skipLaunch] FakturaF.exe wykryty")
    } else {
        if !WinExist("ahk_exe FakturaF.exe") {
            LogMsg("  FakturaF.exe nie działa – próbuję uruchomić")
            if (exePath != "" && FileExist(exePath)) {
                LogMsg("  Run: " exePath)
                Run(exePath)
                LogMsg("  Czekam na FakturaF.exe (max 30s)...")
                if !WinWait("ahk_exe FakturaF.exe",, 30) {
                    LogMsg("  BŁĄD: FakturaF.exe nie pojawił się w 30s")
                    WriteResult(false, "FakturaF.exe nie uruchomiony w 30s")
                    ExitApp(1)
                }
                LogMsg("  Wykryto FakturaF.exe – czekam " D["afterLaunch"] "ms na inicjalizację")
                Sleep(D["afterLaunch"])
            } else {
                LogMsg("  BŁĄD: exe='" exePath "' – brak pliku lub ścieżki")
                WriteResult(false, "FakturaF.exe nie działa i brak ścieżki EXE")
                ExitApp(1)
            }
        } else {
            LogMsg("  FakturaF.exe już działa")
        }
    }
    LogMsg("[KROK 1] WinActivate...")
    WinActivate("ahk_exe FakturaF.exe")
    WinWaitActive("ahk_exe FakturaF.exe",, 15)
    Sleep(500)
    LogMsg("[KROK 1] OK – okno FakturaF aktywne")

    ; =========================================================================
    ; KROK 2 – Nowy dokument (79, 322)
    ; =========================================================================
    LogMsg("[KROK 2] Kliknij Nowy dokument (79, 322)")
    ClickAt(79, 322)
    LogMsg("[KROK 2] OK")

    ; =========================================================================
    ; KROK 3 – Pole dostawcy: wpisz supplierSearch → ↓ → Enter
    ; =========================================================================
    LogMsg("[KROK 3] Pole dostawcy (318, 169) – wpisuję: '" supplierSearch "'")
    ClickAt(318, 169)
    Send("^a")
    Sleep(100)
    LogMsg("[KROK 3] Send: '" supplierSearch "'")
    Send(supplierSearch)
    Sleep(D["afterType"])
    LogMsg("[KROK 3] Send: {Down}")
    Send("{Down}")
    Sleep(500)
    LogMsg("[KROK 3] Send: {Enter}")
    Send("{Enter}")
    LogMsg("[KROK 3] Czekam " D["afterSupplier"] "ms na załadowanie danych dostawcy...")
    Sleep(D["afterSupplier"])
    LogMsg("[KROK 3] OK – dostawca wybrany")

    ; =========================================================================
    ; KROK 4 – Nr faktury (kursor powinien być na polu po Enter z dostawcy)
    ; =========================================================================
    LogMsg("[KROK 4] Wpisuję nr faktury: '" invoiceNr "'")
    Send(invoiceNr)
    Sleep(D["afterType"])
    LogMsg("[KROK 4] Send: {Tab}")
    Send("{Tab}")
    Sleep(D["afterClick"])
    LogMsg("[KROK 4] OK")

    ; =========================================================================
    ; KROK 5 – Klik (170, 450)
    ; =========================================================================
    LogMsg("[KROK 5] Kliknij (170, 450)")
    ClickAt(170, 450)
    LogMsg("[KROK 5] OK")

    ; =========================================================================
    ; KROK 6 – Klik (813, 410)
    ; =========================================================================
    LogMsg("[KROK 6] Kliknij (813, 410) – otwieranie okna waluty")
    ClickAt(813, 410)
    LogMsg("[KROK 6] OK")

    ; =========================================================================
    ; KROK 7 – Klik (864,441) i (881,445)
    ; =========================================================================
    LogMsg("[KROK 7] Kliknij (864, 441) – kod waluty USD")
    ClickAt(864, 441)
    LogMsg("[KROK 7] Kliknij (881, 445) – zatwierdź walutę")
    ClickAt(881, 445)
    LogMsg("[KROK 7] OK")

    ; =========================================================================
    ; KROK 8 – Kurs USD: klik (880, 472) → wpisz kurs
    ; =========================================================================
    rateStr := FormatPrice4(usdRate)
    LogMsg("[KROK 8] Kurs USD=" usdRate " → rateStr='" rateStr "'")
    LogMsg("[KROK 8] Kliknij (880, 472)")
    ClickAt(880, 472)
    Send("^a")
    Sleep(100)
    LogMsg("[KROK 8] Send: '" rateStr "'")
    Send(rateStr)
    Sleep(D["afterType"])
    LogMsg("[KROK 8] Send: {Tab}")
    Send("{Tab}")
    Sleep(D["afterClick"])
    LogMsg("[KROK 8] OK")

    ; =========================================================================
    ; KROK 9 – Zakładka pozycji (73, 322)
    ; =========================================================================
    LogMsg("[KROK 9] Kliknij zakładkę Pozycje (73, 322)")
    ClickAt(73, 322)
    LogMsg("[KROK 9] OK")

    ; =========================================================================
    ; KROK 10 – F7 (Dodaj z kartoteki)
    ; =========================================================================
    LogMsg("[KROK 10] Send: {F7} (Dodaj z kartoteki)")
    Send("{F7}")
    Sleep(D["afterF7"])
    WinActivate("ahk_exe FakturaF.exe")
    Sleep(300)
    LogMsg("[KROK 10] OK")

    ; =========================================================================
    ; KROKI 11–15+16 – PĘTLA POZYCJI
    ; =========================================================================
    LogSep()
    LogMsg("[PĘTLA] START – " items.Length " pozycji do wpisania")
    LogSep()

    Loop items.Length {
        item     := items[A_Index]
        kod5     := SubStr(item["kod"], 1, 5)
        nazwa    := item["nazwa"]
        ilosc    := item["ilosc"]
        priceUSD := item.Has("priceUSD") ? item["priceUSD"] : 0.0
        invPLN   := Round(priceUSD * usdRate, 2)
        qtyStr   := FormatQty(ilosc)

        LogSep()
        LogMsg("[POZYCJA " A_Index "/" items.Length "]")
        LogMsg("  kod5     = '" kod5 "'")
        LogMsg("  nazwa    = '" SubStr(nazwa, 1, 60) "'")
        LogMsg("  ilosc    = " ilosc " → '" qtyStr "'")
        LogMsg("  priceUSD = " priceUSD)
        LogMsg("  invPLN   = " invPLN " (= " priceUSD " × " usdRate " zaokr. do 2dp)")

        ; --- KROK 11: F3 → kod (5 cyfr) → Enter (1. szukaj) -----------------
        LogMsg("  [F3] Otwieram wyszukiwarkę kartoteki...")
        Send("{F3}")
        Sleep(D["afterF3"])

        LogMsg("  [TYPE] Wpisuję kod: '" kod5 "'")
        Send(kod5)
        Sleep(D["afterType"])

        LogMsg("  [ENTER-1] Uruchamiam wyszukiwanie...")
        Send("{Enter}")
        LogMsg("  Czekam " D["searchWait"] "ms na wyniki...")
        Sleep(D["searchWait"])

        ; --- Detekcja produktu -----------------------------------------------
        productFound := CheckProductFound(kod5)
        LogMsg("  [DETECT] Produkt w katalogu: " (productFound ? "TAK ✓" : "NIE ✗"))

        if !productFound {
            ; -----------------------------------------------------------------
            ; KROK 16 – Nowy produkt (F6)
            ; -----------------------------------------------------------------
            LogMsg("  [KROK 16] → F6: dodaję nowy produkt")
            Send("{Escape}")
            Sleep(300)
            HandleNewProduct(kod5, nazwa, ilosc, invPLN)
            g_New++
            AddResult(kod5, nazwa, ilosc, true, "nowy produkt (F6)")
            LogMsg("  [KROK 16] OK – nowy produkt dodany")
            continue
        }

        ; --- ENTER-2: potwierdź wybór z listy --------------------------------
        LogMsg("  [ENTER-2] Potwierdzam wybór produktu z listy")
        Send("{Enter}")
        Sleep(D["afterEnter"])

        ; --- KROK 12: Wpisz ilość --------------------------------------------
        LogMsg("  [KROK 12] Wpisuję ilość: " ilosc " → '" qtyStr "'")
        Send(qtyStr)
        Sleep(D["afterType"])

        ; Przejdź do pola ceny (Tab)
        LogMsg("  [TAB] Przechodzę do pola ceny")
        Send("{Tab}")
        Sleep(400)

        ; --- KROK 13: Sprawdź i ewentualnie popraw cenę ----------------------
        sysPriceStr := ReadFieldViaClipboard()
        sysPrice    := 0.0
        try
            sysPrice := Float(StrReplace(StrReplace(sysPriceStr, " ", ""), ",", "."))
        catch
            sysPrice := 0.0

        LogMsg("  [KROK 13] Cena systemowa raw='" sysPriceStr "' → " sysPrice)
        LogMsg("  [KROK 13] Cena faktury   PLN=" invPLN)
        LogMsg("  [KROK 13] Różnica: " Abs(sysPrice - invPLN))

        if (Abs(sysPrice - invPLN) > 0.005) {
            newPriceStr := FormatPrice2(invPLN)
            LogMsg("  [KROK 13] KOREKTA CENY: " sysPrice " → " invPLN " ('" newPriceStr "')")
            Send("^a")
            Send(newPriceStr)
            Sleep(D["afterType"])
        } else {
            LogMsg("  [KROK 13] Cena OK – zgodna z fakturą (" sysPrice ")")
        }
        Send("{Enter}")
        Sleep(300)

        ; --- KROK 14: F12 ----------------------------------------------------
        LogMsg("  [KROK 14] Send: {F12} – zatwierdzam pozycję")
        Send("{F12}")
        Sleep(D["afterF12"])

        g_Added++
        AddResult(kod5, nazwa, ilosc, true, "")
        LogMsg("  [OK] Pozycja dodana: kod=" kod5 " qty=" ilosc " PLN=" invPLN)
    }

    ; =========================================================================
    ; KROK 17 – Podsumowanie i rabat
    ; =========================================================================
    LogSep()
    LogMsg("[KROK 17] Przechodzę do podsumowania i wpisuję rabat " discount "%")
    LogMsg("[KROK 17a] Kliknij (83, 275)")
    ClickAt(83, 275)
    LogMsg("[KROK 17b] Kliknij (741, 380)")
    ClickAt(741, 380)
    LogMsg("[KROK 17c] Kliknij pole rabatu (892, 412)")
    ClickAt(892, 412)
    Send("^a")
    Sleep(100)
    ; BUG7 FIX: zawsze wysyłamy jako string bez części ułamkowej
    discountStr := String(Integer(Round(discount)))
    LogMsg("[KROK 17c] Wpisuję rabat: " discountStr)
    Send(discountStr)
    Sleep(D["afterType"])
    LogMsg("[KROK 17d] Kliknij OK rabatu (933, 410)")
    ClickAt(933, 410)
    LogMsg("[KROK 17] OK – rabat wpisany")

    ; =========================================================================
    ; KROK 21 – F12 (zapis końcowy)
    ; =========================================================================
    LogMsg("[KROK 21] Send: {F12} – zapis końcowy faktury")
    Send("{F12}")
    Sleep(D["afterSave"])

    elapsed := Round((A_TickCount - g_StartTime) / 1000)
    LogSep()
    LogMsg("=== BOT ZAKOŃCZONY SUKCESEM ===")
    LogMsg("  Dodano (z katalogu): " g_Added)
    LogMsg("  Dodano (nowe F6)   : " g_New)
    LogMsg("  Błędy              : " g_Errors)
    LogMsg("  Czas               : " elapsed "s")
    LogSep()
    WriteResult(true, "")
} catch Error as e {
    LogMsg("BŁĄD KRYTYCZNY: " e.Message)
    LogMsg("  Stack: " e.Stack)
    g_Errors++
    WriteResult(false, e.Message)
    ExitApp(1)
}

ExitApp(0)

; ============================================================================
; FUNKCJE
; ============================================================================

; ── DETEKCJA PRODUKTU W KATALOGU ─────────────────────────────────────────────
;
; Po F3 + kod + 1. Enter:
;   - Pole wyszukiwania (góra okna) zawiera kod → 1 wystąpienie
;   - Wiersz siatki z wynikiem zawiera kod → 2. wystąpienie
;   - Jeśli 2+ wystąpień: produkt JEST w katalogu
;   - Jeśli tylko 1: NIE MA w katalogu (pusta lista, jak na screenshocie)
; ─────────────────────────────────────────────────────────────────────────────
CheckProductFound(kod5) {
    global D

    ; Metoda 1: WinGetText – zlicz wystąpienia kodu
    wText := WinGetText("ahk_exe FakturaF.exe")
    count := 0
    pos   := 1
    while (foundPos := InStr(wText, kod5,, pos)) {
        count++
        pos := foundPos + 1
    }
    LogMsg("    [DETECT M1] WinGetText: kod5='" kod5 "' wystąpień=" count
           " (len=" StrLen(wText) ")")

    if (count >= 2) {
        LogMsg("    [DETECT M1] → ZNALEZIONO (>= 2 wystąpień)")
        return true
    }

    ; Metoda 2: ControlGetText na typowych kontrolkach siatki
    gridClasses := ["TAdvStringGrid", "TStringGrid", "TDBGrid",
                    "TDrawGrid", "TCustomGrid", "TListView", "TDBAdvGrid"]
    for idx, cls in gridClasses {
        try {
            ctrlName := cls "1"
            text := ControlGetText(ctrlName, "ahk_exe FakturaF.exe")
            if StrLen(text) > 0 {
                LogMsg("    [DETECT M2] ControlGetText(" ctrlName ") len="
                       StrLen(text) " start='" SubStr(text, 1, 40) "'")
                if InStr(text, kod5) {
                    LogMsg("    [DETECT M2] → ZNALEZIONO (kod w kontrolce)")
                    return true
                } else {
                    LogMsg("    [DETECT M2] → NIE ZNALEZIONO (brak kodu w kontrolce)")
                    return false
                }
            }
        }
    }

    ; Fallback – konserwatywny: zakładamy że NIE MA
    ; (bezpieczniejsze: F6 tworzy nowy produkt zamiast pominąć istniejący)
    LogMsg("    [DETECT FALLBACK] Brak danych z żadnej metody – przyjmuję: NIE ZNALEZIONO")
    return false
}

; ── KROK 16: NOWY PRODUKT (F6) ───────────────────────────────────────────────
HandleNewProduct(kod5, nazwa, ilosc, invPLN) {
    global D

    LogMsg("    [F6] Otwieram formularz nowego produktu")
    Send("{F6}")
    Sleep(800)

    ; Tłumaczenie nazwy na Polski
    LogMsg("    [TRANSLATE] Tłumaczę: '" SubStr(nazwa, 1, 60) "'")
    polishName := TranslateToPolish(nazwa)
    LogMsg("    [TRANSLATE] Wynik   : '" SubStr(polishName, 1, 60) "'")

    ; Wpisz kod
    LogMsg("    [F6] Wpisuję kod: '" kod5 "'")
    Send("^a")
    Sleep(100)
    Send(kod5)
    Sleep(D["afterType"])
    Send("{Tab}")
    Sleep(300)

    ; Wpisz przetłumaczoną nazwę
    LogMsg("    [F6] Wpisuję nazwę PL: '" SubStr(polishName, 1, 60) "'")
    Send("^a")
    Sleep(100)
    Send(polishName)
    Sleep(D["afterType"])

    ; Ilość – klik (694, 543)
    LogMsg("    [F6] Kliknij pole ilości (694, 543)")
    ClickAt(694, 543)
    Send("^a")
    Sleep(100)
    qtyStr := FormatQty(ilosc)
    LogMsg("    [F6] Wpisuję ilość: " ilosc " → '" qtyStr "'")
    Send(qtyStr)
    Sleep(D["afterType"])
    Send("{Enter}")
    Sleep(300)

    ; Cena
    priceStr := FormatPrice2(invPLN)
    LogMsg("    [F6] Wpisuję cenę: " invPLN " → '" priceStr "'")
    Send("^a")
    Sleep(100)
    Send(priceStr)
    Sleep(D["afterType"])
    Send("{Enter}")
    Sleep(300)

    ; Zatwierdź pozycję
    LogMsg("    [F6] {F12} – zatwierdzam nową pozycję")
    Send("{F12}")
    Sleep(D["afterF12"])
}

; ── TŁUMACZENIE VIA FLASK API ─────────────────────────────────────────────────
TranslateToPolish(text) {
    url  := "http://127.0.0.1:5000/api/translate"
    body := '{"text":' JSON.stringify(text) ',"to":"pl"}'

    LogMsg("    [HTTP] POST " url)
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open("POST", url, false)
        whr.SetRequestHeader("Content-Type", "application/json")
        whr.SetTimeouts(3000, 3000, 10000, 10000)
        whr.Send(body)
        LogMsg("    [HTTP] Status: " whr.Status)
        if (whr.Status = 200) {
            resp := JSON.parse(whr.ResponseText)
            if resp.Has("translated") {
                LogMsg("    [HTTP] Odpowiedź OK: '" SubStr(resp["translated"], 1, 60) "'")
                return resp["translated"]
            }
        }
        LogMsg("    [HTTP] Nieoczekiwana odpowiedź: " SubStr(whr.ResponseText, 1, 100))
    }
    catch Error as e {
        LogMsg("    [HTTP] Błąd tłumaczenia: " e.Message)
    }
    LogMsg("    [HTTP] Fallback – używam oryginału")
    return text
}

; ── KLIKNIJ ABSOLUTNE KOORDYNATY + LOG ───────────────────────────────────────
ClickAt(x, y) {
    global D
    LogMsg("    [CLICK] (" x ", " y ")")
    Click(x, y)
    Sleep(D["afterClick"])
}

; ── ODCZYTAJ WARTOŚĆ POLA PRZEZ SCHOWEK ──────────────────────────────────────
ReadFieldViaClipboard() {
    prevClip    := A_Clipboard
    A_Clipboard := ""
    Send("^a^c")
    ; BUG8 FIX: sprawdzamy czy clipboard się faktycznie zmienił (ClipWait zwraca 0 = timeout)
    ok := ClipWait(2.0)
    if !ok
        LogMsg("    [CLIP] UWAGA: timeout czekania na clipboard – cena może być błędna")
    result      := A_Clipboard
    A_Clipboard := prevClip
    LogMsg("    [CLIP] Odczytano: '" result "'" (ok ? "" : " [TIMEOUT]"))
    return result
}

; ── FORMATOWANIE ILOŚCI ───────────────────────────────────────────────────────
FormatQty(ilosc) {
    if (Mod(ilosc, 1) = 0)
        return Integer(ilosc)
    return StrReplace(Format("{:.3f}", ilosc), ".", ",")
}

; ── FORMATOWANIE CENY (2 dp, przecinek) ───────────────────────────────────────
FormatPrice2(price) {
    return StrReplace(Format("{:.2f}", price), ".", ",")
}

; ── FORMATOWANIE KURSU (4 dp, przecinek) ─────────────────────────────────────
FormatPrice4(val) {
    return StrReplace(Format("{:.4f}", val), ".", ",")
}

; ── DODAJ WYNIK POZYCJI ───────────────────────────────────────────────────────
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

; ── ZAPISZ RESULT.JSON ────────────────────────────────────────────────────────
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
    LogMsg("Zapisano result.json: " items.Length " pozycji, success=" (success ? "true" : "false"))
}

; ── LOGOWANIE ─────────────────────────────────────────────────────────────────
LogMsg(msg) {
    global AhkLog
    ts := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    line := ts " [AHK] " msg
    AhkLog.WriteLine(line)
    OutputDebug(line)
}

LogSep() {
    LogMsg("─────────────────────────────────────────────────────────")
}

; ============================================================================
; KLASY POMOCNICZE
; ============================================================================

class JsonBool {
    __New(v) => this.val := (v ? true : false)
}

; ── Minimalna implementacja JSON dla AHK v2 ──────────────────────────────────
; BUG2 FIX: _parseNumber sprawdza puste n
; BUG3 FIX: _parseObject/Array mają guard p > StrLen(s) – brak nieskończonej pętli
; BUG5 FIX: _parseString ternary zastąpiony if-else (brak problemów z kontynuacją linii)
; BUG6 FIX: _parseString obsługuje \uXXXX (skip 4 hex cyfr)
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
        if p > StrLen(s)
            return ""
        c := SubStr(s, p, 1)
        if (c = '"')
            return JSON._parseString(s, &p)
        if (c = '{')
            return JSON._parseObject(s, &p)
        if (c = '[')
            return JSON._parseArray(s, &p)
        if (c = 't') {
            p += 4
            return true
        }
        if (c = 'f') {
            p += 5
            return false
        }
        if (c = 'n') {
            p += 4
            return ""
        }
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
            if (c = '"') {
                p++
                return result
            }
            if (c = '\') {
                p++
                ec := SubStr(s, p, 1)
                ; BUG5 FIX: if-else zamiast wieloliniowego ternary
                ; BUG6 FIX: \uXXXX – pomijamy 4 cyfry hex, nie dodajemy 'u' do wyniku
                if (ec = 'n')
                    result .= '`n'
                else if (ec = 't')
                    result .= '`t'
                else if (ec = 'r')
                    result .= '`r'
                else if (ec = 'u') {
                    p += 4  ; pomiń 4 cyfry hex (\uXXXX)
                    ; (pomijamy znak Unicode – brak wsparcia w AHK string)
                }
                else if (ec = '"' || ec = '\' || ec = '/')
                    result .= ec
                else
                    result .= ec
            } else {
                result .= c
            }
            p++
        }
        return result
    }

    static _parseNumber(s, &p) {
        start := p
        ; Obsłuż znak minus na początku
        if (SubStr(s, p, 1) = '-')
            p++
        while (p <= StrLen(s) && InStr("0123456789.eE+", SubStr(s, p, 1)))
            p++
        n := SubStr(s, start, p - start)
        ; BUG2 FIX: puste n (błędny JSON) → zwróć 0 zamiast crash
        if StrLen(n) = 0
            return 0
        try {
            return InStr(n, '.') || InStr(n, 'e') || InStr(n, 'E') ? Float(n) : Integer(n)
        } catch {
            return 0
        }
    }

    static _parseObject(s, &p) {
        p++
        obj := Map()
        JSON._skipWS(s, &p)
        if (SubStr(s, p, 1) = '}') {
            p++
            return obj
        }
        ; BUG3 FIX: guard na koniec stringa – zapobiega nieskończonej pętli
        loop {
            if p > StrLen(s)
                break
            JSON._skipWS(s, &p)
            if p > StrLen(s)
                break
            key := JSON._parseString(s, &p)
            JSON._skipWS(s, &p)
            p++   ; ':'
            val := JSON._parseValue(s, &p)
            obj[key] := val
            JSON._skipWS(s, &p)
            if p > StrLen(s)
                break
            c := SubStr(s, p, 1)
            p++
            if (c = '}')
                return obj
            ; c = ',' → kontynuuj, inne znaki → przerywamy (błędny JSON)
            if (c != ',')
                break
        }
        return obj
    }

    static _parseArray(s, &p) {
        p++
        arr := []
        JSON._skipWS(s, &p)
        if (SubStr(s, p, 1) = ']') {
            p++
            return arr
        }
        ; BUG4 FIX: guard na koniec stringa – zapobiega nieskończonej pętli
        loop {
            if p > StrLen(s)
                break
            arr.Push(JSON._parseValue(s, &p))
            JSON._skipWS(s, &p)
            if p > StrLen(s)
                break
            c := SubStr(s, p, 1)
            p++
            if (c = ']')
                return arr
            ; c = ',' → kontynuuj, inne znaki → przerywamy (błędny JSON)
            if (c != ',')
                break
        }
        return arr
    }

    static _serializeValue(val) {
        if (Type(val) = "JsonBool")
            return (val.val ? "true" : "false")
        t := Type(val)
        if (t = "String")
            return '"' StrReplace(StrReplace(StrReplace(val, '\', '\\'), '"', '\"'), '`n', '\n') '"'
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
            if (parts != "") parts .= ","
            parts .= '"' k '":' JSON._serializeValue(v)
        }
        return "{" parts "}"
    }

    static _serializeArray(arr) {
        if (arr.Length = 0)
            return "[]"
        parts := ""
        for v in arr {
            if (parts != "") parts .= ","
            parts .= JSON._serializeValue(v)
        }
        return "[" parts "]"
    }
}
