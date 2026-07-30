#Requires AutoHotkey v2.0
#SingleInstance Force

; =====================================================================
; =====================================================================
;  NETWORK PHARMACIST AHK — V2.3.2
; =====================================================================
;
;  Merged automation script combining:
;    - PDMP Helper v2.2.5
;    - Refusal to Fill v2.2.0
;    - Tachyon Ticket v2.3.10
;    - NDC OOS Ticket v2.0.1
;
;  HOTKEYS:
;    Ctrl+Shift+Alt+B : PDMP Batch Loop
;    Ctrl+Shift+Alt+D : PDMP Duplicate Tab
;    Ctrl+Shift+Alt+R : Refusal to Fill
;    Ctrl+Shift+Alt+T : Tachyon Ticket
;    Ctrl+Shift+Alt+O : NDC OOS Ticket
;    F12              : Emergency Stop (wipes all PHI, reloads)
;
; =====================================================================
; =====================================================================

SetTitleMatchMode(2)
SetWinDelay(10)
SendMode("Input")        ; SendInput — fastest mode; default in v2 but pinned for clarity
SetKeyDelay(-1, -1)      ; No artificial delay (only affects SendEvent if it ever runs)

global g_Version := "2.3.2"
A_IconTip := "Network Pharmacist AHK v" . g_Version

; =====================================================================
;  HIPAA PRE-FLIGHT: Refuse to run if Windows Clipboard History or
;  Cloud Clipboard is enabled. Both features sync clipboard contents
;  to Microsoft cloud. This script copies entire page-scrapes (which
;  include PHI) to A_Clipboard during normal operation, so they MUST
;  be disabled before the script runs.
; =====================================================================
CheckClipboardHistory() {
    BlockReason := ""
    Try {
        HistVal := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Clipboard", "EnableClipboardHistory", 0)
        If (HistVal == 1)
            BlockReason := "Windows Clipboard History is ENABLED."
    } Catch {
        ; Key absent -> feature off by default. Continue.
    }
    Try {
        CloudVal := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Clipboard", "CloudClipboardAutomaticUpload", 0)
        If (CloudVal == 1) {
            If (BlockReason != "")
                BlockReason .= "`n"
            BlockReason .= "Windows Cloud Clipboard sync is ENABLED."
        }
    } Catch {
    }

    If (BlockReason != "") {
        Msg := "HIPAA pre-flight check FAILED.`n`n" . BlockReason . "`n`n"
        Msg .= "These features sync clipboard contents to Microsoft cloud servers. "
        Msg .= "This script copies entire pages (including PHI) to the clipboard "
        Msg .= "during normal operation, so they MUST be disabled.`n`n"
        Msg .= "How to disable:`n"
        Msg .= "  Settings -> System -> Clipboard`n"
        Msg .= "    - Turn OFF 'Clipboard history'`n"
        Msg .= "    - Turn OFF 'Sync across your devices'`n`n"
        Msg .= "After disabling, restart this script.`n`n"
        Msg .= "Script will now exit."
        MsgBox(Msg, "HIPAA Pre-Flight - Script Will Not Start", 16)
        Msg := ""
        ExitApp()
    }
}
CheckClipboardHistory()

; =====================================================================
;  GLOBAL VARIABLES — PDMP Helper
; =====================================================================
global p_FirstName    := ""
global p_LastName     := ""
global p_DOB          := ""
global p_StateName    := ""
global g_BatchRunning := false
global g_PostWipeClipboard := ""

; =====================================================================
;  GLOBAL VARIABLES — Refusal to Fill (core)
; =====================================================================
global p_PatientFullName := ""
global p_MedName         := ""
global p_Batch           := ""
global p_Identifier      := ""
global p_OrderID         := ""
global p_LastFillDate    := ""
global p_LastFillDays    := ""
global p_LastPharmacy    := ""
global p_NextFillDate    := ""
global g_FormSubmitted   := false

; =====================================================================
;  GLOBAL VARIABLES — Refusal to Fill (extended / non-Amazon path)
; =====================================================================
global p_ShipDate            := ""
global p_ArrivalDate         := ""
global p_ArrivalParsed       := ""
global p_ShipmentURL         := ""
global p_PillPackProfileID   := ""
global p_SelectedNDC         := ""
global p_SelectedRxNumber    := ""
global p_AdminURL            := ""
global p_CustomerPhone       := ""
global p_RTFReason           := ""
global p_RTFReasonDetail     := ""
global g_DetailVisible       := false
global g_RTFIsEarlyControl   := false
global g_RTFDecision         := ""
global g_RTFReasonOptions    := ["Early Control", "Drug of Concern", "Quantity/Day Supply Mismatch", "Directions not Appropriate", "DNF"]

; Med list and GUI selection state (Refusal)
global g_RefusalMedList      := []
global g_RefusalSelectedIdx  := 0
global g_RefusalGUIDropCount := 1
global g_RefusalGUIAction    := ""
global g_RefusalGUISaved     := []
global g_RefusalGUIDDLs      := []

; =====================================================================
;  GLOBAL VARIABLES — Tachyon Ticket
; =====================================================================
global g_RxNumber := ""
global g_GroupID := ""
global g_PatientName := ""
global g_Batch := ""
global g_PersonID := ""
global g_PrescriptionType := ""
global g_ShipDate := ""
global g_ArrivalDate := ""
global g_ArrivalParsed := ""
global g_ReasonForChange := ""
global g_ReasonIndex := 0
global g_MedicationName := ""
global g_ChangeNeeded := ""
global g_ShipmentURL := ""

; Tachyon Ticket — Blueprint routing & non-Amazon (UFR) flow
global g_BlueprintName := ""
global g_IsAmazon := false
global g_AdminURL := ""
global g_PillPackProfileID := ""
global g_UFRReason := ""
global g_UFRReasonIndex := 0
global g_StartedOnShipment := false
global g_IsRedShipment := false

; Tachyon Ticket — Shipment-start med selection state
global g_ShipMedList         := []
global g_ShipSelectedIdx     := 0
global g_ShipGUIDropCount    := 1
global g_ShipGUIAction       := ""
global g_ShipGUISaved        := []
global g_ShipGUIDDLs         := []

; =====================================================================
;  GLOBAL VARIABLES — Tachyon Ticket dropdown options
; =====================================================================
global g_ReasonOptions := [
    "NDC Out of Stock",
    "Directions Change",
    "Quantity/Package size mismatch",
    "Days Supply Mismatch",
    "Clinical Review Requested",
    "Address Validation Required",
    "Duplicative dispense for one order",
    "Shipment Hold Review",
    "Delivery Shipping Day Missed Promise",
    "Missed Delivery OR Wrong Ship Option"
]

global g_PrescriptionOptions := [
    "Bulk",
    "Packet"
]

global g_UFROptions := [
    "Follow up with Customer",
    "Follow up with Pharmacy",
    "Time Change Request (HOA) Due to Packager Limitation",
    "Customer Demographic Spelling Errror",
    "Duplicate Prescription",
    "Follow Up with Prescriber",
    "Incorrect Directions",
    "Quantity/Day Supply Mismatch",
    "Shipment Hold Review",
    "Time Change Request (HOA)",
    "Unit of Measure",
    "Update Quantity",
    "Shipping Block",
    "Prugen Out of stock"
]

; =====================================================================
;  GLOBAL VARIABLES — NDC OOS Ticket
; =====================================================================
global g_SelectedNDC            := ""
global g_SelectedMedName        := ""
global g_SelectedRxNumber       := ""
global g_MedList                := []
global g_SelectedMedIdx         := 0
global g_GUIDropCount           := 1
global g_GUIAction              := ""
global g_GUISavedSelections     := []
global g_GUIDDLs                := []

; =====================================================================
;  GLOBAL VARIABLES — HIPAA Watchdog
; =====================================================================
global g_ScriptStartTime := 0
global g_MaxRuntimeMs    := 7 * 60 * 1000   ; 7-minute hard cap
global g_PauseStartTime  := 0               ; 0 = not currently paused
global g_MaxIdleMs       := 5 * 60 * 1000   ; 5-minute idle cap at dialogs

; =====================================================================
;  ON-EXIT / ERROR / SHUTDOWN HANDLERS
;  CombinedExit covers all workflows' cleanup.
; =====================================================================
OnExit(CombinedExit)
OnError(GlobalErrorHandler)
OnMessage(0x11, ShutdownHandler)     ; WM_QUERYENDSESSION
OnMessage(0x16, ShutdownHandler)     ; WM_ENDSESSION

CombinedExit(ExitReason, ExitCode) {
    global g_PostWipeClipboard
    StopWatchdog()
    savedClip := g_PostWipeClipboard
    WipeAllData()
    if (savedClip != "") {
        A_Clipboard := savedClip
        g_PostWipeClipboard := ""
    }
    Return 0
}

GlobalErrorHandler(Exception, Mode) {
    Try StopWatchdog()
    Try WipeAllData()
    Try CloseAllKnownDialogs()
    Return 0
}

ShutdownHandler(wParam, lParam, msg, hwnd) {
    Try StopWatchdog()
    Try WipeAllData()
    Return 1
}

; =====================================================================
;  FUNCTION: GUI X-button crash handler. Immediately wipes all data
;  and terminates the script when the user closes any GUI via X.
; =====================================================================
GUICrash(guiObj, *) {
    Try guiObj.Destroy()
    StopWatchdog()
    WipeAllData()
    Reload()
}

; =====================================================================
;  FUNCTION: Close every dialog ANY workflow is known to open.
;  Called by the watchdog, GlobalErrorHandler, and ShutdownHandler.
;  Failures are silently ignored — best-effort cleanup.
; =====================================================================
CloseAllKnownDialogs() {
    ; --- Refusal to Fill dialogs ---
    Try WinClose("Refusal to Fill - Input Required")
    Try WinClose("Refusal to Fill - Medication Selection")
    Try WinClose("Refusal to Fill - Complete")
    Try WinClose("RTF - Non-Early Control Decision")
    Try WinClose("RTF - Non-Amazon Decision")
    Try WinClose("Refusal to Fill - Callback Complete")
    Try WinClose("Template Ready - Copied to Clipboard")
    ; --- Tachyon Ticket dialogs ---
    Try WinClose("Ticket Review")
    Try WinClose("Clinical Review - Input")
    Try WinClose("Upstream Fix Request - Input")
    Try WinClose("Calendar Skipped")
    Try WinClose("Blueprint Click Failed")
    Try WinClose("Customer Selection Warning")
    Try WinClose("Tachyon Ticket - Medication Selection")
    Try WinClose("Multiple Medications Detected")
    ; --- NDC OOS Ticket dialogs ---
    Try WinClose("NDC OOS Ticket - Input")
    Try WinClose("Complete - Verify Before Submit")
    ; --- Shared / common dialogs ---
    Try WinClose("Validation Error")
    Try WinClose("Non-Amazon Customer")
    Try WinClose("Login Paused")
    Try WinClose("Login Required")
    Try WinClose("Re-login Required")
    Try WinClose("Field Verify Failed")
    Try WinClose("Capture Failed")
    Try WinClose("Extraction Failed")
    Try WinClose("Timeout")
    Try WinClose("Stopping")
    Try WinClose("Error")
    Try WinClose("Warning")
    Try WinClose("Watchdog - Data Wiped")
    Try WinClose("Emergency Stop")
    Try WinClose("HIPAA Pre-Flight - Script Will Not Start")
}

; =====================================================================
;  WATCHDOG CONTROL
; =====================================================================
StartWatchdog() {
    global g_ScriptStartTime, g_PauseStartTime
    g_ScriptStartTime := A_TickCount
    g_PauseStartTime  := 0
    SetTimer(WatchdogCheck, 5000)
}

StopWatchdog() {
    global g_ScriptStartTime, g_PauseStartTime
    g_ScriptStartTime := 0
    g_PauseStartTime  := 0
    SetTimer(WatchdogCheck, 0)
}

PauseStart() {
    global g_PauseStartTime
    If (g_PauseStartTime == 0)
        g_PauseStartTime := A_TickCount
}

PauseEnd() {
    global g_PauseStartTime
    g_PauseStartTime := 0
}

WatchdogCheck() {
    global g_ScriptStartTime, g_MaxRuntimeMs, g_PauseStartTime, g_MaxIdleMs
    If (g_ScriptStartTime == 0)
        Return

    HardExceeded := (A_TickCount - g_ScriptStartTime > g_MaxRuntimeMs)
    IdleExceeded := (g_PauseStartTime > 0) && (A_TickCount - g_PauseStartTime > g_MaxIdleMs)

    If (HardExceeded || IdleExceeded) {
        StopWatchdog()
        WipeAllData()
        CloseAllKnownDialogs()
        Reason := HardExceeded
            ? "script exceeded 7 minutes total runtime"
            : "script sat at a blocking dialog for 5 minutes (idle timeout)"
        MsgBox("Watchdog timeout: " . Reason . ".`n`nAll sensitive data has been wiped for HIPAA compliance. Please re-run the script.", "Watchdog - Data Wiped", 48)
        ExitApp()
    }
}

; =====================================================================
;  FUNCTION: Build a same-length junk string for overwriting PHI
;  buffers before assigning "". Defense-in-depth: displaces the value
;  from the allocator's preferred slot.
; =====================================================================
ScrubString(s) {
    n := StrLen(s)
    If (n == 0)
        Return ""
    junk := ""
    Loop n
        junk .= "X"
    Return junk
}

; =====================================================================
;  FUNCTION: Wipe ALL sensitive data from ALL workflows (HIPAA)
;  Each PHI string is overwritten with same-length junk before being
;  cleared. Clipboard triple-write displaces the prior PHI snapshot.
; =====================================================================
WipeAllData() {
    ; --- PDMP globals ---
    global p_FirstName, p_LastName, p_DOB, p_StateName, g_BatchRunning

    ; --- Refusal to Fill globals ---
    global p_PatientFullName, p_MedName, p_Batch, p_Identifier, p_OrderID
    global p_LastFillDate, p_LastFillDays, p_LastPharmacy, p_NextFillDate, g_FormSubmitted
    global p_ShipDate, p_ArrivalDate, p_ArrivalParsed, p_ShipmentURL
    global p_PillPackProfileID, p_SelectedNDC, p_SelectedRxNumber, p_AdminURL, p_CustomerPhone, p_RTFReason, p_RTFReasonDetail
    global g_DetailVisible, g_RTFIsEarlyControl, g_RTFDecision
    global g_RefusalMedList, g_RefusalSelectedIdx
    global g_RefusalGUIDropCount, g_RefusalGUIAction, g_RefusalGUISaved, g_RefusalGUIDDLs

    ; --- Tachyon Ticket globals ---
    global g_RxNumber, g_GroupID, g_PatientName, g_Batch, g_PersonID
    global g_PrescriptionType, g_ShipDate, g_ArrivalDate, g_ArrivalParsed
    global g_ReasonForChange, g_ReasonIndex
    global g_MedicationName, g_ChangeNeeded, g_ShipmentURL
    global g_BlueprintName, g_IsAmazon, g_AdminURL, g_PillPackProfileID
    global g_UFRReason, g_UFRReasonIndex, g_IsRedShipment, g_StartedOnShipment
    global g_ShipMedList, g_ShipSelectedIdx, g_ShipGUIDropCount
    global g_ShipGUIAction, g_ShipGUISaved, g_ShipGUIDDLs

    ; --- NDC OOS Ticket globals ---
    global g_SelectedNDC, g_SelectedMedName, g_SelectedRxNumber
    global g_MedList, g_SelectedMedIdx
    global g_GUIDropCount, g_GUIAction, g_GUISavedSelections, g_GUIDDLs

    Try StopWatchdog()

    ; === OVERWRITE PHASE: scrub PHI-bearing strings with same-length junk ===

    ; PDMP
    Try p_FirstName        := ScrubString(p_FirstName)
    Try p_LastName         := ScrubString(p_LastName)
    Try p_DOB              := ScrubString(p_DOB)
    Try p_StateName        := ScrubString(p_StateName)

    ; Refusal to Fill
    Try p_PatientFullName  := ScrubString(p_PatientFullName)
    Try p_MedName          := ScrubString(p_MedName)
    Try p_Batch            := ScrubString(p_Batch)
    Try p_Identifier       := ScrubString(p_Identifier)
    Try p_OrderID          := ScrubString(p_OrderID)
    Try p_LastFillDate     := ScrubString(p_LastFillDate)
    Try p_LastFillDays     := ScrubString(p_LastFillDays)
    Try p_LastPharmacy     := ScrubString(p_LastPharmacy)
    Try p_NextFillDate     := ScrubString(p_NextFillDate)
    Try p_ShipDate         := ScrubString(p_ShipDate)
    Try p_ArrivalDate      := ScrubString(p_ArrivalDate)
    If (Type(p_ArrivalParsed) == "Object")
        Try p_ArrivalParsed.formatted := ScrubString(p_ArrivalParsed.formatted)
    Try p_ShipmentURL      := ScrubString(p_ShipmentURL)
    Try p_PillPackProfileID := ScrubString(p_PillPackProfileID)
    Try p_SelectedNDC      := ScrubString(p_SelectedNDC)
    Try p_SelectedRxNumber := ScrubString(p_SelectedRxNumber)
    Try p_AdminURL         := ScrubString(p_AdminURL)
    Try p_CustomerPhone    := ScrubString(p_CustomerPhone)
    Try p_RTFReason        := ScrubString(p_RTFReason)
    Try p_RTFReasonDetail  := ScrubString(p_RTFReasonDetail)

    ; Tachyon Ticket
    Try g_RxNumber         := ScrubString(g_RxNumber)
    Try g_GroupID          := ScrubString(g_GroupID)
    Try g_PatientName      := ScrubString(g_PatientName)
    Try g_Batch            := ScrubString(g_Batch)
    Try g_PersonID         := ScrubString(g_PersonID)
    Try g_ShipDate         := ScrubString(g_ShipDate)
    Try g_ArrivalDate      := ScrubString(g_ArrivalDate)
    If (Type(g_ArrivalParsed) == "Object")
        Try g_ArrivalParsed.formatted := ScrubString(g_ArrivalParsed.formatted)
    Try g_MedicationName   := ScrubString(g_MedicationName)
    Try g_ChangeNeeded     := ScrubString(g_ChangeNeeded)
    Try g_ShipmentURL      := ScrubString(g_ShipmentURL)
    Try g_AdminURL         := ScrubString(g_AdminURL)
    Try g_PillPackProfileID := ScrubString(g_PillPackProfileID)
    Try g_UFRReason        := ScrubString(g_UFRReason)

    ; NDC OOS Ticket
    Try g_SelectedNDC          := ScrubString(g_SelectedNDC)
    Try g_SelectedMedName      := ScrubString(g_SelectedMedName)
    Try g_SelectedRxNumber     := ScrubString(g_SelectedRxNumber)

    ; === CLEAR PHASE ===

    ; PDMP
    p_FirstName    := ""
    p_LastName     := ""
    p_DOB          := ""
    p_StateName    := ""
    g_BatchRunning := false

    ; Refusal to Fill
    p_PatientFullName   := ""
    p_MedName           := ""
    p_Batch             := ""
    p_Identifier        := ""
    p_OrderID           := ""
    p_LastFillDate      := ""
    p_LastFillDays      := ""
    p_LastPharmacy      := ""
    p_NextFillDate      := ""
    g_FormSubmitted     := false
    p_ShipDate          := ""
    p_ArrivalDate       := ""
    p_ArrivalParsed     := ""
    p_ShipmentURL       := ""
    p_PillPackProfileID := ""
    p_SelectedNDC       := ""
    p_SelectedRxNumber  := ""
    p_AdminURL          := ""
    p_CustomerPhone     := ""
    p_RTFReason         := ""
    p_RTFReasonDetail   := ""
    g_DetailVisible     := false
    g_RTFIsEarlyControl := false
    g_RTFDecision       := ""
    g_RefusalMedList      := []
    g_RefusalSelectedIdx  := 0
    g_RefusalGUIDropCount := 1
    g_RefusalGUIAction    := ""
    g_RefusalGUISaved     := []
    g_RefusalGUIDDLs      := []

    ; Tachyon Ticket
    g_RxNumber          := ""
    g_GroupID           := ""
    g_PatientName       := ""
    g_Batch             := ""
    g_PersonID          := ""
    g_PrescriptionType  := ""
    g_ShipDate          := ""
    g_ArrivalDate       := ""
    g_ArrivalParsed     := ""
    g_ReasonForChange   := ""
    g_ReasonIndex       := 0
    g_MedicationName    := ""
    g_ChangeNeeded      := ""
    g_ShipmentURL       := ""
    g_BlueprintName     := ""
    g_IsAmazon          := false
    g_AdminURL          := ""
    g_PillPackProfileID := ""
    g_UFRReason         := ""
    g_UFRReasonIndex    := 0
    g_IsRedShipment     := false
    g_StartedOnShipment := false
    g_ShipMedList       := []
    g_ShipSelectedIdx   := 0
    g_ShipGUIDropCount  := 1
    g_ShipGUIAction     := ""
    g_ShipGUISaved      := []
    g_ShipGUIDDLs       := []

    ; NDC OOS Ticket
    g_SelectedNDC            := ""
    g_SelectedMedName        := ""
    g_SelectedRxNumber       := ""
    g_MedList                := []
    g_SelectedMedIdx         := 0
    g_GUIDropCount           := 1
    g_GUIAction              := ""
    g_GUISavedSelections     := []
    g_GUIDDLs                := []

    ; --- Clipboard scrub: multiple writes force Chrome's clipboard format
    ;     watchers to drop the previous PHI snapshot. Best-effort. ---
    A_Clipboard := ""
    A_Clipboard := "                                "  ; displace prior buffer
    A_Clipboard := ""
}

; =====================================================================
;  HIPAA HELPER: Flag current clipboard contents as confidential.
;  Marks the clipboard so Windows Clipboard History and Cloud Clipboard
;  exclude this entry.
; =====================================================================
MarkClipboardExcluded() {
    static fmt_exclude := DllCall("RegisterClipboardFormat", "Str", "ExcludeClipboardContentFromMonitorProcessing", "UInt")
    static fmt_hist    := DllCall("RegisterClipboardFormat", "Str", "CanIncludeInClipboardHistory", "UInt")
    static fmt_cloud   := DllCall("RegisterClipboardFormat", "Str", "CanUploadToCloudClipboard", "UInt")
    static GMEM_MOVEABLE := 0x2, GMEM_ZEROINIT := 0x40

    ; Retry OpenClipboard briefly — another app may hold it for a few ms
    opened := false
    Loop 10 {
        if DllCall("OpenClipboard", "Ptr", 0) {
            opened := true
            break
        }
        Sleep(5)
    }
    if (!opened)
        return false

    ; Marker 1 — legacy "do not monitor" flag (NULL handle is intentional)
    DllCall("SetClipboardData", "UInt", fmt_exclude, "Ptr", 0)

    ; Marker 2 — exclude from Clipboard History (DWORD value 0 = exclude)
    hHist := DllCall("GlobalAlloc", "UInt", GMEM_MOVEABLE | GMEM_ZEROINIT, "Ptr", 4, "Ptr")
    DllCall("SetClipboardData", "UInt", fmt_hist, "Ptr", hHist)

    ; Marker 3 — exclude from Cloud Clipboard upload (DWORD value 0 = exclude)
    hCloud := DllCall("GlobalAlloc", "UInt", GMEM_MOVEABLE | GMEM_ZEROINIT, "Ptr", 4, "Ptr")
    DllCall("SetClipboardData", "UInt", fmt_cloud, "Ptr", hCloud)

    DllCall("CloseClipboard")
    return true
}

; =====================================================================
;  FUNCTION: Execute JavaScript via address-bar javascript: injection.
;  Pastes code from clipboard into the address bar after typing
;  "javascript:" prefix. Works on both laptop and desktop Chrome.
; =====================================================================
ExecJS(code) {
    Send("!d")
    Sleep(150)
    SendText("javascript:")
    Sleep(50)
    A_Clipboard := code
    Send("^v")
    Sleep(50)
    Send("{Enter}")
    Sleep(1500)
    A_Clipboard := ""
}

; Fast variant for short, simple JS commands (e.g., blur).
ExecJSFast(code) {
    Send("!d")
    Sleep(150)
    SendText("javascript:")
    Sleep(50)
    A_Clipboard := code
    Send("^v")
    Sleep(50)
    Send("{Enter}")
    Sleep(400)
    A_Clipboard := ""
}

; =====================================================================
;  FUNCTION: Re-activate Chrome after any user-facing dialog/GUI.
;  Ensures automation targets the correct window even if the user
;  clicked into another application while the dialog was open.
; =====================================================================
ReactivateChrome() {
    Try WinActivate("ahk_exe chrome.exe")
    Sleep(300)
}

; =====================================================================
;  FUNCTION: Clear text selection reliably
;  Uses {Right} instead of {Esc} to avoid cancelling in-progress
;  Chrome page loads.
; =====================================================================
ClearSelection() {
    Send("{Right}")
    Sleep(50)
    Send("^{Home}")
    Sleep(50)
}

; =====================================================================
;  FUNCTION: Verify text in active input field matches expected value.
;  Returns true if match, false otherwise. Caller handles abort logic.
; =====================================================================
VerifyFieldContent(expectedValue) {
    A_Clipboard := ""
    Send("^a")
    Sleep(50)
    Send("^c")
    If ClipWait(0.5) {
        ReadBack := Trim(A_Clipboard)
        A_Clipboard := ""
        result := (ReadBack == expectedValue)
        ReadBack := ""
        Send("{End}")
        Sleep(50)
        Return result
    }
    A_Clipboard := ""
    Return false
}

; =====================================================================
;  FUNCTION: Navigate to a field via Find (with Find Next) and type
;  a value into it. Uses SendText for fields that reject paste (e.g.,
;  date pickers, dropdowns with type-to-select).
; =====================================================================
FindAndType(label, value) {
    Send("^f")
    Sleep(50)
    SendText(label)
    Sleep(50)
    Send("{Enter}")
    Sleep(50)
    Send("{Esc}")
    Sleep(50)
    Send("{Tab}")
    Sleep(150)
    SendText(value)
    Sleep(200)
}

; =====================================================================
;  FUNCTION: Navigate to a field via Find (no Find Next) and type a
;  value. Uses SendText. For fields where Find highlights the label
;  directly (no need for Find Next).
; =====================================================================
FindAndTypeNoEnter(label, value) {
    Send("^f")
    Sleep(50)
    SendText(label)
    Sleep(50)
    Send("{Esc}")
    Sleep(50)
    Send("{Tab}")
    Sleep(150)
    SendText(value)
    Sleep(200)
}

; =====================================================================
;  FUNCTION: Navigate to a field via Find (with Find Next) and paste
;  a value via clipboard. Used for standard text inputs that accept
;  paste (Shipment ID, NDC, Person ID, etc.).
; =====================================================================
FindAndPasteField(label, value) {
    Send("^f")
    Sleep(50)
    SendText(label)
    Sleep(50)
    Send("{Enter}")
    Sleep(50)
    Send("{Esc}")
    Sleep(50)
    Send("{Tab}")
    Sleep(150)
    A_Clipboard := value
    MarkClipboardExcluded()
    Send("^v")
    Sleep(50)
    A_Clipboard := ""
    Sleep(200)
}

; =====================================================================
;  FUNCTION: Navigate to a textarea via Find and paste a value.
;  Appends to end of the found line via {End}.
; =====================================================================
FindAndPaste(label, value) {
    Send("^f")
    Sleep(50)
    SendText(label)
    Sleep(50)
    Send("{Esc}")
    Sleep(50)
    Send("{End}")
    Sleep(100)
    A_Clipboard := " " . value
    MarkClipboardExcluded()
    Sleep(50)
    Send("^v")
    Sleep(50)
    A_Clipboard := ""
    Sleep(200)
}

; =====================================================================
;  FUNCTION: Move down one line and paste a value (for continuation
;  lines in description boxes).
; =====================================================================
DownAndPaste(value) {
    Send("{Down}")
    Sleep(100)
    A_Clipboard := " " . value
    MarkClipboardExcluded()
    Sleep(50)
    Send("^v")
    Sleep(50)
    A_Clipboard := ""
    Sleep(200)
}

; =====================================================================
;  FUNCTION: Move down one line, go to end, and paste a value.
; =====================================================================
DownEndAndPaste(value) {
    Send("{Down}")
    Sleep(100)
    Send("{End}")
    Sleep(100)
    A_Clipboard := " " . value
    MarkClipboardExcluded()
    Sleep(50)
    Send("^v")
    Sleep(50)
    A_Clipboard := ""
    Sleep(200)
}

; =====================================================================
;  FUNCTION: Navigate to a dropdown via Find and select by index.
;  Opens dropdown, goes to Home, then Down (index-1) times.
; =====================================================================
FindAndSelectDropdown(label, index) {
    Send("^f")
    Sleep(50)
    SendText(label)
    Sleep(50)
    Send("{Enter}")
    Sleep(50)
    Send("{Esc}")
    Sleep(50)
    Send("{Tab}")
    Sleep(150)
    Send("{Enter}")
    Sleep(200)
    Send("{Home}")
    Sleep(100)
    Loop (index - 1) {
        Send("{Down}")
        Sleep(100)
    }
    Send("{Enter}")
    Sleep(200)
}

; =====================================================================
;  FUNCTION: Navigate to a calendar via Find, open it, navigate to
;  the target date, and confirm.  Returns false if parsedDate is not
;  an object (parse failed).
; =====================================================================
FindAndSetCalendar(label, parsedDate) {
    Send("^f")
    Sleep(50)
    SendText(label)
    Sleep(50)
    Send("{Esc}")
    Sleep(50)
    Send("{Tab}")
    Sleep(200)
    Send("{Enter}")
    Sleep(500)
    If (Type(parsedDate) != "Object") {
        MsgBox("Date was not parsed -- skipping calendar navigation.`n`nManually set the " . label . " before submitting the ticket.", "Calendar Skipped", 48)
        ReactivateChrome()
        Send("{Esc}")
        Sleep(200)
        Return false
    }
    DaysDiff := CalculateDaysDifference(parsedDate.year, parsedDate.month, parsedDate.day)
    NavigateCalendar(DaysDiff)
    Sleep(200)
    Send("{Enter}")
    Sleep(200)
    Return true
}

; =====================================================================
;  FUNCTION: Click blueprint and verify task type via page scrape.
;  Clicks up to 3 times, scrapes up to 2 times per click.
;  Returns true if ExpectedTask found on page, false otherwise.
; =====================================================================
ClickBlueprintAndVerify(JSPayload, ExpectedTask, isAmazon) {
    Loop 3 {
        Send("!d")
        Sleep(100)
        SendText("javascript:")
        Sleep(50)
        A_Clipboard := JSPayload
        Send("^v")
        Sleep(50)
        Send("{Enter}")
        Sleep(300)
        A_Clipboard := ""

        If (ExpectedTask != "") {
            Loop 2 {
                Send("!d")
                Sleep(100)
                SendText("javascript:")
                Sleep(50)
                A_Clipboard := "document.activeElement.blur();void(0);"
                Send("^v")
                Sleep(50)
                Send("{Enter}")
                Sleep(100)
                A_Clipboard := ""

                A_Clipboard := ""
                Send("^a")
                Sleep(50)
                Send("^c")
                If ClipWait(0.5) {
                    PageText := A_Clipboard
                    A_Clipboard := ""
                    If InStr(PageText, ExpectedTask) {
                        PageText := ""
                        ClearSelection()
                        Return true
                    }
                    PageText := ""
                } Else {
                    A_Clipboard := ""
                }
                Sleep(300)
            }
        } Else {
            ClearSelection()
            Return true
        }
    }
    Return false
}

; =====================================================================
;  FUNCTION: Select customer name from profile lookup results.
;  Clicks the first visible name after "CUSTOMER PROFILE FOUND",
;  then verifies "Customer first name" field is populated.
;  Retries up to 10 times with 300ms between attempts.
;  Returns true if customer name was confirmed, false otherwise.
; =====================================================================
SelectCustomerName(isAmazon) {
    static JSClick := "(function(){var walker=document.createTreeWalker(document.body,NodeFilter.SHOW_ELEMENT,null,false);var foundLabel=false;while(walker.nextNode()){var el=walker.currentNode;var txt=(el.innerText||'').trim();if(!foundLabel&&txt.indexOf('CUSTOMER PROFILE FOUND')!==-1&&el.children.length===0){foundLabel=true;continue;}if(foundLabel&&el.offsetHeight>0&&el.children.length===0){var t=(el.innerText||'').trim();if(t&&t.length>1&&t!=='CUSTOMER PROFILE FOUND'){el.click();return;}}}})();void(0);"

    Sleep(isAmazon ? 900 : 2000)

    Loop 10 {
        Send("!d")
        Sleep(100)
        SendText("javascript:")
        Sleep(50)
        A_Clipboard := JSClick
        Send("^v")
        Sleep(50)
        Send("{Enter}")
        Sleep(300)
        A_Clipboard := ""

        Send("^f")
        Sleep(150)
        SendText("Customer first name")
        Sleep(150)
        Send("{Esc}")
        Sleep(150)
        Send("{Tab}")
        Sleep(150)

        A_Clipboard := ""
        Send("^a")
        Sleep(50)
        Send("^c")

        If ClipWait(0.3) {
            FieldVal := Trim(A_Clipboard)
            A_Clipboard := ""
            If (FieldVal != "" && !RegExMatch(FieldVal, "^https?://")) {
                FieldVal := ""
                Return true
            }
            FieldVal := ""
        } Else {
            A_Clipboard := ""
        }

        Sleep(300)
    }
    Return false
}

; =====================================================================
;  FUNCTION: Clear selection and refocus page body via address-bar
;  javascript: injection. Works on both laptop and desktop Chrome.
; =====================================================================
ClearSelectionJS() {
    Send("!d")
    Sleep(150)
    SendText("javascript:")
    Sleep(50)
    A_Clipboard := "window.getSelection().removeAllRanges();void(0);"
    Send("^v")
    Sleep(50)
    Send("{Enter}")
    Sleep(200)
    A_Clipboard := ""
}

; =====================================================================
;  FUNCTION: Non-disruptive login detection via Chrome window title
; =====================================================================
CheckChromeTitleForLogin(Context) {
    Try {
        ChromeTitle := WinGetTitle("ahk_exe chrome.exe")
    } Catch {
        Return false
    }

    If (InStr(ChromeTitle, "Okta")
        || InStr(ChromeTitle, "Sign-In")
        || InStr(ChromeTitle, "Sign In")
        || InStr(ChromeTitle, "Verify")
        || InStr(ChromeTitle, "PillPack User Authentication")) {
        PauseStart()
        MsgBox("Sign-in required (" . Context . ").`n`n1. Enter your password and complete sign-in.`n2. Wait until the destination page is fully loaded.`n3. Click OK ONLY after the correct page is visible.", "Login Required", 64+262144)
        PauseEnd()
        Sleep(300)
        WinActivate("ahk_exe chrome.exe")
        Sleep(300)
        Return true
    }
    Return false
}

; =====================================================================
;  FUNCTION: Check Chrome address bar URL for Okta login redirect
; =====================================================================
CheckForOktaLogin(Context) {
    Send("!d")
    Sleep(150)
    A_Clipboard := ""
    Send("^c")
    Sleep(100)
    TempURL := ""
    If ClipWait(1)
        TempURL := A_Clipboard
    A_Clipboard := ""
    Send("{Esc}")
    Sleep(150)

    If InStr(TempURL, "okta") {
        TempURL := ""
        PauseStart()
        MsgBox("Sign-in required (" . Context . ").`n`n1. Enter your password and complete sign-in.`n2. Wait until the destination page is fully loaded.`n3. Click OK ONLY after the correct page is visible.", "Login Required", 64+262144)
        PauseEnd()
        Sleep(300)
        WinActivate("ahk_exe chrome.exe")
        Sleep(300)
        Return true
    }
    TempURL := ""
    Return false
}

; =====================================================================
;  SHARED NAVIGATION HELPERS (V2.2.0)
;  Extracted from repeated patterns across RTF, Tachyon, and NDC OOS.
; =====================================================================

; =====================================================================
;  FUNCTION: Capture URL from Chrome address bar (retry up to 4 times).
;  Only validates ^https?:// — caller checks content (shipment, etc.).
; =====================================================================
CaptureURLFromAddressBar(&CapturedURL) {
    CapturedURL := ""
    Loop 4 {
        Send("!d")
        Sleep(250)
        A_Clipboard := ""
        Send("^a")
        Sleep(50)
        Send("^c")
        If ClipWait(1.5) {
            TryURL := Trim(A_Clipboard)
            A_Clipboard := ""
            If RegExMatch(TryURL, "^https?://") {
                CapturedURL := TryURL
                TryURL := ""
                Return true
            }
            TryURL := ""
        } Else {
            A_Clipboard := ""
        }
        Send("{Esc}")
        Sleep(200)
    }
    Send("{Esc}")
    Sleep(100)
    Return false
}

; =====================================================================
;  FUNCTION: Wait for a PillPack shipment page to become readable.
;  Caller is responsible for page focus injection beforehand if needed.
; =====================================================================
WaitForShipmentPage(&PageText) {
    PageText := ""
    Loop 15 {
        Sleep(A_Index = 1 ? 300 : 500)

        A_Clipboard := ""
        Send("^a")
        Sleep(50)
        Send("^c")

        If ClipWait(0.3) {
            PageText := A_Clipboard
            If RegExMatch(PageText, "^https?://[^\r\n]+$") {
                PageText := ""
                A_Clipboard := ""
                ClearSelectionJS()
                Continue
            }
            If InStr(PageText, "NDC")
                || InStr(PageText, "Dispensed")
                || InStr(PageText, "Ship Date")
                || InStr(PageText, "Scheduled")
                || InStr(PageText, "Unscheduled") {
                MarkClipboardExcluded()
                ClearSelection()
                Return true
            }
        } Else {
            CheckForOktaLogin("PillPack Admin")
        }
    }
    PageText := ""
    Return false
}

; =====================================================================
;  FUNCTION: Wait for Customer Details page to load and scrape.
; =====================================================================
WaitForDetailsPage(&DetailsText) {
    DetailsText := ""
    Loop 62 {
        Sleep(A_Index = 1 ? 1500 : 300)

        If CheckChromeTitleForLogin("PillPack Admin - Customer Details")
            Continue

        A_Clipboard := ""
        Send("^a")
        Sleep(50)
        Send("^c")

        If ClipWait(0.3) {
            DetailsText := A_Clipboard
            If RegExMatch(DetailsText, "^https?://[^\r\n]+$") {
                DetailsText := ""
                A_Clipboard := ""
                ClearSelectionJS()
                Continue
            }
            If InStr(DetailsText, "Batch") && InStr(DetailsText, "MRN") {
                MarkClipboardExcluded()
                ClearSelection()
                Return true
            }
            ClearSelection()
        } Else {
            ClearSelection()
            CheckForOktaLogin("PillPack Admin - Customer Details")
        }
    }
    DetailsText := ""
    Return false
}

; =====================================================================
;  FUNCTION: Wait for the Tachyon page to load (with sign-in handling).
; =====================================================================
WaitForTachyonPage() {
    Loop 30 {
        Sleep(A_Index = 1 ? 1500 : 300)

        If CheckChromeTitleForLogin("Tachyon")
            Continue

        A_Clipboard := ""
        Send("^a")
        Sleep(50)
        Send("^c")

        If ClipWait(0.3) {
            PageText := A_Clipboard

            If RegExMatch(PageText, "^https?://[^\r\n]+$") {
                PageText := ""
                A_Clipboard := ""
                ClearSelectionJS()
                Continue
            }

            If InStr(PageText, "CREATE NEW ISSUE")
                || InStr(PageText, "Search by Person Id")
                || InStr(PageText, "All Agents - AP")
                || InStr(PageText, "All Agents - PP")
                || InStr(PageText, "Search by PillPack User Id") {
                PageText := ""
                ClearSelection()
                Return true
            }

            If InStr(PageText, "Sign In") || InStr(PageText, "Username") || InStr(PageText, "Verify with your password") || InStr(PageText, "PillPack User Authentication") || InStr(PageText, "Powered by Okta") || InStr(PageText, "Keep me signed in") {
                PageText := ""
                PauseStart()
                MsgBox("Sign-in required!`n`n1. Please log in manually.`n2. Click OK on this popup ONLY AFTER you see the Tachyon page.", "Login Paused", 64+262144)
                PauseEnd()
                Sleep(300)
                WinActivate("ahk_exe chrome.exe")
                Sleep(300)
            }

            PageText := ""
            ClearSelection()
        } Else {
            ClearSelection()
            CheckForOktaLogin("Tachyon")
        }
    }
    Return false
}

; =====================================================================
;  FUNCTION: Wait for the "Create New Issue" page to finish loading.
; =====================================================================
WaitForCreateIssuePage() {
    Sleep(50)
    Loop 15 {
        Sleep(500)
        A_Clipboard := ""
        Send("^a")
        Sleep(50)
        Send("^c")
        If ClipWait(0.3) {
            PageText := A_Clipboard
            ClearSelection()
            If InStr(PageText, "Search for blueprint") && InStr(PageText, "Select the blueprint") {
                PageText := ""
                Return true
            }
            PageText := ""
        } Else {
            ClearSelection()
        }
    }
    Return false
}

; =====================================================================
;  FUNCTION: Find customer name on page and open link in new tab.
; =====================================================================
OpenCustomerInNewTab(name) {
    Send("^f")
    Sleep(50)
    SendText(name)
    Sleep(50)
    Send("{Esc}")
    Sleep(50)
    Send("^+{Enter}")
}

; =====================================================================
;  FUNCTION: Open a new tab and navigate to the correct Tachyon URL.
; =====================================================================
OpenTachyonTab(isAmazon) {
    Send("^t")
    Sleep(200)
    If (isAmazon) {
        A_Clipboard := "https://cs.wolfgang.a2z.com/tachyon/PHARMACY/queues/360006946111?program=PHARMACY&app=TACHYON"
    } Else {
        A_Clipboard := "https://cs.wolfgang.a2z.com/tachyon/PILLPACK/queues/114094341813"
    }
    Sleep(50)
    Send("^v")
    Sleep(100)
    Send("{Enter}")
    A_Clipboard := ""
}

; =====================================================================
;  FUNCTION: Click "CREATE NEW ISSUE" button via Find+Enter sequence.
; =====================================================================
ClickCreateNewIssue() {
    Send("^f")
    Sleep(50)
    SendText("CREATE NEW ISSUE")
    Sleep(50)
    Send("{Esc}")
    Sleep(50)
    Send("{Enter}")
}

; =====================================================================
;  FUNCTION: Search for blueprint by name and verify field content.
;  Returns true if verified, false on mismatch.
; =====================================================================
SearchAndVerifyBlueprint(name) {
    Send("^f")
    Sleep(50)
    SendText("Search")
    Sleep(50)
    Send("{Esc}")
    Sleep(50)
    Send("{Tab}")
    Sleep(150)

    A_Clipboard := name
    Send("^v")
    Sleep(50)
    A_Clipboard := ""
    Sleep(150)

    If !VerifyFieldContent(name)
        Return false

    Sleep(200)
    Send("{Enter}")
    Sleep(500)
    Send("{Esc}")
    Sleep(200)
    Return true
}

; =====================================================================
;  FUNCTION: Scrub profile ID globals, type and verify a profile ID.
;  Returns true on success, false on failure. Caller handles error.
; =====================================================================
TypeAndVerifyProfileID(label, value) {
    global g_PersonID, g_PillPackProfileID

    Try g_PersonID := ScrubString(g_PersonID)
    g_PersonID := ""
    Try g_PillPackProfileID := ScrubString(g_PillPackProfileID)
    g_PillPackProfileID := ""

    FindAndPasteField(label, value)

    If !VerifyFieldContent(value) {
        Try value := ScrubString(value)
        Return false
    }
    Try value := ScrubString(value)
    Return true
}

; =====================================================================
;  FUNCTION: Set the Priority dropdown to "High" if dueDate is today
;  or earlier. If dueDate is tomorrow or later (or unparseable), leave
;  Priority at its default (Low) and do nothing.
;  dueDate format: "mm/dd/yyyy"
; =====================================================================
SetPriorityIfDueToday(dueDate) {
    If (dueDate == "")
        Return
    If !RegExMatch(dueDate, "^(\d{1,2})/(\d{1,2})/(\d{4})$", &M)
        Return
    dMonth := Integer(M[1])
    dDay   := Integer(M[2])
    dYear  := Integer(M[3])
    DaysDiff := CalculateDaysDifference(dYear, dMonth, dDay)
    If (DaysDiff > 0)
        Return

    Send("^f")
    Sleep(50)
    SendText("Priority")
    Sleep(50)
    Send("{Esc}")
    Sleep(50)
    Send("{Tab}")
    Sleep(150)
    Send("{Enter}")
    Sleep(200)
    Send("{Home}")
    Sleep(100)
    Loop 2 {
        Send("{Down}")
        Sleep(100)
    }
    Send("{Enter}")
    Sleep(200)
}

; =====================================================================
;  FUNCTION: Convert "Wed, Apr 29" or similar to mm/dd/yyyy
; =====================================================================
ParseDateFromText(rawDate, defaultYear := "") {
    If (defaultYear == "")
        defaultYear := A_Year

    months := Map(
        "Jan", 1, "Feb", 2, "Mar", 3, "Apr", 4,
        "May", 5, "Jun", 6, "Jul", 7, "Aug", 8,
        "Sep", 9, "Oct", 10, "Nov", 11, "Dec", 12
    )

    result := {formatted: "", year: 0, month: 0, day: 0}

    If RegExMatch(rawDate, "i)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2})", &Match) {
        result.month := months[Match[1]]
        result.day   := Integer(Match[2])
        result.year  := Integer(defaultYear)
        monthDiff := result.month - Integer(A_Mon)
        If (monthDiff < -6)
            result.year := Integer(defaultYear) + 1
        Else If (monthDiff > 6)
            result.year := Integer(defaultYear) - 1
        result.formatted := Format("{:02}/{:02}/{}", result.month, result.day, result.year)
    }

    Return result
}

; =====================================================================
;  FUNCTION: Calculate days between today and a target date
; =====================================================================
CalculateDaysDifference(targetYear, targetMonth, targetDay) {
    todayYear  := Integer(A_Year)
    todayMonth := Integer(A_Mon)
    todayDay   := Integer(A_MDay)

    daysInMonth := [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

    If (Mod(todayYear, 4) == 0 && (Mod(todayYear, 100) != 0 || Mod(todayYear, 400) == 0))
        daysInMonthToday := [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    Else
        daysInMonthToday := daysInMonth.Clone()

    If (Mod(targetYear, 4) == 0 && (Mod(targetYear, 100) != 0 || Mod(targetYear, 400) == 0))
        daysInMonthTarget := [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    Else
        daysInMonthTarget := daysInMonth.Clone()

    todayDOY := todayDay
    Loop (todayMonth - 1)
        todayDOY += daysInMonthToday[A_Index]

    targetDOY := targetDay
    Loop (targetMonth - 1)
        targetDOY += daysInMonthTarget[A_Index]

    If (targetYear == todayYear) {
        Return targetDOY - todayDOY
    } Else If (targetYear > todayYear) {
        todayIsLeap := (Mod(todayYear, 4) == 0 && (Mod(todayYear, 100) != 0 || Mod(todayYear, 400) == 0))
        daysLeftThisYear := (todayIsLeap ? 366 : 365) - todayDOY
        Return daysLeftThisYear + targetDOY
    } Else {
        targetIsLeap := (Mod(targetYear, 4) == 0 && (Mod(targetYear, 100) != 0 || Mod(targetYear, 400) == 0))
        daysLeftTargetYear := (targetIsLeap ? 366 : 365) - targetDOY
        Return -(daysLeftTargetYear + todayDOY)
    }
}

; =====================================================================
;  FUNCTION: Calculate Due By date for non-Amazon tickets.
;  Returns (arrival - 1 day, skipping Sunday) as "mm/dd/yyyy".
;  If parsedArrival is not a valid object, returns "" (caller falls back).
; =====================================================================
CalculateDueByDate(parsedArrival) {
    If (Type(parsedArrival) != "Object")
        Return ""

    y := parsedArrival.year
    m := parsedArrival.month
    d := parsedArrival.day

    ; Subtract 1 day
    d -= 1
    If (d < 1) {
        m -= 1
        If (m < 1) {
            m := 12
            y -= 1
        }
        daysInMonth := [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        If (m == 2 && Mod(y, 4) == 0 && (Mod(y, 100) != 0 || Mod(y, 400) == 0))
            d := 29
        Else
            d := daysInMonth[m]
    }

    ; If result is Sunday (day of week = 1), subtract one more day to Saturday
    ; Zeller-like day-of-week: 0=Sun,1=Mon,...6=Sat (matching AHK's A_WDay - 1)
    ; Use the YYYYMMDD timestamp with FormatTime to get the weekday
    stamp := Format("{:04}{:02}{:02}000000", y, m, d)
    wday := FormatTime(stamp, "WDay")  ; 1=Sun, 2=Mon, ... 7=Sat
    If (Integer(wday) == 1) {
        ; Sunday -> go back to Saturday
        d -= 1
        If (d < 1) {
            m -= 1
            If (m < 1) {
                m := 12
                y -= 1
            }
            daysInMonth := [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
            If (m == 2 && Mod(y, 4) == 0 && (Mod(y, 100) != 0 || Mod(y, 400) == 0))
                d := 29
            Else
                d := daysInMonth[m]
        }
    }

    Return Format("{:02}/{:02}/{}", m, d, y)
}

; =====================================================================
;  FUNCTION: Navigate a calendar picker by days difference from today
; =====================================================================
NavigateCalendar(daysDifference) {
    If (daysDifference == 0)
        Return

    If (daysDifference > 0) {
        weeks := daysDifference // 7
        days  := Mod(daysDifference, 7)
        Loop weeks {
            Send("{Down}")
            Sleep(40)
        }
        Loop days {
            Send("{Right}")
            Sleep(40)
        }
    } Else {
        absDiff := Abs(daysDifference)
        weeks   := absDiff // 7
        days    := Mod(absDiff, 7)
        Loop weeks {
            Send("{Up}")
            Sleep(40)
        }
        Loop days {
            Send("{Left}")
            Sleep(40)
        }
    }
}

; =====================================================================
;  PDMP Window Group
; =====================================================================
GroupAdd("PDMPGroup", "Patient Request")
GroupAdd("PDMPGroup", "CURES")
GroupAdd("PDMPGroup", "Cures")
GroupAdd("PDMPGroup", "PMP")
GroupAdd("PDMPGroup", "CSD Individual Search")

; === WORKFLOW-SPECIFIC FUNCTIONS AND HOTKEYS FOLLOW ===

; =====================================================================
;  PDMP-SPECIFIC FUNCTIONS
; =====================================================================

; =====================================================================
;  STATE-NAME DISAMBIGUATION HELPERS (PDMP)
; =====================================================================
GetStateDetectionRegex(stateName) {
    ; Word-aware regex used by both detection and target-position
    ; location. "i)" enables case-insensitive matching in AHK's PCRE.
    ; Virginia needs a lookbehind to exclude the word-bounded
    ; substring inside "West Virginia." Every other state is safe
    ; with plain \b boundaries.
    if (stateName = "Virginia")
        return "i)(?<!West )\bVirginia\b"
    return "i)\b" . stateName . "\b"
}

LocateStateTargetPos(PageText, stateName, interconnectPos, footerPos) {
    ; Returns the 1-based position of the correct state option in
    ; PageText, or 0 if not found. Searches only within the
    ; InterConnect section so that home-state references in the
    ; header and footer can't be mistaken for the target.
    if (interconnectPos = 0)
        return 0
    if (footerPos > interconnectPos)
        sectionText := SubStr(PageText, interconnectPos, footerPos - interconnectPos)
    else
        sectionText := SubStr(PageText, interconnectPos)

    localPos := RegExMatch(sectionText, GetStateDetectionRegex(stateName))
    if (localPos = 0)
        return 0
    return interconnectPos + localPos - 1
}

CountSubstringMatchesBefore(haystack, needle, beforePos) {
    ; Counts case-insensitive substring occurrences of needle in
    ; haystack with starting position strictly less than beforePos.
    ; This mirrors what Chrome's Find sees, and the return value is
    ; the number of {Enter} (Find Next) presses needed to advance
    ; past every false match. Capped at 1000 iterations as a
    ; runaway guard — no realistic PMP page approaches that.
    if (needle = "" || beforePos <= 1)
        return 0
    count := 0
    searchPos := 1
    needleLen := StrLen(needle)
    Loop 1000 {
        searchPos := InStr(haystack, needle, false, searchPos)
        if (searchPos = 0 || searchPos >= beforePos)
            break
        count++
        searchPos += needleLen
    }
    return count
}

; =====================================================================
;  SCRAPE PILLPACK (PDMP)
; =====================================================================
ScrapePillPack() {
    global p_FirstName, p_LastName, p_DOB, p_StateName

    A_Clipboard := ""
    Send("^a")
    Sleep(30)
    Send("^c")

    if !ClipWait(2) {
        MsgBox("Error: Could not copy text. The clipboard remained empty.")
        return false
    }
    MarkClipboardExcluded()              ; HIPAA: flag this clip before reading it

    PageText := SubStr(A_Clipboard, 1, 15000)
    Send("^{Home}")

    p_FirstName := ""
    p_LastName  := ""
    p_DOB       := ""
    p_StateName := ""

    ; --- Name parsing: allows (), periods, commas before the Schedule dash ---
    if RegExMatch(PageText, "([a-zA-Z\-' \(\)\.,]+?)\s*[-—–‐]\s*(?:Unscheduled|Scheduled)", &NameMatch) {
        FullName := NameMatch[1]
        FullName := RegExReplace(FullName, "\([^)]*\)", "")     ; drop "(TJ)" etc.
        FullName := RegExReplace(FullName, "[,.]", "")          ; drop ", Jr." commas/periods
        FullName := RegExReplace(FullName, "\s+", " ")
        FullName := Trim(FullName)

        NameArray := StrSplit(FullName, " ")

        if (NameArray.Length > 0) {
            p_FirstName := NameArray[1]
            if (NameArray.Length > 1) {
                Loop NameArray.Length {
                    Index := NameArray.Length - A_Index + 1
                    if (Index = 1)
                        break
                    CurrentWord := NameArray[Index]
                    if RegExMatch(CurrentWord, "i)^(Jr|Sr|II|III|IV|V|VI|VII|VIII)$")
                        continue                                 ; skip suffix
                    p_LastName := CurrentWord
                    break
                }
            }
        }
    }

    ; --- DOB parsing (now REQUIRED — used as anchor for state search) ---
    DobPos := RegExMatch(PageText, "m)^DOB:\s*([a-zA-Z]{3})\s*(\d{1,2}),\s*(\d{4})", &DobMatch)
    if (DobPos = 0) {
        WipeAllData()
        MsgBox("STOPPING: Could not find the patient's Date of Birth on the page.")
        return false
    }
    MonthMap := Map("Jan","01","Feb","02","Mar","03","Apr","04","May","05","Jun","06","Jul","07","Aug","08","Sep","09","Oct","10","Nov","11","Dec","12")
    NormalizedMonth := StrTitle(DobMatch[1])                     ; handles JAN/jan/Jan in one call
    p_DOB := MonthMap[NormalizedMonth] . "/" . Format("{:02}", DobMatch[2]) . "/" . DobMatch[3]

    if (StrLen(p_DOB) != 10) {
        WipeAllData()
        MsgBox("STOPPING: Date of Birth did not parse correctly. Please check the source page.")
        return false
    }

    ; --- State parsing: pick ZIP+state pair nearest the DOB anchor ---
    BestStateCode := ""
    SmallestDistance := 9999999
    SearchPos := 1
    while (SearchPos := RegExMatch(PageText, ",\s*([A-Z]{2})\s+\d{5}\b", &StateMatch, SearchPos)) {
        Distance := Abs(SearchPos - DobPos)
        if (Distance < SmallestDistance) {
            SmallestDistance := Distance
            BestStateCode := StateMatch[1]
        }
        SearchPos += StrLen(StateMatch[0])
    }

    if (BestStateCode != "") {
        ; Added "DC", "District of Columbia" to the State Map
        StateMap := Map("AL","Alabama","AK","Alaska","AZ","Arizona","AR","Arkansas","CA","California","CO","Colorado","CT","Connecticut","DC","District of Columbia","DE","Delaware","FL","Florida","GA","Georgia","HI","Hawaii","ID","Idaho","IL","Illinois","IN","Indiana","IA","Iowa","KS","Kansas","KY","Kentucky","LA","Louisiana","ME","Maine","MD","Maryland","MA","Massachusetts","MI","Michigan","MN","Minnesota","MS","Mississippi","MO","Missouri","MT","Montana","NE","Nebraska","NV","Nevada","NH","New Hampshire","NJ","New Jersey","NM","New Mexico","NY","New York","NC","North Carolina","ND","North Dakota","OH","Ohio","OK","Oklahoma","OR","Oregon","PA","Pennsylvania","RI","Rhode Island","SC","South Carolina","SD","South Dakota","TN","Tennessee","TX","Texas","UT","Utah","VT","Vermont","VA","Virginia","WA","Washington","WV","West Virginia","WI","Wisconsin","WY","Wyoming")
        if StateMap.Has(BestStateCode)
            p_StateName := StateMap[BestStateCode]
    }

    ; --- Final validation ---
    if (p_FirstName = "") {
        PageText := ScrubString(PageText)
        PageText := ""
        WipeAllData()
        MsgBox("STOPPING: Could not find the patient's First Name.")
        return false
    }
    if (p_LastName = "") {
        PageText := ScrubString(PageText)
        PageText := ""
        WipeAllData()
        MsgBox("STOPPING: Could not determine the patient's Last Name. Only one name token was found.")
        return false
    }
    if (p_StateName = "") {
        PageText := ScrubString(PageText)
        PageText := ""
        WipeAllData()
        MsgBox("STOPPING: Could not determine the patient's State. No ZIP+state pattern was found on the page.")
        return false
    }

    PageText := ScrubString(PageText)
    PageText := ""
    A_Clipboard := ""
    return true
}

; =====================================================================
;  PASTE PDMP (PDMP)
; =====================================================================
PastePDMP() {
    global p_FirstName, p_LastName, p_DOB, p_StateName

    if (p_FirstName = "") {
        MsgBox("No data to paste!")
        return false
    }

    A_Clipboard := ""
    Send("{Esc}")
    Sleep(30)
    Send("^a")
    Sleep(30)
    Send("^c")
    if !ClipWait(1) {
        MsgBox("STOPPING: Could not read the PDMP page.")
        return false
    }
    MarkClipboardExcluded()

    PageText := A_Clipboard
    Send("^{Home}")

    if (StrLen(PageText) < 100) {
        MsgBox("STOPPING: Please click on a blank white area of the PDMP page. The script got trapped.")
        return false
    }

    ; =================================================================
    ;  ROUTE 0: UTAH (Controlled Substance Database)
    ; =================================================================
    if InStr(PageText, "Controlled Substance Database") {
        ; Utah CSD searchable states (home state Utah + interstate list)
        static UtahSearchable := Map(
            "Utah", true,
            "Arizona", true,
            "Colorado", true,
            "Connecticut", true,
            "Idaho", true,
            "Illinois", true,
            "Iowa", true,
            "Kansas", true,
            "Kentucky", true,
            "Maryland", true,
            "Minnesota", true,
            "Montana", true,
            "Nebraska", true,
            "Nevada", true,
            "New Mexico", true,
            "New York", true,
            "North Dakota", true,
            "Oregon", true,
            "Pennsylvania", true,
            "South Carolina", true,
            "South Dakota", true,
            "Texas", true,
            "Virginia", true,
            "Washington", true,
            "Wisconsin", true,
            "Wyoming", true
        )

        PageText := ScrubString(PageText)
        PageText := ""

        if !UtahSearchable.Has(p_StateName) {
            MsgBox("STOPPING: Patient's state is not searchable in this database.")
            return false
        }

        ; --- Find "Last" field ---
        Send("^f")
        Sleep(150)
        SendText("Last")
        Sleep(100)
        Send("{Enter}")
        Sleep(100)
        Send("{Enter}")
        Sleep(100)
        Send("{Esc}")
        Sleep(150)
        Send("{Tab}")
        Sleep(150)

        SendText(p_LastName)
        Sleep(100)
        Send("{Tab}")
        Sleep(150)
        SendText(p_FirstName)
        Sleep(100)
        Send("{Tab}")
        Sleep(150)
        SendText(p_DOB)
        Sleep(150)

        ; --- Interstate state selection (skip for Utah home state) ---
        if (p_StateName != "Utah") {
            Send("^f")
            Sleep(150)
            SendText("interstate")
            Sleep(100)
            Send("{Enter}")
            Sleep(100)
            Send("{Esc}")
            Sleep(150)
            Send("{Tab}")
            Sleep(150)
            SendText(p_StateName)
            Sleep(150)
            Send("{Enter}")
            Sleep(200)
        }

        ; --- Submit ---
        Send("^f")
        Sleep(150)
        SendText("Submit")
        Sleep(100)
        Send("{Enter}")
        Sleep(100)
        Send("{Esc}")
        Sleep(150)
        Send("{Space}")

        ; --- Wait for results page (250ms polling, 15s max) ---
        Loaded := false
        NeedsSelectAll := false
        ResultsText := ""
        Loop 60 {
            Sleep(250)
            A_Clipboard := ""
            Send("^a")
            Sleep(30)
            Send("^c")
            if ClipWait(0.5) {
                MarkClipboardExcluded()
                ResultsText := A_Clipboard
                if InStr(ResultsText, "***Please click on rows to select as many patients as you would like to view.") {
                    Loaded := true
                    NeedsSelectAll := true
                    break
                } else if InStr(ResultsText, "Selected Patient") {
                    Loaded := true
                    break
                } else if (StrLen(ResultsText) > 200 && !InStr(ResultsText, "Submit")) {
                    Loaded := true
                    break
                }
            }
        }

        if (!Loaded) {
            ResultsText := ScrubString(ResultsText)
            ResultsText := ""
            A_Clipboard := ""
            MsgBox("STOPPING: Utah CSD results page took too long to load.")
            return false
        }

        ResultsText := ScrubString(ResultsText)
        ResultsText := ""

        if (NeedsSelectAll) {
            Send("^f")
            Sleep(150)
            SendText("select all")
            Sleep(100)
            Send("{Enter}")
            Sleep(100)
            Send("{Esc}")
            Sleep(150)
            Send("{Space}")
            Sleep(300)

            Send("^f")
            Sleep(150)
            SendText("view selected")
            Sleep(100)
            Send("{Enter}")
            Sleep(100)
            Send("{Esc}")
            Sleep(150)
            Send("{Space}")
        }

        A_Clipboard := ""
        return true

    ; =================================================================
    ;  ROUTE 1: CALIFORNIA (CURES)
    ; =================================================================
    } else if InStr(PageText, "CURES") {
        PageText := ScrubString(PageText)
        PageText := ""

        if (p_StateName != "California") {
            ; Generic — no patient state revealed in dialog
            MsgBox("STOPPING: Patient is not in CA but you are on the CA CURES page.")
            return false
        }

        ; --- Find "Last Name" field ---
        Send("^f")
        Sleep(75)                        ; v2.1.4: 50->75ms — Find bar must be ready before typing
        SendText("Last Name")
        Sleep(50)
        Send("{Esc}")
        Sleep(75)                        ; v2.1.4: 20->75ms — wait for focus to return to page
        Send("{Tab}")
        Sleep(50)

        ; Empty-tab check
        A_Clipboard := ""
        Send("^a")
        Sleep(30)
        Send("^c")
        if ClipWait(0.05) {
            if (Trim(A_Clipboard) != "") {
                MarkClipboardExcluded()
                MsgBox("STOPPING: No empty PDMP tabs found! This tab already contains data.")
                return false
            }
        }

        SendText(p_LastName)
        Sleep(30)
        Send("{Tab}")
        Sleep(50)
        SendText(p_FirstName)
        Sleep(30)
        Send("{Tab}")
        Sleep(50)

        ; --- DOB typed via SendText (no clipboard exposure) ---
        ; v2.1.4: +30ms pre-focus pause — ensures field is active before first keystroke
        Sleep(30)
        SendText(p_DOB)
        Sleep(50)

        Send("{Tab 2}")
        Sleep(30)
        Send("{Enter}")

        ; --- Wait for results page (responsive 250ms polling, 15s max) ---
        Loaded := false
        ResultsText := ""
        Loop 60 {
            Sleep(250)
            A_Clipboard := ""
            Send("^a")
            Sleep(30)
            Send("^c")
            if ClipWait(0.5) {
                MarkClipboardExcluded()
                ResultsText := A_Clipboard
                if InStr(ResultsText, "Patients Found Matching Search Criteria") || InStr(ResultsText, "Select All") {
                    Loaded := true
                    Send("^{Home}")
                    break
                }
            }
        }
        Sleep(100)

        if (!Loaded) {
            ResultsText := ScrubString(ResultsText)
            ResultsText := ""
            A_Clipboard := ""
            MsgBox("STOPPING: CA CURES results page took too long to load.")
            return false
        }

        if RegExMatch(ResultsText, "i)Patients Found Matching Search Criteria:\s*0") {
            ResultsText := ScrubString(ResultsText)
            ResultsText := ""
            A_Clipboard := ""
            return true
        }

        ResultsText := ScrubString(ResultsText)
        ResultsText := ""

        ; --- Select All ---
        Send("^f")
        Sleep(75)                        ; v2.1.4: 50->75ms
        SendText("Select All")
        Sleep(50)
        Send("{Esc}")
        Sleep(75)                        ; v2.1.4: 20->75ms
        Send("{Tab}")
        Sleep(30)
        Send("+{Tab}")
        Sleep(30)
        Send("{Enter}")
        Sleep(600)                       ; v2.1.4: 300->600ms — page must settle before Generate

        ; --- Generate Report ---
        Send("^f")
        Sleep(75)                        ; v2.1.4: 50->75ms
        SendText("Generate Report")
        Sleep(50)
        Send("{Esc}")
        Sleep(75)                        ; v2.1.4: 20->75ms
        Send("{Tab}")
        Sleep(30)
        Send("+{Tab}")
        Sleep(30)
        Send("{Enter}")

        A_Clipboard := ""
        return true

    ; =================================================================
    ;  ROUTE 2: STANDARD PDMP
    ; =================================================================
    } else {
        IsHomeState := false
        IsInterconnectState := false
        IsFloridaPDMP := InStr(PageText, "E-FORCSE") > 0
        AgreementKeyword := ""

        StateRegex := GetStateDetectionRegex(p_StateName)

        ; --- InterConnect detection (page scrape) ---
        InterconnectPos := InStr(PageText, "PMP Interconnect")
        FooterPos := InStr(PageText, "Powered By")

        if (InterconnectPos > 0) {
            if (FooterPos > InterconnectPos)
                InterconnectText := SubStr(PageText, InterconnectPos, FooterPos - InterconnectPos)
            else
                InterconnectText := SubStr(PageText, InterconnectPos)
            if RegExMatch(InterconnectText, StateRegex)
                IsInterconnectState := true
        }

        ; --- URL-based home state detection: only if InterConnect
        ;     did not find the patient's state. Scrapes the URL from
        ;     Chrome's address bar and checks if the domain contains
        ;     the state name or abbreviation as a URL segment. ---
        if (!IsInterconnectState) {
            A_Clipboard := ""
            Send("^l")
            Sleep(75)
            Send("^c")
            Sleep(50)
            PageUrl := ""
            if ClipWait(0.5)
                PageUrl := Trim(A_Clipboard)
            Send("{Esc}")
            Sleep(50)

            if (PageUrl != "") {
                UrlLower := StrLower(PageUrl)
                StateNameLower := StrLower(p_StateName)
                StateNameHyphen := StrReplace(StateNameLower, " ", "-")
                StateNameJoined := StrReplace(StateNameLower, " ", "")
                ReverseStateMap := Map("Alabama","AL","Alaska","AK","Arizona","AZ","Arkansas","AR","California","CA","Colorado","CO","Connecticut","CT","District of Columbia","DC","Delaware","DE","Florida","FL","Georgia","GA","Hawaii","HI","Idaho","ID","Illinois","IL","Indiana","IN","Iowa","IA","Kansas","KS","Kentucky","KY","Louisiana","LA","Maine","ME","Maryland","MD","Massachusetts","MA","Michigan","MI","Minnesota","MN","Mississippi","MS","Missouri","MO","Montana","MT","Nebraska","NE","Nevada","NV","New Hampshire","NH","New Jersey","NJ","New Mexico","NM","New York","NY","North Carolina","NC","North Dakota","ND","Ohio","OH","Oklahoma","OK","Oregon","OR","Pennsylvania","PA","Rhode Island","RI","South Carolina","SC","South Dakota","SD","Tennessee","TN","Texas","TX","Utah","UT","Vermont","VT","Virginia","VA","Washington","WA","West Virginia","WV","Wisconsin","WI","Wyoming","WY")
                StateAbbrLower := ""
                if ReverseStateMap.Has(p_StateName)
                    StateAbbrLower := StrLower(ReverseStateMap[p_StateName])
                UrlSegment := "(://|[.\-/])"
                if RegExMatch(UrlLower, UrlSegment . StateNameLower . "[\.\-/]") || RegExMatch(UrlLower, UrlSegment . StateNameHyphen . "[\.\-/]") || RegExMatch(UrlLower, UrlSegment . StateNameJoined . "[\.\-/]") || (StateAbbrLower != "" && RegExMatch(UrlLower, "(://|\.)" . StateAbbrLower . "[\./]"))
                    IsHomeState := true
            }
            ; Return focus to page content after address bar interaction
            Send("{F6}")
            Sleep(75)
        }

        if (!IsHomeState && !IsInterconnectState) {
            PageText := ScrubString(PageText)
            PageText := ""
            ; Generic — no patient state revealed in dialog
            MsgBox("STOPPING: Patient's state is not searchable in this database.")
            return false
        }

        ; --- Agreement keyword detection ---
        ; Find the keyword that appears exactly once between the InterConnect
        ; state list and "Powered By" footer. FL (E-FORCSE) skips checkbox.
        if (!IsFloridaPDMP) {
            CheckZone := (FooterPos > InterconnectPos)
                ? SubStr(PageText, InterconnectPos, FooterPos - InterconnectPos)
                : SubStr(PageText, InterconnectPos)
            for , kw in ["agree", "certify", "acknowledge"] {
                FoundPos := InStr(CheckZone, kw, false)
                if (FoundPos > 0 && InStr(CheckZone, kw, false, FoundPos + StrLen(kw)) = 0) {
                    AgreementKeyword := kw
                    break
                }
            }
            if (AgreementKeyword = "") {
                PageText := ScrubString(PageText)
                PageText := ""
                MsgBox("STOPPING: Could not detect agreement checkbox keyword on this PDMP page.")
                return false
            }
        }

        ; --- Adaptive field order detection ---
        ; Determine the order of First Name, Last Name, DOB on this page.
        FNPos := InStr(PageText, "First Name")
        LNPos := InStr(PageText, "Last Name")
        DOBPos := InStr(PageText, "Date of Birth")
        if (FNPos = 0 || LNPos = 0 || DOBPos = 0) {
            PageText := ScrubString(PageText)
            PageText := ""
            MsgBox("STOPPING: Could not locate all required fields (First Name, Last Name, Date of Birth) on the PDMP page.")
            return false
        }
        FieldOrder := []
        FieldOrder.Push({pos: FNPos, label: "First Name", value: p_FirstName})
        FieldOrder.Push({pos: LNPos, label: "Last Name", value: p_LastName})
        FieldOrder.Push({pos: DOBPos, label: "Date of Birth", value: p_DOB})
        ; Sort by position (simple 3-element sort)
        Loop 2 {
            Loop (FieldOrder.Length - A_Index) {
                if (FieldOrder[A_Index].pos > FieldOrder[A_Index + 1].pos) {
                    tmp := FieldOrder[A_Index]
                    FieldOrder[A_Index] := FieldOrder[A_Index + 1]
                    FieldOrder[A_Index + 1] := tmp
                }
            }
        }
        ; Detect "Partial Spelling" checkboxes between fields
        ; Tab 2 if checkbox present, Tab 1 if not
        Zone1 := SubStr(PageText, FieldOrder[1].pos, FieldOrder[2].pos - FieldOrder[1].pos)
        Tab1Count := InStr(Zone1, "Partial Spelling") ? 2 : 1
        Zone2 := SubStr(PageText, FieldOrder[2].pos, FieldOrder[3].pos - FieldOrder[2].pos)
        Tab2Count := InStr(Zone2, "Partial Spelling") ? 2 : 1

        if (IsHomeState) {
            PageText := ScrubString(PageText)
            PageText := ""
        }

        ; --- Find first field ---
        Send("^f")
        Sleep(75)
        SendText(FieldOrder[1].label)
        Sleep(50)
        Send("{Esc}")
        Sleep(75)
        Send("{Tab}")
        Sleep(50)

        ; Empty-tab check
        A_Clipboard := ""
        Send("^a")
        Sleep(30)
        Send("^c")
        if ClipWait(0.05) {
            if (Trim(A_Clipboard) != "") {
                MarkClipboardExcluded()
                MsgBox("STOPPING: No empty PDMP tabs found! This tab already contains data.")
                return false
            }
        }

        SendText(FieldOrder[1].value)
        Sleep(30)
        Send("{Tab " . Tab1Count . "}")
        Sleep(50)
        SendText(FieldOrder[2].value)
        Sleep(30)
        Send("{Tab " . Tab2Count . "}")
        Sleep(50)

        ; --- Third field typed via SendText (no clipboard exposure) ---
        Sleep(30)
        SendText(FieldOrder[3].value)
        Sleep(50)

        ; --- InterConnect state selection ---
        if (!IsHomeState && IsInterconnectState) {
            ; v2.1.6: Predict how many false matches Chrome's Find
            ; will hit before the correct option, so we can advance
            ; past them deterministically with Find Next ({Enter}).
            ; Uses the already-scraped PageText — no extra clipboard
            ; trip, no extra latency in the clean case.
            TargetPos := LocateStateTargetPos(PageText, p_StateName, InterconnectPos, FooterPos)
            SkipCount := 0
            if (TargetPos > 0)
                SkipCount := CountSubstringMatchesBefore(PageText, p_StateName, TargetPos)
            PageText := ScrubString(PageText)
            PageText := ""

            Send("^f")
            Sleep(75)                    ; v2.1.4: 50->75ms
            SendText(p_StateName)
            Sleep(50)
            Loop SkipCount {
                Send("{Enter}")          ; v2.1.6: Find Next — advance past false matches
                Sleep(50)
            }
            Send("{Esc}")
            Sleep(75)                    ; v2.1.4: 20->75ms
            Send("+{Tab}")
            Sleep(30)
            Send("{Space}")
            Sleep(75)                    ; v2.1.4: 30->75ms — radio button state must register
        }

        ; --- Agreement + Submit ---
        Send("^{End}")
        Sleep(75)
        if (IsFloridaPDMP) {
            ; FL has no separate checkbox — "By clicking search, you agree"
            Send("+{Tab}")
            Sleep(30)
            Send("{Enter}")
        } else {
            Send("^f")
            Sleep(75)
            SendText(AgreementKeyword)
            Sleep(50)
            Send("{Esc}")
            Sleep(75)
            Send("+{Tab}")
            Sleep(30)
            Send("{Space}")
            Sleep(30)
            Send("^f")
            Sleep(75)
            SendText("search")
            Sleep(50)
            Send("{Esc}")
            Sleep(75)
            Send("{Space}")
        }
        return true
    }
}

; =====================================================================
;  FUNCTION: Scrub PDMP PHI variables (HIPAA Compliance)
; =====================================================================
ScrubPHI(*) {
    global p_FirstName, p_LastName, p_DOB, p_StateName, g_BatchRunning
    p_FirstName    := ""
    p_LastName     := ""
    p_DOB          := ""
    p_StateName    := ""
    g_BatchRunning := false
    A_Clipboard    := ""
}

; =====================================================================
;  REFUSAL TO FILL — MED SELECTION GUI CALLBACKS
; =====================================================================

RefusalMedOK(btn, info) {
    global g_RefusalMedList, g_RefusalSelectedIdx
    global p_MedName, p_SelectedNDC, p_SelectedRxNumber
    global g_RefusalGUIDropCount, g_RefusalGUIAction

    g_RefusalGUIAction := "ok"
    saved := btn.Gui.Submit()

    selectedMeds := []
    Loop g_RefusalGUIDropCount {
        propName := "DDL" . A_Index
        val := saved.%propName%
        If (val == "" || val == "None" || val == "-- Select a medication --")
            Continue
        For idx, med in g_RefusalMedList {
            If (med.name . "  |  NDC: " . med.ndc == val) {
                selectedMeds.Push(med)
                Break
            }
        }
    }

    p_MedName          := ""
    p_SelectedNDC      := ""
    p_SelectedRxNumber := ""
    For i, med in selectedMeds {
        sep := (i > 1) ? " & " : ""
        p_MedName          .= sep . med.name
        p_SelectedNDC      .= sep . med.ndc
        p_SelectedRxNumber .= sep . med.rx
    }
    g_RefusalSelectedIdx := selectedMeds.Length
    Try btn.Gui.Destroy()
}

RefusalMedCancel(btn, info := "") {
    global g_RefusalSelectedIdx, g_RefusalGUIAction
    If g_RefusalGUIAction == "add"
        Return
    g_RefusalGUIAction    := "cancel"
    g_RefusalSelectedIdx  := -1
    If (btn is Gui) {
        Try btn.Destroy()
    } Else {
        Try btn.Gui.Destroy()
    }
}

RefusalMedAdd(btn, info) {
    global g_RefusalGUIAction, g_RefusalGUISaved, g_RefusalGUIDropCount, g_RefusalGUIDDLs
    g_RefusalGUIAction := "add"
    g_RefusalGUISaved  := []

    If IsObject(g_RefusalGUIDDLs) {
        For ddl in g_RefusalGUIDDLs {
            txt := ""
            Try txt := ddl.Text
            g_RefusalGUISaved.Push(txt)
        }
    }

    While (g_RefusalGUISaved.Length < g_RefusalGUIDropCount)
        g_RefusalGUISaved.Push("")

    g_RefusalGUIDDLs := []
    Try btn.Gui.Destroy()
}

; =====================================================================
;  F12 — COMBINED EMERGENCY STOP
;  Disarms the Refusal watchdog, wipes all PHI from both scripts,
;  closes any open dialogs, and reloads the combined script cleanly.
; =====================================================================
F12:: {
    StopWatchdog()
    WipeAllData()
    CloseAllKnownDialogs()
    MsgBox("EMERGENCY STOP ACTIVATED.`n`nAll sensitive data has been wiped.", "Emergency Stop", 48)
    Reload()
}

; =====================================================================
;  ^+!B — PDMP BATCH LOOP   (Genovation Key 3)
; =====================================================================
^+!b:: {
    global p_FirstName, p_LastName, p_DOB, p_StateName, g_BatchRunning, g_PostWipeClipboard

    if (g_BatchRunning) {
        MsgBox("STOPPING: The batch is already running!`nPress F12 to terminate it first.")
        return
    }
    g_BatchRunning := true

    try {
        Loop {
            if WinExist("Admin - PillPack") {
                WinActivate("Admin - PillPack")
                if !WinWaitActive("Admin - PillPack", , 1) {
                    MsgBox("STOPPING: PillPack window did not focus in time.")
                    break
                }
            } else {
                MsgBox("STOPPING: Could not find the PillPack window.")
                break
            }

            if !ScrapePillPack()
                break

            ; State-restricted patients — generic dialog (no PHI in message)
            if (p_StateName = "Texas") {
                MsgBox("STOPPING: Restricted state. A TX-licensed pharmacist must utilize the TX PDMP database to perform this search.")
                break
            } else if (p_StateName = "Missouri") {
                g_PostWipeClipboard := "MO PDMP currently unavailable, approving shipment based on internal clinical review."
                MsgBox("STOPPING: MO Customer — MO PDMP only available to a licensed MO RPH's, if unavailable then approve shipment based on internal clinical review. Perform an internal review by checking previous shipments, then paste the shipment note from your clipboard.`n`nIf a Refusal to Fill or Pause is applicable, follow that process instead.", "MO Customer", 48)
                break
            }

            if WinExist("ahk_group PDMPGroup") {
                WinActivate("ahk_group PDMPGroup")
                if !WinWaitActive("ahk_group PDMPGroup", , 1) {
                    MsgBox("STOPPING: PDMP window did not focus in time.")
                    break
                }
            } else {
                MsgBox("STOPPING: Could not find the PDMP window. Ensure it is open.")
                break
            }

            if !PastePDMP()
                break

            Send("^{Tab}")
            Sleep(30)

            WinActivate("Admin - PillPack")
            if !WinWaitActive("Admin - PillPack", , 1) {
                MsgBox("STOPPING: PillPack window did not focus in time after PDMP submission.")
                break
            }
            Send("^{Tab}")
            Sleep(30)
        }
    } finally {
        ; Always scrub PHI — runs whether we broke cleanly or an error escaped
        ScrubPHI()
        if (g_PostWipeClipboard != "") {
            A_Clipboard := g_PostWipeClipboard
            g_PostWipeClipboard := ""
        }
    }
}

; =====================================================================
;  ^+!D — DUPLICATE TAB (PDMP)
; =====================================================================
^+!d:: {
    CoordMode("Mouse", "Screen")
    MouseGetPos(&mouseX, &mouseY)

    userInput := InputBox("How many times do you want to duplicate this tab?", "Duplicate Tab")
    if (userInput.Result = "Cancel")
        return

    numCopies := userInput.Value
    if IsInteger(numCopies) {
        numCopies := Integer(numCopies)
        if (numCopies > 0) {
            Loop numCopies {
                MouseClick("Right", mouseX, mouseY)
                Sleep(150)
                Send("{Down 6}")
                Sleep(30)
                Send("{Enter}")
                Sleep(40)
            }
        } else {
            MsgBox("Please enter a number greater than 0.", "Invalid Input")
        }
    } else {
        MsgBox("Please enter a valid whole number.", "Invalid Input")
    }
}

; =====================================================================
;  ^+!R — REFUSAL TO FILL WORKFLOW (v2.2.0)
;
;  Flow:
;    1. Scrape PillPack page (URL, name, ship/arrival dates, meds)
;    2. If >1 med or packet meds exist: multi-med warning + selection GUI
;    3. Input form GUI (fill dates, pharmacy)
;    4. Open customer details tab, extract Batch
;    5. AMAZON path: unchanged (Wolfgang template)
;    6. NON-AMAZON path: Tachyon "Refusal to Fill" ticket
; =====================================================================
^+!r:: {
    global p_PatientFullName, p_MedName, p_Batch, p_Identifier, p_OrderID
    global p_LastFillDate, p_LastFillDays, p_LastPharmacy, p_NextFillDate, g_FormSubmitted
    global p_ShipDate, p_ArrivalDate, p_ArrivalParsed, p_ShipmentURL
    global p_PillPackProfileID, p_SelectedNDC, p_SelectedRxNumber, p_AdminURL, p_CustomerPhone, p_RTFReason, p_RTFReasonDetail
    global g_DetailVisible, g_RTFIsEarlyControl, g_RTFDecision
    global g_RefusalMedList, g_RefusalSelectedIdx
    global g_RefusalGUIDropCount, g_RefusalGUIAction, g_RefusalGUISaved, g_RefusalGUIDDLs

    StartWatchdog()

    ; -----------------------------------------------------------------
    ; STEP 1: Capture the current URL from the address bar
    ; -----------------------------------------------------------------
    If !CaptureURLFromAddressBar(&p_ShipmentURL) {
        MsgBox("STOPPING: Could not read a valid URL from Chrome address bar.", "Error", 16)
        WipeAllData()
        Return
    }

    If !(InStr(p_ShipmentURL, "shipment") || InStr(p_ShipmentURL, "pillpack") || InStr(p_ShipmentURL, "amazon")) {
        MsgBox("STOPPING: This does not appear to be a shipment page.", "Wrong Page", 16)
        WipeAllData()
        Return
    }

    ; -----------------------------------------------------------------
    ; STEP 2: Scrape the PillPack page
    ; -----------------------------------------------------------------
    Send("!d")
    Sleep(50)
    SendText("javascript:window.focus();void(0);")
    Sleep(50)
    Send("{Enter}")
    Sleep(200)

    PageText := ""

    If !WaitForShipmentPage(&PageText) {
        MsgBox("STOPPING: Could not read the shipment page content.", "Error", 16)
        WipeAllData()
        Return
    }

    ; -----------------------------------------------------------------
    ; STEP 3: Extract Patient Name, Ship Date, Arrival Date
    ; -----------------------------------------------------------------
    p_PatientFullName := ""
    p_ShipDate        := ""
    p_ArrivalDate     := ""
    p_ArrivalParsed   := ""

    If RegExMatch(PageText, "i)([a-zA-Z\-']+(?:[ \t]+[a-zA-Z\-']+)*)\s*(?:\([^)]+\))?\s*[-—–‐]\s*(?:Unscheduled|Scheduled)", &NameMatch) {
        p_PatientFullName := Trim(NameMatch[1])
        p_PatientFullName := RegExReplace(p_PatientFullName, "\s+", " ")
    }

    If RegExMatch(PageText, "i)Ship\s*Date[:\s]+([A-Za-z,\s\d]+?)(?:\r|\n|Arrival)", &ShipMatch) {
        RawShipDate := Trim(ShipMatch[1])
        ShipParsed  := ParseDateFromText(RawShipDate)
        p_ShipDate  := ShipParsed.formatted
    }

    If RegExMatch(PageText, "i)Arrival\s*Date[:\s]+([A-Za-z,\s\d]+?)(?:\s*\(|Shipping|\r|\n)", &ArrivalMatch) {
        RawArrivalDate  := Trim(ArrivalMatch[1])
        p_ArrivalParsed := ParseDateFromText(RawArrivalDate)
        p_ArrivalDate   := p_ArrivalParsed.formatted
    }

    If (p_PatientFullName = "") {
        PageText := ScrubString(PageText)
        PageText := ""
        MsgBox("STOPPING: Could not find the patient's name on this page.", "Extraction Failed", 16)
        WipeAllData()
        Return
    }

    ; -----------------------------------------------------------------
    ; STEP 4: Parse medications from Bulk and Packet sections
    ; -----------------------------------------------------------------
    g_RefusalMedList := []

    ; --- Bulk meds ---
    BulkStart := InStr(PageText, "Dispensed Bulk Meds")
    If (BulkStart > 0) {
        BulkEnd := InStr(PageText, "Pack Insert Prompts", , BulkStart)
        BulkSection := (BulkEnd > 0)
            ? SubStr(PageText, BulkStart, BulkEnd - BulkStart)
            : SubStr(PageText, BulkStart)

        startPos := 1
        While RegExMatch(BulkSection, "([^\r\n]+?) NDC:\s*(\d{11})", &MedMatch, startPos) {
            medName := Trim(MedMatch[1])
            ndc     := MedMatch[2]

            If InStr(medName, "Med Description") || InStr(medName, "Description") || StrLen(medName) < 3 {
                startPos := MedMatch.Pos + MedMatch.Len
                Continue
            }

            followRaw  := SubStr(BulkSection, MedMatch.Pos + MedMatch.Len, 600)
            nextNDCPos := RegExMatch(followRaw, "\bNDC:\s*\d{11}")
            followText := (nextNDCPos > 0) ? SubStr(followRaw, 1, nextNDCPos - 1) : followRaw

            rxNum := ""
            If RegExMatch(followText, "(\d{7,10})\s*/", &RxMatch)
                rxNum := RxMatch[1]
            Else If RegExMatch(followText, "Rx\s*#?\s*:?\s*(\d{7,10})", &RxMatch)
                rxNum := RxMatch[1]
            Else If RegExMatch(followText, "(?<!\d)(\d{7,10})(?!\d)", &RxMatch)
                rxNum := RxMatch[1]

            isHold := (InStr(followText, "Held by Wms Rx Updates Bot") > 0)
            medName := Trim(RegExReplace(medName, "\s*\([^)]*\)\s*$", ""))
            medName := RegExReplace(medName, "\s+", " ")
            ; Deduplicate by NDC
            isDupe := false
            For existing in g_RefusalMedList {
                If (existing.ndc == ndc) {
                    isDupe := true
                    Break
                }
            }
            If (!isDupe)
                g_RefusalMedList.Push({name: medName, ndc: ndc, rx: rxNum, isHold: isHold})

            startPos := MedMatch.Pos + MedMatch.Len
        }
        BulkSection := ScrubString(BulkSection)
        BulkSection := ""
    }

    ; --- Packet meds ---
    HasPacketMeds := false
    PacketStart := InStr(PageText, "Dispensed Packet Meds")
    If (PacketStart > 0) {
        HasPacketMeds := true
        PacketEnd := InStr(PageText, "Expected Bulk Meds", , PacketStart)
        PacketSection := (PacketEnd > 0)
            ? SubStr(PageText, PacketStart, PacketEnd - PacketStart)
            : SubStr(PageText, PacketStart)

        startPos := 1
        While RegExMatch(PacketSection, "([^\r\n]+?) NDC:\s*(\d{11})", &MedMatch, startPos) {
            medName := Trim(MedMatch[1])
            ndc     := MedMatch[2]

            If InStr(medName, "Med Description") || InStr(medName, "Description") || StrLen(medName) < 3 {
                startPos := MedMatch.Pos + MedMatch.Len
                Continue
            }

            followRaw  := SubStr(PacketSection, MedMatch.Pos + MedMatch.Len, 600)
            nextNDCPos := RegExMatch(followRaw, "\bNDC:\s*\d{11}")
            followText := (nextNDCPos > 0) ? SubStr(followRaw, 1, nextNDCPos - 1) : followRaw

            rxNum := ""
            If RegExMatch(followText, "(\d{7,10})\s*/", &RxMatch)
                rxNum := RxMatch[1]
            Else If RegExMatch(followText, "Rx\s*#?\s*:?\s*(\d{7,10})", &RxMatch)
                rxNum := RxMatch[1]
            Else If RegExMatch(followText, "(?<!\d)(\d{7,10})(?!\d)", &RxMatch)
                rxNum := RxMatch[1]

            medName := Trim(RegExReplace(medName, "\s*\([^)]*\)\s*$", ""))
            medName := RegExReplace(medName, "\s+", " ")
            ; Deduplicate by NDC
            isDupe := false
            For existing in g_RefusalMedList {
                If (existing.ndc == ndc) {
                    isDupe := true
                    Break
                }
            }
            If (!isDupe)
                g_RefusalMedList.Push({name: medName, ndc: ndc, rx: rxNum, isHold: false})

            startPos := MedMatch.Pos + MedMatch.Len
        }
        PacketSection := ScrubString(PacketSection)
        PacketSection := ""
    }

    PageText := ScrubString(PageText)
    PageText := ""

    If (g_RefusalMedList.Length == 0) {
        MsgBox("STOPPING: No medications with NDC numbers were found on this page.", "No Meds Found", 16)
        WipeAllData()
        Return
    }

    ; -----------------------------------------------------------------
    ; STEP 5: Multi-med / Packet detection -> Med Selection GUI
    ; If only 1 med and no packets: auto-select it, skip GUI.
    ; -----------------------------------------------------------------
    If (g_RefusalMedList.Length == 1 && !HasPacketMeds) {
        med := g_RefusalMedList[1]
        p_MedName          := med.name
        p_SelectedNDC      := med.ndc
        p_SelectedRxNumber := med.rx
        g_RefusalSelectedIdx := 1
    } Else {
        ; Show warning about multiple medications
        PauseStart()
        SplitResult := MsgBox("This shipment has more than 1 medication (or has packet medications attached).`n`nEnsure other medications should not be split off before continuing.`n`nPress OK to continue, or Cancel to abort.", "Multiple Medications Detected", 49)
        PauseEnd()
        If (SplitResult = "Cancel") {
            WipeAllData()
            Return
        }
        ReactivateChrome()

        ; Build med display strings
        medDisplayStrings := []
        For idx, med in g_RefusalMedList
            medDisplayStrings.Push(med.name . "  |  NDC: " . med.ndc)

        ; Detect hold meds for auto-selection
        holdMeds := []
        For idx, med in g_RefusalMedList {
            If med.isHold
                holdMeds.Push(idx)
        }
        holdCount := holdMeds.Length

        items_required := []
        If holdCount == 0
            items_required.Push("-- Select a medication --")
        For s in medDisplayStrings
            items_required.Push(s)

        items_optional := ["None"]
        For s in medDisplayStrings
            items_optional.Push(s)

        canAddMore := (g_RefusalMedList.Length > 1)

        g_RefusalGUIAction    := ""
        g_RefusalGUIDropCount := Max(holdCount, 1)
        g_RefusalGUISaved     := []
        g_RefusalGUIDDLs      := []
        g_RefusalSelectedIdx  := 0

        ; ---- GUI build/rebuild loop ----
        While true {
            MedGui := Gui("+AlwaysOnTop", "Refusal to Fill - Medication Selection")
            MedGui.SetFont("s11 Bold")
            MedGui.Add("Text", "w490 Center", "Refusal to Fill - Select Medication")
            MedGui.SetFont("s10 Norm")
            MedGui.Add("Text", "w490", "")

            If (g_RefusalGUIDropCount == 1 && holdCount == 0)
                MedGui.Add("Text", "w490", "Select the medication to complete the refusal for:")
            Else If (g_RefusalGUIDropCount == 1 && holdCount == 1)
                MedGui.Add("Text", "w490", "Inventory hold detected. Medication auto-selected - confirm and click OK:")
            Else
                MedGui.Add("Text", "w490", "Select the medication(s) to complete the refusal for:")

            g_RefusalGUIDDLs := []
            Loop g_RefusalGUIDropCount {
                i       := A_Index
                isFirst := (i == 1)
                items   := isFirst ? items_required : items_optional

                ddl := MedGui.Add("DropDownList", "w490 vDDL" . i, items)
                g_RefusalGUIDDLs.Push(ddl)

                selText := ""
                If (i <= g_RefusalGUISaved.Length && g_RefusalGUISaved[i] != "")
                    selText := g_RefusalGUISaved[i]
                Else If (i <= holdMeds.Length) {
                    hm      := g_RefusalMedList[holdMeds[i]]
                    selText := hm.name . "  |  NDC: " . hm.ndc
                }

                If (selText != "") {
                    For k, s in items {
                        If (s == selText) {
                            ddl.Choose(k)
                            Break
                        }
                    }
                } Else {
                    ddl.Choose(1)
                }
            }

            MedGui.Add("Text", "w490", "")

            BtnOK     := MedGui.Add("Button", "w100", "OK")
            BtnCancel := MedGui.Add("Button", "x+20 w100", "Cancel")
            If (canAddMore && g_RefusalGUIDropCount < g_RefusalMedList.Length)
                BtnAdd := MedGui.Add("Button", "x+20 w150", "Add Medication")

            BtnOK.OnEvent("Click", RefusalMedOK)
            BtnCancel.OnEvent("Click", RefusalMedCancel)
            If (canAddMore && g_RefusalGUIDropCount < g_RefusalMedList.Length)
                BtnAdd.OnEvent("Click", RefusalMedAdd)
            MedGui.OnEvent("Close", GUICrash)

            g_RefusalGUIAction := ""
            PauseStart()
            MedGui.Show()
            While WinExist("Refusal to Fill - Medication Selection")
                Sleep(100)
            PauseEnd()

            If g_RefusalGUIAction == "add" {
                g_RefusalGUIDropCount++
                Continue
            }
            Break
        }

        ; Post-GUI validation
        If (g_RefusalSelectedIdx == -1) {
            WipeAllData()
            Return
        }

        If (p_MedName == "" || p_SelectedNDC == "") {
            MsgBox("No medication was selected. Aborting.", "Error", 16)
            WipeAllData()
            Return
        }
    }

    If (p_SelectedRxNumber == "") {
        PauseStart()
        MsgBox("WARNING: Could not extract an Rx number for the selected medication.`n`nYou will need to fill the Rx# field manually on the ticket.", "Extraction Warning", 48)
        PauseEnd()
        ReactivateChrome()
    }

    ; p_SelectedNDC is not used after med selection; clear now
    Try p_SelectedNDC := ScrubString(p_SelectedNDC)
    p_SelectedNDC := ""

    ; -----------------------------------------------------------------
    ; STEP 6: Input Form GUI (fill dates, pharmacy)
    ; -----------------------------------------------------------------
    g_FormSubmitted := false
    g_DetailVisible := false

    InputGui := Gui(, "Refusal to Fill - Input Required")
    InputGui.SetFont("s10")

    InputGui.Add("Text", "w450", "Cancel " p_MedName " due to refusal to fill?")
    InputGui.Add("Text", "w450", "")

    InputGui.Add("Text", "w450", "RTF - Reason:")
    DDLReason := InputGui.Add("DropDownList", "vRTFReason w300", g_RTFReasonOptions)
    BtnAddDetail := InputGui.Add("Button", "x+10 w100", "Add Detail")
    DetailLabel := InputGui.Add("Text", "xm w450 Hidden", "Detail (appended to reason):")
    DetailEdit  := InputGui.Add("Edit", "vRTFDetail w450 Hidden", "")

    InputGui.Add("Text", "xm w450", "")
    InputGui.Add("Text", "w450", "Last Fill Date:")
    InputGui.Add("DateTime", "vLastFillDate w200 ChooseNone", "MM/dd/yyyy")

    InputGui.Add("Text", "w450", "")
    InputGui.Add("Text", "w450", "Last Fill Days Supply:")
    InputGui.Add("Edit", "vLastFillDays w100", "")

    InputGui.Add("Text", "w450", "")
    InputGui.Add("Text", "w450", "Last Pharmacy Filled At:")
    InputGui.Add("Edit", "vLastPharmacy w300", "")

    InputGui.Add("Text", "w450", "")
    NextFillLabel := InputGui.Add("Text", "w450 Hidden", "Next Available Fill Date:")
    NextFillCtrl  := InputGui.Add("DateTime", "vNextFillDate w200 Hidden ChooseNone", "MM/dd/yyyy")

    InputGui.Add("Text", "w450", "")

    BtnCalc := InputGui.Add("Button", "w150", "Calculate Next Fill")
    BtnSubmit := InputGui.Add("Button", "x+10 Default w100", "Continue")
    BtnCancel := InputGui.Add("Button", "x+10 w100", "Cancel")

    DDLReason.OnEvent("Change", OnReasonChange)
    BtnAddDetail.OnEvent("Click", OnAddDetail)
    BtnCalc.OnEvent("Click", OnCalcNextFill)
    BtnSubmit.OnEvent("Click", OnSubmit)
    BtnCancel.OnEvent("Click", OnCancel)
    InputGui.OnEvent("Close", GUICrash)

    OnReasonChange(ctrl, info) {
        global g_DetailVisible
        if (ctrl.Text != "Early Control" && ctrl.Text != "") {
            if (!g_DetailVisible) {
                DetailLabel.Visible := true
                DetailEdit.Visible  := true
                g_DetailVisible := true
                BtnAddDetail.Text := "Remove Detail"
            }
        } else {
            if (g_DetailVisible) {
                DetailLabel.Visible := false
                DetailEdit.Visible  := false
                DetailEdit.Value    := ""
                g_DetailVisible := false
                BtnAddDetail.Text := "Add Detail"
            }
        }
    }

    OnAddDetail(btn, info) {
        global g_DetailVisible
        If (!g_DetailVisible) {
            DetailLabel.Visible := true
            DetailEdit.Visible  := true
            g_DetailVisible := true
            btn.Text := "Remove Detail"
        } Else {
            DetailLabel.Visible := false
            DetailEdit.Visible  := false
            DetailEdit.Value    := ""
            g_DetailVisible := false
            btn.Text := "Add Detail"
        }
    }

    OnCalcNextFill(btn, info) {
        InputGui.Submit(false)
        RawFillDate := InputGui["LastFillDate"].Value
        RawDays     := InputGui["LastFillDays"].Value

        If (RawFillDate = "") {
            MsgBox("Please select a Last Fill Date first.", "Validation Error", 48)
            return
        }
        If !RegExMatch(RawDays, "^\d+$") {
            MsgBox("Please enter a valid Days Supply first.", "Validation Error", 48)
            return
        }

        DaysSupply := Integer(RawDays)
        EarlyDays := Integer(DaysSupply * 0.20)
        If (EarlyDays > 7)
            EarlyDays := 7
        ElapsedNeeded := DaysSupply - EarlyDays

        FillYear  := Integer(SubStr(RawFillDate, 1, 4))
        FillMonth := Integer(SubStr(RawFillDate, 5, 2))
        FillDay   := Integer(SubStr(RawFillDate, 7, 2))

        DaysInMonth := [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        ResultYear  := FillYear
        ResultMonth := FillMonth
        ResultDay   := FillDay + ElapsedNeeded

        Loop {
            IsLeap := (Mod(ResultYear, 4) == 0 && (Mod(ResultYear, 100) != 0 || Mod(ResultYear, 400) == 0))
            MaxDay := DaysInMonth[ResultMonth]
            If (ResultMonth == 2 && IsLeap)
                MaxDay := 29
            If (ResultDay <= MaxDay)
                Break
            ResultDay -= MaxDay
            ResultMonth++
            If (ResultMonth > 12) {
                ResultMonth := 1
                ResultYear++
            }
        }

        NextFillStamp := Format("{:04}{:02}{:02}000000", ResultYear, ResultMonth, ResultDay)
        NextFillCtrl.Value := NextFillStamp

        NextFillLabel.Visible := true
        NextFillCtrl.Visible  := true
        InputGui.Flash()
    }

    OnSubmit(btn, info) {
        global p_LastFillDate, p_LastFillDays, p_LastPharmacy, p_NextFillDate, p_RTFReason, p_RTFReasonDetail, g_FormSubmitted, g_DetailVisible

        Saved := btn.Gui.Submit()

        ; Validate RTF Reason selection
        if (Saved.RTFReason = "") {
            MsgBox("Please select an RTF - Reason.", "Validation Error", "Icon!")
            btn.Gui.Show()
            return
        }

        IsEarlyControl := (Saved.RTFReason = "Early Control")

        ; Non-Early Control reasons require Detail
        If (!IsEarlyControl && g_DetailVisible && Trim(Saved.RTFDetail) = "") {
            MsgBox("Please enter a Detail for this reason.", "Validation Error", "Icon!")
            btn.Gui.Show()
            return
        }
        If (!IsEarlyControl && !g_DetailVisible) {
            MsgBox("Please click 'Add Detail' and enter a reason detail.", "Validation Error", "Icon!")
            btn.Gui.Show()
            return
        }

        ; Convert DateTime values to MM/DD/YYYY strings
        LastFillVal := ""
        If (Saved.LastFillDate != "")
            LastFillVal := FormatTime(Saved.LastFillDate, "MM/dd/yyyy")

        NextFillVal := ""
        If (Saved.NextFillDate != "")
            NextFillVal := FormatTime(Saved.NextFillDate, "MM/dd/yyyy")

        ; Early Control requires all fill fields
        If (IsEarlyControl) {
            if (LastFillVal = "") {
                MsgBox("Please select a Last Fill Date.", "Validation Error", "Icon!")
                btn.Gui.Show()
                return
            }

            if !RegExMatch(Saved.LastFillDays, "^\d+$") {
                MsgBox("Invalid Days Supply. Please enter a number.", "Validation Error", "Icon!")
                btn.Gui.Show()
                return
            }

            if (Trim(Saved.LastPharmacy) = "") {
                MsgBox("Please enter the Last Pharmacy.", "Validation Error", "Icon!")
                btn.Gui.Show()
                return
            }

            if (NextFillVal = "") {
                MsgBox("Please select a Next Available Fill Date.", "Validation Error", "Icon!")
                btn.Gui.Show()
                return
            }
        }

        p_RTFReason       := Saved.RTFReason
        p_RTFReasonDetail := Trim(Saved.RTFDetail)
        p_LastFillDate    := LastFillVal
        p_LastFillDays    := Saved.LastFillDays
        p_LastPharmacy    := Saved.LastPharmacy
        p_NextFillDate    := NextFillVal
        g_FormSubmitted   := true

        btn.Gui.Destroy()
    }

    OnCancel(btn, info := "") {
        global g_FormSubmitted
        g_FormSubmitted := false

        if (btn is Gui)
            btn.Destroy()
        else
            btn.Gui.Destroy()
    }

    PauseStart()
    InputGui.Show()
    while WinExist("Refusal to Fill - Input Required")
        Sleep(100)
    PauseEnd()

    if (!g_FormSubmitted) {
        WipeAllData()
        return
    }

    ReactivateChrome()

    ; Append detail to RTF reason if provided
    If (p_RTFReasonDetail != "")
        p_RTFReason .= " - " . p_RTFReasonDetail

    ; Track whether this is Early Control (determines Amazon path behavior)
    g_RTFIsEarlyControl := (p_RTFReason = "Early Control")

    ; -----------------------------------------------------------------
    ; STEP 7: Open Customer Name in New Tab
    ; -----------------------------------------------------------------
    OpenCustomerInNewTab(p_PatientFullName)
    Try p_PatientFullName := ScrubString(p_PatientFullName)
    p_PatientFullName := ""

    ; -----------------------------------------------------------------
    ; STEP 8: Wait for Details Tab to Load & Scrape
    ; -----------------------------------------------------------------
    DetailsText := ""

    if !WaitForDetailsPage(&DetailsText) {
        DetailsText := ""
        MsgBox("STOPPING: The new Details tab took too long to load.", "Timeout", 16)
        WipeAllData()
        return
    }

    ; -----------------------------------------------------------------
    ; STEP 9: Grab the Admin URL
    ; -----------------------------------------------------------------
    Send("!d")
    Sleep(100)
    A_Clipboard := ""
    Send("^c")
    Sleep(50)
    If ClipWait(1) {
        p_AdminURL  := A_Clipboard
        A_Clipboard := ""
    }
    Send("{Esc}")
    Sleep(50)

    ; -----------------------------------------------------------------
    ; STEP 10: Extract Batch, Phone, and Identifier
    ; -----------------------------------------------------------------
    p_Batch         := ""
    p_Identifier    := ""
    p_CustomerPhone := ""

    if RegExMatch(DetailsText, "i)Batch[\s:]+([A-Za-z0-9_]+)", &BatchMatch) {
        p_Batch := Trim(BatchMatch[1])
    }

    ; Extract phone number (format on admin page: xxx-xxx-xxxx)
    if RegExMatch(DetailsText, "(\d{3})-(\d{3})-(\d{4})", &PhoneMatch) {
        p_CustomerPhone := PhoneMatch[1] . PhoneMatch[2] . PhoneMatch[3]
    }

    if (p_Batch = "") {
        DetailsText := ScrubString(DetailsText)
        DetailsText := ""
        MsgBox("STOPPING: Could not locate the 'Batch' type.", "Extraction Failed", 16)
        WipeAllData()
        return
    }

    if (p_Batch = "AMAZON") {
        if RegExMatch(DetailsText, "i)Person ID[\s:]+([^\r\n]+)", &IDMatch) {
            p_Identifier := Trim(IDMatch[1])
        }
        DetailsText := ScrubString(DetailsText)
        DetailsText := ""
        ; Amazon path never uses AdminURL; clear it now
        Try p_AdminURL := ScrubString(p_AdminURL)
        p_AdminURL := ""
    } else {
        DetailsText := ScrubString(DetailsText)
        DetailsText := ""
        ; Extract PillPack Profile ID from Admin URL
        CleanURL := p_AdminURL
        QPos := InStr(CleanURL, "?")
        If QPos
            CleanURL := SubStr(CleanURL, 1, QPos - 1)
        HPos := InStr(CleanURL, "#")
        If HPos
            CleanURL := SubStr(CleanURL, 1, HPos - 1)
        LastSlash := InStr(CleanURL, "/", false, -1)
        If LastSlash > 0
            p_PillPackProfileID := Trim(SubStr(CleanURL, LastSlash + 1))
        CleanURL := ""

        if RegExMatch(p_AdminURL, "([^/]+)/?$", &UrlMatch) {
            p_Identifier := Trim(UrlMatch[1])
        }
        ; AdminURL no longer needed after extraction
        Try p_AdminURL := ScrubString(p_AdminURL)
        p_AdminURL := ""
    }

    if (p_Identifier = "" && p_Batch = "AMAZON") {
        MsgBox("STOPPING: Could not locate the Person ID.", "Extraction Failed", 16)
        WipeAllData()
        return
    }

    if (p_Batch != "AMAZON" && p_PillPackProfileID = "") {
        MsgBox("STOPPING: Could not extract the PillPack Profile ID from the admin URL.", "Extraction Failed", 16)
        WipeAllData()
        Return
    }

    ; =================================================================
    ; BRANCH: AMAZON vs NON-AMAZON
    ; =================================================================

    if (p_Batch = "AMAZON") {

        ; =============================================================
        ; AMAZON PATH — NON-EARLY CONTROL DECISION POINT
        ; If reason is NOT Early Control, ask user: Tachyon Ticket or Cancel
        ; =============================================================
        If (!g_RTFIsEarlyControl) {
            g_RTFDecision := ""

            DecisionGui := Gui("+AlwaysOnTop", "RTF - Non-Early Control Decision")
            DecisionGui.SetFont("s10")
            DecisionGui.Add("Text", "w600", "RTF - Reason other than 'Early Control' was selected for an Amazon Pharmacy customer.")
            DecisionGui.Add("Text", "w600", "")
            DecisionGui.Add("Text", "w600", "Select an action:")
            DecisionGui.Add("Text", "w600", "")

            BtnTachyon := DecisionGui.Add("Button", "w180", "Clinical Review Ticket")
            BtnCallback := DecisionGui.Add("Button", "x+10 w220", "Prescriber/Pharmacy Callback")
            BtnCancelOrder := DecisionGui.Add("Button", "x+10 w140", "Cancel Order")
            BtnAbort := DecisionGui.Add("Button", "xm w100", "Abort")

            BtnTachyon.OnEvent("Click", OnDecisionTachyon)
            BtnCallback.OnEvent("Click", OnDecisionCallback)
            BtnCancelOrder.OnEvent("Click", OnDecisionCancel)
            BtnAbort.OnEvent("Click", OnDecisionAbort)
            DecisionGui.OnEvent("Close", GUICrash)

            OnDecisionTachyon(btn, info) {
                global g_RTFDecision
                g_RTFDecision := "tachyon"
                btn.Gui.Destroy()
            }

            OnDecisionCallback(btn, info) {
                global g_RTFDecision
                g_RTFDecision := "callback"
                btn.Gui.Destroy()
            }

            OnDecisionCancel(btn, info) {
                global g_RTFDecision
                g_RTFDecision := "cancel"
                btn.Gui.Destroy()
            }

            OnDecisionAbort(btn, info := "") {
                global g_RTFDecision
                g_RTFDecision := "abort"
                If (btn is Gui)
                    btn.Destroy()
                Else
                    btn.Gui.Destroy()
            }

            PauseStart()
            DecisionGui.Show()
            While WinExist("RTF - Non-Early Control Decision")
                Sleep(100)
            PauseEnd()

            If (g_RTFDecision == "abort") {
                WipeAllData()
                Return
            }

            ReactivateChrome()

            If (g_RTFDecision == "tachyon") {
                ; ==========================================================
                ; TACHYON TICKET PATH: Submit Clinical Review ticket
                ; with FFU - Reason for Change = "Clinical Review Requested"
                ; ==========================================================

                ; Scrub data not needed for this path (Early Control / cancel only)
                Try p_LastFillDate := ScrubString(p_LastFillDate)
                p_LastFillDate := ""
                Try p_LastFillDays := ScrubString(p_LastFillDays)
                p_LastFillDays := ""
                Try p_LastPharmacy := ScrubString(p_LastPharmacy)
                p_LastPharmacy := ""
                Try p_NextFillDate := ScrubString(p_NextFillDate)
                p_NextFillDate := ""
                Try p_CustomerPhone := ScrubString(p_CustomerPhone)
                p_CustomerPhone := ""
                Try p_SelectedNDC := ScrubString(p_SelectedNDC)
                p_SelectedNDC := ""
                Try p_RTFReasonDetail := ScrubString(p_RTFReasonDetail)
                p_RTFReasonDetail := ""
                p_Batch := ""

                ; Open Clinical Review (Bulk) blueprint directly
                Send("^t")
                Sleep(200)
                A_Clipboard := "https://cs.wolfgang.a2z.com/tachyon/PHARMACY/createNewIssue/9F695381-7AB0-4D30-B78C-5378605B465B"
                Sleep(50)
                Send("^v")
                Sleep(100)
                Send("{Enter}")
                A_Clipboard := ""

                ; Wait for the ticket form to load
                Loaded := false
                Loop 30 {
                    Sleep(A_Index = 1 ? 1500 : 500)
                    If CheckChromeTitleForLogin("Tachyon")
                        Continue
                    A_Clipboard := ""
                    Send("^a")
                    Sleep(50)
                    Send("^c")
                    If ClipWait(0.5) {
                        PageText := A_Clipboard
                        A_Clipboard := ""
                        If InStr(PageText, "Task: Red Bulk Shipment Change") {
                            PageText := ""
                            ClearSelection()
                            Loaded := true
                            break
                        }
                        If InStr(PageText, "Sign In") || InStr(PageText, "Username") || InStr(PageText, "Verify with your password") || InStr(PageText, "Powered by Okta") {
                            PageText := ""
                            PauseStart()
                            MsgBox("Sign-in required!`n`n1. Please log in manually.`n2. Click OK on this popup ONLY AFTER you see the ticket form.", "Login Paused", 64+262144)
                            PauseEnd()
                            Sleep(300)
                            WinActivate("ahk_exe chrome.exe")
                            Sleep(300)
                        }
                        PageText := ""
                        ClearSelection()
                    } Else {
                        ClearSelection()
                    }
                }
                If (!Loaded) {
                    MsgBox("STOPPING: Timed out waiting for the Clinical Review ticket form to load.", "Timeout", 16)
                    WipeAllData()
                    Return
                }

                ; Type Person ID and verify
                LocalID := p_Identifier
                Try p_Identifier := ScrubString(p_Identifier)
                p_Identifier := ""

                FindAndPasteField("Person ID", LocalID)
                If !VerifyFieldContent(LocalID) {
                    Try LocalID := ScrubString(LocalID)
                    LocalID := ""
                    MsgBox("Customer ID paste verification FAILED.`n`nAborting and wiping all data.", "Field Verify Failed", 16)
                    WipeAllData()
                    Return
                }
                Try LocalID := ScrubString(LocalID)
                LocalID := ""

                ; Select customer name
                If !SelectCustomerName(true) {
                    MsgBox("WARNING: Could not confirm the customer name was selected.`n`nVerify the 'Customer first name' field is populated before submitting.", "Customer Selection Warning", 48)
                    ReactivateChrome()
                }

                Sleep(300)

                ; Priority -> High if due today or earlier
                SetPriorityIfDueToday(p_ShipDate)

                ; Due By = Ship Date
                FindAndType("Due", p_ShipDate)

                ; FFU - Reason for Change = "Clinical Review Requested" (index 5)
                FindAndSelectDropdown("reason", 5)

                ; FFU - Original Promise Date (calendar navigation to arrival date)
                FindAndSetCalendar("FFU - Original Promise Date", p_ArrivalParsed)
                p_ArrivalParsed := ""
                p_ArrivalDate   := ""
                Try p_ShipDate := ScrubString(p_ShipDate)
                p_ShipDate := ""

                ; Medication Name
                FindAndPaste("Medication Name", p_MedName)

                ; Rx#
                If (p_SelectedRxNumber != "")
                    DownAndPaste(p_SelectedRxNumber)
                Else {
                    Send("{Down}")
                    Sleep(100)
                }
                Try p_SelectedRxNumber := ScrubString(p_SelectedRxNumber)
                p_SelectedRxNumber := ""
                Try p_MedName := ScrubString(p_MedName)
                p_MedName := ""

                ; Change Needed = RTF reason
                FindAndPaste("Change Needed", "RTF - " . p_RTFReason)
                Try p_RTFReason := ScrubString(p_RTFReason)
                p_RTFReason := ""

                ; Shipment link and/or Order ID
                FindAndPaste("Shipment link and/or Order ID", p_ShipmentURL)
                Try p_ShipmentURL := ScrubString(p_ShipmentURL)
                p_ShipmentURL := ""
                Sleep(200)

                ; Final notification
                PauseStart()
                MsgBox("Clinical Review ticket populated.`n`n*** VERIFY ALL INFORMATION IS CORRECT BEFORE SUBMITTING ***`n`nEnsure you DENY the PDMP check for this patient.", "Refusal to Fill - Complete", 48+262144)
                PauseEnd()

                WipeAllData()
                Return
            }

            If (g_RTFDecision == "callback") {
                ; ==========================================================
                ; PRESCRIBER/PHARMACY CALLBACK PATH (Amazon)
                ; Blueprint: Prescriber\Pharmacy Call Back
                ; (Communications > Call Back Request > Resource RPh)
                ; ==========================================================

                ; Scrub data not needed for this path
                Try p_LastFillDate := ScrubString(p_LastFillDate)
                p_LastFillDate := ""
                Try p_LastFillDays := ScrubString(p_LastFillDays)
                p_LastFillDays := ""
                Try p_LastPharmacy := ScrubString(p_LastPharmacy)
                p_LastPharmacy := ""
                Try p_NextFillDate := ScrubString(p_NextFillDate)
                p_NextFillDate := ""
                Try p_CustomerPhone := ScrubString(p_CustomerPhone)
                p_CustomerPhone := ""
                Try p_SelectedNDC := ScrubString(p_SelectedNDC)
                p_SelectedNDC := ""
                Try p_RTFReasonDetail := ScrubString(p_RTFReasonDetail)
                p_RTFReasonDetail := ""
                p_Batch := ""

                ; --- Prescriber lookup: find Rx on admin page, open in new tab ---
                PrescriberName := ""
                PrescriberPhone := ""
                Send("^f")
                Sleep(75)
                SendText(p_SelectedRxNumber)
                Sleep(50)
                Send("{Enter}")
                Sleep(50)
                Send("{Esc}")
                Sleep(200)
                Send("{Enter}")
                Sleep(2500)

                ; Scrape the newly opened Rx detail tab
                A_Clipboard := ""
                Send("^a")
                Sleep(50)
                Send("^c")
                if ClipWait(2) {
                    MarkClipboardExcluded()
                    RxPageText := A_Clipboard
                    if !(InStr(RxPageText, "Data Entry") && InStr(RxPageText, "Edit insurances") && InStr(RxPageText, "DocuPack Document")) {
                        RxPageText := ScrubString(RxPageText)
                        RxPageText := ""
                        A_Clipboard := ""
                        Send("^w")
                        Sleep(500)
                        PauseStart()
                        Result := MsgBox("The Rx detail page did not open correctly.`n`nExpected to see 'Data Entry', 'Edit insurances', and 'DocuPack Document' on the page.`n`nContinue without auto-filling prescriber info?`n(You will need to enter it manually on the ticket.)", "Rx Page Not Found", 49+262144)
                        PauseEnd()
                        if (Result = "Cancel") {
                            WipeAllData()
                            Return
                        }
                        ReactivateChrome()
                    } else {
                        if RegExMatch(RxPageText, "i)Prescriber\s+([^\r\n]+)", &PrescMatch)
                            PrescriberName := Trim(PrescMatch[1])
                        if RegExMatch(RxPageText, "i)Phone\s*Number[:\s]+([^\r\n]+)", &PhoneMatch)
                            PrescriberPhone := Trim(PhoneMatch[1])
                        RxPageText := ScrubString(RxPageText)
                        RxPageText := ""
                        A_Clipboard := ""
                        Send("^w")
                        Sleep(500)
                    }
                } else {
                    A_Clipboard := ""
                    Send("^w")
                    Sleep(500)
                }

                ; Verify prescriber info was found
                if (PrescriberName = "" || PrescriberPhone = "") {
                    MissingFields := ""
                    if (PrescriberName = "")
                        MissingFields .= "- Prescriber Name`n"
                    if (PrescriberPhone = "")
                        MissingFields .= "- Phone Number`n"
                    PauseStart()
                    Result := MsgBox("Could not scrape the following from the Rx detail page:`n`n" . MissingFields . "`nContinue without auto-filling prescriber info?`n(You will need to enter it manually on the ticket.)", "Prescriber Lookup Failed", 49+262144)
                    PauseEnd()
                    if (Result = "Cancel") {
                        WipeAllData()
                        Return
                    }
                    ReactivateChrome()
                }

                ; Open Prescriber/Pharmacy Call Back blueprint directly
                Send("^t")
                Sleep(200)
                A_Clipboard := "https://cs.wolfgang.a2z.com/tachyon/PHARMACY/createNewIssue/360121418232"
                Sleep(50)
                Send("^v")
                Sleep(100)
                Send("{Enter}")
                A_Clipboard := ""

                ; Wait for the ticket form to load (check for "Task: Prescriber/Pharmacy Callback")
                Loaded := false
                Loop 30 {
                    Sleep(A_Index = 1 ? 1500 : 500)
                    If CheckChromeTitleForLogin("Tachyon")
                        Continue
                    A_Clipboard := ""
                    Send("^a")
                    Sleep(50)
                    Send("^c")
                    If ClipWait(0.5) {
                        PageText := A_Clipboard
                        A_Clipboard := ""
                        If InStr(PageText, "Task: Prescriber/Pharmacy Callback") {
                            PageText := ""
                            ClearSelection()
                            Loaded := true
                            break
                        }
                        If InStr(PageText, "Sign In") || InStr(PageText, "Username") || InStr(PageText, "Verify with your password") || InStr(PageText, "Powered by Okta") {
                            PageText := ""
                            PauseStart()
                            MsgBox("Sign-in required!`n`n1. Please log in manually.`n2. Click OK on this popup ONLY AFTER you see the ticket form.", "Login Paused", 64+262144)
                            PauseEnd()
                            Sleep(300)
                            WinActivate("ahk_exe chrome.exe")
                            Sleep(300)
                        }
                        PageText := ""
                        ClearSelection()
                    } Else {
                        ClearSelection()
                    }
                }
                If (!Loaded) {
                    MsgBox("STOPPING: Timed out waiting for the Prescriber/Pharmacy Callback ticket form to load.", "Timeout", 16)
                    WipeAllData()
                    Return
                }

                ; Type Person ID and verify
                LocalID := p_Identifier
                Try p_Identifier := ScrubString(p_Identifier)
                p_Identifier := ""

                FindAndPasteField("Person ID", LocalID)
                If !VerifyFieldContent(LocalID) {
                    Try LocalID := ScrubString(LocalID)
                    LocalID := ""
                    MsgBox("Customer ID paste verification FAILED.`n`nAborting and wiping all data.", "Field Verify Failed", 16)
                    WipeAllData()
                    Return
                }
                Try LocalID := ScrubString(LocalID)
                LocalID := ""

                ; Select customer name
                If !SelectCustomerName(true) {
                    MsgBox("WARNING: Could not confirm the customer name was selected.`n`nVerify the 'Customer first name' field is populated before submitting.", "Customer Selection Warning", 48)
                    ReactivateChrome()
                }

                Sleep(300)

                ; Priority -> High if due today or earlier
                SetPriorityIfDueToday(p_ShipDate)

                ; Due By = Ship Date
                FindAndType("Due", p_ShipDate)
                Try p_ShipDate := ScrubString(p_ShipDate)
                p_ShipDate := ""
                p_ArrivalParsed := ""
                p_ArrivalDate := ""

                ; PPO - Source of ticket creation = "Fulfillment" (index 5)
                FindAndSelectDropdown("PPO - Source", 5)

                ; Comment field — Find pre-existing labels, append data
                FindAndPaste("Name of Prescriber/Pharmacy:", PrescriberName)
                FindAndPaste("Prescriber/Pharmacy phone:", PrescriberPhone)
                FindAndPaste("Customer link of who the call referred to:", p_ShipmentURL)
                FindAndPaste("Additional information:", "")
                DownAndPaste(p_ShipmentURL)
                DownAndPaste(p_MedName . " | Rx#: " . p_SelectedRxNumber)
                DownAndPaste("RTF - " . p_RTFReason)
                Try PrescriberName := ScrubString(PrescriberName)
                PrescriberName := ""
                Try PrescriberPhone := ScrubString(PrescriberPhone)
                PrescriberPhone := ""

                Try p_ShipmentURL := ScrubString(p_ShipmentURL)
                p_ShipmentURL := ""
                Try p_MedName := ScrubString(p_MedName)
                p_MedName := ""
                Try p_RTFReason := ScrubString(p_RTFReason)
                p_RTFReason := ""
                Try p_SelectedRxNumber := ScrubString(p_SelectedRxNumber)
                p_SelectedRxNumber := ""

                Sleep(200)

                ; Final notification
                PauseStart()
                MsgBox("Prescriber/Pharmacy Callback ticket populated.`n`n*** VERIFY ALL INFORMATION IS CORRECT BEFORE SUBMITTING ***`n`nEnsure you DENY the PDMP check for this patient.", "Refusal to Fill - Callback Complete", 48+262144)
                PauseEnd()

                WipeAllData()
                Return
            }

            ; ==========================================================
            ; CANCELLATION PATH: Navigate to Orders, select order, stop
            ; with instructions (no customer communication generated)
            ; ==========================================================
            ; Scrub ticket-only fields not needed for cancellation
            Try p_SelectedRxNumber := ScrubString(p_SelectedRxNumber)
            p_SelectedRxNumber := ""
            Try p_SelectedNDC := ScrubString(p_SelectedNDC)
            p_SelectedNDC := ""
            Try p_RTFReasonDetail := ScrubString(p_RTFReasonDetail)
            p_RTFReasonDetail := ""
            Try p_PillPackProfileID := ScrubString(p_PillPackProfileID)
            p_PillPackProfileID := ""
            ; Falls through to the Wolfgang dashboard flow below,
            ; but will skip the template generation at the end.
        }

        ; =============================================================
        ; AMAZON PATH (Early Control or Cancellation path)
        ; =============================================================
        ; Clear non-Amazon-specific data that won't be used
        Try p_ShipmentURL := ScrubString(p_ShipmentURL)
        p_ShipmentURL := ""
        Try p_ShipDate := ScrubString(p_ShipDate)
        p_ShipDate := ""
        p_ArrivalDate   := ""
        p_ArrivalParsed := ""

        Send("^t")
        Sleep(200)
        A_Clipboard := "https://cs.wolfgang.a2z.com/dashboard"
        Send("^v")
        Sleep(100)
        Send("{Enter}")
        A_Clipboard := ""

        LoadedDashboard := false

        Loop 45 {
            Sleep(1500)

            A_Clipboard := ""
            Send("^a")
            Sleep(150)
            Send("^c")

            if ClipWait(1) {
                PageText := A_Clipboard
                A_Clipboard := ""

                ClearSelection()

                if (RegExMatch(PageText, "^https?://[^\r\n]+$")) {
                    PageText := ""
                    ClearSelectionJS()
                    Continue
                }

                if InStr(PageText, "Active contacts") || InStr(PageText, "NON-CALL") {
                    PageText := ""
                    LoadedDashboard := true
                    break
                } else if InStr(PageText, "Sign in with your account") || InStr(PageText, "Username") {
                    PageText := ""
                    PauseStart()
                    MsgBox("Sign-in required!`n`n1. Please log in manually.`n2. Click OK on this popup ONLY AFTER you see the Dashboard.", "Login Paused", "Iconi")
                    PauseEnd()
                    Sleep(500)
                    WinActivate("ahk_exe chrome.exe")
                    Sleep(500)
                } else {
                    PageText := ""
                }
            } else {
                ClearSelection()
            }
        }

        if (!LoadedDashboard) {
            MsgBox("STOPPING: Timed out waiting for the Wolfgang Dashboard.", "Timeout", 16)
            WipeAllData()
            return
        }

        MaxRetries := 3
        RetryCount := 0
        LoadedPatientDash := false

        LocalID := p_Identifier
        Try p_Identifier := ScrubString(p_Identifier)
        p_Identifier := ""

        while (!LoadedPatientDash && RetryCount < MaxRetries) {

            Sleep(300)
            Send("^f")
            Sleep(50)
            SendText("START NON-CALL")
            Sleep(50)
            Send("{Esc}")
            Sleep(50)
            Send("{Enter}")

            Sleep(1500)

            Send("^f")
            Sleep(50)
            SendText("Enter Amazon PID")
            Sleep(50)
            Send("{Esc}")
            Sleep(50)
            Send("{Tab}")
            Sleep(100)

            SendText(LocalID)
            Sleep(150)

            A_Clipboard := ""
            Send("^a")
            Sleep(50)
            Send("^c")
            PasteOK := false
            If ClipWait(0.5) {
                ReadBack := Trim(A_Clipboard)
                A_Clipboard := ""
                If (ReadBack != "" && ReadBack == LocalID)
                    PasteOK := true
                Try ReadBack := ScrubString(ReadBack)
                ReadBack := ""
            }
            A_Clipboard := ""

            If (!PasteOK) {
                Try LocalID := ScrubString(LocalID)
                LocalID := ""
                MsgBox("Identifier paste verification FAILED.`n`nAborting and wiping all data.", "Field Verify Failed", 16)
                WipeAllData()
                Return
            }

            Send("{Enter}")

            Loop 15 {
                Sleep(1300)
                A_Clipboard := ""
                Send("^a")
                Sleep(150)
                Send("^c")

                if ClipWait(1) {
                    PageText := A_Clipboard
                    A_Clipboard := ""

                    ClearSelection()

                    if (RegExMatch(PageText, "^https?://[^\r\n]+$")) {
                        PageText := ""
                        ClearSelectionJS()
                        Continue
                    }

                    if InStr(PageText, "Health Profile") && InStr(PageText, "Orders") {
                        PageText := ""
                        LoadedPatientDash := true
                        break
                    }

                    if InStr(PageText, "You are logged out") || InStr(PageText, "TAKE ME TO THE DASHBOARD") {
                        PageText := ""
                        RetryCount++

                        PauseStart()
                        MsgBox("Session expired!`n`n1. Click 'TAKE ME TO THE DASHBOARD'`n2. Log back in manually`n3. Click OK ONLY AFTER you see the Dashboard.`n`nRetry attempt " RetryCount " of " MaxRetries, "Re-login Required", "Iconi")
                        PauseEnd()
                        Sleep(300)
                        WinActivate("ahk_exe chrome.exe")
                        Sleep(300)

                        DashboardReady := false
                        Loop 30 {
                            Sleep(1500)
                            A_Clipboard := ""
                            Send("^a")
                            Sleep(150)
                            Send("^c")

                            if ClipWait(1) {
                                DashText := A_Clipboard
                                A_Clipboard := ""

                                ClearSelection()

                                if (RegExMatch(DashText, "^https?://[^\r\n]+$")) {
                                    DashText := ""
                                    ClearSelectionJS()
                                    Continue
                                }

                                if InStr(DashText, "Active contacts") || InStr(DashText, "NON-CALL") {
                                    DashText := ""
                                    DashboardReady := true
                                    break
                                }
                                DashText := ""
                            } else {
                                ClearSelection()
                            }
                        }

                        if (!DashboardReady) {
                            Try LocalID := ScrubString(LocalID)
                            LocalID := ""
                            MsgBox("STOPPING: Timed out waiting for Dashboard after re-login.", "Timeout", 16)
                            WipeAllData()
                            return
                        }

                        break
                    }
                    PageText := ""
                } else {
                    ClearSelection()
                }
            }
        }

        Try LocalID := ScrubString(LocalID)
        LocalID := ""

        if (!LoadedPatientDash) {
            MsgBox("STOPPING: Could not load patient profile after " MaxRetries " attempts.", "Stopping", 16)
            WipeAllData()
            return
        }

        ; Click 'Orders' Tab (via ExecJS)
        ClearSelection()
        Sleep(500)

        ExecJS("[...document.querySelectorAll('*')].find(e=>e.textContent.trim()==='Orders').click();void(0);")

        ; Wait for Orders Page to Load
        Sleep(300)
        ClearSelectionJS()

        LoadedOrders := false
        PageText := ""

        Loop 15 {
            Sleep(500)
            A_Clipboard := ""
            Send("^a")
            Sleep(150)
            Send("^c")

            if ClipWait(1) {
                PageText := A_Clipboard
                A_Clipboard := ""

                ClearSelection()

                if (RegExMatch(PageText, "^https?://[^\r\n]+$")) {
                    PageText := ""
                    ClearSelectionJS()
                    Continue
                }

                if InStr(PageText, "Order Search") {
                    PageText := ""
                    LoadedOrders := true
                    Sleep(500)
                    break
                }
                PageText := ""
            } else {
                ClearSelection()
            }
        }

        if (!LoadedOrders) {
            MsgBox("STOPPING: Timed out waiting for the Orders page to load.", "Timeout", 16)
            WipeAllData()
            return
        }

        ; Click medication row in Orders table via JS with title signal (retry up to 3 times)
        ; Strip trailing parenthetical only; uppercase for case-insensitive match
        MedForJS := StrUpper(RegExReplace(p_MedName, "\s*\([^)]*\)\s*$", ""))
        MedForJS := Trim(MedForJS)
        MedForJS := StrReplace(MedForJS, "\", "\\")
        MedForJS := StrReplace(MedForJS, "'", "\'")
        MedForJS := StrReplace(MedForJS, "`n", "")
        MedForJS := StrReplace(MedForJS, "`r", "")
        MedForJS := StrReplace(MedForJS, Chr(0), "")

        JSPayload := "(function(){window.getSelection().removeAllRanges();var rows=document.querySelectorAll('tr[role=`"row`"]');for(var i=0;i<rows.length;i++){if(rows[i].textContent.toUpperCase().includes('" . MedForJS . "')){rows[i].click();document.title='__WF_ROW_OK_'+Date.now();return;}}document.title='__WF_ROW_NOTFOUND_'+Date.now();})();void(0);"

        Try MedForJS := ScrubString(MedForJS)
        MedForJS := ""

        RowClickOK := false
        Loop 3 {
            ClearSelection()
            Sleep(150)
            ClearSelectionJS()
            Sleep(300)

            ExecJS(JSPayload)

            Loop 20 {
                Sleep(500)
                Try {
                    ChromeTitle := WinGetTitle("ahk_exe chrome.exe")
                    If InStr(ChromeTitle, "__WF_ROW_OK_") {
                        RowClickOK := true
                        ChromeTitle := ""
                        Break
                    }
                    If InStr(ChromeTitle, "__WF_ROW_NOTFOUND_") {
                        ChromeTitle := ""
                        Break
                    }
                    ChromeTitle := ""
                }
            }

            If (RowClickOK)
                Break
            Sleep(1000)
        }
        JSPayload := ""

        If (!RowClickOK) {
            PauseStart()
            ManualResult := MsgBox("Could not click the medication row in the Orders table after 3 attempts.`n`nThe medication may not be visible on the page.`n`nWould you still like to generate the customer communication?", "Row Click Failed", 4+48)
            PauseEnd()
            If (ManualResult = "No") {
                WipeAllData()
                Return
            }
            ReactivateChrome()

            ; Prompt user for Order ID manually
            ManualOrderGui := Gui("+AlwaysOnTop", "Enter Order ID")
            ManualOrderGui.SetFont("s10")
            ManualOrderGui.Add("Text", "w350", "Enter the Order ID:")
            ManualOrderGui.Add("Edit", "vManualOrderID w300", "")
            ManualOrderGui.Add("Text", "w350", "")
            BtnGenerate := ManualOrderGui.Add("Button", "Default w180", "Generate Communication")
            BtnCancelOrder := ManualOrderGui.Add("Button", "x+20 w100", "Cancel")

            g_FormSubmitted := false

            BtnGenerate.OnEvent("Click", OnManualGenerate)
            BtnCancelOrder.OnEvent("Click", OnManualCancel)
            ManualOrderGui.OnEvent("Close", GUICrash)

            OnManualGenerate(btn, info) {
                global p_OrderID, g_FormSubmitted
                Saved := btn.Gui.Submit()
                If (Trim(Saved.ManualOrderID) = "") {
                    MsgBox("Please enter an Order ID.", "Validation Error", 48)
                    btn.Gui.Show()
                    return
                }
                p_OrderID := Trim(Saved.ManualOrderID)
                g_FormSubmitted := true
                btn.Gui.Destroy()
            }

            OnManualCancel(btn, info := "") {
                global g_FormSubmitted
                g_FormSubmitted := false
                If (btn is Gui)
                    btn.Destroy()
                Else
                    btn.Gui.Destroy()
            }

            PauseStart()
            ManualOrderGui.Show()
            While WinExist("Enter Order ID")
                Sleep(100)
            PauseEnd()

            If (!g_FormSubmitted) {
                WipeAllData()
                Return
            }
            ReactivateChrome()

        } Else {

            ; Wait for Order Details & Extract Order ID
            Sleep(1000)

            LoadedOrderDetails := false
            p_OrderID := ""

            Loop 15 {
                Sleep(500)
                A_Clipboard := ""
                Send("^a")
                Sleep(150)
                Send("^c")

                if ClipWait(1) {
                    PageText := A_Clipboard
                    A_Clipboard := ""

                    ClearSelection()

                    if (RegExMatch(PageText, "^https?://[^\r\n]+$")) {
                        PageText := ""
                        ClearSelectionJS()
                        Continue
                    }

                    if InStr(PageText, "order in fulfillment") || InStr(PageText, "Order delivered") {
                        if RegExMatch(PageText, "i)Bulk\s+(\S+)", &OrderMatch) {
                            p_OrderID := Trim(OrderMatch[1])
                            OrderPageText := PageText
                            PageText  := ""
                            LoadedOrderDetails := true
                            break
                        }
                    }
                    PageText := ""
                } else {
                    ClearSelection()
                }
            }

            if (!LoadedOrderDetails || p_OrderID = "") {
                MsgBox("STOPPING: Could not extract the Order ID from the order details page.", "Extraction Failed", 16)
                WipeAllData()
                return
            }

            ; Canary: Verify medication name appears between "Order Placed" and "All Claims"
            OrderPlacedPos := InStr(OrderPageText, "Order Placed")
            AllClaimsPos   := InStr(OrderPageText, "All Claims")
            If (OrderPlacedPos > 0 && AllClaimsPos > OrderPlacedPos) {
                BetweenSection := SubStr(OrderPageText, OrderPlacedPos, AllClaimsPos - OrderPlacedPos)
                If !InStr(BetweenSection, p_MedName) {
                    BetweenSection := ""
                    OrderPageText := ""
                    MsgBox("STOPPING: Medication name mismatch.`n`nExpected to find:`n" . p_MedName . "`n`nbetween 'Order Placed' and 'All Claims' on the order details page, but it was not found.`nThe wrong medication may have been selected.", "Medication Canary Failed", 16)
                    WipeAllData()
                    return
                }
                BetweenSection := ""
            }
            OrderPageText := ""

            ClearSelectionJS()
        }

        ; Branch: Early Control -> template | Non-Early Control cancellation -> no template
        If (g_RTFIsEarlyControl) {
            ; Generate Template & Display to User
            Template := "After reviewing order # " . p_OrderID . " for " . p_MedName . ", it turns out we had to cancel your order. This happened because this medication is a controlled substance and it was filled at " . p_LastPharmacy . " on " . p_LastFillDate . " for " . p_LastFillDays . " days supply, so you're not due for a refill until " . p_NextFillDate . ". We can't refill this medication early without getting approval from your prescriber. If you need this medication processed before " . p_NextFillDate . ", contact Amazon Pharmacy directly at 855-745-5725. We can then reach out to your prescriber for approval. If you're OK with waiting, place another order for " . p_MedName . " on " . p_NextFillDate . " and the medication will be processed on that day. The delivery timeline will be presented when the medication completes processing and is available for checkout."

            A_Clipboard := Template
            MarkClipboardExcluded()

            ResultGui := Gui(, "Template Ready - Copied to Clipboard")
            ResultGui.SetFont("s10")

            ResultGui.Add("Text", "w600", "The following message has been copied to your clipboard:")
            ResultGui.Add("Text", "w600", "")
            ResultGui.Add("Edit", "w600 h200 ReadOnly", Template)
            ResultGui.Add("Text", "w600", "")
            ResultGui.Add("Text", "w600 cGreen", "Template copied to clipboard! Press Ctrl+V to paste.")

            Template := ScrubString(Template)
            Template := ""

            BtnClose := ResultGui.Add("Button", "Default w100", "Close")
            BtnClose.OnEvent("Click", CloseAndWipe)
            ResultGui.OnEvent("Close", GUICrash)

            CloseAndWipe(*) {
                ResultGui.Destroy()
                WipeAllData()
            }

            PauseStart()
            ResultGui.Show()

        } Else {
            ; Non-Early Control cancellation: scrub Early Control-only fields
            Try p_LastFillDate := ScrubString(p_LastFillDate)
            p_LastFillDate := ""
            Try p_LastFillDays := ScrubString(p_LastFillDays)
            p_LastFillDays := ""
            Try p_LastPharmacy := ScrubString(p_LastPharmacy)
            p_LastPharmacy := ""
            Try p_NextFillDate := ScrubString(p_NextFillDate)
            p_NextFillDate := ""
            Try p_CustomerPhone := ScrubString(p_CustomerPhone)
            p_CustomerPhone := ""

            ; order is selected, inform user to generate message manually
            PauseStart()
            MsgBox("Order has been selected for cancellation.`n`nRTF Reason: " . p_RTFReason . "`n`nNo pre-built customer communication is available for this reason.`nPlease generate an appropriate message to the customer based on the specific issue.`n`nCancel the order and communicate with the customer accordingly.", "Refusal to Fill - Manual Communication Required", 48+262144)
            PauseEnd()
            WipeAllData()
        }

    } else {

        ; =============================================================
        ; NON-AMAZON PATH: Decision — RTF Ticket or Callback
        ; =============================================================
        g_RTFDecision := ""

        NonAmazonGui := Gui("+AlwaysOnTop", "RTF - Non-Amazon Decision")
        NonAmazonGui.SetFont("s10")
        NonAmazonGui.Add("Text", "w500", "Non-Amazon customer — select ticket type:")
        NonAmazonGui.Add("Text", "w500", "")

        BtnRTF := NonAmazonGui.Add("Button", "w200", "Refusal to Fill")
        BtnCB := NonAmazonGui.Add("Button", "x+10 w220", "Prescriber/Pharmacy Callback")
        BtnAbortNA := NonAmazonGui.Add("Button", "xm w100", "Abort")

        BtnRTF.OnEvent("Click", OnNARTF)
        BtnCB.OnEvent("Click", OnNACallback)
        BtnAbortNA.OnEvent("Click", OnNAAbort)
        NonAmazonGui.OnEvent("Close", GUICrash)

        OnNARTF(btn, info) {
            global g_RTFDecision
            g_RTFDecision := "rtf"
            btn.Gui.Destroy()
        }

        OnNACallback(btn, info) {
            global g_RTFDecision
            g_RTFDecision := "callback"
            btn.Gui.Destroy()
        }

        OnNAAbort(btn, info := "") {
            global g_RTFDecision
            g_RTFDecision := "abort"
            If (btn is Gui)
                btn.Destroy()
            Else
                btn.Gui.Destroy()
        }

        PauseStart()
        NonAmazonGui.Show()
        While WinExist("RTF - Non-Amazon Decision")
            Sleep(100)
        PauseEnd()

        If (g_RTFDecision == "abort") {
            WipeAllData()
            Return
        }

        ReactivateChrome()

        If (g_RTFDecision == "callback") {
            ; ==========================================================
            ; PRESCRIBER/PHARMACY CALLBACK PATH (Non-Amazon)
            ; Blueprint: Prescriber\Pharmacy Call Back
            ; (Communications > Call Back Request > Resource RPh)
            ; ==========================================================

            ; Scrub data not needed for this path
            Try p_LastFillDate := ScrubString(p_LastFillDate)
            p_LastFillDate := ""
            Try p_LastFillDays := ScrubString(p_LastFillDays)
            p_LastFillDays := ""
            Try p_LastPharmacy := ScrubString(p_LastPharmacy)
            p_LastPharmacy := ""
            Try p_NextFillDate := ScrubString(p_NextFillDate)
            p_NextFillDate := ""
            Try p_CustomerPhone := ScrubString(p_CustomerPhone)
            p_CustomerPhone := ""
            Try p_SelectedNDC := ScrubString(p_SelectedNDC)
            p_SelectedNDC := ""
            Try p_RTFReasonDetail := ScrubString(p_RTFReasonDetail)
            p_RTFReasonDetail := ""

            ; --- Prescriber lookup: find Rx on admin page, open in new tab ---
            PrescriberName := ""
            PrescriberPhone := ""
            Send("^f")
            Sleep(150)
            SendText(p_SelectedRxNumber)
            Sleep(150)
            Send("{Enter}")
            Sleep(150)
            Send("{Esc}")
            Sleep(300)
            Send("{Enter}")
            Sleep(2500)

            ; Scrape the newly opened Rx detail tab
            A_Clipboard := ""
            Send("^a")
            Sleep(50)
            Send("^c")
            if ClipWait(2) {
                MarkClipboardExcluded()
                RxPageText := A_Clipboard
                if !(InStr(RxPageText, "Data Entry") && InStr(RxPageText, "Edit insurances") && InStr(RxPageText, "DocuPack Document")) {
                    RxPageText := ScrubString(RxPageText)
                    RxPageText := ""
                    A_Clipboard := ""
                    Send("^w")
                    Sleep(500)
                    PauseStart()
                    Result := MsgBox("The Rx detail page did not open correctly.`n`nExpected to see 'Data Entry', 'Edit insurances', and 'DocuPack Document' on the page.`n`nContinue without auto-filling prescriber info?`n(You will need to enter it manually on the ticket.)", "Rx Page Not Found", 49+262144)
                    PauseEnd()
                    if (Result = "Cancel") {
                        WipeAllData()
                        Return
                    }
                    ReactivateChrome()
                } else {
                    if RegExMatch(RxPageText, "i)Prescriber\s+([^\r\n]+)", &PrescMatch)
                        PrescriberName := Trim(PrescMatch[1])
                    if RegExMatch(RxPageText, "i)Phone\s*Number[:\s]+([^\r\n]+)", &PhoneMatch)
                        PrescriberPhone := Trim(PhoneMatch[1])
                    RxPageText := ScrubString(RxPageText)
                    RxPageText := ""
                    A_Clipboard := ""
                    Send("^w")
                    Sleep(500)
                }
            } else {
                A_Clipboard := ""
                Send("^w")
                Sleep(500)
            }

            ; Verify prescriber info was found
            if (PrescriberName = "" || PrescriberPhone = "") {
                MissingFields := ""
                if (PrescriberName = "")
                    MissingFields .= "- Prescriber Name`n"
                if (PrescriberPhone = "")
                    MissingFields .= "- Phone Number`n"
                PauseStart()
                Result := MsgBox("Could not scrape the following from the Rx detail page:`n`n" . MissingFields . "`nContinue without auto-filling prescriber info?`n(You will need to enter it manually on the ticket.)", "Prescriber Lookup Failed", 49+262144)
                PauseEnd()
                if (Result = "Cancel") {
                    WipeAllData()
                    Return
                }
                ReactivateChrome()
            }

            ; Open Prescriber/Pharmacy Call Back blueprint directly (non-Amazon)
            Send("^t")
            Sleep(200)
            A_Clipboard := "https://cs.wolfgang.a2z.com/tachyon/PILLPACK/createNewIssue/360098758174"
            Sleep(50)
            Send("^v")
            Sleep(100)
            Send("{Enter}")
            A_Clipboard := ""

            ; Wait for the ticket form to load (check for "Task: Prescriber/Pharmacy Callback")
            Loaded := false
            Loop 30 {
                Sleep(A_Index = 1 ? 1500 : 500)
                If CheckChromeTitleForLogin("Tachyon")
                    Continue
                A_Clipboard := ""
                Send("^a")
                Sleep(50)
                Send("^c")
                If ClipWait(0.5) {
                    PageText := A_Clipboard
                    A_Clipboard := ""
                    If InStr(PageText, "Task: Prescriber/Pharmacy Callback") {
                        PageText := ""
                        ClearSelection()
                        Loaded := true
                        break
                    }
                    If InStr(PageText, "Sign In") || InStr(PageText, "Username") || InStr(PageText, "Verify with your password") || InStr(PageText, "Powered by Okta") {
                        PageText := ""
                        PauseStart()
                        MsgBox("Sign-in required!`n`n1. Please log in manually.`n2. Click OK on this popup ONLY AFTER you see the ticket form.", "Login Paused", 64+262144)
                        PauseEnd()
                        Sleep(300)
                        WinActivate("ahk_exe chrome.exe")
                        Sleep(300)
                    }
                    PageText := ""
                    ClearSelection()
                } Else {
                    ClearSelection()
                }
            }
            If (!Loaded) {
                MsgBox("STOPPING: Timed out waiting for the Prescriber/Pharmacy Callback ticket form to load.", "Timeout", 16)
                WipeAllData()
                Return
            }

            ; Type PillPack Profile ID and verify
            LocalProfileID := p_PillPackProfileID
            Try p_PillPackProfileID := ScrubString(p_PillPackProfileID)
            p_PillPackProfileID := ""

            FindAndPasteField("PillPack Profile ID", LocalProfileID)
            If !VerifyFieldContent(LocalProfileID) {
                Try LocalProfileID := ScrubString(LocalProfileID)
                LocalProfileID := ""
                MsgBox("Customer ID paste verification FAILED.`n`nAborting and wiping all data.", "Field Verify Failed", 16)
                WipeAllData()
                Return
            }
            Try LocalProfileID := ScrubString(LocalProfileID)
            LocalProfileID := ""

            ; Select customer name
            If !SelectCustomerName(false) {
                MsgBox("WARNING: Could not confirm the customer name was selected.`n`nVerify the 'Customer first name' field is populated before submitting.", "Customer Selection Warning", 48)
                ReactivateChrome()
            }

            Sleep(300)

            ; Priority -> High if due today or earlier (use arrival date for non-Amazon)
            DueByVal := ""
            If (Type(p_ArrivalParsed) == "Object" && p_ArrivalParsed.formatted != "")
                DueByVal := p_ArrivalParsed.formatted
            If (DueByVal == "")
                DueByVal := p_ShipDate
            SetPriorityIfDueToday(DueByVal)
            FindAndType("Due", DueByVal)
            DueByVal := ""
            Try p_ShipDate := ScrubString(p_ShipDate)
            p_ShipDate := ""
            p_ArrivalParsed := ""
            p_ArrivalDate := ""

            ; PPO - Source of ticket creation = "Fulfillment" (index 5)
            FindAndSelectDropdown("PPO - Source", 5)

            ; Comment field — non-Amazon has all labels on one line
            ; Find ":Additional information:", End, then Down to new lines
            Send("^f")
            Sleep(50)
            SendText(":Additional information:")
            Sleep(50)
            Send("{Esc}")
            Sleep(50)
            Send("{End}")
            Sleep(100)
            Send("{Enter}")
            Sleep(50)
            A_Clipboard := PrescriberName
            MarkClipboardExcluded()
            Send("^v")
            Sleep(50)
            A_Clipboard := ""
            Sleep(100)
            Send("{Enter}")
            Sleep(50)
            A_Clipboard := PrescriberPhone
            MarkClipboardExcluded()
            Send("^v")
            Sleep(50)
            A_Clipboard := ""
            Sleep(100)
            Send("{Enter}")
            Sleep(50)
            A_Clipboard := p_ShipmentURL
            MarkClipboardExcluded()
            Send("^v")
            Sleep(50)
            A_Clipboard := ""
            Sleep(100)
            Send("{Enter}")
            Sleep(50)
            A_Clipboard := p_MedName . " | Rx#: " . p_SelectedRxNumber
            MarkClipboardExcluded()
            Send("^v")
            Sleep(50)
            A_Clipboard := ""
            Sleep(100)
            Send("{Enter}")
            Sleep(50)
            A_Clipboard := "RTF - " . p_RTFReason
            MarkClipboardExcluded()
            Send("^v")
            Sleep(50)
            A_Clipboard := ""
            Try PrescriberName := ScrubString(PrescriberName)
            PrescriberName := ""
            Try PrescriberPhone := ScrubString(PrescriberPhone)
            PrescriberPhone := ""

            Try p_ShipmentURL := ScrubString(p_ShipmentURL)
            p_ShipmentURL := ""
            Try p_MedName := ScrubString(p_MedName)
            p_MedName := ""
            Try p_RTFReason := ScrubString(p_RTFReason)
            p_RTFReason := ""
            Try p_SelectedRxNumber := ScrubString(p_SelectedRxNumber)
            p_SelectedRxNumber := ""
            Try p_Identifier := ScrubString(p_Identifier)
            p_Identifier := ""

            Sleep(200)

            ; Final notification
            PauseStart()
            MsgBox("Prescriber/Pharmacy Callback ticket populated.`n`n*** VERIFY ALL INFORMATION IS CORRECT BEFORE SUBMITTING ***`n`nEnsure you DENY the PDMP check for this patient.", "Refusal to Fill - Callback Complete", 48+262144)
            PauseEnd()

            WipeAllData()
            Return
        }

        ; =============================================================
        ; NON-AMAZON PATH: Tachyon "Refusal to Fill" Ticket
        ; =============================================================

        ; --- Open Refusal to Fill blueprint directly (Non-Amazon) ---
        Send("^t")
        Sleep(200)
        A_Clipboard := "https://cs.wolfgang.a2z.com/tachyon/PILLPACK/createNewIssue/360179179473"
        Sleep(50)
        Send("^v")
        Sleep(100)
        Send("{Enter}")
        A_Clipboard := ""

        ; Wait for the ticket form to load
        Loaded := false
        Loop 30 {
            Sleep(A_Index = 1 ? 1500 : 500)
            If CheckChromeTitleForLogin("Tachyon")
                Continue
            A_Clipboard := ""
            Send("^a")
            Sleep(50)
            Send("^c")
            If ClipWait(0.5) {
                PageText := A_Clipboard
                A_Clipboard := ""
                If InStr(PageText, "Task: Refusal to Fill") {
                    PageText := ""
                    ClearSelection()
                    Loaded := true
                    break
                }
                If InStr(PageText, "Sign In") || InStr(PageText, "Username") || InStr(PageText, "Verify with your password") || InStr(PageText, "Powered by Okta") {
                    PageText := ""
                    PauseStart()
                    MsgBox("Sign-in required!`n`n1. Please log in manually.`n2. Click OK on this popup ONLY AFTER you see the ticket form.", "Login Paused", 64+262144)
                    PauseEnd()
                    Sleep(300)
                    WinActivate("ahk_exe chrome.exe")
                    Sleep(300)
                }
                PageText := ""
                ClearSelection()
            } Else {
                ClearSelection()
            }
        }
        If (!Loaded) {
            MsgBox("STOPPING: Timed out waiting for the Refusal to Fill ticket form to load.", "Timeout", 16)
            WipeAllData()
            Return
        }

        ; Navigate to Profile ID field and type it
        LocalProfileID := p_PillPackProfileID
        Try p_PillPackProfileID := ScrubString(p_PillPackProfileID)
        p_PillPackProfileID := ""

        FindAndPasteField("PillPack Profile ID", LocalProfileID)

        If !VerifyFieldContent(LocalProfileID) {
            Try LocalProfileID := ScrubString(LocalProfileID)
            LocalProfileID := ""
            MsgBox("Customer ID paste verification FAILED.`n`nAborting and wiping all data.", "Field Verify Failed", 16)
            WipeAllData()
            Return
        }
        Try LocalProfileID := ScrubString(LocalProfileID)
        LocalProfileID := ""

        ; Select customer name from profile lookup results
        If !SelectCustomerName(false) {
            MsgBox("WARNING: Could not confirm the customer name was selected after all retries.`n`nVerify the 'Customer first name' field is populated before submitting the ticket.", "Customer Selection Warning", 48)
            ReactivateChrome()
        }

        Sleep(300)

        ; --- Compute Due By (arrival date for non-Amazon), then set Priority ---
        DueByVal := ""
        If (Type(p_ArrivalParsed) == "Object" && p_ArrivalParsed.formatted != "")
            DueByVal := p_ArrivalParsed.formatted
        If (DueByVal == "")
            DueByVal := p_ShipDate
        SetPriorityIfDueToday(DueByVal)
        FindAndType("Due By", DueByVal)
        DueByVal := ""
        Try p_ShipDate := ScrubString(p_ShipDate)
        p_ShipDate := ""

        ; --- Fill "RTF - Next Fill Date" calendar (skip if empty) ---
        If (p_NextFillDate != "") {
            NextFillParsed := ""
            If RegExMatch(p_NextFillDate, "^(\d{1,2})/(\d{1,2})/(\d{4})$", &NFMatch)
                NextFillParsed := {year: Integer(NFMatch[3]), month: Integer(NFMatch[1]), day: Integer(NFMatch[2])}
            FindAndSetCalendar("RTF - Next Fill Date", NextFillParsed)
        }
        p_ArrivalDate   := ""
        p_ArrivalParsed := ""

        ; --- Fill form fields ---

        ; RTF - Reason
        FindAndTypeNoEnter("RTF - Reason", p_RTFReason)

        ; Customer Phone Number
        If (p_CustomerPhone != "") {
            FindAndTypeNoEnter("Customer Phone", p_CustomerPhone)
            Try p_CustomerPhone := ScrubString(p_CustomerPhone)
            p_CustomerPhone := ""
        }

        ; --- Fill comment fields (labels pre-exist in form) ---

        ; Customer ID
        FindAndPaste("Customer ID", p_Identifier)
        Try p_Identifier := ScrubString(p_Identifier)
        p_Identifier := ""

        ; Rx# (next line below Customer ID)
        If (p_SelectedRxNumber != "")
            DownEndAndPaste(p_SelectedRxNumber)
        Else {
            Send("{Down}")
            Sleep(100)
        }
        Try p_SelectedRxNumber := ScrubString(p_SelectedRxNumber)
        p_SelectedRxNumber := ""

        ; Drug Name and Strength
        FindAndPaste("Drug Name and Strength", p_MedName)
        Try p_MedName := ScrubString(p_MedName)
        p_MedName := ""

        ; Red Flag(s) Identified
        FindAndPaste("Red Flag(s) Identified", p_RTFReason)
        Try p_RTFReason := ScrubString(p_RTFReason)
        p_RTFReason := ""

        ; Last Filled Date (skip if empty)
        If (p_LastFillDate != "") {
            FindAndPaste("Last Filled Date", p_LastFillDate)
            Try p_LastFillDate := ScrubString(p_LastFillDate)
            p_LastFillDate := ""
        }

        ; Last Filled Day Supply (skip if empty)
        If (p_LastFillDays != "") {
            FindAndPaste("Last Filled Day Supply", p_LastFillDays)
            Try p_LastFillDays := ScrubString(p_LastFillDays)
            p_LastFillDays := ""
        }

        ; Last Filled Pharmacy (skip if empty)
        If (p_LastPharmacy != "") {
            FindAndPaste("Last Filled Pharmacy", p_LastPharmacy)
            Try p_LastPharmacy := ScrubString(p_LastPharmacy)
            p_LastPharmacy := ""
        }

        ; Next Fill Date if applicable (skip if empty)
        If (p_NextFillDate != "") {
            FindAndPaste("Next Fill Date if applicable", p_NextFillDate)
            Try p_NextFillDate := ScrubString(p_NextFillDate)
            p_NextFillDate := ""
        }

        ; Shipment Link
        FindAndPaste("Shipment Link", p_ShipmentURL)
        Try p_ShipmentURL := ScrubString(p_ShipmentURL)
        p_ShipmentURL := ""

        ; --- Final GUI: inform user to verify and deny PDMP ---
        PauseStart()
        MsgBox("Ticket form populated - DO NOT SUBMIT YET.`n`n*** VERIFY ALL INFORMATION IS CORRECT ***`n`nAfter submitting the ticket, ensure you DENY the PDMP check for this patient.", "Refusal to Fill - Complete", 48+262144)
        PauseEnd()

        WipeAllData()
    }
}

; === TACHYON TICKET SECTION FOLLOWS ===

; ==============================================================================
; Tachyon Ticket GUI Callbacks
; ==============================================================================
ShipStartMedOK(btn, info) {
    global g_ShipMedList, g_ShipSelectedIdx
    global g_MedicationName, g_RxNumber, g_PrescriptionType
    global g_ShipGUIDropCount, g_ShipGUIAction

    g_ShipGUIAction := "ok"
    saved := btn.Gui.Submit()

    selectedMeds := []
    Loop g_ShipGUIDropCount {
        propName := "DDL" . A_Index
        val := saved.%propName%
        If (val == "" || val == "None" || val == "-- Select a medication --")
            Continue
        For idx, med in g_ShipMedList {
            If (med.name . "  |  NDC: " . med.ndc == val) {
                selectedMeds.Push(med)
                Break
            }
        }
    }

    g_MedicationName   := ""
    g_RxNumber         := ""
    g_PrescriptionType := ""
    For i, med in selectedMeds {
        sep := (i > 1) ? " & " : ""
        g_MedicationName .= sep . med.name
        g_RxNumber       .= sep . med.rx
        If (i == 1)
            g_PrescriptionType := med.section
    }
    g_ShipSelectedIdx := selectedMeds.Length
    Try btn.Gui.Destroy()
}

ShipStartMedCancel(btn, info := "") {
    global g_ShipSelectedIdx, g_ShipGUIAction
    If g_ShipGUIAction == "add"
        Return
    g_ShipGUIAction   := "cancel"
    g_ShipSelectedIdx := -1
    If (btn is Gui) {
        Try btn.Destroy()
    } Else {
        Try btn.Gui.Destroy()
    }
}

ShipStartMedAdd(btn, info) {
    global g_ShipGUIAction, g_ShipGUISaved, g_ShipGUIDropCount, g_ShipGUIDDLs
    g_ShipGUIAction := "add"
    g_ShipGUISaved  := []

    If IsObject(g_ShipGUIDDLs) {
        For ddl in g_ShipGUIDDLs {
            txt := ""
            Try txt := ddl.Text
            g_ShipGUISaved.Push(txt)
        }
    }

    While (g_ShipGUISaved.Length < g_ShipGUIDropCount)
        g_ShipGUISaved.Push("")

    g_ShipGUIDDLs := []
    Try btn.Gui.Destroy()
}

; ==============================================================================
; PROJECT: Extract Group ID from Nexia -> Open PillPack Admin -> Tachyon Ticket
; Trigger: Ctrl + Shift + Alt + t (^+!t)
; ==============================================================================
^+!t::
{
    global g_RxNumber, g_GroupID, g_PatientName, g_Batch, g_PersonID
    global g_PrescriptionType, g_ShipDate, g_ArrivalDate, g_ArrivalParsed
    global g_ReasonForChange, g_ReasonIndex, g_ReasonOptions, g_PrescriptionOptions
    global g_MedicationName, g_ChangeNeeded, g_ShipmentURL
    global g_BlueprintName, g_IsAmazon, g_AdminURL, g_PillPackProfileID
    global g_UFRReason, g_UFRReasonIndex, g_UFROptions
    global g_IsRedShipment, g_StartedOnShipment
    global g_ShipMedList, g_ShipSelectedIdx
    global g_ShipGUIDropCount, g_ShipGUIAction, g_ShipGUISaved, g_ShipGUIDDLs

    StartWatchdog()

    ; -------------------------------------------------------------------------
    ; STEP 0 & 1: PRE-FLIGHT CHECK & DATA EXTRACTION
    ; Auto-detects: Nexia (View Rx / Track Rxs) or Chrome shipment page
    ; -------------------------------------------------------------------------
    KeyWait("Control")
    KeyWait("Alt")
    KeyWait("Shift")

    ActiveTitle := WinGetTitle("A")
    StartedOnDetails := false
    g_StartedOnShipment := false

    If InStr(ActiveTitle, "View Rx") {
        StartedOnDetails := true
    } Else If InStr(ActiveTitle, "Track Rxs") {
        StartedOnDetails := false
    } Else If WinActive("ahk_exe chrome.exe") {
        ; Check if current Chrome page is a shipment/pillpack page
        urlCaptured := false
        Loop 4 {
            Send("!d")
            Sleep(250)
            A_Clipboard := ""
            Send("^a")
            Sleep(50)
            Send("^c")
            If ClipWait(1.5) {
                TryURL := Trim(A_Clipboard)
                A_Clipboard := ""
                If RegExMatch(TryURL, "^https?://") {
                    If (InStr(TryURL, "shipment") || InStr(TryURL, "pillpack")) {
                        g_StartedOnShipment := true
                        g_ShipmentURL := TryURL
                        CleanURL := TryURL
                        QPos := InStr(CleanURL, "?")
                        If QPos
                            CleanURL := SubStr(CleanURL, 1, QPos - 1)
                        HPos := InStr(CleanURL, "#")
                        If HPos
                            CleanURL := SubStr(CleanURL, 1, HPos - 1)
                        LastSlash := InStr(CleanURL, "/", false, -1)
                        If LastSlash > 0
                            g_GroupID := Trim(SubStr(CleanURL, LastSlash + 1))
                        CleanURL := ""
                    }
                    urlCaptured := true
                    Break
                }
                TryURL := ""
            } Else {
                A_Clipboard := ""
            }
            Send("{Esc}")
            Sleep(200)
        }
        Send("{Esc}")
        Sleep(100)

        If (!urlCaptured) {
            MsgBox("STOPPING: Could not read a valid URL from Chrome address bar.", "Error", 16)
            WipeAllData()
            Return
        }

        If (!g_StartedOnShipment) {
            MsgBox("STOPPING: Chrome is active but this does not appear to be a shipment page.`n`nYou must be on a Nexia screen (View Rx / Track Rxs) or a PillPack shipment page.", "Wrong Screen", 16)
            WipeAllData()
            Return
        }
    } Else {
        MsgBox("STOPPING: You must be on the 'Track Rxs' or 'View Rx' screen in Nexia, or a PillPack shipment page in Chrome.`n`nCurrent Window: " . ActiveTitle, "Wrong Screen", 16)
        WipeAllData()
        Return
    }

    If (StartedOnDetails) {
        PageText := WinGetText("A")

        If RegExMatch(PageText, "([A-Za-z0-9]{24})\s*[\r\n]+\s*Group:", &MatchGrp)
            g_GroupID := MatchGrp[1]

        If RegExMatch(PageText, "(\d{9,10})(?:[/-]\d{2})?\s*[\r\n]+\s*Rx/Refill:", &MatchRx)
            g_RxNumber := MatchRx[1]

        If RegExMatch(PageText, "Medication:\s*[\r\n]+([A-Z][A-Z0-9 \/\.%,()&+\-]+\d+[A-Z0-9 \/\.%,()&+\-]+(?:INJ|TAB|CAP|ML|MG|MCG|CREAM|GEL|PEN|PATCH|SOLUTION|SOLN|SUSPENSION|SUSP|TABLET|CAPSULE|SFTGL|OINT|SYRUP|LIQ|DROP|LOTION|SUPP|INHALER|SPRAY|VIAL|KIT|SENSOR|STRIP|LANCET|SYRINGE|NEEDLE|DEVICE|DEV|FOAM|POWDER|SHAMPOO|WASH|PASTE|PUMP|MISC|PACK|PK|CART|SYS)(?:[A-Z ]*)?)", &MatchMed)
            g_MedicationName := Trim(MatchMed[1])
        PageText := ""
    }
    Else If (!g_StartedOnShipment) {
        A_Clipboard := ""
        Send("^c")

        If !ClipWait(2) {
            MsgBox("Failed to copy selected line. Make sure you clicked the line to highlight it.", "Error", 16)
            WipeAllData()
            Return
        }
        MarkClipboardExcluded()

        CopiedData := A_Clipboard
        A_Clipboard := ""

        If RegExMatch(CopiedData, "(\d{9,10})[/-]\d{2}\t([A-Za-z0-9]{24})", &Match) {
            g_RxNumber := Match[1]
            g_GroupID := Match[2]
        } Else If RegExMatch(CopiedData, "\b([A-Za-z0-9]{24})\b", &Match) {
            g_GroupID := Match[1]
        }

        If RegExMatch(CopiedData, "([A-Z][A-Z0-9 \/\.%,()&+\-]+\d+[A-Z0-9 \/\.%,()&+\-]+(?:INJ|TAB|CAP|ML|MG|MCG|CREAM|GEL|PEN|PATCH|SOLUTION|SOLN|SUSPENSION|SUSP|TABLET|CAPSULE|SFTGL|OINT|SYRUP|LIQ|DROP|LOTION|SUPP|INHALER|SPRAY|VIAL|KIT|SENSOR|STRIP|LANCET|SYRINGE|NEEDLE|DEVICE|DEV|FOAM|POWDER|SHAMPOO|WASH|PASTE|PUMP|MISC|PACK|PK|CART|SYS)(?:[A-Z ]*)?)", &MedMatch) {
            g_MedicationName := Trim(MedMatch[1])
        }

        CopiedData := ScrubString(CopiedData)
        CopiedData := ""
    }

    ; =========================================================================
    ; SHIPMENT-START BRANCH: Scrape page, parse meds, show selection GUI
    ; (Skips Steps 2-5 of the Nexia flow; merges at Step 6)
    ; =========================================================================
    If (g_StartedOnShipment) {
        ; Focus page and scrape
        Send("!d")
        Sleep(50)
        SendText("javascript:window.focus();void(0);")
        Sleep(50)
        Send("{Enter}")
        Sleep(200)

        PageText := ""

        If !WaitForShipmentPage(&PageText) {
            MsgBox("STOPPING: Could not read the shipment page content.", "Error", 16)
            WipeAllData()
            Return
        }

        ; Extract patient name, ship date, arrival date
        g_PatientName := ""
        g_ShipDate    := ""
        g_ArrivalParsed := ""
        g_IsRedShipment := false

        If RegExMatch(PageText, "i)([a-zA-Z\-']+(?:[ \t]+[a-zA-Z\-']+)*)\s*(?:\([^)]+\))?\s*[-—–‐]\s*(?:Unscheduled|Scheduled)", &NameMatch) {
            g_PatientName := Trim(NameMatch[1])
            g_PatientName := RegExReplace(g_PatientName, "\s+", " ")
        }

        If RegExMatch(PageText, "i)Ship\s*Date[:\s]+([A-Za-z,\s\d]+?)(?:\r|\n|Arrival)", &ShipMatch) {
            RawShipDate := Trim(ShipMatch[1])
            ShipParsed  := ParseDateFromText(RawShipDate)
            g_ShipDate  := ShipParsed.formatted
        }

        If RegExMatch(PageText, "i)Arrival\s*Date[:\s]+([A-Za-z,\s\d]+?)(?:\s*\(|Shipping|\r|\n)", &ArrivalMatch) {
            RawArrivalDate  := Trim(ArrivalMatch[1])
            g_ArrivalParsed := ParseDateFromText(RawArrivalDate)
        }

        If RegExMatch(PageText, "i)Status[:\s]+([^\r\n]+)", &StatusMatch) {
            StatusVal := Trim(StatusMatch[1])
            g_IsRedShipment := (InStr(StatusVal, "Red") > 0)
            StatusVal := ""
        }

        ; Parse medications from Bulk and Packet sections
        g_ShipMedList := []

        BulkStart := InStr(PageText, "Dispensed Bulk Meds")
        If (BulkStart > 0) {
            BulkEnd := InStr(PageText, "Pack Insert Prompts", , BulkStart)
            BulkSection := (BulkEnd > 0)
                ? SubStr(PageText, BulkStart, BulkEnd - BulkStart)
                : SubStr(PageText, BulkStart)

            startPos := 1
            While RegExMatch(BulkSection, "([^\r\n]+?) NDC:\s*(\d{11})", &MedMatch, startPos) {
                medName := Trim(MedMatch[1])
                ndc     := MedMatch[2]
                If InStr(medName, "Med Description") || InStr(medName, "Description") || StrLen(medName) < 3 {
                    startPos := MedMatch.Pos + MedMatch.Len
                    Continue
                }
                followRaw  := SubStr(BulkSection, MedMatch.Pos + MedMatch.Len, 600)
                nextNDCPos := RegExMatch(followRaw, "\bNDC:\s*\d{11}")
                followText := (nextNDCPos > 0) ? SubStr(followRaw, 1, nextNDCPos - 1) : followRaw
                rxNum := ""
                If RegExMatch(followText, "(\d{7,10})\s*/", &RxMatch)
                    rxNum := RxMatch[1]
                Else If RegExMatch(followText, "Rx\s*#?\s*:?\s*(\d{7,10})", &RxMatch)
                    rxNum := RxMatch[1]
                Else If RegExMatch(followText, "(?<!\d)(\d{7,10})(?!\d)", &RxMatch)
                    rxNum := RxMatch[1]
                isHold := (InStr(followText, "Held by Wms Rx Updates Bot") > 0)
                medName := Trim(RegExReplace(medName, "\([^)]*\)", ""))
                medName := RegExReplace(medName, "\s+", " ")
                isDupe := false
                For existing in g_ShipMedList {
                    If (existing.ndc == ndc) {
                        isDupe := true
                        Break
                    }
                }
                If (!isDupe)
                    g_ShipMedList.Push({name: medName, ndc: ndc, rx: rxNum, isHold: isHold, section: "Bulk"})
                startPos := MedMatch.Pos + MedMatch.Len
            }
            BulkSection := ScrubString(BulkSection)
            BulkSection := ""
        }

        HasPacketMeds := false
        PacketStart := InStr(PageText, "Dispensed Packet Meds")
        If (PacketStart > 0) {
            HasPacketMeds := true
            PacketEnd := InStr(PageText, "Expected Bulk Meds", , PacketStart)
            PacketSection := (PacketEnd > 0)
                ? SubStr(PageText, PacketStart, PacketEnd - PacketStart)
                : SubStr(PageText, PacketStart)

            startPos := 1
            While RegExMatch(PacketSection, "([^\r\n]+?) NDC:\s*(\d{11})", &MedMatch, startPos) {
                medName := Trim(MedMatch[1])
                ndc     := MedMatch[2]
                If InStr(medName, "Med Description") || InStr(medName, "Description") || StrLen(medName) < 3 {
                    startPos := MedMatch.Pos + MedMatch.Len
                    Continue
                }
                followRaw  := SubStr(PacketSection, MedMatch.Pos + MedMatch.Len, 600)
                nextNDCPos := RegExMatch(followRaw, "\bNDC:\s*\d{11}")
                followText := (nextNDCPos > 0) ? SubStr(followRaw, 1, nextNDCPos - 1) : followRaw
                rxNum := ""
                If RegExMatch(followText, "(\d{7,10})\s*/", &RxMatch)
                    rxNum := RxMatch[1]
                Else If RegExMatch(followText, "Rx\s*#?\s*:?\s*(\d{7,10})", &RxMatch)
                    rxNum := RxMatch[1]
                Else If RegExMatch(followText, "(?<!\d)(\d{7,10})(?!\d)", &RxMatch)
                    rxNum := RxMatch[1]
                medName := Trim(RegExReplace(medName, "\([^)]*\)", ""))
                medName := RegExReplace(medName, "\s+", " ")
                isDupe := false
                For existing in g_ShipMedList {
                    If (existing.ndc == ndc) {
                        isDupe := true
                        Break
                    }
                }
                If (!isDupe)
                    g_ShipMedList.Push({name: medName, ndc: ndc, rx: rxNum, isHold: false, section: "Packet"})
                startPos := MedMatch.Pos + MedMatch.Len
            }
            PacketSection := ScrubString(PacketSection)
            PacketSection := ""
        }

        PageText := ScrubString(PageText)
        PageText := ""

        If (g_ShipMedList.Length == 0) {
            MsgBox("STOPPING: No medications with NDC numbers found on this page.", "No Meds Found", 16)
            WipeAllData()
            Return
        }

        If (g_PatientName == "") {
            MsgBox("STOPPING: Could not find the customer's name on this page.", "Extraction Failed", 16)
            WipeAllData()
            Return
        }

        ; Med selection: auto-select if 1 med + no packets, else show GUI
        If (g_ShipMedList.Length == 1 && !HasPacketMeds) {
            med := g_ShipMedList[1]
            g_MedicationName   := med.name
            g_RxNumber         := med.rx
            g_PrescriptionType := med.section
            g_ShipSelectedIdx  := 1
        } Else {
            PauseStart()
            SplitResult := MsgBox("This shipment has more than 1 medication (or has packet medications).`n`nEnsure other medications should not be split off before continuing.`n`nPress OK to continue, or Cancel to abort.", "Multiple Medications Detected", 49)
            PauseEnd()
            If (SplitResult = "Cancel") {
                WipeAllData()
                Return
            }
            ReactivateChrome()

            medDisplayStrings := []
            For idx, med in g_ShipMedList
                medDisplayStrings.Push(med.name . "  |  NDC: " . med.ndc)

            holdMeds := []
            For idx, med in g_ShipMedList {
                If med.isHold
                    holdMeds.Push(idx)
            }
            holdCount := holdMeds.Length

            items_required := []
            If holdCount == 0
                items_required.Push("-- Select a medication --")
            For s in medDisplayStrings
                items_required.Push(s)

            items_optional := ["None"]
            For s in medDisplayStrings
                items_optional.Push(s)

            canAddMore := (g_ShipMedList.Length > 1)
            g_ShipGUIAction    := ""
            g_ShipGUIDropCount := Max(holdCount, 1)
            g_ShipGUISaved     := []
            g_ShipGUIDDLs      := []
            g_ShipSelectedIdx  := 0

            While true {
                MedGui := Gui("+AlwaysOnTop", "Tachyon Ticket - Medication Selection")
                MedGui.SetFont("s11 Bold")
                MedGui.Add("Text", "w490 Center", "Select Medication for Tachyon Ticket")
                MedGui.SetFont("s10 Norm")
                MedGui.Add("Text", "w490", "")

                If (g_ShipGUIDropCount == 1 && holdCount == 0)
                    MedGui.Add("Text", "w490", "Select the medication to submit the ticket for:")
                Else If (g_ShipGUIDropCount == 1 && holdCount == 1)
                    MedGui.Add("Text", "w490", "Inventory hold detected. Medication auto-selected:")
                Else
                    MedGui.Add("Text", "w490", "Select the medication(s) to submit the ticket for:")

                g_ShipGUIDDLs := []
                Loop g_ShipGUIDropCount {
                    i       := A_Index
                    isFirst := (i == 1)
                    items   := isFirst ? items_required : items_optional

                    ddl := MedGui.Add("DropDownList", "w490 vDDL" . i, items)
                    g_ShipGUIDDLs.Push(ddl)

                    selText := ""
                    If (i <= g_ShipGUISaved.Length && g_ShipGUISaved[i] != "")
                        selText := g_ShipGUISaved[i]
                    Else If (i <= holdMeds.Length) {
                        hm      := g_ShipMedList[holdMeds[i]]
                        selText := hm.name . "  |  NDC: " . hm.ndc
                    }

                    If (selText != "") {
                        For k, s in items {
                            If (s == selText) {
                                ddl.Choose(k)
                                Break
                            }
                        }
                    } Else {
                        ddl.Choose(1)
                    }
                }

                MedGui.Add("Text", "w490", "")
                BtnOK     := MedGui.Add("Button", "w100", "OK")
                BtnCancel := MedGui.Add("Button", "x+20 w100", "Cancel")
                If (canAddMore && g_ShipGUIDropCount < g_ShipMedList.Length)
                    BtnAdd := MedGui.Add("Button", "x+20 w150", "Add Medication")

                BtnOK.OnEvent("Click", ShipStartMedOK)
                BtnCancel.OnEvent("Click", ShipStartMedCancel)
                If (canAddMore && g_ShipGUIDropCount < g_ShipMedList.Length)
                    BtnAdd.OnEvent("Click", ShipStartMedAdd)
                MedGui.OnEvent("Close", GUICrash)

                g_ShipGUIAction := ""
                PauseStart()
                MedGui.Show()
                While WinExist("Tachyon Ticket - Medication Selection")
                    Sleep(100)
                PauseEnd()

                If g_ShipGUIAction == "add" {
                    g_ShipGUIDropCount++
                    Continue
                }
                Break
            }

            If (g_ShipSelectedIdx == -1) {
                WipeAllData()
                Return
            }
            If (g_MedicationName == "") {
                MsgBox("No medication was selected. Aborting.", "Error", 16)
                WipeAllData()
                Return
            }
        }

        If (g_RxNumber == "") {
            MsgBox("WARNING: Could not extract Rx number. You may need to fill it manually.", "Extraction Warning", 48)
            ReactivateChrome()
        }

        ; Shipment-start flow complete — jump to Step 6 (open customer name tab)
        Goto ShipmentMergeStep6
    }

    ; -------------------------------------------------------------------------
    ; STEP 2: FINAL DATA VALIDATION (Nexia flow)
    ; -------------------------------------------------------------------------
    If (g_GroupID == "") {
        MsgBox("Could not extract the 24-character Group ID.`n`nAborting.", "Extraction Failed", 16)
        WipeAllData()
        Return
    }

    If (g_MedicationName == "") {
        MsgBox("WARNING: Could not extract Medication Name. You may need to enter it manually.", "Extraction Warning", 48)
    }

    ; -------------------------------------------------------------------------
    ; STEP 2b: OPEN DETAILS SCREEN (If we started on Track Rxs)
    ; -------------------------------------------------------------------------
    If (!StartedOnDetails) {
        Send("{F8}")
        If WinWait("Parent or Portion Rx Selection", , 2) {
            WinActivate("Parent or Portion Rx Selection")
            Sleep(200)
            ControlClick("Parent Rx", "Parent or Portion Rx Selection")
            Sleep(1000)
        } Else {
            Sleep(1000)
        }
    }

    ; -------------------------------------------------------------------------
    ; STEP 3: Open Chrome with the PillPack shipment URL (save URL for later)
    ; -------------------------------------------------------------------------
    g_ShipmentURL := "https://admin.pillpack.com/admin/shipments/" . g_GroupID
    g_GroupID := ""

    If WinExist("ahk_exe chrome.exe") {
        WinActivate("ahk_exe chrome.exe")
        Sleep(300)
        Send("^t")
        Sleep(200)
        Send("!d")    ; V2.3.8: belt-and-suspenders address bar focus
        Sleep(100)
        SendText(g_ShipmentURL)
        Send("{Enter}")
    } Else {
        Run("chrome.exe " . g_ShipmentURL)
    }

    ; -------------------------------------------------------------------------
    ; STEP 4: Wait for Shipment Page to Load & Scrape (with login check via URL)
    ; -------------------------------------------------------------------------
    Sleep(500)

    Send("!d")
    Sleep(100)
    Send("^c")
    Sleep(100)

    If ClipWait(1) {
        CurrentURL := A_Clipboard
        A_Clipboard := ""
        If InStr(CurrentURL, "okta") {
            PauseStart()
            MsgBox("PillPack Sign-in required!`n`n1. Please log in manually.`n2. Click OK on this popup ONLY AFTER you see the shipment page.", "Login Paused", 64+262144)
            PauseEnd()
            Sleep(300)
            WinActivate("ahk_exe chrome.exe")
            Sleep(300)
        }
        CurrentURL := ""
    }

    WinActivate("ahk_exe chrome.exe")
    Sleep(150)
    Send("!d")
    Sleep(50)
    SendText("javascript:window.focus();void(0);")
    Sleep(50)
    Send("{Enter}")
    Sleep(200)

    PageText := ""

    If !WaitForShipmentPage(&PageText) {
        MsgBox("STOPPING: The shipment page took too long to load, or focus was lost.", "Timeout", 16)
        WipeAllData()
        Return
    }

    ; -------------------------------------------------------------------------
    ; STEP 5: Extract Patient Name, Ship Date, Arrival Date, and RED status
    ; -------------------------------------------------------------------------
    g_PatientName := ""
    g_ShipDate := ""
    g_ArrivalDate := ""
    g_ArrivalParsed := ""
    g_IsRedShipment := false

    If RegExMatch(PageText, "i)([a-zA-Z\-']+(?:[ \t]+[a-zA-Z\-']+)*)\s*(?:\([^)]+\))?\s*[-—–‐]\s*(?:Unscheduled|Scheduled)", &NameMatch) {
        g_PatientName := Trim(NameMatch[1])
        g_PatientName := RegExReplace(g_PatientName, "\s+", " ")
    }

    If RegExMatch(PageText, "i)Ship\s*Date[:\s]+([A-Za-z,\s\d]+?)(?:\r|\n|Arrival)", &ShipMatch) {
        RawShipDate := Trim(ShipMatch[1])
        ShipParsed := ParseDateFromText(RawShipDate)
        g_ShipDate := ShipParsed.formatted
    }

    If RegExMatch(PageText, "i)Arrival\s*Date[:\s]+([A-Za-z,\s\d]+?)(?:\s*\(|Shipping|\r|\n)", &ArrivalMatch) {
        RawArrivalDate := Trim(ArrivalMatch[1])
        g_ArrivalParsed := ParseDateFromText(RawArrivalDate)
        g_ArrivalDate := g_ArrivalParsed.formatted
    }

    ; Extract shipment status for RED shipment detection.
    ; Only the boolean result is stored -- raw text cleared immediately (HIPAA).
    If RegExMatch(PageText, "i)Status[:\s]+([^\r\n]+)", &StatusMatch) {
        StatusVal := Trim(StatusMatch[1])
        g_IsRedShipment := (InStr(StatusVal, "Red") > 0)
        StatusVal := ""
    }

    ; Determine Prescription Type from page content (Bulk vs Packet)
    If (g_PrescriptionType == "" && g_MedicationName != "") {
        BulkPos := InStr(PageText, "Dispensed Bulk Meds")
        PacketPos := InStr(PageText, "Dispensed Packet Meds")
        If (BulkPos > 0 && InStr(PageText, g_MedicationName, , BulkPos))
            g_PrescriptionType := "Bulk"
        Else If (PacketPos > 0 && InStr(PageText, g_MedicationName, , PacketPos))
            g_PrescriptionType := "Packet"
    }
    If (g_PrescriptionType == "") {
        If InStr(PageText, "Dispensed Bulk Meds")
            g_PrescriptionType := "Bulk"
        Else If InStr(PageText, "Dispensed Packet Meds")
            g_PrescriptionType := "Packet"
    }

    PageText := ScrubString(PageText)
    PageText := ""

    If (g_PatientName == "") {
        MsgBox("STOPPING: Could not find the customer's name on this page.", "Extraction Failed", 16)
        WipeAllData()
        Return
    }

    If (g_ShipDate == "") {
        MsgBox("WARNING: Could not extract Ship Date. You may need to enter it manually.", "Warning", 48)
        ReactivateChrome()
    }

    If (g_ArrivalDate == "") {
        MsgBox("WARNING: Could not extract Arrival Date. You may need to enter it manually.", "Warning", 48)
        ReactivateChrome()
    }
    g_ArrivalDate := ""

    ; -------------------------------------------------------------------------
    ; STEP 6: Open Customer Name in New Tab
    ; -------------------------------------------------------------------------
    ShipmentMergeStep6:
    OpenCustomerInNewTab(g_PatientName)
    g_PatientName := ""

    ; -------------------------------------------------------------------------
    ; STEP 7: Wait for Customer Details Tab to Load & Scrape
    ; -------------------------------------------------------------------------
    DetailsText := ""

    If !WaitForDetailsPage(&DetailsText) {
        MsgBox("STOPPING: The customer details tab took too long to load.", "Timeout", 16)
        WipeAllData()
        Return
    }

    ; -------------------------------------------------------------------------
    ; STEP 7b: Capture the Admin URL
    ; -------------------------------------------------------------------------
    Send("!d")
    Sleep(100)
    A_Clipboard := ""
    Send("^c")
    Sleep(50)
    If ClipWait(1) {
        g_AdminURL := A_Clipboard
        A_Clipboard := ""
    }
    Send("{Esc}")
    Sleep(50)

    ; -------------------------------------------------------------------------
    ; STEP 8: Extract Batch
    ; -------------------------------------------------------------------------
    g_Batch := ""

    If RegExMatch(DetailsText, "i)Batch[\s:]+([A-Za-z0-9_]+)", &BatchMatch) {
        g_Batch := Trim(BatchMatch[1])
    }

    If (g_Batch == "") {
        MsgBox("STOPPING: Could not locate the 'Batch' type.", "Extraction Failed", 16)
        DetailsText := ScrubString(DetailsText)
        DetailsText := ""
        WipeAllData()
        Return
    }

    ; -------------------------------------------------------------------------
    ; STEP 9: Branch based on Batch type
    ; -------------------------------------------------------------------------
    If (g_Batch == "AMAZON") {
        g_IsAmazon := true
        g_Batch := ""
        If RegExMatch(DetailsText, "i)Person ID[\s:]+([^\r\n]+)", &IDMatch) {
            g_PersonID := Trim(IDMatch[1])
        }
        DetailsText := ScrubString(DetailsText)
        DetailsText := ""
        g_AdminURL := ""
        If (g_PersonID == "") {
            MsgBox("STOPPING: Could not locate the Person ID.", "Extraction Failed", 16)
            WipeAllData()
            Return
        }
    } Else {
        g_IsAmazon := false
        g_Batch := ""
        DetailsText := ScrubString(DetailsText)
        DetailsText := ""

        CleanURL := g_AdminURL
        QPos := InStr(CleanURL, "?")
        If (QPos)
            CleanURL := SubStr(CleanURL, 1, QPos - 1)
        HPos := InStr(CleanURL, "#")
        If (HPos)
            CleanURL := SubStr(CleanURL, 1, HPos - 1)
        LastSlash := InStr(CleanURL, "/", false, -1)
        If (LastSlash > 0)
            g_PillPackProfileID := Trim(SubStr(CleanURL, LastSlash + 1))
        CleanURL := ""

        If (g_PillPackProfileID == "" || g_AdminURL == "") {
            MsgBox("STOPPING: Could not extract the PillPack Profile ID from the admin URL.`n`nAdmin URL captured: " . (g_AdminURL == "" ? "(none)" : "yes"), "Extraction Failed", 16)
            WipeAllData()
            Return
        }
    }

    ; -------------------------------------------------------------------------
    ; STEP 10: Ask if ticket is still warranted
    ; -------------------------------------------------------------------------
    Result := MsgBox("After reviewing notes on the shipment and admin pages,`nis a ticket still warranted?", "Ticket Review", 4+32+262144)
    If (Result != "Yes") {
        WipeAllData()
        Return
    }

    ; -------------------------------------------------------------------------
    ; STEP 11: Ask user for inputs (GUI branches on Amazon vs non-Amazon)
    ; -------------------------------------------------------------------------
    If WinExist("ahk_exe chrome.exe") {
        WinActivate("ahk_exe chrome.exe")
        If !WinWaitActive("ahk_exe chrome.exe", , 2) {
            MsgBox("STOPPING: Could not refocus Chrome. Aborting.", "Focus Error", 16)
            WipeAllData()
            Return
        }
    }

    g_ReasonForChange := ""
    g_ReasonIndex := 0
    g_ChangeNeeded := ""
    g_UFRReason := ""
    g_UFRReasonIndex := 0

    If (g_IsAmazon) {
        InputGui := Gui("+AlwaysOnTop", "Clinical Review - Input")
        InputGui.SetFont("s10")

        ; Show Prescription Type only if it couldn't be auto-determined
        If (g_PrescriptionType == "") {
            InputGui.Add("Text", "w400", "Select Prescription Type:")
            PrescriptionDropdown := InputGui.Add("DropDownList", "w400 vPrescriptionChoice", g_PrescriptionOptions)
            PrescriptionDropdown.Choose(1)
        }

        InputGui.Add("Text", "w400", "")
        InputGui.Add("Text", "w400", "Select the FFU - Reason for Change:")
        ReasonDropdown := InputGui.Add("DropDownList", "w400 vReasonChoice", g_ReasonOptions)
        ReasonDropdown.Choose(1)

        InputGui.Add("Text", "w400", "")
        InputGui.Add("Text", "w400", "Enter the Change Needed:")
        InputGui.Add("Edit", "w400 h60 vChangeNeeded", "")

        InputGui.Add("Text", "w400", "")
        BtnOK := InputGui.Add("Button", "w100", "OK")
        BtnOK.OnEvent("Click", InputOK_Amazon)
        BtnCancel := InputGui.Add("Button", "x+20 w100", "Cancel")
        BtnCancel.OnEvent("Click", InputCancel_Amazon)
        InputGui.OnEvent("Close", GUICrash)

        InputOK_Amazon(btn, info) {
            global g_PrescriptionType, g_ReasonForChange, g_ReasonIndex, g_ReasonOptions, g_ChangeNeeded
            saved := btn.Gui.Submit()
            If (g_PrescriptionType == "")
                g_PrescriptionType := saved.HasProp("PrescriptionChoice") ? saved.PrescriptionChoice : "Bulk"
            g_ReasonForChange := saved.ReasonChoice
            g_ChangeNeeded := saved.ChangeNeeded
            For idx, opt in g_ReasonOptions {
                If (opt == g_ReasonForChange) {
                    g_ReasonIndex := idx
                    Break
                }
            }
        }

        InputCancel_Amazon(btn, info := "") {
            global g_ReasonForChange, g_ReasonIndex, g_ChangeNeeded
            g_ReasonForChange := ""
            g_ReasonIndex := -1
            g_ChangeNeeded := ""
            If (btn is Gui)
                btn.Destroy()
            Else
                btn.Gui.Destroy()
        }

        InputGui.Show()
        While WinExist("Clinical Review - Input")
            Sleep(100)

        If (g_ReasonIndex == -1) {
            WipeAllData()
            Return
        }
        If (g_ReasonForChange == "" || g_ReasonIndex == 0) {
            MsgBox("No reason selected. Aborting.", "Error", 16)
            WipeAllData()
            Return
        }
        If (g_ChangeNeeded == "") {
            MsgBox("No change needed entered. Aborting.", "Error", 16)
            WipeAllData()
            Return
        }

        If (g_ReasonForChange == "Days Supply Mismatch") {
            g_BlueprintName := "Day Supply or Quantity\Package Mismatch"
        } Else If (g_ReasonForChange == "NDC Out of Stock") {
            g_BlueprintName := "NDC Out of Stock"
        } Else {
            g_BlueprintName := "Clinical Review"
        }

    } Else {
        InputGui := Gui("+AlwaysOnTop", "Upstream Fix Request - Input")
        InputGui.SetFont("s10")

        InputGui.Add("Text", "w400", "Select the UFR - Error main issue:")
        UFRDropdown := InputGui.Add("DropDownList", "w400 vUFRChoice", g_UFROptions)
        UFRDropdown.Choose(1)

        InputGui.Add("Text", "w400", "")
        InputGui.Add("Text", "w400", "Enter the Explanation of the change needed:")
        InputGui.Add("Edit", "w400 h60 vChangeNeeded", "")

        InputGui.Add("Text", "w400", "")
        BtnOK := InputGui.Add("Button", "w100", "OK")
        BtnOK.OnEvent("Click", InputOK_UFR)
        BtnCancel := InputGui.Add("Button", "x+20 w100", "Cancel")
        BtnCancel.OnEvent("Click", InputCancel_UFR)
        InputGui.OnEvent("Close", GUICrash)

        InputOK_UFR(btn, info) {
            global g_UFRReason, g_UFRReasonIndex, g_UFROptions, g_ChangeNeeded
            saved := btn.Gui.Submit()
            g_UFRReason := saved.UFRChoice
            g_ChangeNeeded := saved.ChangeNeeded
            For idx, opt in g_UFROptions {
                If (opt == g_UFRReason) {
                    g_UFRReasonIndex := idx
                    Break
                }
            }
        }

        InputCancel_UFR(btn, info := "") {
            global g_UFRReason, g_UFRReasonIndex, g_ChangeNeeded
            g_UFRReason := ""
            g_UFRReasonIndex := 0
            g_ChangeNeeded := ""
            If (btn is Gui)
                btn.Destroy()
            Else
                btn.Gui.Destroy()
        }

        InputGui.Show()
        While WinExist("Upstream Fix Request - Input")
            Sleep(100)

        If (g_UFRReason == "" || g_UFRReasonIndex == 0) {
            MsgBox("No UFR-Error selected. Aborting.", "Error", 16)
            WipeAllData()
            Return
        }
        If (g_ChangeNeeded == "") {
            MsgBox("No explanation entered. Aborting.", "Error", 16)
            WipeAllData()
            Return
        }

        g_BlueprintName := "Upstream Fix Request"
    }

    ReactivateChrome()

    ; =========================================================================
    ; STEP 12: Open blueprint directly via URL
    ; =========================================================================
    DirectURL := ""
    ExpectedTask := ""
    If (g_IsAmazon) {
        ExpectedTask := "Task: Red Bulk Shipment Change"
        If (g_BlueprintName == "Day Supply or Quantity\Package Mismatch") {
            DirectURL := "https://cs.wolfgang.a2z.com/tachyon/PHARMACY/createNewIssue/46AB000-0170-4A6F-834B-D9CDDDF2A8CC"
        } Else If (g_BlueprintName == "NDC Out of Stock") {
            DirectURL := "https://cs.wolfgang.a2z.com/tachyon/PHARMACY/createNewIssue/23D85000-0170-4A6F-834B-F9CDDDF2A8BB"
        } Else If (g_PrescriptionType == "Packet") {
            DirectURL := "https://cs.wolfgang.a2z.com/tachyon/PHARMACY/createNewIssue/a4b28a6885854cc0be9b2bb18def80c9"
        } Else {
            DirectURL := "https://cs.wolfgang.a2z.com/tachyon/PHARMACY/createNewIssue/9F695381-7AB0-4D30-B78C-5378605B465B"
        }
    } Else {
        ExpectedTask := "Task: Fulfillment Fix-up"
        DirectURL := "https://cs.wolfgang.a2z.com/tachyon/PILLPACK/createNewIssue/360207892034"
    }

    Send("^t")
    Sleep(200)
    A_Clipboard := DirectURL
    Sleep(50)
    Send("^v")
    Sleep(100)
    Send("{Enter}")
    A_Clipboard := ""
    DirectURL := ""

    ; -------------------------------------------------------------------------
    ; STEP 13: Wait for ticket form to load
    ; -------------------------------------------------------------------------
    Loaded := false
    Loop 30 {
        Sleep(A_Index = 1 ? 1500 : 500)
        If CheckChromeTitleForLogin("Tachyon")
            Continue
        A_Clipboard := ""
        Send("^a")
        Sleep(50)
        Send("^c")
        If ClipWait(0.5) {
            PageText := A_Clipboard
            A_Clipboard := ""
            If InStr(PageText, ExpectedTask) {
                PageText := ""
                ClearSelection()
                Loaded := true
                break
            }
            If InStr(PageText, "Sign In") || InStr(PageText, "Username") || InStr(PageText, "Verify with your password") || InStr(PageText, "Powered by Okta") {
                PageText := ""
                PauseStart()
                MsgBox("Sign-in required!`n`n1. Please log in manually.`n2. Click OK on this popup ONLY AFTER you see the ticket form.", "Login Paused", 64+262144)
                PauseEnd()
                Sleep(300)
                WinActivate("ahk_exe chrome.exe")
                Sleep(300)
            }
            PageText := ""
            ClearSelection()
        } Else {
            ClearSelection()
        }
    }
    If (!Loaded) {
        MsgBox("STOPPING: Timed out waiting for the ticket form to load.`n`nExpected: " . ExpectedTask, "Timeout", 16)
        WipeAllData()
        Return
    }

    ; -------------------------------------------------------------------------
    ; STEP 20: Navigate to Person ID / PillPack Profile ID field and type
    ; -------------------------------------------------------------------------
    IDFieldLabel := g_IsAmazon ? "Person ID" : "PillPack Profile ID"
    LocalCustID := g_IsAmazon ? g_PersonID : g_PillPackProfileID

    If !TypeAndVerifyProfileID(IDFieldLabel, LocalCustID) {
        LocalCustID := ""
        MsgBox("Customer ID paste verification FAILED.`n`nAborting and wiping all data.", "Field Verify Failed", 16)
        WipeAllData()
        Return
    }
    LocalCustID := ""

    ; -------------------------------------------------------------------------
    ; STEP 21: Select customer name from profile lookup results
    ; -------------------------------------------------------------------------
    If !SelectCustomerName(g_IsAmazon) {
        MsgBox("WARNING: Could not confirm the customer name was selected after all retries.`n`nVerify the 'Customer first name' field is populated before submitting the ticket.", "Customer Selection Warning", 48)
        ReactivateChrome()
    }

    Sleep(300)

    If (g_IsAmazon) {
        ; =====================================================================
        ; AMAZON FORM FILLING (Clinical Review / Day Supply blueprint)
        ; Unchanged from V2.3.4.
        ; =====================================================================

        ; STEP 22: "Priority" dropdown -> High if due today or earlier
        SetPriorityIfDueToday(g_ShipDate)

        ; STEP 23: "Due By" -- Ship Date
        FindAndType("Due", g_ShipDate)
        g_ShipDate := ""

        ; STEP 23: "FFU - Reason for Change" dropdown
        FindAndSelectDropdown("reason", g_ReasonIndex)
        g_ReasonIndex := 0
        g_ReasonForChange := ""

        ; STEP 24: "FFU - Original Promise Date" calendar
        FindAndSetCalendar("FFU - Original Promise Date", g_ArrivalParsed)
        g_ArrivalParsed := ""
        g_ArrivalDate := ""

        ; STEP 25: Medication Name
        FindAndPaste("Medication Name", g_MedicationName)
        g_MedicationName := ""

        ; STEP 26: Rx#
        DownAndPaste(g_RxNumber)
        g_RxNumber := ""

        ; STEP 27: Change Needed
        FindAndPaste("Change Needed", g_ChangeNeeded)
        g_ChangeNeeded := ""

        ; STEP 28: Shipment link and/or Order ID
        FindAndPaste("Shipment link and/or Order ID", g_ShipmentURL)
        g_ShipmentURL := ""
        g_PrescriptionType := ""

    } Else {
        ; =====================================================================
        ; NON-AMAZON FORM FILLING (Upstream Fix Request blueprint)
        ;
        ; Field order on the rendered form:
        ;   Subject      (auto-filled by blueprint -- not touched)
        ;   Due By       (Step 22-UFR: g_ShipDate)
        ;   UFR - Error  (Step 23-UFR: index-based dropdown)
        ;   UFR - Is this a RED shipment?  (Step 24-UFR: Yes / No dropdown)
        ;   Description box:
        ;     Medication name / Rx# / Shipment Link  (Step 25-UFR)
        ;     Admin Link                             (Step 26-UFR, Down arrow)
        ;     Explanation of the change needed       (Step 27-UFR, Down arrow)
        ; =====================================================================

        ; STEP 22-UFR-0: Compute Due By (arrival date for non-Amazon), then set Priority
        DueByVal := ""
        If (Type(g_ArrivalParsed) == "Object" && g_ArrivalParsed.formatted != "")
            DueByVal := g_ArrivalParsed.formatted
        If (DueByVal == "")
            DueByVal := g_ShipDate
        SetPriorityIfDueToday(DueByVal)
        FindAndType("Due", DueByVal)
        DueByVal := ""
        g_ShipDate := ""

        ; STEP 23-UFR: "UFR - Error" main-issue dropdown
        FindAndSelectDropdown("UFR - Error", g_UFRReasonIndex)
        g_UFRReasonIndex := 0
        g_UFRReason := ""

        ; STEP 24-UFR: "UFR - Is this a RED shipment?" Yes / No dropdown
        ; Observed dropdown order: [blank] -> Yes -> No
        RedIndex := g_IsRedShipment ? 2 : 3
        FindAndSelectDropdown("UFR - Is this a RED shipment", RedIndex)
        g_IsRedShipment := false

        ; STEP 25-UFR: Medication Name / Rx# / Shipment Link (first description line)
        FindAndPaste("Medication name", g_MedicationName . " / " . g_RxNumber . " / " . g_ShipmentURL)
        g_MedicationName := ""
        g_RxNumber := ""
        g_ShipmentURL := ""

        ; STEP 26-UFR: Admin Link (Down arrow to second description line)
        DownEndAndPaste(g_AdminURL)
        g_AdminURL := ""

        ; STEP 27-UFR: Explanation of the change needed (Down arrow to third line)
        DownEndAndPaste(g_ChangeNeeded)
        A_Clipboard := ""
        g_ChangeNeeded := ""
        Sleep(200)

        ; Clear remaining Amazon-only variables unused in this branch
        g_ArrivalDate := ""
        g_ArrivalParsed := ""
        g_ReasonForChange := ""
        g_ReasonIndex := 0
    }

    ; -------------------------------------------------------------------------
    ; STEP 29: Complete
    ; -------------------------------------------------------------------------
    WipeAllData()
    MsgBox("Ticket form populated.`n`n*** VERIFY ALL INFORMATION IS CORRECT BEFORE SUBMITTING ***`n`nReview every field (customer, dates, reason, medication, Rx#, etc.) and confirm the blueprint matches the case.", "Complete - Verify Before Submit", 64+262144)
}

; === NDC OOS TICKET SECTION FOLLOWS ===

; ==============================================================================
; NDC OOS Ticket GUI Callbacks
; ==============================================================================

NDCInputOK(btn, info) {
    global g_MedList, g_SelectedMedIdx
    global g_SelectedNDC, g_SelectedMedName, g_SelectedRxNumber
    global g_GUIDropCount, g_GUIAction
    g_GUIAction := "ok"
    saved := btn.Gui.Submit()   ; hides GUI, collects all vVariable values

    ; Collect every dropdown that holds a real med selection
    selectedMeds := []
    Loop g_GUIDropCount {
        propName := "DDL" . A_Index
        val := saved.%propName%
        If (val == "" || val == "None" || val == "-- Select a medication --")
            Continue
        For idx, med in g_MedList {
            If (med.name . "  |  NDC: " . med.ndc == val) {
                selectedMeds.Push(med)
                Break
            }
        }
    }

    ; Build " & "-joined combined strings
    g_SelectedMedName  := ""
    g_SelectedNDC      := ""
    g_SelectedRxNumber := ""
    For i, med in selectedMeds {
        sep := (i > 1) ? " & " : ""
        g_SelectedMedName  .= sep . med.name
        g_SelectedNDC      .= sep . med.ndc
        g_SelectedRxNumber .= sep . med.rx
    }
    g_SelectedMedIdx := selectedMeds.Length   ; 0 = nothing chosen; >0 = valid
    Try btn.Gui.Destroy()
}

NDCInputCancel(btn, info := "") {
    global g_SelectedMedIdx, g_GUIAction
    ; Guard: if g_GUIAction is already "add", Destroy() inside NDCInputAdd
    ; may have fired this handler as a side-effect.  Ignore it in that case.
    If g_GUIAction == "add"
        Return
    g_GUIAction      := "cancel"
    g_SelectedMedIdx := -1
    If (btn is Gui) {
        Try btn.Destroy()
    } Else {
        Try btn.Gui.Destroy()
    }
}

NDCInputAdd(btn, info) {
    global g_GUIAction, g_GUISavedSelections, g_GUIDropCount, g_GUIDDLs
    ; Set action flag FIRST so NDCInputCancel's guard fires correctly if
    ; the Destroy() call below triggers the Close event handler.
    g_GUIAction          := "add"
    g_GUISavedSelections := []

    ; Read each dropdown directly from the stored control references.
    ; Wrap each read in Try so a single bad control can't abort the hotkey.
    If IsObject(g_GUIDDLs) {
        For ddl in g_GUIDDLs {
            txt := ""
            Try
                txt := ddl.Text
            g_GUISavedSelections.Push(txt)
        }
    }

    ; Pad if for any reason we collected fewer entries than expected
    While (g_GUISavedSelections.Length < g_GUIDropCount)
        g_GUISavedSelections.Push("")

    ; Drop the references; the GUI is about to be destroyed
    g_GUIDDLs := []
    Try btn.Gui.Destroy()
}

; ==============================================================================
; PROJECT: Read OS Shipment Page -> Tachyon NDC OOS Ticket
; Trigger: Ctrl + Shift + Alt + O  (^+!o)
; ==============================================================================
^+!o::
{
    global g_GroupID, g_PatientName, g_PersonID
    global g_ShipDate, g_ArrivalDate, g_ArrivalParsed
    global g_ShipmentURL, g_IsAmazon, g_AdminURL, g_PillPackProfileID
    global g_IsRedShipment, g_ChangeNeeded
    global g_SelectedNDC, g_SelectedMedName, g_SelectedRxNumber
    global g_MedList, g_SelectedMedIdx
    global g_BlueprintName, g_ReasonIndex
    global g_GUIDropCount, g_GUIAction, g_GUISavedSelections, g_GUIDDLs

    StartWatchdog()

    ; -------------------------------------------------------------------------
    ; STEP 0: PRE-FLIGHT CHECK & AUTO-DETECTION
    ; Accepts: Nexia (View Rx / Track Rxs) OR Chrome shipment page
    ; -------------------------------------------------------------------------
    KeyWait("Control")
    KeyWait("Alt")
    KeyWait("Shift")

    ActiveTitle := WinGetTitle("A")
    StartedFromNexia := false

    If InStr(ActiveTitle, "View Rx") || InStr(ActiveTitle, "Track Rxs") {
        StartedFromNexia := true

        ; Extract GroupID from Nexia page
        If InStr(ActiveTitle, "View Rx") {
            PageText := WinGetText("A")
            If RegExMatch(PageText, "([A-Za-z0-9]{24})\s*[\r\n]+\s*Group:", &MatchGrp)
                g_GroupID := MatchGrp[1]
            PageText := ScrubString(PageText)
            PageText := ""
        } Else {
            ; Track Rxs — user must have a line highlighted
            A_Clipboard := ""
            Send("^c")
            If !ClipWait(2) {
                MsgBox("STOPPING: Could not copy selected line. Make sure you clicked the line to highlight it.", "Error", 16)
                WipeAllData()
                Return
            }
            MarkClipboardExcluded()
            CopiedData := A_Clipboard
            A_Clipboard := ""

            If RegExMatch(CopiedData, "(\d{9,10})[/-]\d{2}\t([A-Za-z0-9]{24})", &Match)
                g_GroupID := Match[2]
            Else If RegExMatch(CopiedData, "\b([A-Za-z0-9]{24})\b", &Match)
                g_GroupID := Match[1]
            CopiedData := ScrubString(CopiedData)
            CopiedData := ""
        }

        If (g_GroupID == "") {
            MsgBox("STOPPING: Could not extract the 24-character Group ID from Nexia.", "Extraction Failed", 16)
            WipeAllData()
            Return
        }

        ; Build shipment URL and open in Chrome
        g_ShipmentURL := "https://admin.pillpack.com/admin/shipments/" . g_GroupID

        If WinExist("ahk_exe chrome.exe") {
            WinActivate("ahk_exe chrome.exe")
            Sleep(300)
            Send("^t")
            Sleep(200)
            Send("!d")
            Sleep(100)
            SendText(g_ShipmentURL)
            Send("{Enter}")
        } Else {
            Run("chrome.exe " . g_ShipmentURL)
        }

        ; Wait for shipment page to load
        Sleep(500)
        Send("!d")
        Sleep(100)
        Send("^c")
        Sleep(100)
        If ClipWait(1) {
            CurrentURL := A_Clipboard
            A_Clipboard := ""
            If InStr(CurrentURL, "okta") {
                PauseStart()
                MsgBox("PillPack Sign-in required!`n`n1. Please log in manually.`n2. Click OK ONLY AFTER you see the shipment page.", "Login Paused", 64+262144)
                PauseEnd()
                Sleep(300)
                WinActivate("ahk_exe chrome.exe")
                Sleep(300)
            }
            CurrentURL := ""
        }

    } Else If WinActive("ahk_exe chrome.exe") {
        ; Started from Chrome — capture URL directly (original flow)
        Sleep(150)

        If !CaptureURLFromAddressBar(&g_ShipmentURL) {
            MsgBox("STOPPING: Could not read a valid URL from Chrome address bar.", "Error", 16)
            WipeAllData()
            Return
        }

        If !(InStr(g_ShipmentURL, "shipment") || InStr(g_ShipmentURL, "pillpack") || InStr(g_ShipmentURL, "amazon")) {
            MsgBox("STOPPING: This does not appear to be a shipment page.`n`nCaptured URL:`n" . g_ShipmentURL, "Wrong Page", 16)
            WipeAllData()
            Return
        }

        CleanURL := g_ShipmentURL
        QPos := InStr(CleanURL, "?")
        If QPos
            CleanURL := SubStr(CleanURL, 1, QPos - 1)
        HPos := InStr(CleanURL, "#")
        If HPos
            CleanURL := SubStr(CleanURL, 1, HPos - 1)
        LastSlash := InStr(CleanURL, "/", false, -1)
        If LastSlash > 0
            g_GroupID := Trim(SubStr(CleanURL, LastSlash + 1))
        CleanURL := ""

        If (g_GroupID == "") {
            MsgBox("STOPPING: Could not extract a Shipment/Group ID from the URL.", "Extraction Failed", 16)
            WipeAllData()
            Return
        }

    } Else {
        MsgBox("STOPPING: You must be on a Nexia screen (View Rx / Track Rxs) or a PillPack shipment page in Chrome.`n`nCurrent Window: " . ActiveTitle, "Wrong Screen", 16)
        WipeAllData()
        Return
    }

    ; -------------------------------------------------------------------------
    ; STEP 2: Focus page and read all text content
    ; -------------------------------------------------------------------------
    Send("!d")
    Sleep(50)
    SendText("javascript:window.focus();void(0);")
    Sleep(50)
    Send("{Enter}")
    Sleep(200)

    PageText := ""

    If !WaitForShipmentPage(&PageText) {
        MsgBox("STOPPING: Could not read the shipment page content.`n`nMake sure Chrome is focused on the correct shipment page.", "Error", 16)
        WipeAllData()
        Return
    }

    ; -------------------------------------------------------------------------
    ; STEP 3: Extract Patient Name, Ship Date, Arrival Date, RED status
    ; -------------------------------------------------------------------------
    g_PatientName   := ""
    g_ShipDate      := ""
    g_ArrivalDate   := ""
    g_ArrivalParsed := ""
    g_IsRedShipment := false

    If RegExMatch(PageText, "i)([a-zA-Z\-']+(?:[ \t]+[a-zA-Z\-']+)*)\s*(?:\([^)]+\))?\s*[-—–‐]\s*(?:Unscheduled|Scheduled)", &NameMatch) {
        g_PatientName := Trim(NameMatch[1])
        g_PatientName := RegExReplace(g_PatientName, "\s+", " ")
    }

    If RegExMatch(PageText, "i)Ship\s*Date[:\s]+([A-Za-z,\s\d]+?)(?:\r|\n|Arrival)", &ShipMatch) {
        RawShipDate  := Trim(ShipMatch[1])
        ShipParsed   := ParseDateFromText(RawShipDate)
        g_ShipDate   := ShipParsed.formatted
    }

    If RegExMatch(PageText, "i)Arrival\s*Date[:\s]+([A-Za-z,\s\d]+?)(?:\s*\(|Shipping|\r|\n)", &ArrivalMatch) {
        RawArrivalDate  := Trim(ArrivalMatch[1])
        g_ArrivalParsed := ParseDateFromText(RawArrivalDate)
        g_ArrivalDate   := g_ArrivalParsed.formatted
    }

    If RegExMatch(PageText, "i)Status[:\s]+([^\r\n]+)", &StatusMatch) {
        StatusVal       := Trim(StatusMatch[1])
        g_IsRedShipment := (InStr(StatusVal, "Red") > 0)
        StatusVal       := ""
    }

    If (g_PatientName == "") {
        MsgBox("STOPPING: Could not find the customer name on this page.`n`nVerify you are on the correct shipment page.", "Extraction Failed", 16)
        WipeAllData()
        Return
    }

    If (g_ShipDate == "")
        MsgBox("WARNING: Could not extract Ship Date. You may need to enter date fields manually.", "Warning", 48)

    If (g_ArrivalDate == "")
        MsgBox("WARNING: Could not extract Arrival Date. You will need to set the Due Date calendar manually.", "Warning", 48)

    ; -------------------------------------------------------------------------
    ; STEP 4: Scope page text to "Dispensed Bulk Meds" and parse meds
    ; -------------------------------------------------------------------------
    BulkStart := InStr(PageText, "Dispensed Bulk Meds")
    If (BulkStart == 0) {
        PageText := ScrubString(PageText)
        PageText := ""
        MsgBox("STOPPING: No 'Dispensed Bulk Meds' section was found on this page.`n`nThis shipment may not have any bulk medications, or you may be on the wrong page.", "No Bulk Meds Found", 16)
        WipeAllData()
        Return
    }

    BulkEnd     := InStr(PageText, "Pack Insert Prompts", , BulkStart)
    BulkSection := (BulkEnd > 0)
        ? SubStr(PageText, BulkStart, BulkEnd - BulkStart)
        : SubStr(PageText, BulkStart)

    PageText := ScrubString(PageText)
    PageText := ""

    g_MedList := []

    startPos := 1
    While RegExMatch(BulkSection, "([^\r\n]+?) NDC:\s*(\d{11})", &MedMatch, startPos) {
        medName := Trim(MedMatch[1])
        ndc     := MedMatch[2]

        If InStr(medName, "Med Description") || InStr(medName, "Description") {
            startPos := MedMatch.Pos + MedMatch.Len
            Continue
        }
        If StrLen(medName) < 3 {
            startPos := MedMatch.Pos + MedMatch.Len
            Continue
        }

        followRaw  := SubStr(BulkSection, MedMatch.Pos + MedMatch.Len, 600)
        nextNDCPos := RegExMatch(followRaw, "\bNDC:\s*\d{11}")
        followText := (nextNDCPos > 0) ? SubStr(followRaw, 1, nextNDCPos - 1) : followRaw

        rxNum := ""
        If RegExMatch(followText, "(\d{10})\s*/", &RxMatch)
            rxNum := RxMatch[1]

        isHold := (InStr(followText, "Held by Wms Rx Updates Bot") > 0)

        isDupe := false
        For existing in g_MedList {
            If (existing.ndc == ndc) {
                isDupe := true
                Break
            }
        }
        If (!isDupe)
            g_MedList.Push({name: medName, ndc: ndc, rx: rxNum, isHold: isHold})

        startPos := MedMatch.Pos + MedMatch.Len
    }

    BulkSection := ScrubString(BulkSection)
    BulkSection := ""

    If (g_MedList.Length == 0) {
        MsgBox("STOPPING: A 'Dispensed Bulk Meds' section was found but contained no medications with NDC numbers.`n`nVerify you are on the correct shipment page.", "No Meds in Section", 16)
        WipeAllData()
        Return
    }

    ; -------------------------------------------------------------------------
    ; STEP 5: Show user input GUI (or auto-select if 1 med on hold)
    ; -------------------------------------------------------------------------

    holdMeds := []
    For idx, med in g_MedList {
        If med.isHold
            holdMeds.Push(idx)
    }
    holdCount := holdMeds.Length

    ; Auto-select: 1 medication total AND it's on inventory hold -> skip GUI
    If (g_MedList.Length == 1 && holdCount == 1) {
        med := g_MedList[1]
        g_SelectedMedName  := med.name
        g_SelectedNDC      := med.ndc
        g_SelectedRxNumber := med.rx
        g_SelectedMedIdx   := 1
        If (g_SelectedRxNumber == "")
            MsgBox("WARNING: Could not extract an Rx number for the selected medication. You may need to fill it in manually.", "Extraction Warning", 48)
        Goto NDCOOSPostGUI
    }

    medDisplayStrings := []
    For idx, med in g_MedList
        medDisplayStrings.Push(med.name . "  |  NDC: " . med.ndc)

    items_required := []
    If holdCount == 0
        items_required.Push("-- Select a medication --")
    For s in medDisplayStrings
        items_required.Push(s)

    items_optional := ["None"]
    For s in medDisplayStrings
        items_optional.Push(s)

    canAddMore := (g_MedList.Length > 1)

    g_GUIAction          := ""
    g_GUIDropCount       := Max(holdCount, 1)
    g_GUISavedSelections := []
    g_GUIDDLs            := []
    g_SelectedMedIdx     := 0

    ; ---- GUI build/rebuild loop ----
    While true {

        InputGui := Gui("+AlwaysOnTop", "NDC OOS Ticket - Input")
        InputGui.SetFont("s11 Bold")
        InputGui.Add("Text", "w490 Center", "Creating Bulk NDC OOS Ticket")
        InputGui.SetFont("s10 Norm")
        InputGui.Add("Text", "w490", "")

        medWord := (g_MedList.Length > 1) ? "medication(s)" : "medication"
        If (g_GUIDropCount == 1 && holdCount == 0)
            InputGui.Add("Text", "w490", "No inventory hold detected. Select the bulk " . medWord . " to submit the ticket for:")
        Else If (g_GUIDropCount == 1 && holdCount == 1)
            InputGui.Add("Text", "w490", "Inventory hold detected. Bulk medication auto-selected - confirm and click OK:")
        Else
            InputGui.Add("Text", "w490", "Select the bulk medication(s) to submit the NDC OOS ticket for:")

        ; Reset and re-populate the live dropdown reference list for this build
        g_GUIDDLs := []
        Loop g_GUIDropCount {
            i       := A_Index
            isFirst := (i == 1)
            items   := isFirst ? items_required : items_optional

            ddl := InputGui.Add("DropDownList", "w490 vDDL" . i, items)
            g_GUIDDLs.Push(ddl)

            selText := ""
            If (i <= g_GUISavedSelections.Length && g_GUISavedSelections[i] != "")
                selText := g_GUISavedSelections[i]
            Else If (i <= holdMeds.Length) {
                hm      := g_MedList[holdMeds[i]]
                selText := hm.name . "  |  NDC: " . hm.ndc
            }

            If (selText != "") {
                For k, s in items {
                    If (s == selText) {
                        ddl.Choose(k)
                        Break
                    }
                }
            } Else {
                ddl.Choose(1)
            }
        }

        InputGui.Add("Text", "w490", "")

        BtnOK     := InputGui.Add("Button", "w100", "OK")
        BtnCancel := InputGui.Add("Button", "x+20 w100", "Cancel")
        If (canAddMore && g_GUIDropCount < g_MedList.Length)
            BtnAdd := InputGui.Add("Button", "x+20 w150", "Add Medication")

        BtnOK.OnEvent("Click", NDCInputOK)
        BtnCancel.OnEvent("Click", NDCInputCancel)
        If (canAddMore && g_GUIDropCount < g_MedList.Length)
            BtnAdd.OnEvent("Click", NDCInputAdd)
        InputGui.OnEvent("Close", GUICrash)

        g_GUIAction := ""
        PauseStart()
        InputGui.Show()
        While WinExist("NDC OOS Ticket - Input")
            Sleep(100)
        PauseEnd()

        If g_GUIAction == "add" {
            g_GUIDropCount++
            Continue
        }
        Break
    }

    ; ---- Post-GUI validation ----
    If (g_SelectedMedIdx == -1) {
        WipeAllData()
        Return
    }

    If (g_SelectedNDC == "") {
        MsgBox("No medication was selected. Aborting.", "Error", 16)
        WipeAllData()
        Return
    }

    If (g_SelectedRxNumber == "")
        MsgBox("WARNING: Could not extract an Rx number for the selected medication(s). You may need to fill it in manually.", "Extraction Warning", 48)

    NDCOOSPostGUI:
    g_ChangeNeeded := "NDC is OOS. No ETA, please reach out to customer to see how they would like to proceed."
    g_ReasonIndex  := 1

    ; -------------------------------------------------------------------------
    ; STEP 6: Refocus Chrome and open customer's admin details in a new tab
    ; -------------------------------------------------------------------------
    If WinExist("ahk_exe chrome.exe") {
        WinActivate("ahk_exe chrome.exe")
        If !WinWaitActive("ahk_exe chrome.exe", , 2) {
            MsgBox("STOPPING: Could not refocus Chrome. Aborting.", "Focus Error", 16)
            WipeAllData()
            Return
        }
    }

    OpenCustomerInNewTab(g_PatientName)
    g_PatientName := ""

    ; -------------------------------------------------------------------------
    ; STEP 7: Wait for Customer Details tab to load
    ; -------------------------------------------------------------------------
    DetailsText := ""

    If !WaitForDetailsPage(&DetailsText) {
        MsgBox("STOPPING: The customer details tab took too long to load.", "Timeout", 16)
        WipeAllData()
        Return
    }

    ; -------------------------------------------------------------------------
    ; STEP 7b: Capture the Admin URL
    ; -------------------------------------------------------------------------
    Send("!d")
    Sleep(100)
    A_Clipboard := ""
    Send("^c")
    Sleep(50)
    If ClipWait(1) {
        g_AdminURL  := A_Clipboard
        A_Clipboard := ""
    }
    Send("{Esc}")
    Sleep(50)

    ; -------------------------------------------------------------------------
    ; STEP 8: Extract Batch type -> determine Amazon or non-Amazon
    ; -------------------------------------------------------------------------
    BatchVal := ""
    If RegExMatch(DetailsText, "i)Batch[\s:]+([A-Za-z0-9_]+)", &BatchMatch)
        BatchVal := Trim(BatchMatch[1])

    If (BatchVal == "") {
        MsgBox("STOPPING: Could not locate the 'Batch' type on the customer details page.", "Extraction Failed", 16)
        DetailsText := ScrubString(DetailsText)
        DetailsText := ""
        WipeAllData()
        Return
    }

    If (BatchVal == "AMAZON") {
        g_IsAmazon := true
        If RegExMatch(DetailsText, "i)Person ID[\s:]+([^\r\n]+)", &IDMatch)
            g_PersonID := Trim(IDMatch[1])
        DetailsText := ScrubString(DetailsText)
        DetailsText := ""

        If (g_PersonID == "") {
            MsgBox("STOPPING: Could not locate the Person ID on the customer details page.", "Extraction Failed", 16)
            WipeAllData()
            Return
        }
    } Else {
        g_IsAmazon  := false
        DetailsText := ScrubString(DetailsText)
        DetailsText := ""

        CleanURL  := g_AdminURL
        QPos      := InStr(CleanURL, "?")
        If QPos
            CleanURL := SubStr(CleanURL, 1, QPos - 1)
        HPos := InStr(CleanURL, "#")
        If HPos
            CleanURL := SubStr(CleanURL, 1, HPos - 1)
        LastSlash := InStr(CleanURL, "/", false, -1)
        If LastSlash > 0
            g_PillPackProfileID := Trim(SubStr(CleanURL, LastSlash + 1))
        CleanURL := ""

        If (g_PillPackProfileID == "" || g_AdminURL == "") {
            MsgBox("STOPPING: Could not extract the PillPack Profile ID from the admin URL.`n`nAdmin URL captured: " . (g_AdminURL == "" ? "(none)" : "yes"), "Extraction Failed", 16)
            WipeAllData()
            Return
        }
    }
    BatchVal := ""
    ; AdminURL no longer needed after Profile ID extraction
    g_AdminURL := ""

    g_BlueprintName := g_IsAmazon ? "NDC Out of Stock" : "Bulk Inventory Issue"

    ; =========================================================================
    ; STEP 9: Open blueprint directly via URL
    ; =========================================================================
    DirectURL := g_IsAmazon
        ? "https://cs.wolfgang.a2z.com/tachyon/PHARMACY/createNewIssue/23D85000-0170-4A6F-834B-F9CDDDF2A8BB"
        : "https://cs.wolfgang.a2z.com/tachyon/PILLPACK/createNewIssue/360098504593"
    ExpectedTask := g_IsAmazon ? "Task: Red Bulk Shipment Change" : "Task: Bulk Inventory Issue"

    Send("^t")
    Sleep(200)
    A_Clipboard := DirectURL
    Sleep(50)
    Send("^v")
    Sleep(100)
    Send("{Enter}")
    A_Clipboard := ""
    DirectURL := ""

    ; -------------------------------------------------------------------------
    ; STEP 10: Wait for ticket form to load
    ; -------------------------------------------------------------------------
    Loaded := false
    Loop 30 {
        Sleep(A_Index = 1 ? 1500 : 500)
        If CheckChromeTitleForLogin("Tachyon")
            Continue
        A_Clipboard := ""
        Send("^a")
        Sleep(50)
        Send("^c")
        If ClipWait(0.5) {
            PageText := A_Clipboard
            A_Clipboard := ""
            If InStr(PageText, ExpectedTask) {
                PageText := ""
                ClearSelection()
                Loaded := true
                break
            }
            If InStr(PageText, "Sign In") || InStr(PageText, "Username") || InStr(PageText, "Verify with your password") || InStr(PageText, "Powered by Okta") {
                PageText := ""
                PauseStart()
                MsgBox("Sign-in required!`n`n1. Please log in manually.`n2. Click OK on this popup ONLY AFTER you see the ticket form.", "Login Paused", 64+262144)
                PauseEnd()
                Sleep(300)
                WinActivate("ahk_exe chrome.exe")
                Sleep(300)
            }
            PageText := ""
            ClearSelection()
        } Else {
            ClearSelection()
        }
    }
    If (!Loaded) {
        MsgBox("STOPPING: Timed out waiting for the " . g_BlueprintName . " ticket form to load.", "Timeout", 16)
        WipeAllData()
        Return
    }

    ; -------------------------------------------------------------------------
    ; STEP 15: Navigate to Person ID / PillPack Profile ID field and type
    ; -------------------------------------------------------------------------
    IDFieldLabel := g_IsAmazon ? "Person ID" : "PillPack Profile ID"
    LocalCustID := g_IsAmazon ? g_PersonID : g_PillPackProfileID

    If !TypeAndVerifyProfileID(IDFieldLabel, LocalCustID) {
        LocalCustID := ""
        MsgBox("Customer ID paste verification FAILED.`n`nAborting and wiping all data.", "Field Verify Failed", 16)
        WipeAllData()
        Return
    }
    LocalCustID := ""

    ; -------------------------------------------------------------------------
    ; STEP 16: Select customer name from profile lookup results
    ; -------------------------------------------------------------------------
    If !SelectCustomerName(g_IsAmazon) {
        MsgBox("WARNING: Could not confirm the customer name was selected after all retries.`n`nVerify the 'Customer first name' field is populated before submitting the ticket.", "Customer Selection Warning", 48)
        ReactivateChrome()
    }

    Sleep(300)

    ; =========================================================================
    ; STEP 17: Fill form fields
    ; =========================================================================
    If (g_IsAmazon) {

        ; STEP 17-A0: "Priority" -> High if due today or earlier
        SetPriorityIfDueToday(g_ShipDate)

        ; STEP 17-A1: "Due By" (Ship Date)
        FindAndType("Due", g_ShipDate)
        g_ShipDate := ""

        ; STEP 17-A2: "FFU - Reason for Change" dropdown (index 1 = first option)
        FindAndSelectDropdown("reason", 1)

        ; STEP 17-A2a: "Shipment ID" field
        FindAndPasteField("Shipment ID", g_GroupID)
        g_GroupID := ""

        ; STEP 17-A2b: "NDC11" field
        FindAndPasteField("NDC11", g_SelectedNDC)
        g_SelectedNDC := ""

        ; STEP 17-A3: "FFU - Original Promise Date" calendar
        FindAndSetCalendar("FFU - Original Promise Date", g_ArrivalParsed)
        g_ArrivalParsed := ""
        g_ArrivalDate   := ""

        ; STEP 17-A4: Medication Name
        FindAndPaste("Medication Name", g_SelectedMedName)
        g_SelectedMedName := ""

        ; STEP 17-A5: Rx#
        DownAndPaste(g_SelectedRxNumber)
        g_SelectedRxNumber := ""

        ; STEP 17-A6: Change Needed
        FindAndPaste("Change Needed", g_ChangeNeeded)
        g_ChangeNeeded := ""

        ; STEP 17-A7: Shipment link and/or Order ID
        FindAndPaste("Shipment link and/or Order ID", g_ShipmentURL)
        g_ShipmentURL := ""

    } Else {

        ; g_SelectedRxNumber is not used in the non-Amazon path; clear it now
        g_SelectedRxNumber := ""
        ; g_ShipmentURL is not used in the non-Amazon form; clear it now
        g_ShipmentURL := ""

        ; STEP 17-B0: Compute Due By (arrival date for non-Amazon), then set Priority
        DueByVal := ""
        If (Type(g_ArrivalParsed) == "Object" && g_ArrivalParsed.formatted != "")
            DueByVal := g_ArrivalParsed.formatted
        If (DueByVal == "")
            DueByVal := g_ShipDate
        SetPriorityIfDueToday(DueByVal)
        FindAndType("Due By", DueByVal)
        DueByVal := ""
        g_ShipDate := ""

        ; STEP 17-B2: "Due Date" calendar (Arrival Date)
        FindAndSetCalendar("Due Date", g_ArrivalParsed)
        g_ArrivalParsed := ""
        g_ArrivalDate   := ""

        ; STEP 17-B3: NDC field
        FindAndPasteField("NDC", g_SelectedNDC)

        ; STEP 17-B4: Shipment ID field
        FindAndPasteField("Shipment ID", g_GroupID)
        g_GroupID := ""

        ; STEP 17-B5: "NDC:Bulk Medication:" template line
        FindAndPaste("NDC:Bulk Medication", g_SelectedNDC . "  |  " . g_SelectedMedName)
        g_SelectedNDC     := ""
        g_SelectedMedName := ""

        ; STEP 17-B6: "Change Needed" text
        Send("^f")
        Sleep(50)
        SendText("What date are we expecting")
        Sleep(50)
        Send("{Esc}")
        Sleep(50)
        Send("{End}")
        Sleep(100)
        Send("{Enter}")
        Sleep(200)
        A_Clipboard := g_ChangeNeeded
        MarkClipboardExcluded()
        Sleep(50)
        Send("^v")
        Sleep(50)
        A_Clipboard := ""
        g_ChangeNeeded := ""
        Sleep(200)

        g_ShipmentURL := ""
    }

    ; -------------------------------------------------------------------------
    ; STEP 18: Complete
    ; -------------------------------------------------------------------------
    WipeAllData()
    MsgBox("Ticket form populated.`n`n*** VERIFY ALL INFORMATION IS CORRECT BEFORE SUBMITTING ***`n`nReview every field (customer, dates, NDC, medication, Rx#, etc.) and confirm the blueprint matches the case.", "Complete - Verify Before Submit", 64+262144)
}
