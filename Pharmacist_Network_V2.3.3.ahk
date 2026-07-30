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

; =====================================================================; ... (file continues unchanged) ...

; =====================================================================; New hotkey: Ctrl+Shift+Alt+N — RX Not Dispensable (Option B)
; Opens the Tachyon Create Issue flow, selects the "RX Not Dispensable"
; blueprint, pre-fills the template and available fields from script
; variables, and leaves the form open for agent review before submit.
; =====================================================================
^+!n:: {
    global g_RxNumber, g_ShipmentURL, g_MedicationName, g_PatientName, g_DOB

    ; Open the standard Tachyon ticket flow (existing hotkey handles navigation)
    Send("^+!t")
    Sleep(1200)
    ReactivateChrome()
    Sleep(600)

    ; JS payload: find any element containing the blueprint text and click it    JSPayload := "(function(){var txt='RX Not Dispensable';var els=document.querySelectorAll('button,div,span,a');for(var i=0;i<els.length;i++){try{if(els[i].textContent && els[i].textContent.trim().indexOf(txt)!=-1){els[i].click();return;} }catch(e){} }})();"

    ; Try to click blueprint and verify expected task text appears    ClickBlueprintAndVerify(JSPayload, "RX Not Dispensable", false)
    Sleep(600)

    ; Pre-fill the Instructions box with the blueprint template    template := "Instructions to CCC Agent:`n`nCustomer name: {{ticket.requester.name}}`nDate of Birth: {{ticket.requester.custom_fields.birthdate}}`nFacility: {{ticket.requester.custom_fields.facility_name}}`n`nShipment in question (please provide a shipment link):`n`nNot Dispensable Prescription #:`n`nMedication Name(s):"
    FindAndPaste("Instructions to CCC Agent:", template)
    Sleep(300)

    ; Pre-fill specific fields if script variables are available    if (g_RxNumber != "") {
        FindAndPasteField("Not Dispensable Prescription #", g_RxNumber)
        Sleep(150)
    }
    if (g_ShipmentURL != "") {
        FindAndPasteField("Shipment in question", g_ShipmentURL)
        Sleep(150)
    }
    if (g_MedicationName != "") {
        FindAndPasteField("Medication Name(s)", g_MedicationName)
        Sleep(150)
    }

    MsgBox("RX Not Dispensable blueprint selected and fields pre-filled. Review the ticket and submit when ready.", "RX Not Dispensable - Prefill Complete", 64)
}
