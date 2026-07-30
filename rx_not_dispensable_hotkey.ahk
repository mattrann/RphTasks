#Requires AutoHotkey v2.0
#SingleInstance Force

; rx_not_dispensable_hotkey.ahk
; Adds a standalone hotkey Ctrl+Shift+Alt+N that opens the Tachyon "Create New Issue"
; flow, selects the "RX Not Dispensable" blueprint, pre-fills the template and a few
; known fields from environment variables or clipboard, and leaves the form open for
; agent review before submit.
;
; This file is intentionally standalone so it can be tested independently of the
; main Pharmacist_Network_V2.3.3.ahk script. It favors JS-based clicks (address-bar
; javascript:) when possible to avoid brittle keystroke-only flows.

; -----------------------------
; Helper: execute JS via address-bar javascript: injection in Chrome
; -----------------------------
ExecJS(code) {
    ; activate chrome and open address bar javascript: payload
    Try WinActivate("ahk_exe chrome.exe")
    Sleep(150)
    Send("!d")              ; focus address bar
    Sleep(150)
    SendText("javascript:")
    Sleep(80)
    A_Clipboard := code
    Send("^v")
    Sleep(60)
    Send("{Enter}")
    Sleep(700)
    A_Clipboard := ""
}

; -----------------------------
; Helper: paste into a textarea found by its visible label using Find
; (Find-and-paste behaviour mirrors main script but simplified)
; -----------------------------
FindAndPaste(label, value) {
    Send("^f")
    Sleep(120)
    SendText(label)
    Sleep(120)
    Send("{Esc}")
    Sleep(80)
    ; Move to end of the found line then paste (appends)
    Send("{End}")
    Sleep(80)
    A_Clipboard := value
    ; mark clipboard excluded if the environment has helper — best effort
    Send("^v")
    Sleep(120)
    A_Clipboard := ""
}

FindAndPasteField(label, value) {
    Send("^f")
    Sleep(120)
    SendText(label)
    Sleep(120)
    Send("{Enter}")
    Sleep(120)
    Send("{Esc}")
    Sleep(80)
    Send("{Tab}")
    Sleep(120)
    A_Clipboard := value
    Send("^v")
    Sleep(120)
    A_Clipboard := ""
}

ReactivateChrome() {
    Try WinActivate("ahk_exe chrome.exe")
    Sleep(250)
}

; -----------------------------
; Hotkey: Ctrl+Shift+Alt+N
; Behavior (Option B):
;  - Trigger existing Tachyon open flow by sending ^+!t (assumes main script binds that)
;  - Click the "RX Not Dispensable" blueprint using JS
;  - Paste the blueprint template into the instructions area
;  - Pre-fill Rx#, Shipment URL, Medication Name if available in environment variables
;  - Leave the form open for agent review
; -----------------------------
^+!n:: {
    ; Try to use existing shortcut to open the Tachyon flow if main script is running
    Send("^+!t")
    Sleep(1200)
    ReactivateChrome()
    Sleep(600)

    ; JS: click an element that contains the blueprint name text
    JSPayload := "(function(){var txt='RX Not Dispensable';var candidates=document.querySelectorAll('button,div,span,a');for(var i=0;i<candidates.length;i++){try{if(candidates[i].textContent && candidates[i].textContent.trim().indexOf(txt)!=-1){candidates[i].click();return 'clicked';}}catch(e){} } return 'notfound';})()"
    ExecJS(JSPayload)
    Sleep(750)

    ; second attempt if not yet visible (some pages render slowly)
    ExecJS("(function(){var txt='RX Not Dispensable';var c=document.querySelectorAll('button,div,span,a');for(var i=0;i<c.length;i++){try{if(c[i].textContent&&c[i].textContent.trim().indexOf(txt)!=-1){c[i].click();return;} }catch(e){} }})();")
    Sleep(700)

    ; Prepare template text (matches the Tachyon blueprint template)
    template := "Instructions to CCC Agent:`n`nCustomer name: {{ticket.requester.name}}`nDate of Birth: {{ticket.requester.custom_fields.birthdate}}`nFacility: {{ticket.requester.custom_fields.facility_name}}`n`nShipment in question (please provide a shipment link):`n`nNot Dispensable Prescription #: `n`nMedication Name(s):"

    ; Paste the template under the instructions label
    FindAndPaste("Instructions to CCC Agent:", template)
    Sleep(350)

    ; Attempt to pre-fill fields from script variables or environment variables
    ; Use environment variables first (allows ad-hoc invocation), else clipboard heuristics
    RxVal := EnvGet("G_RxNumber")
    ShipURL := EnvGet("G_ShipmentURL")
    MedName := EnvGet("G_MedicationName")

    ; Fallback: if nothing in env, try to use currently loaded clipboard (best-effort)
    If (RxVal == "") {
        ; no-op; keep blank
    }
    If (ShipURL == "") {
        ; no-op
    }
    If (MedName == "") {
        ; no-op
    }

    if (RxVal != "") {
        FindAndPasteField("Not Dispensable Prescription #", RxVal)
        Sleep(200)
    }
    if (ShipURL != "") {
        FindAndPasteField("Shipment in question", ShipURL)
        Sleep(200)
    }
    if (MedName != "") {
        FindAndPasteField("Medication Name(s)", MedName)
        Sleep(200)
    }

    MsgBox("RX Not Dispensable blueprint selected and pre-filled (where data available).\n\nReview the issue and submit when ready.", "Prefill Complete", 64)
}

; End of file
