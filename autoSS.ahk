#Requires AutoHotkey v2.0

; This script captures a screenshot of the active window and pastes it into Microsoft Word when Print Screen is pressed.

; Variable to store the previous window's ID
previousWindow := ""

^PrintScreen::
{
    global previousWindow  ; Declare as global to access outside the hotkey

    ; Store the ID of the currently active window
    previousWindow := WinActive("A")
    ; Take a screenshot of the active window
    Send "{Alt Down}{PrintScreen}{Alt Up}"  ; Alt + PrintScreen captures the active window
    ClipWait(1)  ; Wait until the clipboard has content

   id := WinGetList("ahk_class OpusApp") ; Class name for Microsoft Word
    for thisID in id {
        thisTitle := WinGetTitle(thisID)
        
        ; Check if this window is a document (not the main Word window)
        if (InStr(thisTitle, "Document") > 0) {
            latestWordWindow := thisID
            break ; Exit loop after finding the first document
        }
    }

    ; If a Word document is found, activate it and paste
    if (latestWordWindow != "") {
        WinActivate(latestWordWindow) ; Activate the latest Word document
        Send("^v") ; Paste the screenshot into the document
        Send("{Enter}") ; Optional: add a new line after pasting
    } else {
        MsgBox("No open Word documents found.")
    }
	; Return to the previous window
    if (previousWindow)
    {
        WinActivate(previousWindow)  ; Activate the previously active window
    }
}