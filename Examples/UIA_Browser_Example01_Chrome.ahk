#Requires AutoHotkey v2
;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\UIA.ahk
;#include <UIA_Browser> ; Uncomment if you have moved UIA_Browser.ahk to your main Lib folder
#include ..\Lib\UIA_Browser.ahk

/**
 * This example starts a new Chrome window and then gets the Document element content into the Clipboard.
 */

; Run in Incognito mode to avoid any extensions interfering.
Run "chrome.exe -incognito" 
WinWaitActive "ahk_exe chrome.exe"
Sleep 500
; Initialize UIA_Browser, use Last Found Window (returned by WinWaitActive)
cUIA := UIA_Browser() 
A_Clipboard := ""
; Get the current document element (this excludes the URL bar, navigation buttons etc) 
; and dump all the information about it in the clipboard. 
; Use Ctrl+V to paste it somewhere, such as in Notepad.
A_Clipboard := cUIA.GetCurrentDocumentElement().DumpAll() 
ClipWait 1
if A_Clipboard
	MsgBox "Page information successfully dumped. Use Ctrl+V to paste the info somewhere, such as in Notepad."
else
	MsgBox "Something went wrong and nothing was dumped in the clipboard!"
ExitApp
