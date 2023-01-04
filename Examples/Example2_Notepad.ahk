;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

Run "notepad.exe"
WinWaitActive "ahk_exe notepad.exe"
npEl := UIA.ElementFromHandle("ahk_exe notepad.exe") ; Get the element for the Notepad window
documentEl := npEl.FindFirst([{Type:"Document"}, {Type:"Edit"}]) ; Find the first Document or Edit control (in Notepad there is only one). 
documentEl.Highlight(2000) ; Highlight the found element
documentEl.Value := "Lorem ipsum" ; Set the value for the document control. 

; This could also be chained together as: 
; UIA.ElementFromHandle("ahk_exe notepad.exe").FindFirst([{Type:"Document"}, {Type:"Edit"}]).Highlight(2000).Value := "Lorem ipsum"
ExitApp