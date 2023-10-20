#Requires AutoHotkey v2
;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

Run "notepad.exe"
WinWaitActive "ahk_exe notepad.exe"
; Get the element for the Notepad window
npEl := UIA.ElementFromHandle("ahk_exe notepad.exe") 
; Display all the sub-elements for the Notepad window. Press OK to continue
MsgBox npEl.DumpAll()
; Find the first Document or Edit control (in Notepad there is only one). In older versions of Windows this was an Edit type control, in newer ones it's Document.
try documentEl := npEl.FindElement([{Type:"Document"}, {Type:"Edit"}]) 
catch {
    ; Windows 11 has broken Notepad so that the Document element isn't findable; instead get it by the ClassNN
    Sleep 40
    documentEl := UIA.ElementFromHandle(ControlGetHwnd("RichEditD2DPT1"))
}
; Set the value of the document control, same as documentEl.SetValue("Lorem ipsum")
documentEl.Value := "Lorem ipsum" 
MsgBox "Press OK to test saving." ; Wait for the user to press OK
; Find the "File" menu item
fileEl := npEl.FindElement({Type:"MenuItem", Name:"File"})
fileEl.Highlight()
fileEl.Click()
; The last three lines could be combined into:
; fileEl.FindElement({Type:"MenuItem", Name:"File"}).Highlight(2000).Click()
saveEl := npEl.WaitElement({Name:"Save", mm:2}) ; Wait for the "Save" menu item to exist
saveEl.Highlight()
; And now click Save
saveEl.Click() 
ExitApp

