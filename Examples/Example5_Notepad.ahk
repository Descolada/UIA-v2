;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

Run "notepad.exe"
WinWaitActive "ahk_exe notepad.exe"
npEl := UIA.ElementFromHandle("ahk_exe notepad.exe") ; Get the element for the Notepad window
MsgBox npEl.DumpAll() ; Display all the sub-elements for the Notepad window. Press OK to continue
documentEl := npEl.FindFirst({Type:"Document"}) ; Find the first Document control (in Notepad there is only one). This assumes the user is running a relatively recent Windows and UIA interface version 2+ is available. In UIA interface v1 this control was Edit, so an alternative option instead of "Document" would be "UIA.__Version > 1 ? "Document" : "Edit""
documentEl.Value := "Lorem ipsum" ; Set the value of the document control, same as documentEl.SetValue("Lorem ipsum")
MsgBox "Press OK to test saving." ; Wait for the user to press OK
fileEl := npEl.FindFirst({Type:"MenuItem", Name:"File"}) ; Find the "File" menu item
fileEl.Highlight(2000)
fileEl.Click()
; The last three lines could be combined into:
; fileEl.FindFirst({Type:"MenuItem", Name:"File"}).Highlight(2000).Click()
saveEl := npEl.WaitExist({Name:"Save", mm:2}) ; Wait for the "Save" menu item to exist
saveEl.Highlight(2000)
saveEl.Click() ; And now click Save
ExitApp

