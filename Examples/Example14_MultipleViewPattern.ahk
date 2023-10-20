#Requires AutoHotkey v2
;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

Run "explore C:\Windows"
WinWaitActive("Windows",,1)
WinMove(100, 200, 1000, , "A")
explorerEl := UIA.ElementFromHandle("A")
if !explorerEl {
	MsgBox "Drive C: element not found! Exiting app..."
	ExitApp
}
listEl := explorerEl.FindElement({Type:"List"})

MsgBox "MultipleView properties: "
	. "`nCurrentCurrentView: " (currentView := listEl.CurrentView)

supportedViews := listEl.GetSupportedViews()
viewNames := ""
for view in supportedViews {
	viewNames .= listEl.GetViewName(view) " (" view ")`n"
}
MsgBox "This MultipleView supported views:`n" viewNames
MsgBox "Press OK to set MultipleView to view 4."
listEl.SetCurrentView(4)

Sleep 500
MsgBox "Press OK to reset back to view " currentView "."
listEl.SetCurrentView(currentView)

ExitApp
