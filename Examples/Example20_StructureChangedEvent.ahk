#Requires AutoHotkey v2
;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

Run "explore C:\"
CDriveName := DriveGetLabel("C:") " (C:)"
WinWaitActive(CDriveName,,1)
explorerEl := UIA.ElementFromHandle("A")
if !explorerEl {
	MsgBox "Drive C: element not found! Exiting app..."
	ExitApp
}
MsgBox "Press OK to create a new EventHandler for the StructureChanged event.`nTo test this, interact with the Explorer window, and a tooltip should pop up.`n`nTo exit the script, press F5."
handler := UIA.CreateStructureChangedEventHandler(StructureChangedEventHandler)
UIA.AddStructureChangedEventHandler(handler, explorerEl)
OnExit(ExitFunc) ; Set up an OnExit call to clean up the handler when exiting the script
return

StructureChangedEventHandler(sender, changeType, runtimeId) {
    ToolTip "Sender: " sender.Dump() 
        . "`nChange type: " changeType
        . "`nRuntime Id: " UIA.RuntimeIdToString(runtimeId)
	SetTimer ToolTip, -3000
}

ExitFunc(*) {
	UIA.RemoveAllEventHandlers()
}

F5::ExitApp
