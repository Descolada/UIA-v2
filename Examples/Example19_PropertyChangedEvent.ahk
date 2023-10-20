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
MsgBox "Press OK to create a new EventHandler for the PropertyChanged event (property UIA_NamePropertyId).`nTo test this, click on any file/folder, and a tooltip should pop up.`n`nTo exit the script, press F5."
handler := UIA.CreatePropertyChangedEventHandler(PropertyChangedEventHandler)
UIA.AddPropertyChangedEventHandler(handler, explorerEl, UIA.Property.Name) ; Multiple properties can be specified in an array
OnExit(ExitFunc) ; Set up an OnExit call to clean up the handler when exiting the script
return

PropertyChangedEventHandler(sender, propertyId, newValue) {
    ToolTip "Sender: " sender.Dump() 
        . "`nPropertyId: " propertyId
        . "`nNew value: " newValue
	SetTimer ToolTip, -3000
}

ExitFunc(*) {
	UIA.RemoveAllEventHandlers()
}

F5::ExitApp
