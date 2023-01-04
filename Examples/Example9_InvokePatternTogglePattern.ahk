;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

Run "explore C:\"
WinWaitActive DriveGetLabel("C:") " (C:)",, 1
explorerEl := UIA.ElementFromHandle("A")
if !explorerEl {
	MsgBox "Drive C: element not found! Exiting app..."
	ExitApp
}
fileEl := explorerEl.FindFirst({Type:"Button", Name:"File tab"})
MsgBox "Invoke pattern doesn't have any properties. Press OK to call Invoke on the `"File`" button..."
fileEl.Invoke()

Sleep 1000
MsgBox "Press OK to navigate to the View tab to test TogglePattern..." ; Not part of this demonstration
explorerEl.FindFirst({Type:"TabItem", Name:"View"}).Select() ; Not part of this demonstration

hiddenItemsCB := explorerEl.FindFirst({Type:"CheckBox", Name:"Hidden items"})
Sleep 500
MsgBox "TogglePattern properties for `"Hidden items`" checkbox: "
	. "`nCurrentToggleState: " hiddenItemsCB.ToggleState

MsgBox "Press OK to toggle"
hiddenItemsCB.Toggle()
Sleep 500
MsgBox "Press OK to toggle again"
hiddenItemsCB.TogglePattern.Toggle() ; This way Toggle() will be called specifically from TogglePattern

; hiddenItemsCB.ToggleState := 1 ; ToggleState can also be used to set the state

ExitApp
