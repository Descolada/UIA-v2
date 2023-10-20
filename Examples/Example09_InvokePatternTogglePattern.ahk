#Requires AutoHotkey v2
;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

if VerCompare(A_OSVersion, ">=10.0.22000") {
    MsgBox "This example works only in Windows 10. Press OK to Exit."
    ExitApp
}

Run "explore C:\"
WinWaitActive DriveGetLabel("C:") " (C:)",, 1
WinMove(100, 200, 1000, , "A")
explorerEl := UIA.ElementFromHandle("A")
if !explorerEl {
	MsgBox "Drive C: element not found! Exiting app..."
	ExitApp
}
fileEl := explorerEl.FindElement({Type:"Button", Name:"File tab"})
MsgBox "Invoke pattern doesn't have any properties. Press OK to call Invoke on the `"File`" button..."
fileEl.Invoke()

Sleep 1000
MsgBox "Press OK to navigate to the View tab to test TogglePattern..." ; Not part of this demonstration
if !explorerEl.FindElement({Name:"Lower ribbon", cs:0}) ; Not part of this demonstration
	try explorerEl.FindElement({T:0,N:"Minimize the Ribbon"}).Invoke() ; Not part of this demonstration
explorerEl.FindElement({Type:"TabItem", Name:"View"}).Select() ; Not part of this demonstration

hiddenItemsCB := explorerEl.FindElement({Type:"CheckBox", Name:"Hidden items"})
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
