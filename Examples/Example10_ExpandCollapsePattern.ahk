#Requires AutoHotkey v2
;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

Run "explore C:\"
CDriveName := DriveGetLabel("C:") " (C:)"
WinWaitActive(CDriveName,,1)
WinMove(100, 200, 1000, , "A")
explorerEl := UIA.ElementFromHandle("A")
if !explorerEl {
	MsgBox "Drive C: element not found! Exiting app..."
	ExitApp
}
CDriveEl := explorerEl.FindElement({Type:"TreeItem", Name:CDriveName, matchmode:"Substring"})

Sleep 500
MsgBox "ExpandCollapsePattern properties: "
	. "`nCurrentExpandCollapseState: " (state := CDriveEl.ExpandCollapseState) " (" UIA.ExpandCollapseState[state] ")"
Sleep 500
MsgBox "Press OK to expand drive C: element"
CDriveEl.Expand()
Sleep 500
MsgBox "Press OK to collapse drive C: element"
CDriveEl.Collapse()

ExitApp
