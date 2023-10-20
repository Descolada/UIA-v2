#Requires AutoHotkey v2
;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

Run "explore C:\"
CDriveName := DriveGetLabel("C:") " (C:)"
WinWaitActive(CDriveName,,1)
WinMove(200, 100, 1000, 800, CDriveName)
explorerEl := UIA.ElementFromHandle(CDriveName)
if !explorerEl {
	MsgBox "Drive C: element not found! Exiting app..."
	ExitApp
}
; WindowPattern
Sleep 500
MsgBox "WindowPattern properties: "
	. "`nCurrentCanMaximize: " explorerEl.CanMaximize
	. "`nCurrentCanMinimize: " explorerEl.CanMinimize
	. "`nCurrentIsModal: " explorerEl.IsModal
	. "`nCurrentIsTopmost: " explorerEl.IsTopmost
	. "`nCurrentWindowVisualState: " (visualState := explorerEl.WindowVisualState) " (" UIA.WindowVisualState[visualState] ")"
	. "`nCurrentWindowInteractionState: " (interactionState := explorerEl.WindowInteractionState) " (" UIA.WindowInteractionState[interactionState] ")"
Sleep 50
MsgBox "Press OK to try minimizing"
explorerEl.SetWindowVisualState(UIA.WindowVisualState.Minimized)

Sleep 500
MsgBox "Press OK to bring window back to normal"
explorerEl.SetWindowVisualState(UIA.WindowVisualState.Normal)

; TransformPattern
Sleep 500
MsgBox "TransformPattern properties: "
	. "`nCurrentCanMove: " explorerEl.CanMove
	. "`nCurrentCanResize: " explorerEl.CanResize
	. "`nCurrentCanRotate: " explorerEl.CanRotate

MsgBox "Press OK to move to coordinates x100 y200"
explorerEl.Move(100,200)

Sleep 500
MsgBox "Press OK to resize to w600 h400"
explorerEl.Resize(600,400)

Sleep 500
MsgBox "Press OK to close window"
WinMove(100, 200, 1000, 800, CDriveName)
explorerEl.Close()
ExitApp
