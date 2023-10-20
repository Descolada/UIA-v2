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
treeEl := explorerEl.FindElement({Type:"Tree"})

MsgBox "For this example, make sure that the folder tree on the left side in File Explorer has some scrollable elements (make the window small enough)."
Sleep 500
MsgBox "ScrollPattern properties: "
	. "`nCurrentHorizontalScrollPercent: " treeEl.HorizontalScrollPercent
	. "`nCurrentVerticalScrollPercent: " treeEl.VerticalScrollPercent
	. "`nCurrentHorizontalViewSize: " treeEl.HorizontalViewSize
	. "`nCurrentHorizontallyScrollable: " treeEl.HorizontallyScrollable
	. "`nCurrentVerticallyScrollable: " treeEl.VerticallyScrollable
Sleep 50
MsgBox "Press OK to set scroll percent to 50% vertically and 0% horizontally."
treeEl.SetScrollPercent(50) ; Equivalent to treeEl.VerticalScrollPercent := 50
Sleep 500
MsgBox "Press OK to scroll a Page Up equivalent upwards vertically."
treeEl.Scroll(UIA.ScrollAmount.LargeDecrement) ; LargeDecrement is equivalent to pressing the PAGE UP key or clicking on a blank part of a scroll bar. SmallDecrement is equivalent to pressing an arrow key or clicking the arrow button on a scroll bar.

Sleep 500
MsgBox "Press OK to scroll drive C: into view."
CDriveEl := explorerEl.FindElement({Type:"TreeItem", Name:CDriveName, matchmode:"Substring"})
if !CDriveEl {
	MsgBox "C: drive element not found! Exiting app..."
	ExitApp
}
CDriveEl.ScrollIntoView()

ExitApp