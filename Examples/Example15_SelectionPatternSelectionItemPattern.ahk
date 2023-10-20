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
listEl := explorerEl.FindElement({Type:"List"})

MsgBox "SelectionPattern properties: "
	. "`nCurrentCanSelectMultiple: " listEl.CanSelectMultiple
	. "`nCurrentIsSelectionRequired: " listEl.IsSelectionRequired

currentSelectionEls := listEl.GetCurrentSelection()
currentSelections := ""
for index, selection in currentSelectionEls
	currentSelections .= index ": " selection.Dump() "`n"

windowsListItem := explorerEl.FindElement({Name:"Windows", Type:"ListItem"})
MsgBox "ListItemPattern properties for Windows folder list item:"
	. "`nCurrentIsSelected: " windowsListItem.IsSelected
	. "`nCurrentSelectionContainer: " windowsListItem.SelectionContainer.Dump()

MsgBox "Press OK to select `"Windows`" folder list item."
windowsListItem.Select()
MsgBox "Press OK to add to selection `"Program Files`" folder list item."
explorerEl.FindElement({Name:"Program Files", Type:"ListItem"}).AddToSelection()
MsgBox "Press OK to remove selection from `"Windows`" folder list item."
windowsListItem.RemoveFromSelection()

ExitApp
