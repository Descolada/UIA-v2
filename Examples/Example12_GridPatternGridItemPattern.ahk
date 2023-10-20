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
MsgBox "GridPattern properties: "
	. "`nCurrentRowCount: " listEl.RowCount
	. "`nCurrentColumnCount: " listEl.ColumnCount
Sleep 500
MsgBox "Getting grid item from row 4, column 1 (0-based indexing)"
editEl := listEl.GetItem(3,0).Highlight()
MsgBox "Got this element: `n" editEl.Dump()

MsgBox "GridItemPattern properties: "
	. "`nCurrentRow: " editEl.Row
	. "`nCurrentColumn: " editEl.Column
	. "`nCurrentRowSpan: " editEl.RowSpan
	. "`nCurrentColumnSpan: " editEl.ColumnSpan
	; editEl.ContainingGrid should return listEl

ExitApp
