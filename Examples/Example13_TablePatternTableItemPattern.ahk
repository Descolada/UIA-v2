#Requires AutoHotkey v2
;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

Run "explore C:\Windows"
WinWaitActive("Windows",,1)
WinMove(100, 200, 1000, , "A")
explorerEl := UIA.ElementFromHandle("A")
if !explorerEl {
	MsgBox "Windows folder window not found! Exiting app..."
	ExitApp
}
Sleep 500
listEl := explorerEl.FindElement({Type:"List"})

MsgBox "TablePattern properties: "
	. "`nCurrentRowOrColumnMajor: " listEl.RowOrColumnMajor

rowHeaders := listEl.GetRowHeaders()
rowHeadersDump := ""
for header in rowHeaders
	rowHeadersDump .= header.Dump() "`n"
MsgBox "TablePattern elements from GetRowHeaders:`n" rowHeadersDump ; Should be empty, there aren't any row headers
columnHeaders := listEl.GetColumnHeaders()
columnHeadersDump := ""
for header in columnHeaders
	columnHeadersDump .= header.Dump() "`n"
MsgBox "TablePattern elements from GetColumnHeaders:`n" columnHeadersDump

editEl := listEl.GetItem(3,0) ; To test the TableItem pattern, we need to get an element supporting that using Grid pattern...
rowHeaderItems := editEl.GetRowHeaderItems()
rowHeaderItemsDump := ""
for headerItem in rowHeaderItems
	rowHeaderItemsDump .= headerItem.Dump() "`n"
MsgBox "TableItemPattern elements from GetRowHeaderItems:`n" rowHeaderItemsDump ; Should be empty, there aren't any row headers
columnHeaderItems := editEl.GetColumnHeaderItems()
columnHeaderItemsDump := ""
for headerItem in columnHeaderItems
	columnHeaderItemsDump .= headerItem.Dump() "`n"
MsgBox "TableItemPattern elements from GetCurrentColumnHeaderItems:`n" columnHeaderItemsDump
ExitApp
