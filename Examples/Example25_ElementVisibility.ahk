#Requires AutoHotkey v2
;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

/*
    UIAutomation usually cannot "see" elements if they are not visible on the screen.
    Different programs handle this differently. Some just set Element.IsOffscreen property to 0,
    others don't display the element at all in the UIA tree.

    This example shows one way how an element (in this case the System32 folder) can be
    searched for by interacting with the window.

    // Credit for this example goes to user neogna2.
*/

SetTitleMatchMode 2
Run "explorer C:\Windows"
if !explorerHwnd := WinWaitActive("Windows",,1)
    ExitApp
; Decrease window height so that folder "System32" is out of view
WinMove( , , 1000, 800, explorerHwnd)
explorerEl := UIA.ElementFromHandle("A")
if !explorerEl
    ExitApp
listEl := explorerEl.FindElement({Type:"List"})
if !listItem := explorerEl.ElementExist({Name:"System32", Type:"ListItem"}) {
    if "OK" != MsgBox("Press OK to scroll until 'System32' is in view and then select it.")
        ExitApp
    Loop {
        if listItem := explorerEl.ElementExist({Name:"System32", Type:"ListItem"}) {
            listItem.AddToSelection()
            break
        }
        listEl.Scroll("LargeIncrement")
    } Until Round(listEl.VerticalScrollPercent) = 100
}
ExitApp

F5::ExitApp