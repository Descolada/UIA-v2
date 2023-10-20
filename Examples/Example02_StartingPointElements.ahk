#Requires AutoHotkey v2
;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

/*
    One of the best ways of accessing UI elements (buttons, text etc) is by first getting the window
    element for the starting element with ElementFromHandle. For that we can use the same notation 
    as for any other AHK function.
*/
Run "notepad.exe"
WinWaitActive "ahk_exe notepad.exe"
npEl := UIA.ElementFromHandle("ahk_exe notepad.exe") ; Get the element for the Notepad window
MsgBox "Notepad window element with all descendants: `n`n" npEl.DumpAll() ; Display all the sub-elements for the Notepad window. 

/*
    ElementFromHandle doesn't only access windows, but can use any handle: control handles can be
    used to get a part of the window. This is sometimes necessary when for some reason UIAutomation
    hasn't been implemented properly and not all elements are displayed in the UIA tree.
*/
try editHandle := ControlGetHwnd("Edit1", "ahk_exe notepad.exe")
catch
    editHandle := ControlGetHwnd("RichEditD2DPT1")

editEl := UIA.ElementFromHandle(editHandle)
MsgBox "Edit control element with all descendants: `n`n" editEl.DumpAll()

/*
    A special case for ElementFromHandle using a control is ElementFromChromium, which gets the
    element for the renderer control. This special case exists because Chromium-based (browser-based)
    applications frequently have the problem of not being UI-accessible from the main window.

    For this example you need to have Chrome open.
*/
if WinExist("ahk_exe chrome.exe") {
    WinActivate("ahk_exe chrome.exe")
    WinWaitActive("ahk_exe chrome.exe")
}
chromiumEl := UIA.ElementFromChromium("ahk_exe chrome.exe")
MsgBox "Chromium control element without descendants: `n`n" chromiumEl.Dump()

/*
    Elements can also be gotten from any point on the screen with ElementFromPoint. This usually won't
    return the window element under the cursor, but a sub-element of the window that the mouse
    is hovering over. Omitting x and y arguments from ElementFromPoint will use the current mouse position.
*/

MsgBox "Starting ElementFromPoint example. `nPress OK to start, F5 to exit."
Loop 
    try ToolTip UIA.ElementFromPoint().Dump()

ExitApp
F5::ExitApp