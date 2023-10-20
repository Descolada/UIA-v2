#Requires AutoHotkey v2
;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

cacheRequest := UIA.CreateCacheRequest(["Type", "Name", "Value"],, "Subtree")
/*
    ; Instead we could also define a cacherequest like this:
    cacheRequest := UIA.CreateCacheRequest()
    ; Set TreeScope to include the starting element and all descendants as well
    cacheRequest.TreeScope := 5 
    ; Add some properties to be cached
    cacheRequest.AddProperty("Type") 
    cacheRequest.AddProperty("Name")
    cacheRequest.AddProperty("Value")

    ; Or like this:
    cacheRequest := UIA.CreateCacheRequest({properties:["Type", "Name", "Value"], scope:"Subtree"})
*/

Run "notepad.exe"
WinWaitActive "ahk_exe notepad.exe"

MsgBox("Type something in Notepad: note that the document content won't change in the tooltip.`nPress F5 to refresh the cache - then the document content will also update in the tooltip.")

; Get element and also build the cache
npEl:= UIA.ElementFromHandle("ahk_exe notepad.exe", cacheRequest)
docEl := npEl.FindElement([{Type:"Document"}, {Type:"Edit"}],,,,,cacheRequest)
; We now have a cached "snapshot" of the window from which we can access our desired elements faster.
Loop {
    ToolTip "Cached window name: " npEl.CachedName "`nCached document content: " docEl.CachedValue
}

F5::
{
    global npEl, docEl, cacheRequest
    npEl := npEl.BuildUpdatedCache(cacheRequest)
    docEl := docEl.BuildUpdatedCache(cacheRequest)
}
Esc::ExitApp