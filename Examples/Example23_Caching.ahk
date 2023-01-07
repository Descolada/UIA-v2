;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

cacheRequest := UIA.CreateCacheRequest()
; Set TreeScope to include the starting element and all descendants as well
cacheRequest.TreeScope := 5 
; Add all the necessary properties that DumpAll uses: ControlType, LocalizedControlType, AutomationId, Name, Value, ClassName, AcceleratorKey
cacheRequest.AddProperty("ControlType") 
cacheRequest.AddProperty("LocalizedControlType")
cacheRequest.AddProperty("AutomationId")
cacheRequest.AddProperty("Name")
cacheRequest.AddProperty("Value")
cacheRequest.AddProperty("ClassName")
cacheRequest.AddProperty("AcceleratorKey")

; To use cached patterns, first add the pattern
cacheRequest.AddPattern("Window") 
; Also need to add any pattern properties we wish to use
cacheRequest.AddProperty("WindowCanMaximize") 

Run "notepad.exe"
WinWaitActive "ahk_exe notepad.exe"

; Get element and also build the cache
npEl:= UIA.ElementFromHandle("ahk_exe notepad.exe",0, cacheRequest)
; We now have a cached "snapshot" of the window from which we can access our desired elements faster.
MsgBox npEl.CachedDumpAll()
MsgBox npEl.CachedWindowPattern.CachedCanMaximize

ExitApp