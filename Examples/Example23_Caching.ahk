#NoEnv
#Warn
#SingleInstance force
SetTitleMatchMode, 2

;#include <UIA_Interface> ; Uncomment if you have moved UIA_Interface.ahk to your main Lib folder
#include ..\Lib\UIA_Interface.ahk

UIA := UIA_Interface()
cacheRequest := UIA.CreateCacheRequest()
cacheRequest.TreeScope := 5 ; Set TreeScope to include the starting element and all descendants as well
cacheRequest.AddProperty("ControlType") ; Add all the necessary properties that DumpAll uses: ControlType, LocalizedControlType, AutomationId, Name, Value, ClassName, AcceleratorKey
cacheRequest.AddProperty("LocalizedControlType")
cacheRequest.AddProperty("AutomationId")
cacheRequest.AddProperty("Name")
cacheRequest.AddProperty("Value")
cacheRequest.AddProperty("ClassName")
cacheRequest.AddProperty("AcceleratorKey")

cacheRequest.AddPattern("Window") ; To use cached patterns, first add the pattern
cacheRequest.AddProperty("WindowCanMaximize") ; Also need to add any pattern properties we wish to use

Run, notepad.exe
WinWaitActive, ahk_exe notepad.exe

npEl:= UIA.ElementFromHandleBuildCache("ahk_exe notepad.exe", cacheRequest) ; Get element and also build the cache
MsgBox, % npEl.CachedDumpAll()
MsgBox, % npEl.CachedWindowPattern.CachedCanMaximize

ExitApp