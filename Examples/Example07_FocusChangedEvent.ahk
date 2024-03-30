#Requires AutoHotkey v2
;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

EventHandler(el) {
	try {
		ToolTip "Caught event!`nElement name: " el.Name
	}
}

ExitFunc(*) {
	global h
	UIA.RemoveFocusChangedEventHandler(h) ; Remove the event handler. Alternatively use UIA.RemoveAllEventHandlers() to remove all handlers
}

browserExe := "chrome.exe"
Run browserExe " -incognito"
WinWaitActive "ahk_exe " browserExe

global h := UIA.CreateFocusChangedEventHandler(EventHandler) ; Create a new FocusChanged event handler that calls the function EventHandler (required arguments: element)
UIA.AddFocusChangedEventHandler(h) ; Add a new FocusChangedEventHandler
OnExit(ExitFunc) ; Set up an OnExit call to clean up the handler when exiting the script
return

F5::ExitApp
