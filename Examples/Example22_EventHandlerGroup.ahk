#Requires AutoHotkey v2
;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

Run "calc.exe"
Sleep 1000
cEl := UIA.ElementFromHandle("A")

ehGroup := UIA.CreateEventHandlerGroup()
h1 := UIA.CreateAutomationEventHandler(AutomationEventHandler)
h2 := UIA.CreateNotificationEventHandler(NotificationEventHandler)
ehGroup.AddAutomationEventHandler(h1, UIA.Event.AutomationFocusChanged)
ehGroup.AddNotificationEventHandler(h2)
UIA.AddEventHandlerGroup(ehGroup, cEl)

OnExit(ExitFunc) ; Set up an OnExit call to clean up the handler when exiting the script
return

AutomationEventHandler(sender, eventId) {
	ToolTip "Sender: " sender.Dump()
		. "`nEvent Id: " eventId
	Sleep 500
	SetTimer ToolTip, -3000
}

NotificationEventHandler(sender, notificationKind, notificationProcessing, displayString, activityId) {
    ToolTip "Sender: " sender.Dump() 
        . "`nNotification kind: " notificationKind " (" UIA.NotificationKind[notificationKind] ")"
	    . "`nNotification processing: " notificationProcessing " (" UIA.NotificationProcessing[notificationProcessing] ")"
	    . "`nDisplay string: " displayString
	    . "`nActivity Id: " activityId
	Sleep 500
	SetTimer ToolTip, -3000
}

ExitFunc(*) {
	UIA.RemoveAllEventHandlers()
}

F5::ExitApp
