#Requires AutoHotkey v2
;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

Run "calc.exe"
Sleep 1000
cEl := UIA.ElementFromHandle("A")
MsgBox "Press OK to create a new EventHandler for the Notification event.`nTo test this, interact with the Calculator window, and a tooltip should pop up.`n`nTo exit the script, press F5."
handler := UIA.CreateNotificationEventHandler(NotificationEventHandler)
UIA.AddNotificationEventHandler(handler, cEl)
OnExit(ExitFunc) ; Set up an OnExit call to clean up the handler when exiting the script
return

NotificationEventHandler(sender, notificationKind, notificationProcessing, displayString, activityId) {
    ToolTip "Sender: " sender.Dump() 
        . "`nNotification kind: " notificationKind " (" UIA.NotificationKind[notificationKind] ")"
	    . "`nNotification processing: " notificationProcessing " (" UIA.NotificationProcessing[notificationProcessing] ")"
	    . "`nDisplay string: " displayString
	    . "`nActivity Id: " activityId
	SetTimer ToolTip, -3000
}

ExitFunc(*) {
	UIA.RemoveAllEventHandlers()
}

F5::ExitApp
