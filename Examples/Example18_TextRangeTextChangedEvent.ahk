#Requires AutoHotkey v2
;#include <UIA> ; Uncomment if you have moved UIA_Interface.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

MsgBox "To test this file, create a new Word document (with the title `"Document1 - Word`") and write some sample text in it."

program := "Document1 - Word"
WinActivate program
WinWaitActive program
wordEl := UIA.ElementFromHandle(program)
bodyEl := wordEl.FindElement({AutomationId:"Body"}) ; First get the body element of Word
document := bodyEl.DocumentRange ; Get the TextRange for the whole document

MsgBox "Current text inside Word body element:`n" document.GetText() ; Display the text from the TextRange
WinActivate program

MsgBox "We can get text from a specific attribute, such as text within a `"bullet list`".`nTo test this, create a bullet list (with filled bullets) in Word and press OK."
MsgBox "Found the following text in bullet list:`n" document.FindAttribute(UIA.TextAttribute.BulletStyle, UIA.BulletStyle.FilledRoundBullet).GetText()

Loop {
	out := InputBox("Search text in Word by font. Type some example text in Word.`nThen write a font (such as `"Calibri`") and press OK`n`nNote that this is case-sensitive, and fonts start with a capital letter`n(`"calibri`" is not the same as `"Calibri`")", "Find")
	if out.Result != "OK"
		break
	else if !out.value
		MsgBox "You need to type a font to search!"
	else if (found := document.FindAttribute(UIA.TextAttribute.FontName, out.value))
		MsgBox "Found the following text:`n" found.GetText()
	else
		MsgBox "No text with the font " out.value " found!"
}

MsgBox "Press OK to create a new EventHandler for the TextChangedEvent.`nTo test this, type some new text inside Word, and a tooltip should pop up.`n`nTo exit the script, press F5."
handler := UIA.CreateAutomationEventHandler(TextChangedEventHandler) ; Create a new event handler that points to the function TextChangedEventHandler, which must accept two arguments: element and eventId.
UIA.AddAutomationEventHandler(handler, UIA.Event.Text_TextChanged, wordEl) ; Add a new automation handler for the TextChanged event. Note that we can only use wordEl here, not bodyEd, because the event is handled for the whole window.
OnExit(ExitFunc) ; Set up an OnExit call to clean up the handler when exiting the script

return

TextChangedEventHandler(el, eventId) {
	try {
		ToolTip "You changed text in Word:`n`n" el.DocumentRange.GetText()
		SetTimer RemoveToolTip, -2000
	}
}

ExitFunc(*) {
	global handler, wordEl
	UIA.RemoveAutomationEventHandler(handler, UIA.Event.Text_TextChanged, wordEl) ; Remove the event handler. Alternatively use UIA.RemoveAllEventHandlers() to remove all handlers
}

RemoveToolTip() {
	ToolTip
}

F5::ExitApp
