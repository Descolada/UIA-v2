#Requires AutoHotkey v2
;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

TextSelectionChangedEventHandler(el, eventId) {
	textPattern := el.TextPattern
	selectionArray := textPattern.GetSelection() ; Gets the currently selected text in Notepad as an array of TextRanges (some elements support selecting multiple pieces of text at the same time, thats why an array is returned)
	selectedRange := selectionArray[1] ; Our range of interest is the first selection (TextRange)
	wholeRange := textPattern.DocumentRange ; For comparison, get the whole range (TextRange) of the document
	selectionStart := selectedRange.CompareEndpoints(UIA.TextPatternRangeEndpoint.Start, wholeRange, UIA.TextPatternRangeEndpoint.Start) ; Compare the start point of the selection to the start point of the whole document
	selectionEnd := selectedRange.CompareEndpoints(UIA.TextPatternRangeEndpoint.End, wholeRange, UIA.TextPatternRangeEndpoint.Start) ; Compare the end point of the selection to the start point of the whole document

	ToolTip "Selected text: " selectedRange.GetText() "`nSelection start location: " selectionStart "`nSelection end location: " selectionEnd  ; Display the selected text and locations of selection
}

ExitFunc(*) {
	global handler, NotepadEl
	try UIA.RemoveAutomationEventHandler(UIA.Event.Text_TextSelectionChanged, NotepadEl, handler) ; Remove the event handler. Alternatively use UIA.RemoveAllEventHandlers() to remove all handlers. If the Notepad window doesn't exist any more, this throws an error.
}

; Some sample text to play around with
lorem := "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."

Run "notepad.exe"
WinWaitActive "ahk_exe notepad.exe"

try {
NotepadEl := UIA.ElementFromHandle("ahk_exe notepad.exe")
DocumentControl := NotepadEl.FindElement([{Type:"Document"}, {Type:"Edit"}]) ; If UIA Interface version is 1, then the ControlType is Edit instead of Document!
} catch {
    ; Windows 11 has broken Notepad, so the Document element isn't findable; instead get the focused element
    Sleep 40
    DocumentControl := UIA.ElementFromHandle(ControlGetHwnd("RichEditD2DPT1"))
}
DocumentControl.Value := lorem ; Set the value to our sample text

handler := UIA.CreateAutomationEventHandler(TextSelectionChangedEventHandler) ; Create a new event handler that points to the function TextSelectionChangedEventHandler, which must accept two arguments: element and eventId.
UIA.AddAutomationEventHandler(handler, NotepadEl, UIA.Event.Text_TextSelectionChanged) ; Add a new automation handler for the TextSelectionChanged event
OnExit(ExitFunc) ; Set up an OnExit call to clean up the handler when exiting the script
return

F5::ExitApp
