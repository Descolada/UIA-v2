#Requires AutoHotkey v2
;#include <UIA> ; Uncomment if you have moved UIA_Interface.ahk to your main Lib folder
#include ..\Lib\UIA.ahk

lorem := "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."

Run "notepad.exe"
;WinActivate, "ahk_exe notepad.exe"
WinWaitActive "ahk_exe notepad.exe"
NotepadEl := UIA.ElementFromHandle("ahk_exe notepad.exe")
editEl := NotepadEl.FindElement([{Type:"Document"}, {Type:"Edit"}]) ; Get the Edit or Document element (differs between UIAutomation versions)
editEl.Value := lorem ; Set the text to our sample text
textPattern := editEl.TextPattern

MsgBox "TextPattern properties:"
	. "`nDocumentRange: returns the TextRange for all the text inside the element"
	. "`nSupportedTextSelection: " editEl.SupportedTextSelection
	. "`n`nTextPattern methods:"
	. "`nRangeFromPoint(x,y): retrieves an empty TextRange nearest to the specified screen coordinates"
	. "`nRangeFromChild(child): retrieves a text range enclosing a child element such as an image, hyperlink, Microsoft Excel spreadsheet, or other embedded object."
	. "`nGetSelection(): returns the currently selected text"
	. "`nGetVisibleRanges(): retrieves an array of disjoint text ranges from a text-based control where each text range represents a contiguous span of visible text"

wholeRange := editEl.DocumentRange ; Get the TextRange for all the text inside the Edit element

MsgBox "To select a certain phrase inside the text, use FindText() method to get the corresponding TextRange, then Select() to select it.`n`nPress OK to select the text `"dolor sit amet`""
WinActivate "ahk_exe notepad.exe"
wholeRange.FindText("dolor sit amet").Select()
Sleep 1000

; For the next example we need to clone the TextRange, because some methods change the supplied TextRange directly (here we don't want to change our original wholeRange TextRange). An alternative would be to use wholeRange, and after moving the endpoints and selecting the new range, we could call ExpandToEnclosingUnit() to reset the endpoints and get the whole TextRange back 
textSpan := wholeRange.Clone()

MsgBox "To select a span of text, we need to move the endpoints of the TextRange. This can be done with MoveEndpointByUnit.`n`nPress OK to select the text with startpoint of 28 characters from start`nand 390 characters from the end of the sample text"
WinActivate "ahk_exe notepad.exe"
textSpan.MoveEndpointByUnit(UIA.TextPatternRangeEndpoint.Start, UIA.TextUnit.Character, 28) ; Move 28 characters from the start of the sample text
textSpan.MoveEndpointByUnit(UIA.TextPatternRangeEndpoint.End, UIA.TextUnit.Character, -390) ; Move 390 characters backwards from the end of the sample text
textSpan.Select()
Sleep 1000

MsgBox "We can also get the location of texts. Press OK to test it"
br := wholeRange.GetBoundingRectangles()
for k, v in br {
	RangeTip(v.x, v.y, v.w, v.h)
	Sleep 1000
}
RangeTip()

ExitApp

RangeTip(x:="", y:="", w:="", h:="", color:="Red", d:=2) { ; from the FindText library, credit goes to feiyue
  static HighlightGui := []
  if x="" {
    for r in HighlightGui
        r.Destroy()
    HighlightGui := []
    return
  }
  Loop 4 
    HighlightGui.Push(Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale +E0x08000000"))
  Loop 4 {
    i:=A_Index
    , x1:=(i=2 ? x+w : x-d)
    , y1:=(i=3 ? y+h : y-d)
    , w1:=(i=1 or i=3 ? w+2*d : d)
    , h1:=(i=2 or i=4 ? h+2*d : d)
    HighlightGui[i].BackColor := color
    HighlightGui[i].Show("NA x" . x1 . " y" . y1 . " w" . w1 . " h" . h1)
  }
}