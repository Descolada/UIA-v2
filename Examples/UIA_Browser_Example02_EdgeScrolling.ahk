#Requires AutoHotkey v2
;#include <UIA_Interface> ; Uncomment if you have moved UIA_Interface.ahk to your main Lib folder
#include ..\Lib\UIA.ahk
;#include <UIA_Browser> ; Uncomment if you have moved UIA_Browser.ahk to your main Lib folder
#include ..\Lib\UIA_Browser.ahk

/**
 * A small example for Edge, demostrating form filling (the search box) and scrolling the page.
 */

; Run in Incognito mode to avoid any extensions interfering. 
Run "msedge.exe -inprivate" 
WinWaitActive "ahk_exe msedge.exe"
Sleep 500
; Initialize UIA_Browser, use Last Found Window
cUIA := UIA_Browser() 
; Wait the "New inprivate tab" (case insensitive) page to load with a timeout of 5 seconds
cUIA.WaitPageLoad("New inprivate tab", 5000) 
; Set the URL to google and navigate
cUIA.Navigate("google.com") 

; First lets make sure the selected language is correct. 
; This waits an element to exist where ClassName is "neDYw tHlp8d" and ControlType is Button OR an element with a MenuItem type.
if (langBut := cUIA.WaitElement([{ClassName:"neDYw tHlp8d", Type:"Button"}, {Type:"MenuItem"}],1000)) { 
	; Check that it is collapsed
	if (langBut.ExpandCollapseState == UIA.ExpandCollapseState.Collapsed)
		langBut.Expand()
	; Select the English language
	cUIA.WaitElement({Name:"English", Type:"MenuItem"}).Click() 
	; If the "I agree" or "Accept all" button exists, then click it to get rid of the consent form
	cUIA.WaitElement({Type:"Button", or:[{Name:"Accept all"}, {Name:"I agree"}]}).Click() 
}
; Looking for a partial name match "Searc" OR the ClassName for the search box (found using UIAViewer), using matchMode=Substring. 
; WaitElement instead of FindElement is used here, because if the "I agree" button was clicked then this element might not exist right away, so lets first wait for it to exist.
searchBox := cUIA.WaitElement({or:[{Name:"Searc"}, {ClassName:"gLFyf"}], Type:"ComboBox", matchmode:"Substring"}) 
; Set the search box text to "autohotkey forums"
searchBox.Value := "autohotkey forums" 
; Click the search button to search (either Name "Google Search" OR the ClassName for it)
cUIA.FindElement([{Name:"Google Search"}, {ClassName:"gNO89b"}]).Click() 
cUIA.WaitPageLoad()

; Now that the Google search results have loaded, lets scroll the page to the end.
; First get the document element
docEl := cUIA.GetCurrentDocumentElement() 
ToolTip "Current scroll percent: " docEl.VerticalScrollPercent
; Lets scroll down in steps of 10%
for percent in [10, 20, 30, 40, 50, 60, 70, 80, 90, 100] { 
	docEl.VerticalScrollPercent := percent
	Sleep 500
	ToolTip "Current scroll percent: " docEl.VerticalScrollPercent
}
Sleep 3000
ToolTip
ExitApp
