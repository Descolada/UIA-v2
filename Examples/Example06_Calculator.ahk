#Requires AutoHotkey v2
;#include <UIA> ; Uncomment if you have moved UIA_Interface.ahk to your main Lib folder
#include ..\Lib\UIA.ahk
SetTitleMatchMode 3

Run "calc.exe"
WinWaitActive "Calculator"
; Get the element for the Calculator window
cEl := UIA.ElementFromHandle("Calculator") 
; All the calculator buttons are of "Button" ControlType, and if the system language is English then the Name 
; of the elements are the English words for the buttons (eg button 5 is named "Five", 
; = sign is named "Equals")

; Wait for the "Six" button by name and click it
cEl.WaitElement({Name:"Six"}).Click() 
; Specify both name "Five" and control type "Button"
cEl.FindElement({Name:"Five", Type:"Button"}).Click() 
cEl.FindElement({Type:"Button", Name:"Plus"}).Click() 
; The type can be specified as "Button", UIA.Type.Button, or 50000 (which is the value of UIA.Type.Button)
cEl.FindElement({Name:"Four", Type:UIA.Type.Button}).Click() 
cEl.FindElement({Name:"Equals", Type:"Button"}).Click()
ExitApp
