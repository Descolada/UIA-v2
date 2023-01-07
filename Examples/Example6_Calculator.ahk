;#include <UIA> ; Uncomment if you have moved UIA_Interface.ahk to your main Lib folder
#include ..\Lib\UIA.ahk
SetTitleMatchMode 3

Run "calc.exe"
WinWaitActive "Kalkulaator"
; Get the element for the Calculator window
cEl := UIA.ElementFromHandle("Kalkulaator") 
; All the calculator buttons are of "Button" ControlType, and if the system language is English then the Name 
; of the elements are the English words for the buttons (eg button 5 is named "Five", 
; = sign is named "Equals")

; Wait for the "Six" button by name and click it
cEl.WaitElement({Name:"Kuus"}).Click() 
; Specify both name "Five" and control type "Button"
cEl.FindElement({Name:"Viis", Type:"Button"}).Click() 
cEl.FindElement({Type:"Button", Name:"Pluss"}).Click() 
; The type can be specified as "Button", UIA.ControlType.Button, or 50000 (which is the value of UIA.ControlType.Button)
cEl.FindElement({Name:"Neli", Type:UIA.ControlType.Button}).Click() 
cEl.FindElement({Name:"Võrdub", Type:"Button"}).Click()
ExitApp
