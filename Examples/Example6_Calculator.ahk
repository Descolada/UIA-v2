;#include <UIA> ; Uncomment if you have moved UIA_Interface.ahk to your main Lib folder
#include ..\Lib\UIA.ahk
SetTitleMatchMode 3

Run "calc.exe"
WinWaitActive "Kalkulaator"
cEl := UIA.ElementFromHandle("Kalkulaator") ; Get the element for the Calculator window
; All the calculator buttons are of "Button" ControlType, and if the system language is English then the Name of the elements are the English words for the buttons (eg button 5 is named "Five", = sign is named "Equals")
cEl.WaitExist({Name:"Kuus"}).Click() ; Wait for the "Six" button by name and click it
cEl.FindFirst({Name:"Viis", Type:"Button"}).Click() ; Specify both name "Five" and control type "Button"
cEl.FindElement("Button", {Name:"Pluss"}).Click() ; An alternative method to FindFirstBy("Name=Plus")
cEl.FindFirst({Name:"Neli", Type:UIA.ControlType.Button}).Click() ; The type can be specified as "Button", UIA.ButtonControlTypeId, UIA_Enum.UIA_ButtonControlTypeId, or 50000 (which is the value of UIA.ButtonControlTypeId)
cEl.FindFirst({Name:"Võrdub", Type:"Button"}).Click()
ExitApp
