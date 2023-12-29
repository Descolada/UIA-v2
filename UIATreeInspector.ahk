#include <UIA>

TreeInspector()

class TreeInspector {
    static SettingsFolderPath := A_AppData "\UIATreeInspector"
    static SettingsFilePath := A_AppData "\UIATreeInspector\settings.ini"
    __New() {
        local v, pattern, value
        OnError this.ErrorHandler.Bind(this)
        CoordMode "Mouse", "Screen"
        DetectHiddenWindows True
        this.Stored := {hWnd:0, WinList:Map(), FilteredTreeView:Map(), TreeView:Map(), Controls:Map(), HighlightedElement:0}
        this.Capturing := False, this.MacroSidebarVisible := False, this.MacroSidebarWidth := 350, this.MacroSidebarMinWidth := 290, this.GuiMinWidth := 840, this.GuiMinHeight := 400, this.Focused := 1
        this.LoadSettings()
        this.cacheRequest := UIA.CreateCacheRequest()
        ; Don't even get the live element, because we don't need it. Gives a significant speed improvement.
        this.cacheRequest.AutomationElementMode := UIA.AutomationElementMode.None
        ; Set TreeScope to include the starting element and all descendants as well
        this.cacheRequest.TreeScope := 5

        this.gViewer := Gui((this.AlwaysOnTop ? "AlwaysOnTop " : "") "Resize +MinSize" this.GuiMinWidth "x" this.GuiMinHeight, "UIATreeInspector")
        this.gViewer.OnEvent("Close", this.gViewer_Close.Bind(this))
        this.gViewer.OnEvent("Size", this.gViewer_Size.Bind(this))

        this.gViewer.Add("Text", "w140", "Windows and Controls").SetFont("bold")
        this.TVWins := this.gViewer.Add("TreeView", "w300 h465")
        this.TVWins.OnEvent("ItemSelect", this.TVWins_ItemSelect.Bind(this))
        this.ButRefreshTVWins := this.gViewer.Add("Button", "xm+0 y500 w100", "&Refresh list")
        this.ButRefreshTVWins.OnEvent("Click", this.RefreshTVWins.Bind(this))
        this.CBWinVisible := this.gViewer.Add("Checkbox", "xm+110 y505 w50" (this.WinVisible ? " Checked" : ""), "Visible")
        this.CBWinVisible.OnEvent("Click", (Ctrl, *) => (this.WinVisible := Ctrl.Value, this.RefreshTVWins()))
        this.CBWinTitle := this.gViewer.Add("Checkbox", "xm+160 y505 w40" (this.WinTitle ? " Checked" : ""), "Title")
        this.CBWinTitle.OnEvent("Click", (Ctrl, *) => (this.WinTitle := Ctrl.Value, this.RefreshTVWins()))
        this.CBWinActivate := this.gViewer.Add("Checkbox", "xm+200 y505 w60" (this.WinActivate ? " Checked" : ""), "Activate")
        this.CBWinActivate.OnEvent("Click", (Ctrl, *) => (this.WinActivate := Ctrl.Value))

        this.gViewer.Add("Text", "x320 y5 w100", "Window Info").SetFont("bold")
        this.LVWin := this.gViewer.Add("ListView", "x320 y25 h135 w250", ["Property", "Value"])
        this.LVWin.OnEvent("ContextMenu", LV_CopyTextMethod := this.LV_CopyText.Bind(this))
        this.LVWin.ModifyCol(1,60)
        this.LVWin.ModifyCol(2,180)
        for v in ["Title", "Text", "Id", "Location", "Class(NN)", "Process", "PID"]
            this.LVWin.Add(,v,"")
        this.gViewer.Add("Text", "w100", "Properties").SetFont("bold")
        this.LVProps := this.gViewer.Add("ListView", "h200 w250", ["Property", "Value"])
        this.LVProps.OnEvent("ContextMenu", LV_CopyTextMethod)
        this.LVProps.ModifyCol(1,100)
        this.LVProps.ModifyCol(2,140)
        this.DisplayedProps := ["Type", "LocalizedType", "Name", "Value", "AutomationId", "BoundingRectangle", "ClassName", "FullDescription", "HelpText", "AccessKey", "AcceleratorKey", "HasKeyboardFocus", "IsKeyboardFocusable", "ItemType", "ProcessId", "IsEnabled", "IsPassword", "IsOffscreen", "FrameworkId", "IsRequiredForForm", "ItemStatus", "RuntimeId"]
        Loop DisplayedPropsLength := this.DisplayedProps.Length {
            v := this.DisplayedProps[i := DisplayedPropsLength-A_Index+1]
            try this.cacheRequest.AddProperty(v) ; Throws if not available, 
            catch
                this.DisplayedProps.RemoveAt(i) ; Remove property if it is not available
        }
        for v in this.DisplayedProps
            this.LVProps.Add(,v = "BoundingRectangle" ? "Location" : v,"")
        for pattern in [UIA.Property.OwnProps()*] {
            if pattern ~= "Is([\w]+Pattern.?)Available"
                try this.cacheRequest.AddProperty(UIA.Property.%pattern%)
                catch
                    UIA.Property.DeleteProp(pattern) ; Remove pattern if it is not available
        }

        (this.TextTVPatterns := this.gViewer.Add("Text", "w100", "Patterns")).SetFont("bold")
        this.TVPatterns := this.gViewer.Add("TreeView", "h85 w250")
        this.TVPatterns.OnEvent("DoubleClick", this.TVPatterns_DoubleClick.Bind(this))

        this.SBMain := this.gViewer.Add("StatusBar",, "  Right-click to change additional settings")
        this.SBMain.OnEvent("Click", this.SBMain_Click.Bind(this))
        this.SBMain.OnEvent("ContextMenu", this.SBMain_ContextMenu.Bind(this))

        this.gViewer.Add("Text", "x580 y5 w100", "UIA Tree").SetFont("bold")
        this.TVUIA := this.gViewer.Add("TreeView", "x580 y25 w300 h465")
        this.TVUIA.OnEvent("Click", this.TVUIA_Click.Bind(this))
        this.TVUIA.OnEvent("ContextMenu", this.TVUIA_ContextMenu.Bind(this))
        this.TVUIA.Add("Select window or control")
        this.TextFilterTVUIA := this.gViewer.Add("Text", "x275 y503", "&Filter:")
        this.EditFilterTVUIA := this.gViewer.Add("Edit", "x305 y500 w100")
        this.EditFilterTVUIA.OnEvent("Change", this.EditFilterTVUIA_Change.Bind(this))
        
        this.GroupBoxMacro := this.gViewer.Add("GroupBox", "x1200 y20 w" (this.MacroSidebarWidth-20), "Macro creator")
        (this.TextMacroAction := this.gViewer.Add("Text", "x1200 y40 w40", "Action:")).SetFont("bold")
        this.DDLMacroAction := this.gViewer.Add("DDL", "Choose1 x1200 y38 w120", ["No element selected"])
        (this.ButMacroAddElement := this.gViewer.Add("Button","x1200 y37 w90 h20", "&Add element")).SetFont("bold")
        this.ButMacroAddElement.OnEvent("Click", this.ButMacroAddElement_Click.Bind(this))
        (this.EditMacroScript := this.gViewer.Add("Edit", "-Wrap HScroll x1200 y65 h410 w" (this.MacroSidebarWidth-40), "#include <UIA>`n`n")).SetFont("s10") ; Setting a font here disables UTF-8-BOM
        (this.ButMacroScriptRun := this.gViewer.Add("Button", "x1180 y120 w70", "&Test script")).SetFont("bold")
        this.ButMacroScriptRun.OnEvent("Click", this.ButMacroScriptRun_Click.Bind(this))
        this.ButMacroScriptCopy := this.gViewer.Add("Button", "x1220 y120 w70", "&Copy")
        this.ButMacroScriptCopy.OnEvent("Click", (*) => (A_Clipboard := this.EditMacroScript.Text, ToolTip("Macro code copied to Clipboard!"), SetTimer(ToolTip, -3000)))
        this.ButToggleMacroSidebar := this.gViewer.Add("Button", "x490 y500 w120", "Show macro &sidebar =>")
        this.ButToggleMacroSidebar.OnEvent("Click", this.ButToggleMacroSidebar_Click.Bind(this))
        xy := ""
        if this.RememberGuiPosition {
            xy := StrSplit(this.RememberGuiPosition, ","), monitor := 0
            Loop MonitorGetCount() {
                MonitorGetWorkArea(A_Index, &Left, &Top, &Right, &Bottom)
                if xy[1] > (Left-50) && xy[2] > (Top-50) && xy[1] < (Right-50) && xy[2] < (Bottom-30) {
                    monitor := A_Index
                    break
                }
            }
            xy := monitor ? "x" xy[1] " y" xy[2] " " : ""
        }
        this.gViewer.Show(xy "w900 h550")
        this.gViewer_Size(this.gViewer,0,900,550)
        this.RefreshTVWins()
        this.FocusHook := DllCall("SetWinEventHook", "UInt", 0x8005, "UInt", 0x8005, "Ptr",0,"Ptr", CallbackCreate(this.HandleFocusChangedEvent.Bind(this), "F", 7),"UInt", 0, "UInt",0, "UInt",0)
    }
    __Delete() {
        DllCall("UnhookWinEvent", "Ptr", this.FocusHook)
    }
    HandleFocusChangedEvent(hWinEventHook, Event, hWnd, idObject, idChild, dwEventThread, dwmsEventTime) {
        winHwnd := DllCall("GetAncestor", "UInt", hWnd, "UInt", 2)
        try winTitle := WinGetTitle(winHwnd)
        catch
            winTitle := ""
        if winHwnd = this.gViewer.Hwnd || winTitle == "UIATreeInspector" {
            if !this.Focused {
                this.Focused := 1
                if IsObject(this.Stored.HighlightedElement)
                    this.Stored.HighlightedElement.Highlight(0, "Blue", 4)
            }
        } else {
            if this.Focused {
                this.Focused := 0
                if IsObject(this.Stored.HighlightedElement)
                    this.Stored.HighlightedElement.Highlight("clear")
            }
        }
        return 0
    }
    SaveSettings() {
        if !FileExist(A_AppData "\UIATreeInspector")
            DirCreate(A_AppData "\UIATreeInspector")
        IniWrite(this.PathIgnoreNames, TreeInspector.SettingsFilePath, "Path", "IgnoreNames")
        IniWrite(this.PathType, TreeInspector.SettingsFilePath, "Path", "Type")
        IniWrite(this.AlwaysOnTop, TreeInspector.SettingsFilePath, "General", "AlwaysOnTop")
        IniWrite(this.DPIAwareness, UIA.Viewer.SettingsFilePath, "General", "DPIAwareness")
        IniWrite(this.RememberGuiPosition, UIA.Viewer.SettingsFilePath, "General", "RememberGuiPosition")
        IniWrite(this.WinVisible, TreeInspector.SettingsFilePath, "WinTree", "Visible")
        IniWrite(this.WinTitle, TreeInspector.SettingsFilePath, "WinTree", "Title")
        IniWrite(this.WinActivate, TreeInspector.SettingsFilePath, "WinTree", "Activate")
    }
    LoadSettings() {
        this.PathIgnoreNames := IniRead(TreeInspector.SettingsFilePath, "Path", "IgnoreNames", 1)
        this.PathType := IniRead(TreeInspector.SettingsFilePath, "Path", "Type", "")
        this.AlwaysOnTop := IniRead(TreeInspector.SettingsFilePath, "General", "AlwaysOnTop", 1)
        this.DPIAwareness := IniRead(UIA.Viewer.SettingsFilePath, "General", "DPIAwareness", 0)
        this.RememberGuiPosition := IniRead(UIA.Viewer.SettingsFilePath, "General", "RememberGuiPosition", "")
        this.WinVisible := IniRead(TreeInspector.SettingsFilePath, "WinTree", "Visible", 1)
        this.WinTitle := IniRead(TreeInspector.SettingsFilePath, "WinTree", "Title", 1)
        this.WinActivate := IniRead(TreeInspector.SettingsFilePath, "WinTree", "Activate", 1)
    }
    ErrorHandler(Exception, Mode) => (OutputDebug(Format("{1} ({2}) : ({3}) {4}`n", Exception.File, Exception.Line, Exception.What, Exception.Message) (HasProp(Exception, "Extra") ? "    Specifically: " Exception.Extra "`n" : "") "Stack:`n" Exception.Stack "`n`n"), 1)
    gViewer_Close(GuiObj, *) {
        if this.RememberGuiPosition
            WinGetPos(&X, &Y,,,GuiObj.Hwnd), this.RememberGuiPosition := X "," Y, this.SaveSettings()
        ExitApp()
    }
    ; Resizes window controls when window is resized
    gViewer_Size(GuiObj, MinMax, Width, Height) {
        static RedrawFunc := WinRedraw.Bind(GuiObj.Hwnd)
        this.TVUIA.GetPos(&TV_Pos_X, &TV_Pos_Y, &TV_Pos_W, &TV_Pos_H)
        this.MoveControls(this.MacroSidebarVisible ? {Control:this.TVUIA,h:(TV_Pos_H:=Height-TV_Pos_Y-60)} : {Control:this.TVUIA,w:(TV_Pos_W:=Width-TV_Pos_X-10),h:(TV_Pos_H:=Height-TV_Pos_Y-60)})
        TV_Pos_R := TV_Pos_X+TV_Pos_W
        this.LVProps.GetPos(&LVPropsX, &LVPropsY, &LVPropsWidth, &LVPropsHeight)
        this.ButToggleMacroSidebar.GetPos(,,&ButToggleMacroSidebarW)
        this.MoveControls(
            {Control:this.TextFilterTVUIA, x:TV_Pos_X, y:Height-47}, 
            {Control:this.ButToggleMacroSidebar, x:TV_Pos_X+TV_Pos_W-ButToggleMacroSidebarW, y:Height-50}, 
            {Control:this.EditFilterTVUIA, x:TV_Pos_X+30, y:Height-50},
            {Control:this.LVProps,h:Height-LVPropsY-170}, 
            {Control:this.TextTVPatterns,y:Height-165}, 
            {Control:this.TVPatterns,y:Height-145},
            {Control:this.ButRefreshTVWins,y:Height-50},
            {Control:this.CBWinVisible,y:Height-45},
            {Control:this.CBWinTitle,y:Height-45},
            {Control:this.CBWinActivate,y:Height-45},
            {Control:this.TVWins,h:(TV_Pos_H:=Height-TV_Pos_Y-60)})
        if this.MacroSidebarVisible
            this.MacroSidebarWidth := Width-(TV_Pos_X+TV_Pos_W)-10
        this.MoveControls(
            {Control:this.GroupBoxMacro,x:TV_Pos_R+15, w:Width-TV_Pos_R-30, h:TV_Pos_H+35}, 
            {Control:this.TextMacroAction,x:TV_Pos_R+25}, 
            {Control:this.DDLMacroAction,x:TV_Pos_R+70}, 
            {Control:this.ButMacroAddElement,x:TV_Pos_R+this.MacroSidebarWidth-105}, 
            {Control:this.EditMacroScript,x:TV_Pos_R+25,w:Width-TV_Pos_R-50,h:TV_Pos_H-50}, 
            {Control:this.ButMacroScriptRun,x:TV_Pos_R+85,y:TV_Pos_Y+TV_Pos_H-2}, 
            {Control:this.ButMacroScriptCopy,x:TV_Pos_R+this.MacroSidebarWidth-145,y:TV_Pos_Y+TV_Pos_H-2})
        DllCall("RedrawWindow", "ptr", GuiObj.Hwnd, "ptr", 0, "ptr", 0, "uint", 0x0081) ; Reduces flicker compared to RedrawFunc
    }
    MoveControls(ctrls*) {
        for ctrl in ctrls
            ctrl.Control.Move(ctrl.HasOwnProp("x") ? ctrl.x : unset, ctrl.HasOwnProp("y") ? ctrl.y : unset, ctrl.HasOwnProp("w") ? ctrl.w : unset, ctrl.HasOwnProp("h") ? ctrl.h : unset)
    }
    RefreshTVWins(*) {
        DetectHiddenWindows(!this.WinVisible)
        this.TVWins.Delete()
        for hWnd in WinGetList() {
            wTitle := WinGetTitle(hWnd)
            if (this.WinTitle && (wTitle == ""))
                continue
            try wExe := WinGetProcessName(hWnd)
            catch
                wExe := "ERROR"
            parent := this.TVWins.Add(wTitle " (" wExe ")")
            this.Stored.WinList[parent] := hWnd
            for ctrl in WinGetControlsHwnd(hWnd) {
                try {
                    classNN := ControlGetClassNN(ctrl, hWnd)
                    item := this.TVWins.Add(classNN, parent)
                    this.Stored.WinList[item] := ctrl
                    this.Stored.Controls[ctrl] := hWnd
                }
            }
        }
    }
    ; Show/hide macros sidebar
    ButToggleMacroSidebar_Click(GuiCtrlObj?, Info?) {
        local w
        this.MacroSidebarVisible := !this.MacroSidebarVisible
        GuiCtrlObj.Text := this.MacroSidebarVisible ? "Hide macro &sidebar <=" : "Show macro &sidebar =>"
        this.gViewer.GetPos(,, &w)
        this.gViewer.Opt("+MinSize" (this.MacroSidebarVisible ? w + this.MacroSidebarMinWidth : this.GuiMinWidth) "x" this.GuiMinHeight)
        this.gViewer.Move(,,w+(this.MacroSidebarVisible ? this.MacroSidebarWidth : -this.MacroSidebarWidth))
    }
    ; Handles adding elements with actions to the macro Edit
    ButMacroAddElement_Click(GuiCtrlObj?, Info?) {
        local match
        if !this.Stored.HasOwnProp("CapturedElement")
            return
        processName := WinGetProcessName(this.Stored.hWnd)
        winElVariable := RegExMatch(processName, "^[^ .\d]+", &match:="") ? RegExReplace(match[], "[^\p{L}0-9_#@$]") "El" : "winEl" ; Leaves only letters, numbers, and symbols _#@$ (allowed AHK characters)
        winTitle := "`"" WinGetTitle(this.Stored.hWnd) " ahk_exe " processName "`""
        winElText := winElVariable " := UIA.ElementFromHandle(" (this.Stored.Controls.Has(ctrl := this.Stored.hWnd) ? "ControlGetHwnd(`"" ControlGetClassNN(ctrl, this.Stored.Controls[ctrl]) "`", " winTitle ")" : winTitle) ")"
        if !InStr(this.EditMacroScript.Text, winElText) || RegExMatch(this.EditMacroScript.Text, "\Q" winElText "\E(?=[\w\W]*\QwinEl := UIA.ElementFromHandle(`"ahk_exe\E)")
            this.EditMacroScript.Text := RTrim(this.EditMacroScript.Text, "`r`n`t ") "`r`n`r`n" winElText
        else
            this.EditMacroScript.Text := RTrim(this.EditMacroScript.Text, "`r`n`t ")
        winElVariable := winElVariable (SubStr(this.SBMain.Text, 9) ? ".ElementFromPath(" SubStr(this.SBMain.Text, 9) ")" : "") (this.DDLMacroAction.Text ? "." this.DDLMacroAction.Text : "")
        if InStr(this.DDLMacroAction.Text, "Dump")
            winElVariable := "MsgBox(" winElVariable ")"
        this.EditMacroScript.Text := this.EditMacroScript.Text "`r`n" RegExReplace(winElVariable, "(?<!``)`"", "`"") "`r`n"
    }
    ; Tries to run the code in the macro Edit
    ButMacroScriptRun_Click(GuiCtrlObj?, Info?) {
        static tempFileName := "~UIATreeInspectorMacro.tmp"
        if IsObject(this.Stored.HighlightedElement)
            this.Stored.HighlightedElement.Highlight("clear"), this.Stored.HighlightedElement := 0
        DetectHiddenWindows 1
        WinHide(this.gViewer)
        try FileDelete(tempFileName)
        try {
            FileAppend(StrReplace(this.EditMacroScript.Text, "`r"), tempFileName, "UTF-8") 
            Run(A_AhkPath " /force /cp65001 `"" A_ScriptDir "\" tempFileName "`"",,,&pid)
            if WinWait("ahk_pid " pid,, 3)
                WinWaitClose(, , 30)
        }
        if IsSet(pid) && WinExist("ahk_pid " pid)
            WinKill
        try FileDelete(tempFileName)
        WinShow(this.gViewer)
        DetectHiddenWindows 0
    }
    ; Handles right-clicking a listview (copies to clipboard)
    LV_CopyText(GuiCtrlObj, Info, *) {
        local LVData, out := "", Property
        LVData := Info > GuiCtrlObj.GetCount()
            ? ListViewGetContent("", GuiCtrlObj)
            : ListViewGetContent("Selected", GuiCtrlObj)
        for LVData in StrSplit(LVData, "`n") {
            LVData := StrSplit(LVData, "`t",,2)
            if LVData.Length < 2
                continue
            switch LVData[1], 0 {
                case "Type":
                    LVData[2] := "`"" RTrim(SubStr(LVData[2],8), ")") "`""
                case "Location":
                    LVData[2] := "{" RegExReplace(LVData[2], "(\w:) (\d+)(?= )", "$1$2,") "}"
            }
            Property := -1
            try Property := UIA.Property.%LVData[1]%
            out .= ", " (GuiCtrlObj.Hwnd = this.LVWin.Hwnd ? "" : LVData[1] ":") (UIA.PropertyVariantTypeBSTR.Has(Property) ? "`"" StrReplace(StrReplace(LVData[2], "``", "````"), "`"", "```"") "`"" : LVData[2])
        }
        ToolTip("Copied: " (A_Clipboard := SubStr(out, 3)))
        SetTimer(ToolTip, -3000)
    }
    ; Handles running pattern methods, first trying to find the live element by RuntimeId
    TVPatterns_DoubleClick(GuiCtrlObj, Info) {
        if !Info
            return
        Item := GuiCtrlObj.GetText(Info)
        if !InStr(Item, "()")
            return
        Item := SubStr(Item, 1, -2)
        if !(CurrentEl := UIA.ElementFromHandle(this.Stored.hWnd).ElementExist({RuntimeId:this.Stored.CapturedElement.CachedRuntimeId}))
            return MsgBox("Live element not found!",,"4096")
        if Item ~= "Value|Scroll(?!Into)" {
            this.gViewer.Opt("-AlwaysOnTop")
            Ret := InputBox("Insert value", Item, "W200 H120")
            this.gViewer.Opt("+AlwaysOnTop")
            if Ret.Result != "OK"
                return
        }
        parent := DllCall("GetAncestor", "UInt", this.Stored.hWnd, "UInt", 2)
        WinActivate(parent)
        WinWaitActive(parent,1)
        try CurrentEl.%GuiCtrlObj.GetText(GuiCtrlObj.GetParent(Info)) "Pattern"%.%Item%(IsSet(Ret) ? Ret.Value : unset)
    }
    ; Copies the UIA path to clipboard when statusbar is clicked
    SBMain_Click(GuiCtrlObj, Info, *) {
        if InStr(this.SBMain.Text, "Path:") {
            ToolTip("Copied: " (A_Clipboard := SubStr(this.SBMain.Text, 9)))
            SetTimer(ToolTip, -3000)
        }
    }
    ; StatusBar context menu creation
    SBMain_ContextMenu(GuiCtrlObj, Item, IsRightClick, X, Y) {
        SBMain_Menu := Menu()
        if InStr(this.SBMain.Text, "Path:") {
            SBMain_Menu.Add("Copy UIA path", (*) => (ToolTip("Copied: " (A_Clipboard := this.Stored.CapturedElement.Path)), SetTimer(ToolTip, -3000)))
            SBMain_Menu.Add("Copy condition path", (*) => (ToolTip("Copied: " (A_Clipboard := this.Stored.CapturedElement.ConditionPath)), SetTimer(ToolTip, -3000)))
            SBMain_Menu.Add("Copy numeric path", (*) => (ToolTip("Copied: " (A_Clipboard := this.Stored.CapturedElement.NumericPath)), SetTimer(ToolTip, -3000)))
            SBMain_Menu.Add()
        }
        SBMain_Menu.Add("Display UIA path (relatively reliable, shortest)", (*) => (this.PathType := this.PathType = "" ? "" : "", this.Stored.HasOwnProp("CapturedElement") && this.Stored.CapturedElement.HasOwnProp("Path") ? this.SBMain.SetText("  Path: " (this.PathType = "Numeric" ? this.Stored.CapturedElement.NumericPath : this.Stored.CapturedElement.Path)) : 1))
        SBMain_Menu.Add("Display numeric path (least reliable, short)", (*) => (this.PathType := this.PathType = "Numeric" ? "" : "Numeric", this.Stored.HasOwnProp("CapturedElement") && this.Stored.CapturedElement.HasOwnProp("Path") ? this.SBMain.SetText("  Path: " (this.PathType = "Numeric" ? this.Stored.CapturedElement.NumericPath : this.Stored.CapturedElement.Path)) : 1))
        SBMain_Menu.Add("Display condition path (most reliable, longest)", (*) => (this.PathType := this.PathType = "Condition" ? "" : "Condition", this.Stored.HasOwnProp("CapturedElement") && this.Stored.CapturedElement.HasOwnProp("Path") ? this.SBMain.SetText("  Path: " (this.PathType = "Condition" ? this.Stored.CapturedElement.ConditionPath : this.Stored.CapturedElement.Path)) : 1))
        SBMain_Menu.Add("Ignore Name properties in condition path", (*) => (this.PathIgnoreNames := !this.PathIgnoreNames))
        if this.PathIgnoreNames
            SBMain_Menu.Check("Ignore Name properties in condition path")
        if this.PathType = ""
            SBMain_Menu.Check("Display UIA path (relatively reliable, shortest)")
        if this.PathType = "Numeric"
            SBMain_Menu.Check("Display numeric path (least reliable, short)")
        if this.PathType = "Condition"
            SBMain_Menu.Check("Display condition path (most reliable, longest)")
        SBMain_Menu.Add()
        SBMain_Menu.Add("UIATreeInspector always on top", (*) => (this.AlwaysOnTop := !this.AlwaysOnTop, this.gViewer.Opt((this.AlwaysOnTop ? "+" : "-") "AlwaysOnTop")))
        if this.AlwaysOnTop
            SBMain_Menu.Check("UIATreeInspector always on top")
        SBMain_Menu.Add("Remember UIATreeInspector position", (*) => (this.RememberGuiPosition := !this.RememberGuiPosition, this.SaveSettings()))
        if this.RememberGuiPosition
            SBMain_Menu.Check("Remember UIATreeInspector position")
        SBMain_Menu.Add("Enable DPI awareness", (*) => (this.DPIAwareness := !this.DPIAwareness, this.DPIAwareness ? UIA.SetMaximumDPIAwareness() : UIA.DPIAwareness := -2))
        if this.DPIAwareness
            SBMain_Menu.Check("Enable DPI awareness")
        SBMain_Menu.Add()
        SBMain_Menu.Add("Save settings", (*) => (this.SaveSettings(), ToolTip("Settings saved!"), SetTimer(ToolTip, -2000)))
        SBMain_Menu.Show()
    }
    ; Updates the GUI with the selected item
    TVWins_ItemSelect(GuiCtrlObj, Info) {
        local hWnd := this.Stored.WinList[Info], parent := DllCall("GetAncestor", "UInt", hWnd, "UInt", 2)
        this.Stored.hWnd := hWnd
        if UIA.ProcessIsElevated(WinGetPID(parent)) > 0 && !A_IsAdmin {
            if MsgBox("The inspected window is running with elevated privileges.`nUIATreeInspector must be running in UIAccess mode or as administrator to inspect it.`n`nRun UIATreeInspector as administrator to inspect it?",, 0x1000 | 0x30 | 0x4) = "Yes" {
                try {
                    Run('*RunAs "' (A_IsCompiled ? A_ScriptFullPath '" /restart' : A_AhkPath '" /restart "' A_ScriptFullPath '"'))
                    ExitApp
                }
            }
        }
        this.cacheRequest.TreeScope := 1
        try this.Stored.CapturedElement := UIA.ElementFromHandle(hWnd, this.cacheRequest)
        this.cacheRequest.TreeScope := 5
        propsOrder := ["Title", "Text", "Id", "Location", "Class(NN)", "Process", "PID"]
        if this.WinActivate {
            WinActivate(parent)
            WinWaitActive(parent, 1)
            WinActivate(this.gViewer.Hwnd)
        }
        WinGetPos(&wX, &wY, &wW, &wH, hWnd)
        props := Map("Title", WinGetTitle(hWnd), "Text", WinGetText(hWnd), "Id", hWnd, "Location", "x: " wX " y: " wY " w: " wW " h: " wH, "Class(NN)", WinGetClass(hWnd), "Process", WinGetProcessName(hWnd), "PID", WinGetPID(hWnd))
        this.LVWin.Delete()
        for propName in propsOrder
            this.LVWin.Add(,propName,props[propName])
        this.PopulatePropsPatterns(this.Stored.CapturedElement)
        this.ConstructTreeView()
    }
    ; Populates the listview with UIA element properties
    PopulatePropsPatterns(Element) {
        local v, value, pattern, parent, proto, match, X, Y, W, H
        if IsObject(this.Stored.HighlightedElement)
            this.Stored.HighlightedElement.Highlight("clear")
        this.Stored.HighlightedElement := Element
        try { ; Show the Highlight only if the window is visible and
            WinGetPos(&X, &Y, &W, &H, this.Stored.hWnd)
            if IsObject(this.Stored.HighlightedElement) && (elBR := this.Stored.HighlightedElement.CachedBoundingRectangle) && UIA.IntersectRect(X, Y, X+W, Y+H, elBR.l, elBR.t, elBR.r, elBR.b)
                Element.Highlight(0, "Blue", 4) ; Indefinite show
        }
        this.LVProps.Delete()
        this.TVPatterns.Delete()
        for v in this.DisplayedProps {
            try prop := Element.Cached%v%
            switch v, 0 {
                case "Type":
                    try name := UIA.Type[prop]
                    catch
                        name := "Unknown"
                    try this.LVProps.Add(, v, prop " (" name ")")
                case "BoundingRectangle":
                    prop := prop ? prop : {l:0,t:0,r:0,b:0}
                    try this.LVProps.Add(, "Location", "x: " prop.l " y: " prop.t " w: " (prop.r - prop.l) " h: " (prop.b - prop.t))
                case "RuntimeId":
                    continue ; Don't display this for now, since it might confuse users into using it as a search property.
                    ; try this.LVProps.Add(, v, UIA.RuntimeIdToString(prop)) ; Uncomment for debugging purposes
                default:
                    try this.LVProps.Add(, v, prop)
            }
            prop := ""
        }
        lastAction := this.DDLMacroAction.Text
        this.DDLMacroAction.Delete()
        this.DDLMacroAction.Add(['', 'Click()', 'Click("left")', 'ControlClick()', 'SetFocus()', 'ShowContextMenu()', 'Highlight()', 'Dump()','DumpAll()'])
        for pattern, value in UIA.Property.OwnProps() {
            if RegExMatch(pattern, "Is([\w]+)Pattern(\d?)Available", &match:=0) && Element.GetCachedPropertyValue(value) {
                parent := this.TVPatterns.Add(match[1] (match.Count > 1 ? match[2] : ""))
                if !IsObject(UIA.IUIAutomation%match[1]%Pattern)
                    continue
                proto := UIA.IUIAutomation%match[1]%Pattern.Prototype
                switch match[1], 0 {
                    case "Invoke":
                        this.DDLMacroAction.Add(['Invoke()'])
                    case "ExpandCollapse":
                        this.DDLMacroAction.Add(['Expand()', 'Collapse()'])
                    case "Value":
                        this.DDLMacroAction.Add(['Value := "value"'])
                    case "Toggle":
                        this.DDLMacroAction.Add(['Toggle()'])
                    case "SelectionItem":
                        this.DDLMacroAction.Add(['Select()', 'AddToSelection()', 'RemoveFromSelection()'])
                    case "ScrollItem":
                        this.DDLMacroAction.Add(['ScrollIntoView()'])
                }
                for name in proto.OwnProps() {
                    if name ~= "i)^(_|Cached)"
                        continue
                    this.TVPatterns.Add(name (proto.GetOwnPropDesc(name).HasOwnProp("call") ? "()" : ""), parent)
                }
            }
        }
        try this.DDLMacroAction.Choose(lastAction)
        catch
            this.DDLMacroAction.Choose(7)
    }
    ; Handles selecting elements in the UIA tree, highlights the selected element
    TVUIA_Click(GuiCtrlObj, Info) {
        if this.Capturing
            return
        try Element := this.EditFilterTVUIA.Value ? this.Stored.FilteredTreeView[Info] : this.Stored.TreeView[Info]
        if IsSet(Element) && Element {
            if IsObject(this.Stored.HighlightedElement) {
                if this.SafeCompareElements(Element, this.Stored.HighlightedElement)
                    return (this.Stored.HighlightedElement.Highlight("clear"), this.Stored.HighlightedElement := 0)
            }
            this.Stored.CapturedElement := Element
            try this.SBMain.SetText("  Path: " (this.PathType = "Numeric" ? Element.NumericPath : this.PathType = "Condition" ? Element.ConditionPath : Element.Path))
            this.PopulatePropsPatterns(Element)
        }
    }
    ; Permits copying the Dump of UIA element(s) to clipboard
    TVUIA_ContextMenu(GuiCtrlObj, Item, IsRightClick, X, Y) {
        TVUIA_Menu := Menu()
        try Element := this.EditFilterTVUIA.Value ? this.Stored.FilteredTreeView[Item] : this.Stored.TreeView[Item]
        if IsSet(Element)
            TVUIA_Menu.Add("Copy to Clipboard", (*) => (ToolTip("Copied Dump() output to Clipboard!"), A_Clipboard := Element.CachedDump(), SetTimer((*) => ToolTip(), -3000)))
        TVUIA_Menu.Add("Copy Tree to Clipboard", (*) => (ToolTip("Copied DumpAll() output to Clipboard!"), A_Clipboard := UIA.ElementFromHandle(this.Stored.hWnd, this.cacheRequest).DumpAll(), SetTimer((*) => ToolTip(), -3000)))
        TVUIA_Menu.Show()
    }
    ; Handles filtering the UIA elements inside the TreeView when the text hasn't been changed in 500ms.
    ; Sorts the results by UIA properties.
    EditFilterTVUIA_Change(GuiCtrlObj, Info, *) {
        static TimeoutFunc := "", ChangeActive := False
        if !this.Stored.TreeView.Count
            return
        if (Info != "DoAction") || ChangeActive {
            if !TimeoutFunc
                TimeoutFunc := this.EditFilterTVUIA_Change.Bind(this, GuiCtrlObj, "DoAction")
            SetTimer(TimeoutFunc, -500)
            return
        }
        ChangeActive := True
        this.Stored.FilteredTreeView := Map(), parents := Map()
        if !(searchPhrase := this.EditFilterTVUIA.Value) {
            this.ConstructTreeView()
            ChangeActive := False
            return
        }
        this.TVUIA.Delete()
        temp := this.TVUIA.Add("Searching...")
        Sleep -1
        this.TVUIA.Opt("-Redraw")
        this.TVUIA.Delete()
        for index, Element in this.Stored.TreeView {
            for prop in this.DisplayedProps {
                try {
                    if InStr(Element.Cached%Prop%, searchPhrase) {
                        if !parents.Has(prop)
                            parents[prop] := this.TVUIA.Add(prop,, "Expand")
                        this.Stored.FilteredTreeView[this.TVUIA.Add(this.GetShortDescription(Element), parents[prop], "Expand")] := Element
                    }
                }
            }
        }
        if !this.Stored.FilteredTreeView.Count
            this.TVUIA.Add("No results found matching `"" searchPhrase "`"")
        this.TVUIA.Opt("+Redraw")
        TimeoutFunc := "", ChangeActive := False
    }
    ; Populates the TreeView with the UIA tree when capturing and the mouse is held still
    ConstructTreeView() {
        local k, v
        this.TVUIA.Delete()
        this.TVUIA.Add("Constructing Tree, please wait...")
        Sleep -1
        this.TVUIA.Opt("-Redraw")
        this.TVUIA.Delete()
        this.Stored.TreeView := Map()
        try this.RecurseTreeView(UIA.ElementFromHandle(this.Stored.hWnd, this.cacheRequest))
        catch {
            this.Stored.TreeView := []
            this.TVUIA.Add("Error: unspecified error (window not found?)")
        }
        
        this.TVUIA.Opt("+Redraw")
        this.SBMain.SetText("  Path: ")
        if !this.Stored.CapturedElement.HasOwnProp("Path") {
            this.Stored.CapturedElement.DefineProp("Path", {Value:""})
            this.Stored.CapturedElement.DefineProp("NumericPath", {Value:""})
            this.Stored.CapturedElement.DefineProp("ConditionPath", {Value:""})
        }
        for k, v in this.Stored.TreeView {
            if this.SafeCompareElements(this.Stored.CapturedElement, v) {
                this.TVUIA.Modify(k, "Vis Select"), this.SBMain.SetText("  Path: " (this.PathType = "Numeric" ? v.NumericPath : this.PathType = "Condition" ? v.ConditionPath : v.Path))
                , this.Stored.CapturedElement.Path := v.Path
                , this.Stored.CapturedElement.NumericPath := v.NumericPath
                , this.Stored.CapturedElement.ConditionPath := v.ConditionPath
            }
        }
    }
    ; Stores the UIA tree with corresponding path values for each element
    RecurseTreeView(Element, parent:=0, path:="", conditionpath := "", numpath:="") {
        local info, child, type, k, paths := Map(), childInfo := [], children := Element.CachedChildren
        Element.DefineProp("Path", {value:"`"" path "`""})
        Element.DefineProp("ConditionPath", {value:conditionpath})
        Element.DefineProp("NumericPath", {value:numpath})
        this.Stored.TreeView[TWEl := this.TVUIA.Add(this.GetShortDescription(Element), parent, "Expand")] := Element
        ; First count up all multiple-condition-index conditions and type-index conditions
        ; This is to know whether the condition is the last of the matching ones, so we can use index -1
        ; This gives an important speed difference over regular indexing
        for child in children {
            compactCondition := this.GetCompactCondition(child, &paths, &typeCondition := "", &type:="")
            childInfo.Push([compactCondition, paths[compactCondition], typeCondition, paths[typeCondition], type])
        }
        ; Now create the final conditions and recurse the tree
        for k, child in children {
            info := childInfo[k], compactCondition := info[1], conditionIndex := info[2]
            if conditionIndex > 1 && conditionIndex = paths[compactCondition]
                conditionIndex := -1
            compactCondition .= conditionIndex = 1 ? "}" : ", i:" conditionIndex "}"
            typeIndex := info[4]
            if typeIndex > 1 && typeIndex = paths[info[3]]
                typeIndex := -1
            this.RecurseTreeView(child, TWEl, path UIA.EncodePath([typeIndex = 1 ? {Type:info[5]} : {Type:info[5], i:typeIndex}]), conditionpath (conditionpath?", ":"") compactCondition, numpath (numpath?",":"") k)
        }
    }
    ; CompareElements sometimes fails to match elements, so this compares some properties instead
    SafeCompareElements(e1, e2) {
        if e1.CachedDump() == e2.CachedDump() {
            try {
                if UIA.RuntimeIdToString(e1.CachedRuntimeId) == UIA.RuntimeIdToString(e2.CachedRuntimeId)
                    return 1
            }
            br_e1 := e1.CachedBoundingRectangle, br_e2 := e2.CachedBoundingRectangle
            return br_e1.l = br_e2.l && br_e1.t = br_e2.t && br_e1.r = br_e2.r && br_e1.b = br_e2.b
        }
        return 0
    }
    ; Creates a short description string for the UIA tree elements
    GetShortDescription(Element) {
        local elDesc := " `"`""
        try elDesc := " `"" Element.CachedName "`""
        try elDesc := Element.CachedLocalizedType elDesc
        catch
            elDesc := "`"`"" elDesc
        return elDesc
    }
    GetCompactCondition(Element, &pathsMap, &t := "", &type := "", &automationId := "", &className := "", &name := "") {
        local n := "", c := "", a := ""
        type := Element.CachedType
        t := "{T:" (type-50000)
        pathsMap[t] := pathsMap.Has(t) ? pathsMap[t] + 1 : 1
        try a := StrReplace(automationId := Element.CachedAutomationId, "`"", "```"")
        if a != "" && !IsInteger(a) { ; Ignore Integer AutomationIds, since they seem to be auto-generated in Chromium apps
            a := t ",A:`"" a "`""
            pathsMap[a] := pathsMap.Has(a) ? pathsMap[a] + 1 : 1 ; This actually shouldn't be needed, if AutomationId's are unique
        }
        try c := StrReplace(className := Element.CachedClassName, "`"", "```"")
        if c != "" {
            c := t ",CN:`"" c "`""
            pathsMap[c] := pathsMap.Has(c) ? pathsMap[c] + 1 : 1
        }
        try n := StrReplace(name := Element.CachedName, "`"", "```"")
        if !this.PathIgnoreNames && n != "" { ; Consider Name last, because it can change (eg. window title)
            n := t ",N:`"" n "`""
            pathsMap[n] := pathsMap.Has(n) ? pathsMap[n] + 1 : 1
        }
        if pathsMap[t] = 1
            return t
        if a != "" && !IsInteger(a) {
            return c != "" ? (pathsMap[a] <= pathsMap[c] ? a : c) : a
        } else if c != ""
            return c
        else if !this.PathIgnoreNames && n != "" && (pathsMap[n] < pathsMap[t])
            return n
        return t
    }
}