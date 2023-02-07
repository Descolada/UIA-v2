/*
	Introduction:
	UIA_Browser implements some methods to help automate browsers with UIAutomation framework.

	Initiate new instance of UIA_Browser with
		cUIA := UIA_Browser(wTitle="")
			wTitle: the title of the browser
		Example: cUIA := UIA_Browser("ahk_exe chrome.exe")
	
	Instances for specific browsers may be initiated with UIA_Chrome, UIA_Edge, UIA_Mozilla (arguments are the same as for UIA_Browser).
	These are usually auto-detected by UIA_Browser, so do not have to be used.

	Available properties for UIA_Browser:
	BrowserId
		ahk_id of the browser window
	BrowserType
		"Chrome", "Edge", "Mozilla" or "Unknown"
	BrowserElement
		The browser window element, which can also be accessed by just calling an element method from UIA_Browser (cUIA.FindFirst would call FindFirst method on the BrowserElement, is equal to cUIA.BrowserElement.FindFirst)
	MainPaneElement
		Element for the upper part of the browser containing the URL bar, tabs, extensions etc
	URLEditElement
		Element for the address bar

	UIA_Browser methods:
	GetCurrentMainPaneElement()
		Refreshes UIA_Browser.MainPaneElement and also returns it
	GetCurrentDocumentElement()
		Returns the current document/content element of the browser. For Mozilla, the tab name which content to get can be specified.
	GetAllText()
		Gets all text from the browser element (Name properties for all child elements)
	GetAllLinks()
		Gets all link elements from the browser (returns an array of elements)
	WaitTitleChange(targetTitle:="", timeOut:=-1)
		Waits the browser title to change to targetTitle (by default just waits for the title to change), timeOut is in milliseconds (default is indefinite waiting)
	WaitPageLoad(targetTitle:="", timeOut:=-1, sleepAfter:=500, titleMatchMode:=3, titleCaseSensitive:=True) 
		Waits the browser page to load to targetTitle, default timeOut is indefinite waiting, sleepAfter additionally sleeps for 200ms after the page has loaded. 
	Back()
		Presses the Back button
	Forward()
		Presses the Forward button
	Reload()
		Presses the Reload button
	Home()
		Presses the Home button if it exists. 
	GetCurrentURL(fromAddressBar:=False)
		Gets the current URL. fromAddressBar=True gets it straight from the URL bar element, which is not a very good method, because the text might be changed by the user and doesn't start with "http(s)://". Default of fromAddressBar=False will cause the real URL to be fetched, but the browser must be visible for it to work (if is not visible, it will be automatically activated).
	SetURL(newUrl, navigateToNewUrl:=False)
		Sets the URL bar to newUrl, optionally also navigates to it if navigateToNewUrl=True
	Navigate(url, targetTitle:="", waitLoadTimeOut:=-1, sleepAfter:=500)
		Navigates to URL and waits page to load
	NewTab()
		Presses the New tab button.
	GetTab(searchPhrase:="", matchMode:=3, caseSensitive:=True)
		Returns a tab element with text of searchPhrase, or if empty then the currently selected tab. matchMode follows SetTitleMatchMode scheme: 1=tab name must must start with tabName; 2=can contain anywhere; 3=exact match; RegEx
	TabExist(searchPhrase:="", matchMode:=3, caseSensitive:=True)
		Checks whether a tab element with text of searchPhrase exists: if a matching tab is found then the element is returned, otherwise 0 is returned. matchMode follows SetTitleMatchMode scheme: 1=tab name must must start with tabName; 2=can contain anywhere; 3=exact match; RegEx
	GetTabs(searchPhrase:="", matchMode:=3, caseSensitive:=True)
		Returns all tab elements with text of searchPhrase, or if empty then all tabs. matchMode follows SetTitleMatchMode scheme: 1=tab name must must start with tabName; 2=can contain anywhere; 3=exact match; RegEx
	GetAllTabNames()
		Gets all the titles of tabs
	SelectTab(tabName, matchMode:=3, caseSensitive:=True) 
		Selects a tab with the text of tabName. matchMode follows SetTitleMatchMode scheme: 1=tab name must must start with tabName; 2=can contain anywhere; 3=exact match; RegEx
	CloseTab(tabElementOrName:="", matchMode:=3, caseSensitive:=True)
		Close tab by either providing the tab element or the name of the tab. If tabElementOrName is left empty, the current tab will be closed.
	IsBrowserVisible()
		Returns True if any of the 4 corners of the browser are visible.
	Send(text)
		Uses ControlSend to send text to the browser.
	GetAlertText()
		Gets the text from an alert box
	CloseAlert()
		Closes an alert box
	JSExecute(js)
		Executes Javascript code using the address bar
		NOTE: In Firefox this is done by default through the console, which is a slow and inefficient method.
			A better way is to create a new bookmark with URL "javascript:%s" and keyword "javascript". This
			allows executing javascript through the address bar with "javascript alert("hello")".
			To make JSExecute use this method, either create UIA_Mozilla with JavascriptExecutionMethod set to "Bookmark",
			or set cUIA.JavascriptExecutionMethod := "Bookmark". Default value is "Console".
	JSReturnThroughClipboard(js)
		Executes Javascript code using the address bar and returns the return value of the code using the clipboard (resetting it back afterwards)
	JSReturnThroughTitle(js, timeOut:=500)
		Executes Javascript code using the address bar and returns the return value of the code using the browsers title (resetting it back afterwards). This might be unreliable, so the clipboard method is recommended instead.
	JSSetTitle(newTitle)
		Uses Javascript through the address bar to change the title of the browser
	JSGetElementPos(selector, useRenderWidgetPos:=False)
		Uses Javascript's querySelector to get a Javascript element and then its position. useRenderWidgetPos=True uses position of the Chrome_RenderWidgetHostHWND1 control to locate the position element relative to the window, otherwise it uses UIA_Browsers CurrentDocumentElement position.
	JSClickElement(selector)
		Uses Javascript's querySelector to get and click a Javascript element
	ControlClickJSElement(selector, WhichButton:="", ClickCount:="", Options:="", useRenderWidgetPos:=False)
		Uses Javascript's querySelector to get a Javascript element and then ControlClicks that position. useRenderWidgetPos=True uses position of the Chrome_RenderWidgetHostHWND1 control to locate the position element relative to the window, otherwise it uses UIA_Browsers CurrentDocumentElement position.
	ClickJSElement(selector, WhichButton:="", ClickCount:=1, DownOrUp:="", Relative:="", useRenderWidgetPos:=False)
		Uses Javascript's querySelector to get a Javascript element and then Clicks that position. useRenderWidgetPos=True uses position of the Chrome_RenderWidgetHostHWND1 control to locate the position element relative to the window, otherwise it uses UIA_Browsers CurrentDocumentElement position.
*/


/*
	If implementing new browser classes, then necessary methods/properties for main browser functions are:

	this.GetCurrentMainPaneElement() -- fetches MainPaneElement, NavigationBarElement, TabBarElement, URLEditElement 
		// this might be necessary to implement for speed reasons, and is automatically called by InitiateUIA method
	this.GetCurrentDocumentElement() -- fetches Document element for the current page // might be necessary to implement
	this.GetCurrentReloadButton()

	this.MainPaneElement -- element that doesn't contain page content: this element includes URL bar, navigation buttons, setting buttons etc
	this.NavigationBarElement -- smallest element (usually a Toolbar element) that contains the URL bar and navigation buttons
	this.TabBarElement -- contains only tabs
	this.URLEditElement -- the URL bar element
	this.ReloadButton
*/

class UIA_Chrome extends UIA_Browser {
	__New(wTitle:="") {
		this.BrowserType := "Chrome"
		this.InitiateUIA(wTitle)
	}
	; Refreshes UIA_Browser.MainPaneElement and returns it
	GetCurrentMainPaneElement() { 
		try this.BrowserElement.FindElement({Type:"Document"}) 
		catch {
			WinActivate this.BrowserId
			WinWaitActive this.BrowserId,,1
		}
		if !this.BrowserElement.WaitElement({Type:"Document"}, 3000)
			throw TargetError("UIA_Browser was unable to find the Document element for browser. Make sure the browser is at least partially visible or active before calling UIA_Browser()", -2)
		Loop 2 {
			 this.URLEditElement := this.BrowserElement.FindFirstWithOptions(this.EditControlCondition, 2, this.BrowserElement)
			try {
				if !this.URLEditElement
					this.URLEditElement := UIA.CreateTreeWalker(this.EditControlCondition).GetLastChildElement(this.BrowserElement)
				this.NavigationBarElement := UIA.CreateTreeWalker(this.ToolbarControlCondition).GetParentElement(this.URLEditElement)
				this.MainPaneElement := UIA.TreeWalkerTrue.GetParentElement(this.NavigationBarElement)
				if !this.NavigationBarElement
					this.NavigationBarElement := this.BrowserElement
				if !this.MainPaneElement
					this.MainPaneElement := this.BrowserElement
				if !(this.TabBarElement := UIA.CreateTreeWalker(this.TabControlCondition).GetPreviousSiblingElement(this.NavigationBarElement))
					this.TabBarElement := this.MainPaneElement
				this.ReloadButton := "", this.ReloadButtonDescription := "", this.ReloadButtonFullDescription := "", this.ReloadButtonName := ""
				Loop 2 {
					try {
						this.ReloadButton := UIA.TreeWalkerTrue.GetNextSiblingElement(UIA.TreeWalkerTrue.GetNextSiblingElement(UIA.TreeWalkerTrue.GetFirstChildElement(this.NavigationBarElement)))
						this.ReloadButtonDescription := this.ReloadButton.LegacyIAccessiblePattern.Description
						this.ReloadButtonName := this.ReloadButton.Name
					}
					if (this.ReloadButtonDescription || this.ReloadButtonName)
						break
					Sleep 200
				}
				return this.MainPaneElement
			} catch {
				WinActivate "ahk_id " this.BrowserId
				WinWaitActive "ahk_id " this.BrowserId,,1
			}
		}
		; If all goes well, this part is not reached
	}
}

class UIA_Edge extends UIA_Browser {
	__New(wTitle:="") {
		this.BrowserType := "Edge"
		this.InitiateUIA(wTitle)
	}

	; Refreshes UIA_Browser.MainPaneElement and returns it
	GetCurrentMainPaneElement() { 
		local k, v
		try this.BrowserElement.FindElement({Type:"Document"}) 
		catch {
			WinActivate this.BrowserId
			WinWaitActive this.BrowserId,,1
		}
		if !this.BrowserElement.WaitElement({Type:"Document"}, 3000)
			throw TargetError("UIA_Browser was unable to find the Document element for browser. Make sure the browser is at least partially visible or active before calling UIA_Browser()", -2)
		Loop 2 {
			try {
				if !(this.URLEditElement := this.BrowserElement.FindFirst(this.EditControlCondition)) {
					this.ToolbarElements := this.BrowserElement.FindAll(this.ToolbarControlCondition), topCoord := 10000000
					for k, v in this.ToolbarElements {
						if ((bT := v.BoundingRectangle.t) && (bt < topCoord))
							topCoord := bT, this.NavigationBarElement := v
					}
					this.URLEditElement := this.NavigationBarElement.FindFirst(this.EditControlCondition)
					if this.URLEditElement.GetChildren().Length
						this.URLEditElement := (el := this.URLEditElement.FindFirst(this.EditControlCondition)) ? el : this.URLEditElement
				} Else {
					this.NavigationBarElement := UIA.CreateTreeWalker(this.ToolbarControlCondition).GetParentElement(this.URLEditElement)
				}
				this.MainPaneElement := UIA.TreeWalkerTrue.GetParentElement(this.NavigationBarElement)
				if !this.NavigationBarElement
					this.NavigationBarElement := this.BrowserElement
				if !this.MainPaneElement
					this.MainPaneElement := this.BrowserElement
				if !(this.TabBarElement := UIA.CreateTreeWalker(this.TabControlCondition).GetPreviousSiblingElement(this.NavigationBarElement))
					this.TabBarElement := this.MainPaneElement
				this.ReloadButton := "", this.ReloadButtonDescription := "", this.ReloadButtonFullDescription := "", this.ReloadButtonName := ""
				Loop 2 {
					try {
						this.ReloadButton := UIA.TreeWalkerTrue.GetNextSiblingElement(UIA.TreeWalkerTrue.GetNextSiblingElement(UIA.TreeWalkerTrue.GetFirstChildElement(this.NavigationBarElement)))
						this.ReloadButtonFullDescription := this.ReloadButton.FullDescription
						this.ReloadButtonName := this.ReloadButton.Name
					}
					if (this.ReloadButtonDescription || this.ReloadButtonName)
						break
					Sleep 200
				}
				return this.MainPaneElement
			} catch {
				WinActivate "ahk_id " this.BrowserId
				WinWaitActive "ahk_id " this.BrowserId,,1
			}
		}
		; If all goes well, this part is not reached
	}
}

class UIA_Mozilla extends UIA_Browser {
	__New(wTitle:="", javascriptExecutionMethod:="Console") {
		this.JavascriptExecutionMethod := javascriptExecutionMethod
		this.BrowserType := "Mozilla"
		this.InitiateUIA(wTitle)
	}
	; Refreshes UIA_Browser.MainPaneElement and returns it
	GetCurrentMainPaneElement() { 
		try this.BrowserElement.FindElement({AutomationId:"panel", mm:2},2)
		catch {
			WinActivate this.BrowserId
			WinWaitActive this.BrowserId,,1
		}
		Loop 2 {
			try {
				this.TabBarElement := this.ToolbarTreeWalker.GetNextSiblingElement(this.ToolbarTreeWalker.GetFirstChildElement(this.BrowserElement))
				this.NavigationBarElement := this.ToolbarTreeWalker.GetNextSiblingElement(this.TabBarElement)
				this.URLEditElement := this.NavigationBarElement.FindFirst(this.EditControlCondition)
				this.MainPaneElement := UIA.TreeWalkerTrue.GetParentElement(this.NavigationBarElement)
				if !this.NavigationBarElement
					this.NavigationBarElement := this.BrowserElement
				if !this.MainPaneElement
					this.MainPaneElement := this.BrowserElement
				this.ReloadButton := "", this.ReloadButtonDescription := "", this.ReloadButtonFullDescription := "", this.ReloadButtonName := ""
				Loop 2 {
					try {
						this.ReloadButton := UIA.TreeWalkerTrue.GetNextSiblingElement(UIA.TreeWalkerTrue.GetNextSiblingElement(UIA.TreeWalkerTrue.GetFirstChildElement(this.NavigationBarElement)))
						this.ReloadButtonFullDescription := this.ReloadButton.FullDescription
						this.ReloadButtonName := this.ReloadButton.Name
					}
					if (this.ReloadButtonDescription || this.ReloadButtonName)
						break
					Sleep 200
				}
				return this.MainPaneElement
			} catch {
				WinActivate this.BrowserId
				WinWaitActive this.BrowserId,,1
			}
		}
		; If all goes well, this part is not reached
	}

	; Returns the current document/content element of the browser
	GetCurrentDocumentElement() {
		Loop 2 {
			try {
				this.DocumentPanelElement := this.BrowserElement.FindElement({Type:"Custom", IsOffscreen:0},2)
				return UIA.TreeWalkerTrue.GetFirstChildElement(UIA.TreeWalkerTrue.GetFirstChildElement(this.DocumentPanelElement))
			} catch TargetError {
				WinActivate this.BrowserId
				WinWaitActive this.BrowserId,,1
			}
		}
	}

	; Presses the New tab button. 
	NewTab() { 
		this.TabBarElement.FindFirst(this.ButtonControlCondition,4).Click()
	}

	; Sets the URL bar to newUrl, optionally also navigates to it if navigateToNewUrl=True
	SetURL(newUrl, navigateToNewUrl := False) { 
		this.URLEditElement.SetFocus()
		this.URLEditElement.Value := newUrl " "
		if (navigateToNewUrl&&InStr(this.URLEditElement.Value, newUrl)) {
			ControlSend "{LCtrl up}{LAlt up}{LShift up}{RCtrl up}{RAlt up}{RShift up}{LCtrl down}{Enter}{LCtrl up}", , this.BrowserId
		}
	}

	JSExecute(js) {
		if this.JavascriptExecutionMethod = "Bookmark" {
			this.SetURL("javascript " js, True)
			return
		}
		ControlSend "{LCtrl up}{LAlt up}{LShift up}{RCtrl up}{RAlt up}{RShift up}", , this.BrowserId
		ControlSend "{ctrl down}{shift down}k{ctrl up}{shift up}", , this.BrowserId
		if !this.BrowserElement.WaitElement({Name:"Switch to multi-line editor mode (Ctrl + B)", Type:"Button"},5000)
			return
		ClipSave := ClipboardAll()
		A_Clipboard := js
		WinActivate this.BrowserId
		WinWaitActive this.BrowserId
		ControlSend "{ctrl down}v{ctrl up}", , this.BrowserId
		Sleep 20
		ControlSend "{ctrl down}{enter}{ctrl up}", , this.BrowserId
		sleep 40
		ControlSend "{ctrl down}{shift down}i{ctrl up}{shift up}", , this.BrowserId
		A_Clipboard := ClipSave
	}

	; Gets text from an alert-box
	GetAlertText(closeAlert:=True, timeOut:=3000) {
		this.GetCurrentDocumentElement()
		local startTime := A_TickCount, text := ""
		if !(alertEl := UIA.TreeWalkerTrue.GetNextSiblingElement(UIA.TreeWalkerTrue.GetFirstChildElement(this.DocumentPanelElement)))
			return
		
		while ((A_tickCount - startTime) < timeOut) {
			try {
				dialogEl := alertEl.FindElement({AutomationId:"commonDialogWindow"})
				OKBut := dialogEl.FindFirst(this.ButtonControlCondition)
				break
			} catch
				Sleep 100
		}
		try text := dialogEl.FindFirst(this.TextControlCondition).Name
		if closeAlert
			try OKBut.Click()
		return text
	}

	CloseAlert() {
		this.GetCurrentDocumentElement()
		try UIA.TreeWalkerTrue.GetNextSiblingElement(UIA.TreeWalkerTrue.GetFirstChildElement(this.DocumentPanelElement)).FindElement({AutomationId:"commonDialogWindow"}).FindFirst(this.ButtonControlCondition).Click()
	}

	; Close tab by either providing the tab element or the name of the tab. If tabElementOrName is left empty, the current tab will be closed.
	CloseTab(tabElementOrName:="", matchMode:=3, caseSensitive:=True) { 
		if (tabElementOrName != "") {
			if IsObject(tabElementOrName) {
				if (tabElementOrName.Type == UIA.Type.TabItem)
					tabElementOrName.Click()
			} else {
				try this.TabBarElement.FindElement({Name:tabElementOrName, Type:"TabItem", mm:matchMode, cs:caseSensitive}).Click()
			}
		}
		ControlSend "{LCtrl up}{LAlt up}{LShift up}{RCtrl up}{RAlt up}{RShift up}", , this.BrowserId
		ControlSend "{Ctrl down}w{Ctrl up}", , this.BrowserId
	}
}

class UIA_Browser {
	InitiateUIA(wTitle:="") {
		this.BrowserId := WinExist(wTitle)
		if !this.BrowserId
			throw TargetError("UIA_Browser: failed to find the browser!", -1)
		this.TextControlCondition := UIA.CreatePropertyCondition(UIA.Property.Type, UIA.Type.Text)
		this.DocumentControlCondition := UIA.CreatePropertyCondition(UIA.Property.Type, UIA.Type.Document)
		this.ButtonControlCondition := UIA.CreatePropertyCondition(UIA.Property.Type, UIA.Type.Button)
		this.EditControlCondition := UIA.CreatePropertyCondition(UIA.Property.Type, UIA.Type.Edit)
		this.ToolbarControlCondition := UIA.CreatePropertyCondition(UIA.Property.Type, UIA.Type.ToolBar)
		this.TabControlCondition := UIA.CreatePropertyCondition(UIA.Property.Type, UIA.Type.Tab)
		this.ToolbarTreeWalker := UIA.CreateTreeWalker(this.ToolbarControlCondition)
		this.BrowserElement := UIA.ElementFromHandle(this.BrowserId)
		this.GetCurrentMainPaneElement()
	}
	; Initiates UIA and hooks to the browser window specified with wTitle. 
	__New(wTitle:="") { 
		this.BrowserId := WinExist(wTitle)
		if !this.BrowserId
			throw TargetError("UIA_Browser: failed to find the browser!", -1)
		wExe := WinGetProcessName("ahk_id" this.BrowserId)
		wClass := WinGetClass("ahk_id" this.BrowserId)
		this.BrowserType := (wExe = "chrome.exe") ? "Chrome" : (wExe = "msedge.exe") ? "Edge" : InStr(wClass, "Mozilla") ? "Mozilla" : "Unknown"
		if (this.BrowserType != "Unknown") {
			this.base := UIA_%(this.BrowserType)%.Prototype
			this.__New(wTitle)
		} else 
			this.InitiateUIA(wTitle)
	}
	
	__Get(member, params) {
		if this.HasOwnProp("BrowserElement") {
			try return this.BrowserElement.%member%
			catch PropertyError {
			} catch Any as err
				throw %Type(err)%(err.Message, -1, err.Extra)
		}
		try return UIA.%member%
		catch PropertyError
			throw PropertyError("This class does not contain property `"" member "`"", -1)
		catch Any as err
			throw %Type(err)%(err.Message, -1, err.Extra)
	}
	
	__Call(member, params) {
		if this.HasOwnProp("BrowserElement") {
			try return this.BrowserElement.%member%(params*)
			catch MethodError {
			} catch Any as err
				throw %Type(err)%(err.Message, -1, err.Extra)
		}
		try return UIA.%member%(params*)
		catch MethodError
			throw MethodError("This class does not contain method `"" member "`"", -1)
		catch Any as err
			throw %Type(err)%(err.Message, -1, err.Extra)
	}

	__Set(member, params, value) {
		if this.HasOwnProp("BrowserElement")
			if this.BrowserElement.HasOwnProp(member)
				this.BrowserElement.%member% := Value
		if UIA.HasOwnProp(member)
			UIA.%member% := Value
		this.DefineProp(member, {Value:value})
	}

    __Item[params*] {
        get => this.BrowserElement[params*]
	}
	
	; Refreshes UIA_Browser.MainPaneElement and returns it
	GetCurrentMainPaneElement() { 
		try this.BrowserElement.FindElement({Type:"Document"}) 
		catch {
			WinActivate this.BrowserId
			WinWaitActive this.BrowserId,,1
		}
		if !this.BrowserElement.WaitElement({Type:"Document"}, 3000)
			throw TargetError("UIA_Browser was unable to find the Document element for browser. Make sure the browser is at least partially visible or active before calling UIA_Browser()", -2)

		; Finding the correct Toolbar ends up to be quite tricky. 
		; In Chrome the toolbar element is located in the tree after the content element, 
		; so if the content contains a toolbar then that will be returned. 
		; Two workarounds I can think of: either look for the Toolbar by name ("Address and search bar" 
		; both in Chrome and edge), or by location (it must be the topmost toolbar). I opted for a 
		; combination of two, so if finding by name fails, all toolbar elements are evaluated.
		Loop 2 {
			try this.URLEditElement := (this.BrowserType = "Chrome") ? this.BrowserElement.FindFirstWithOptions(this.EditControlCondition, 2, this.BrowserElement) : this.BrowserElement.FindFirst(this.EditControlCondition)
			try {
				if (this.BrowserType = "Chrome") && !this.URLEditElement
					this.URLEditElement := UIA.CreateTreeWalker(this.EditControlCondition).GetLastChildElement(this.BrowserElement)
				this.NavigationBarElement := UIA.CreateTreeWalker(this.ToolbarControlCondition).GetParentElement(this.URLEditElement)
				this.MainPaneElement := UIA.TreeWalkerTrue.GetParentElement(this.NavigationBarElement)
				if !this.NavigationBarElement
					this.NavigationBarElement := this.BrowserElement
				if !this.MainPaneElement
					this.MainPaneElement := this.BrowserElement
				if !(this.TabBarElement := UIA.CreateTreeWalker(this.TabControlCondition).GetPreviousSiblingElement(this.NavigationBarElement))
					this.TabBarElement := this.MainPaneElement
				this.GetCurrentReloadButton()
				this.ReloadButton := "", this.ReloadButtonDescription := "", this.ReloadButtonFullDescription := "", this.ReloadButtonName := ""
				Loop 2 {
					try {
						this.ReloadButtonDescription := this.ReloadButton.LegacyIAccessiblePattern.Description
						this.ReloadButtonFullDescription := this.ReloadButton.FullDescription
						this.ReloadButtonName := this.ReloadButton.Name
					}
					if (this.ReloadButtonDescription || this.ReloadButtonName)
						break
					Sleep 200
				}

				return this.MainPaneElement
			} catch {
				WinActivate "ahk_id " this.BrowserId
				WinWaitActive "ahk_id " this.BrowserId,,1
			}
		}
		; If all goes well, this part is not reached
	}
	
	; Returns the current document/content element of the browser
	GetCurrentDocumentElement() { 
		if (this.BrowserType = "Mozilla")
			return (this.CurrentDocumentElement := this.BrowserElement.FindElement({Name:this.GetTab().Name, Type:"Document"}))
		return (this.CurrentDocumentElement := this.BrowserElement.FindFirst(this.DocumentControlCondition))
	}

	GetCurrentReloadButton() {
		try {
			if this.ReloadButton && this.ReloadButton.Name
				return this.ReloadButton
		}
		ButtonWalker := UIA.CreateTreeWalker(this.ButtonControlCondition)
		this.ReloadButton := ButtonWalker.GetNextSiblingElement(ButtonWalker.GetNextSiblingElement(ButtonWalker.GetFirstChildElement(this.NavigationBarElement)))
		return this.ReloadButton
	}
	
	; Uses Javascript to set the title of the browser.
	JSSetTitle(newTitle) {
		this.JSExecute("document.title=`"" newTitle "`"; void(0);")
	}
	
	JSExecute(js) {
		this.SetURL("javascript:" js, True)
	}
	
	JSAlert(js, closeAlert:=True, timeOut:=3000) {
		this.JSExecute("alert(" js ");")
		return this.GetAlertText(closeAlert, timeOut)
	}
	
	; Executes Javascript code through the address bar and returns the return value through the clipboard.
	JSReturnThroughClipboard(js) {
		saveClip := ClipboardAll()
		A_Clipboard := ""
		this.JSExecute("copyToClipboard(" js ");function copyToClipboard(text) {const elem = document.createElement('textarea');elem.value = text;document.body.appendChild(elem);elem.select();document.execCommand('copy');document.body.removeChild(elem);}")
		ClipWait 2
		returnText := A_Clipboard
		A_Clipboard := saveClip
		return returnText
	}
	
	; Executes Javascript code through the address bar and returns the return value through the browser windows title.
	JSReturnThroughTitle(js, timeOut:=500) {
		this.JSExecute("origTitle=document.title;document.title=(" js ");void(0);setTimeout(function() {document.title=origTitle;void(0);}, " timeOut ")")
		local startTime := A_TickCount, origTitle := WinGetTitle("ahk_id " this.BrowserId), newTitle
		Loop {
			newTitle := WinGetTitle("ahk_id " this.BrowserId)
			Sleep 40
		} Until ((origTitle != newTitle) || (A_TickCount - startTime > timeOut))
		return (origTitle == newTitle) ? "" : RegexReplace(newTitle, "(?: - Personal)? - [^-]+$")
	}
	
	; Uses Javascript's querySelector to get a Javascript element and then its position. useRenderWidgetPos=True uses position of the Chrome_RenderWidgetHostHWND1 control to locate the position element relative to the window, otherwise it uses UIA_Browsers CurrentDocumentElement position.
    JSGetElementPos(selector, useRenderWidgetPos:=False) { ; based on code by AHK Forums user william_ahk
        js := "
        (LTrim
			(() => {
				let bounds = document.querySelector("%selector%").getBoundingClientRect().toJSON();
				let zoom = window.devicePixelRatio.toFixed(2);
				for (const key in bounds) {
					bounds[key] = bounds[key] * zoom;
				}
				return JSON.stringify(bounds);
			})()
        )"
        bounds_str := this.JSReturnThroughClipboard(js)
        RegexMatch(bounds_str, "`"x`":(\d+).?\d*?,`"y`":(\d+).?\d*?,`"width`":(\d+).?\d*?,`"height`":(\d+).?\d*?", &size)
		if useRenderWidgetPos {
			ControlGetPos &win_x, &win_y, &win_w, &win_h, "Chrome_RenderWidgetHostHWND1", this.BrowserId
			return {x:size[1]+win_x,y:size[2]+win_y,w:size[3],h:size[4]}
		} else {
			br := this.GetCurrentDocumentElement().GetPos("window")
			return {x:size[1]+br.x,y:size[2]+br.y,w:size[3],h:size[4]}
		}
    }
	
	; Uses Javascript's querySelector to get and click a Javascript element. Compared with ClickJSElement method, this method has the advantage of skipping the need to wait for a return value from the clipboard, but it might be more unreliable (some elements might not support Javascript's "click()" properly).
	JSClickElement(selector) {
        this.JSExecute("document.querySelector(`"" selector "`").click();")
	}
    
	; Uses Javascript's querySelector to get a Javascript element and then ControlClicks that position. useRenderWidgetPos=True uses position of the Chrome_RenderWidgetHostHWND1 control to locate the position element relative to the window, otherwise it uses UIA_Browsers CurrentDocumentElement position.
    ControlClickJSElement(selector, WhichButton?, ClickCount?, Options?, useRenderWidgetPos:=False) {
        bounds := this.JSGetElementPos(selector, useRenderWidgetPos)
        ControlClick("X" (bounds.x + bounds.w // 2) " Y" (bounds.y + bounds.h // 2), this.browserId,, WhichButton?, ClickCount?, Options?)
    }

	; Uses Javascript's querySelector to get a Javascript element and then Clicks that position. useRenderWidgetPos=True uses position of the Chrome_RenderWidgetHostHWND1 control to locate the position element relative to the window, otherwise it uses UIA_Browsers CurrentDocumentElement position.
    ClickJSElement(selector, WhichButton:="", ClickCount:=1, DownOrUp:="", Relative:="", useRenderWidgetPos:=False) {
        bounds := this.JSGetElementPos(selector, useRenderWidgetPos)
        Click((bounds.x + bounds.w / 2) " " (bounds.y + bounds.h / 2) " " WhichButton (ClickCount ? " " ClickCount : "") (DownOrUp ? " " DownOrUp : "") (Relative ? " " Relative : ""))
    }
	
	; Gets text from an alert-box created with for example javascript:alert('message')
	GetAlertText(closeAlert:=True, timeOut:=3000) {
		local startTime := A_TickCount, text := ""
		if !this.HasOwnProp("DialogTreeWalker")
			this.DialogTreeWalker := UIA.CreateTreeWalker(UIA.CreateOrCondition(UIA.CreatePropertyCondition(UIA.Property.Type, UIA.Type.Custom), UIA.CreatePropertyCondition(UIA.Property.Type, UIA.Type.Window)))
		startTime := A_TickCount
		while ((A_tickCount - startTime) < timeOut) {
			try {
				if IsObject(dialogEl := this.DialogTreeWalker.GetLastChildElement(this.BrowserElement)) && IsObject(OKBut := dialogEl.FindFirst(this.ButtonControlCondition))
					break
			}
			Sleep 100
		}
		try
			text := this.BrowserType = "Edge" ? dialogEl.FindFirstWithOptions(this.TextControlCondition, 2, dialogEl).Name : dialogEl.FindFirst(this.TextControlCondition).Name
		if closeAlert {
			Sleep 500
			try OKBut.Click()
		}
		return text
	}
	
	CloseAlert() {
		if !this.HasOwnProp("DialogTreeWalker")
			this.DialogTreeWalker := UIA.CreateTreeWalker(UIA.CreateOrCondition(UIA.CreatePropertyCondition(UIA.Property.Type, UIA.Type.Custom), UIA.CreatePropertyCondition(UIA.Property.Type, UIA.Type.Window)))
		try {
			dialogEl := this.DialogTreeWalker.GetLastChildElement(this.BrowserElement)
			OKBut := dialogEl.FindFirst(this.ButtonControlCondition)
			OKBut.Click()
		}
	}
	
	; Gets all text from the browser element (Name properties for all Text elements)
	GetAllText() { 
		local TextArray, Text, k, v
		if !this.IsBrowserVisible()
			WinActivate this.BrowserId
			
		TextArray := this.BrowserElement.FindAll(this.TextControlCondition)
		Text := ""
		for k, v in TextArray
			Text .= v.Name "`n"
		return Text
	}
	; Gets all link elements from the browser
	GetAllLinks() {
		static LinkCondition := UIA.CreatePropertyCondition(UIA.Property.Type, UIA.Type.Hyperlink)
		if !this.IsBrowserVisible()
			WinActivate this.BrowserId			
		return this.BrowserElement.FindAll(LinkCondition)
	}
	
	__CompareTitles(compareTitle, winTitle, matchMode:="", caseSensitive:=True) {
		if !matchMode
			matchMode := A_TitleMatchMode
		if (matchMode == 1) {
			if (caseSensitive ? (SubStr(winTitle, 1, StrLen(compareTitle)) == compareTitle) : (SubStr(winTitle, 1, StrLen(compareTitle)) = compareTitle))
				return 1
		} else if (matchMode == 2) {
			if InStr(winTitle, compareTitle, caseSensitive)
				return 1
		} else if (matchMode == 3) {
			if (caseSensitive ? (compareTitle == winTitle) : (compareTitle = winTitle))
				return 1
		} else if (matchMode = "RegEx") {
			if RegexMatch(winTitle, compareTitle)
				return 1
		}
		return 0
	}
	
	; Waits the browser title to change to targetTitle (by default just waits for the title to change), timeOut is in milliseconds (default is indefinite waiting)
	WaitTitleChange(targetTitle:="", timeOut:=-1) { 
		local origTitle := WinGetTitle("ahk_id" this.BrowserId), startTime := A_TickCount, newTitle := origTitle
		while ((((A_TickCount - startTime) < timeOut) || (timeOut = -1)) && (targetTitle ? !this.__CompareTitles(targetTitle, newTitle) : (origTitle == newTitle))) {
			Sleep 200
			newTitle := WinGetTitle("A")
		}
	}
	
	; Waits the browser page to load to targetTitle, default timeOut is indefinite waiting, sleepAfter additionally sleeps for 200ms after the page has loaded. 
	WaitPageLoad(targetTitle:="", timeOut:=-1, sleepAfter:=500, titleMatchMode:="", titleCaseSensitive:=False) {
		local legacyPattern := "", startTime := A_TickCount, wTitle, ReloadButtonName := "", ReloadButtonDescription := "", ReloadButtonFullDescription := "" 
		Sleep 200 ; Give some time for the Reload button to change after navigating
		if this.ReloadButtonDescription
			try legacyPattern := this.ReloadButton.LegacyIAccessiblePattern
		while ((A_TickCount - startTime) < timeOut) || (timeOut = -1) {
			if this.BrowserType = "Mozilla"
				this.GetCurrentReloadButton()
			try ReloadButtonName := this.ReloadButton.Name
			try ReloadButtonDescription := legacyPattern.Description
			try ReloadButtonFullDescription := this.ReloadButton.FullDescription
			if ((this.ReloadButtonName ? InStr(ReloadButtonName, this.ReloadButtonName) : 1) 
			   && (this.ReloadButtonDescription && legacyPattern ? InStr(ReloadButtonDescription, this.ReloadButtonDescription) : 1)
			   && (this.ReloadButtonFullDescription ? InStr(ReloadButtonFullDescription, this.ReloadButtonFullDescription) : 1)) {
				if targetTitle {
					wTitle := WinGetTitle(this.BrowserId)
					if this.__CompareTitles(targetTitle, wTitle, titleMatchMode, titleCaseSensitive)
						break
				} else
					break
			}
			Sleep 40
		}
		if ((A_TickCount - startTime) < timeOut) || (timeOut = -1)
			Sleep sleepAfter
	}
	
	; Presses the Back button
	Back() { 
		UIA.TreeWalkerTrue.GetFirstChildElement(this.NavigationBarElement).Click()
	}
	
	; Presses the Forward button
	Forward() { 
		UIA.TreeWalkerTrue.GetNextSiblingElement(UIA.TreeWalkerTrue.GetFirstChildElement(this.NavigationBarElement)).Click()
	}

	; Presses the Reload button
	Reload() { 
		this.GetCurrentReloadButton().Click()
	}

	; Presses the Home button if it exists.
	Home() { 
		if homeBut := UIA.TreeWalkerTrue.GetNextSiblingElement(this.ReloadButton)
			return homeBut.Click()
		;NameCondition := UIA.CreatePropertyCondition(UIA.NamePropertyId, this.CustomNames.HomeButtonName ? this.CustomNames.HomeButtonName : butName)
		;this.NavigationBarElement.FindFirst(UIA.CreateAndCondition(NameCondition, this.ButtonControlCondition)).Click()
	}
	
	; Gets the current URL. fromAddressBar=True gets it straight from the URL bar element, which is not a very good method, because the text might be changed by the user and doesn't start with "http(s)://". Default of fromAddressBar=False will cause the real URL to be fetched, but the browser must be visible for it to work (if is not visible, it will be automatically activated).
	GetCurrentURL(fromAddressBar:=False) { 
		if fromAddressBar {
			URL := this.URLEditElement.Value
			return URL ? (RegexMatch(URL, "^https?:\/\/") ? URL : "https://" URL) : ""
		} else {
			; This can be used in Chrome and Edge, but works only if the window is active
			if (!this.IsBrowserVisible() && (this.BrowserType != "Mozilla"))
				WinActivate this.BrowserId
			return this.GetCurrentDocumentElement().Value
		}
	}
	
	; Sets the URL bar to newUrl, optionally also navigates to it if navigateToNewUrl=True
	SetURL(newUrl, navigateToNewUrl := False) { 
		this.URLEditElement.SetFocus()
		this.URLEditElement.ValuePattern.SetValue(newUrl " ")
		if !InStr(this.URLEditElement.Value, newUrl) {
			legacyPattern := this.URLEditElement.LegacyIAccessiblePattern
			legacyPattern.SetValue(newUrl " ")
			legacyPattern.Select()
		}
		if (navigateToNewUrl&&InStr(this.URLEditElement.Value, newUrl)) {
			ControlSend "{LCtrl up}{LAlt up}{LShift up}{RCtrl up}{RAlt up}{RShift up}{Enter}", , "ahk_id" this.BrowserId ; Or would it be better to use BlockInput instead of releasing modifier keys?
		}
	}

	; Navigates to URL and waits page to load
	Navigate(url, targetTitle:="", waitLoadTimeOut:=-1, sleepAfter:=500) {
		this.SetURL(url, True)
		this.WaitPageLoad(targetTitle,waitLoadTimeOut,sleepAfter)
	}
	
	; Presses the New tab button. The button name might differ if the browser language is not set to English and can be specified with butName
	NewTab() {
		local el
		try el := this.TabBarElement.FindFirstWithOptions(this.ButtonControlCondition,2,this.TabBarElement)
		if IsSet(el) && el
			el.Click()
		else {
			UIA.CreateTreeWalker(this.ButtonControlCondition).GetLastChildElement(this.TabBarElement).Click()
		}
	}
	
	GetAllTabs() { 
		return this.TabBarElement.FindAll(UIA.CreatePropertyCondition(UIA.Property.Type, UIA.Type.TabItem))
	}
	; Gets all tab elements matching searchPhrase, matchMode and caseSensitive
	; If searchPhrase is omitted then all tabs will be returned
	GetTabs(searchPhrase:="", matchMode:=3, caseSensitive:=True) {
		return (searchPhrase == "") ? this.TabBarElement.FindAll(UIA.CreatePropertyCondition(UIA.Property.Type, UIA.Type.TabItem)) : this.TabBarElement.FindElements({Name:searchPhrase, Type:"TabItem", mm:matchMode, cs:caseSensitive})
	}

	; Gets all the titles of tabs
	GetAllTabNames() { 
		local names := [], k, v
		for k, v in this.GetTabs() {
			names.Push(v.Name)
		}
		return names
	}
	
	; Returns a tab element with text of searchPhrase, or if empty then the currently selected tab. matchMode follows SetTitleMatchMode scheme: 1=tab name must must start with tabName; 2=can contain anywhere; 3=exact match; RegEx
	GetTab(searchPhrase:="", matchMode:=3, caseSensitive:=True) { 
		return (searchPhrase == "") ? this.TabBarElement.FindElement({Type:"TabItem", SelectionItemIsSelected:1}) : this.TabBarElement.FindElement({Name:searchPhrase, Type:"TabItem", mm:matchMode, cs:caseSensitive})
	}
	; Checks whether a tab element with text of searchPhrase exists: if a matching tab is found then the element is returned, otherwise 0 is returned. matchMode follows SetTitleMatchMode scheme: 1=tab name must must start with tabName; 2=can contain anywhere; 3=exact match; RegEx
	TabExist(searchPhrase:="", matchMode:=3, caseSensitive:=True) {
		try return this.GetTab(searchPhrase, matchMode, caseSensitive)
		return 0
	}
	
	; Selects a tab with the text of tabName. matchMode follows SetTitleMatchMode scheme: 1=tab name must must start with tabName; 2=can contain anywhere; 3=exact match; RegEx
	SelectTab(tabName, matchMode:=3, caseSensitive:=True) { 
		local selectedTab
		try {
			selectedTab := this.TabBarElement.FindElement({Name:tabName, Type:"TabItem", mm:matchMode, cs:caseSensitive})
			selectedTab.Click()
		} catch TargetError
			throw TargetError("Tab with name " tabName " was not found (MatchMode: " matchMode ", CaseSensitive: " caseSensitive ")")
		return selectedTab
	}
	
	; Close tab by either providing the tab element or the name of the tab. If tabElementOrName is left empty, the current tab will be closed.
	CloseTab(tabElementOrName:="", matchMode:=3, caseSensitive:=True) { 
		if IsObject(tabElementOrName) {
			if (tabElementOrName.Type == UIA.Type.TabItem)
				try UIA.TreeWalkerTrue.GetLastChildElement(tabElementOrName).Click()
		} else {
			if (tabElementOrName == "") {
				try UIA.TreeWalkerTrue.GetLastChildElement(this.GetTab()).Click()
			} else {
				try {
					targetTab := this.TabBarElement.FindElement({Name:tabElementOrName, Type:"TabItem", mm:matchMode, cs:caseSensitive})
					UIA.TreeWalkerTrue.GetLastChildElement().Click(targetTab)
				} catch
					throw TargetError("Tab with name " tabElementOrName " was not found (MatchMode: " matchMode ", CaseSensitive: " caseSensitive ")")
			}
		}
	}
	
	; Returns True if any of window 4 corners are visible
	IsBrowserVisible() { 
		local X, Y, W, H
		WinGetPos &X, &Y, &W, &H, "ahk_id" this.BrowserId
		if ((this.BrowserId == this.WindowFromPoint(X, Y)) || (this.BrowserId == this.WindowFromPoint(X, Y+H-1)) || (this.BrowserId == this.WindowFromPoint(X+W-1, Y)) || (this.BrowserId == this.WindowFromPoint(X+W-1, Y+H-1)))
			return True
		return False
	}

	Send(text) {
		ControlSend "{LCtrl up}{LAlt up}{LShift up}{RCtrl up}{RAlt up}{RShift up}{Enter}" , this.BrowserId
		ControlSend text, , this.BrowserId
	}
	
	WindowFromPoint(X, Y) { ; by SKAN and Linear Spoon
		return DllCall( "GetAncestor", "UInt"
			   , DllCall( "WindowFromPoint", "UInt64", (X & 0xFFFFFFFF) | (Y << 32))
			   , "UInt", 2 ) ; GA_ROOT
	}

	PrintArray(arr) {
		local ret := "", k, v
		for k, v in arr
			ret .= "Key: " k " Value: " (HasMethod(v)? v.name:IsObject(v)?this.PrintArray(v):v) "`n"
		return ret
	}
}