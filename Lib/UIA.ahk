/*
	Introduction & credits
	This library implements Microsoft's UI Automation framework. More information is here: https://docs.microsoft.com/en-us/windows/win32/winauto/entry-uiauto-win32
	Authors: Descolada, thqby, neptercn (v1), jethrow (AHK v1 UIA library)

	A lot of modifications have been added to the original UIA framework, such as custom methods for elements (eg element.Click())
*/

/* 
	Usage
	UIA needs to be initialized with UIA_Interface() function, which returns a UIA_Interface object:
	UIA := UIA_Interface()
	After calling this function, all UIA_Interface class properties and methods can be accessed through it. 
	In addition some extra variables are initialized: 
		IUIAutomationVersion contains the version number of IUIAutomation interface
		TrueCondition contains a TrueCondition
	Note that a new UIA object can't be created with the "new" keyword. 
	
	UIAutomation constants and enumerations are available from the UIA_Enum class (see a more thorough description at the class header).
	Microsoft documentation for constants and enumerations:
		UI Automation Constants: https://docs.microsoft.com/en-us/windows/win32/winauto/uiauto-entry-constants
		UI Automation Enumerations: https://docs.microsoft.com/en-us/windows/win32/winauto/uiauto-entry-enumerations
	
	For more information, see the AHK Forums post on UIAutomation: https://www.autohotkey.com/boards/viewtopic.php?f=6&t=104999
*/


/* 	
	Questions:
	- if method returns a SafeArray, should we return a Wrapped SafeArray, Raw SafeArray, or AHK Array. Currently we return wrapped AHK arrays for SafeArrays. Although SafeArrays are more convenient to loop over, this causes more confusion in users who are not familiar with SafeArrays (questions such as why are they 0-indexed not 1-indexed, why doesnt for k, v in SafeArray work properly etc). 
	- on UIA Interface conversion methods, how should the data be returned? wrapped/extracted or raw? should raw data be a ByRef param?
	- do variants need cleared? what about SysAllocString BSTRs? As per Microsoft documentation (https://docs.microsoft.com/en-us/cpp/atl-mfc-shared/allocating-and-releasing-memory-for-a-bstr?view=msvc-170), when we pass a BSTR into IUIAutomation, then IUIAutomation should take care of freeing it. But when we receive a UIA_Variant and use UIA_VariantData, then we should clear the BSTR.
	- ObjRelease: if IUIA returns an interface then it automatically increases the ref count for the object it inherits from, and when released decreases it. So do all returned objects (UIA_Element, UIA_Pattern, UIA_TextRange) need to be released? Currently we release these objects as well, but jethrow's version didn't. 
	- do RECT structs need destroyed?
	- if returning wrapped data & raw is ByRef, will the wrapped data being released destroy the raw data?
	- returning variant data other than vt=3|8|9|13|0x2000
	- Cached Members?
	- UIA Element existance - dependent on window being visible (non minimized), and also sometimes Elements are lazily generated (eg Microsoft Teams, when a meeting is started then the toolbar buttons (eg Mute, react) aren't visible to UIA, but hovering over them with the cursor or calling ElementFromPoint causes Teams to generate and make them visible to UIA.
	- better way of supporting differing versions of IUIAutomation (version 2, 3, 4)
	- Get methods vs property getter: currently we use properties when the item stores data, fetching the data is "cheap" and when it doesn't have side-effects, and in computationally expensive cases use Get...(). 
	- should ElementFromHandle etc methods have activateChromiumAccessibility set to True or False? Currently is True, because Chromium apps are very common, and checking whether its on should be relatively fast.
*/
global IUIAutomationMaxVersion := 8

class UIA {
    static __New() {
        static __IID := {
            IUIAutomation:"{30cbe57d-d9d0-452a-ab13-7ac5ac4825ee}",
            IUIAutomation2:"{34723aff-0c9d-49d0-9896-7ab52df8cd8a}",
            IUIAutomation3:"{73d768da-9b51-4b89-936e-c209290973e7}",
            IUIAutomation4:"{1189c02a-05f8-4319-8e21-e817e3db2860}",
            IUIAutomation5:"{25f700c8-d816-4057-a9dc-3cbdee77e256}",
            IUIAutomation6:"{aae072da-29e3-413d-87a7-192dbf81ed10}",
            IUIAutomation7:"{29de312e-83c6-4309-8808-e8dfcb46c3c2}"
        }
        global IUIAutomationMaxVersion
		DllCall("user32.dll\SystemParametersInfo", "uint", 0x0046, "uint", 0, "ptr*", &screenreader:=0) ; SPI_GETSCREENREADER
		if !screenreader
			DllCall("user32.dll\SystemParametersInfo", "uint", 0x0047, "uint", 1, "int", 0, "uint", 2) ; SPI_SETSCREENREADER
        this.IUIAutomationVersion := IUIAutomationMaxVersion, this.ptr := 0
		while (--this.IUIAutomationVersion > 1) {
			if !__IID.HasOwnProp("IUIAutomation" this.IUIAutomationVersion)
				continue
			try {
                this.ptr := ComObjValue(this.__ := ComObject("{e22ad333-b25f-460c-83d0-0581107395c9}", __IID.IUIAutomation%this.IUIAutomationVersion%))
                break
            }
		}
		; If all else fails, try the first IUIAutomation version
        if !this.ptr
            this.ptr := ComObjValue(this.__ := ComObject("{ff48dba4-60ef-4201-aa87-54103eef594e}", "{30cbe57d-d9d0-452a-ab13-7ac5ac4825ee}"))
        UIA.TrueCondition := UIA.CreateTrueCondition()
        UIA.TreeWalkerTrue := UIA.CreateTreeWalker(UIA.TrueCondition)
    }
    static TrueCondition := ""
    static TreeWalkerTrue := ""

    ; IUIAutomation constants and enumerations.
    ; Access properties with UIA.property.subproperty (UIA.ControlType.Button)
    ; To get the property name from value, use the array style: UIA.property[value] (UIA.ControlType[50000])

    static __PropertyFromValue(obj, value) {
        for k, v in obj.OwnProps()
            if value = v
                return k
        throw UnsetItemError("Property item `"" value "`" not found!", -1)
    }
    static __PropertyValueGetter := {get: (obj, value) => UIA.__PropertyFromValue(obj, value)}
    static ControlType := {Button:50000,Calendar:50001,CheckBox:50002,ComboBox:50003,Edit:50004,Hyperlink:50005,Image:50006,ListItem:50007,List:50008,Menu:50009,MenuBar:50010,MenuItem:50011,ProgressBar:50012,RadioButton:50013,ScrollBar:50014,Slider:50015,Spinner:50016,StatusBar:50017,Tab:50018,TabItem:50019,Text:50020,ToolBar:50021,ToolTip:50022,Tree:50023,TreeItem:50024,Custom:50025,Group:50026,Thumb:50027,DataGrid:50028,DataItem:50029,Document:50030,SplitButton:50031,Window:50032,Pane:50033,Header:50034,HeaderItem:50035,Table:50036,TitleBar:50037,Separator:50038,SemanticZoom:50039,AppBar:50040, 50000:"Button",50001:"Calendar",50002:"CheckBox",50003:"ComboBox",50004:"Edit",50005:"Hyperlink",50006:"Image",50007:"ListItem",50008:"List",50009:"Menu",50010:"MenuBar",50011:"MenuItem",50012:"ProgressBar",50013:"RadioButton",50014:"ScrollBar",50015:"Slider",50016:"Spinner",50017:"StatusBar",50018:"Tab",50019:"TabItem",50020:"Text",50021:"ToolBar",50022:"ToolTip",50023:"Tree",50024:"TreeItem",50025:"Custom",50026:"Group",50027:"Thumb",50028:"DataGrid",50029:"DataItem",50030:"Document",50031:"SplitButton",50032:"Window",50033:"Pane",50034:"Header",50035:"HeaderItem",50036:"Table",50037:"TitleBar",50038:"Separator",50039:"SemanticZoom",50040:"AppBar"}.DefineProp("__Item", this.__PropertyValueGetter)

    static Pattern := { Invoke: 10000, Selection: 10001, Value: 10002, RangeValue: 10003, Scroll: 10004, ExpandCollapse: 10005, Grid: 10006, GridItem: 10007, MultipleView: 10008, Window: 10009, SelectionItem: 10010, Dock: 10011, Table: 10012, TableItem: 10013, Text: 10014, Toggle: 10015, Transform: 10016, ScrollItem: 10017, LegacyIAccessible: 10018, ItemContainer: 10019, VirtualizedItem: 10020, SynchronizedInput: 10021, ObjectModel: 10022, Annotation: 10023, Text2:10024, Styles: 10025, Spreadsheet: 10026, SpreadsheetItem: 10027, Transform2: 10028, TextChild: 10029, Drag: 10030, DropTarget: 10031, TextEdit: 10032, CustomNavigation: 10033, Selection2: 10034
    , 10000: "Invoke", 10001: "Selection", 10002: "Value", 10003: "RangeValue", 10004: "Scroll", 10005: "ExpandCollapse", 10006: "Grid", 10007: "GridItem", 10008: "MultipleView", 10009: "Window", 10010: "SelectionItem", 10011: "Dock", 10012: "Table", 10013: "TableItem", 10014: "Text", 10015: "Toggle", 10016: "Transform", 10017: "ScrollItem", 10018: "LegacyIAccessible", 10019: "ItemContainer", 10020: "VirtualizedItem", 10021: "SynchronizedInput", 10022: "ObjectModel", 10023: "Annotation", 10024: "Text2", 10025: "Styles", 10026: "Spreadsheet", 10027: "SpreadsheetItem", 10028:"Transform2", 10029: "TextChild", 10030: "Drag", 10031: "DropTarget", 10032: "TextEdit", 10033: "CustomNavigation", 10034: "Selection2"}.DefineProp("__Item", this.__PropertyValueGetter)

    static Event := {ToolTipOpened:20000,ToolTipClosed:20001,StructureChanged:20002,MenuOpened:20003,AutomationPropertyChanged:20004,AutomationFocusChanged:20005,AsyncContentLoaded:20006,MenuClosed:20007,LayoutInvalidated:20008,Invoke_Invoked:20009,SelectionItem_ElementAddedToSelection:20010,SelectionItem_ElementRemovedFromSelection:20011,SelectionItem_ElementSelected:20012,Selection_Invalidated:20013,Text_TextSelectionChanged:20014,Text_TextChanged:20015,Window_WindowOpened:20016,Window_WindowClosed:20017,MenuModeStart:20018,MenuModeEnd:20019,InputReachedTarget:20020,InputReachedOtherElement:20021,InputDiscarded:20022,SystemAlert:20023,LiveRegionChanged:20024,HostedFragmentRootsInvalidated:20025,Drag_DragStart:20026,Drag_DragCancel:20027,Drag_DragComplete:20028,DropTarget_DragEnter:20029,DropTarget_DragLeave:20030,DropTarget_Dropped:20031,TextEdit_TextChanged:20032,TextEdit_ConversionTargetChanged:20033,Changes:20034,Notification:20035,ActiveTextPositionChanged:20036,20000:"ToolTipOpened",20001:"ToolTipClosed",20002:"StructureChanged",20003:"MenuOpened",20004:"AutomationPropertyChanged",20005:"AutomationFocusChanged",20006:"AsyncContentLoaded",20007:"MenuClosed",20008:"LayoutInvalidated",20009:"Invoke_Invoked",20010:"SelectionItem_ElementAddedToSelection",20011:"SelectionItem_ElementRemovedFromSelection",20012:"SelectionItem_ElementSelected",20013:"Selection_Invalidated",20014:"Text_TextSelectionChanged",20015:"Text_TextChanged",20016:"Window_WindowOpened",20017:"Window_WindowClosed",20018:"MenuModeStart",20019:"MenuModeEnd",20020:"InputReachedTarget",20021:"InputReachedOtherElement",20022:"InputDiscarded",20023:"SystemAlert",20024:"LiveRegionChanged",20025:"HostedFragmentRootsInvalidated",20026:"Drag_DragStart",20027:"Drag_DragCancel",20028:"Drag_DragComplete",20029:"DropTarget_DragEnter",20030:"DropTarget_DragLeave",20031:"DropTarget_Dropped",20032:"TextEdit_TextChanged",20033:"TextEdit_ConversionTargetChanged",20034:"Changes",20035:"Notification",20036:"ActiveTextPositionChanged"}.DefineProp("__Item", this.__PropertyValueGetter)

    static Property := {RuntimeId:30000,BoundingRectangle:30001,ProcessId:30002,ControlType:30003,Type:30003,LocalizedControlType:30004,Name:30005,AcceleratorKey:30006,AccessKey:30007,HasKeyboardFocus:30008,IsKeyboardFocusable:30009,IsEnabled:30010,AutomationId:30011,ClassName:30012,HelpText:30013,ClickablePoint:30014,Culture:30015,IsControlElement:30016,IsContentElement:30017,LabeledBy:30018,IsPassword:30019,NativeWindowHandle:30020,ItemType:30021,IsOffscreen:30022,Orientation:30023,FrameworkId:30024,IsRequiredForForm:30025,ItemStatus:30026,IsDockPatternAvailable:30027,IsExpandCollapsePatternAvailable:30028,IsGridItemPatternAvailable:30029,IsGridPatternAvailable:30030,IsInvokePatternAvailable:30031,IsMultipleViewPatternAvailable:30032,IsRangeValuePatternAvailable:30033,IsScrollPatternAvailable:30034,IsScrollItemPatternAvailable:30035,IsSelectionItemPatternAvailable:30036,IsSelectionPatternAvailable:30037,IsTablePatternAvailable:30038,IsTableItemPatternAvailable:30039,IsTextPatternAvailable:30040,IsTogglePatternAvailable:30041,IsTransformPatternAvailable:30042,IsValuePatternAvailable:30043,IsWindowPatternAvailable:30044,ValueValue:30045,Value:30045,ValueIsReadOnly:30046,RangeValueValue:30047,RangeValueIsReadOnly:30048,RangeValueMinimum:30049,RangeValueMaximum:30050,RangeValueLargeChange:30051,RangeValueSmallChange:30052,ScrollHorizontalScrollPercent:30053,ScrollHorizontalViewSize:30054,ScrollVerticalScrollPercent:30055,ScrollVerticalViewSize:30056,ScrollHorizontallyScrollable:30057,ScrollVerticallyScrollable:30058,SelectionSelection:30059,SelectionCanSelectMultiple:30060,SelectionIsSelectionRequired:30061,GridRowCount:30062,GridColumnCount:30063,GridItemRow:30064,GridItemColumn:30065,GridItemRowSpan:30066,GridItemColumnSpan:30067,GridItemContainingGrid:30068,DockDockPosition:30069,ExpandCollapseExpandCollapseState:30070,MultipleViewCurrentView:30071,MultipleViewSupportedViews:30072,WindowCanMaximize:30073,WindowCanMinimize:30074,WindowWindowVisualState:30075,WindowWindowInteractionState:30076,WindowIsModal:30077,WindowIsTopmost:30078,SelectionItemIsSelected:30079,SelectionItemSelectionContainer:30080,TableRowHeaders:30081,TableColumnHeaders:30082,TableRowOrColumnMajor:30083,TableItemRowHeaderItems:30084,TableItemColumnHeaderItems:30085,ToggleToggleState:30086,TransformCanMove:30087,TransformCanResize:30088,TransformCanRotate:30089,IsLegacyIAccessiblePatternAvailable:30090,LegacyIAccessibleChildId:30091,LegacyIAccessibleName:30092,LegacyIAccessibleValue:30093,LegacyIAccessibleDescription:30094,LegacyIAccessibleRole:30095,LegacyIAccessibleState:30096,LegacyIAccessibleHelp:30097,LegacyIAccessibleKeyboardShortcut:30098,LegacyIAccessibleSelection:30099,LegacyIAccessibleDefaultAction:30100,AriaRole:30101,AriaProperties:30102,IsDataValidForForm:30103,ControllerFor:30104,DescribedBy:30105,FlowsTo:30106,ProviderDescription:30107,IsItemContainerPatternAvailable:30108,IsVirtualizedItemPatternAvailable:30109,IsSynchronizedInputPatternAvailable:30110,OptimizeForVisualContent:30111,IsObjectModelPatternAvailable:30112,AnnotationAnnotationTypeId:30113,AnnotationAnnotationTypeName:30114,AnnotationAuthor:30115,AnnotationDateTime:30116,AnnotationTarget:30117,IsAnnotationPatternAvailable:30118,IsTextPattern2Available:30119,StylesStyleId:30120,StylesStyleName:30121,StylesFillColor:30122,StylesFillPatternStyle:30123,StylesShape:30124,StylesFillPatternColor:30125,StylesExtendedProperties:30126,IsStylesPatternAvailable:30127,IsSpreadsheetPatternAvailable:30128,SpreadsheetItemFormula:30129,SpreadsheetItemAnnotationObjects:30130,SpreadsheetItemAnnotationTypes:30131,IsSpreadsheetItemPatternAvailable:30132,Transform2CanZoom:30133,IsTransformPattern2Available:30134,LiveSetting:30135,IsTextChildPatternAvailable:30136,IsDragPatternAvailable:30137,DragIsGrabbed:30138,DragDropEffect:30139,DragDropEffects:30140,IsDropTargetPatternAvailable:30141,DropTargetDropTargetEffect:30142,DropTargetDropTargetEffects:30143,DragGrabbedItems:30144,Transform2ZoomLevel:30145,Transform2ZoomMinimum:30146,Transform2ZoomMaximum:30147,FlowsFrom:30148,IsTextEditPatternAvailable:30149,IsPeripheral:30150,IsCustomNavigationPatternAvailable:30151,PositionInSet:30152,SizeOfSet:30153,Level:30154,AnnotationTypes:30155,AnnotationObjects:30156,LandmarkType:30157,LocalizedLandmarkType:30158,FullDescription:30159,FillColor:30160,OutlineColor:30161,FillType:30162,VisualEffects:30163,OutlineThickness:30164,CenterPoint:30165,Rotation:30166,Size:30167,IsSelectionPattern2Available:30168,Selection2FirstSelectedItem:30169,Selection2LastSelectedItem:30170,Selection2CurrentSelectedItem:30171,Selection2ItemCount:30173,IsDialog:30174}.DefineProp("__Item", this.__PropertyValueGetter)

    static PropertyVariantType := {30000:0x2003,30001:0x2005,30002:3,30003:3,30004:8,30005:8,30006:8,30007:8,30008:0xB,30009:0xB,30010:0xB,30011:8,30012:8,30013:8,30014:0x2005,30015:3,30016:0xB,30017:0xB,30018:0xD,30019:0xB,30020:3,30021:8,30022:0xB,30023:3,30024:8,30025:0xB,30026:8,30027:0xB,30028:0xB,30029:0xB,30030:0xB,30031:0xB,30032:0xB,30033:0xB,30034:0xB,30035:0xB,30036:0xB,30037:0xB,30038:0xB,30039:0xB,30040:0xB,30041:0xB,30042:0xB,30043:0xB,30044:0xB,30045:8,30046:0xB,30047:5,30048:0xB,30049:5,30050:5,30051:5,30052:5,30053:5,30054:5,30055:5,30056:5,30057:0xB,30058:0xB,30059:0x200D,30060:0xB,30061:0xB,30062:3,30063:3,30064:3,30065:3,30066:3,30067:3,30068:0xD,30069:3,30070:3,30071:3,30072:0x2003,30073:0xB,30074:0xB,30075:3,30076:3,30077:0xB,30078:0xB,30079:0xB,30080:0xD,30081:0x200D,30082:0x200D,30083:0x2003,30084:0x200D,30085:0x200D,30086:3,30087:0xB,30088:0xB,30089:0xB,30090:0xB,30091:3,30092:8,30093:8,30094:8,30095:3,30096:3,30097:8,30098:8,30099:0x200D,30100:8}, type2:={30101:8,30102:8,30103:0xB,30104:0xD,30105:0xD,30106:0xD,30107:8,30108:0xB,30109:0xB,30110:0xB,30111:0xB,30112:0xB,30113:3,30114:8,30115:8,30116:8,30117:0xD,30118:0xB,30119:0xB,30120:3,30121:8,30122:3,30123:8,30124:8,30125:3,30126:8,30127:0xB,30128:0xB,30129:8,30130:0x200D,30131:0x2003,30132:0xB,30133:0xB,30134:0xB,30135:3,30136:0xB,30137:0xB,30138:0xB,30139:8,30140:0x2008,30141:0xB,30142:8,30143:0x2008,30144:0x200D,30145:5,30146:5,30147:5,30148:0x200D,30149:0xB,30150:0xB,30151:0xB,30152:3,30153:3,30154:3,30155:0x2003,30156:0x2003,30157:3,30158:8,30159:8,30160:3,30161:0x2003,30162:3,30163:3,30164:0x2005,30165:0x2005,30166:5,30167:0x2005,30168:0xB}

    static TextAttribute := {AnimationStyle:40000,BackgroundColor:40001,BulletStyle:40002,CapStyle:40003,Culture:40004,FontName:40005,FontSize:40006,FontWeight:40007,ForegroundColor:40008,HorizontalTextAlignment:40009,IndentationFirstLine:40010,IndentationLeading:40011,IndentationTrailing:40012,IsHidden:40013,IsItalic:40014,IsReadOnly:40015,IsSubscript:40016,IsSuperscript:40017,MarginBottom:40018,MarginLeading:40019,MarginTop:40020,MarginTrailing:40021,OutlineStyles:40022,OverlineColor:40023,OverlineStyle:40024,StrikethroughColor:40025,StrikethroughStyle:40026,Tabs:40027,TextFlowDirections:40028,UnderlineColor:40029,UnderlineStyle:40030,AnnotationTypes:40031,AnnotationObjects:40032,StyleName:40033,StyleId:40034,Link:40035,IsActive:40036,SelectionActiveEnd:40037,CaretPosition:40038,CaretBidiMode:40039,LineSpacing:40040,BeforeParagraphSpacing:40041,AfterParagraphSpacing:40042,SayAsInterpretAs:40043,40000:"AnimationStyle",40001:"BackgroundColor",40002:"BulletStyle",40003:"CapStyle",40004:"Culture",40005:"FontName",40006:"FontSize",40007:"FontWeight",40008:"ForegroundColor",40009:"HorizontalTextAlignment",40010:"IndentationFirstLine",40011:"IndentationLeading",40012:"IndentationTrailing",40013:"IsHidden",40014:"IsItalic",40015:"IsReadOnly",40016:"IsSubscript",40017:"IsSuperscript",40018:"MarginBottom",40019:"MarginLeading",40020:"MarginTop",40021:"MarginTrailing",40022:"OutlineStyles",40023:"OverlineColor",40024:"OverlineStyle",40025:"StrikethroughColor",40026:"StrikethroughStyle",40027:"Tabs",40028:"TextFlowDirections",40029:"UnderlineColor",40030:"UnderlineStyle",40031:"AnnotationTypes",40032:"AnnotationObjects",40033:"StyleName",40034:"StyleId",40035:"Link",40036:"IsActive",40037:"SelectionActiveEnd",40038:"CaretPosition",40039:"CaretBidiMode",40040:"LineSpacing",40041:"BeforeParagraphSpacing",40042:"AfterParagraphSpacing",40043:"SayAsInterpretAs"}.DefineProp("__Item", this.__PropertyValueGetter)

    static AttributeVariantType := {40000:3,40001:3,40002:3,40003:3,40004:3,40005:8,40006:5,40007:3,40008:3,40009:3,40010:5,40011:5,40012:5,40013:0xB,40014:0xB,40015:0xB,40016:0xB,40017:0xB,40018:5,40019:5,40020:5,40021:5,40022:3,40023:3,40024:3,40025:3,40026:3,40027:0x2005,40028:3,40029:3,40030:3,40031:0x2003,40032:0x200D,40033:8,40034:3,40035:0xD,40036:0xB,40037:3,40038:3,40039:3,40040:8,40041:5,40042:5,40043:8}

    static AnnotationType := {Unknown:60000,SpellingError:60001,GrammarError:60002,Comment:60003,FormulaError:60004,TrackChanges:60005,Header:60006,Footer:60007,Highlighted:60008,Endnote:60009,Footnote:60010,InsertionChange:60011,DeletionChange:60012,MoveChange:60013,FormatChange:60014,UnsyncedChange:60015,EditingLockedChange:60016,ExternalChange:60017,ConflictingChange:60018,Author:60019,AdvancedProofingIssue:60020,DataValidationError:60021,CircularReferenceError:60022,Mathematics:60023, 60000:"Unknown",60001:"SpellingError",60002:"GrammarError",60003:"Comment",60004:"FormulaError",60005:"TrackChanges",60006:"Header",60007:"Footer",60008:"Highlighted",60009:"Endnote",60010:"Footnote",60011:"InsertionChange",60012:"DeletionChange",60013:"MoveChange",60014:"FormatChange",60015:"UnsyncedChange",60016:"EditingLockedChange",60017:"ExternalChange",60018:"ConflictingChange",60019:"Author",60020:"AdvancedProofingIssue",60021:"DataValidationError",60022:"CircularReferenceError",60023:"Mathematics"}.DefineProp("__Item", this.__PropertyValueGetter)

    static Style := {Custom:70000,Heading1:70001,Heading2:70002,Heading3:70003,Heading4:70004,Heading5:70005,Heading6:70006,Heading7:70007,Heading8:70008,Heading9:70009,Title:70010,Subtitle:70011,Normal:70012,Emphasis:70013,Quote:70014,BulletedList:70015,NumberedList:70016,70000:"Custom",70001:"Heading1",70002:"Heading2",70003:"Heading3",70004:"Heading4",70005:"Heading5",70006:"Heading6",70007:"Heading7",70008:"Heading8",70009:"Heading9",70010:"Title",70011:"Subtitle",70012:"Normal",70013:"Emphasis",70014:"Quote",70015:"BulletedList",70016:"NumberedList"}.DefineProp("__Item", this.__PropertyValueGetter)

    static LandmarkType := {Custom:80000,Form:80001,Main:80002,Navigation:80003,Search:80004,80000:"Custom",80001:"Form",80002:"Main",80003:"Navigation",80004:"Search"}.DefineProp("__Item", this.__PropertyValueGetter)
    
    static HeadingLevel := {None:80050, 1:80051, 2:80052, 3:80053, 4:80054, 5:80055, 6:80056, 7:80057, 8:80058, 9:80059, 80050:"None", 80051:"1", 80052:"2", 80053:"3", 80054:"4", 80055:"5", 80056:"6", 80057:"7", 80058:"8", 80059:"9"}.DefineProp("__Item", this.__PropertyValueGetter)

    static Change := {Summary:90000, 90000:"Summary"}.DefineProp("__Item", this.__PropertyValueGetter)

    static Metadata := {SayAsInterpretAs:100000, 100000:"SayAsInterpretAs"}.DefineProp("__Item", this.__PropertyValueGetter)

    static AsyncContentLoadedState := {Beginning:0, Progress:1, Completed:2, 0:"Beginning", 1:"Progress", 2:"Completed"}.DefineProp("__Item", this.__PropertyValueGetter)

    static AutomationIdentifierType := {Property:0,Pattern:1,Event:2,ControlType:3,TextAttribute:4,LandmarkType:5,Annotation:6,Changes:7,Style:7,0:"Property", 1:"Pattern", 2:"Event", 3:"ControlType", 4:"TextAttribute", 5:"LandmarkType", 6:"Annotation", 7:"Changes", 8:"Style"}.DefineProp("__Item", this.__PropertyValueGetter)

    static ConditionType := {True:0,False:1,Property:2,And:3,Or:4,Not:5,0:"True", 1:"False", 2:"Property", 3:"And", 4:"Or", 5:"Not"}.DefineProp("__Item", this.__PropertyValueGetter)

    static EventArgsType := {Simple:0,PropertyChanged:1,StructureChanged:2,AsyncContentLoaded:3,WindowClosed:4,TextEditTextChanged:5,Changes:6,Notification:7,ActiveTextPositionChanged:8,StructuredMarkup:9, 0:"Simple", 1:"PropertyChanged", 2:"StructureChanged", 3:"AsyncContentLoaded", 4:"WindowClosed", 5:"TextEditTextChanged", 6:"Changes", 7:"Notification", 8:"ActiveTextPositionChanged", 9:"StructuredMarkup"}

    static AutomationElementMode := {None:0, Full:1, 0:"None", 1:"Full"}.DefineProp("__Item", this.__PropertyValueGetter)

    static CoalesceEventsOptions := {Disabled:0, Enabled:1, 0:"Disabled", 1:"Enabled"}.DefineProp("__Item", this.__PropertyValueGetter)

    static ConnectionRecoveryBehaviorOptions := {Disabled:0, Enabled:1, 0:"Disabled", 1:"Enabled"}.DefineProp("__Item", this.__PropertyValueGetter)

    static PropertyConditionFlags := {None:0, IgnoreCase:1, MatchSubstring:2, 0:"None", 1:"IgnoreCase", 2:"MatchSubstring"}.DefineProp("__Item", this.__PropertyValueGetter)

    static TreeScope := {None: 0, Element: 1, Children: 2, Descendants: 4, Subtree: 7, Parent: 8, Ancestors: 16, 0x0:"None", 0x1:"Element", 0x2:"Children", 0x4:"Descendants", 0x8:"Parent", 0x10:"Ancestors", 0x7:"Subtree"
    }.DefineProp("__Item", this.__PropertyValueGetter)

    static TreeTraversalOptions := {Default:0, PostOrder:1, LastToFirstOrder:2}.DefineProp("__Item", this.__PropertyValueGetter)
	; enum ActiveEnd Contains possible values for the SelectionActiveEnd text attribute, which indicates the location of the caret relative to a text range that represents the currently selected text.
    static ActiveEnd := {None:0,Start:1,End:2}.DefineProp("__Item", this.__PropertyValueGetter)
	; enum AnimationStyle Contains values for the AnimationStyle text attribute.
    static AnimationStyle := {None:0,LasVegasLights:1,BlinkingBackground:2,SparkleText:3,MarchingBlackAnts:4,MarchingRedAnts:5,Shimmer:6,Other:-1}.DefineProp("__Item", this.__PropertyValueGetter)
	; enum BulletStyle Contains values for the BulletStyle text attribute.
    static BulletStyle := {None:0,HollowRoundBullet:1,FilledRoundBullet:2,HollowSquareBullet:3,FilledSquareBullet:4,DashBullet:5,Other:-1}.DefineProp("__Item", this.__PropertyValueGetter)
	; enum CapStyle Contains values that specify the value of the CapStyle text attribute.
    static CapStyle := {None:0,SmallCap:1,AllCap:2,AllPetiteCaps:3,PetiteCaps:4,Unicase:5,Titling:6,Other:-1}.DefineProp("__Item", this.__PropertyValueGetter)
	; enum CaretBidiMode Contains possible values for the CaretBidiMode text attribute, which indicates whether the caret is in text that flows from left to right, or from right to left.
    static CaretBidiMode := {LTR:0,RTL:1}.DefineProp("__Item", this.__PropertyValueGetter)
	; enum CaretPosition Contains possible values for the CaretPosition text attribute, which indicates the location of the caret relative to a line of text in a text range.
    static CaretPosition := {Unknown:0,EndOfLine:1,BeginningOfLine:2}.DefineProp("__Item", this.__PropertyValueGetter)
	; enum DockPosition Contains values that specify the location of a docking window represented by the Dock control pattern.
    static DockPosition := {Top:0,Left:1,Bottom:2,Right:3,Fill:4,None:5}.DefineProp("__Item", this.__PropertyValueGetter)
	; enum ExpandCollapseState Contains values that specify the state of a UI element that can be expanded and collapsed.	
    static ExpandCollapseState := {Collapsed:0,Expanded:1,PartiallyExpanded:2,LeafNode:3}.DefineProp("__Item", this.__PropertyValueGetter)
	; enum FillType Contains values for the FillType attribute.
    static FillType := {None:0,Color:1,Gradient:2,Picture:3,Pattern:4}.DefineProp("__Item", this.__PropertyValueGetter)
    ; enum FlowDirection Contains values for the TextFlowDirections text attribute.
    static FlowDirection := {Default:0,RightToLeft:1,BottomToTop:2,Vertical:4}.DefineProp("__Item", this.__PropertyValueGetter)
	;enum LiveSetting Contains possible values for the LiveSetting property. This property is implemented by provider elements that are part of a live region.
    static LiveSetting := {Off:0,Polite:1,Assertive:2}.DefineProp("__Item", this.__PropertyValueGetter)
	; enum NavigateDirection Contains values used to specify the direction of navigation within the Microsoft UI Automation tree.
    static NavigateDirection := {Parent:0,NextSibling:1,PreviousSibling:2,FirstChild:3,LastChild:4}.DefineProp("__Item", this.__PropertyValueGetter)
    ; enum NotificationKind Defines values that indicate the type of a notification event, and a hint to the listener about the processing of the event. 
    static NotificationKind := {ItemAdded:0,ItemRemoved:1,ActionCompleted:2,ActionAborted:3,Other:4}.DefineProp("__Item", this.__PropertyValueGetter)
    ; enum NotificationProcessing Defines values that indicate how a notification should be processed.
    static NotificationProcessing := {ImportantAll:0,ImportantMostRecent:1,All:2,MostRecent:3,CurrentThenMostRecent:4}.DefineProp("__Item", this.__PropertyValueGetter)
	; enum OrientationType Contains values that specify the orientation of a control.
    static OrientationType := {None:0,Horizontal:1,Vertical:2}.DefineProp("__Item", this.__PropertyValueGetter)
	; enum OutlineStyles Contains values for the OutlineStyle text attribute.
    static OutlineStyles := {None:0,Outline:1,Shadow:2,Engraved:4,Embossed:8}.DefineProp("__Item", this.__PropertyValueGetter)
    ;enum ProviderOptions
    static ProviderOptions := {ClientSideProvider:1,ServerSideProvider:2,NonClientAreaProvider:4,OverrideProvider:8,ProviderOwnsSetFocus:10,UseComThreading:20,RefuseNonClientSupport:40,HasNativeIAccessible:80,UseClientCoordinates:100}.DefineProp("__Item", this.__PropertyValueGetter)
	; enum RowOrColumnMajor Contains values that specify whether data in a table should be read primarily by row or by column.
    static RowOrColumnMajor := {RowMajor:0,ColumnMajor:1,Indeterminate:2}.DefineProp("__Item", this.__PropertyValueGetter)
    ; enum SayAsInterpretAs Defines the values that indicate how a text-to-speech engine should interpret specific data.
    static SayAsInterpretAs := {None:0,Spell:1,Cardinal:2,Ordinal:3,Number:4,Date:5,Time:6,Telephone:7,Currency:8,Net:9,Url:10,Address:11,Name:13,Media:14,Date_MonthDayYear:15,Date_DayMonthYear:16,Date_YearMonthDay:17,Date_YearMonth:18,Date_MonthYear:19,Date_DayMonth:20,Date_MonthDay:21,Date_Year:22,Time_HoursMinutesSeconds12:23,Time_HoursMinutes12:24,Time_HoursMinutesSeconds24:25,Time_HoursMinutes24:26}.DefineProp("__Item", this.__PropertyValueGetter)
    ; enum ScrollAmount Contains values that specify the direction and distance to scroll.	
    static ScrollAmount := {LargeDecrement:0,SmallDecrement:1,NoAmount:2,LargeIncrement:3,SmallIncrement:4}.DefineProp("__Item", this.__PropertyValueGetter)
    ; enum StructureChangeType Contains values that specify the type of change in the Microsoft UI Automation tree structure.
    static StructureChangeType := {ChildAdded:0,ChildRemoved:1,ChildrenInvalidated:2,ChildrenBulkAdded:3,ChildrenBulkRemoved:4,ChildrenReordered:5}.DefineProp("__Item", this.__PropertyValueGetter)
    ; enum SupportedTextSelection Contains values that specify the supported text selection attribute.	
    static SupportedTextSelection := {None:0,Single:1,Multiple:2}.DefineProp("__Item", this.__PropertyValueGetter)
    ; enum SynchronizedInputType Contains values that specify the type of synchronized input.
    static SynchronizedInputType := {KeyUp:1,KeyDown:2,LeftMouseUp:4,LeftMouseDown:8,RightMouseUp:10,RightMouseDown:20}.DefineProp("__Item", this.__PropertyValueGetter)
    ; enum TextDecorationLineStyle Contains values that specify the OverlineStyle, StrikethroughStyle, and UnderlineStyle text attributes.
	static TextDecorationLineStyle := {None:0, Single:1, WordsOnly:2, Double:3, Dot:4, Dash:5, DashDot:6, DashDotDot:7, Wavy:8, ThickSingle:9, DoubleWavy:11, ThickWavy:12, LongDash:13, ThickDash:14, ThickDashDot:15, ThickDashDotDot:16, ThickDot:17, ThickLongDash:18, Other:-1}.DefineProp("__Item", this.__PropertyValueGetter)
    ; enum TextEditChangeType Describes the text editing change being performed by controls when text-edit events are raised or handled.
    static TextEditChangeType := {None:0,AutoCorrect:1,Composition:2,CompositionFinalized:3,AutoComplete:4}.DefineProp("__Item", this.__PropertyValueGetter)
    ; enum TextPatternRangeEndpoint Contains values that specify the endpoints of a text range.
    static TextPatternRangeEndpoint := {Start:0,End:1}.DefineProp("__Item", this.__PropertyValueGetter)
    ; enum TextUnit Contains values that specify units of text for the purposes of navigation.
    static TextUnit := {Character:0,Format:1,Word:2,Line:3,Paragraph:4,Page:5,Document:6}.DefineProp("__Item", this.__PropertyValueGetter)
    ; enum ToggleState Contains values that specify the toggle state of a Microsoft UI Automation element that implements the Toggle control pattern.
    static ToggleState := {Off:0,On:1,Indeterminate:2}.DefineProp("__Item", this.__PropertyValueGetter)
    ; enum ZoomUnit Contains possible values for the IUIAutomationTransformPattern2::ZoomByUnit method, which zooms the viewport of a control by the specified unit.
    static ZoomUnit := {NoAmount:0,LargeDecrement:1,SmallDecrement:2,LargeIncrement:3,SmallIncrement:4}.DefineProp("__Item", this.__PropertyValueGetter)
    ; enum WindowVisualState Contains values that specify the visual state of a window.
    static WindowVisualState := {Normal:0,Maximized:1,Minimized:2}.DefineProp("__Item", this.__PropertyValueGetter)
    ; enum WindowInteractionState Contains values that specify the current state of the window for purposes of user interaction.
    static WindowInteractionState := {Running:0,Closing:1,ReadyForUserInteraction:2,BlockedByModalWindow:3,NotResponding:4}.DefineProp("__Item", this.__PropertyValueGetter)

    ; BSTR wrapper, convert BSTR to AHK string and free it
    static BSTR(ptr) {
        static _ := DllCall("LoadLibrary", "str", "oleaut32.dll")
        if ptr {
            s := StrGet(ptr), DllCall("oleaut32\SysFreeString", "ptr", ptr)
            return s
        }
    }
    ; X can be pt64 as well, in which case Y should be omitted
	static WindowFromPoint(X, Y?) { ; by SKAN and Linear Spoon
		return DllCall("GetAncestor", "UInt", DllCall("user32.dll\WindowFromPoint", "Int64", IsSet(Y) ? (Y << 32 | X) : X), "UInt", 2)
	}

    class Variant {
        __New(Value := unset, VarType := 0xC) {
            static SIZEOF_VARIANT := 8 + (2 * A_PtrSize)
            this.var := Buffer(SIZEOF_VARIANT, 0)
            this.owner := True
            if IsSet(Value) {
                if (Type(Value) == "ComVar") {
                    this.var := Value.var, this.ref := Value.ref, this.obj := Value, this.owner := False
                    return
                } 
                if (IsObject(Value)) {
                    this.ref := ComValue(0x400C, this.var.ptr)
                    if Value is Array {
                        if Value.Length {
                            switch Type(Value[1]) {
                                case "Integer": VarType := 3
                                case "String": VarType := 8
                                case "Float": VarType := 5
                                case "ComValue", "ComObject": VarType := ComObjType(Value[1])
                                default: VarType := 0xC
                            }
                        } else
                            VarType := 0xC
                        ComObjFlags(obj := ComObjArray(VarType, Value.Length), -1), i := 0, this.ref[] := obj
                        for v in Value
                            obj[i++] := v
                        return
                    }
                }
            }
            this.ref := ComValue(0x4000 | VarType, this.var.Ptr + (VarType = 0xC ? 0 : 8))
            this.ref[] := Value
        }
        __Delete() => (this.owner ? DllCall("oleaut32\VariantClear", "ptr", this.var) : 0)
        __Item {
            get => this.Type=0xB?-this.ref[]:this.ref[]
            set => this.ref[] := this.Type=0xB?(!value?0:-1):value
        }
        Ptr => this.var.Ptr
        Size => this.var.Size
        Type {
            get => NumGet(this.var, "ushort")
            set {
                if (!this.IsVariant)
                    throw PropertyError("VarType is not VT_VARIANT, Type is read-only.", -2)
                NumPut("ushort", Value, this.var)
            }
        }
        IsVariant => ComObjType(this.ref) & 0xC
    }

    ; Construction and deconstruction VARIANT struct
    class ComVar {
        /**
         * Construction VARIANT struct, `ptr` property points to the address, `__Item` property returns var's Value
         * @param vVal Values that need to be wrapped, supports String, Integer, Double, Array, ComValue, ComObjArray
         * ### example
         * `var1 := ComVar('string'), MsgBox(var1[])`
         * 
         * `var2 := ComVar([1,2,3,4], , true)`
         * 
         * `var3 := ComVar(ComValue(0xb, -1))`
         * @param vType Variant's type, VT_VARIANT(default)
         * @param convert Convert AHK's array to ComObjArray
         */
        __New(vVal := unset, vType := 0xC, convert := false) {
            static size := 8 + 2 * A_PtrSize
            this.var := Buffer(size, 0), this.owner := true
            this.ref := ComValue(0x4000 | vType, this.var.Ptr + (vType = 0xC ? 0 : 8))
            if IsSet(vVal) {
                if (Type(vVal) == "ComVar") {
                    this.var := vVal.var, this.ref := vVal.ref, this.obj := vVal, this.owner := false
                } else {
                    if (IsObject(vVal)) {
                        if (vType != 0xC)
                            this.ref := ComValue(0x400C, this.var.ptr)
                        if convert && (vVal is Array) {
                            switch Type(vVal[1]) {
                                case "Integer": vType := 3
                                case "String": vType := 8
                                case "Float": vType := 5
                                case "ComValue", "ComObject": vType := ComObjType(vVal[1])
                                default: vType := 0xC
                            }
                            ComObjFlags(obj := ComObjArray(vType, vVal.Length), -1), i := 0, this.ref[] := obj
                            for v in vVal
                                obj[i++] := v
                        } else
                            this.ref[] := vVal
                    } else
                        this.ref[] := vVal
                }
            }
        }
        __Delete() => (this.owner ? DllCall("oleaut32\VariantClear", "ptr", this.var) : 0)
        __Item {
            get => this.ref[]
            set => this.ref[] := value
        }
        Ptr => this.var.Ptr
        Size => this.var.Size
        Type {
            get => NumGet(this.var, "ushort")
            set {
                if (!this.IsVariant)
                    throw PropertyError("VarType is not VT_VARIANT, Type is read-only.", -2)
                NumPut("ushort", Value, this.var)
            }
        }
        IsVariant => ComObjType(this.ref) & 0xC
    }
    ; NativeArray is C style array, zero-based index, it has `__Item` and `__Enum` property
    class NativeArray {
        __New(ptr, count, type := "ptr") {
            static _ := DllCall("LoadLibrary", "str", "ole32.dll")
            static bits := { UInt: 4, UInt64: 8, Int: 4, Int64: 8, Short: 2, UShort: 2, Char: 1, UChar: 1, Double: 8, Float: 4, Ptr: A_PtrSize, UPtr: A_PtrSize }
            this.size := (this.count := count) * (bit := bits.%type%), this.ptr := ptr || DllCall("ole32\CoTaskMemAlloc", "uint", this.size, "ptr")
            this.DefineProp("__Item", { get: (s, i) => NumGet(s, i * bit, type) })
            this.DefineProp("__Enum", { call: (s, i) => (i = 1 ?
                    (i := 0, (&v) => i < count ? (v := NumGet(s, i * bit, type), ++i) : false) :
                        (i := 0, (&k, &v) => (i < count ? (k := i, v := NumGet(s, i * bit, type), ++i) : false))
                ) })
        }
        __Delete() => DllCall("ole32\CoTaskMemFree", "ptr", this)
    }
        
    ; The base class for IUIAutomation objects that return releasable pointers
    class IUIAutomationBase {
        __New(ptr) {
            if !(this.ptr := ptr)
                throw ValueError('Invalid IUnknown interface pointer', -2, this.__Class)
        }
        __Delete() => this.Release()
        __Item => (ObjAddRef(this.ptr), ComValue(0xd, this.ptr))
        AddRef() => ObjAddRef(this.ptr)
        Release() => this.ptr ? ObjRelease(this.ptr) : 0

        __Get(Name, Params) {
            if this.base.HasOwnProp(NewName := StrReplace(Name, "Current"))
                return this.%NewName%
            else 
                throw Error("Property " Name " not found in " this.__Class " Class.",-1,Name)
        }
    
        __Call(Name, Params) {
            if this.base.HasOwnProp(NewName := StrReplace(Name, "Current"))
                return this.%NewName%(Params*)
            else 
                throw Error("Method " Name " not found in " this.__Class " Class.",-1,Name)
        }
    }

    static RuntimeIdToString(runtimeId) {
        str := ""
        for v in runtimeId
            str .= "." Format("{:X}", v)
        return LTrim(str, ".")
    }

    static RuntimeIdFromString(str) {
        t := StrSplit(str, ".")
        arr := ComObjArray(3, t.Length)
        for v in t
            arr[A_Index - 1] := Integer("0x" v)
        return arr
    }

    /**
     * Create Property Condition from AHK Object
     * @param conditionObject Object or Map or Array contains multiple Property Conditions. 
     *     Everything inside {} is an "and" condition
     *     Everything inside [] is an "or" condition
     *     Object key "not" creates a not condition
     * 
     *     matchmode key defines the MatchMode (default: 3): 1=must start with; 2=can contain anywhere in string; 3=exact match; RegEx
     *     casesensitive key defines case sensitivity (default: case-sensitive/True): True=case sensitive; False=case insensitive
     * 
     * #### Examples:
     * `{Name:"Something"}` => Name must match "Something" (case-sensitive)
     * `{Type:"Button", Name:"Something"}` => Name must match "Something" AND ControlType must be Button
     * `{Type:"Button", or:[Name:"Something", Name:"Else"]}` => Name must match "Something" OR "Else", AND ControlType must be Button
     * 
     * @returns IUIAutomationCondition
     */
    static CreateCondition(conditionObject) {
        return UIA.__ConditionBuilder(conditionObject)
    }

    static __ConditionBuilder(obj, &sanitizeMatchmode?) {
        obj := obj.Clone()
        sanitizeMM := False
        switch Type(obj) {
            case "Object":
                obj.DeleteProp("index"), obj.DeleteProp("i")
                operator := obj.DeleteProp("operator") || "and"
                cs := obj.DeleteProp("casesensitive") || obj.DeleteProp("cs") || 1
                mm := obj.DeleteProp("matchmode") || obj.DeleteProp("mm") || 3
                if IsSet(sanitizeMatchmode) {
                    if (mm = "RegEx" || mm = 1)
                        sanitizeMatchmode := True, sanitizeMM := True
                } else {
                    if !((mm == 3) || (mm == 2))
                        throw TypeError("MatchMode can only be 3 or 2 when creating UIA conditions. MatchMode 1 and RegEx are allowed with FindFirst, FindAll, and TreeWalking methods.")
                }
                flags := ((mm = 3 ? 0 : 2) | (!cs)) || obj.DeleteProp("flags") || 0
                count := ObjOwnPropCount(obj), obj := obj.OwnProps()
            case "Array":
                operator := "or", flags := 0, count := obj.Length
            default:
                throw TypeError("Invalid parameter type", -3)
        }
        if count = 0
            return UIA.TrueCondition
        arr := ComObjArray(0xd, count), i := 0
        for k, v in obj {
            if IsObject(v) {
                t := UIA.__ConditionBuilder(v, &sanitizeMatchmode?)
                if k = "not" || operator = "not"
                    t := UIA.CreateNotCondition(t)
                arr[i++] := t[]
                continue
            }
            k := IsNumber(k) ? Integer(k) : UIA.Property.%k%
            if k = 30003 && !IsInteger(v)
                try v := UIA.ControlType.%v%
            if sanitizeMM && RegexMatch(UIA.Property[k], "i)Name|AutomationId|Value|ClassName|FrameworkId") {
                t := mm = 1 ? UIA.CreatePropertyConditionEx(k, v, !cs | 2) : UIA.CreateNotCondition(UIA.CreatePropertyCondition(k, ""))
                arr[i++] := t[]
            } else if (k >= 30000) {
                t := flags ? UIA.CreatePropertyConditionEx(k, v, flags) : UIA.CreatePropertyCondition(k, v)
                arr[i++] := t[]
            }
        }
        if count = 1
            return t
        switch operator, false {
            case "and":
                return UIA.CreateAndConditionFromArray(arr)
            case "or":
                return UIA.CreateOrConditionFromArray(arr)
            default:
                return UIA.CreateFalseCondition()
        }
    }

	; This can be used when a Chromium apps content isn't accessible by normal methods (ElementFromHandle, GetRootElement etc). fromFocused=True uses the focused element as a reference point, fromFocused=False uses ElementFromPoint
	static GetChromiumContentElement(winTitle:="", &fromFocused:=True) {
		WinActivate winTitle
		WinWaitActive winTitle,,1
		WinGetPos &X, &Y, &W, &H, winTitle
        fromFocused := fromFocused ? UIA.GetFocusedElement() : UIA.ElementFromPoint(x+w//2, y+h//2)
		chromiumTW := UIA.CreateTreeWalker("ControlType=Document") ; Create a TreeWalker to find the Document element (the content)
		try focusedEl := chromiumTW.NormalizeElement(fromFocused) ; Get the first parent that is a Window element
		return focusedEl
	}
	; Tries to get the Chromium content from Chrome_RenderWidgetHostHWND1 control
	static ElementFromChromium(winTitle:="", activateChromiumAccessibility:=True, timeOut:=500, cacheRequest?) {
		try cHwnd := ControlGetHwnd("Chrome_RenderWidgetHostHWND1", winTitle)
		if !IsSet(cHwnd) || !cHwnd
			return
		cEl := UIA.ElementFromHandle(cHwnd,False,cacheRequest?)
		if activateChromiumAccessibility {
			SendMessage(WM_GETOBJECT := 0x003D, 0, 1,, cHwnd)
			if cEl {
				_ := cEl.Name ; it doesn't work without calling CurrentName (at least in Skype)
				if (cEl.ControlType == 50030) {
					waitTime := A_TickCount + timeOut
					while (!cEl.Value && (A_TickCount < waitTime))
						Sleep 20
				}
			}
		}
		return cEl
	}
	; In some setups Chromium-based renderers don't react to UIA calls by enabling accessibility, so we need to send the WM_GETOBJECT message to the renderer control to enable accessibility. Thanks to users malcev and rommmcek for this tip. Explanation why this works: https://www.chromium.org/developers/design-documents/accessibility/#TOC-How-Chrome-detects-the-presence-of-Assistive-Technology 
	static ActivateChromiumAccessibility(winTitle:="", cacheRequest?) {
		static activatedHwnds := Map(), WM_GETOBJECT := 0x003D
        hwnd := IsInteger(winTitle) ? winTitle : WinExist(winTitle)
        if activatedHwnds.Has(hwnd)
            return 1
        activatedHwnds[hWnd] := 1, cHwnd := 0
        return UIA.ElementFromChromium(hwnd)
	}

    ; ---------- IUIAutomation ----------

    ; Compares two UI Automation elements to determine whether they represent the same underlying UI element.
    static CompareElements(el1, el2) => (ComCall(3, this, "ptr", el1, "ptr", el2, "int*", &areSame := 0), areSame)

    ; Compares two integer arrays containing run-time identifiers (IDs) to determine whether their content is the same and they belong to the same UI element.
    static CompareRuntimeIds(runtimeId1, runtimeId2) => (ComCall(4, this, "ptr", runtimeId1, "ptr", runtimeId2, "int*", &areSame := 0), areSame)

    ; Retrieves the UI Automation element that represents the desktop.
    static GetRootElement() => (ComCall(5, this, "ptr*", &root := 0), UIA.IUIAutomationElement(root))

    ; Retrieves a UI Automation element for the specified window.
    static ElementFromHandle(hwnd:="", activateChromiumAccessibility:=True, cacheRequest?) {
		if !IsInteger(hwnd)
			hwnd := WinExist(hwnd)
		if (activateChromiumAccessibility && IsObject(retEl := UIA.ActivateChromiumAccessibility(hwnd, cacheRequest?)))
			return retEl
        return IsSet(cacheRequest) ? UIA.ElementFromHandleBuildCache(hwnd, cacheRequest) : (ComCall(6, this, "ptr", hwnd, "ptr*", &element := 0), UIA.IUIAutomationElement(element))
    }

    ; Retrieves the UI Automation element at the specified point on the desktop.
    static ElementFromPoint(x?, y?, &activateChromiumAccessibility:=True) {
		if !(IsSet(x) && IsSet(y))
			DllCall("user32.dll\GetCursorPos", "int64P", &pt64:=0)
		else
            pt64 := y << 32 | x
		if (activateChromiumAccessibility && (hwnd := DllCall("GetAncestor", "UInt", DllCall("user32.dll\WindowFromPoint", "int64",  pt64), "UInt", 2))) { ; hwnd from point by SKAN
			activateChromiumAccessibility := UIA.ActivateChromiumAccessibility(hwnd)
		}
		return (ComCall(7, this, "int64", pt64, "ptr*", &element := 0), UIA.IUIAutomationElement(element))
    }

    ; Retrieves the UI Automation element that has the input focus.
    static GetFocusedElement() => (ComCall(8, this, "ptr*", &element := 0), UIA.IUIAutomationElement(element))

    ; Retrieves the UI Automation element that has the input focus, prefetches the requested properties and control patterns, and stores the prefetched items in the cache.
    static GetRootElementBuildCache(cacheRequest) => (ComCall(9, this, "ptr", cacheRequest, "ptr*", &root := 0), UIA.IUIAutomationElement(root))

    ; Retrieves a UI Automation element for the specified window, prefetches the requested properties and control patterns, and stores the prefetched items in the cache.
    static ElementFromHandleBuildCache(hwnd, cacheRequest) => (ComCall(10, this, "ptr", hwnd, "ptr", cacheRequest, "ptr*", &element := 0), UIA.IUIAutomationElement(element))

    ; Retrieves the UI Automation element at the specified point on the desktop, prefetches the requested properties and control patterns, and stores the prefetched items in the cache.
    static ElementFromPointBuildCache(pt, cacheRequest) => (ComCall(11, this, "int64", pt, "ptr", cacheRequest, "ptr*", &element := 0), UIA.IUIAutomationElement(element))

    ; Retrieves the UI Automation element that has the input focus, prefetches the requested properties and control patterns, and stores the prefetched items in the cache.
    static GetFocusedElementBuildCache(cacheRequest) => (ComCall(12, this, "ptr", cacheRequest, "ptr*", &element := 0), UIA.IUIAutomationElement(element))

    ; Retrieves a tree walker object that can be used to traverse the Microsoft UI Automation tree.
    static CreateTreeWalker(pCondition) => (ComCall(13, this, "ptr", pCondition, "ptr*", &walker := 0), UIA.IUIAutomationTreeWalker(walker))

    ; Retrieves an IUIAutomationTreeWalker interface used to discover control elements.
    static ControlViewWalker() => (ComCall(14, this, "ptr*", &walker := 0), UIA.IUIAutomationTreeWalker(walker))

    ; Retrieves an IUIAutomationTreeWalker interface used to discover content elements.
    static ContentViewWalker() => (ComCall(15, this, "ptr*", &walker := 0), UIA.IUIAutomationTreeWalker(walker))

    ; Retrieves a tree walker object used to traverse an unfiltered view of the UI Automation tree.
    static RawViewWalker() => (ComCall(16, this, "ptr*", &walker := 0), UIA.IUIAutomationTreeWalker(walker))

    ; Retrieves a predefined IUIAutomationCondition interface that selects all UI elements in an unfiltered view.
    static RawViewCondition() => (ComCall(17, this, "ptr*", &condition := 0), UIA.IUIAutomationCondition(condition))

    ; Retrieves a predefined IUIAutomationCondition interface that selects control elements.
    static ControlViewCondition() => (ComCall(18, this, "ptr*", &condition := 0), UIA.IUIAutomationCondition(condition))

    ; Retrieves a predefined IUIAutomationCondition interface that selects content elements.
    static ContentViewCondition() => (ComCall(19, this, "ptr*", &condition := 0), UIA.IUIAutomationCondition(condition))

    ; Creates a cache request.
    ; After obtaining the IUIAutomationCacheRequest interface, use its methods to specify properties and control patterns to be cached when a UI Automation element is obtained.
    static CreateCacheRequest() => (ComCall(20, this, "ptr*", &cacheRequest := 0), UIA.IUIAutomationCacheRequest(cacheRequest))

    ; Retrieves a predefined condition that selects all elements.
    static CreateTrueCondition() => (ComCall(21, this, "ptr*", &newCondition := 0), UIA.IUIAutomationBoolCondition(newCondition))

    ; Creates a condition that is always false.
    ; This method exists only for symmetry with IUIAutomation,,CreateTrueCondition. A false condition will never enable a match with UI Automation elements, and it cannot usefully be combined with any other condition.
    static CreateFalseCondition() => (ComCall(22, this, "ptr*", &newCondition := 0), UIA.IUIAutomationBoolCondition(newCondition))

    ; Creates a condition that selects elements that have a property with the specified value.
    static CreatePropertyCondition(propertyId, value?, flags?) {
        if !IsSet(value) && IsObject(propertyId)
            return UIA.CreateCondition(propertyId)
        if !IsInteger(propertyId)
            propertyId := UIA.Property.%propertyId%
        if propertyId = 30003 && !IsNumber(value)
            try value := UIA.ControlType.%value%
        if IsSet(flags)
            return UIA.CreatePropertyConditionEx(propertyId, value, flags)
        if A_PtrSize = 4
            v := UIA.ComVar(value, , true), ComCall(23, this, "int", propertyId, "int64", NumGet(v, 'int64'), "int64", NumGet(v, 8, "int64"), "ptr*", &newCondition := 0)
        else
            ComCall(23, this, "int", propertyId, "ptr", UIA.ComVar(value, , true), "ptr*", &newCondition := 0)
        return UIA.IUIAutomationPropertyCondition(newCondition)
    }

    ; Creates a condition that selects elements that have a property with the specified value, using optional flags.
    static CreatePropertyConditionEx(propertyId, value, flags := 0) {
        if !IsInteger(propertyId)
            propertyId := UIA.Property.%propertyId%
        if propertyId = 30003
            try value := UIA.ControlType.%value%
        if A_PtrSize = 4
            v := UIA.ComVar(value, , true), ComCall(24, this, "int", propertyId, "int64", NumGet(v, 'int64'), "int64", NumGet(v, 8, "int64"), "int", flags, "ptr*", &newCondition := 0)
        else
            ComCall(24, this, "int", propertyId, "ptr", UIA.ComVar(value, , true), "int", flags, "ptr*", &newCondition := 0)
        return UIA.IUIAutomationPropertyCondition(newCondition)
    }

    ; The Create**Condition** method calls AddRef on each pointers. This means you can call Release on those pointers after the call to Create**Condition** returns without invalidating the pointer returned from Create**Condition**. When you call Release on the pointer returned from Create**Condition**, UI Automation calls Release on those pointers.

    ; Creates a condition that selects elements that match both of two conditions.
    static CreateAndCondition(condition1, condition2) => (ComCall(25, this, "ptr", condition1, "ptr", condition2, "ptr*", &newCondition := 0), UIA.IUIAutomationAndCondition(newCondition))

    ; Creates a condition that selects elements based on multiple conditions, all of which must be true.
    static CreateAndConditionFromArray(conditions) => (ComCall(26, this, "ptr", conditions, "ptr*", &newCondition := 0), UIA.IUIAutomationAndCondition(newCondition))

    ; Creates a condition that selects elements based on multiple conditions, all of which must be true.
    static CreateAndConditionFromNativeArray(conditions, conditionCount) => (ComCall(27, this, "ptr", conditions, "int", conditionCount, "ptr*", &newCondition := 0), UIA.IUIAutomationAndCondition(newCondition))

    ; Creates a combination of two conditions where a match exists if either of the conditions is true.
    static CreateOrCondition(condition1, condition2) => (ComCall(28, this, "ptr", condition1, "ptr", condition2, "ptr*", &newCondition := 0), UIA.IUIAutomationOrCondition(newCondition))

    ; Creates a combination of two or more conditions where a match exists if any of the conditions is true.
    static CreateOrConditionFromArray(conditions) => (ComCall(29, this, "ptr", conditions, "ptr*", &newCondition := 0), UIA.IUIAutomationOrCondition(newCondition))

    ; Creates a combination of two or more conditions where a match exists if any one of the conditions is true.
    static CreateOrConditionFromNativeArray(conditions, conditionCount) => (ComCall(30, this, "ptr", conditions, "ptr", conditionCount, "ptr*", &newCondition := 0), UIA.IUIAutomationOrCondition(newCondition))

    ; Creates a condition that is the negative of a specified condition.
    static CreateNotCondition(condition) => (ComCall(31, this, "ptr", condition, "ptr*", &newCondition := 0), UIA.IUIAutomationNotCondition(newCondition))

    ; Note,  Before implementing an event handler, you should be familiar with the threading issues described in Understanding Threading Issues. http,//msdn.microsoft.com/en-us/library/ee671692(v=vs.85).aspx
    ; A UI Automation client should not use multiple threads to add or remove event handlers. Unexpected behavior can result if one event handler is being added or removed while another is being added or removed in the same client process.
    ; It is possible for an event to be delivered to an event handler after the handler has been unsubscribed, if the event is received simultaneously with the request to unsubscribe the event. The best practice is to follow the Component Object Model (COM) standard and avoid destroying the event handler object until its reference count has reached zero. Destroying an event handler immediately after unsubscribing for events may result in an access violation if an event is delivered late.

    ; Registers a method that handles Microsoft UI Automation events.
    static AddAutomationEventHandler(element, eventId, handler, scope:=0x4, cacheRequest?) => ComCall(32, this, "int", eventId, "ptr", element, "int", scope, "ptr", cacheRequest ?? 0, "ptr", handler)

    ; Removes the specified UI Automation event handler.
    static RemoveAutomationEventHandler(element, eventId, handler) => ComCall(33, this, "int", eventId, "ptr", element, "ptr", handler)

    ; Registers a method that handles property-changed events.
    ; The UI item specified by element might not support the properties specified by the propertyArray parameter.
    ; This method serves the same purpose as IUIAutomation,,AddPropertyChangedEventHandler, but takes a normal array of property identifiers instead of a SAFEARRAY.
    static AddPropertyChangedEventHandlerNativeArray(element, propertyArray, propertyCount, handler, scope:=0x4, cacheRequest?) => ComCall(34, this, "ptr", element, "int", scope, "ptr", cacheRequest ?? 0, "ptr", handler, "ptr", propertyArray, "int", propertyCount)

    ; Registers a method that handles property-changed events.
    ; The UI item specified by element might not support the properties specified by the propertyArray parameter.
    static AddPropertyChangedEventHandler(element, propertyArray, handler, scope:=0x4, cacheRequest?) {
        if !IsObject(propertyArray)
			propertyArray := [propertyArray] 
		SafeArray:=ComObjArray(0x3,propertyArray.Length)
		for i, propertyId in propertyArray
			SafeArray[i-1]:=propertyId
        ComCall(35, this, "ptr", element, "int", scope, "ptr", cacheRequest ?? 0, "ptr", handler, "ptr", SafeArray)
    }

    ; Removes a property-changed event handler.
    static RemovePropertyChangedEventHandler(element, handler) => ComCall(36, this, "ptr", element, "ptr", handler)

    ; Registers a method that handles structure-changed events.
    static AddStructureChangedEventHandler(element, handler, scope:=0x4, cacheRequest?) => ComCall(37, this, "ptr", element, "int", scope, "ptr", cacheRequest ?? 0, "ptr", handler)

    ; Removes a structure-changed event handler.
    static RemoveStructureChangedEventHandler(element, handler) => ComCall(38, this, "ptr", element, "ptr", handler)

    ; Registers a method that handles focus-changed events.
    ; Focus-changed events are system-wide; you cannot set a narrower scope.
    static AddFocusChangedEventHandler(handler, cacheRequest?) => ComCall(39, this, "ptr", cacheRequest ?? 0, "ptr", handler)

    ; Removes a focus-changed event handler.
    static RemoveFocusChangedEventHandler(handler) => ComCall(40, this, "ptr", handler)

    ; Removes all registered Microsoft UI Automation event handlers.
    static RemoveAllEventHandlers() => ComCall(41, this)

    ; Converts an array of integers to a SAFEARRAY.
    static IntNativeArrayToSafeArray(array, arrayCount) => (ComCall(42, this, "ptr", array, "int", arrayCount, "ptr*", &safeArray := 0), ComValue(0x2003, safeArray))

    ; Converts a SAFEARRAY of integers to an array.
    static IntSafeArrayToNativeArray(intArray) => (ComCall(43, this, "ptr", intArray, "ptr*", &array := 0, "int*", &arrayCount := 0), UIA.NativeArray(array, arrayCount, "int"))

    ; Creates a VARIANT that contains the coordinates of a rectangle.
    ; The returned VARIANT has a data type of VT_ARRAY | VT_R8.
    static RectToVariant(rc) => (ComCall(44, this, "ptr", rc, "ptr", var := UIA.ComVar()), var)

    ; Converts a VARIANT containing rectangle coordinates to a RECT.
    static VariantToRect(var) {
        if A_PtrSize = 4
            ComCall(45, this, "int64", NumGet(var, "int64"), "int64", NumGet(var, 8, "int64"), "ptr", rc := UIA.NativeArray(0, 4, "Int"))
        else
            ComCall(45, this, "ptr", var, "ptr", rc := UIA.NativeArray(0, 4, "Int"))
        return rc
    }

    ; Converts a SAFEARRAY containing rectangle coordinates to an array of type RECT.
    static SafeArrayToRectNativeArray(rects) => (ComCall(46, this, "ptr", rects, "ptr*", &rectArray := 0, "int*", &rectArrayCount := 0), UIA.NativeArray(rectArray, rectArrayCount, "int"))

    ; Creates a instance of a proxy factory object.
    ; Use the IUIAutomationProxyFactoryMapping interface to enter the proxy factory into the table of available proxies.
    static CreateProxyFactoryEntry(factory) => (ComCall(47, this, "ptr", factory, "ptr*", &factoryEntry := 0), UIA.IUIAutomationProxyFactoryEntry(factoryEntry))

    ; Retrieves an object that represents the mapping of Window classnames and associated data to individual proxy factories. This property is read-only.
    static ProxyFactoryMapping() => (ComCall(48, this, "ptr*", &factoryMapping := 0), UIA.IUIAutomationProxyFactoryMapping(factoryMapping))

    ; The programmatic name is intended for debugging and diagnostic purposes only. The string is not localized.
    ; This property should not be used in string comparisons. To determine whether two properties are the same, compare the property identifiers directly.

    ; Retrieves the registered programmatic name of a property.
    static GetPropertyProgrammaticName(property) => (ComCall(49, this, "int", property, "ptr*", &name := 0), UIA.BSTR(name))

    ; Retrieves the registered programmatic name of a control pattern.
    static GetPatternProgrammaticName(pattern) => (ComCall(50, this, "int", pattern, "ptr*", &name := 0), UIA.BSTR(name))

    ; This method is intended only for use by Microsoft UI Automation tools that need to scan for properties. It is not intended to be used by UI Automation clients.
    ; There is no guarantee that the element will support any particular control pattern when asked for it later.

    ; Retrieves the control patterns that might be supported on a UI Automation element.
    static PollForPotentialSupportedPatterns(pElement, &patternIds, &patternNames) {
        ComCall(51, this, "ptr", pElement, "ptr*", &patternIds := 0, "ptr*", &patternNames := 0)
        patternIds := ComValue(0x2003, patternIds), patternNames := ComValue(0x2008, patternNames)
    }

    ; Retrieves the properties that might be supported on a UI Automation element.
    static PollForPotentialSupportedProperties(pElement, &propertyIds, &propertyNames) {
        ComCall(52, this, "ptr", pElement, "ptr*", &propertyIds := 0, "ptr*", &propertyNames := 0)
        propertyIds := ComValue(0x2003, propertyIds), propertyNames := ComValue(0x2008, propertyNames)
    }

    ; Checks a provided VARIANT to see if it contains the Not Supported identifier.
    ; After retrieving a property for a UI Automation element, call this method to determine whether the element supports the retrieved property. CheckNotSupported is typically called after calling a property retrieving method such as GetPropertyValue.
    static CheckNotSupported(value) {
        if A_PtrSize = 4
            value := UIA.ComVar(value, , true), ComCall(53, this, "int64", NumGet(value, "int64"), "int64", NumGet(value, 8, "int64"), "int*", &isNotSupported := 0)
        else
            ComCall(53, this, "ptr", UIA.ComVar(value, , true), "int*", &isNotSupported := 0)
        return isNotSupported
    }

    ; Retrieves a static token object representing a property or text attribute that is not supported. This property is read-only.
    ; This object can be used for comparison with the results from UIA.IUIAutomationElement,,GetPropertyValue or IUIAutomationTextRange,,GetAttributeValue.
    static ReservedNotSupportedValue() => (ComCall(54, this, "ptr*", &notSupportedValue := 0), ComValue(0xd, notSupportedValue))

    ; Retrieves a static token object representing a text attribute that is a mixed attribute. This property is read-only.
    ; The object retrieved by IUIAutomation,,ReservedMixedAttributeValue can be used for comparison with the results from IUIAutomationTextRange,,GetAttributeValue to determine if a text range contains more than one value for a particular text attribute.
    static ReservedMixedAttributeValue() => (ComCall(55, this, "ptr*", &mixedAttributeValue := 0), ComValue(0xd, mixedAttributeValue))

    ; This method enables UI Automation clients to get UIA.IUIAutomationElement interfaces for accessible objects implemented by a Microsoft Active Accessiblity server.
    ; This method may fail if the server implements UI Automation provider interfaces alongside Microsoft Active Accessibility support.
    ; The method returns E_INVALIDARG if the underlying implementation of the Microsoft UI Automation element is not a native Microsoft Active Accessibility server; that is, if a client attempts to retrieve the IAccessible interface for an element originally supported by a proxy object from Oleacc.dll, or by the UIA-to-MSAA Bridge.

    ; Retrieves a UI Automation element for the specified accessible object from a Microsoft Active Accessibility server.
    static ElementFromIAccessible(accessible, childId) => (ComCall(56, this, "ptr", accessible, "int", childId, "ptr*", &element := 0), UIA.IUIAutomationElement(element))

    ; Retrieves a UI Automation element for the specified accessible object from a Microsoft Active Accessibility server, prefetches the requested properties and control patterns, and stores the prefetched items in the cache.
    static ElementFromIAccessibleBuildCache(accessible, childId, cacheRequest) => (ComCall(57, this, "ptr", accessible, "int", childId, "ptr", cacheRequest, "ptr*", &element := 0), UIA.IUIAutomationElement(element))

    ; ---------- IUIAutomation2 ----------
	
	; Specifies whether calls to UI Automation control pattern methods automatically set focus to the target element. Default is True. 
	static AutoSetFocus {
		get => (ComCall(58, this,  "int*", &out:=0), out)
		set => ComCall(59, this,  "int", Value)
	}
	; Specifies the length of time that UI Automation will wait for a provider to respond to a client request for an automation element. Default is 20000ms (20 seconds), minimum seems to be 50ms.
	static ConnectionTimeout {
		get => (ComCall(60, this,  "int*", &out:=0), out)
		set => ComCall(61, this,  "int", Value) ; Minimum seems to be 50 (ms?)
	}
	; Specifies the length of time that UI Automation will wait for a provider to respond to a client request for information about an automation element. Default is 2000ms (2 seconds), minimum seems to be 50ms.
	static TransactionTimeout {
		get => (ComCall(62, this,  "int*", &out:=0), out)
		set => ComCall(63, this,  "int", Value)
	}

    ; ---------- IUIAutomation3 ----------

	static AddTextEditTextChangedEventHandler(element, textEditChangeType, handler, scope:=0x4, cacheRequest:=0) => (ComCall(64, this,  "ptr", element, "int", scope, "int", textEditChangeType, "ptr", cacheRequest, "ptr", handler))
	static RemoveTextEditTextChangedEventHandler(element, handler) => ComCall(65, this,  "ptr", element, "ptr", handler)

    ; ---------- IUIAutomation4 ----------

	static AddChangesEventHandler(element, changeTypes, handler, scope:=0x4, cacheRequest:=0) {
        if !IsObject(changeTypes)
            changeTypes := [changeTypes]
        nativeArray := UIA.NativeArray(0, changeTypes.Length, "int")
        for k, v in changeTypes
            NumPut("int", v, nativeArray, (k-1)*4)
        ComCall(66, this,  "ptr", element, "int", scope, "ptr", nativeArray, "int", changeTypes.Length, "ptr", cacheRequest, "ptr", handler)
    }
	static RemoveChangesEventHandler(element, handler) => (ComCall(67, this,  "ptr", element, "ptr", handler))

    ; ---------- IUIAutomation5 ----------

	static AddNotificationEventHandler(element, handler, scope:=0x4, cacheRequest:=0) => (ComCall(68, this,  "ptr", element, "uint", scope, "ptr", cacheRequest, "ptr", handler))
	static RemoveNotificationEventHandler(element, handler) => (ComCall(69, this,  "ptr", element, "ptr", handler))

    ; ---------- IUIAutomation6 ----------

	; Indicates whether an accessible technology client adjusts provider request timeouts when the provider is non-responsive.
	ConnectionRecoveryBehavior {
		get => (ComCall(73, this,  "int*", &out), out)
		set => ComCall(74, this,  "int", value) 
	}
	; Gets or sets whether an accessible technology client receives all events, or a subset where duplicate events are detected and filtered.
	CoalesceEvents {
		get => (ComCall(75, this,  "int*", &out:=0), out)
		set => ComCall(76, this,  "int", value)
	}

	; Registers one or more event listeners in a single method call.
	static CreateEventHandlerGroup() => (ComCall(70, this,  "ptr*", &out:=0), UIA.IUIAutomationEventHandlerGroup(out))
	; Registers a collection of event handler methods specified with the IUIAutomation6 CreateEventHandlerGroup.
	static AddEventHandlerGroup(element, handlerGroup) => (ComCall(71, this,  "ptr", element, "ptr", handlerGroup))
	static RemoveEventHandlerGroup(element, handlerGroup) => (ComCall(72, this,  "ptr", element, "ptr", handlerGroup))
	; Registers a method that handles when the active text position changes.
	static AddActiveTextPositionChangedEventHandler(element, handler, scope:=0x4, cacheRequest:=0) => (ComCall(77, this,  "ptr", element, "int", scope, "ptr", cacheRequest, "ptr", handler))
	static RemoveActiveTextPositionChangedEventHandler(element, handler) => (ComCall(78, this,  "ptr", element, "ptr", handler))

    ; ---------- IUIAutomation7 ----------
    ; Has no properties/methods

/*
	Exposes methods and properties for a UI Automation element, which represents a UI item.
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationelement
*/
class IUIAutomationElement extends UIA.IUIAutomationBase {
    cachedPatterns := Map() ; Used to cache patterns when they are called directly from an element

    /**
     * Enables array-like use of UIA elements to access child elements. 
     * If value is an integer then the nth corresponding child will be returned. 
     *     Eg. Element[2] ==> returns second child of Element
     * If value is a string, then it will be parsed as a comma-separated path which
     * allows both indexes (nth child), traversing tree sideways with +-, and Type values.
     *     Eg. Element["+1,Button2,4"] ==> gets next sibling element in tree, then from its children
     *     the second button, then its fourth child
     * If value is an object, then it will be used in a FindFirst call with scope set to Children.
     *     Eg. Element[{Type:"Button"}] will return the first child with ControlType Button.
     * @returns {UIA.IUIAutomationElement}
     */
    __Item[params*] {
        get {
            el := this
            for _, param in params {
                if IsInteger(param)
                    el := el.GetChildren()[param]
                else if IsObject(param)
                    el := el.FindFirst(param, 2)
                else if Type(param) = "String"
                    el := el.FindByPath(param)
                else
                    TypeError("Invalid item type!", -1)
            }
            return el
        }
    }
    /**
     * Enables enumeration of UIA elements, usually in a for loop. 
     * Usage:
     * for [index, ] child in Element
     */
    __Enum(varCount) {
        maxLen := this.Length, i := 0, children := this.GetChildren()
        EnumElements(&element) {
            if ++i > maxLen
                return false
            element := children[i]
            return true
        }
        EnumIndexAndElements(&index, &element) {
            if ++i > maxLen
                return false
            index := i
            element := children[i]
            return true
        }
        return (varCount = 1) ? EnumElements : EnumIndexAndElements
    }
    /**
     * Getter for element properties and element supported pattern properties. 
     * This allows for syntax such as:
     *     Element.Name == Element.CurrentName
     *     Element.ValuePattern == Element.GetPattern("Value")
     */
    __Get(Name, Params) {
        NewName := StrReplace(Name, "Current",,,,1)
        if (SubStr(Name, 1, 6) = "Cached") {
            if UIA.Property.HasOwnProp(PropName := SubStr(Name, 7))
                return this.GetCachedPropertyValue(UIA.Property.%PropName%)
        } else if this.base.HasOwnProp(NewName)
            return this.%NewName%
        else if UIA.Property.HasOwnProp(NewName)
            return this.GetPropertyValue(UIA.Property.%NewName%)
        else if (NewName ~= "i)Pattern\d?") {
            Name := RegexReplace(NewName, "i)Pattern$")
            if this.CachedPatterns.Has(Name)
                return this.CachedPatterns[Name]
            else if UIA.Pattern.HasOwnProp(Name)
                return this.CachedPatterns[Name] := this.GetPattern(UIA.Pattern.%Name%)
            else if ((SubStr(Name, 1, 6) = "Cached") && UIA.Pattern.HasOwnProp(pattern := SubStr(Name, 7)))
                return this.CachedPatterns[Name] := this.GetCachedPattern(UIA.Pattern.%pattern%)
        } else {
            for pName, pVal in UIA.Pattern.OwnProps()
                if IsInteger(pVal) && UIA.HasProp("IUIAutomation" pName "Pattern") {
                    if UIA.IUIAutomation%pName%Pattern.Prototype.HasProp(NewName) 
                        return (this.CachedPatterns.Has(pName) ? this.CachedPatterns[pName] : this.CachedPatterns[pName] := this.GetPattern(pVal)).%NewName%
                    else if UIA.IUIAutomation%pName%Pattern.Prototype.HasProp(Name) 
                        return (this.CachedPatterns.Has(pName) ? this.CachedPatterns[pName] : this.CachedPatterns[pName] := this.GetPattern(pVal)).%Name%
                }
            throw Error("Property " Name " not found in " this.__Class " Class.",-1,Name)
        }
    }
    ; Setter for UIA element and pattern properties. 
    __Set(Name, Params, Value) {
        NewName := StrReplace(Name, "Current",,,,1)
        if this.base.HasOwnProp(NewName)
            return this.%NewName% := Value
        for pName, pVal in UIA.Pattern.OwnProps()
            if IsInteger(pVal) && UIA.HasProp("IUIAutomation" pName "Pattern")
                if UIA.IUIAutomation%pName%Pattern.Prototype.HasProp(NewName)
                    return (this.CachedPatterns.Has(pName) ? this.CachedPatterns[pName] : this.CachedPatterns[pName] := this.GetPattern(pVal)).%NewName% := Value
                else if UIA.IUIAutomation%pName%Pattern.Prototype.HasProp(Name) {
                    return (this.CachedPatterns.Has(pName) ? this.CachedPatterns[pName] : this.CachedPatterns[pName] := this.GetPattern(pVal)).%Name% := Value
            }
        this.DefineProp(Name, {Value:Value})
    }
    /**
     * Meta-function for calling methods from supperted patterns. 
     * This allows for syntax such as:
     *     Element.Invoke() == Element.GetPattern("Value").Invoke()
     */
    __Call(Name, Params) {
        if this.base.HasOwnProp(NewName := StrReplace(Name, "Current",,,,1))
            return this.%NewName%(Params*)
        for pName, pVal in UIA.Pattern.OwnProps()
            if IsInteger(pVal) && UIA.HasProp("IUIAutomation" pName "Pattern") {
                if UIA.IUIAutomation%pName%Pattern.Prototype.HasMethod(Name) 
                    return (this.CachedPatterns.Has(pName) ? this.CachedPatterns[pName] : this.CachedPatterns[pName] := this.GetPattern(pVal)).%Name%(Params*)
                else if UIA.IUIAutomation%pName%Pattern.Prototype.HasMethod(NewName)
                    return (this.CachedPatterns.Has(pName) ? this.CachedPatterns[pName] : this.CachedPatterns[pName] := this.GetPattern(pVal)).%NewName%(Params*)
            }
        throw Error("Method " Name " not found in " this.__Class " Class.",-1,Name)
    }
    ; Returns all direct children of the element
    Children => this.GetChildren()
    ; Returns the number of children of the element
    Length => this.GetChildren().Length
    /**
     * Returns an object containing the location of the element
     * @returns {Object} {x: screen x-coordinate, y: screen y-coordinate, w: width, h: height}
     */
    Location => (br := this.BoundingRectangle, {x: br.l, y: br.t, w: br.r-br.l, h: br.b-br.t})

	; Gets or sets the current value of the element, if the ValuePattern is supported.
	Value { 
		get {
			try return this.GetPropertyValue(UIA.Property.ValueValue)
		}
		set {
            try return this.GetPattern(UIA.Pattern.Value).SetValue(value)
            try return this.GetPattern(UIA.Pattern.LegacyIAccessible).SetValue(value)
            Throw Error("Setting the value failed! Is ValuePattern or LegacyIAccessiblePattern supported?",,-1)
		}
	}
	CachedValue => this.GetCachedPropertyValue(UIA.Property.ValueValue)
    ; Checks whether this object still exists
    Exists {
        get {
			try return (this.Name this.Value (this.BoundingRectangle.t ? 1 : "")) != ""
			return 1
        }
    }

	/**
     * Returns the children of this element, optionally filtering by a condition
     * @param c Optional UIA condition object
     * @param scope Optional UIA.TreeScope value: Element, Children, Descendants, Subtree (=Element+Descendants). Default is Children.
     * @returns {UIA.IUIAutomationElementArray}
     */
	GetChildren(c?, scope:=0x2) => this.FindAll(c ?? UIA.TrueCondition, scope)

	; Get all child elements using TreeWalker
	TWGetChildren() { 
		arr := []
		if !IsObject(nextChild := UIA.TreeWalkerTrue.GetFirstChildElement(this))
			return ""
		arr.Push(nextChild)
		while IsObject(nextChild := UIA.TreeWalkerTrue.GetNextSiblingElement(nextChild))
			arr.Push(nextChild)
		return arr
	}
	/**
     * Returns info about the element: ControlType, Name, Value, LocalizedControlType, AutomationId, AcceleratorKey. 
     * @param scope Optional UIA.TreeScope value: Element, Children, Descendants, Subtree (=Element+Descendants). Default is Element.
     * @param maxDepth Optional maximal depth of tree levels traversal. Default is unlimited.
     * @returns {String}
     */
	Dump(scope:=1, maxDepth:=-1) { 
        out := ""
        if scope&1
            out := "Type: " (ctrlType := this.ControlType) " (" UIA.ControlType.%ctrlType% ")" ((name := this.Name) == "" ? "" : " Name: `"" name "`"") ((val := this.Value) == "" ? "" : " Value: `"" val "`"") ((lct := this.LocalizedControlType) == "" ? "" : " LocalizedControlType: `"" lct "`"") ((aid := this.AutomationId) == "" ? "" : " AutomationId: `"" aid "`"") ((cm := this.ClassName) == "" ? "" : " ClassName: `"" cm "`"") ((ak := this.AcceleratorKey) == "" ? "" : " AcceleratorKey: `"" ak "`"") "`n"
        if scope&4
            return RTrim(RecurseTree(this, out), "`n")
        if scope&2 {
            for n, oChild in this.Children
                out .= n ": " oChild.Dump() "`n"
        }
        return RTrim(out, "`n")

        RecurseTree(element, tree, path:="") {
            if maxDepth > 0 {
                StrReplace(path, "," , , , &count)
                if count >= (maxDepth-1)
                    return tree
            }
            For i, child in element.GetChildren() {
                tree .= path (path?",":"") i ": " child.Dump() "`n"
                tree := RecurseTree(child, tree, path (path?",":"") i)
            }
            return tree
        }
	}
    ToString() => this.Dump()
	/**
     * Returns info about the element and its descendants: ControlType, Name, Value, LocalizedControlType, AutomationId, AcceleratorKey. 
     * @param maxDepth Optional maximal depth of tree levels traversal. Default is unlimited.
     * @returns {String}
     */
    DumpAll(maxDepth:=-1) => this.Dump(5, maxDepth)
	CachedDump() { 
		return "Type: " (ctrlType := this.CachedControlType) " (" UIA.ControlType.%ctrlType% ")" ((name := this.CachedName) == "" ? "" : " Name: `"" name "`"") ((val := this.CachedValue) == "" ? "" : " Value: `"" val "`"") ((lct := this.CachedLocalizedControlType) == "" ? "" : " LocalizedControlType: `"" lct "`"") ((aid := this.CachedAutomationId) == "" ? "" : " AutomationId: `"" aid "`"") ((cm := this.CachedClassName) == "" ? "" : " ClassName: `"" cm "`"") ((ak := this.CachedAcceleratorKey) == "" ? "" : " AcceleratorKey: `"" ak "`"")
	}
    /**
     * @param relativeTo CoordMode to be used: client, window or screen. Default is A_CoordModeMouse
     * @returns {x:x coordinate, y:y coordinate, w:width, h:height}
     */
	GetPos(relativeTo:="") { 
		relativeTo := (relativeTo == "") ? A_CoordModeMouse : relativeTo
		br := this.BoundingRectangle
		if (relativeTo = "screen")
			return {x:br.l, y:br.t, w:(br.r-br.l), h:(br.b-br.t)}
		else if (relativeTo = "window") {
            DllCall("user32\GetWindowRect", "Int", this.GetWinId(), "Ptr", RECT := Buffer(16))
            return {x:(br.l-NumGet(RECT, 0, "Int")), y:(br.t-NumGet(RECT, 4, "Int")), w:br.r-br.l, h:br.b-br.t}
        } else if (relativeTo = "client") {
            pt := Buffer(8), NumPut("int",br.l,pt), NumPut("int", br.t,pt,4)
            DllCall("ScreenToClient", "Int", this.GetWinId(), "Ptr", pt)
            return {x:NumGet(pt,0,"int"), y:NumGet(pt,4,"int"), w:br.r-br.l, h:br.b-br.t}
        } else
            throw Error(relativeTo "is not a valid CoordMode",-1)		
	}

	; Get the parent window hwnd from the element
	GetWinId() { 
        static TW := UIA.CreateTreeWalker(UIA.CreateNotCondition(UIA.CreatePropertyCondition(UIA.Property.NativeWindowHandle, 0)))
        try {
            hwnd := TW.NormalizeElement(this).GetCurrentPropertyValue(UIA.Property.NativeWindowHandle)
            return DllCall("GetAncestor", "UInt", hwnd, "UInt", 2) ; hwnd from point by SKAN
        }
	}

    /**
     * Tries to click the element. The method depends on WhichButton variable: by default it is attempted
     * to use any "click"-like methods, such as InvokePattern Invoke(), TogglePattern Toggle(), SelectionItemPattern Select().
     * @param WhichButton If WhichButton is left empty (default), then any "click"-like pattern methods
     *     will be used (Invoke(), Toggle(), Select() etc. If WhichButton is a number, then Sleep 
     *     will be called afterwards with that number of milliseconds. 
     *     Eg. Element.Click(200) will sleep 200ms after "clicking".
     * If WhichButton is "left" or "right", then the native Click() will be used to move the cursor to
     *     the center of the element and perform a click.
     * @param ClickCount Is used if WhichButton isn't a number or left empty, that is if AHK Click()
     * will be used. In this case if ClickCount is a number <10, then that number of clicks will be performed.
     * If ClickCount is >=10, then Sleep will be called with that number of ms. Both ClickCount and sleep time
     * can be combined by separating with a space.
     * Eg. Element.Click("left", 1000) will sleep 1000ms after clicking.
     *     Element.Click("left", 2) will double-click the element
     *     Element.Click("left" "2 1000") will double-click the element and then sleep for 1000ms
     * @param DownOrUp If AHK Click is used, then this will either press the mouse down, or release it.
     * @param Relative If Relative is "Rel" or "Relative" then X and Y coordinates are treated as offsets from the current mouse position. 
     * Otherwise it expects offset values for both X and Y (eg "-5 10" would offset X by -5 and Y by +10 from the center of the element).
     * @param NoActivate If AHK Click is used, then this will determine whether the window is activated
     * before clicking if the clickable point isn't visible on the screen. Default is no activating.
     */
    Click(WhichButton:="", ClickCount:=1, DownOrUp:="", Relative:="", NoActivate:=False) {
        if WhichButton = "" or IsInteger(WhichButton) {
            SleepTime := WhichButton ? WhichButton : -1
			if (this.GetCurrentPropertyValue(UIA.Property.IsInvokePatternAvailable)) {
				this.InvokePattern.Invoke()
				Sleep SleepTime
				return 1
			}
			if (this.GetCurrentPropertyValue(UIA.Property.IsTogglePatternAvailable)) {
				togglePattern := this.GetCurrentPatternAs("Toggle"), toggleState := togglePattern.CurrentToggleState
				togglePattern.Toggle()
				if (togglePattern.CurrentToggleState != toggleState) {
					Sleep sleepTime
					return 1
				}
			}
			if (this.GetCurrentPropertyValue(UIA.Property.IsExpandCollapsePatternAvailable)) {
				if ((expandState := (pattern := this.ExpandCollapsePattern).ExpandCollapseState) == 0)
					pattern.Expand()
				Else
					pattern.Collapse()
				if (pattern.ExpandCollapseState != expandState) {
					Sleep sleepTime
					return 1
				}
			} 
			if (this.GetCurrentPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)) {
				selectionPattern := this.SelectionItemPattern, selectionState := selectionPattern.IsSelected
				selectionPattern.Select()
				if (selectionPattern.IsSelected != selectionState) {
					Sleep sleepTime
					return 1
				}
			}
			if (this.GetCurrentPropertyValue(UIA.Property.IsLegacyIAccessiblePatternAvailable)) {
				this.LegacyIAccessiblePattern.DoDefaultAction()
				Sleep sleepTime
				return 1
			}
			return 0
        }	
        rel := [0,0], pos := this.Location, cCount := 1, SleepTime := -1
        if (Relative && !InStr(Relative, "rel"))
            rel := StrSplit(Relative, " "), Relative := ""
        if IsInteger(WhichButton)
            SleepTime := WhichButton, WhichButton := "left"
        if !IsInteger(ClickCount) && InStr(ClickCount, " ") {
            sCount := StrSplit(ClickCount, " ")
            cCount := sCount[1], SleepTime := sCount[2]
        } else if ClickCount > 9 {
            SleepTime := cCount, cCount := 1
        }
        if (!NoActivate && (UIA.WindowFromPoint(pos.x+pos.w//2+rel[1], pos.y+pos.h//2+rel[2]) != (wId := this.GetWinId()))) {
            WinActivate(wId)
            WinWaitActive(wId)
        }
        saveCoordMode := A_CoordModeMouse
        CoordMode("Mouse", "Screen")
        Click(pos.x+pos.w//2+rel[1] " " pos.y+pos.h//2+rel[2] " " WhichButton (ClickCount ? " " ClickCount : "") (DownOrUp ? " " DownOrUp : "") (Relative ? " " Relative : ""))
        CoordMode("Mouse", saveCoordMode)
        Sleep(SleepTime)
    }

    ; ControlClicks the element after getting relative coordinates with GetLocation("client"). 
    ; 
    /**
     * Uses AHK ControlClick to click the element.
     * @param WhichButton determines which button to use to click (left, right, middle).0
     * If WhichButton is a number, then a Sleep will be called afterwards. 
     * Eg. ControlClick(200) will sleep 200ms after clicking. 
     * @param ClickCount How many times to click. Default is 1.
     * @param Options Additional ControlClick Options (see AHK documentations).
     */
    ControlClick(WhichButton:="left", ClickCount:=1, Options:="") { 
        pos := this.GetPos("client")
        ControlClick("X" pos.x+pos.w//2 " Y" pos.y+pos.h//2, this.GetWinId(),, IsInteger(WhichButton) ? "left" : WhichButton, ClickCount, Options)
        if IsInteger(WhichButton)
            Sleep(WhichButton)
    }
    /**
     * Highlights the element for a chosen period of time.
     * @param showTime Can be one of the following:
     *     Unset - removes the highlighting. This is the default value.
     *     0 - Indefinite highlighting
     *     Positive integer (eg 2000) - will highlight and pause for the specified amount of time in ms
     *     Negative integer - will highlight for the specified amount of time in ms, but script execution will continue
     * @param color The color of the highlighting. Default is red.
     * @param d The border thickness of the highlighting in pixels. Default is 2.
     * @returns {UIA.IUIAutomationElement}
     */
    Highlight(showTime:=unset, color:="Red", d:=2) {
        if !this.HasOwnProp("HighlightGui")
            this.HighlightGui := []
        if !IsSet(showTime) {
            for _, r in this.HighlightGui
                r.Destroy()
            this.HighlightGui := []
            return this
        }
        try loc := this.BoundingRectangle
        if !IsSet(loc) || !IsObject(loc)
            return this
        Loop 4 {
            this.HighlightGui.Push(Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale +E0x08000000"))
        }
        Loop 4
        {
            i:=A_Index
            , x1:=(i=2 ? loc.r : loc.l-d)
            , y1:=(i=3 ? loc.b : loc.t-d)
            , w1:=(i=1 or i=3 ? (loc.r-loc.l)+2*d : d)
            , h1:=(i=2 or i=4 ? (loc.b-loc.t)+2*d : d)
            this.HighlightGui[i].BackColor := color
            this.HighlightGui[i].Show("NA x" . x1 . " y" . y1 . " w" . w1 . " h" . h1)
        }
        if showTime > 0 {
            Sleep(showTime)
            this.Highlight()
        } else if showTime < 0
            SetTimer(ObjBindMethod(this, "Highlight"), -Abs(showTime))
        return this
    }

	/*
		FindFirst using search criteria. 
		expr: 
			Takes a value in the form of "PropertyId=matchvalue" to match a specific property with the value matchValue. PropertyId can be most properties from UIA_Enum.UIA_PropertyId method (for example Name, ControlType, AutomationId etc). 
			
			Example1: "Name=Username:" would use FindFirst with UIA_Enum.UIA_NamePropertyId matching the name "Username:"
			Example2: "ControlType=Button would FindFirst using UIA_Enum.UIA_ControlTypePropertyId and matching for UIA_Enum.UIA_ButtonControlTypeId. Alternatively "ControlType=50000" can be used (direct value for UIA_ButtonControlTypeId which is 50000)
			
			Criteria can be combined with AND, OR, &&, ||:
			Example3: "Name=Username: AND ControlType=Button" would FindFirst an element with the name property of "Username:" and control type of button.

			Flags can be modified for each individual condition by specifying FLAGS=n after the condition (and before and/or operator). 0=no flags; 1=ignore case (case insensitive matching); 2=match substring; 3=ignore case and match substring

			If matchMode==3 or matching substrings is supported (Windows 10 17763 and above) and matchMode==2, then parentheses are supported. 
			Otherwise parenthesis are not supported, and criteria are evaluated left to right, so "a AND b OR c" would be evaluated as "(a and b) or c".
			
			Negation can be specified with NOT:
			Example4: "NOT ControlType=Edit" would return the first element that is not an edit element
		
		scope:
			Scope by default is UIA_TreeScope_Descendants. 
			
		matchMode:
			If using Name PropertyId as a criteria, this follows the SetTitleMatchMode scheme: 
				1=name must must start with the specified name
				2=can contain anywhere
				3=exact match
				RegEx=using regular expression. In this case the Name can't be empty.
		
		caseSensitive:
			If matching for a string, this will specify case-sensitivity.

	*/
    /*
	FindFirstBy(expr, scope:=0x4, matchMode:=3, caseSensitive:=True, cacheRequest:="") { 
		static MatchSubstringSupported := !InStr(A_OSVersion, "WIN") && (StrSplit(A_OSVersion, ".")[3] >= 17763)
		if ((matchMode == 3) || (matchMode==2 && MatchSubstringSupported)) {
			return this.FindFirst(UIA.CreateCondition(expr, ((matchMode==2)?2:0)|!caseSensitive), scope, cacheRequest)
		}
		pos := 1, match := "", createCondition := "", operator := "", bufName := []
		while (pos := RegexMatch(expr, "i) *(NOT|!)? *(\w+?) *=(?: *(\d+|'.*?(?<=[^\\]|[^\\]\\\\)')|(.*?))(?: FLAGS=(\d))?( AND | OR | && | \|\| |$)", &match, pos+StrLen(match))) {
			if !match
				break
			if ((StrLen(match[3]) > 1) && (SubStr(match[3],1,1) == "'") && (SubStr(match[3],0,1) == "'"))
				match[3] := StrReplace(RegexReplace(SubStr(match[3],2,StrLen(match[3])-2), "(?<=[^\\]|[^\\]\\\\)\\'", "'"), "\\", "\") ; Remove single quotes and escape characters
			else if match[4]
				match[3] := match[4]
			if ((isNamedProperty := RegexMatch(match[2], "i)Name|AutomationId|Value|ClassName|FrameworkId")) && !bufName[1] && ((matchMode != 2) || ((matchMode == 2) && !MatchSubstringSupported)) && (matchMode != 3)) { ; Check if we have a condition that needs FindAll to be searched, and save it. Apply only for the first one encountered.
				bufName[1] := (match[1] ? "NOT " : "") match[2], bufName[2] := match[3], bufName[3] := match[5]
				Continue
			}
			newCondition := UIA.CreateCondition(match[2], match[3], match[5] ? match[5] : ((((matchMode==2) && isNamedProperty)?2:0)|!caseSensitive))
			if match[1]
				newCondition := UIA.CreateNotCondition(newCondition)
			fullCondition := (operator == " AND " || operator == " && ") ? UIA.CreateAndCondition(fullCondition, newCondition) : (operator == " OR " || operator == " || ") ? UIA.CreateOrCondition(fullCondition, newCondition) : newCondition
			operator := match[6] ; Save the operator to be used in the next loop
		}
		if (bufName[1]) { ; If a special case was encountered requiring FindAll
			notProp := InStr(bufName[1], "NOT "), property := StrReplace(StrReplace(bufName[1], "NOT "), "Current"), value := bufName[2], caseSensitive := bufName[3] ? !(bufName[3]&1) : caseSensitive 
			if (property = "value")
				property := "ValueValue"
			if (MatchSubstringSupported && (matchMode==1)) { ; Check if we can speed up the search by limiting to substrings when matchMode==1
				propertyCondition := UIA.CreatePropertyConditionEx(UIA.Property.%property%, value, 2|!caseSensitive)
				if notProp
					propertyCondition := UIA.CreateNotCondition(propertyCondition)
			} else 
				propertyCondition := UIA.CreateNotCondition(UIA.CreatePropertyCondition(UIA.Property.%property%, ""))
			fullCondition := IsObject(fullCondition) ? UIA.CreateAndCondition(propertyCondition, fullCondition) : propertyCondition
			for _, element in this.FindAll(fullCondition, scope, cacheRequest) {
				curValue := element.%property%
				if notProp {
					if (((matchMode == 1) && !InStr(SubStr(curValue, 1, StrLen(value)), value, caseSensitive)) || ((matchMode == 2) && !InStr(curValue, value, caseSensitive)) || (InStr(matchMode, "RegEx") && !RegExMatch(curValue, value)))
						return element
				} else {
					if (((matchMode == 1) && InStr(SubStr(curValue, 1, StrLen(value)), value, caseSensitive)) || ((matchMode == 2) && InStr(curValue, value, caseSensitive)) || (InStr(matchMode, "RegEx") && RegExMatch(curValue, value)))
						return element
				}
			}
		} else {
			return this.FindFirst(fullCondition, scope, cacheRequest)
		}
	}
    */

    /**
     * Wait element to exist.
     * @param condition Property condition(s) or path for the search (see UIA.PropertyCondition).
     * @param timeout Waiting time for element to appear. Default: indefinite wait
     * @param scope TreeScope search scope for the search
     * @returns Found element if successful, 0 if timeout happens
     */
    WaitExist(condition, timeout := -1, scope := 4) {
        endtime := A_TickCount + timeout
        While ((timeout == -1) || (A_Tickcount < endtime)) && !(el := (IsObject(condition) ? this.FindFirst(condition, scope) : this.FindByPath(condition)))
            Sleep 20
        return el
    }

    /**
     * Wait element to not exist (disappear).
     * @param condition Property condition(s) or path for the search (see UIA.PropertyCondition).
     * @param timeout Waiting time for element to disappear. Default: indefinite wait
     * @param scope TreeScope search scope for the search
     * @returns 1 if element disappeared, 0 otherwise
     */
    WaitNotExist(condition, timeout := -1, scope := 4) {
        endtime := A_TickCount + timeout
        While (timeout == -1) || (A_Tickcount < endtime) {
            if !(IsObject(condition) ? this.FindFirst(condition, scope) : this.FindByPath(condition))
                return 1
        }
        return 0
    }

    /**
     * Find or wait target control element.
     * @param ControlType target control type, such as 'Button' or UIA.ControlType.Button
     *     Can also be used to target nth element of the type: Button2 will find the second button matching the criteria.
     * @param properties The condition object or 'Name' property.
     * @param timeOut Waiting time for control element to appear. Default: no waiting.
     * @param scope TreeScope search scope for the search
     * @returns {IUIAutomationElement}
     */
    FindElement(ControlType, properties := unset, timeOut := 0, scope := 4) {
        index := 1
        if !IsNumber(ControlType) {
            if RegExMatch(ControlType, "i)^([a-z]+) *(\d+)$", &m)
                index := Integer(m[2]), ControlType := m[1]
            try ControlType := UIA.ControlType.%ControlType%
            catch
                throw ValueError("Invalid ControlType")
        }
        properties := IsSet(properties) ? (Type(properties) = "Object" ? properties.Clone() : properties) : {}
        switch Type(properties) {
            case "String":
                properties := { Name: properties }
            case "Array":
                properties := { or: properties }
            case "Object":
            default:
                throw ValueError("'properties' can only be a String, Array, or Object!")
        }
        properties.ControlType := ControlType
        cond := UIA.CreateCondition(properties)
        endtime := A_TickCount + timeOut
        loop {
            try {
                if (index = 1)
                    return this.FindFirst(cond, scope)
                else {
                    els := this.FindAll(cond, scope)
                    if (index <= els.Length)
                        return els.GetElement(index)
                    throw TargetError("Target element not found.")
                }
            } catch TargetError {
                if (A_TickCount > endtime)
                    return
            }
        }
    }

	FindByPath(searchPath:="", c?) {
		el := this, PathTW := (IsSet(c) ? UIA.CreateTreeWalker(c) : UIA.TreeWalkerTrue)
		searchPath := StrReplace(StrReplace(searchPath, " "), ".", ",")
		Loop Parse searchPath, "," {
			if IsInteger(A_LoopField) {
				children := el.GetChildren(c?, 0x2)
				if !children || !children.Length > A_LoopField
					throw ValueError("Step " A_index " was out of bounds")
                el := children[A_LoopField]
			} else if RegexMatch(A_LoopField, "i)([\w+-]+)(\d+)?", &m:="") {
                if !m[2]
                    m[2] := 1
				if m[1] = "p" {
					Loop m[2] {
						if !(el := PathTW.GetParentElement(el))
							throw ValueError("Step " A_index " with P" m[2] " was out of bounds (GetParentElement failed)")
					}
				} else if m[1] = "+" || m[1] = "-" {
					if (m[1] == "+") {
						Loop m[2] {
							if !(el := PathTW.GetNextSiblingElement(el))
								throw ValueError("Step " A_index " with `"" m[1] m[2] "`" was out of bounds (GetNextSiblingElement failed)")
						}
					} else if (m[1] == "-") {
						Loop m[2] {
							if !(el := PathTW.GetPreviousSiblingElement(el))
								throw ValueError("Step " A_index " with `"" m[1] m[2] "`" was out of bounds (GetPreviousSiblingElement failed)")
						}
					}
				} else if UIA.ControlType.HasOwnProp(m[1]) {
                    el := el.FindFirst({Type:m[1], i:m[2]}, 2)
                }
			}
		}
		return el
	}

    GetAllPropertyValues() {
        infos := {}
        for k, v in UIA.Property.OwnProps() {
            v := this.GetPropertyValue(v)
            if (v is ComObjArray) {
                arr := []
                for t in v
                    arr.Push(t)
                v := arr
            }
            infos.%k% := v
        }
        return infos
    }

    GetControlID() {
        cond := UIA.CreatePropertyCondition(UIA.Property.ControlType, controltype := this.GetPropertyValue(UIA.Property.ControlType))
        runtimeid := UIA.RuntimeIdToString(this.GetRuntimeId())
        runtimeid := RegExReplace(runtimeid, "^(\w+\.\w+)\..*$", "$1")
        rootele := UIA.GetRootElement().FindFirst(UIA.CreatePropertyCondition(UIA.Property.RuntimeId, UIA.RuntimeIdFromString(runtimeid)))
        for el in rootele.FindAll(cond) {
            if (UIA.CompareElements(this, el))
                return UIA.ControlType[controltype] A_Index
        }
    }

    ValidateCondition(obj) {
        mm := 3, cs := 1
        switch Type(obj) {
            case "Object":
                mm := obj.HasOwnProp("matchmode") ? obj.matchmode : obj.HasOwnProp("mm") ? obj.mm : 3
                cs := obj.HasOwnProp("casesensitive") ? obj.casesensitive : obj.HasOwnProp("cs") ? obj.cs : 1
                obj := obj.OwnProps()
                operator := "and"
            case "Array":
                operator := "or"
                for k, v in obj {
                    if IsObject(v) && this.ValidateCondition(v)
                        return true
                }
                return false
        }
        for k, v in obj {
            if IsObject(v) { 
                if (k = "not" ? this.ValidateCondition(v) : !this.ValidateCondition(v))
                    return False
                continue
            }
            try k := IsInteger(k) ? Integer(k) : UIA.Property.%k%
            if !IsInteger(k)
                continue
            if k = 30003 && !IsInteger(v)
                try v := UIA.ControlType.%v%
            prop := ""
            try prop := UIA.Property[k]
            currentValue := this.GetPropertyValue(k)
            switch mm, False {
                case "RegEx":
                    if !RegExMatch(currentValue, v)
                        return False
                case 1:
                    if !((cs && SubStr(currentValue, 1, StrLen(v)) == v) || (!cs && SubStr(currentValue, 1, StrLen(v)) = v))
                        return False
                case 2:
                    if !InStr(currentValue, v, cs)
                        return False
                case 3:
                    if !(cs ? currentValue == v : currentValue = v)
                        return False
            }
        }
        return True
    }

    ; Sets the keyboard focus to this UI Automation element.
    SetFocus() => ComCall(3, this)

    ; Retrieves the unique identifier assigned to the UI element.
    ; The identifier is only guaranteed to be unique to the UI of the desktop on which it was generated. Identifiers can be reused over time.
    ; The format of run-time identifiers might change in the future. The returned identifier should be treated as an opaque value and used only for comparison; for example, to determine whether a Microsoft UI Automation element is in the cache.
    GetRuntimeId() => (ComCall(4, this, "ptr*", &runtimeId := 0), ComValue(0x2003, runtimeId))

    ; The scope of the search is relative to the element on which the method is called. Elements are returned in the order in which they are encountered in the tree.
    ; This function cannot search for ancestor elements in the Microsoft UI Automation tree; that is, TreeScope_Ancestors is not a valid value for the scope parameter.
    ; When searching for top-level windows on the desktop, be sure to specify TreeScope_Children in the scope parameter, not TreeScope_Descendants. A search through the entire subtree of the desktop could iterate through thousands of items and lead to a stack overflow.
    ; If your client application might try to find elements in its own user interface, you must make all UI Automation calls on a separate thread.

    ; Retrieves the first child or descendant element that matches the specified condition.
    FindFirst(condition, scope := 4, cacheRequest?) {
        if !InStr(Type(condition), "Condition") {
            index := condition.HasOwnProp("index") ? condition.index : condition.HasOwnProp("i") ? condition.i : 1
            /*
                If MatchMode is 1:
                    Use MatchMode 2, but then return all the results and filter using the condition
                If MatchMode is RegEx:
                    Find all !"" elements, then filter using the condition
            */
            IUIAcondition := UIA.__ConditionBuilder(condition, &sanitized:=False)
            if sanitized {
                unfilteredEls := this.FindAll(IUIAcondition, scope), counter := 0
                for el in unfilteredEls {
                    if el.ValidateCondition(condition) && ++counter = index {
                        return el
                    }
                }
                return ""
            }
            condition := IUIAcondition
            ; Were any conditions encountered where we need to filter conditions?
            if index != 1 {
                try return (IsSet(cacheRequest) ? this.FindAllBuildCache(condition, scope, cacheRequest) : this.FindAll(condition, scope)).GetElement(index)
                catch
                    return ""
            }
        }
        if IsSet(cacheRequest)
            return this.FindFirstBuildCache(condition, scope, cacheRequest)
        else if (ComCall(5, this, "int", scope, "ptr", condition, "ptr*", &found := 0), found)
            return UIA.IUIAutomationElement(found)
        return ""
    }

    ; Returns all UI Automation elements that satisfy the specified condition.
    FindAll(condition, scope := 4, cacheRequest?) {
        if !InStr(Type(condition), "Condition") {
            IUIAcondition := UIA.__ConditionBuilder(condition, &sanitized:=False)
            if sanitized {
                unfilteredEls := this.FindAll(IUIAcondition, scope), filteredEls := []
                for el in unfilteredEls {
                    if el.ValidateCondition(condition)
                        filteredEls.Push(el)
                }
                return filteredEls
            }
            condition := IUIAcondition
        }
        if IsSet(cacheRequest)
            return this.FindAllBuildCache(condition, scope, cacheRequest)
        else if (ComCall(6, this, "int", scope, "ptr", condition, "ptr*", &found := 0), found)
            return UIA.IUIAutomationElementArray(found)
        return ""
    }

    ; Retrieves the first child or descendant element that matches the specified condition, prefetches the requested properties and control patterns, and stores the prefetched items in the cache.
    FindFirstBuildCache(condition, scope := 4, cacheRequest?) {
        if !InStr(Type(condition), "Condition")
            condition := UIA.CreateCondition(condition)
        if (ComCall(7, this, "int", scope, "ptr", condition, "ptr", cacheRequest, "ptr*", &found := 0), found)
            return UIA.IUIAutomationElement(found)
        return ""
    }

    ; Returns all UI Automation elements that satisfy the specified condition, prefetches the requested properties and control patterns, and stores the prefetched items in the cache.
    FindAllBuildCache(condition, scope := 4, cacheRequest?) {
        if !InStr(Type(condition), "Condition")
            condition := UIA.CreateCondition(condition)
        if (ComCall(8, this, "int", scope, "ptr", condition, "ptr", cacheRequest, "ptr*", &found := 0), found)
            return UIA.IUIAutomationElementArray(found)
        return ""
    }

    ; Retrieves a  UI Automation element with an updated cache.
    ; The original UI Automation element is unchanged. The  UIA.IUIAutomationElement interface refers to the same element and has the same runtime identifier.
    BuildUpdatedCache(cacheRequest) => (ComCall(9, this, "ptr", cacheRequest, "ptr*", &updatedElement := 0), UIA.IUIAutomationElement(updatedElement))

    ; Microsoft UI Automation properties of the double type support Not a Number (NaN) values. When retrieving a property of the double type, a client can use the _isnan function to determine whether the property is a NaN value.

    ; Retrieves the current value of a property for this UI Automation element.
    GetPropertyValue(propertyId) {
		if !IsNumber(propertyId)
			try propertyId := UIA.Property.%propertyId%
        ComCall(10, this, "int", propertyId, "ptr", val := UIA.ComVar())
        return val[]
    } 

    ; Retrieves a property value for this UI Automation element, optionally ignoring any default value.
    ; Passing FALSE in the ignoreDefaultValue parameter is equivalent to calling UIA.IUIAutomationElement,,GetPropertyValue.
    ; If the Microsoft UI Automation provider for the element itself supports the property, the value of the property is returned. Otherwise, if ignoreDefaultValue is FALSE, a default value specified by UI Automation is returned.
    ; This method returns a failure code if the requested property was not previously cached.
    GetPropertyValueEx(propertyId, ignoreDefaultValue) => (ComCall(11, this, "int", propertyId, "int", ignoreDefaultValue, "ptr", val := UIA.ComVar()), val[])

    ; Retrieves a property value from the cache for this UI Automation element.
    GetCachedPropertyValue(propertyId) => (ComCall(12, this, "int", propertyId, "ptr", val := UIA.ComVar()), val[])

    ; Retrieves a property value from the cache for this UI Automation element, optionally ignoring any default value.
    GetCachedPropertyValueEx(propertyId, ignoreDefaultValue, retVal) => (ComCall(13, this, "int", propertyId, "int", ignoreDefaultValue, "ptr", val := UIA.ComVar()), val[])

    ; Retrieves the control pattern interface of the specified pattern on this UI Automation element.
    GetPatternAs(patternId, riid) {	; not completed
        if IsNumber(patternId)
            name := UIA.Pattern.%patternId%
        else
            patternId := UIA.Pattern.%(name := patternId)%
        ComCall(14, this, "int", patternId, "ptr", riid, "ptr*", &patternObject := 0)
        return UIA.IUIAutomation%name%Pattern(patternObject)
    }

    ; Retrieves the control pattern interface of the specified pattern from the cache of this UI Automation element.
    GetCachedPatternAs(patternId, riid) {	; not completed
        try {
            if IsNumber(patternId)
                name := UIA.Pattern.%patternId%
            else
                patternId := UIA.Pattern.%(name := patternId)%
            ComCall(15, this, "int", patternId, "ptr", riid, "ptr*", &patternObject := 0)
            return UIA.IUIAutomation%name%Pattern(patternObject)
        }
    }

    ; Retrieves the IUnknown interface of the specified control pattern on this UI Automation element.
    ; This method gets the specified control pattern based on its availability at the time of the call.
    ; For some forms of UI, this method will incur cross-process performance overhead. Applications can reduce overhead by caching control patterns and then retrieving them by using UIA.IUIAutomationElement,,GetCachedPattern.
    GetPattern(patternId) {
        try {
            if IsNumber(patternId)
                name := UIA.Pattern.%patternId%
            else
                patternId := UIA.Pattern.%(name := StrReplace(patternId, "Pattern"))%
            
            ComCall(16, this, "int", patternId, "ptr*", &patternObject := 0)
            return UIA.IUIAutomation%RegExReplace(name, "\d+$")%Pattern(patternObject)
        } catch
            Throw Error("Failed to get pattern `"" name "`n!", -1)
    }

    ; Retrieves from the cache the IUnknown interface of the specified control pattern of this UI Automation element.
    GetCachedPattern(patternId) {
        try {
            if IsNumber(patternId)
                name := UIA.Pattern.%patternId%
            else
                patternId := UIA.Pattern.%(name := patternId)%
            ComCall(17, this, "int", patternId, "ptr*", &patternObject := 0)
            return UIA.IUIAutomation%name%Pattern(patternObject)
        }
    }

    ; Retrieves from the cache the parent of this UI Automation element.
    GetCachedParent() => (ComCall(18, this, "ptr*", &parent := 0), UIA.IUIAutomationElement(parent))

    ; Retrieves the cached child elements of this UI Automation element.
    ; The view of the returned collection is determined by the TreeFilter property of the IUIAutomationCacheRequest that was active when this element was obtained.
    ; Children are cached only if the scope of the cache request included TreeScope_Subtree, TreeScope_Children, or TreeScope_Descendants.
    ; If the cache request specified that children were to be cached at this level, but there are no children, the value of this property is 0. However, if no request was made to cache children at this level, an attempt to retrieve the property returns an error.
    GetCachedChildren() => (ComCall(19, this, "ptr*", &children := 0), UIA.IUIAutomationElementArray(children))

    ; Retrieves the identifier of the process that hosts the element.
    ProcessId => (ComCall(20, this, "int*", &retVal := 0), retVal)

    ; Retrieves the control type of the element.
    ; Control types describe a known interaction model for UI Automation elements without relying on a localized control type or combination of complex logic rules. This property cannot change at run time unless the control supports the IUIAutomationMultipleViewPattern interface. An example is the Win32 ListView control, which can change from a data grid to a list, depending on the current view.
    ControlType => (ComCall(21, this, "int*", &retVal := 0), retVal)

    ; Retrieves a localized description of the control type of the element.
    LocalizedControlType => (ComCall(22, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves the name of the element.
    Name => (ComCall(23, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves the accelerator key for the element.
    AcceleratorKey => (ComCall(24, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves the access key character for the element.
    ; An access key is a character in the text of a menu, menu item, or label of a control such as a button that activates the attached menu function. For example, the letter "O" is often used to invoke the Open file common dialog box from a File menu. Microsoft UI Automation elements that have the access key property set always implement the Invoke control pattern.
    AccessKey => (ComCall(25, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Indicates whether the element has keyboard focus.
    HasKeyboardFocus => (ComCall(26, this, "int*", &retVal := 0), retVal)

    ; Indicates whether the element can accept keyboard focus.
    IsKeyboardFocusable => (ComCall(27, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates whether the element is enabled.
    IsEnabled => (ComCall(28, this, "int*", &retVal := 0), retVal)

    ; Retrieves the Microsoft UI Automation identifier of the element.
    ; The identifier is unique among sibling elements in a container, and is the same in all instances of the application.
    AutomationId => (ComCall(29, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves the class name of the element.
    ; The value of this property is implementation-defined. The property is useful in testing environments.
    ClassName => (ComCall(30, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves the help text for the element. This information is typically obtained from tooltips.
    ; Caution  Do not retrieve the CachedHelpText property from a control that is based on the SysListview32 class. Doing so could cause the system to become unstable and data to be lost. A client application can discover whether a control is based on SysListview32 by retrieving the CachedClassName or ClassName property from the control.
    HelpText => (ComCall(31, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves the culture identifier for the element.
    Culture => (ComCall(32, this, "int*", &retVal := 0), retVal)

    ; Indicates whether the element is a control element.
    IsControlElement => (ComCall(33, this, "int*", &retVal := 0), retVal)

    ; Indicates whether the element is a content element.
    ; A content element contains data that is presented to the user. Examples of content elements are the items in a list box or a button in a dialog box. Non-content elements, also called peripheral elements, are typically used to manipulate the content in a composite control; for example, the button on a drop-down control.
    IsContentElement => (ComCall(34, this, "int*", &retVal := 0), retVal)

    ; Indicates whether the element contains a disguised password.
    ; This property enables applications such as screen-readers to determine whether the text content of a control should be read aloud.
    IsPassword => (ComCall(35, this, "int*", &retVal := 0), retVal)

    ; Retrieves the window handle of the element.
    NativeWindowHandle => (ComCall(36, this, "ptr*", &retVal := 0), retVal)

    ; Retrieves a description of the type of UI item represented by the element.
    ; This property is used to obtain information about items in a list, tree view, or data grid. For example, an item in a file directory view might be a "Document File" or a "Folder".
    ItemType => (ComCall(37, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Indicates whether the element is off-screen.
    IsOffscreen => (ComCall(38, this, "int*", &retVal := 0), retVal)

    ; Retrieves a value that indicates the orientation of the element.
    ; This property is supported by controls such as scroll bars and sliders that can have either a vertical or a horizontal orientation.
    Orientation => (ComCall(39, this, "int*", &retVal := 0), retVal)

    ; Retrieves the name of the underlying UI framework. The name of the UI framework, such as "Win32", "WinForm", or "DirectUI".
    FrameworkId => (ComCall(40, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Indicates whether the element is required to be filled out on a form.
    IsRequiredForForm => (ComCall(41, this, "int*", &retVal := 0), retVal)

    ; Retrieves the description of the status of an item in an element.
    ; This property enables a client to ascertain whether an element is conveying status about an item. For example, an item associated with a contact in a messaging application might be "Busy" or "Connected".
    ItemStatus => (ComCall(42, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves the coordinates of the rectangle that completely encloses the element, in screen coordinates.
    BoundingRectangle => (ComCall(43, this, "ptr", retVal := UIA.NativeArray(0, 4, "int")), { l: retVal[0], t: retVal[1], r: retVal[2], b: retVal[3] })

    ; This property maps to the Accessible Rich Internet Applications (ARIA) property.

    ; Retrieves the element that contains the text label for this element.
    ; This property could be used to retrieve, for example, the static text label for a combo box.
    LabeledBy => (ComCall(44, this, "ptr*", &retVal := 0), UIA.IUIAutomationElement(retVal))

    ; Retrieves the Accessible Rich Internet Applications (ARIA) role of the element.
    AriaRole => (ComCall(45, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves the ARIA properties of the element.
    AriaProperties => (ComCall(46, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Indicates whether the element contains valid data for a form.
    IsDataValidForForm => (ComCall(47, this, "int*", &retVal := 0), retVal)

    ; Retrieves an array of elements for which this element serves as the controller.
    ControllerFor => (ComCall(48, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))

    ; Retrieves an array of elements that describe this element.
    DescribedBy => (ComCall(49, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))

    ; Retrieves an array of elements that indicates the reading order after the current element.
    FlowsTo => (ComCall(50, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))

    ; Retrieves a description of the provider for this element.
    ProviderDescription => (ComCall(51, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves the cached ID of the process that hosts the element.
    CachedProcessId => (ComCall(52, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates the control type of the element.
    CachedControlType => (ComCall(53, this, "int*", &retVal := 0), retVal)

    ; Retrieves the cached localized description of the control type of the element.
    CachedLocalizedControlType => (ComCall(54, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves the cached name of the element.
    CachedName => (ComCall(55, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves the cached accelerator key for the element.
    CachedAcceleratorKey => (ComCall(56, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves the cached access key character for the element.
    CachedAccessKey => (ComCall(57, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; A cached value that indicates whether the element has keyboard focus.
    CachedHasKeyboardFocus => (ComCall(58, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates whether the element can accept keyboard focus.
    CachedIsKeyboardFocusable => (ComCall(59, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates whether the element is enabled.
    CachedIsEnabled => (ComCall(60, this, "int*", &retVal := 0), retVal)

    ; Retrieves the cached UI Automation identifier of the element.
    CachedAutomationId => (ComCall(61, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves the cached class name of the element.
    CachedClassName => (ComCall(62, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ;
    CachedHelpText => (ComCall(63, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves the cached help text for the element.
    CachedCulture => (ComCall(64, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates whether the element is a control element.
    CachedIsControlElement => (ComCall(65, this, "int*", &retVal := 0), retVal)

    ; A cached value that indicates whether the element is a content element.
    CachedIsContentElement => (ComCall(66, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates whether the element contains a disguised password.
    CachedIsPassword => (ComCall(67, this, "int*", &retVal := 0), retVal)

    ; Retrieves the cached window handle of the element.
    CachedNativeWindowHandle => (ComCall(68, this, "ptr*", &retVal := 0), retVal)

    ; Retrieves a cached string that describes the type of item represented by the element.
    CachedItemType => (ComCall(69, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves a cached value that indicates whether the element is off-screen.
    CachedIsOffscreen => (ComCall(70, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates the orientation of the element.
    CachedOrientation => (ComCall(71, this, "int*", &retVal := 0), retVal)

    ; Retrieves the cached name of the underlying UI framework associated with the element.
    CachedFrameworkId => (ComCall(72, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves a cached value that indicates whether the element is required to be filled out on a form.
    CachedIsRequiredForForm => (ComCall(73, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached description of the status of an item within an element.
    CachedItemStatus => (ComCall(74, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves the cached coordinates of the rectangle that completely encloses the element.
    CachedBoundingRectangle => (ComCall(75, this, "ptr", retVal := UIA.NativeArray(0, 4, "int")), { l: retVal[0], t: retVal[1], r: retVal[2], b: retVal[3] })

    ; Retrieves the cached element that contains the text label for this element.
    CachedLabeledBy => (ComCall(76, this, "ptr*", &retVal := 0), UIA.IUIAutomationElement(retVal))

    ; Retrieves the cached ARIA role of the element.
    CachedAriaRole => (ComCall(77, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves the cached ARIA properties of the element.
    CachedAriaProperties => (ComCall(78, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves a cached value that indicates whether the element contains valid data for the form.
    CachedIsDataValidForForm => (ComCall(79, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached array of UI Automation elements for which this element serves as the controller.
    CachedControllerFor => (ComCall(80, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))

    ; Retrieves a cached array of elements that describe this element.
    CachedDescribedBy => (ComCall(81, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))

    ; Retrieves a cached array of elements that indicate the reading order after the current element.
    CachedFlowsTo => (ComCall(82, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))

    ; Retrieves a cached description of the provider for this element.
    CachedProviderDescription => (ComCall(83, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves a point on the element that can be clicked.
    ; A client application can use this method to simulate clicking the left or right mouse button. For example, to simulate clicking the right mouse button to display the context menu for a control,
    ; • Call the GetClickablePoint method to find a clickable point on the control.
    ; • Call the SendInput function to send a right-mouse-down, right-mouse-up sequence.
    GetClickablePoint() {
        if (ComCall(84, this, "int64*", &clickable := 0, "int*", &gotClickable := 0), gotClickable)
            return { x: clickable & 0xffff, y: clickable >> 32 }
        throw TargetError('The element has no clickable point')
    }

    ;; UIA.IUIAutomationElement2
    OptimizeForVisualContent => (ComCall(85, this, "int*", &retVal := 0), retVal)
    CachedOptimizeForVisualContent => (ComCall(86, this, "int*", &retVal := 0), retVal)
    LiveSetting => (ComCall(87, this, "int*", &retVal := 0), retVal)
    CachedLiveSetting => (ComCall(88, this, "int*", &retVal := 0), retVal)
    FlowsFrom => (ComCall(89, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))
    CachedFlowsFrom => (ComCall(90, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))

    ;; UIA.IUIAutomationElement3
    ShowContextMenu() => ComCall(91, this)
    IsPeripheral => (ComCall(92, this, "int*", &retVal := 0), retVal)
    CachedIsPeripheral => (ComCall(93, this, "int*", &retVal := 0), retVal)

    ;; UIA.IUIAutomationElement4
    PositionInSet => (ComCall(94, this, "int*", &retVal := 0), retVal)
    SizeOfSet => (ComCall(95, this, "int*", &retVal := 0), retVal)
    Level => (ComCall(96, this, "int*", &retVal := 0), retVal)
    AnnotationTypes => (ComCall(97, this, "ptr*", &retVal := 0), ComValue(0x2003, retVal))
    AnnotationObjects => (ComCall(98, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))
    CachedPositionInSet => (ComCall(99, this, "int*", &retVal := 0), retVal)
    CachedSizeOfSet => (ComCall(100, this, "int*", &retVal := 0), retVal)
    CachedLevel => (ComCall(101, this, "int*", &retVal := 0), retVal)
    CachedAnnotationTypes => (ComCall(102, this, "ptr*", &retVal := 0), ComValue(0x2003, retVal))
    CachedAnnotationObjects => (ComCall(103, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))

    ;; UIA.IUIAutomationElement5
    LandmarkType => (ComCall(104, this, "int*", &retVal := 0), retVal)
    LocalizedLandmarkType => (ComCall(105, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    CachedLandmarkType => (ComCall(106, this, "int*", &retVal := 0), retVal)
    CachedLocalizedLandmarkType => (ComCall(107, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ;; UIA.IUIAutomationElement6
    FullDescription => (ComCall(108, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    CachedFullDescription => (ComCall(109, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ;; UIA.IUIAutomationElement7
    ; IUIAutomationCondition, TreeTraversalOptions, UIA.IUIAutomationElement, TreeScope
    FindFirstWithOptions(condition, traversalOptions:=0, root:=0, scope := 4) {
        if (ComCall(110, this, "int", scope, "ptr", InStr(Type(condition), "Condition") ? condition : UIA.CreateCondition(condition), "int", traversalOptions, "ptr", root, "ptr*", &found := 0), found)
            return UIA.IUIAutomationElement(found)
        throw TargetError("Target element not found.")
    }
    FindAllWithOptions(condition, traversalOptions:=0, root:=0, scope := 4) {
        if (ComCall(111, this, "int", scope, "ptr", InStr(Type(condition), "Condition") ? condition : UIA.CreateCondition(condition), "int", traversalOptions, "ptr", root, "ptr*", &found := 0), found)
            return UIA.IUIAutomationElementArray(found)
        throw TargetError("Target elements not found.")
    }

    ; TreeScope, IUIAutomationCondition, IUIAutomationCacheRequest, TreeTraversalOptions, UIA.IUIAutomationElement
    FindFirstWithOptionsBuildCache(condition, traversalOptions:=0, root:=0, scope := 4, cacheRequest?) {
        if (ComCall(112, this, "int", scope, "ptr", InStr(Type(condition), "Condition") ? condition : UIA.CreateCondition(condition), "ptr", cacheRequest ?? 0, "int", traversalOptions, "ptr", root, "ptr*", &found := 0), found)
            return UIA.IUIAutomationElement(found)
        throw TargetError("Target element not found.")
    }
    FindAllWithOptionsBuildCache(condition, traversalOptions:=0, root:=0, scope := 4, cacheRequest?) {
        if (ComCall(113, this, "int", scope, "ptr", InStr(Type(condition), "Condition") ? condition : UIA.CreateCondition(condition), "ptr", cacheRequest ?? 0, "int", traversalOptions, "ptr", root, "ptr*", &found := 0), found)
            return UIA.IUIAutomationElementArray(found)
        throw TargetError("Target elements not found.")
    }
    GetMetadataValue(targetId, metadataId) => (ComCall(114, this, "int", targetId, "int", metadataId, "ptr", returnVal := UIA.ComVar()), returnVal[])

    ;; UIA.IUIAutomationElement8
    HeadingLevel => (ComCall(115, this, "int*", &retVal := 0), retVal)
    CachedHeadingLevel => (ComCall(116, this, "int*", &retVal := 0), retVal)

    ;; UIA.IUIAutomationElement9
    IsDialog => (ComCall(117, this, "int*", &retVal := 0), retVal)
    CachedIsDialog => (ComCall(118, this, "int*", &retVal := 0), retVal)
}

class IUIAutomationElementArray extends UIA.IUIAutomationBase {
    __Item[index] {
        get => this.GetElement(index-1)
    }
    __Enum(varCount) {
        maxLen := this.Length-1, i := 0
        EnumElements(&element) {
            if i > maxLen
                return false
            element := this.GetElement(i++)
            return true
        }
        EnumIndexAndElements(&index, &element) {
            if i > maxLen
                return false
            element := this.GetElement(i++)
            index := i
            return true
        }
        return (varCount = 1) ? EnumElements : EnumIndexAndElements
    }
    ; Retrieves the number of elements in the collection.
    Length => (ComCall(3, this, "int*", &length := 0), length)

    ; Retrieves a Microsoft UI Automation element from the collection.
    GetElement(index) => (ComCall(4, this, "int", index, "ptr*", &element := 0), UIA.IUIAutomationElement(element))
}

/*
	Exposes properties and methods that UI Automation client applications use to view and navigate the UI Automation elements on the desktop.
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationtreewalker
*/
class IUIAutomationTreeWalker extends UIA.IUIAutomationBase {
    ; The structure of the Microsoft UI Automation tree changes as the visible UI elements on the desktop change.
    ; An element can have additional child elements that do not match the current view condition and thus are not returned when navigating the element tree.

    ; Retrieves the parent element of the specified UI Automation element.
    GetParentElement(element) => (ComCall(3, this, "ptr", element, "ptr*", &parent := 0), UIA.IUIAutomationElement(parent))

    ; Retrieves the first child element of the specified UI Automation element.
    GetFirstChildElement(element) => (ComCall(4, this, "ptr", element, "ptr*", &first := 0), UIA.IUIAutomationElement(first))

    ; Retrieves the last child element of the specified UI Automation element.
    GetLastChildElement(element) => (ComCall(5, this, "ptr", element, "ptr*", &last := 0), UIA.IUIAutomationElement(last))

    ; Retrieves the next sibling element of the specified UI Automation element, and caches properties and control patterns.
    GetNextSiblingElement(element) => (ComCall(6, this, "ptr", element, "ptr*", &next := 0), UIA.IUIAutomationElement(next))

    ; Retrieves the previous sibling element of the specified UI Automation element, and caches properties and control patterns.
    GetPreviousSiblingElement(element) => (ComCall(7, this, "ptr", element, "ptr*", &previous := 0), UIA.IUIAutomationElement(previous))

    ; Retrieves the ancestor element nearest to the specified Microsoft UI Automation element in the tree view.
    ; The element is normalized by navigating up the ancestor chain in the tree until an element that satisfies the view condition (specified by a previous call to IUIAutomationTreeWalker,,Condition) is reached. If the root element is reached, the root element is returned, even if it does not satisfy the view condition.
    ; This method is useful for applications that obtain references to UI Automation elements by hit-testing. The application might want to work only with specific types of elements, and can use IUIAutomationTreeWalker,,Normalize to make sure that no matter what element is initially retrieved (for example, when a scroll bar gets the input focus), only the element of interest (such as a content element) is ultimately retrieved.
    NormalizeElement(element) => (ComCall(8, this, "ptr", element, "ptr*", &normalized := 0), UIA.IUIAutomationElement(normalized))

    ; Retrieves the parent element of the specified UI Automation element, and caches properties and control patterns.
    GetParentElementBuildCache(element, cacheRequest) => (ComCall(9, this, "ptr", element, "ptr", cacheRequest, "ptr*", &parent := 0), UIA.IUIAutomationElement(parent))

    ; Retrieves the first child element of the specified UI Automation element, and caches properties and control patterns.
    GetFirstChildElementBuildCache(element, cacheRequest) => (ComCall(10, this, "ptr", element, "ptr", cacheRequest, "ptr*", &first := 0), UIA.IUIAutomationElement(first))

    ; Retrieves the last child element of the specified UI Automation element, and caches properties and control patterns.
    GetLastChildElementBuildCache(element, cacheRequest) => (ComCall(11, this, "ptr", element, "ptr", cacheRequest, "ptr*", &last := 0), UIA.IUIAutomationElement(last))

    ; Retrieves the next sibling element of the specified UI Automation element, and caches properties and control patterns.
    GetNextSiblingElementBuildCache(element, cacheRequest) => (ComCall(12, this, "ptr", element, "ptr", cacheRequest, "ptr*", &next := 0), UIA.IUIAutomationElement(next))

    ; Retrieves the previous sibling element of the specified UI Automation element, and caches properties and control patterns.
    GetPreviousSiblingElementBuildCache(element, cacheRequest) => (ComCall(13, this, "ptr", element, "ptr", cacheRequest, "ptr*", &previous := 0), UIA.IUIAutomationElement(previous))

    ; Retrieves the ancestor element nearest to the specified Microsoft UI Automation element in the tree view, prefetches the requested properties and control patterns, and stores the prefetched items in the cache.
    NormalizeElementBuildCache(element, cacheRequest) => (ComCall(14, this, "ptr", element, "ptr", cacheRequest, "ptr*", &normalized := 0), UIA.IUIAutomationElement(normalized))

    ; Retrieves the condition that defines the view of the UI Automation tree. This property is read-only.
    ; The condition that defines the view. This is the interface that was passed to CreateTreeWalker.
    Condition => (ComCall(15, this, "ptr*", &condition := 0), UIA.IUIAutomationCondition(condition))
}

/*
	Represents a condition based on a property value that is used to find UI Automation elements.
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationpropertycondition
*/
class IUIAutomationPropertyCondition extends UIA.IUIAutomationCondition {
    static __IID := "{99ebf2cb-5578-4267-9ad4-afd6ea77e94b}"
    PropertyId => (ComCall(3, this, "int*", &propertyId := 0), propertyId)
    PropertyValue => (ComCall(4, this, "ptr", propertyValue := UIA.ComVar()), propertyValue[])
    PropertyConditionFlags => (ComCall(5, this, "int*", &flags := 0), flags)
}

/*
	Exposes properties and methods that Microsoft UI Automation client applications can use to retrieve information about an AND-based property condition.
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationandcondition
*/
class IUIAutomationAndCondition extends UIA.IUIAutomationCondition {
    static __IID := "{a7d0af36-b912-45fe-9855-091ddc174aec}"
    ChildCount => (ComCall(3, this, "int*", &childCount := 0), childCount)
    GetChildrenAsNativeArray() => (ComCall(4, this, "ptr*", &childArray := 0, "int*", &childArrayCount := 0), UIA.NativeArray(childArray, childArrayCount))
    GetChildren() => (ComCall(5, this, "ptr*", &childArray := 0), UIA.IUIAutomationCondition.__SafeArrayToConditionArray(childArray))
}

/*
	Represents a condition made up of multiple conditions, at least one of which must be true.
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationorcondition
*/
class IUIAutomationOrCondition extends UIA.IUIAutomationAndCondition {
    static __IID := "{8753f032-3db1-47b5-a1fc-6e34a266c712}"
    ChildCount => (ComCall(3, this, "int*", &childCount := 0), childCount)
    GetChildrenAsNativeArray() => (ComCall(4, this, "ptr*", &childArray := 0, "int*", &childArrayCount := 0), UIA.NativeArray(childArray, childArrayCount))
    GetChildren() => (ComCall(5, this, "ptr*", &childArray := 0), UIA.IUIAutomationCondition.__SafeArrayToConditionArray(childArray))
}

/*
	Represents a condition that can be either TRUE=1 (selects all elements) or FALSE=0(selects no elements).
	Microsoft documentation: 
*/
class IUIAutomationBoolCondition extends UIA.IUIAutomationCondition {
    static __IID := "{1B4E1F2E-75EB-4D0B-8952-5A69988E2307}"
    Value => (ComCall(3, this, "int*", &boolVal := 0), boolVal)
}

/*
	Represents a condition that is the negative of another condition.
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationnotcondition
*/
class IUIAutomationNotCondition extends UIA.IUIAutomationCondition {
    static __IID := "{f528b657-847b-498c-8896-d52b565407a1}"
    GetChild() => (ComCall(3, this, "ptr*", &condition := 0), UIA.IUIAutomationCondition.__QueryCondition(condition))
}

class IUIAutomationCondition extends UIA.IUIAutomationBase {
    static __QueryCondition(pCond) {
        for n in ["Property", "Bool", "And", "Or", "Not"] {
            try {
                if ComObjQuery(pCond, UIA.IUIAutomation%n%Condition.__IID) 
                    return UIA.IUIAutomation%n%Condition(pCond)
            }
        }
    }
    static __SafeArrayToConditionArray(pSafeArr) {
        safeArray := ComValue(0x2003,pSafeArr), out := []
        for k in safeArray {
            if cond := UIA.IUIAutomationCondition.__QueryCondition(k)
                out.Push(cond)
        }
        return out
    }
}

/*
	Exposes properties and methods of a cache request. Client applications use this interface to specify the properties and control patterns to be cached when a Microsoft UI Automation element is obtained.
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationcacherequest
*/
class IUIAutomationCacheRequest extends UIA.IUIAutomationBase {
    ; Adds a property to the cache request.
    AddProperty(propertyId) => ComCall(3, this, "int", propertyId)

    ; Adds a control pattern to the cache request. Adding a control pattern that is already in the cache request has no effect.
    AddPattern(patternId) => ComCall(4, this, "int", patternId)

    ; Creates a copy of the cache request.
    Clone() => (ComCall(5, this, "ptr*", &clonedRequest := 0), UIA.IUIAutomationCacheRequest(clonedRequest))

    TreeScope {
        get => (ComCall(6, this, "int*", &scope := 0), scope)
        set => ComCall(7, this, "int", Value)
    }

    TreeFilter {
        get => (ComCall(8, this, "ptr*", &filter := 0), UIA.IUIAutomationCondition(filter))
        set => ComCall(9, this, "ptr", Value)
    }

    AutomationElementMode {
        get => (ComCall(10, this, "int*", &mode := 0), mode)
        set => ComCall(11, this, "int", Value)
    }
}

/*	event handle sample
* HandleAutomationEvent(pself,sender,eventId) ; UIA.IUIAutomationElement , EVENTID
* HandleFocusChangedEvent(pself,sender) ; UIA.IUIAutomationElement
* HandlePropertyChangedEvent(pself,sender,propertyId,newValue) ; UIA.IUIAutomationElement, PROPERTYID, VARIANT
* HandleStructureChangedEvent(pself,sender,changeType,runtimeId) ; UIA.IUIAutomationElement, StructureChangeType, SAFEARRAY
*/

class IUIAutomationEventHandler {
	static __IID := "{146c3c17-f12e-4e22-8c27-f894b9b79c69}"

	HandleAutomationEvent(pSelf, sender, eventId) {
		this.EventHandler.Call(UIA.IUIAutomationElement(sender), eventId)
	}
}
class IUIAutomationFocusChangedEventHandler {
	static __IID := "{c270f6b5-5c69-4290-9745-7a7f97169468}"

	HandleFocusChangedEvent(pSelf, sender) {
		this.EventHandler.Call(UIA.IUIAutomationElement(sender))
	}
}

/*
	Exposes a method to handle Microsoft UI Automation events that occur when a property is changed.
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationpropertychangedeventhandler
*/
class IUIAutomationPropertyChangedEventHandler { ; UNTESTED
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696119(v=vs.85).aspx
	static __IID := "{40cd37d4-c756-4b0c-8c6f-bddfeeb13b50}"

	HandlePropertyChangedEvent(pSelf, sender, propertyId, newValue) {
        val := ComValue(0x400C, newValue)[], DllCall("oleaut32\VariantClear", "ptr", newValue)
        this.EventHandler.Call(UIA.IUIAutomationElement(sender), propertyId, val)
    }
}
/*
	Exposes a method to handle events that occur when the Microsoft UI Automation tree structure is changed.
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationstructurechangedeventhandler
*/
class IUIAutomationStructureChangedEventHandler {
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ee696197(v=vs.85).aspx
	static __IID := "{e81d1b4e-11c5-42f8-9754-e7036c79f054}"

	HandleStructureChangedEvent(pSelf, sender, changeType, runtimeId) {
        this.EventHandler.Call(UIA.IUIAutomationElement(sender), changeType, ComValue(0x2003, runtimeId))
        DllCall("oleaut32\VariantClear", "ptr", runtimeId)
	}
}
/*
	Exposes a method to handle events that occur when Microsoft UI Automation reports a text-changed event from text edit controls
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationtextedittextchangedeventhandler
*/
class IUIAutomationTextEditTextChangedEventHandler { ; UNTESTED
	;~ http://msdn.microsoft.com/en-us/library/windows/desktop/dn302202(v=vs.85).aspx
	static __IID := "{92FAA680-E704-4156-931A-E32D5BB38F3F}"

	HandleTextEditTextChangedEvent(pSelf, sender, changeType, eventStrings) {
        val := ComValue(0x400C, eventStrings)[], DllCall("oleaut32\VariantClear", "ptr", eventStrings)
        this.EventHandler.Call(UIA.IUIAutomationElement(sender), changeType, val)
	}
}

/*
	Exposes a method to handle one or more Microsoft UI Automation change events
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationchangeseventhandler
*/
class IUIAutomationChangesEventHandler { ; UNTESTED
	static __IID := "{58EDCA55-2C3E-4980-B1B9-56C17F27A2A0}"

	HandleChangesEvent(pSelf, sender, uiaChanges, changesCount) {
        changes := {id:NumGet(uiaChanges,"Int"), payload:ComValue(0x400C, pPayload := NumGet(uiaChanges,8,"uint64"))[], extraInfo:ComValue(0x400C, pExtraInfo := NumGet(uiaChanges,16+2*A_PtrSize,"uint64")[])}
        DllCall("oleaut32\VariantClear", "ptr", pPayload), DllCall("oleaut32\VariantClear", "ptr", pExtraInfo)
        this.EventHandler.Call(UIA.IUIAutomationElement(sender), changes, changesCount)
	}
}
/*
	Exposes a method to handle Microsoft UI Automation notification events
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationnotificationeventhandler
*/
class IUIAutomationNotificationEventHandler {
	static __IID := "{C7CB2637-E6C2-4D0C-85DE-4948C02175C7}"

	HandleNotificationEvent(pSelf, sender, notificationKind, notificationProcessing, displayString, activityId) {
        this.EventHandler.Call(UIA.IUIAutomationElement(sender), notificationKind, notificationProcessing, UIA.BSTR(displayString), UIA.BSTR(activityId))
	}
}

class IUIAutomationEventHandlerGroup extends UIA.IUIAutomationBase {
	static __IID := "{C9EE12F2-C13B-4408-997C-639914377F4E}"

	AddActiveTextPositionChangedEventHandler(handler, scope:=0x4, cacheRequest:=0) => ComCall(3, this, "int", scope, "ptr", cacheRequest, "ptr", handler)
    AddAutomationEventHandler(eventId, handler, scope:=0x4, cacheRequest:=0) => ComCall(4, this, "uint", eventId, "int", scope, "ptr", cacheRequest, "ptr", handler)
    AddChangesEventHandler(changeTypes, handler, scope:=0x4, cacheRequest:=0) {
        if !IsObject(changeTypes)
            changeTypes := [changeTypes]
        nativeArray := UIA.NativeArray(0, changeTypes.Length, "int")
        for k, v in changeTypes
            NumPut("int", v, nativeArray, (k-1)*4)
        return ComCall(5, this, "int", scope, "ptr", nativeArray, "int", changeTypes.Length, "int", cacheRequest, "ptr", handler)
    }
	AddNotificationEventHandler(handler, scope:=0x4, cacheRequest:=0) => ComCall(6, this, "int", scope, "ptr", cacheRequest, "ptr", handler)
	AddPropertyChangedEventHandler(propertyArray, handler, scope:=0x1,cacheRequest:=0) {
        if !IsObject(propertyArray)
            propertyArray := [propertyArray]
		SafeArray:=ComObjArray(0x3,propertyArray.Length)
		for i,propertyId in propertyArray
			SafeArray[i-1]:=propertyId
		return ComCall(7, this, "int",scope, "ptr", cacheRequest,"ptr", handler,"ptr", SafeArray)
	}
	AddStructureChangedEventHandler(handler, scope:=0x4, cacheRequest:=0) => ComCall(8, this "int", scope, "ptr",cacheRequest, "ptr", handler)
	AddTextEditTextChangedEventHandler(textEditChangeType, handler:="", scope:=0x4, cacheRequest:=0) => ComCall(9, this, "int", scope, "int", textEditChangeType, "ptr", cacheRequest, "ptr", handler)
}

static CreateEventHandler(funcObj, handlerType:="") {
    if funcObj is String
        try funcObj := %funcObj%
    if !HasMethod(funcObj, "Call")
        throw TypeError("Invalid function provided", -2)

    buf := Buffer(A_PtrSize * 5)
    handler := UIA.IUIAutomation%handlerType%EventHandler()
    handler.Buffer := buf, handler.Ptr := buf.ptr
    handlerFunc := handler.Handle%(handlerType ? handlerType : "Automation")%Event
    NumPut("ptr", buf.Ptr + A_PtrSize, "ptr", cQI:=CallbackCreate(QueryInterface, "F", 3), "ptr", cAF:=CallbackCreate(AddRef, "F", 1), "ptr", cR:=CallbackCreate(Release, "F", 1), "ptr", cF:=CallbackCreate(handlerFunc.Bind(handler), "F", handlerFunc.MaxParams-1), buf)
    handler.DefineProp("__Delete", { call: (*) => (CallbackFree(cQI), CallbackFree(cAF), CallbackFree(cR), CallbackFree(cF)) })
    handler.EventHandler := funcObj
    
    return handler

    QueryInterface(pSelf, pRIID, pObj){ ; Credit: https://github.com/neptercn/UIAutomation/blob/master/UIA.ahk
		DllCall("ole32\StringFromIID","ptr",pRIID,"wstr*",&str)
		return (str="{00000000-0000-0000-C000-000000000046}")||(str="{146c3c17-f12e-4e22-8c27-f894b9b79c69}")||(str="{40cd37d4-c756-4b0c-8c6f-bddfeeb13b50}")||(str="{e81d1b4e-11c5-42f8-9754-e7036c79f054}")||(str="{c270f6b5-5c69-4290-9745-7a7f97169468}")||(str="{92FAA680-E704-4156-931A-E32D5BB38F3F}")||(str="{58EDCA55-2C3E-4980-B1B9-56C17F27A2A0}")||(str="{C7CB2637-E6C2-4D0C-85DE-4948C02175C7}")?NumPut("ptr",pSelf,pObj)*0:0x80004002 ; E_NOINTERFACE
	}
    AddRef(pSelf) {
    }
    Release(pSelf) {
    }
}

class IUIAutomationAnnotationPattern extends UIA.IUIAutomationBase {
    AnnotationTypeId => (ComCall(3, this, "int*", &retVal := 0), retVal)
    AnnotationTypeName => (ComCall(4, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    Author => (ComCall(5, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    DateTime => (ComCall(6, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    Target => (ComCall(7, this, "ptr*", &retVal := 0), UIA.IUIAutomationElement(retVal))
    CachedAnnotationTypeId => (ComCall(8, this, "int*", &retVal := 0), retVal)
    CachedAnnotationTypeName => (ComCall(9, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    CachedAuthor => (ComCall(10, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    CachedDateTime => (ComCall(11, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    CachedTarget => (ComCall(11, this, "ptr*", &retVal := 0), UIA.IUIAutomationElement(retVal))
}



class IUIAutomationCustomNavigationPattern extends UIA.IUIAutomationBase {
    Navigate(direction) => (ComCall(3, this, "int", direction, "ptr*", &pRetVal := 0), UIA.IUIAutomationElement(pRetVal))
}

/*
	Provides access to a control that enables child elements to be arranged horizontally and vertically, relative to each other.
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationdockpattern
*/
class IUIAutomationDockPattern extends UIA.IUIAutomationBase {
	static	__IID := "{fde5ef97-1464-48f6-90bf-43d0948e86ec}"

    ; ---------- DockPattern properties ----------

    ; Retrieves the `dock position` of this element within its docking container.
    DockPosition => (ComCall(4, this, "int*", &retVal := 0), retVal)
    ; Retrieves the `cached dock` position of this element within its docking container.
    CachedDockPosition => (ComCall(5, this, "int*", &retVal := 0), retVal)

    ; ---------- DockPattern methods ----------

    /**
     * Sets the dock position of this element. 
     * @param dockPos One of UIA.DockPosition values: Top = 0, Left = 1, Bottom = 2, Right = 3, Fill = 4, None = 5
     */
    SetDockPosition(dockPos) => ComCall(3, this, "int", dockPos)
}

/*
	Provides access to information exposed by a UI Automation provider for an element that can be dragged as part of a drag-and-drop operation.
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationdragpattern
*/
class IUIAutomationDragPattern extends UIA.IUIAutomationBase {
    static __IID := "{1DC7B570-1F54-4BAD-BCDA-D36A722FB7BD}"

    ; ---------- DragPattern properties ----------

    IsGrabbed => (ComCall(3, this, "int*", &retVal := 0), retVal)
    CachedIsGrabbed => (ComCall(4, this, "int*", &retVal := 0), retVal)
    DropEffect => (ComCall(5, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    CachedDropEffect => (ComCall(6, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    DropEffects => (ComCall(7, this, "ptr*", &retVal := 0), ComValue(0x2008, retVal))
    CachedDropEffects => (ComCall(8, this, "ptr*", &retVal := 0), ComValue(0x2008, retVal))

    ; ---------- DragPattern methods ----------
    /**
     * Retrieves a collection of elements that represent the full set of items that the user is dragging as part of a drag operation.
     * @returns {IUIAutomationElementArray}
     */
    GetGrabbedItems() => (ComCall(9, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))
    GetCachedGrabbedItems() => (ComCall(10, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))
}

/*
	Provides access to drag-and-drop information exposed by a Microsoft UI Automation provider for an element that can be the drop target of a drag-and-drop operation.
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationdroptargetpattern
*/
class IUIAutomationDropTargetPattern extends UIA.IUIAutomationBase {
    static __IID := "{69A095F7-EEE4-430E-A46B-FB73B1AE39A5}"

    ; ---------- DropTargetPattern properties ----------

    DropTargetEffect => (ComCall(3, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    CachedDropTargetEffect => (ComCall(4, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    DropTargetEffects => (ComCall(5, this, "ptr*", &retVal := 0), ComValue(0x2008, retVal))
    CachedDropTargetEffects => (ComCall(6, this, "ptr*", &retVal := 0), ComValue(0x2008, retVal))
}

class IUIAutomationExpandCollapsePattern extends UIA.IUIAutomationBase {
    ; This is a blocking method that returns after the element has been collapsed.
    ; There are cases when a element that is marked as a leaf node might not know whether it has children until either the IUIAutomationExpandCollapsePattern,,Collapse or the IUIAutomationExpandCollapsePattern,,Expand method is called. This behavior is possible with a tree view control that does delayed loading of its child items. For example, Microsoft Windows Explorer might display the expand icon for a node even though there are currently no child items; when the icon is clicked, the control polls for child items, finds none, and removes the expand icon. In these cases clients should listen for a property-changed event on the IUIAutomationExpandCollapsePattern,,ExpandCollapseState property.

    ; Displays all child nodes, controls, or content of the element.
    Expand() => ComCall(3, this)

    ; Hides all child nodes, controls, or content of the element.
    Collapse() => ComCall(4, this)

    ; Retrieves a value that indicates the state, expanded or collapsed, of the element.
    ExpandCollapseState => (ComCall(5, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates the state, expanded or collapsed, of the element.
    CachedExpandCollapseState => (ComCall(6, this, "int*", &retVal := 0), retVal)
}

class IUIAutomationGridItemPattern extends UIA.IUIAutomationBase {
    ; Retrieves the element that contains the grid item.
    ContainingGrid => (ComCall(3, this, "ptr*", &retVal := 0), UIA.IUIAutomationElement(retVal))

    ; Retrieves the zero-based index of the row that contains the grid item.
    Row => (ComCall(4, this, "int*", &retVal := 0), retVal)

    ; Retrieves the zero-based index of the column that contains the item.
    Column => (ComCall(5, this, "int*", &retVal := 0), retVal)

    ; Retrieves the number of rows spanned by the grid item.
    RowSpan => (ComCall(6, this, "int*", &retVal := 0), retVal)

    ; Retrieves the number of columns spanned by the grid item.
    ColumnSpan => (ComCall(7, this, "int*", &retVal := 0), retVal)

    ; Retrieves the cached element that contains the grid item.
    CachedContainingGrid => (ComCall(8, this, "ptr*", &retVal := 0), UIA.IUIAutomationElement(retVal))

    ; Retrieves the cached zero-based index of the row that contains the item.
    CachedRow => (ComCall(9, this, "int*", &retVal := 0), retVal)

    ; Retrieves the cached zero-based index of the column that contains the grid item.
    CachedColumn => (ComCall(10, this, "int*", &retVal := 0), retVal)

    ; Retrieves the cached number of rows spanned by a grid item.
    CachedRowSpan => (ComCall(11, this, "int*", &retVal := 0), retVal)

    ; Retrieves the cached number of columns spanned by the grid item.
    CachedColumnSpan => (ComCall(12, this, "int*", &retVal := 0), retVal)
}

class IUIAutomationGridPattern extends UIA.IUIAutomationBase {
    ; Retrieves a UI Automation element representing an item in the grid.
    GetItem(row, column) => (ComCall(3, this, "int", row, "int", column, "ptr*", &element := 0), UIA.IUIAutomationElement(element))

    ; Hidden rows and columns, depending on the provider implementation, may be loaded in the Microsoft UI Automation tree and will therefore be reflected in the row count and column count properties. If the hidden rows and columns have not yet been loaded they are not counted.

    ; Retrieves the number of rows in the grid.
    RowCount => (ComCall(4, this, "int*", &retVal := 0), retVal)

    ; The number of columns in the grid.
    ColumnCount => (ComCall(5, this, "int*", &retVal := 0), retVal)

    ; Retrieves the cached number of rows in the grid.
    CachedRowCount => (ComCall(6, this, "int*", &retVal := 0), retVal)

    ; Retrieves the cached number of columns in the grid.
    CachedColumnCount => (ComCall(7, this, "int*", &retVal := 0), retVal)
}

class IUIAutomationInvokePattern extends UIA.IUIAutomationBase {
    ; Invokes the action of a control, such as a button click.
    ; Calls to this method should return immediately without blocking. However, this behavior depends on the implementation.
    Invoke() => ComCall(3, this)
}

class IUIAutomationItemContainerPattern extends UIA.IUIAutomationBase {
    ; IUIAutomationItemContainerPattern

    ; Retrieves an element within a containing element, based on a specified property value.
    ; The provider may return an actual UIA.IUIAutomationElement interface or a placeholder if the matching element is virtualized.
    ; This method returns E_INVALIDARG if the property requested is not one that the container supports searching over. It is expected that most containers will support Name property, and if appropriate for the container, AutomationId and IsSelected.
    ; This method can be slow, because it may need to traverse multiple objects to find a matching one. When used in a loop to return multiple items, no specific order is defined so long as each item is returned only once (that is, the loop should terminate). This method is also item-centric, not UI-centric, so items with multiple UI representations need to be hit only once.
    ; When the propertyId parameter is specified as 0 (zero), the provider is expected to return the next item after pStartAfter. If pStartAfter is specified as NULL with a propertyId of 0, the provider should return the first item in the container. When propertyId is specified as 0, the value parameter should be VT_EMPTY.
    FindItemByProperty(pStartAfter, propertyId, value) {
        if A_PtrSize = 4
            value := UIA.ComVar(value, , true), ComCall(3, this, "ptr", pStartAfter, "int", propertyId, "int64", NumGet(value, "int64"), "int64", NumGet(value, 8, "int64"), "ptr*", &pFound := 0)
        else
            ComCall(3, this, "ptr", pStartAfter, "int", propertyId, "ptr", UIA.ComVar(value, , true), "ptr*", &pFound := 0)
        if (pFound)
            return UIA.IUIAutomationElement(pFound)
        throw TargetError("Target elements not found.")
    }
}

class IUIAutomationLegacyIAccessiblePattern extends UIA.IUIAutomationBase {

    ; IUIAutomationLegacyIAccessiblePattern

    ; Performs a Microsoft Active Accessibility selection.
    Select(flagsSelect:=3) => ComCall(3, this, "int", flagsSelect)

    ; Performs the Microsoft Active Accessibility default action for the element.
    DoDefaultAction() => ComCall(4, this)

    ; Sets the Microsoft Active Accessibility value property for the element. This method is supported only for some elements (usually edit controls).
    SetValue(szValue) => ComCall(5, this, "wstr", szValue)

    ; Retrieves the Microsoft Active Accessibility child identifier for the element. If the element is not a child element, CHILDID_SELF (0) is returned.
    ChildId => (ComCall(6, this, "int*", &pRetVal := 0), pRetVal)

    ; Retrieves the Microsoft Active Accessibility name property of the element. The name of an element can be used to find the element in the element tree when the automation ID property is not supported on the element.
    Name => (ComCall(7, this, "ptr*", &pszName := 0), UIA.BSTR(pszName))

    ; Retrieves the Microsoft Active Accessibility value property.
    Value => (ComCall(8, this, "ptr*", &pszValue := 0), UIA.BSTR(pszValue))

    ; Retrieves the Microsoft Active Accessibility description of the element.
    Description => (ComCall(9, this, "ptr*", &pszDescription := 0), UIA.BSTR(pszDescription))

    ; Retrieves the Microsoft Active Accessibility role identifier of the element.
    Role => (ComCall(10, this, "uint*", &pdwRole := 0), pdwRole)

    ; Retrieves the Microsoft Active Accessibility state identifier for the element.
    State => (ComCall(11, this, "uint*", &pdwState := 0), pdwState)

    ; Retrieves the Microsoft Active Accessibility help string for the element.
    Help => (ComCall(12, this, "ptr*", &pszHelp := 0), UIA.BSTR(pszHelp))

    ; Retrieves the Microsoft Active Accessibility keyboard shortcut property for the element.
    KeyboardShortcut => (ComCall(13, this, "ptr*", &pszKeyboardShortcut := 0), UIA.BSTR(pszKeyboardShortcut))

    ; Retrieves the Microsoft Active Accessibility property that identifies the selected children of this element.
    GetSelection() => (ComCall(14, this, "ptr*", &pvarSelectedChildren := 0), UIA.IUIAutomationElementArray(pvarSelectedChildren))

    ; Retrieves the Microsoft Active Accessibility default action for the element.
    DefaultAction => (ComCall(15, this, "ptr*", &pszDefaultAction := 0), UIA.BSTR(pszDefaultAction))

    ; Retrieves the cached Microsoft Active Accessibility child identifier for the element.
    CachedChildId => (ComCall(16, this, "int*", &pRetVal := 0), pRetVal)

    ; Retrieves the cached Microsoft Active Accessibility name property of the element.
    CachedName => (ComCall(17, this, "ptr*", &pszName := 0), UIA.BSTR(pszName))

    ; Retrieves the cached Microsoft Active Accessibility value property.
    CachedValue => (ComCall(18, this, "ptr*", &pszValue := 0), UIA.BSTR(pszValue))

    ; Retrieves the cached Microsoft Active Accessibility description of the element.
    CachedDescription => (ComCall(19, this, "ptr*", &pszDescription := 0), UIA.BSTR(pszDescription))

    ; Retrieves the cached Microsoft Active Accessibility role of the element.
    CachedRole => (ComCall(20, this, "uint*", &pdwRole := 0), pdwRole)

    ; Retrieves the cached Microsoft Active Accessibility state identifier for the element.
    CachedState => (ComCall(21, this, "uint*", &pdwState := 0), pdwState)

    ; Retrieves the cached Microsoft Active Accessibility help string for the element.
    CachedHelp => (ComCall(22, this, "ptr*", &pszHelp := 0), UIA.BSTR(pszHelp))

    ; Retrieves the cached Microsoft Active Accessibility keyboard shortcut property for the element.
    CachedKeyboardShortcut => (ComCall(23, this, "ptr*", &pszKeyboardShortcut := 0), UIA.BSTR(pszKeyboardShortcut))

    ; Retrieves the cached Microsoft Active Accessibility property that identifies the selected children of this element.
    GetCachedSelection() => (ComCall(24, this, "ptr*", &pvarSelectedChildren := 0), UIA.IUIAutomationElementArray(pvarSelectedChildren))

    ; Retrieves the Microsoft Active Accessibility default action for the element.
    CachedDefaultAction => (ComCall(25, this, "ptr*", &pszDefaultAction := 0), UIA.BSTR(pszDefaultAction))

    ; Retrieves an IAccessible object that corresponds to the Microsoft UI Automation element.
    ; This method returns NULL if the underlying implementation of the UI Automation element is not a native Microsoft Active Accessibility server; that is, if a client attempts to retrieve the IAccessible interface for an element originally supported by a proxy object from OLEACC.dll, or by the UIA-to-MSAA Bridge.
    GetIAccessible() => (ComCall(26, this, "ptr*", &ppAccessible := 0), ComValue(0xd, ppAccessible))
}

class IUIAutomationMultipleViewPattern extends UIA.IUIAutomationBase {
    static	__IID := "{8d253c91-1dc5-4bb5-b18f-ade16fa495e8}"

    ; ---------- MultipleViewPattern properties ----------

    ; Retrieves the name of a control-specific view.
    GetViewName(view) => (ComCall(3, this, "int", view, "ptr*", &name := 0), UIA.BSTR(name))

    ; Sets the view of the control.
    SetView(view) => ComCall(4, this, "int", view)

    ; Retrieves the control-specific identifier of the current view of the control.
    View => this.CurrentView
    CurrentView => (ComCall(5, this, "int*", &retVal := 0), retVal)

    ; Retrieves a collection of control-specific view identifiers.
    GetSupportedViews() => (ComCall(6, this, "ptr*", &retVal := 0), ComValue(0x2003, retVal))

    ; Retrieves the cached control-specific identifier of the current view of the control.
    CachedView => this.CachedCurrentView
    CachedCurrentView => (ComCall(7, this, "int*", &retVal := 0), retVal)

    ; Retrieves a collection of control-specific view identifiers from the cache.
    GetCachedSupportedViews() => (ComCall(8, this, "ptr*", &retVal := 0), ComValue(0x2003, retVal))
}

/*
	Provides access to the underlying object model implemented by a control or application.
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationobjectmodelpattern
*/
class IUIAutomationObjectModelPattern extends UIA.IUIAutomationBase {
    static	__IID := "{71c284b3-c14d-4d14-981e-19751b0d756d}"
    ; Retrieves an interface used to access the underlying object model of the provider.
    GetUnderlyingObjectModel() => (ComCall(3, this, "ptr*", &retVal := 0), ComValue(0xd, retVal))
}

class IUIAutomationProxyFactory extends UIA.IUIAutomationBase {
    CreateProvider(hwnd, idObject, idChild) => (ComCall(3, this, "ptr", hwnd, "int", idObject, "int", idChild, "ptr*", &provider := 0), ComValue(0xd, provider))
    ProxyFactoryId => (ComCall(4, this, "ptr*", &factoryId := 0), UIA.BSTR(factoryId))
}

class IUIAutomationProxyFactoryEntry extends UIA.IUIAutomationBase {
    ProxyFactory() => (ComCall(3, this, "ptr*", &factory := 0), UIA.IUIAutomationProxyFactory(factory))
    ClassName {
        get => (ComCall(4, this, "ptr*", &classname := 0), UIA.BSTR(classname))
        set => (ComCall(9, this, "wstr", Value))
    }
    ImageName {
        get => (ComCall(5, this, "ptr*", &imageName := 0), UIA.BSTR(imageName))
        set => (ComCall(10, this, "wstr", Value))
    }
    AllowSubstringMatch {
        get => (ComCall(6, this, "int*", &allowSubstringMatch := 0), allowSubstringMatch)
        set => (ComCall(11, this, "int", Value))
    }
    CanCheckBaseClass {
        get => (ComCall(7, this, "int*", &canCheckBaseClass := 0), canCheckBaseClass)
        set => (ComCall(12, this, "int", Value))
    }
    NeedsAdviseEvents {
        get => (ComCall(8, this, "int*", &adviseEvents := 0), adviseEvents)
        set => (ComCall(13, this, "int", Value))
    }
    SetWinEventsForAutomationEvent(eventId, propertyId, winEvents) => ComCall(14, this, "int", eventId, "Int", propertyId, "ptr", winEvents)
    GetWinEventsForAutomationEvent(eventId, propertyId) => (ComCall(15, this, "int", eventId, "Int", propertyId, "ptr*", &winEvents := 0), ComValue(0x200d, winEvents))
}

class IUIAutomationProxyFactoryMapping extends UIA.IUIAutomationBase {
    Count => (ComCall(3, this, "uint*", &count := 0), count)
    GetTable() => (ComCall(4, this, "ptr*", &table := 0), ComValue(0x200d, table))
    GetEntry(index) => (ComCall(5, this, "int", index, "ptr*", &entry := 0), UIA.IUIAutomationProxyFactoryEntry(entry))
    SetTable(factoryList) => ComCall(6, this, "ptr", factoryList)
    InsertEntries(before, factoryList) => ComCall(7, this, "uint", before, "ptr", factoryList)
    InsertEntry(before, factory) => ComCall(8, this, "uint", before, "ptr", factory)
    RemoveEntry(index) => ComCall(9, this, "uint", index)
    ClearTable() => ComCall(10, this)
    RestoreDefaultTable() => ComCall(11, this)
}

class IUIAutomationRangeValuePattern extends UIA.IUIAutomationBase {
    ; Sets the value of the control.
    SetValue(val) => ComCall(3, this, "double", val)

    ; Retrieves the value of the control.
    Value => (ComCall(4, this, "double*", &retVal := 0), retVal)

    ; Indicates whether the value of the element can be changed.
    IsReadOnly => (ComCall(5, this, "int*", &retVal := 0), retVal)

    ; Retrieves the maximum value of the control.
    Maximum => (ComCall(6, this, "double*", &retVal := 0), retVal)

    ; Retrieves the minimum value of the control.
    Minimum => (ComCall(7, this, "double*", &retVal := 0), retVal)

    ; The LargeChange and SmallChange property can support a Not a Number (NaN) value. When retrieving this property, a client can use the _isnan function to determine whether the property is a NaN value.

    ; Retrieves the value that is added to or subtracted from the value of the control when a large change is made, such as when the PAGE DOWN key is pressed.
    LargeChange => (ComCall(8, this, "double*", &retVal := 0), retVal)

    ; Retrieves the value that is added to or subtracted from the value of the control when a small change is made, such as when an arrow key is pressed.
    SmallChange => (ComCall(9, this, "double*", &retVal := 0), retVal)

    ; Retrieves the cached value of the control.
    CachedValue => (ComCall(10, this, "double*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates whether the value of the element can be changed.
    CachedIsReadOnly => (ComCall(11, this, "int*", &retVal := 0), retVal)

    ; Retrieves the cached maximum value of the control.
    CachedMaximum => (ComCall(12, this, "double*", &retVal := 0), retVal)

    ; Retrieves the cached minimum value of the control.
    CachedMinimum => (ComCall(13, this, "double*", &retVal := 0), retVal)

    ; Retrieves, from the cache, the value that is added to or subtracted from the value of the control when a large change is made, such as when the PAGE DOWN key is pressed.
    CachedLargeChange => (ComCall(14, this, "double*", &retVal := 0), retVal)

    ; Retrieves, from the cache, the value that is added to or subtracted from the value of the control when a small change is made, such as when an arrow key is pressed.
    CachedSmallChange => (ComCall(15, this, "double*", &retVal := 0), retVal)
}

class IUIAutomationScrollItemPattern extends UIA.IUIAutomationBase {
    ; Scrolls the content area of a container object to display the UI Automation element within the visible region (viewport) of the container.
    ; This method does not provide the ability to specify the position of the element within the viewport.
    ScrollIntoView() => ComCall(3, this)
}

class IUIAutomationScrollPattern extends UIA.IUIAutomationBase {
    ; Scrolls the visible region of the content area horizontally and vertically.
    ; Default values for horizontalAmount and horizontalAmount is UIA.ScrollAmount.NoAmount
    Scroll(horizontalAmount:=-1, verticalAmount:=-1) => ComCall(3, this, "int", horizontalAmount, "int", verticalAmount)

    ; Sets the horizontal and vertical scroll positions as a percentage of the total content area within the UI Automation element.
    ; This method is useful only when the content area of the control is larger than the visible region.
    ; Default values for horizontalPercent and verticalPercent is UIA.ScrollAmount.NoAmount
    SetScrollPercent(horizontalPercent:=-1, verticalPercent:=-1) => ComCall(4, this, "double", horizontalPercent, "double", verticalPercent)

    ; Retrieves the horizontal scroll position.
    HorizontalScrollPercent => (ComCall(5, this, "double*", &retVal := 0), retVal)

    ; Retrieves the vertical scroll position.
    VerticalScrollPercent => (ComCall(6, this, "double*", &retVal := 0), retVal)

    ; Retrieves the horizontal size of the viewable region of a scrollable element.
    HorizontalViewSize => (ComCall(7, this, "double*", &retVal := 0), retVal)

    ; Retrieves the vertical size of the viewable region of a scrollable element.
    VerticalViewSize => (ComCall(8, this, "double*", &retVal := 0), retVal)

    ; Indicates whether the element can scroll horizontally.
    ; This property can be dynamic. For example, the content area of the element might not be larger than the current viewable area, meaning that the property is FALSE. However, resizing the element or adding child items can increase the bounds of the content area beyond the viewable area, making the property TRUE.
    HorizontallyScrollable => (ComCall(9, this, "int*", &retVal := 0), retVal)

    ; Indicates whether the element can scroll vertically.
    VerticallyScrollable => (ComCall(10, this, "int*", &retVal := 0), retVal)

    ; Retrieves the cached horizontal scroll position.
    CachedHorizontalScrollPercent => (ComCall(11, this, "double*", &retVal := 0), retVal)

    ; Retrieves the cached vertical scroll position.
    CachedVerticalScrollPercent => (ComCall(12, this, "double*", &retVal := 0), retVal)

    ; Retrieves the cached horizontal size of the viewable region of a scrollable element.
    CachedHorizontalViewSize => (ComCall(13, this, "double*", &retVal := 0), retVal)

    ; Retrieves the cached vertical size of the viewable region of a scrollable element.
    CachedVerticalViewSize => (ComCall(14, this, "double*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates whether the element can scroll horizontally.
    CachedHorizontallyScrollable => (ComCall(15, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates whether the element can scroll vertically.
    CachedVerticallyScrollable => (ComCall(16, this, "int*", &retVal := 0), retVal)
}

class IUIAutomationSelectionItemPattern extends UIA.IUIAutomationBase {
    ; Clears any selected items and then selects the current element.
    Select() => ComCall(3, this)

    ; Adds the current element to the collection of selected items.
    AddToSelection() => ComCall(4, this)

    ; Removes this element from the selection.
    ; An error code is returned if this element is the only one in the selection and the selection container requires at least one element to be selected.
    RemoveFromSelection() => ComCall(5, this)

    ; Indicates whether this item is selected.
    IsSelected => (ComCall(6, this, "int*", &retVal := 0), retVal)

    ; Retrieves the element that supports IUIAutomationSelectionPattern and acts as the container for this item.
    SelectionContainer => (ComCall(7, this, "ptr*", &retVal := 0), UIA.IUIAutomationElement(retVal))

    ; A cached value that indicates whether this item is selected.
    CachedIsSelected => (ComCall(8, this, "int*", &retVal := 0), retVal)

    ; Retrieves the cached element that supports IUIAutomationSelectionPattern and acts as the container for this item.
    CachedSelectionContainer => (ComCall(9, this, "ptr*", &retVal := 0), UIA.IUIAutomationElement(retVal))
}

class IUIAutomationSelectionPattern extends UIA.IUIAutomationBase {
    ; Retrieves the selected elements in the container.
    GetSelection() => (ComCall(3, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))

    ; Indicates whether more than one item in the container can be selected at one time.
    CanSelectMultiple => (ComCall(4, this, "int*", &retVal := 0), retVal)

    ; Indicates whether at least one item must be selected at all times.
    IsSelectionRequired => (ComCall(5, this, "int*", &retVal := 0), retVal)

    ; Retrieves the cached selected elements in the container.
    GetCachedSelection() => (ComCall(6, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))

    ; Retrieves a cached value that indicates whether more than one item in the container can be selected at one time.
    CachedCanSelectMultiple => (ComCall(7, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates whether at least one item must be selected at all times.
    CachedIsSelectionRequired => (ComCall(8, this, "int*", &retVal := 0), retVal)
}

class IUIAutomationSpreadsheetPattern extends UIA.IUIAutomationBase {
    GetItemByName(name) => (ComCall(3, this, "wstr", name, "ptr*", &element := 0), UIA.IUIAutomationElement(element))
}

class IUIAutomationSpreadsheetItemPattern extends UIA.IUIAutomationBase {
    Formula => (ComCall(3, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    GetAnnotationObjects() => (ComCall(4, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))
    GetAnnotationTypes() => (ComCall(5, this, "ptr*", &retVal := 0), ComValue(0x2003, retVal))
    CachedFormul => (ComCall(6, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    GetCachedAnnotationObjects() => (ComCall(7, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))
    GetCachedAnnotationTypes() => (ComCall(8, this, "ptr*", &retVal := 0), ComValue(0x2003, retVal))
}

class IUIAutomationStylesPattern extends UIA.IUIAutomationBase {
    StyleId => (ComCall(3, this, "int*", &retVal := 0), retVal)
    StyleName => (ComCall(4, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    FillColor => (ComCall(5, this, "int*", &retVal := 0), retVal)
    FillPatternStyle => (ComCall(6, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    Shape => (ComCall(7, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    FillPatternColor => (ComCall(8, this, "int*", &retVal := 0), retVal)
    ExtendedProperties => (ComCall(9, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    GetExtendedPropertiesAsArray() {
        ComCall(10, this, "ptr*", &propertyArray := 0, "int*", &propertyCount := 0), arr := []
        for p in UIA.NativeArray(propertyArray, propertyCount)
            arr.Push({ PropertyName: UIA.BSTR(NumGet(p, "ptr")), PropertyValue: UIA.BSTR(NumGet(p, A_PtrSize, "ptr")) })
        return arr
    }
    CachedStyleId => (ComCall(11, this, "int*", &retVal := 0), retVal)
    CachedStyleName => (ComCall(12, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    CachedFillColor => (ComCall(13, this, "int*", &retVal := 0), retVal)
    CachedFillPatternStyle => (ComCall(14, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    CachedShape => (ComCall(15, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    CachedFillPatternColor => (ComCall(16, this, "int*", &retVal := 0), retVal)
    CachedExtendedProperties => (ComCall(17, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))
    GetCachedExtendedPropertiesAsArray() {
        ComCall(18, this, "ptr*", &propertyArray := 0, "int*", &propertyCount := 0), arr := []
        for p in UIA.NativeArray(propertyArray, propertyCount)
            arr.Push({ PropertyName: UIA.BSTR(NumGet(p, "ptr")), PropertyValue: UIA.BSTR(NumGet(p, A_PtrSize, "ptr")) })
        return arr
    }
}

class IUIAutomationSynchronizedInputPattern extends UIA.IUIAutomationBase {
    ; Causes the Microsoft UI Automation provider to start listening for mouse or keyboard input.
    ; When matching input is found, the provider checks whether the target element matches the current element. If they match, the provider raises the UIA_InputReachedTargetEventId event; otherwise it raises the UIA_InputReachedOtherElementEventId or UIA_InputDiscardedEventId event.
    ; After receiving input of the specified type, the provider stops checking for input and continues as normal.
    ; If the provider is already listening for input, this method returns E_INVALIDOPERATION.
    StartListening(inputType) => ComCall(3, this, "int", inputType)

    ; Causes the Microsoft UI Automation provider to stop listening for mouse or keyboard input.
    Cancel() => ComCall(4, this)
}

class IUIAutomationTableItemPattern extends UIA.IUIAutomationBase {
    ; Retrieves the row headers associated with a table item or cell.
    GetRowHeaderItems() => (ComCall(3, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))

    ; Retrieves the column headers associated with a table item or cell.
    GetColumnHeaderItems() => (ComCall(4, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))

    ; Retrieves the cached row headers associated with a table item or cell.
    GetCachedRowHeaderItems() => (ComCall(5, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))

    ; Retrieves the cached column headers associated with a table item or cell.
    GetCachedColumnHeaderItems() => (ComCall(6, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))
}

class IUIAutomationTablePattern extends UIA.IUIAutomationBase {
    ; Retrieves a collection of UI Automation elements representing all the row headers in a table.
    GetRowHeaders() => (ComCall(3, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))

    ; Retrieves a collection of UI Automation elements representing all the column headers in a table.
    GetColumnHeaders() => (ComCall(4, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))

    ; Retrieves the primary direction of traversal for the table.
    RowOrColumnMajor => (ComCall(5, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached collection of UI Automation elements representing all the row headers in a table.
    GetCachedRowHeaders() => (ComCall(6, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))

    ; Retrieves a cached collection of UI Automation elements representing all the column headers in a table.
    GetCachedColumnHeaders() => (ComCall(7, this, "ptr*", &retVal := 0), UIA.IUIAutomationElementArray(retVal))

    ; Retrieves the cached primary direction of traversal for the table.
    CachedRowOrColumnMajor => (ComCall(8, this, "int*", &retVal := 0), retVal)
}

class IUIAutomationTextChildPattern extends UIA.IUIAutomationBase {
    TextContainer => (ComCall(3, this, "ptr*", &container := 0), UIA.IUIAutomationElement(container))
    TextRange => (ComCall(4, this, "ptr*", &range := 0), UIA.IUIAutomationTextRange(range))
}

class IUIAutomationTextEditPattern extends UIA.IUIAutomationBase {
    GetActiveComposition() => (ComCall(3, this, "ptr*", &range := 0), UIA.IUIAutomationTextRange(range))
    GetConversionTarget() => (ComCall(4, this, "ptr*", &range := 0), UIA.IUIAutomationTextRange(range))
}

class IUIAutomationTextPattern extends UIA.IUIAutomationBase {
    ; Retrieves the degenerate (empty) text range nearest to the specified screen coordinates.
    /*
    * A text range that wraps a child object is returned if the screen coordinates are within the coordinates of an image, hyperlink, Microsoft Excel spreadsheet, or other embedded object.
    * Because hidden text is not ignored, this method retrieves a degenerate range from the visible text closest to the specified coordinates.
    * The implementation of RangeFromPoint in Windows Internet Explorer 9 does not return the expected result. Instead, clients should,
    * 1. Call the GetVisibleRanges method to retrieve an array of visible text ranges.Call the GetVisibleRanges method to retrieve an array of visible text ranges.
    * 2. Call the GetVisibleRanges method to retrieve an array of visible text ranges.For each text range in the array, call IUIAutomationTextRange,,GetBoundingRectangles to retrieve the bounding rectangles.
    * 3. Call the GetVisibleRanges method to retrieve an array of visible text ranges.Check the bounding rectangles to find the text range that occupies the particular screen coordinates.
    */
    RangeFromPoint(pt) => (ComCall(3, this, "int64", pt, "ptr*", &range := 0), UIA.IUIAutomationTextRange(range))

    ; Retrieves a text range enclosing a child element such as an image, hyperlink, Microsoft Excel spreadsheet, or other embedded object.
    ; If there is no text in the range that encloses the child element, a degenerate (empty) range is returned.
    ; The child parameter is either a child of the element associated with a IUIAutomationTextPattern or from the array of children of a IUIAutomationTextRange.
    RangeFromChild(child) => (ComCall(4, this, "ptr", child, "ptr*", &range := 0), UIA.IUIAutomationTextRange(range))

    ; Retrieves a collection of text ranges that represents the currently selected text in a text-based control.
    ; If the control supports the selection of multiple, non-contiguous spans of text, the ranges collection receives one text range for each selected span.
    ; If the control contains only a single span of selected text, the ranges collection receives a single text range.
    ; If the control contains a text insertion point but no text is selected, the ranges collection receives a degenerate (empty) text range at the position of the text insertion point.
    ; If the control does not contain a text insertion point or does not support text selection, ranges is set to NULL.
    ; Use the IUIAutomationTextPattern,,SupportedTextSelection property to test whether a control supports text selection.
    GetSelection() => (ComCall(5, this, "ptr*", &ranges := 0), UIA.IUIAutomationTextRangeArray(ranges))

    ; Retrieves an array of disjoint text ranges from a text-based control where each text range represents a contiguous span of visible text.
    ; If the visible text consists of one contiguous span of text, the ranges array will contain a single text range that represents all of the visible text.
    ; If the visible text consists of multiple, disjoint spans of text, the ranges array will contain one text range for each visible span, beginning with the first visible span, and ending with the last visible span. Disjoint spans of visible text can occur when the content of a text-based control is partially obscured by an overlapping window or other object, or when a text-based control with multiple pages or columns has content that is partially scrolled out of view.
    ; IUIAutomationTextPattern,,GetVisibleRanges retrieves a degenerate (empty) text range if no text is visible, if all text is scrolled out of view, or if the text-based control contains no text.
    GetVisibleRanges() => (ComCall(6, this, "ptr*", &ranges := 0), UIA.IUIAutomationTextRangeArray(ranges))

    ; Retrieves a text range that encloses the main text of a document.
    ; Some auxiliary text such as headers, footnotes, or annotations might not be included.
    DocumentRange => (ComCall(7, this, "ptr*", &range := 0), UIA.IUIAutomationTextRange(range))

    ; Retrieves a value that specifies the type of text selection that is supported by the control.
    SupportedTextSelection => (ComCall(8, this, "int*", &supportedTextSelection := 0), supportedTextSelection)

    ; ------------- TextPattern2 ------------

    RangeFromAnnotation(annotation) => (ComCall(9, this, "ptr", annotation, "ptr*", &out:=0), UIA.IUIAutomationTextRange(out))
    GetCaretRange(&isActive:="") => (ComCall(10, this, "ptr*", &isActive, "ptr*", &out:=0), UIA.IUIAutomationTextRange(out))
}

/*
	Provides access to a span of continuous text in a container that supports the TextPattern interface. TextRange can be used to select, compare, and retrieve embedded objects from the text span. The interface uses two endpoints to delimit where the text span starts and ends. Disjoint spans of text are represented by a TextRangeArray, which is an array of TextRange interfaces.
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationtextrange
*/
class IUIAutomationTextRange extends UIA.IUIAutomationBase {
    static __IID := "{A543CC6A-F4AE-494B-8239-C814481187A8}"
    ; Retrieves a IUIAutomationTextRange identical to the original and inheriting all properties of the original.
    ; The range can be manipulated independently of the original.
    Clone() => (ComCall(3, this, "ptr*", &clonedRange := 0), UIA.IUIAutomationTextRange(clonedRange))

    ; Retrieves a value that specifies whether this text range has the same endpoints as another text range.
    ; This method compares the endpoints of the two text ranges, not the text in the ranges. The ranges are identical if they share the same endpoints. If two text ranges have different endpoints, they are not identical even if the text in both ranges is exactly the same.
    Compare(range) => (ComCall(4, this, "ptr", range, "int*", &areSame := 0), areSame)

    ; Retrieves a value that specifies whether the start or end endpoint of this text range is the same as the start or end endpoint of another text range.
    CompareEndpoints(srcEndPoint, range, targetEndPoint) => (ComCall(5, this, "int", srcEndPoint, "ptr", range, "int", targetEndPoint, "int*", &compValue := 0), compValue)

    ; Normalizes the text range by the specified text unit. The range is expanded if it is smaller than the specified unit, or shortened if it is longer than the specified unit.
    ; Client applications such as screen readers use this method to retrieve the full word, sentence, or paragraph that exists at the insertion point or caret position.
    ; Despite its name, the ExpandToEnclosingUnit method does not necessarily expand a text range. Instead, it "normalizes" a text range by moving the endpoints so that the range encompasses the specified text unit. The range is expanded if it is smaller than the specified unit, or shortened if it is longer than the specified unit. If the range is already an exact quantity of the specified units, it remains unchanged. The following diagram shows how ExpandToEnclosingUnit normalizes a text range by moving the endpoints of the range.
    ; ExpandToEnclosingUnit defaults to the next largest text unit supported if the specified text unit is not supported by the control. The order, from smallest unit to largest, is as follows, Character Format Word Line Paragraph Page Document
    ; ExpandToEnclosingUnit respects both visible and hidden text.
    ExpandToEnclosingUnit(textUnit:=6) => ComCall(6, this, "int", textUnit)

    ; Retrieves a text range subset that has the specified text attribute value.
    ; The FindAttribute method retrieves matching text regardless of whether the text is hidden or visible. Use UIA_IsHiddenAttributeId to check text visibility.
    FindAttribute(attr, val, backward:=False) {
        if A_PtrSize = 4
            val := UIA.ComVar(val, , true), ComCall(7, this, "int", attr, "int64", NumGet(val, "int64"), "int64", NumGet(val, 8, "int64"), "int", backward, "ptr*", &found := 0)
        else
            ComCall(7, this, "int", attr, "ptr", UIA.ComVar(val, , true), "int", backward, "ptr*", &found := 0)
        if (found)
            return UIA.IUIAutomationTextRange(found)
    }

    ; Retrieves a text range subset that contains the specified text. There is no differentiation between hidden and visible text.
    FindText(text, backward:=False, ignoreCase:=False) {
        if (ComCall(8, this, "wstr", text, "int", backward, "int", ignoreCase, "ptr*", &found := 0), found)
            return UIA.IUIAutomationTextRange(found)
        throw TargetError("Target textrange not found.")
    }

    ; Retrieves the value of the specified text attribute across the entire text range.
    ; The type of value retrieved by this method depends on the attr parameter. For example, calling GetAttributeValue with the attr parameter set to UIA_FontNameAttributeId returns a string that represents the font name of the text range, while calling GetAttributeValue with attr set to UIA_IsItalicAttributeId would return a boolean.
    ; If the attribute specified by attr is not supported, the value parameter receives a value that is equivalent to the IUIAutomation,,ReservedNotSupportedValue property.
    ; A text range can include more than one value for a particular attribute. For example, if a text range includes more than one font, the FontName attribute will have multiple values. An attribute with more than one value is called a mixed attribute. You can determine if a particular attribute is a mixed attribute by comparing the value retrieved from GetAttributeValue with the UIAutomation,,ReservedMixedAttributeValue property.
    ; The GetAttributeValue method retrieves the attribute value regardless of whether the text is hidden or visible. Use UIA_ IsHiddenAttributeId to check text visibility.
    GetAttributeValue(attr) => (ComCall(9, this, "int", attr, "ptr", val := UIA.ComVar()), val[])

    ; Retrieves a collection of bounding rectangles for each fully or partially visible line of text in a text range.
    GetBoundingRectangles() {
        ComCall(10, this, "ptr*", &boundingRects := 0)
        DllCall("oleaut32\SafeArrayGetVartype", "ptr", boundingRects, "ushort*", &baseType:=0)
        retArr := [], sa := ComValue(0x2000 | baseType, boundingRects)
        Loop sa.MaxIndex() / 4 + 1
			retArr.Push({x:Floor(sa[4*(A_Index-1)]),y:Floor(sa[4*(A_Index-1)+1]),w:Floor(sa[4*(A_Index-1)+2]),h:Floor(sa[4*(A_Index-1)+3])})
		return retArr
    } 

    ; Returns the innermost UI Automation element that encloses the text range.
    GetEnclosingElement() => (ComCall(11, this, "ptr*", &enclosingElement := 0), UIA.IUIAutomationElement(enclosingElement))

    ; Returns the plain text of the text range.
    GetText(maxLength := -1) => (ComCall(12, this, "int", maxLength, "ptr*", &text := 0), UIA.BSTR(text))

    ; Moves the text range forward or backward by the specified number of text units. unit needs to be a TextUnit enum.
    Move(unit, count) => (ComCall(13, this, "int", unit, "int", count, "int*", &moved := 0), moved)

    ; Moves one endpoint of the text range the specified number of text units within the document range.
    MoveEndpointByUnit(endpoint, unit, count) {	; TextPatternRangeEndpoint , TextUnit
        ComCall(14, this, "int", endpoint, "int", unit, "int", count, "int*", &moved := 0)	; TextPatternRangeEndpoint,TextUnit
        return moved
    }

    ; Moves one endpoint of the current text range to the specified endpoint of a second text range.
    ; If the endpoint being moved crosses the other endpoint of the same text range, that other endpoint is moved also, resulting in a degenerate (empty) range and ensuring the correct ordering of the endpoints (that is, the start is always less than or equal to the end).
    MoveEndpointByRange(srcEndPoint, range, targetEndPoint) {	; TextPatternRangeEndpoint , IUIAutomationTextRange , TextPatternRangeEndpoint
        ComCall(15, this, "int", srcEndPoint, "ptr", range, "int", targetEndPoint)
    }

    ; Selects the span of text that corresponds to this text range, and removes any previous selection.
    ; If the Select method is called on a text range object that represents a degenerate (empty) text range, the text insertion point moves to the starting endpoint of the text range.
    Select() => ComCall(16, this)

    ; Adds the text range to the collection of selected text ranges in a control that supports multiple, disjoint spans of selected text.
    ; The text insertion point moves to the newly selected text. If AddToSelection is called on a text range object that represents a degenerate (empty) text range, the text insertion point moves to the starting endpoint of the text range.
    AddToSelection() => ComCall(17, this)

    ; Removes the text range from an existing collection of selected text in a text container that supports multiple, disjoint selections.
    ; The text insertion point moves to the area of the removed highlight. Providing a degenerate text range also moves the insertion point.
    RemoveFromSelection() => ComCall(18, this)

    ; Causes the text control to scroll until the text range is visible in the viewport.
    ; The method respects both hidden and visible text. If the text range is hidden, the text control will scroll only if the hidden text has an anchor in the viewport.
    ; A Microsoft UI Automation client can check text visibility by calling IUIAutomationTextRange,,GetAttributeValue with the attr parameter set to UIA_IsHiddenAttributeId.
    ScrollIntoView(alignToTop) => ComCall(19, this, "int", alignToTop)

    ; Retrieves a collection of all embedded objects that fall within the text range.
    GetChildren() => (ComCall(20, this, "ptr*", &children := 0), UIA.IUIAutomationElementArray(children))
;}

;class IUIAutomationTextRange2 extends IUIAutomationTextRange {
;	static __IID := "{BB9B40E0-5E04-46BD-9BE0-4B601B9AFAD4}"
    ShowContextMenu() => ComCall(21, this, "ptr*", &out:="")
;}

;class IUIAutomationTextRange3 extends IUIAutomationTextRange2 { ; UNTESTED
;	static __IID := "{6A315D69-5512-4C2E-85F0-53FCE6DD4BC2}"

	GetEnclosingElementBuildCache(cacheRequest) => (ComCall(22, this, "ptr", cacheRequest, "ptr*", &out:=0), UIA.IUIAutomationElement(out))
    GetChildrenBuildCache(cacheRequest) => (ComCall(23, this, "ptr", cacheRequest, "ptr*", &out:=0), UIA.IUIAutomationElementArray(out))
    GetAttributeValues(attributeIds, attributeIdCount) {
        ComCall(24, this, "ptr", ComObjValue(ComObjArray(8, attributeIds*)), "int", attributeIdCount, "ptr*", &out:=UIA.ComVar())
        return out[]
    }
}

class IUIAutomationTextRangeArray extends UIA.IUIAutomationBase {
    __Item[index] {
        get => this.GetElement(index-1)
    }
    __Enum(varCount) {
        maxLen := this.Length-1, i := 0
        EnumElements(&element) {
            if i > maxLen
                return false
            element := this.GetElement(i++)
            return true
        }
        EnumIndexAndElements(&index, &element) {
            if i > maxLen
                return false
            element := this.GetElement(i++)
            index := i
            return true
        }
        return (varCount = 1) ? EnumElements : EnumIndexAndElements
    }
    ; Retrieves the number of text ranges in the collection.
    Length => (ComCall(3, this, "int*", &length := 0), length)

    ; Retrieves a text range from the collection.
    GetElement(index) => (ComCall(4, this, "int", index, "ptr*", &element := 0), UIA.IUIAutomationTextRange(element))
}

class IUIAutomationTogglePattern extends UIA.IUIAutomationBase {
    ; Cycles through the toggle states of the control.
    ; A control cycles through its states in this order, ToggleState_On, ToggleState_Off and, if supported, ToggleState_Indeterminate.
    Toggle() => ComCall(3, this)

    ; Retrieves the state of the control.
    ToggleState {
        get => (ComCall(4, this, "int*", &retVal := 0), retVal)
        set {
			if (this.ToggleState != value)
				this.Toggle()
        }
    } 

    ; Retrieves the cached state of the control.
    CachedToggleState => (ComCall(5, this, "int*", &retVal := 0), retVal)
}

/*
	Provides access to a control that can be moved, resized, or rotated.
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationtransformpattern 
*/
class IUIAutomationTransformPattern extends UIA.IUIAutomationBase {
    ; An element cannot be moved, resized or rotated such that its resulting screen location would be completely outside the coordinates of its container and inaccessible to the keyboard or mouse. For example, when a top-level window is moved completely off-screen or a child object is moved outside the boundaries of the container's viewport, the object is placed as close to the requested screen coordinates as possible with the top or left coordinates overridden to be within the container boundaries.

    ; Moves the UI Automation element.
    Move(x, y) => ComCall(3, this, "double", x, "double", y)

    ; Resizes the UI Automation element.
    ; When called on a control that supports split panes, this method can have the side effect of resizing other contiguous panes.
    Resize(width, height) => ComCall(4, this, "double", width, "double", height)

    ; Rotates the UI Automation element.
    Rotate(degrees) => ComCall(5, this, "double", degrees)

    ; Indicates whether the element can be moved.
    CanMove => (ComCall(6, this, "int*", &retVal := 0), retVal)

    ; Indicates whether the element can be resized.
    CanResize => (ComCall(7, this, "int*", &retVal := 0), retVal)

    ; Indicates whether the element can be rotated.
    CanRotate => (ComCall(8, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates whether the element can be moved.
    CachedCanMove => (ComCall(9, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates whether the element can be resized.
    CachedCanResize => (ComCall(10, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates whether the element can be rotated.
    CachedCanRotate => (ComCall(11, this, "int*", &retVal := 0), retVal)

    ; --------------- TransformPattern2 ---------------

    CanZoom => (ComCall(14, this, "int*", &retVal := 0), retVal)
    CachedCanZoom => (ComCall(15, this, "int*", &retVal := 0), retVal)
    ZoomLevel => (ComCall(16, this, "double*", &retVal := 0), retVal)
    CachedZoomLevel => (ComCall(17, this, "double*", &retVal := 0), retVal)
    ZoomMinimum => (ComCall(18, this, "double*", &retVal := 0), retVal)
    CachedZoomMinimum => (ComCall(19, this, "double*", &retVal := 0), retVal)
    ZoomMaximum => (ComCall(20, this, "double*", &retVal := 0), retVal)
    CachedZoomMaximum => (ComCall(21, this, "double*", &retVal := 0), retVal)
    Zoom(zoomValue) => ComCall(12, this, "double", zoomValue)
    ZoomByUnit(ZoomUnit) => ComCall(13, this, "uint", ZoomUnit)
}

class IUIAutomationValuePattern extends UIA.IUIAutomationBase {
    ; Sets the value of the element.
    ; The IsEnabled property must be TRUE, and the IUIAutomationValuePattern,,IsReadOnly property must be FALSE.
    SetValue(val) => ComCall(3, this, "wstr", val)

    ; Retrieves the value of the element.
    ; Single-line edit controls support programmatic access to their contents through IUIAutomationValuePattern. However, multiline edit controls do not support this control pattern, and their contents must be retrieved by using IUIAutomationTextPattern.
    ; This property does not support the retrieval of formatting information or substring values. IUIAutomationTextPattern must be used in these scenarios as well.
    Value => (ComCall(4, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Indicates whether the value of the element is read-only.
    IsReadOnly => (ComCall(5, this, "int*", &retVal := 0), retVal)

    ; Retrieves the cached value of the element.
    CachedValue => (ComCall(6, this, "ptr*", &retVal := 0), UIA.BSTR(retVal))

    ; Retrieves a cached value that indicates whether the value of the element is read-only.
    ; This property must be TRUE for IUIAutomationValuePattern,,SetValue to succeed.
    CachedIsReadOnly => (ComCall(7, this, "int*", &retVal := 0), retVal)
}

/*
	Represents a virtualized item, which is an item that is represented by a placeholder automation element in the Microsoft UI Automation tree.
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationvirtualizeditempattern
*/
class IUIAutomationVirtualizedItemPattern extends UIA.IUIAutomationBase {
    ; Creates a full UI Automation element for a virtualized item.
    ; A virtualized item is represented by a placeholder automation element in the UI Automation tree. The Realize method causes the provider to make full information available for the item so that a full UI Automation element can be created for the item.
    Realize() => ComCall(3, this)
}

/*
	Provides access to the fundamental functionality of a window.
	Microsoft documentation: https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationwindowpattern
*/
class IUIAutomationWindowPattern extends UIA.IUIAutomationBase {
    ; Closes the window.
    ; When called on a split pane control, this method closes the pane and removes the associated split. This method may also close all other panes, depending on implementation.
    Close() => ComCall(3, this)

    ; Causes the calling code to block for the specified time or until the associated process enters an idle state, whichever completes first.
    WaitForInputIdle(milliseconds) => (ComCall(4, this, "int", milliseconds, "int*", &success := 0), success)

    ; Minimizes, maximizes, or restores the window.
    SetWindowVisualState(state) => ComCall(5, this, "int", state)

    ; Indicates whether the window can be maximized.
    CanMaximize => (ComCall(6, this, "int*", &retVal := 0), retVal)

    ; Indicates whether the window can be minimized.
    CanMinimize => (ComCall(7, this, "int*", &retVal := 0), retVal)

    ; Indicates whether the window is modal.
    IsModal => (ComCall(8, this, "int*", &retVal := 0), retVal)

    ; Indicates whether the window is the topmost element in the z-order.
    IsTopmost => (ComCall(9, this, "int*", &retVal := 0), retVal)

    ; Retrieves the visual state of the window; that is, whether it is in the normal, maximized, or minimized state.
    WindowVisualState => (ComCall(10, this, "int*", &retVal := 0), retVal)

    ; Retrieves the current state of the window for the purposes of user interaction.
    WindowInteractionState => (ComCall(11, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates whether the window can be maximized.
    CachedCanMaximize => (ComCall(12, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates whether the window can be minimized.
    CachedCanMinimize => (ComCall(13, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates whether the window is modal.
    CachedIsModal => (ComCall(14, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates whether the window is the topmost element in the z-order.
    CachedIsTopmost => (ComCall(15, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates the visual state of the window; that is, whether it is in the normal, maximized, or minimized state.
    CachedWindowVisualState => (ComCall(16, this, "int*", &retVal := 0), retVal)

    ; Retrieves a cached value that indicates the current state of the window for the purposes of user interaction.
    CachedWindowInteractionState => (ComCall(17, this, "int*", &retVal := 0), retVal)
}

}