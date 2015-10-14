Attribute VB_Name = "vim"

'Declare for WinAPI
Public Declare Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Long
Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Declare Function GetKeyState Lib "User32.dll" (ByVal vKey As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'
'Definition for Keyboard event (keybd_event in user32)
Const KEYUP = &H2          'Key up
Const EXTENDED_KEY = &H1   'For using extended keys
'
'Virtual Key codes'{{{
Const LSHIFT = &HA0 'Left Shift
Const RSHIFT = &HA1 'Right Shift
Const LCTRL = &HA2  'Left Ctrl
Const RCTRL = &HA3  'Right Ctrl
Const LMENU = &HA4  'Left Alt
Const RMENU = &HA5  'Right Alt
Const KANJI = &H19  'Kanji'}}}

Dim mode As String
Dim buf As Long

'------------Normal Mode----------------------
Function move_up() '{{{
  keybd_event vbKeyUp, 0, EXTENDED_KEY Or 0, 0
  keybd_event vbKeyUp, 0, EXTENDED_KEY Or KEYUP, 0
End Function '}}}

Function move_down() '{{{
  keybd_event vbKeyDown, 0, EXTENDED_KEY Or 0, 0
  keybd_event vbKeyDown, 0, EXTENDED_KEY Or KEYUP, 0
End Function '}}}

Function move_left() '{{{
  keybd_event vbKeyLeft, 0, EXTENDED_KEY Or 0, 0
  keybd_event vbKeyLeft, 0, EXTENDED_KEY Or KEYUP, 0
End Function '}}}

Function move_right() '{{{
  keybd_event vbKeyRight, 0, EXTENDED_KEY Or 0, 0
  keybd_event vbKeyRight, 0, EXTENDED_KEY Or KEYUP, 0
End Function '}}}

Sub move_head() '{{{
	Dim startCell As Range
	Set startCell = ActiveCell

	Dim dest As Range
	Set dest = cells(ActiveCell.Row, 1)
	If dest.value = "" Then
		Set dest = dest.End(xlToRight)
	End If

	If dest.Column = Columns.Count Then
		Set dest = Cells(dest.Row, 1)
	End If

	dest.Activate
End Sub '}}}

Sub move_tail() '{{{
	Dim dest As Range
	Set dest = cells(ActiveCell.Row, Columns.Count)
	If dest.value = "" Then
		Set dest = dest.End(xlToLeft)
	End If

	dest.Activate
End Sub '}}}

Public Sub gg() '{{{
        cells(1, ActiveCell.Column).Select
End Sub '}}}

Public Sub G() '{{{
        With ActiveSheet.UsedRange
                cells(.Rows(.Rows.count).Row, ActiveCell.Column).Select
        End With
End Sub '}}}

Sub vim_w() '{{{
        ActiveCell.End(xlToRight).Select
End Sub '}}}

Sub vim_b() '{{{
        ActiveCell.End(xlToLeft).Select
End Sub '}}}

Function scroll_up() '{{{
  Dim scroll_width As Integer
  
  Application.ScreenUpdating = False
  selected_range_top = ActiveWindow.VisibleRange.Row
  
  scroll_width = ActiveWindow.VisibleRange.Rows.count / 2
  scroll_target_left = ActiveCell.Column

  scroll_target_top = selected_range_top - scroll_width
  
  If scroll_target_top < 1 Then
    scroll_target_top = 1
  End If
  
  ActiveWindow.SmallScroll Up:=scroll_width

  cells(scroll_target_top, scroll_target_left).Activate
  Application.ScreenUpdating = True
  
End Function '}}}

Function scroll_down() '{{{
  Dim scroll_width As Integer
  
  Application.ScreenUpdating = False
  selected_range_top = ActiveWindow.VisibleRange.Row
  
  scroll_width = ActiveWindow.VisibleRange.Rows.count / 2
  scroll_target_left = ActiveCell.Column

  scroll_target_top = selected_range_top + scroll_width

  ActiveWindow.SmallScroll Down:=scroll_width

  cells(scroll_target_top, scroll_target_left).Activate
  Application.ScreenUpdating = True
  
End Function '}}}

Sub find() '{{{
	Dim obj As Object
	searchString = InputBox("/", "command", "")
	If searchString = "" Then
		Exit Sub
	End If
	Set obj = ActiveSheet.cells.find(what:=searchString, lookat:=xlPart)
	If Not obj Is Nothing Then
		obj.Activate
	Else
		MsgBox "見つかりませんでした｡"
	End If
	'Selection.FindNext(After:=ActiveCell).Activate
End Sub '}}}

Function findNext() '{{{
  Dim t As Range
  Set t = cells.findNext(After:=ActiveCell)
  If t Is Nothing Then
  Else
    t.Activate
  End If
End Function
'}}}

Function findPrevious() '{{{
  Dim t As Range
  Set t = cells.findPrevious(After:=ActiveCell)
  If t Is Nothing Then
  Else
    t.Activate
  End If
End Function '}}}

Function insertRowDown() '{{{
  keyupControlKeys
  releaseShiftKeys
  keybd_event vbKeyDown, 0, 0, 0
  keybd_event vbKeyDown, 0, KEYUP, 0
  keybd_event vbKeyMenu, 0, 0, 0
  keybd_event vbKeyI, 0, 0, 0
  keybd_event vbKeyI, 0, KEYUP, 0
  keybd_event vbKeyR, 0, 0, 0
  keybd_event vbKeyR, 0, KEYUP, 0
  keybd_event vbKeyMenu, 0, KEYUP, 0
  unkeyupControlKeys
  period_buff = "o"
End Function '}}}

Function insertRowUp() '{{{
  keyupControlKeys
  releaseShiftKeys
  keybd_event vbKeyMenu, 0, 0, 0
  keybd_event vbKeyI, 0, 0, 0
  keybd_event vbKeyI, 0, KEYUP, 0
  keybd_event vbKeyR, 0, 0, 0
  keybd_event vbKeyR, 0, KEYUP, 0
  keybd_event vbKeyMenu, 0, KEYUP, 0
  unkeyupControlKeys
  period_buff = "+o"
End Function '}}}

Function insertColumnRight() '{{{
  keyupControlKeys
  keybd_event vbKeyRight, 0, 0, 0
  keybd_event vbKeyRight, 0, KEYUP, 0
  keybd_event vbKeyMenu, 0, 0, 0
  keybd_event vbKeyI, 0, 0, 0
  keybd_event vbKeyI, 0, KEYUP, 0
  keybd_event vbKeyC, 0, 0, 0
  keybd_event vbKeyC, 0, KEYUP, 0
  keybd_event vbKeyMenu, 0, KEYUP, 0
  unkeyupControlKeys
  period_buff = "t"
End Function '}}}

Function insertColumnLeft() '{{{
  keyupControlKeys
  releaseShiftKeys
  keybd_event vbKeyMenu, 0, 0, 0
  keybd_event vbKeyI, 0, 0, 0
  keybd_event vbKeyI, 0, KEYUP, 0
  keybd_event vbKeyC, 0, 0, 0
  keybd_event vbKeyC, 0, KEYUP, 0
  keybd_event vbKeyMenu, 0, KEYUP, 0
  unkeyupControlKeys
  period_buff = "+t"
End Function '}}}

Public Sub n_ESC()'{{{
	Application.CutCopyMode = False
End Sub'}}}

Public Sub n_ESC_ime_off()'{{{
	Application.CutCopyMode = False
	If IMEStatus <> 2 Then
		Call SendKeys("{KANJI}", True)
	End if
End Sub'}}}

Public Sub n_u()'{{{
	'undo履歴を辞書でとる｡順番を覚えさせる｡システムの方はworkbook_changeイベントで｡自分が今どこにいるかも覚えている｡システム変更の場合の値は""｡
	keybd_event vbKeyControl, 0, 0, 0
	keybd_event vbKeyZ, 0, 0, 0
	keybd_event vbKeyZ, 0, KEYUP, 0
	keybd_event vbKeyControl, 0, KEYUP, 0
End Sub'}}}

Public Sub n_yy() '{{{
	Application.ScreenUpdating = False
	Call n_v_()
	Call v_y()
End Sub '}}}

Public Sub yank_value() '{{{
	' ActiveCell.Value
	MsgBox "Todo ･･･"
End Sub '}}}

Public Sub n_dd() '{{{
	Call n_yy()
	Rows(ActiveCell.Row).Delete
End Sub '}}}

Public Sub n_dc() '{{{
	MsgBox "実装中です"
End Sub '}}}

	'----------Mode Chage--------------------------
	Public Sub n_v()'{{{
		buf = ActiveCell.Column
		SetKeyMapDic("visual")
		mode = "visual"
	End Sub'}}}

	Public Sub n_v_()'{{{
		buf = ActiveCell.Column

		Rows(ActiveCell.Row).Select
		SetKeyMapDic("line_visual")
		mode = "line_visual"

	End Sub'}}}

	Public Sub vertival_visual_mode()'{{{
		buf = ActiveCell.Column

		Columns(ActiveCell.Row).Select
		SetKeyMapDic("vertical_visual")
		mode = "vertical_visual"

	End Sub'}}}

	Function insert_mode() '{{{
		keyupControlKeys
		releaseShiftKeys
		keybd_event vbKeyF2, 0, 0, 0
		keybd_event vbKeyF2, 0, KEYUP, 0
		' Application.OnTime Now + TimeValue("00:00:00"), "disableIME"
		unkeyupControlKeys
	End Function '}}}

	Public Sub n_p(Optional registerName As String = """") '{{{
		Application.ScreenUpdating = False
		Dim srcRegSheet As Worksheet '宣言がないとGetdataRangeが型を判定出来ずエラーになる｡
		Set srcRegSheet = Workbooks("register.xlsx").Worksheets(registerName)
		Set srcRange = GetDataRange(srcRegSheet)
		srcRange.Copy

		If srcRegSheet.Cells(2, 4).Value  = "line_visual" Then
			Range(ActiveCell.Row + 1 & ":" & ActiveCell.Row + srcRange.Rows.Count).Insert
			Cells(ActiveCell.Row + 1, 1).Select
		Else
			ActiveCell.Select 'なぜかこれを行わないとvisual_modeに対する貼付けが出来ない｡←多分これを行わないとregsiterbookのrangeを選択している。?
		End If

		ActiveSheet.Paste
		' Application.ScreenUpdating = True
		' 'ctrl+vの送信｡undoのためキーボードで実現
		' ' ActiveSheet.Paste
		' keybd_event vbKeyControl, 0, 0, 0
		' keybd_event vbKeyV, 0, 0, 0
		' keybd_event vbKeyV, 0, KEYUP, 0
		' keybd_event vbKeyControl, 0, KEYUP, 0
		' ' DoEvents
		'
		' 'ctrl+BackSpaceの送信｡選択範囲の解除
		' keybd_event vbKeyShift, 0, 0, 0
		' keybd_event vbKeyBack, 0, EXTENDED_KEY Or 0, 0
		' keybd_event vbKeyBack, 0, EXTENDED_KEY Or KEYUP, 0
		' keybd_event vbKeyShift, 0, KEYUP, 0
	End Sub '}}}

Sub command_vim() '{{{
	Dim AWB As String
	Dim commandString As String

	AWB = ActiveWorkbook.Name
	commandString = InputBox("Please Enter Command you wanna do", "command", "")
	If commandString = "" Then
		Exit Sub
	End If

	commandString = Replace(commandString, "!", "_exclamation")
	Call ExeStringPro(commandString, AWB)
End Sub '}}}

'------------Visual Mode----------------------
	'--------move---------------------'{{{
	Public Sub visual_move(commandString As String)
		'header

		'main
		ExeStringPro(commandString)
		'hooder
	End Sub


Public Sub v_j()'{{{
  keybd_event vbKeyShift, 0, 0, 0
  keybd_event vbKeyDown, 0, EXTENDED_KEY Or 0, 0
  keybd_event vbKeyDown, 0, EXTENDED_KEY Or KEYUP, 0
  keybd_event vbKeyShift, 0, KEYUP, 0
End Sub'}}}

Public Sub v_k()'{{{
  keybd_event vbKeyShift, 0, 0, 0
  keybd_event vbKeyUp, 0, EXTENDED_KEY Or 0, 0
  keybd_event vbKeyUp, 0, EXTENDED_KEY Or KEYUP, 0
  keybd_event vbKeyShift, 0, KEYUP, 0
End Sub'}}}

Public Sub v_h()'{{{
  keybd_event vbKeyShift, 0, 0, 0
  keybd_event vbKeyLeft, 0, EXTENDED_KEY Or 0, 0
  keybd_event vbKeyLeft, 0, EXTENDED_KEY Or KEYUP, 0
  keybd_event vbKeyShift, 0, KEYUP, 0
End Sub'}}}

Public Sub v_l()'{{{
  keybd_event vbKeyShift, 0, 0, 0
  keybd_event vbKeyRight, 0, EXTENDED_KEY Or 0, 0
  keybd_event vbKeyRight, 0, EXTENDED_KEY Or KEYUP, 0
  keybd_event vbKeyShift, 0, KEYUP, 0
End Sub'}}}

Public Sub v_gg()'{{{
	Dim buf As Range
	Set buf = ActiveCell

	Range(Activecell, Cells(1, Selection(Selection.Count).Column)).Select

	buf.Activate
End Sub'}}}

Public Sub v_G() '{{{
	Dim buf As Range
	Set buf = ActiveCell

	With ActiveSheet.UsedRange
		Range(ActiveCell, Cells(.Rows(.Rows.count).Row, Selection(Selection.Count).Column)).Select
	End With

	buf.Activate
End Sub '}}}

Sub v_w() '{{{
	Dim startCell As Range
	Set startCell = ActiveCell

	Dim currentRow As Long
	If startCell.Row = Selection(1).Row Then
		currentRow = Selection(Selection.Count).Row
	Else
		currentRow = Selection(1).Row
	End If

	Dim currentColumn As Long
	If startCell.Column = Selection(1).Column Then
		currentColumn = Selection(Selection.Count).Column
	Else
		currentColumn = Selection(1).Column
	End If
	Dim currentCell As Range
	Set currentCell = Cells(currentRow, currentColumn)

	Dim dest As Range
	Set dest = cells(currentRow, currentCell.End(xlToRight).Column)

	Range(Activecell, dest).Select
	startCell.Activate
End Sub '}}}

Sub v_b() '{{{
End Sub '}}}

Sub v_move_head() '{{{
	Dim startCell As Range
	Set startCell = ActiveCell

	Dim currentRow As Long
	If startCell.Row = Selection(1).Row Then
		currentRow = Selection(Selection.Count).Row
	Else
		currentRow = Selection(1).Row
	End If

	Dim dest As Range
	Set dest = cells(currentRow, 1)
	If dest.value = "" Then
		Set dest = dest.End(xlToRight)
	End If

	If dest.Column = Columns.Count Then
		Set dest = Cells(dest.Row, 1)
	End If

	Range(Activecell, dest).Select
	startCell.Activate
End Sub '}}}

Sub v_move_tail() '{{{
	Dim startCell As Range
	Set startCell = ActiveCell

	Dim currentRow As Long
	If startCell.Row = Selection(1).Row Then
		currentRow = Selection(Selection.Count).Row
	Else
		currentRow = Selection(1).Row
	End If

	Dim dest As Range
	Set dest = cells(currentRow, Columns.Count)
	If dest.value = "" Then
		Set dest = dest.End(xlToLeft)
	End If

	Range(Activecell, dest).Select
	startCell.Activate
End Sub '}}}

Sub v_scroll_up() '{{{
	Dim scroll_width As Integer

	Application.ScreenUpdating = False
	selected_range_top = ActiveWindow.VisibleRange.Row

	scroll_width = ActiveWindow.VisibleRange.Rows.count / 2
	scroll_target_left = ActiveCell.Column

	scroll_target_top = selected_range_top - scroll_width

	If scroll_target_top < 1 Then
		scroll_target_top = 1
	End If

	ActiveWindow.SmallScroll Up:=scroll_width

	cells(scroll_target_top, scroll_target_left).Activate
	Application.ScreenUpdating = True

End Sub '}}}

Sub v_scroll_down() '{{{
	Dim scroll_width As Integer

	Application.ScreenUpdating = False
	selected_range_top = ActiveWindow.VisibleRange.Row

	scroll_width = ActiveWindow.VisibleRange.Rows.count / 2
	scroll_target_left = ActiveCell.Column

	scroll_target_top = selected_range_top + scroll_width

	ActiveWindow.SmallScroll Down:=scroll_width

	cells(scroll_target_top, scroll_target_left).Activate
	Application.ScreenUpdating = True

End Sub '}}}

Sub v_a() '{{{
	ActiveSheet.UsedRange.Select
End Sub '}}}
'}}}

	'--------operator---------------------'{{{
Public Sub visual_operation(commandString As String)
	'header
	Debug.Print "visual_operation start"

	'main
	' ExeStringPro(commandString)
	Application.Run(commandString)
	'hooder

	SetKeyMapDic("normal")
	mode = ""
End Sub
Public Sub v_ESC()'{{{
	Debug.Print "v_ESC"
	SetKeyMapDic("normal")
	mode = ""
	ActiveCell.Select
End Sub'}}}

Public Sub v_y(Optional registerName As String = """")'{{{
	Call registerSelection(registerName)
	Call v_ESC()
End Sub'}}}

Public Sub lv_d(Optional registerName As String = """")'{{{
	Application.ScreenUpdating = False
	Call registerSelection(registerName)
	Selection.Delete Shift:=xlUp
	Call v_ESC()
End Sub'}}}'}}}

Public Sub v_d(Optional registerName As String = """")'{{{
	Application.ScreenUpdating = False
	Call registerSelection(registerName)
	Selection.ClearContents
	Call v_ESC()
End Sub'}}}'}}}

Public Sub v_x(Optional registerName As String = """")'{{{
	Application.ScreenUpdating = False
	Call registerSelection(registerName)
	Selection.Clear
	Call v_ESC()
End Sub'}}}'}}}

Public Sub v_D_(Optional registerName As String = """")'{{{
	Application.ScreenUpdating = False
	Call registerSelection(registerName)
	Selection.Delete Shift:=xlUp
	Call v_ESC()
End Sub'}}}'}}}

'------------Line Visual Mode----------------------


'------------Core Functions----------------------------
Public Sub registerSelection(Optional registerName As String = """")'{{{
	Const destRangeStartRow = 4
	Set destRegSheet = Workbooks("register.xlsx").Worksheets(registerName)

	Dim s As Shape
    For Each s In destRegSheet.Shapes
        s.Delete
    Next
	destRegSheet.Rows(destRangeStartRow & ":" & Rows.count).Clear

	Set destRange =  destRegSheet.Cells(destRangeStartRow,1)

	Selection.Copy(destRange)
	destRegSheet.Cells(2,3).Value = Selection.Rows.Count & ":" & Selection.Columns.Count
	destRegSheet.Cells(2,4).Value = mode
	DoEvents

	' Workbooks("register.xlsx").Save
End Sub'}}}


'------------Supplimental Functions----------------------------
Private Function releaseShiftKeys() '{{{
  If GetKeyState(LSHIFT) > 0 Then
    keybd_event LSHIFT, 0, KEYUP, 0
    DoEvents
  ElseIf GetKeyState(RSHIFT) > 0 Then
    keybd_event RSHIFT, 0, KEYUP, 0
    DoEvents
  Else
    DoEvents
    keybd_event vbKeyShift, 0, KEYUP, 0
  End If
End Function '}}}

Private Function keyupControlKeys() '{{{
  keybd_event LSHIFT, 0, KEYUP, 0
  keybd_event RSHIFT, 0, EXTENDED_KEY Or KEYUP, 0
  keybd_event LCTRL, 0, KEYUP, 0
  keybd_event RCTRL, 0, EXTENDED_KEY Or KEYUP, 0
  keybd_event LMENU, 0, KEYUP, 0
  keybd_event RMENU, 0, EXTENDED_KEY Or KEYUP, 0
End Function '}}}

Private Function unkeyupControlKeys() '{{{
  If GetKeyState(LSHIFT) < 0 Then
  ElseIf GetKeyState(RSHIFT) < 0 Then
    keybd_event RSHIFT, 0, EXTENDED_KEY, 0
  Else
    keybd_event vbKeyShift, 0, KEYUP, 0
  End If
  If GetKeyState(LCTRL) < 0 Then
    keybd_event LCTRL, 0, 0, 0
  ElseIf GetKeyState(RCTRL) < 0 Then
    keybd_event RCTRL, 0, EXTENDED_KEY, 0
  Else
    keybd_event vbKeyControl, 0, KEYUP, 0
  End If
  If GetKeyState(LMENU) < 0 Then
    keybd_event LMENU, 0, 0, 0
  ElseIf GetKeyState(RMENU) < 0 Then
    keybd_event RMENU, 0, EXTENDED_KEY, 0
  Else
    keybd_event vbKeyMenu, 0, KEYUP, 0
  End If
End Function
'}}}

Private Function GetDataRange(sh As Worksheet) As Range'{{{
	Const dataStartRow = 4
	' Set GetDataRange = InterSect(sh.UsedRange, sh.Rows(dataStartRow & ":" & sh.Rows.Count))
	dataSizeOfRows = Split(sh.Cells(2,3).Value, ":")(0)
	dataSizeOfColumns = Split(sh.Cells(2,3).Value, ":")(1)
	Set GetDataRange = sh.Cells(dataStartRow, 1).Resize(dataSizeOfRows, dataSizeOfColumns)
End Function'}}}

Private Sub RegisterToDataRange(srcRange As Range, Optional registerName As String = """")'{{{
	Const destRangeStartRow = 4
	Set destRegSheet = Workbooks("register.xlsx").Worksheets(registerName)
	destRegSheet.Rows(destRangeStartRow & ":" & Rows.count).Clear
	Set destRange =  destRegSheet.Cells(destRangeStartRow,1)

	srcRange.Copy(destRange)
End Sub'}}}

Sub disableIME()'{{{
	If IMEStatus <> vbIMEModeOff Then
		keybd_event KANJI, 0, 0, 0
		keybd_event KANJI, 0, KEYUP, 0
	End If
End Sub'}}}
