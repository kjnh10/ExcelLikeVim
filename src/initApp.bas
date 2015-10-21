Attribute VB_Name = "initApp"
'-----------------------------------------
Public myobject As New ApplicationEvent

'-------main----------
Public Sub InitializeApplication()'{{{
' On Error Goto MyError
On Error Resume Next
	Call AllKeyToAssesKeyFunc
	Call SpecialMapping
	Application.Cursor = xlNorthwestArrow
	' IsExistPython = True 意味なし｡グローバル変数は消える｡
	Call read_setting(Environ("homepath") & "/.vimxrc")
	' If visualmodefeature Then
	If True Then
		Call OpenRegisterBook()
		If Workbooks.Count = 1 Then
			Workbooks.Add
		End If
	End If
On Error Goto 0
MyError:
If Err.Description <> "" Then
	MsgBox  Err.Number & Err.Description
End If

End Sub'}}}

Public Sub InitializeLater()'{{{
	Call SetAppEvent
End Sub'}}}

Public Sub read_setting(filePath As String)'{{{
	filePath = absPath(filePath) 
	Open FilePath For Input As #1
	Do Until EOF(1)
		Line Input #1, buf
		buf = Replace(buf,vbTab,"") 'ignore indent

		If Left(buf,1) = "'" Then 'ignore comment
			Goto NextLoop
		End If

		If buf <> "" Then
			instruction = Split(buf, " ")(0)
			argument_start = Instr(buf, " ")
			If argument_start = 0 Then
				Application.Run instruction
			Else
				Dim argument As String:argument = Mid(buf, Instr(buf, " ") + 1) 'スペース以降をargumentにセット
				If Instr(instruction, "map") = 0 And Instr(instruction, "for") = 0 Then 'map系じゃなければそのまま実行 TODO map系もそのうち
					Application.Run instruction, argument
				End If
			End If
		End If

		NextLoop:
	Loop
	Close #1
End Sub'}}}

'------supplimental functions-------------
Public Sub SpecialMapping()'{{{
	'ここで指定した関数はkeystroke.basが不具合でも働く｡mapping.txtを上書く
	' Application.OnKey "{f11}", "'updateModules ""VimX"", 0'"
	Application.OnKey "{f11}", "'updateModulesOfBook """", False'"
End Sub'}}}

Private Sub OpenRegisterBook()'{{{
	Application.ScreenUpdating = False
	Workbooks.Open FileName:=ThisWorkbook.Path & "\data\register.xlsx", ReadOnly:=True
	Windows("register.xlsx").Visible = False
End Sub'}}}

Public Sub SetAppEvent()'{{{
	Set myobject.appEvent = Application
	Set myobject.pptEvent = New PowerPoint.Application
	Set myobject.wrdEvent = New Word.Application
	MsgBox "setiing AppEvent is done"
End Sub'}}}

Sub Wrap(arg As String)'{{{
	buf = Split(arg, ",")
	a = buf(0):b = buf(1)
	With ThisWorkbook.VBProject.VBComponents("wrapper").CodeModule
		.InsertLines 1, "Sub " & a & "()"
		.InsertLines 2, "End Sub"
		.InsertLines 2, "ExeStringPro(""" & b & """)"
	End With
End Sub'}}}

Sub ClearWrapper(a As String, b As String)'{{{
	With ThisWorkbook.VBProject.VBComponents("wrapper").CodeModule
		.DeleteLines StartLine:=1, count:=.CountOfLines
	End With
End Sub'}}}

Sub AddToInitializeLater(code As String)
	Dim objCode As VBIDE.CodeModule
	Set objCode = ThisWorkbook.VBProject.VBComponents("initApp").CodeModule
	MsgBox objCode.ProcStartLine("InitializeLater", 0)
	MsgBox objCode.ProcCountLines("InitializeLater", 0)
End Sub


