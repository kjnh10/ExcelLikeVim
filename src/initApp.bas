Attribute VB_Name = "initApp"
'-----------------------------------------
Public myobject As New ApplicationEvent

'-------main----------
Public Sub InitializeApplication()'{{{
' On Error Goto MyError
On Error Resume Next
	Call SetReference
	Call AllKeyToAssesKeyFunc
	Call SpecialMapping
	Call SetAppEvent
	' IsExistPython = True 意味なし｡グローバル変数は消える｡
	Call read_setting(Environ("homepath") & "/.vimxrc")
	' If visualmodefeature Then
	If True Then
		Call OpenRegisterBook()
		If Workbooks.Count = 1 Then
			Workbooks.Add
		End If
	End If
	Application.Cursor = xlNorthwestArrow
On Error Goto 0
MyError:
If Err.Description <> "" Then
	MsgBox Err.Description
End If

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
			' argument = Mid(buf, Instr(Instr(buf, " ") + 1, buf, " ") + 1) '2つ目のスペース以降を取得
			argument_start = Instr(buf, " ")
			If argument_start <> 0 Then
				argument = Mid(buf, Instr(buf, " ") + 1) '1つ目のスペース以降を取得
			End If
			If Instr(instruction, "map") = 0 And Instr(instruction, "for") = 0 Then 'map系じゃなければそのまま実行 TODO map系もそのうち
				Debug.Print "instruction:" & instruction & vbCrLf & "argument:" & argument
				If argument_start = 0 Then
					Application.Run instruction
				Else
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
	Debug.Print "Called SetAppEvent"
	Set myobject.appevent = Application
End Sub'}}}

Public Sub SetReference()'{{{
	'unite_command 用 本来はプラグイン側からの呼び出しを出来るようにしたい｡
	Debug.Print AddToReference("C:\Program Files\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB")
End Sub'}}}

Function AddToReference(strFileName As String) As Boolean'{{{
	'指定されたタイプライブラリへの参照を作成します｡
	On Error GoTo MyError
		Dim ref As Reference
		Set ref = ThisWorkbook.VBProject.References.AddFromFile(strFileName)
		AddToReference = True
		Set ref = Nothing
		Exit Function
	MyError:
		Select Case Err.Number
			Case 32813
				Debug.Print strFileName & "は既に参照設定されています。", , "タイプライブラリへの参照"
			Case 29060
				MsgBox "設定ファイルがインストールされていないか、" & vbNewLine & _
					"所定のフォルダーに存在しない場合が考えられます。" & vbNewLine & _
					"よって、参照設定ができません。", , "タイプライブラリへの参照"
			Case Else
				MsgBox "予期せぬエラーが発生しました。" & vbNewLine & _
					Err.Number & vbNewLine & _
					Err.Description, 16, "タイプライブラリへの参照"
		End Select
End Function'}}}
