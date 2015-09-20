VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------
Public WithEvents appevent As Application
Attribute appevent.VB_VarHelpID = -1

Private Sub appevent_WorkbookOpen(ByVal Wb As Workbook) '{{{
	Call UpdateMru(Wb)
	Call cd(Wb.Name)

	For Each b in Workbooks
		If b.Name Like "Book*" Then
			b.Close
		End If
	Next b

	If Wb.FullName <> ThisWorkbook.FullName Then
		'開いたブックが登録されているマクロブックだった場合に読み込む｡TODO
		Dim registerdMacroBookNameList As New Collection
		bookName = Wb.Name

		If False Then
			Application.run "updateModules", bookName
			Application.run "ExeStringPro", "init", bookName
		End If
	End If
End Sub '}}}

Private Sub appevent_WorkbookBeforeClose(Byval wb As Workbook, Cancel As Boolean) '{{{
	If Workbooks.count <= 2 Then
		On Error Resume Next
		Workbooks("register.xlsx").Close savechanges:=False
	End If
End Sub '}}}

Private Sub appevent_SheetActivate(ByVal Sh As Object) '{{{
	Debug.Print "App.clsで監視中です｡SheetActivateイベントが発生しました｡"
	'Debug.print sh.Name
	Call SetKeyMapDic
End Sub '}}}

Private Sub appevent_WorkbookActivate(ByVal Wb As Workbook) '{{{
	'Debug.Print "App.clsで監視中です｡WorkbookActivateイベントが発生しました｡"
	Call SetKeyMapDic
End Sub '}}}

'------------------sub Function/Sub-----------------------------
Private Sub UpdateMru(Wb As Workbook)'{{{
	'ファイルを開いた履歴を書き込む
	openTime = Now
	newRegister = True
	separator = ":::"
	Dim buflines As New Collection

	'If mru_file does not exist, make mru_file.
	Dim mruDir As String: mruDir = ThisWorkbook.Path & "\.cache\"
	If Dir(mruDir & "mru.txt") = "" Then
		If Dir(mruDir, vbDirectory) = "" Then
			MkDir mruDir
		End If
		CreateObject("Scripting.FileSystemObject").CreateTextFile ThisWorkbook.Path & "\.cache\" & "mru.txt"
	End If

	Open ThisWorkbook.Path & "\.cache\" & "mru.txt" For Input As #1
	Do Until EOF(1)
		Line Input #1, buf
		If Split(buf, separator)(0) = Wb.FullName Then
			newRegister = False
			count = Val(Split(buf, separator)(1))
		Else
			buflines.Add Item:=buf
		End If
	Loop
	Close #1
	buflines.Add Item:=Wb.FullName & separator & count + 1 & separator & openTime

	Open ThisWorkbook.Path & "\.cache\mru.txt" For Output As #1
	For Each record In buflines
		Print #1, record
	Next record
	Close #1

	'TODO 利用したいpythonを指定出来るようにする。
	PYTHONPATH = "python"
	On Error Resume Next
		Call Shell(PYTHONPATH & " " & ThisWorkbook.Path & "\src\lib\sort.pyw")
	On Error GoTo 0

End Sub'}}}

Private Function registerdMacroBookNameList() As Collection
End Function
