VERSION 1.0 CLASS
BEGIN
	MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'----------------------------- Consts ---------------
'ライブラリリストの設定 (設置フォルダはワークブックと同じディレクトリ)
Const FILENAME_LIBLIST As String = "libdef.txt" 'ライブラリリストのファイル名
Const REGISTRY_BOOK AS String =  "register.xlsx"

'----------------------------- Workbook_open() ---------------
Private Sub Workbook_Open() '{{{
	Call clearAllModules()
	Call import_xbundle()
	Application.run "UpdateModules", "VimX", 1
	Application.run ("InitializeApplication")
	Application.OnTime Now * TimeValue("00:00:05"), "SetAppevent"
End Sub '}}}

Private Sub Workbook_Close() '{{{
	'clearしてsaveした方がよいかも
	' Call clearAllModules
End Sub '}}}

Private Sub clearAllModules() '{{{
	For Each component In ThisWorkbook.VBProject.VBComponents
		'標準モジュール(Type=1) / クラスモジュール(Type=2) / form(Type=3)を全て削除
		If component.Type = 1 Or component.Type = 2 Or component.Type = 3 Then
			ThisWorkbook.VBProject.VBComponents.Remove component
		End If
	Next component
End Sub '}}}

Private Sub import_xbundle() '{{{
	Set myFSO = CreateObject("Scripting.FileSystemObject")
	modPath = absPath("./src/core/CodeManager.bas")
	Dim moduleName As String: moduleName = myFSO.GetBaseName(modPath)

	If Not isMemberOfVBEComponets(ThisWorkbook, moduleName) Then '存在しない場合は新規登録｡
		ThisWorkbook.VBProject.VBComponents.Import modPath
	Else
		With ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule
			.DeleteLines StartLine:=1, count:=.CountOfLines
			.AddFromFile modPath
		End With
	End If
End Sub '}}}

Private Sub OpenRegistryBook() '{{{
	SmartOpenBook ThisWorkbook.Path & "\.cache\" & REGISTRY_BOOK
	Application.Windows(REGISTRY_BOOK).Visible = False
End Sub '}}}

'----------------------------- common Functions / Subs ---------------
Private Function absPath(ByVal pathFile As String) As String '{{{
	' ファイルパスを絶対パスに変換

	Dim nameOS As String
	nameOS = Application.OperatingSystem

	'replace Win backslash(Chr(92))
	pathFile = Replace(pathFile, Chr(92), Application.PathSeparator)

	'replace Mac ":"Chr(58)
	pathFile = Replace(pathFile, ":", Application.PathSeparator)

	'replace Unix "/"Chr(47)
	pathFile = Replace(pathFile, "/", Application.PathSeparator)


	Select Case Left(pathFile, 1)

		'Case1. . で始まる場合(相対指定)
	Case ".":

		Select Case Left(pathFile, 2)

			' Case1-1. 相対指定 "../" 対応
		Case "..":
			'MsgBox "Case1-1: " & pathFile
			absPath = ThisWorkbook.Path & Application.PathSeparator & pathFile
			Exit Function

			' Case1-2. 相対指定 "./" 対応
		Case Else:
			'MsgBox "Case1-2: " & pathFile
			absPath = ThisWorkbook.Path & Mid(pathFile, 2, Len(pathFile) - 1)
			Exit Function

	End Select

	'Case2. 区切り文字で始まる場合 (絶対指定)
Case Application.PathSeparator:

	' Case2-1. Windows Network Drive ( chr(92) & chr(92) & "hoge")
	'MsgBox "Case2-1: " & pathFile
	If Left(pathFile, 2) = Chr(92) & Chr(92) Then
		absPath = pathFile
		Exit Function

	Else
		' Case2-2. Mac/UNIX Absolute path (/hoge)
		absPath = pathFile
		Exit Function

	End If

  End Select


  'Case3. [A-z][0-9]で始まる場合 (Mac版Officeで正規表現が使えれば select文に入れるべき...)

  ' Case3-1.ドライブレター対応("c:" & chr(92) が "c" & chr(92) & chr(92)になってしまうので書き戻す)
  If nameOS Like "Windows *" And Left(pathFile, 2) Like "[A-z]" & Application.PathSeparator Then
	  'MsgBox "Case3-1" & pathFile
	  absPath = Replace(pathFile, Application.PathSeparator, ":", 1, 1)
	  Exit Function
  End If

  ' Case3-2. 無指定 "filename"対応
  If Left(pathFile, 1) Like "[0-9]" Or Left(pathFile, 1) Like "[A-z]" Then
	  absPath = ThisWorkbook.Path & Application.PathSeparator & pathFile
	  Exit Function
  Else
	  MsgBox "Error[AbsPath]: fail to get absolute path."

  End If

End Function '}}}

Private Function checkExistFile(ByVal pathFile As String) As Boolean '{{{

	On Error GoTo Err_dir
	If dir(pathFile) = "" Then
		checkExistFile = False
	Else
		checkExistFile = True
	End If

	Exit Function

Err_dir:
	checkExistFile = False
End Function '}}}

Private Function list2array(ByVal pathFile As String) As Variant '{{{
	'リストファイルを配列で返す(行頭が'(コメント)の行 & 空行は無視する)
	Dim nameOS As String
	nameOS = Application.OperatingSystem

	'1. リストファイルの読み取り
	Dim fp As Integer
	fp = FreeFile
	Open pathFile For Input As #fp

	'2. リストの配列化
	Dim arrayOutput() As String
	Dim countLine As Integer
	countLine = 0
	ReDim Preserve arrayOutput(countLine) ' 配列0で返す場合があるため

	Do Until EOF(fp)
		'ライブラリリストを1行ずつ処理
		Dim strLine As String
		Line Input #fp, strLine
		isLf = InStr(strLine, vbLf)
		If nameOS Like "Windows *" And Not isLf = 0 Then
			'OSがWindows かつ リストに LFが含まれる場合 (ファイルがUNIX形式)
			'ファイル全体で1行に見えてしまう。
			Dim arrayLineLF As Variant
			strLine = Replace(strLine, vbTab, "") 'tab(インデント)を無視
			arrayLineLF = Split(strLine, vbLf)
			For i = 0 To UBound(arrayLineLF) - 1
				'行頭が '(コメント) ではない & 空行ではない場合
				' If Not left(arrayLineLF(i), 1) = "'" And Len(arrayLineLF(i)) > 0 Then
				If arrayLineLF(i) <> "" Then
					arrayLineLFS = Split(arrayLineLF(i), " ")
					If arrayLineLFS(0) = "bundle" Then
						'配列への追加
						countLine = countLine + 1
						ReDim Preserve arrayOutput(countLine)
						arrayOutput(countLine - 1) = arrayLineLFS(1)
					End If
				End If
			Next i
		Else
			'OSがWindows and ファイルがWindows形式 (変換不要)
			'OSがMacOS X and ファイルがUNIX形式 (変換不要)
			'OSがMacOS X and ファイルがWindows形式
			' vbCrがモジュールファイル名を発見できなくなる。
			strLine = Replace(strLine, vbCr, "")
			arraystrLine = Split(strLine, " ")
			'行頭が '(コメント) ではない & 空行ではない場合
			If Not Left(strLine, 1) = "'" And Len(strLine) > 0 Then
				If arraystrLine(0) = "bundle" Then
					'配列への追加
					countLine = countLine + 1
					ReDim Preserve arrayOutput(countLine)
					arrayOutput(countLine - 1) = arraystrLine(1)
				End If
			End If
		End If
	Loop


	'3. リストファイルを閉じる
	Close #fp

	'4. 戻り値を配列で返す
	list2array = arrayOutput
End Function '}}}

Private Function isMemberOfVBEComponets(book As Workbook, moduleName As String) As Boolean '{{{
	'argument: moduleName like CodeManager
	'return: whether or not module is registered
	For Each Item In book.VBProject.VBComponents
		If Item.Name = moduleName Then
			isMemberOfVBEComponets = True
			Exit Function
		End If
	Next
	isMemberOfVBEComponets = False
End Function '}}}
