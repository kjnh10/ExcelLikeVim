Attribute VB_Name = "vbundle"

'----------------------------- declare variables ----------
Enum Module'{{{
  Standard = 1
  Class = 2
  Forms = 3
  ActiveX = 11
  Document = 100
End Enum'}}}

'----------------------------- main -----------------------
Public Sub read_vimxrc()'{{{
	settingFilePath = Environ("homepath") & "\.vimxrc"

	Open settingFilePath For Input As #1
	Do Until EOF(1)
		Line Input #1, buf
		buf = Replace(buf,vbTab,"") 'tab(インデント)を無視

		If Left(buf,1) = "'" Then
			Goto NextLoop
		End If

		If buf <> "" Then
			instruction = Split(buf, " ")(0)
			argument = Mid(buf, Instr(Instr(buf, " ") + 1, buf, " ") + 1) '2つ目のスペース以降を取得
			' Application.Run instruction, argument
			Debug.Print "instruction:" & instruction & vbCrLf & "argument:" & argument
		End If

		NextLoop:
	Loop
	Close #1
End Sub'}}}

'----------------------------- updatemodules --------------
Public Function bundle()
End Function

Public Function UpdateModulesOfBook(Optional bookPath As String = "", Optional isCalledFromThisWorkbookModule = False) '{{{
	Const moduleListFile As String = "libdef.txt" 'ライブラリリストのファイル名
	Dim msgError As String: msgError = "Error Message"
	Dim updatedModuleNameList As New Collection

	'Get module list to update from libdef.'{{{
	'Set targetBook, targetBookModuleDirectory, libDefPath 
	Dim targetBook As Workbook
	if bookPath = "" Then
		Set targetBook = ThisWorkbook
		targetBookModuleDirectory = ThisWorkbook.Path
		libDefPath = targetBookModuleDirectory & "\" & moduleListFile
	Else
		Set targetBook = Application.Workbooks(bookPath)
		targetBookModuleDirectory = ThisWorkbook.Path & "\src\forbook\" &targetBook.Name
		libDefPath = targetBookModuleDirectory & "\" & moduleListFile
	End if

	Dim targetModuleList As Variant 'list of module path

	If Not checkExistFile(libDefPath) Then
		Msgbox "Error: ライブラリリスト" & libDefPath & "が存在しません。"
		Exit Function
	End If
	targetModuleList = list2array(libDefPath)
	If UBound(targetModuleList) = 0 Then
		Msgbox "Error: ライブラリリストに有効なモジュールの記述が存在しません。"
		Exit Function
	End If'}}}

	'Update modules'{{{
	Set myFSO = CreateObject("Scripting.FileSystemObject")
	For i = 0 To UBound(targetModuleList) - 1
		Dim modulePath As String: modulePath = targetModuleList(i)
		If isCalledFromThisWorkbookModule And myFSO.GetBaseName(absPath(targetModuleList(i))) = "ThisWorkbook" Then
			Debug.Print "Not update ThisWorkbook because it's dangerous updating it when it is called from ThisWorkbook moudule."
		Else
			Call updateSingleModule(targetBook, modulePath, msgError)
		End If
	Next i
	Set myFSO = Nothing '}}}

	If msgError = "Error Message" Then
		Msgbox "All Modules were successfully updated!"
	Else
		Msgbox msgError
	End If
End Function'}}}

Private Function updateSingleModule(targetBook As Workbook, modulePath As String, msgError As String)'{{{
	Set myFSO = CreateObject("Scripting.FileSystemObject")
	On Error GoTo except
		pathModule = absPath(modulePath)
		moduleName = myFSO.GetBaseName(pathModule)
		If Not isMemberOfVBEComponets(targetBook, moduleName) Then '存在しない場合は新規登録｡
			targetBook.VBProject.VBComponents.Import pathModule
		ElseIf moduleName <> "vbundle" And checkExistFile(pathModule) Then 'CodeManagerの書き換えは行わない。
			With targetBook.VBProject.VBComponents(moduleName).CodeModule
				Debug.Print "Started deleting " & moduleName
				'workbook,worksheetモジュールの場合 http://futurismo.biz/archives/2386
				.DeleteLines StartLine:=1, count:=.CountOfLines
				Debug.Print moduleName & "Deleted " & moduleName
				.AddFromFile pathModule
				Debug.Print moduleName & "読み込まれた"

				Select Case targetBook.VBProject.VBComponents(moduleName).type
					Case Module.Standard 'for .bas
						'何もしない｡
					Case Module.Class
						.DeleteLines StartLine:=1, count:=4
					Case Module.Forms
						.DeleteLines StartLine:=1, count:=10
					Case Module.Document
						.DeleteLines StartLine:=1, count:=4
					Case Else
						Debug.Print targetBook.VBProject.VBComponents(moduleName).type
					End Select
			End With
		End If
	except:
		If Err.Description <> "" Then
			msgError = msgError & vbCrLf & Err.Description & ": when updating " & moduleName
		End If
		Set myFSO = Nothing
End Function'}}}

Public Sub UpdateLibList(Optional targetBookModuleDirectory As String = "", Optional registerPattern As String = ".*\.cls$|.*\.bas$|.*\.frm$")'{{{
	'----------------This Function Update Libdef file  inaccordance with current directory structure.-------------------------------
	'TODO これはvbaからでなくともvim上で実行したい｡
	If targetBookModuleDirectory = "" Then targetBookModuleDirectory = ThisWorkbook.path & "\src\forbook\" & ActiveWorkbook.Name
	Dim fu As New FileUtil
	Dim file As Variant
	' 結果格納用変数
	Dim result As Collection
	' your_path配下でsearch_patternに合致した情報を取得
	Set result = fu.getFileListRecursive(targetBookModuleDirectory, registerPattern).Files ' 第二引数を省略した場合は全取得
	' ファイル一覧がフルパスで表示される

	Open targetBookModuleDirectory & "\" & "libdef.txt" For output As #1
	Print #1, "' vim: filetype=vb"
	For Each file In result
		'TODO 相対パスへ変換
		Print #1, Replace(file,"\","/")
	Next
	Close #1
End Sub'}}}'}}}
'----------------------------- export modules --------------
Public Sub EM(Optional bookPath As String = "")'{{{
	'''targetbookのコードを外部へ保存する。'''

	'targetbook,targetbookdirectoryの設定
	If bookPath = "vimx" Then
		Set targetBook = ThisWorkbook
		targetBookModuleDirectory = ThisWorkbook.path
	ElseIf bookPath = "" Then
		Set targetBook = ActiveWorkbook
		targetBookModuleDirectory = ThisWorkbook.path & "\src\forbook\" & targetBook.Name
	Else
		Set targetBook = Application.Workbooks(Dir(bookPath))
		targetBookModuleDirectory = ThisWorkbook.path & "\src\forbook\" & targetBook.Name
	End If

	'targetBookModuleDirectoryが存在しなければ作る｡
	isNewRegistration = False
	If Dir(targetBookModuleDirectory, vbDirectory) = "" Then
		'保存先ディレクトリの作成
		MkDir targetBookModuleDirectory
		isNewRegistration = True
	End If

	'moduleのエクスポート
	For Each vb_component In targetBook.VBProject.VBComponents
		pathToExport = "" '初期化
		If Not vb_component.Name = "CodeManager" Then 'CodeManagerは自身なのでexportを行わない｡importでなければ大丈夫？
			'libdefを参照して置き場所をpathToExportに設定。
			If isNewRegistration = True Then
				pathToExport = targetBookModuleDirectory & "\" & vb_component.Name & getExtention(vb_component)
			Else
				Open targetBookModuleDirectory & "\" & "libdef.txt" For Input As #1
				Do Until EOF(1)
					Line Input #1, buf
					buf = Replace(buf,vbTab,"") 'tab(インデント)を無視
					If Left(buf, 1) = "'" Then 'コメント行を無視
						Exit Do
					End If
					buf = Split(buf, " ")(1)
					If InStr(buf,vb_component.Name) Then 'TODO 正規表現にする
						pathToExport = buf
						Exit Do
					End If
				Loop
				Close #1
			End If

			If pathToExport <> "" Then
				vb_component.Export pathToExport
			End IF
		End If
	Next

	'libdefの更新
	If isNewRegistration Then
		Call UpdateLibList(targetBookModuleDirectory)
	End If
End Sub'}}}

'----------------------------- common Functions / Subs --------------
Private Function isExcelObject(fileName As String) As Boolean'{{{
	Set RE = CreateObject("VBScript.RegExp")
	RE.IgnoreCase = True
	RE.pattern = ".cls$|.frm|ThisWorkbook|Sheet"
	If RE.test(fileName) Then
		isExcelObject = True
	Else
		isExcelObject = False
	End If
End Function'}}}

	Private Function checkRemainigComponents() As Boolean'{{{
	  '標準モジュール/クラスモジュールの合計数が0であればOK
	  Dim cntBAS As Long
	  cntBAS = countBAS()
	
	  Dim cntClass As Long
	  cntClass = countClasses()
	
	  'CodeManagerのみが残っている。
	  If cntBAS <= 1 And cntClass = 0 Then
		  checkRemainigComponents = True
	  Else
		  checkRemainigComponents = False
	  End If
	End Function'}}}

Private Function countBAS() As Long'{{{
  Dim count As Long
  count = countComponents(1) 'Type 1: bas
  countBAS = count
End Function'}}}

Private Function countClasses() As Long'{{{
  Dim count As Long
  count = countComponents(2) 'Type 2: class
  countClasses = count
End Function'}}}

Private Function countComponents(ByVal numType As Integer) As Long'{{{
  '存在する標準モジュール/クラスモジュールの数を数える
  
  Dim i As Long
  Dim count As Long
  count = 0
  
  With targetBook.VBProject
    For i = 1 To .VBComponents.count
      If .VBComponents(i).Type = numType Then
        count = count + 1
      End If
    Next i
  End With

  countComponents = count
End Function'}}}

Private Function checkExistFile(ByVal pathFile As String) As Boolean'{{{
  On Error GoTo Err_dir
  If Dir(pathFile) = "" Then
    checkExistFile = False
  Else
    checkExistFile = True
  End If

  Exit Function

Err_dir:
  checkExistFile = False
End Function'}}}

Private Function list2array(ByVal pathFile As String) As Variant'{{{
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
			strLine = Replace(strLine,vbTab,"") 'tab(インデント)を無視
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
			strLine = Replace(strLine, vbCr, "") ' vbCrがモジュールファイル名を発見できなくなる。
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
End Function'}}}

Private Function getExtention(myComponent) As String'{{{
	Dim extention As String
	Select Case myComponent.Type
		Case Module.Standard
			extention = ".bas"
		Case Module.Class
			extention = ".cls"
		Case Module.Forms
			extention = ".frm"
		Case Module.ActiveX
			extention = ".cls"
		Case Module.Document
			extention = ".cls"
	End Select

	getExtention = extention
End Function'}}}

Public Function absPath(ByVal pathFile As String) As String'{{{
	'------------ ファイルパスを絶対パスに変換 -----------------------
	'省略文字(. .. ~)の展開
	Select Case left(pathFile, 1)
	Case ".": 'Case1. . で始まる場合(相対指定)
		Select Case left(pathFile, 2) ' Case1-1. 相対指定 "../" 対応
		Case "..":
			absPath = ThisWorkbook.Path & Application.PathSeparator & pathFile
			Exit Function
		Case Else: ' Case1-2. 相対指定 "./" 対応
			absPath = ThisWorkbook.Path & Mid(pathFile, 2, Len(pathFile) - 1)
			Exit Function
		End Select
	Case Application.PathSeparator: 'Case2. 区切り文字で始まる場合 (絶対指定)
		If left(pathFile, 2) = Chr(92) & Chr(92) Then ' Case2-1. Windows Network Drive ( chr(92) & chr(92) & "hoge")
			absPath = pathFile
			Exit Function
		Else ' Case2-2. Mac/UNIX Absolute path (/hoge)
			absPath = pathFile
			Exit Function
		End If
	Case "~"
		pathfile = Replace(pathfile, "~", Environ("homepath"))
	End Select

	'区切り文字をOSに合わせて変換
	nameOS = Application.OperatingSystem
	pathFile = Replace(pathFile, Chr(92), Application.PathSeparator) 'replace Win backslash(Chr(92))
	pathFile = Replace(pathFile, ":", Application.PathSeparator) 'replace Mac ":"Chr(58)
	pathFile = Replace(pathFile, "/", Application.PathSeparator) 'replace Unix "/"Chr(47)

	' 'Case3. [A-z][0-9]で始まる場合 (Mac版Officeで正規表現が使えれば select文に入れるべき...)
	' ' Case3-1.ドライブレター対応("c:" & chr(92) が "c" & chr(92) & chr(92)になってしまうので書き戻す)
	If nameOS Like "Windows *" And left(pathFile, 2) Like "[A-z]" & Application.PathSeparator Then
		'MsgBox "Case3-1" & pathFile
		absPath = Replace(pathFile, Application.PathSeparator, ":", 1, 1)
		Exit Function
	End If
	' Case3-2. 無指定 "filename"対応
	If left(pathFile, 1) Like "[0-9]" Or left(pathFile, 1) Like "[A-z]" Then
		absPath = ThisWorkbook.Path & Application.PathSeparator & pathFile
		Exit Function
	Else
		MsgBox "Error[AbsPath]: fail to get absolute path."
	End If
End Function'}}}

Private Function isMemberOfCollection(col As Collection, query) As Boolean'{{{
	For Each item In col
		If item = query Then
			isMemberOfCollection = True
			Exit Function
		End If
	Next
	isMemberOfCollection = False
End Function'}}}

Private Function isMemberOfVBEComponets(book As Workbook, query) As Boolean'{{{
	For Each item In book.VBProject.VBComponents
		If item.Name = query Then
			isMemberOfVBEComponets = True
			Exit Function
		End If
	Next
	isMemberOfVBEComponets = False
End Function'}}}
