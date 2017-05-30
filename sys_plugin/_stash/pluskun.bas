Attribute VB_Name = "pluskun"

' Dim fu As New FileUtil
' fu.getFileListRecursive(path)

Sub aaa() '{{{

End Sub'}}}

Sub eee() '{{{
	' If fullpath <> "" Then
	' 	SmartOpenBook(fullpath)
	' End If
	' 'Headerの修正

	'本体部分の修正
	For Each partName in GetAllParts()
		InitializePart(partName)
		ModifyPart(partName)
		ModifyFileName(partName)
	Next partName

End Sub'}}}

'---------------------------------------------------------------
Sub InitializePart(partName As String)'{{{
	'どの部品にも共通した処理を書く｡
	Dim contents As Range: Set contents = GetContentsOfPart(partName)
	' contents.Offset(0,1).ClearContents '備考1
	contents.Offset(0,2).ClearContents '備考2

	'編集形態
	contents.Offset(0,4).Value = "完全流用" '一旦全て完全流用に
	On Error Resume Next
	contents.Offset(0,4).FormatConditions.Add(xlCellValue, xlEqual, "完全流用").Interior.ColorIndex = 16 '完全流用なら網掛け M10はエラー
	On Error GoTo 0

	contents.Offset(0,5).ClearContents '流用元

	contents.Offset(0,7).NumberFormatLocal = "G/標準"
	contents.Offset(0,7).Value = "=LOOKUP(L" & contents(0).row + 1 &",{""完全流用"",""新規"",""流用改訂"";""流用指示"",""ネイティブ＋赤字あり"",""PDF/X1-a""})" '入稿形態

	contents.Offset(0,8).NumberFormatLocal = "G/標準"
	contents.Offset(0,8).Value = "=LOOKUP(O" & contents(0).row + 1 &",{""PDF/X1-a"",""ネイティブ+赤字あり"",""流用指示"";""WF1"",""WF2"",""WF1""})" '入稿形態
End Sub'}}}

Sub ModifyPart(partName As String)'{{{
	'どの部品にも共通した処理を書く｡
	If partName Like "*添削*" Then
		Set contents = GetContentsOfPart(partName)
		For Each c in contents.Offset(0,1)
			If c.Value like "ウラ" Then
				c.Offset(0,3).Value = "流用改訂"
			Else
				c.Offset(0,3).Value = "新規"
			End If
		Next c

	ElseIf partName = "本冊" Then
		Set contents = GetContentsOfPart(partName)
		For Each c in contents
			Select Case c
				Case "表Ⅰ","表Ⅳ" '新規に変更する行
					c.Offset(0,4).Value = "新規"
				Case "表Ⅱ","表Ⅲ","目次","告知","添削課題トビラ","添削課題活用法","今月のヒント" '流用改訂に変更する行
					c.Offset(0,4).Value = "流用改訂"
			End Select
		Next c

	ElseIf partName = "見返し" Then
		Set contents = GetContentsOfPart(partName)
		' contents(2).Offset(0, 5) = "前年度から色カエ"
		' contents(3).Offset(0, 5) = "前年度から色カエ"
		' contents(4).Offset(0, 4) = "流用改訂"
	End If

End Sub'}}}

Sub ModifyFileName(partName As String)'{{{
	Dim contents As Range: Set contents = GetContentsOfPart(partName)
	'新規､流用ならファイル名を14→15にする｡
	For Each c in contents.Offset(0, 6)
		If c.Offset(0, -2).Value <> "完全流用" Then
			If Left(c.Value, 3) = "011" Then
				c.Value =  Left(c.Value, 3) & "15" & Mid(c.Value, 6)
			End If
		Else
			'TODO 前年度のデータから持ってくる｡
		End If
	Next c
End Sub'}}}

'---------------------------------------------------------------
Function GetContentsOfPart(partName As String) As Range'{{{
'引数：部品名
'返り値：内容列一覧
On Error GoTo ErrorHandling
	Set searchRange = Range(Cells(12, 3), Cells(ActiveSheet.UsedRange.Rows.Count, 3))
	For Each c in searchRange
		If c.Value = partName Then
			Set partNameCell = c
			Exit For
		End If
	Next c

	Do Until i > 100
		Set startCell = partNameCell.Offset(i, 0)
		If startCell.Value = "台" Then
			Set startCell = startCell.Offset(1, 5) 'Offsetの場合は結合セル分は足さない
			Exit Do
		End If
		i = i + 1
	Loop

	Set GetContentsOfPart = Range(startCell, startCell.End(xlDown))
	Exit Function
ErrorHandling:
	Set GetContentsOfPart = Nothing
End Function'}}}

Function GetColumnOfProperty(propertyName As String) As Long'{{{
'引数：属性名
'返り値：
On Error GoTo ErrorHandling
	Set searchRange = Range(Cells(12, 3), Cells(ActiveSheet.UsedRange.Rows.Count, 3))
	For Each c in searchRange
		If c.Value = "台" Then
			FieldRowNo = c.Row + 2
			Exit For
		End If
	Next c

	Set searchRange = Cells(FieldRowNo, 3).Resize(1, 23)
	For Each c in searchRange
		If c.MergeArea(1, 1).Value = propertyName Then
			GetColumnOfProperty = c.Column
			Exit Function
		End If
	Next c
	GetColumnOfProperty = 0 '見つからない場合
ErrorHandling:
	GetColumnOfProperty = 0 'エラーの場合
End Function'}}}

Function GetAllParts() As Collection 'parts一覧の取得'{{{
'引数：
'返り値：部品名のcollecton
On Error GoTo ErrorHandling
	Dim result As New Collection
	Set searchRange = Range(Cells(12, 3), Cells(ActiveSheet.UsedRange.Rows.Count, 3))
	
	For Each c in searchRange
		If c.Value = "製本" Then
			result.Add c.Offset(-1, 0).Value
		End If
	Next c

	Set GetAllParts = result
	For Each a in result
		Debug.Print a
	Next a
ErrorHandling:

End Function'}}}

Function LastyearData() ''{{{

End Function'}}}

Function LastmonthData() ''{{{

End Function'}}}
