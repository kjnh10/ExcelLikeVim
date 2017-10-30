Attribute VB_Name = "UniteCaller"

Public UniteCandidatesList As Collection '
Public unite_source As String
Public unite_argument As String
Public isExistPython As Boolean

Public Sub unite(Optional sourceName As String = "") '{{{
  If sourceName = "" Then
    Msgbox "sourceを指定してください｡使用出来るsource一覧を表示出来るようにする予定｡"
  End If

  On Error GoTo Myerror
  Set UniteCandidatesList = ExeStringPro("GatherCandidates_" & sourceName) 'CandidateListの設定
  On Error GoTo 0
  unite_source = sourceName 'source名の設定

  'TODO soureとcandidateはここで持たすのではなく,formオブジェクトのインスタンス変数として持たせた方が､複数子ウィンドウを立ちあげられよい？
  UniteInterface.Show
  Exit Sub
Myerror:
  MsgBox "sourceNameが不正です｡" & Err.Description
End Sub '}}}

'mru
Function GatherCandidates_mru() As Collection '{{{
  Dim result As New Collection
  Dim reverseResult As New Collection
  Set FSO = CreateObject("Scripting.FileSystemObject")

  Open Udir & ".cache\mru.txt" For Input As #1
  Do Until EOF(1)
    Line Input #1, buf
    FileName = Split(buf, ":::")(0)
    ' If fso.FileExists(filename) Then 時間がかかりすぎるため｡
    If True Then
      result.Add buf
    End If
  Loop
  Close #1

  'sort.pywが使えない場合｡mruファイルは最終行から読めば開かれた順になっているはず｡
  If Not isExistPython Then
    For i = result.Count to 1 Step -1
      reverseResult.Add result(i)
    Next
    Set GatherCandidates_mru = reverseResult
  Else
    Set GatherCandidates_mru = result
  End If
End Function '}}}
Function defaultAction_mru(arg) 'tableにした方がよいか'{{{
  For Each f in Split(arg, vbCrLf)
    SmartOpenBook(f)
  Next f
End Function'}}}

'sheet
Function GatherCandidates_sheet() As Collection '{{{
  Dim result As New Collection
  Dim sh As Worksheet
  Set Wb = ActiveWorkbook
  For Each sh In Wb.Worksheets
    result.Add sh.Name
  Next sh
  Set GatherCandidates_sheet = result
End Function '}}}
Function defaultAction_sheet(arg) 'tableにした方がよいか'{{{
  Worksheets(arg).Activate
End Function'}}}

'book
Function GatherCandidates_book() As Collection '{{{
  Dim result As New Collection
  Dim wb As Workbook

  For Each wb In Workbooks()
    result.Add wb.Name
  Next wb

  Set GatherCandidates_book = result
End Function '}}}
Function defaultAction_book(arg) 'tableにした方がよいか'{{{
  Workbooks(arg).Activate
End Function'}}}

'filter
Function GatherCandidates_filter() As Collection '{{{
  Dim ValueCollection As New Collection
  Set targetColumnRange = InterSect(GetFilterRange, Columns(ActiveCell.Column))
  Set targetColumnRange = targetColumnRange.SpecialCells(xlCellTypeVisible)

  On Error Resume Next
  For Each c in targetColumnRange
    If c.Value <> "" Then
      Debug.Print c.Value
      ValueCollection.Add c.Value, Cstr(c.Value)
    End If
  Next c
  On Error GoTo 0

  Set GatherCandidates_filter = ValueCollection
End Function '}}}
Function defaultAction_filter(SelectionMerged As String) 'tableにした方がよいか'{{{
  Application.ScreenUpdating = False
  ' If ActiveSheet.FilterMode Then
  ' ActiveSheet.ShowAllData
  ' End If
  GetFilterRange.AutoFilter field:= ActiveCell.Column - GetFilterRange.Column + 1, Criteria1:=Split(SelectionMerged, vbCrLf), Operator:=xlFilterValues
  Call gg()
  Call move_down()
End Function '}}}

'project
Function GatherCandidates_project() As Collection '{{{
  Dim ValueCollection As New Collection
  Set targetColumnRange = InterSect(GetFilterRange, Columns(GetFilterRange.Column))

  On Error Resume Next
  For Each c in targetColumnRange
    If c.Value <> "" Then
      ValueCollection.Add c.Value, Cstr(c.Value)
    End If
  Next c
  On Error GoTo 0

  Set GatherCandidates_project = ValueCollection
End Function '}}}
Function defaultAction_project(SelectionMerged As String) 'tableにした方がよいか'{{{
  Application.ScreenUpdating = False
  If ActiveSheet.FilterMode Then
    ActiveSheet.ShowAllData
  End If
  GetFilterRange.AutoFilter field:= GetFilterRange.Column, Criteria1:=Split(SelectionMerged, vbCrLf), Operator:=xlFilterValues
  Call gg()
  Call move_down()
End Function '}}}

