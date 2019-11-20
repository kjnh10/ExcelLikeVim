Attribute VB_Name = "UniteCaller"

Public UniteCandidatesList As Collection '
Public unite_source As String
Public unite_argument As String
Public isExistPython As Boolean

Public Sub unite(Optional sourceName As String = "") '{{{
  If sourceName = "" Then
    Msgbox "no source name specified"
  End If

  On Error GoTo Myerror
  Set UniteCandidatesList = ExeStringPro("GatherCandidates_" & sourceName)
  On Error GoTo 0
  unite_source = sourceName

  UniteInterface.Show
  Exit Sub
Myerror:
  MsgBox "sourceName is invalid" & Err.Description
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
    If True Then
      result.Add buf
    End If
  Loop
  Close #1

  If Not isExistPython Then
    For i = result.Count to 1 Step -1
      reverseResult.Add result(i)
    Next
    Set GatherCandidates_mru = reverseResult
  Else
    Set GatherCandidates_mru = result
  End If
End Function '}}}
Function defaultAction_mru(arg) 'table is better '{{{
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
Function defaultAction_sheet(arg) 'table is better '{{{
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
Function defaultAction_book(arg) 'table is better '{{{
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
Function defaultAction_filter(SelectionMerged As String) 'table is better '{{{
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
Function defaultAction_project(SelectionMerged As String) 'table is better? '{{{
  Application.ScreenUpdating = False
  If ActiveSheet.FilterMode Then
    ActiveSheet.ShowAllData
  End If
  GetFilterRange.AutoFilter field:= GetFilterRange.Column, Criteria1:=Split(SelectionMerged, vbCrLf), Operator:=xlFilterValues
  Call gg()
  Call move_down()
End Function '}}}

