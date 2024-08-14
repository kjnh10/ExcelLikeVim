Attribute VB_Name = "EditOperation"
Sub InteriorColor(number) '{{{
  Selection.Interior.ColorIndex = number
End Sub '}}}

Sub FontColor(number) '{{{
  Debug.Print "FontColor"
  Selection.Font.ColorIndex = number
End Sub '}}}

Sub SetRuledLines() '{{{
  Selection.Borders.LineStyle = xlContinuous
End Sub '}}}

Sub UnsetRuledLines() '{{{
  Selection.Borders.LineStyle = xlLineStyleNone
End Sub '}}}

Sub merge() '{{{
  Selection.merge
End Sub '}}}

Sub unmerge() '{{{
  Selection.unmerge
End Sub '}}}

Sub ex_up() '{{{
  Application.ScreenUpdating = False
  cur_row = ActiveCell.Row
  Rows(cur_row).Copy
  'select target_row
  Dim i As Long
  i = 1
  Do Until ActiveCell.Offset(-i, 0).EntireRow.Hidden = False
    i = i + 1
  Loop
  target_row = ActiveCell.Offset(-i, 0).Row
  target_column = ActiveCell.Offset(-i, 0).Column

  Rows(target_row).Select
  Selection.Insert

  ' remove the cell before moving
  Rows(cur_row + 1).Delete
  cells(target_row, target_column).Select
End Sub '}}}

Sub ex_below() '{{{
  Application.ScreenUpdating = False
  cur_row = ActiveCell.Row
  Rows(cur_row).Copy
  'target_row�̑I��
  Dim i As Long
  i = 1
  Do Until ActiveCell.Offset(i, 0).EntireRow.Hidden = False
    i = i + 1
  Loop
  target_row = ActiveCell.Offset(i, 0).Row
  target_column = ActiveCell.Offset(i, 0).Column

  Rows(target_row + 1).Select
  Selection.Insert
  Rows(cur_row).Delete

  cells(target_row, target_column).Select
End Sub '}}}

Sub ex_right() '{{{
  Application.ScreenUpdating = False
  cur_col = ActiveCell.Column
  Columns(cur_col).Copy
  'select target_row
  Dim i As Long
  i = 1
  Do Until ActiveCell.Offset(0, i).EntireColumn.Hidden = False
    i = i + 1
  Loop
  target_row = ActiveCell.Offset(0, i).Row
  target_column = ActiveCell.Offset(0, i).Column

  Columns(target_column + 1).Select
  Selection.Insert
  Columns(cur_col).Delete

  cells(target_row, target_column).Select
End Sub '}}}

Sub ex_left() '{{{
  Application.ScreenUpdating = False
  cur_col = ActiveCell.Column
  Columns(cur_col).Copy
  'select target_row
  Dim i As Long
  i = 1
  Do Until ActiveCell.Offset(0, -i).EntireColumn.Hidden = False
    i = i + 1
  Loop
  target_row = ActiveCell.Offset(0, -i).Row
  target_column = ActiveCell.Offset(0, -i).Column

  Columns(target_column).Select
  Selection.Insert
  Columns(cur_col + 1).Delete

  cells(target_row, target_column).Select
End Sub '}}}

Sub ZoomInWindow() '{{{
  ActiveWindow.Zoom = ActiveWindow.Zoom + 5
End Sub '}}}

Sub ZoomOutWindow() '{{{
  ActiveWindow.Zoom = ActiveWindow.Zoom - 5
End Sub '}}}

Sub MouseNormal() '{{{
  Application.Cursor = xlDefault
End Sub '}}}

Sub SetSeqNumber(Optional destRange As Range = Nothing) '{{{
  Application.ScreenUpdating = False
  If destRange Is Nothing Then
    Set destRange = Selection
  End If
  Set destRange = destRange.SpecialCells(xlCellTypeVisible)
  n = 1
  For Each r In destRange
    r.value = n
    Selection.NumberFormatLocal = "0_);[��](0)"
    n = n + 1
  Next
End Sub '}}}

Sub SortCurrentColumn() '{{{
  Application.ScreenUpdating = False
  Set targetRange = Selection.CurrentRegion

  With ActiveSheet.Sort
    With .SortFields
      .Clear
      .Add _
      Key:=Columns(ActiveCell.Column), _
      SortOn:=xlSortOnValues, _
      Order:=xlAscending, _
      DataOption:=xlSortNormal
    End With
    .SetRange targetRange
    .Header = xlYes 'check header exists. xlGuess rely on Excel.
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
  End With
End Sub '}}}

'--------sheet_move-------------------
Sub ActivateLeftSheet() '{{{
  sendkeys "^{PGUP}"
End Sub '}}}

Sub ActivateRightSheet() '{{{
  sendkeys "^{PGDN}"
End Sub '}}}

Sub ActivateFirstSheet(Optional where As String) '{{{
  With ActiveWorkbook
    .WorkSheets(1).Activate
  End With
End Sub '}}}

Sub ActivateLastSheet(Optional where As String) '{{{
  With ActiveWorkbook
    .WorkSheets(.WorkSheets.Count).Activate
  End With
End Sub '}}}

'---------auto_filter-----------------
Sub focusFromScratch() '{{{
  Application.ScreenUpdating = False
  If ActiveSheet.FilterMode Then
    ActiveSheet.ShowAllData
  End If
  GetFilterRange.AutoFilter ActiveCell.Column - GetFilterRange.Column + 1, ActiveCell.Value
End Sub '}}}

Sub focus() '{{{
  Application.ScreenUpdating = False
  GetFilterRange.AutoFilter ActiveCell.Column - GetFilterRange.Column + 1, ActiveCell.Value
End Sub '}}}

Sub exclude()'{{{
  Application.ScreenUpdating = False
  Dim filterCondition As Variant
  Dim buf As String

  buf = cells(ActiveCell.Row ,ActiveCell.Column).value

  Debug.Print Cells(Rows.Count, ActiveCell.Column).End(xlUp).Row
  Set targetColumnRange = InterSect(GetFilterRange, Columns(ActiveCell.Column))
  Set targetColumnRange = targetColumnRange.SpecialCells(xlCellTypeVisible)

  Set showedValueCollection = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  For Each c in targetColumnRange
    If c.Value <> buf Then
      showedValueCollection.Add "_" & c.Value, c.Value
    End If
  Next c
  On Error GoTo 0

  filterCondition = showedValueCollection.Keys

  For e = 0 to Ubound(filterCondition)
    filterCondition(e) = Mid(filterCondition(e),2)
  Next e

  GetFilterRange.AutoFilter field:= ActiveCell.Column - GetFilterRange.Column + 1, Criteria1:=filterCondition, Operator:=xlFilterValues
End Sub'}}}

Sub filterOff() '{{{
  Application.ScreenUpdating = False
  GetFilterRange.AutoFilter ActiveCell.Column
End Sub '}}}

Function GetFilterRange() As Range'{{{
  On Error GoTo error
  Set GetFilterRange = ActiveSheet.AutoFilter.Range
  Exit Function
error:
  Set GetFilterRange = ActiveSheet.UsedRange
End Function'}}}

Function smallerFonts() '{{{
  Dim currentFontSize As Long
  On Error GoTo ERROR01
  currentFontSize = Selection.Font.Size
  Selection.Font.Size = currentFontSize - 1
  period_buff = ">"
ERROR01:
End Function '}}}

Function biggerFonts() '{{{
  Dim currentFontSize As Long
  On Error GoTo ERROR01
  currentFontSize = Selection.Font.Size
  Selection.Font.Size = currentFontSize + 1
  period_buff = "<"
ERROR01:
End Function '}}}

Sub sp(Optional clearFilterdRowValue = 0) '{{{ smartpaste
  'TODO: erase data source

  Application.ScreenUpdating = False

  'need Microsoft Forms 2.0 Object Library reference
  Dim V As Variant    'whole data on clipboard
  Dim A As Variant    'one line


  Set destRange = Range(ActiveCell, cells(Rows.count, ActiveCell.Column)) 'ActiveCell to last row
  Set destRange = destRange.SpecialCells(xlCellTypeVisible)   'get visible cells only

  ''{{{
  Dim Dobj As DataObject
  Set Dobj = New DataObject
  With Dobj
    .GetFromClipboard
    On Error Resume Next
    V = .GetText
    On Error GoTo 0
  End With'}}}

  If Not IsEmpty(V) Then
    V = Split(CStr(V), vbCrLf)

    'erase lines hided by filter '{{{
    If clearFilterdRowValue = 1 Then
      referencRangeHeight = UBound(V) + 1
      referencRangeWidth = UBound(Split(CStr(V(0)), vbTab)) + 1
      Debug.Print referencRangeHeight
      Debug.Print referencRangeWidth
      For Each c in ActiveCell.Resize(referencRangeHeight, referencRangeWidth)
        c.Value = ""
      Next c
    End If'}}}

    'TODO: delete source data
    If Application.CutCopyMode = xlCut Then
      Set srcRange = GetCopiedRange(ActiveSheet.Name)
      For Each c in srcRange
        c.Value = ""
      Next c

      Application.CutCopyMode = False
    End If

    '{{{
    Dim i As Integer: i = 0
    Dim r As Range
    For Each r In destRange
      A = Split(CStr(V(i)), vbTab) 'i line
      For j = 0 to Ubound(A)
        If Cstr(Val(A(j))) = A(j) Then
          r.Offset(0, j).Value = Val(A(j))
        Else
          r.Offset(0, j).Value = A(j)
        End If
      Next j
      If Ubound(A) = -1 Then
        r.Offset(0, j).Value = ""
      End If

      i = i + 1
      If i >= UBound(V) Then
        Exit For
      End If
    Next'}}}
  End If

  Set Dobj = Nothing
  Set r = Nothing
End Sub '}}}

Sub sp2() '{{{ smartpaste
  Set srcRange = GetCopiedRange(ActiveSheet.Name)
  For Each r in srcRange.Rows
    Debug.Print r.row
  Next r
End Sub '}}}

Sub num_format_million()
  Selection.NumberFormatLocal = "#,##0,,"
End Sub

Sub topleft()
  Dim backsh As worksheet
  Set backsh = ActiveSheet
  For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    Range("A1").Select
  Next ws
  backsh.Activate
End Sub

'---------diff-----------------
Sub diffsh(targetsh As Worksheet, fromsh As Worksheet)'{{{
  'TODO prompt
  For Each c in fromsh.UsedRange
    If c.Value <> targetsh.Cells(c.Row, c.Column).Value Then
      targetsh.Cells(c.Row, c.Column).Interior.ColorIndex = 29
    End If
  Next c
End Sub'}}}

Sub diffRange(targetRange As Range, fromRange As Range)'{{{
  'TODO
End Sub'}}}

'-----------Supplimental functions------------------------
Function Field_No(fieldName As String, Optional sheetName As String = "", Optional fieldRowNum As Long = 1)'{{{
  If sheetName = "" Then
    set sheet = ActiveSheet
  Else
    set sheet = Worksheets(sheetName)
  End If

  Field_No = sheet.Range(Cells(fieldRowNum,1),Cells(fieldRowNum,50)).Find(What:=fieldName, LookIn:=xlFormulas, LookAt _
  :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
  False, MatchByte:=False, SearchFormat:=False).Column
End Function'}}}

Function GroupNo(groupName as String)'{{{
  GroupNo = ActiveSheet.Columns("A:A").Find(What:=groupName, LookIn:=xlFormulas, LookAt _
  :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
  False, MatchByte:=False, SearchFormat:=False).Row
End Function'}}}

Function AlphabetColumn(num As Long)'{{{
  buf = Cells(1, num).Address(True, False)
  AlphabetColumn = Left(buf, InStr(buf, "$") - 1)
End Function'}}}
