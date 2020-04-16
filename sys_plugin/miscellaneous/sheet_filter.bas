Attribute VB_Name = "sheet_filter"

Sub filter_sheet()
  ' TODO: parameteraize
  For Each ws In Worksheets
    If InStr(ws.Name, "2019") = 0 Then
      Worksheets(ws.Name).Visible = False
    End If
  Next ws
End Sub

Sub show_all_sheet()
  For Each ws In Worksheets
    Worksheets(ws.Name).Visible = True
  Next ws
End Sub
