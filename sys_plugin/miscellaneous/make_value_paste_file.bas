
Attribute VB_Name = "make_value_paste_file"

Sub make_value_paste_file_for_activeworkbook()
  make_value_paste_file ActiveWorkbook
End Sub

Sub make_value_paste_file(Optional wb As Workbook)
  Dim org As Long
  org = Application.Calculation
  Application.Calculation = xlCalculationManual
  For Each ws In wb.Worksheets
    ws.UsedRange.Value = ws.UsedRange.Value
  Next ws
  Application.Calculation = org
  wb.SaveAs fileName:=wb.Path & "\value.xlsx", FileFormat:=xlOpenXMLWorkbook
End Sub
