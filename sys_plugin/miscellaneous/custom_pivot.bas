Attribute VB_Name = "custom_pivot"

Public sub setdefaultprop()
  Dim PT As PivotTable
  Set PT = ActiveCell.PivotTable
  PT.HasAutoFormat = False
  For Each f In PT.PivotFields
    f.ShowAllItems = True
  Next f
End Sub

Public sub numberformat()
  Dim PT As PivotTable
  Set PT = ActiveCell.PivotTable
  For Each f In PT.DataFields
    f.NumberFormat = "#,##0,,"
  Next f
End Sub
