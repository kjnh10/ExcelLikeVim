Attribute VB_Name = "fold"

Sub Expand_All()
    ActiveSheet.Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
End Sub

Sub Collapse_All()
    ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
End Sub