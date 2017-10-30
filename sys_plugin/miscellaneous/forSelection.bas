Attribute VB_Name = "forSelection"
'
Sub dfs() 'doforeslection
  Dim c As Range
  For Each c In Selection
    Call dealing(c)
  Next c
End Sub

Private Sub dealing(ByRef c As Range)
  'Write down process for each cell
  c.Value = "'" & CStr(c.Value)
End Sub
