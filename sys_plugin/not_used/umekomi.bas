Sub mawasitene()
	Application.ScreenUpdating = False
	Dim rowsizes As New Collection
	Dim columnsizes As New Collection

	ActiveSheet.Shapes.SelectAll
	Selection.Placement = xlMove

	Set buf = ActiveSheet.UsedRange
	'行幅の保存'{{{
	r = buf.Rows.Count
	c = buf.Columns.Count
	For i = 1 To r
		rowsizes.Add Rows(i).RowHeight
	Next i

	For j = 1 To c
		columnsizes.Add Columns(j).ColumnWidth
	Next j '}}}

	Rows("1:" & r).RowHeight = 200
	Range(Columns(1), Columns(c)).ColumnWidth = 200

	ActiveSheet.Shapes.SelectAll
	Selection.Placement = xlMoveAndSize

	'行幅の回復
	For i = 1 To r
		Rows(i).RowHeight = rowsizes(i)
	Next i

	For j = 1 To c
		columns(j).ColumnWidth = columnsizes(j)
	Next j '}}}
End Sub
