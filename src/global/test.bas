Attribute VB_Name = "test"
Function Sample2() As DataObject
    Application.ScreenUpdating = False
    ActiveSheet.Rows(1).Copy
        Dim Dobj As DataObject
    Set Dobj = New DataObject
    With Dobj
        .GetFromClipboard    ''ïœêîÇÃÉfÅ[É^ÇDataObjectÇ…äiî[Ç∑ÇÈ
    End With
    Set Sample2 = Dobj
    Application.CutCopyMode = False
End Function
    
Sub sa()
    SetRegDic Sample2(), "*"
End Sub

Sub papin()
    RegDic.Item("*").PutInClipboard
    ActiveSheet.paste
End Sub


Sub checkfile()
	Dim b As Workbook
	For Each b in Workbooks
		Debug.Print b.Name
	Next b
End Sub
