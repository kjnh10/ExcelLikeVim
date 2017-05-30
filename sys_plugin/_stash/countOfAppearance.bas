

Sub kic()
	'countの取得
	'TODO dictionaryで返すようにする｡
	Dim dic As Object
	Set dic = CreateObject("scripting.dictionary")
	For Each c In Selection
		If dic.exists(c.Value) Then
			dic(c.Value) = dic(c.Value) + 1
		Else
			dic(c.Value) = 1
		End If
	Next c

	For Each c In Selection
		c.Offset(0, 1).Value = dic(c.Value)
	Next c
End Sub
