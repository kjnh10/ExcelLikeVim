VERSION 1.0 CLASS
BEGIN
	MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'----------------------------- Workbook_open() ---------------
Private Sub Workbook_Open() '{{{
	Call clearAllModules
	Call import_coreloader
	Application.Run "coreloader.main"
End Sub '}}}

Private Sub Workbook_Close() '{{{
End Sub '}}}

Private Sub clearAllModules(Optional targetBook As Workbook = ThisWorkbook) '{{{
	For Each component In targetBook.VBProject.VBComponents
		If component.Type = 1 Or component.Type = 2 Or component.Type = 3 Then 'Standard(Type=1) / Class(Type=2) / Form(Type=3)
			ThisWorkbook.VBProject.VBComponents.Remove component
		End If
	Next component
End Sub '}}}

Private Sub import_coreloader() '{{{
	Dim srcPath As String: srcPath = ThisWorkbook.Path & "/core/coreloader.bas"
	Dim moduleName As String: moduleName = "coreloader"

	If Not isMemberOfVBEComponets(ThisWorkbook, moduleName) Then
		ThisWorkbook.VBProject.VBComponents.Import srcPath
	Else
		With ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule
			.DeleteLines StartLine:=1, count:=.CountOfLines
			.AddFromFile srcPath
		End With
	End If
End Sub '}}}

'----------------------------- common Functions / Subs ---------------
Private Function isMemberOfVBEComponets(book As Workbook, moduleName As String) As Boolean '{{{
	'Argument: moduleName like CodeManager
	'Return: whether or not module is registered
	For Each Item In book.VBProject.VBComponents
		If Item.Name = moduleName Then
			isMemberOfVBEComponets = True
			Exit Function
		End If
	Next
	isMemberOfVBEComponets = False
End Function '}}}
