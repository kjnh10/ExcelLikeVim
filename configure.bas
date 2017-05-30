Attribute VB_Name = "configure"

Public myobject As ApplicationEvent

Public Sub SetAppEvent() '{{{
	If myobject is Nothing Then
		Set myobject = New ApplicationEvent
		Set myobject.appEvent = Application
		'Set myobject.pptEvent = New PowerPoint.Application
		'Set myobject.wrdEvent = New Word.Application
	End If
	' MsgBox "setiing AppEvent is done"
End Sub '}}}

