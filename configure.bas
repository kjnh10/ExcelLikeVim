Attribute VB_Name = "configure"

Public myobject As ApplicationEvent

Public Sub SetAppEvent() '{{{
  If myobject is Nothing Then
    Set myobject = New ApplicationEvent
    Set myobject.appEvent = Application
  End If
  ' MsgBox "setiing AppEvent is done"
End Sub '}}}

Public Sub init() '{{{
  Call SetAppEvent
  Call keystrokeAsseser.init
  call vimize.main
  application.onkey "{F3}", "coreloader.reload"
End Sub '}}}

