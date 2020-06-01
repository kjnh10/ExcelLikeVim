Attribute VB_Name = "configure"

Public myobject As ApplicationEvent

Public Sub SetAppEvent() '{{{
  If myobject is Nothing Then
    Set myobject = New ApplicationEvent
    Set myobject.appEvent = Application
  End If
  Debug.Print "setiing AppEvent is done"
End Sub '}}}

Public Sub init() '{{{
  Call SetAppEvent
  application.onkey "{F3}", "coreloader.reload"
  application.onkey "^P", "'ExeStringPro ""unite command""'"
End Sub '}}}

