
Attribute VB_Name = "form_resizer"

Public Const GWL_STYLE = -16
Public Const WS_CAPTION = &HC00000
Public Const WS_THICKFRAME = &H40000

#If VBA7 Then
    Public Declare PtrSafe Function GetWindowLong _
        Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function SetWindowLong _
        Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
    Public Declare PtrSafe Function DrawMenuBar _
        Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare PtrSafe Function FindWindowA _
        Lib "user32" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
#Else
    Public Declare Function GetWindowLong _
        Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function SetWindowLong _
        Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
    Public Declare Function DrawMenuBar _
        Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function FindWindowA _
        Lib "user32" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
#End If

Sub ResizeWindowSettings(frm As Object, show As Boolean)

Dim windowStyle As Long
Dim windowHandle As Long

'Get the references to window and style position within the Windows memory
windowHandle = FindWindowA(vbNullString, frm.Caption)
windowStyle = GetWindowLong(windowHandle, GWL_STYLE)

'Determine the style to apply based
If show = False Then
    windowStyle = windowStyle And (Not WS_THICKFRAME)
Else
    windowStyle = windowStyle + (WS_THICKFRAME)
End If

'Apply the new style
SetWindowLong windowHandle, GWL_STYLE, windowStyle

'Recreate the UserForm window with the new style 
DrawMenuBar windowHandle

End Sub