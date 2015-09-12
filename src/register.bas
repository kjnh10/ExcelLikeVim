Attribute VB_Name = "register"

Public RegDic As Scripting.Dictionary

Public Sub SetRegDic(ByVal Content As DataObject, Optional registerName As String = "*")
    If RegDic Is Nothing Then
        Set RegDic = CreateObject("Scripting.Dictionary")
    End If
    Set RegDic.Item(registerName) = Content
End Sub

