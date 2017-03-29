Attribute VB_Name = "migemo"

Function migemize(query As String)
    Dim commandString As String
    commandString = "cmigemo -d ""C:\Users\bc0074854\Program\dict\migemo-dict"" -w " & query
    migemize = Replace(ExecCommand2(commandString), vbCrLf, "")
End Function

Public Function ExecCommand2(sCommand As String) As String
    ' 定数/変数宣言部
    Const TemporaryFolder = 2
    Dim oShell As Object, fso As Object, fdr As Object, ts As Object
    Dim sFileName As String
  
    ' オブジェクト変数に参照をセットします。
    Set oShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fdr = fso.GetSpecialFolder(TemporaryFolder)
      
    ' リダイレクト先のファイル名を生成します。
    Do: sFileName = fso.BuildPath(fdr.Path, fso.GetTempName)
    Loop While fso.FileExists(sFileName)
  
    ' コマンドを実行します。
    oShell.Run "%ComSpec% /c " & sCommand & ">" & sFileName & " 2<&1" _
               , 0, True
  
    ' 戻り値をセットします。
    If fso.FileExists(sFileName) Then
        Set ts = fso.OpenTextFile(sFileName)
        ExecCommand2 = ts.ReadAll
        ts.Close
        Kill sFileName
    End If
  
    ' オブジェクト変数の参照を解放します。
    Set ts = Nothing: Set fdr = Nothing
    Set fso = Nothing: Set oShell = Nothing
End Function

