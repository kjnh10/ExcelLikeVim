Attribute VB_Name = "ExecuteProcedure"

Function ExeStringPro(commandString As String) '{{{
  ' execute function specified by commandString
  Dim commandArray() As String
  commandString = Replace(commandString, "::", "!")
  commandArray = Split(commandString, " ")

  If commandArray(0) = "-a" Then  'execute as is
    commandString = Mid(commandString, Instr(commandString, " ") + 1) 'get after #2 space
    ExecuteAsIs commandString
    Exit Function
  End If

  ' Application.Run ("'KiriMacro_20200107-64bit.xlsm'!jumpmore")
  Set buf = ExeStringPro_core(commandArray)
  If buf(1) = 0 Then
    Call SetVariant(ExeStringPro, buf(2))
  Else
    MsgBox "specified function:" & commandString & " failed. check that the function exists and args are valid."
  End If
End Function '}}}

Function ExeStringPro_core(commandArray) As Variant '{{{
  'return (Err.Number, result)
  Dim buf As New Collection

  'Debug.Print "Start ExeStringPro_core"
  'TODO: case for the number of args is greater than 3
  On Error GoTo MyError
  If UBound(commandArray) = 0 Then
    debug.print commandArray(0)
    Call SetVariant(result, Application.run(commandArray(0)))
  ElseIf UBound(commandArray) = 1 Then
    Call SetVariant(result, Application.run(commandArray(0), commandArray(1)))
  ElseIf UBound(commandArray) = 2 Then
    Call SetVariant(result, Application.run(commandArray(0), commandArray(1), commandArray(2)))
  End If

MyError:
  buf.Add Err.Number 'return 0 if no error
  buf.Add result
  Set ExeStringPro_core = buf
  Set buf = Nothing
End Function '}}}

Sub SetVariant(a As Variant, b As Variant)'{{{
  If IsObject(b) Then
    Set a = b
  Else
    Let a = b
  End If
End Sub'}}}

Function ExecuteAsIs(code As String)'{{{
  'Todo return value but that seems to be a little bit dangerous

  With ThisWorkbook.VBProject.VBComponents("oneliner").CodeModule
    .DeleteLines StartLine:=1, count:=.CountOfLines
    .InsertLines 1, "Sub temp_for_ExecuteAsIs()"
    .InsertLines 2, "End Sub"
    .InsertLines 2, code
  End With
  DoEvents
  Application.Run("temp_for_ExecuteAsIs")
End Function'}}}

Public Function ExecCommand(sCommand As String, sResult As String) As Boolean  '{{{
  'http://www.f3.dion.ne.jp/~element/msaccess/AcTipsGetDosResult.html

  Dim oShell As Object, oExec As Object  
  Set oShell = CreateObject("WScript.Shell")  
  Set oExec = oShell.Exec("%ComSpec% /c " & sCommand)  

  ' wait the process finished
  Do Until oExec.status: DoEvents: Loop  

    ' set result
    If Not oExec.StdErr.AtEndOfStream Then  
      ExecCommand = True  
      sResult = oExec.StdErr.ReadAll  
    ElseIf Not oExec.StdOut.AtEndOfStream Then  
      sResult = oExec.StdOut.ReadAll  
    End If  

    ' release the reference of obeject variable
    Set oExec = Nothing: Set oShell = Nothing  
  End Function  '}}}

