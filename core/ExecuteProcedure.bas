Attribute VB_Name = "ExecuteProcedure"

Function ExeStringPro(commandString As String, Optional bookName As String = "") '{{{
  'bookNameのmoduleを優先で探して実行。見つからなければこのブックのコマンドを探して実行。

  'Debug.Print "Start ExeStringPro"
  Dim commandArray() As String
  Dim AWBcommandArray() As String
  commandArray = Split(commandString, " ")

  If commandArray(0) = "-a" Then
    commandString = Mid(commandString, Instr(commandString, " ") + 1) '2つ目のスペース以降を取得
    ExecuteAsIs commandString
    Exit Function
  End If

  If bookName = "" Then
    On Error Resume Next 'for when there is no book 
    bookName = ActiveWorkbook.Name
    On Error Goto 0
  End If
  AWBcommandArray = commandArray
  AWBcommandArray(0) = bookName & "!" & commandArray(0)

  Set buf = ExeStringPro_core(AWBcommandArray)
  If buf(1) = 0 Then 'Search command within ActiveWorkbook code
    Call SetVariant(ExeStringPro, buf(2))
  Else
    Set buf = ExeStringPro_core(commandArray)
    If buf(1) = 0 Then
      Call SetVariant(ExeStringPro, buf(2))
    Else
      MsgBox "指定された関数" & commandString & "の実行に失敗しました｡ 関数が存在しているか､引数が不正でないか確認して下さい｡"
    End If
  End If
End Function '}}}

Function ExeStringPro_core(commandArray) As Variant '{{{
  'return (Err.Number, result)
  Dim buf As New Collection

  'Debug.Print "Start ExeStringPro_core"
  'TODO:引数が3つ以上ある関数の場合の処理
  On Error GoTo MyError
  If UBound(commandArray) = 0 Then
    Call SetVariant(result, Application.run(commandArray(0)))
  ElseIf UBound(commandArray) = 1 Then
    Call SetVariant(result, Application.run(commandArray(0), commandArray(1)))
  ElseIf UBound(commandArray) = 2 Then
    Call SetVariant(result, Application.run(commandArray(0), commandArray(1), commandArray(2)))
  End If

MyError:
  buf.Add Err.Number 'errorがなければ0が返る。
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
  ' 目　的：DOS コマンドの実行結果を取得します。  
  ' 戻り値：エラーの有無を Boolean 型で返します。  
  ' 　　　　エラー発生時は True、正常終了時は False です。  
  ' 引　数：sCommand-> 必須/入力用です。実行コマンドを文字列型で渡します。  
  ' 　　　　sResult -> 必須/出力用です。実行結果を文字列型で受け取ります。  
  '　　　　　　　　　　失敗した場合はエラー内容を示します。  
  ' 注　意：実行中はコマンドプロンプト ウィンドウが開きます。また実行後は自動的にウィンドウが閉じます。  
  'http://www.f3.dion.ne.jp/~element/msaccess/AcTipsGetDosResult.html

  Dim oShell As Object, oExec As Object  
  Set oShell = CreateObject("WScript.Shell")  
  Set oExec = oShell.Exec("%ComSpec% /c " & sCommand)  

  ' 処理完了を待機します。  
  Do Until oExec.status: DoEvents: Loop  

    ' 戻り値をセットします。  
    If Not oExec.StdErr.AtEndOfStream Then  
      ExecCommand = True  
      sResult = oExec.StdErr.ReadAll  
    ElseIf Not oExec.StdOut.AtEndOfStream Then  
      sResult = oExec.StdOut.ReadAll  
    End If  

    ' オブジェクト変数の参照を解放します。  
    Set oExec = Nothing: Set oShell = Nothing  
  End Function  '}}}

