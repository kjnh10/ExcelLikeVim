Attribute VB_Name = "FileOperation"
'Option Explicit

'File operation
Sub w(Optional fileName As String = "") '{{{
  If fileName = "" Then
    Set wb = ActiveWorkbook
  Else
    Set wb = Workbooks(fileName)
  End If
  wb.Save
End Sub '}}}
Sub wa() '{{{
  For Each wb In Workbooks
    wb.Save
  Next
End Sub '}}}

Sub q(Optional fileName As String = "") '{{{
  If fileName = "" Then
    Set wb = ActiveWorkbook
  Else
    Set wb = Workbooks(fileName)
  End If
  wb.Close

  If Workbooks.count <= 1 Then
    On Error Resume Next
    Workbooks("register.xlsx").Close savechanges:=True
    Application.quit
  End If

End Sub '}}}
Sub q_exclamation() '{{{
  Set atwb = ActiveWorkbook
  atwb.Close savechanges:=False

  If Workbooks.count <= 1 Then
    On Error Resume Next
    Workbooks("register.xlsx").Close savechanges:=True
    Application.quit
  End If
End Sub '}}}
Sub qa() '{{{
  For Each Wb In Workbooks
    Wb.Close
  Next

  If Workbooks.count <= 1 Then
    On Error Resume Next
    Workbooks("register.xlsx").Close savechanges:=True
    Application.quit
  End If
End Sub '}}}
Sub qa_exclamation() '{{{
  For Each Wb In Workbooks
    Wb.Close savechanges:=False
  Next
  If Workbooks.count = 0 Then
    Application.quit
  End If
End Sub '}}}

Sub wq(Optional fileName As String = "") '{{{
  Call w(fileName)
  Call q(fileName)
End Sub '}}}

Sub cos() '{{{
  Set atwb = ActiveWorkbook
  Workbooks.CheckOut (atwb.Path & "\" & atwb.Name)
  atwb.Save
  atwb.CheckIn 'include closing
  If Workbooks.count = 0 Then
    Application.quit
  End If
End Sub '}}}

Sub co() '{{{
  Set atwb = ActiveWorkbook
  Workbooks.CheckOut (atwb.Path & "\" & atwb.Name)
End Sub '}}}

Function Path() '{{{
  MsgBox ActiveWorkbook.Path
  Dim buf As String
  buf = ActiveWorkbook.Path
  SetStrToClipBoard(buf)
End Function '}}}

Function fpath() '{{{
  Dim AWF As String
  AWF = ActiveWorkbook.FullName
  SetStrToClipBoard(AWF)
  MsgBox AWF
End Function '}}}

Function af() '{{{
  ActiveCell.EntireColumn.Autofit
  ActiveCell.EntireRow.Autofit
End Function '}}}

Sub cd(Optional bookName = "") '{{{
  If bookName = "" Then
    Set book = ActiveWorkbook
  Else
    Set book = Workbooks(bookName)
  End If

  On Error GoTo MyError
  If Left(book.FullName, 2) <> "\\" And Left(book.FullName, 2) <> "ht" Then
    ChDrive book.Path
    ChDir book.Path
  Else
    CreateObject("WScript.Shell").CurrentDirectory = book.Path
  End If

  Debug.Print "moved to " & Curdir
  Exit Sub
MyError:
  MsgBox "failed cd " & "\n" & Err.Description
End Sub '}}}

Sub update() '{{{
  If Not ActiveSheet.AutoFilter Is Nothing Then
    ActiveSheet.AutoFilter.ApplyFilter
  End If
  ActiveSheet.Calculate
End Sub '}}}

Public Function SmartOpenBook(filePath) '{{{
  Dim buf As String, Wb As Workbook

  On Error Goto Myerror
  buf = dir(filePath)
  If buf = "" Then
    MsgBox filePath & vbCrLf & " not found", vbExclamation
    Exit Function
  End If

  'check same name book
  For Each Wb In Workbooks
    If Wb.FullName = filePath Then
      Wb.Activate
      Exit Function
    End If
  Next Wb

  DoEvents
  ' Workbooks.Open FileName:=filePath, Notify:=True, AddToMru:=True
  CreateObject("Wscript.Shell").Run Chr(34) & filePath & Chr(34), 5

  Exit Function
Myerror:
  MsgBox Err.Description & vbCrLf & "Alternatively filepath was copied to clipboard"
  SetStrToClipBoard(filePath)
End Function '}}}

'PDF copy
Public Sub PrintPdfDir(dirPath As String) '{{{
  Dim FileName As String
  Dim fileList As New Collection
  FileName = dir(dirPath)
  Do While FileName <> ""
    fileList.Add Item:=dirPath & FileName
    FileName = dir
  Loop
  On Error GoTo err
  For Each f In fileList
    AdobeReader.PrintPdf FilePath:=CStr(f)
  Next f
  Exit Sub
  err:
  MsgBox err.number & vbCr & err.Description, vbCritical
End Sub  '}}}
