Attribute VB_Name = "unite_command"

Function GatherCandidates_command() As Collection'{{{
  ' Declare variables to access the Excel 2007 workbook.'{{{
  Dim objXLWorkbooks As Excel.Workbooks
  Dim objXLABC As Excel.Workbook'}}}
  ' Declare variables to access the macros in the workbook.'{{{
  Dim VBAEditor As VBIDE.VBE
  Dim objProject As VBIDE.VBProject
  Dim objComponent As VBIDE.VBComponent
  Dim objCode As VBIDE.CodeModule'}}}
  ' Declare other miscellaneous variables.'{{{
  Dim iLine As Integer
  Dim sProcName As String
  Dim pk As vbext_ProcKind'}}}

  Dim result As New Collection
  ' For Each objComponent In Application.VBE.ActiveVBProject.VBComponents
  For Each objComponent In ThisWorkbook.VBProject.VBComponents
    ' Find the code module for the project.
    Set objCode = objComponent.CodeModule

    ' Scan through the code module, looking for procedures.
    iLine = 1
    Do While iLine < objCode.CountOfLines
      sProcName = objCode.ProcOfLine(iLine, pk)
      If sProcName <> "" Then
        result.Add objComponent.Name & "." & sProcName
        ' Found a procedure. Display its details, and then skip to the end of the procedure.
        iLine = iLine + objCode.ProcCountLines(sProcName, pk)
      Else
        iLine = iLine + 1
      End If
    Loop
    Set objCode = Nothing
    Set objComponent = Nothing
  Next
  Set GatherCandidates_command = result
  Set result = Nothing
End Function'}}}

Function defaultAction_command(arg) 'table‚É‚µ‚½•û‚ª‚æ‚¢‚©'{{{
  For Each f in Split(arg, vbCrLf)
    ExeStringPro(f)
  Next f
End Function'}}}

Function defaultAction_command_parent(arg) 'table‚É‚µ‚½•û‚ª‚æ‚¢‚©'{{{
  For Each f in Split(arg, vbCrLf)
    Dim g As Variant
    For Each g in Split(unite_argument, vbCrlF)
      ExeStringPro(f & " " & g)
    Next g
  Next f
End Function'}}}

Sub kojikoji(arg)
  Msgbox arg
End Sub
