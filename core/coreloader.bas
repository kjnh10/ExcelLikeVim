Attribute VB_Name = "coreloader"

'Declare SHCreateDirectoryEx' {{{
#If VBA7 And Win64 Then
  Private Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" _
    Alias "SHCreateDirectoryExA" _
    (ByVal hwnd As LongPtr, ByVal pszPath As String, _
    ByVal psa As LongPtr) As Long
#Else
  Private Declare Function SHCreateDirectoryEx Lib "shell32" _
    Alias "SHCreateDirectoryExA" _
    (ByVal hwnd As Long, ByVal pszPath As String, _
    ByVal psa As Long) As Long
#End If  '}}}

Private Enum Module'{{{
Standard = 1
Class = 2
Forms = 3
ActiveX = 11
Document = 100
End Enum'}}}

Public Sub init()'{{{
  Call SetReference

  'default setting
  Call RegisterModule(ThisWorkbook.Path & "\configure.bas", ThisWorkbook)
  Call initModule("configure")

  'user setting
  On Error GoTo except
    Call RegisterModule(uDir & "user_configure.bas", ThisWorkbook)
    Call initModule("user_configure")
  except:
    If Err.Number <> 0 Then
      Debug.print Err.Description
    End If
End Sub'}}}

Public Function Udir() As String  'user_folder'{{{
  Udir = Environ("homedrive") & Environ("homepath") & "\vimx\"  
End Function'}}}

Public Sub reload() '{{{
  'Dont't call with of from a module functions because that module will not be deleted while this function is being executed'
  Application.Run "clearAllModules"
  'load core modules
  Call LoadPluginDir(ThisWorkbook.Path & "\core\", ThisWorkbook.Name, True)
  'load standard plugins
  Call LoadPluginDir(ThisWorkbook.Path & "\sys_plugin", ThisWorkbook.Name, True)
  'load user plugins
  Call LoadPluginDir(Udir & "plugin", ThisWorkbook.Name, True)

  init

  Msgbox "All Modules were successfully updated."
End Sub '}}}

'------------------ Update ------------------------
Public Sub LoadPluginDir(Optional DirPath As String = "", Optional targetBookName As String = "", Optional isCalledFromThisWorkbookModule = False) '{{{
  Dim msgError As String: msgError = "Error Message"

  Set FSO = CreateObject("Scripting.FileSystemObject")
  Set moduleList = ModuleListOfPlugin(DirPath)

  For Each modulePath in moduleList
    filename = FSO.GetFileName(modulePath)
    moduleName = Left(filename, InStr(filename, ".")-1) 'remove extention
    If isCalledFromThisWorkbookModule And moduleName = "ThisWorkbook" Then
      'pass
    ElseIf moduleName <> "coreloader" Then 'to update coreloader.bas, restart Application. ThisWorkbook.import_coreloader does that.
      Call RegisterModule(modulePath, Workbooks(targetBookName), msgError)
    End If
  Next modulePath
  'error handling'{{{
  If msgError = "Error Message" Then
  Else
    Msgbox msgError
  End If '}}}
End Sub '}}}

Private Sub RegisterModule(modulePath, Optional targetBook As Workbook = Nothing, Optional msgError As String = "")'{{{
  On Error GoTo except
    Set myFSO = CreateObject("Scripting.FileSystemObject")
    Dim moduleName As String: moduleName = myFSO.GetBaseName(modulePath)

    If moduleName <> "coreloader" Then
      If isMemberOfVBEComponets(targetBook, moduleName) Then
        ' remove won't get effect soon after executed, so neeed to rename for avoiding duplication when inporting.
        delname = moduleName & "_to_delete"
        ThisWorkbook.VBProject.VBComponents(moduleName).name = delname
        ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents.Item(delname)
      End If

      If Not isMemberOfVBEComponets(targetBook, moduleName) Then
        targetBook.VBProject.VBComponents.Import modulePath
      End If
    End If

  except:
    If Err.Number <> 0 Then
      msgError = msgError & vbCrLf & Err.Description & ": when updating " & moduleName
    End If
    Set myFSO = Nothing

End Sub'}}}

Private Sub initModule(moduleName, Optional msgError As String = "") '{{{
  On Error Resume Next
  Application.Run(moduleName & ".init")
  If Err.Number <> 0 And Err.Number <> 1004 Then
    msgError = msgError & vbCrLf & Err.Description & ":" & Err.Number
  End If
End Sub '}}}

Private Sub clearAllModules() '{{{
  Dim moduleList As Collection: Set moduleList = New Collection
  For Each component In ThisWorkbook.VBProject.VBComponents
    If (component.name <> "coreloader") and (component.Type = 1 Or component.Type = 2 Or component.Type = 3) Then
      'Standard(Type=1) / Class(Type=2) / Form(Type=3)
      moduleList.Add(component.name)
    End If
  Next component
  For Each modName in moduleList
    Set component = ThisWorkbook.VBProject.VBComponents.Item(modName)
    ThisWorkbook.VBProject.VBComponents.Remove component
  Next modName
End Sub '}}}

'------------------ Other -------------------------
Private Sub SetReference()'{{{
  'TODO to be able to specify a book
  'TODO stop hard coding
  'unite_command 用 本来はプラグイン側からの呼び出しを出来るようにしたい｡
  AddToReference("C:\Program Files\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB")
  AddToReference("C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB")

  AddToReference("C:\Program Files\Microsoft Office 15\Root\Office15\MSPPT.OLB")
  AddToReference("C:\Program Files (x86)\Microsoft Office 15\Root\Office15\MSPPT.OLB")

  AddToReference("C:\Program Files\Microsoft Office 15\Root\Office15\MSWORD.OLB")
  AddToReference("C:\Program Files (x86)\Microsoft Office 15\Root\Office15\MSWORD.OLB")
End Sub'}}}

Private Function AddToReference(strFileName As String) As Boolean'{{{
  '指定されたタイプライブラリへの参照を作成します｡
  On Error GoTo MyError
  ' Dim ref As Reference
  ' Set ref = ThisWorkbook.VBProject.References.AddFromFile(strFileName)
  ' AddToReference = True
  ' Set ref = Nothing
  ThisWorkbook.VBProject.References.AddFromFile(strFileName)
  AddToReference = True
  Exit Function
MyError:
  Select Case Err.Number
    Case 32813
      'Debug.Print strFileName & "は既に参照設定されています。", , "タイプライブラリへの参照"
    Case 48
      'Debug.Print strFileName & "は存在しません。"
    Case 29060
      MsgBox "設定ファイルがインストールされていないか、" & vbNewLine & _
      "所定のフォルダーに存在しない場合が考えられます。" & vbNewLine & _
      "よって、参照設定ができません。", , "タイプライブラリへの参照"
    Case Else
      MsgBox "予期せぬエラーが発生しました。" & vbNewLine & _
      Err.Number & vbNewLine & _
      Err.Description, 16, "タイプライブラリへの参照"
  End Select
End Function'}}}

Private Sub printReferencesName()'{{{
  'For Investigation
  For Each r in ThisWorkbook.VBProject.References
    Debug.Print r.FullName
  End Sub'}}}

  '------------------- common Functions / Subs --------------
  Private Function isExcelObject(fileName As String) As Boolean'{{{
    Set RE = CreateObject("VBScript.RegExp")
    RE.IgnoreCase = True
    RE.pattern = ".cls$|.frm|ThisWorkbook|Sheet"
    If RE.test(fileName) Then
      isExcelObject = True
    Else
      isExcelObject = False
    End If
  End Function'}}}

  Private Function getExtention(myComponent) As String'{{{
    Dim extention As String
    Select Case myComponent.Type
      Case Module.Standard
        extention = ".bas"
      Case Module.Class
        extention = ".cls"
      Case Module.Forms
        extention = ".frm"
      Case Module.ActiveX
        extention = ".cls"
      Case Module.Document
        extention = ".cls"
    End Select

    getExtention = extention
  End Function'}}}

  Private Function checkExistFile(ByVal pathFile As String) As Boolean'{{{
    On Error GoTo Err_dir
    If Dir(pathFile) = "" Then
      checkExistFile = False
    Else
      checkExistFile = True
    End If

    Exit Function

Err_dir:
    checkExistFile = False
  End Function'}}}

  Private Function isMemberOfCollection(col As Collection, query) As Boolean'{{{
    For Each item In col
      If item = query Then
        isMemberOfCollection = True
        Exit Function
      End If
    Next
    isMemberOfCollection = False
  End Function'}}}

  Private Function isMemberOfVBEComponets(book As Workbook, moduleName As String) As Boolean '{{{
    'Argument: moduleName like CodeManager
    'Return: whether or not module is registered
    For Each Item In book.VBProject.VBComponents
      If Item.Name = moduleName Then
        isMemberOfVBEComponets = True
        Exit Function
      End If
    Next
    isMemberOfVBEComponets = False
  End Function '}}}

  '再帰的なフォルダの取得'{{{
  Public Function ModuleListOfPlugin(folder_path As String) As Collection
    Set ModuleListOfPlugin = AllFiles(folder_path,  ".*(bas|cls|frm)$", "_stash")
  End Function

  Public Function AllFiles(folder_path As String, pattern As String, Optional excluded_folder_pattern As String = "") As Collection
    Dim result As Collection: Set result = New Collection
    Dim FSO As FileSystemObject: Set FSO = New FileSystemObject

    Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.pattern = pattern
    Call getRecursive(folder_path, RegExp, FSO, result, excluded_folder_pattern)

    Set RegExp = Nothing
    Set FSO = Nothing
    Set AllFiles = result
  End Function

  Private Sub getRecursive(folder_path As String, RegExp, FSO As FileSystemObject, result As Collection, Optional excluded_folder_pattern As String = "")
    ' 現在ディレクトリ内の全ファイルの取得
    Dim file_path As Variant
    For Each file_path In FSO.GetFolder(folder_path).Files
      If RegExp.test(file_path) Then
        DoEvents    ' フリーズ防止用
        Call result.Add(CStr(file_path))
      End If
    Next

    ' サブディレクトリの再帰
    Dim dir As Variant
    For Each dir In FSO.GetFolder(folder_path).SubFolders
      if Not (dir Like "*"& excluded_folder_pattern &"*" And excluded_folder_pattern <> "") then
        Call getRecursive(dir.Path, RegExp, FSO, result, excluded_folder_pattern)
      end if
    Next
  End Sub'}}}

  Public Sub MkDirRecursively(targetPath)
    rc = SHCreateDirectoryEx(0&, targetPath, 0&)
  End Sub
