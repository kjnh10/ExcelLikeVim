Attribute VB_Name = "AdobeRead"
'---------------------------------------------------------------------------------------
' Module    : AdobeReader
' Created   : 2010/02/06 21:14
' Updated   : 2010/09/17 0:23
' Version   : 1.1.0
' Author    : YU-TANG
' Purpose   : Adobe Reader による文書の表示と印刷
' Reference : http://www.f3.dion.ne.jp/~element/msaccess/AcTipsAdobeReader.html
' History   : 2010/02/09 1.0.0 Initioal Release
'             2010/09/17 1.1.0 Ver.9 対応(Search と View オプション)
'---------------------------------------------------------------------------------------
Option Compare Binary
'Option Explicit

' ************ PDF 表示用 ************
' ページモード列挙定数
Public Enum OpenPdfPageMode
    oppmNone                ' 指定なし
    oppmBookmarks           ' しおり
    oppmThumbs              ' サムネール
End Enum

' 表示 列挙定数
Public Enum OpenPdfView
    opvNone                 ' 指定なし
    opvFitPage              ' 全体表示
    opvFitWidth             ' 幅に合わせる
    opvFitHeight            ' 高さに合わせる
    opvFitVisible           ' 描画領域の幅に合わせる
    opvRotateRight = &H10   ' 右90°回転
    opvRotateLeft = &H20    ' 左90°回転
End Enum

' ************ レジストリ関連 ************
' 環境によっては WScript.Shell オブジェクトの生成が禁止されている
' 場合があるようなので、API でレジストリへアクセスします。

'参照
'Shell Lightweight Utility APIs - HEROPA's HomePage
'http://www31.ocn.ne.jp/~heropa/vb123.htm#SHGetValue
Private Enum hKeyConstants
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

' DWORD型のタイプ
Private Enum RegTypeConstants
'    REG_NONE = (0)                         ' 定義されていない種類
    REG_SZ = (1)                           ' NULL で終わる文字列
'    REG_EXPAND_SZ = (2)                    ' 展開前の環境変数への参照 が入った NULL で終わる文字列
'    REG_BINARY = (3)                       ' 任意の形式のバイナリデータ
    REG_DWORD = (4)                        ' 32 ビット値
    REG_DWORD_LITTLE_ENDIAN = (4)          ' リトルエンディアン形式の 32 ビット値
'    REG_DWORD_BIG_ENDIAN = (5)             ' ビッグエンディアン形式の 32 ビット値
'    REG_LINK = (6)                         ' Unicode のシンボリックリンク
'    REG_MULTI_SZ = (7)                     ' NULL で終わる文字列の配列
'    REG_RESOURCE_LIST = (8)                ' デバイスドライバのリソースリスト
End Enum


Private Const ERROR_SUCCESS     As Long = 0

Private Declare Function SHGetValue Lib "SHLWAPI.DLL" Alias "SHGetValueA" _
                                (ByVal hKey As Long, _
                                 ByVal pszSubKey As String, _
                                 ByVal pszValue As String, _
                                 pdwType As Long, _
                                 pvData As Any, _
                                 pcbData As Long) As Long

' ************ ウィンドウ取得関連 ************
'参照
'インスタンス ハンドルからウィンドウのハンドルを検索する方法
'http://support.microsoft.com/kb/242308/ja
Private Const GW_HWNDNEXT = 2

Private Declare Function GetParent Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindow Lib "User32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function IsWindow Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hWnd As Long, lpdwprocessid As Long) As Long
Private Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMillsecounds As Long)
Private Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_COMMAND                As Long = &H111&
Private Const WM_CLOSE                  As Long = &H10&     '終了メッセージ

Private Const MENU_ID_ZOOM_FIT_PAGE     As Long = 6074&     ' [表示]-[ズーム]-[全体表示]
Private Const MENU_ID_ZOOM_FIT_WIDTH    As Long = 6075&     ' [表示]-[ズーム]-[幅に合わせる]
Private Const MENU_ID_ZOOM_FIT_HEIGHT   As Long = 6076&     ' [表示]-[ズーム]-[高さに合わせる]
Private Const MENU_ID_ZOOM_FIT_VISIBLE  As Long = 6077&     ' [表示]-[ズーム]-[描画領域の幅に合わせる]
Private Const MENU_ID_VIEW_ROTATE_RIGHT As Long = 6090&     ' [表示]-[表示を回転]-[右90°回転]
Private Const MENU_ID_VIEW_ROTATE_LEFT  As Long = 6091&     ' [表示]-[表示を回転]-[左90°回転]
Private Const MENU_ID_EDIT_SEARCH       As Long = 6042&     ' [編集]-[簡易検索]

' ************ UTF-8 変換関連 ************
'based on:
'保存形式をUTF-8にしたい
'http://rararahp.cool.ne.jp/cgi-bin/lng/vb/vblng.cgi?print+200508/05080003.txt
Private Declare Function WideCharToMultiByte Lib "kernel32" _
        (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, _
         lpMultiByteStr As Byte, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Private Const CP_UTF8 = 65001

' ************ ShellExecute 関連 ************
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
     ByVal lpOperation As String, _
     ByVal lpFile As String, _
     Optional ByVal lpParameters As String, _
     Optional ByVal lpDirectory As String, _
     Optional ByVal nShowCmd As VbAppWinStyle) As Long

' ************ バージョン情報関連 ************
Private Type VS_FIXEDFILEINFO
        dwSignature As Long
        dwStrucVersion As Long         '  e.g. 0x00000042 = "0.42"
        dwFileVersionMS As Long        '  e.g. 0x00030075 = "3.75"
        dwFileVersionLS As Long        '  e.g. 0x00000031 = "0.31"
        dwProductVersionMS As Long     '  e.g. 0x00030010 = "3.10"
        dwProductVersionLS As Long     '  e.g. 0x00000031 = "0.31"
        dwFileFlagsMask As Long        '  = 0x3F for version "0.42"
        dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
        dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
        dwFileType As Long             '  e.g. VFT_DRIVER
        dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
        dwFileDateMS As Long           '  e.g. 0
        dwFileDateLS As Long           '  e.g. 0
End Type

Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dst As Any, src As Any, ByVal num As Long)


'---------------------------------------------------------------------------------------
' Procedure : OpenPdf
' DateTime  : 2010/02/02 21:20
' Author    : YU-TANG
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function OpenPdf( _
        ByRef FilePath As String, _
        Optional ByVal Page As Long, _
        Optional ByVal Comment As String, _
        Optional ByVal Zoom As String, _
        Optional ByVal PageMode As OpenPdfPageMode = oppmNone, _
        Optional ByVal ScrollBar As Variant, _
        Optional ByVal Search As String, _
        Optional ByVal ToolBar As Variant, _
        Optional ByVal NavPanes As Boolean, _
        Optional ByVal View As OpenPdfView, _
        Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus) As Double
On Error GoTo eh

    Dim App         As String
    Dim Command     As String
    Dim OpenActions As String
    Dim OpenAction  As String
    Dim hInst       As Long         ' Instance handle from Shell function.
    Dim hWndApp     As Long         ' Window handle from GetWinHandle.
    Dim MajorVer    As Long

    ' Adobe Reader のパスを取得
    App = AdobeReaderPath

    ' Adobe Reader のメジャーバージョンを取得
    Call GetVersion(App, MajorVer)

    ' ファイルが存在するかチェック
    IsFileExists FilePath
    
    '以下、OpenActions 引数

    ' -- 1.ページ番号
    If Page > 0 Then
        OpenAction = "Page=" & Page
        GoSub AddOpenAction
    ' -- 2.コメント
        If Comment <> vbNullString Then
            OpenAction = "Comment=" & Comment
            GoSub AddOpenAction
        End If
    End If

    ' -- 3.表示倍率
    If Zoom <> vbNullString Then
        OpenAction = "Zoom=" & Zoom
        GoSub AddOpenAction
    End If

    ' -- 4.ページモード
    Select Case PageMode
        Case oppmBookmarks
            OpenAction = "PageMode=bookmarks"
            GoSub AddOpenAction
            NavPanes = True
        Case oppmThumbs
            OpenAction = "PageMode=thumbs"
            GoSub AddOpenAction
            NavPanes = True
    End Select
    
    ' -- 5.スクロールバー
    If Not IsMissing(ScrollBar) Then
        OpenAction = "ScrollBar=" & IIf(ScrollBar <> 0, "1", "0")
        GoSub AddOpenAction
    End If

    ' -- 6.検索
    If Search <> vbNullString Then
        Select Case MajorVer
            Case Is <= 8    ' Version 8 以前
                OpenAction = "Search=""" & UrlEncodeUTF8(Search) & """"
            Case Else       ' Version 9 以後
                OpenAction = "Search=" & UrlEncodeUTF8(Search)
        End Select
        GoSub AddOpenAction
    End If

    ' -- 7.ツールバー
    If Not IsMissing(ToolBar) Then
        OpenAction = "ToolBar=" & IIf(ToolBar <> 0, "1", "0")
        GoSub AddOpenAction
    End If

    ' -- 8.ナビゲーションパネル
    '  + ページモードで「しおり」か「サムネール」指定時は、
    '    ナビゲーションパネルの指定は無視され、常に表示されます。
    OpenAction = "NavPanes=" & IIf(NavPanes, "1", "0")
    GoSub AddOpenAction

    ' コマンドを生成
    If View <> opvNone Then
        ' NewInstance スイッチ /n を付けないと、既存のインスタンスが
        ' 存在した場合に Window ハンドルを取れないため、表示オプション
        ' が指定された場合は /n を強制的に付加します。
        Command = "'<App>' /n /a '<OpenActions>' '<File>'"
    Else
        Command = "'<App>' /a '<OpenActions>' '<File>'"
    End If
    Command = Replace(Command, "'", """")
    Command = Replace(Command, "<App>", App)
    Command = Replace(Command, "<File>", FilePath)
    If OpenActions <> vbNullString Then
        Command = Replace(Command, "<OpenActions>", OpenActions)
    End If

    ' PDF を開く
    hInst = Shell(Command, WindowStyle)
    OpenPdf = hInst

    ' 表示オプション
    If View <> opvNone Then
        hWndApp = GetWinHandle(hInst)
        If hWndApp <> 0& Then
            Select Case MajorVer
                Case Is <= 8    ' Version 8 以前
                    ' Rotate
                    Select Case View And &HF0
                        Case opvRotateRight     ' 右90°回転
                            PostMessage hWndApp, WM_COMMAND, MENU_ID_VIEW_ROTATE_RIGHT, 0&
                        Case opvRotateLeft      ' 左90°回転
                            PostMessage hWndApp, WM_COMMAND, MENU_ID_VIEW_ROTATE_LEFT, 0&
                    End Select
                    
                    ' Zoom
                    Select Case View And &HF
                        Case opvFitPage         ' 全体表示
                            PostMessage hWndApp, WM_COMMAND, MENU_ID_ZOOM_FIT_PAGE, 0&
                        Case opvFitWidth        ' 幅に合わせる
                            PostMessage hWndApp, WM_COMMAND, MENU_ID_ZOOM_FIT_WIDTH, 0&
                        Case opvFitHeight       ' 高さに合わせる
                            PostMessage hWndApp, WM_COMMAND, MENU_ID_ZOOM_FIT_HEIGHT, 0&
                        Case opvFitVisible      ' 描画領域の幅に合わせる
                            PostMessage hWndApp, WM_COMMAND, MENU_ID_ZOOM_FIT_VISIBLE, 0&
                    End Select

                Case Else       ' Version 9 以後
                    ' Rotate
                    Select Case View And &HF0
                        Case opvRotateRight     ' 右90°回転
                            PostMessage hWndApp, WM_COMMAND, MENU_ID_VIEW_ROTATE_RIGHT + 8, 0&
                        Case opvRotateLeft      ' 左90°回転
                            PostMessage hWndApp, WM_COMMAND, MENU_ID_VIEW_ROTATE_LEFT + 8, 0&
                    End Select
                    
                    ' Zoom
                    Select Case View And &HF
                        Case opvFitPage         ' 全体表示
                            PostMessage hWndApp, WM_COMMAND, MENU_ID_ZOOM_FIT_PAGE + 8, 0&
                        Case opvFitWidth        ' 幅に合わせる
                            PostMessage hWndApp, WM_COMMAND, MENU_ID_ZOOM_FIT_WIDTH + 8, 0&
                        Case opvFitHeight       ' 高さに合わせる
                            PostMessage hWndApp, WM_COMMAND, MENU_ID_ZOOM_FIT_HEIGHT + 8, 0&
                        Case opvFitVisible      ' 描画領域の幅に合わせる
                            PostMessage hWndApp, WM_COMMAND, MENU_ID_ZOOM_FIT_VISIBLE + 8, 0&
                    End Select
            End Select  ' MajorVer
        End If  ' hWndApp <> 0&
    End If  ' View <> opvNone

    Exit Function

AddOpenAction:
    If OpenActions <> vbNullString Then
        OpenActions = OpenActions & "&"
    End If
    OpenActions = OpenActions & OpenAction
    Return

eh:
    If Err.Number = 16 Then
        ' Shell 実行時に '式が複雑すぎます' エラー。
        ' 起きたり起きなかったり。原因不明だが、動作に支障はないので無視。
    Else
        ' 呼び出し元に通知するため、改めて実行時エラーを発生させる。
        Dim num  As Long:   num = Err.Number
        Dim desc As String: desc = Err.Description
        On Error GoTo 0
        Err.Raise num, "OpenPdf", desc
    End If

End Function

'---------------------------------------------------------------------------------------
' Procedure : PrintPdf
' DateTime  : 2010/02/04 00:21
' Author    : YU-TANG
' Purpose   :
' Return    : ShowPrintSettings 引数 = True の場合は、内部的に使用した Shell 関数の
'             戻り値をそのまま返却します。詳細はヘルプを参照してください。
'             ShowPrintSettings 引数 = False の場合は、内部的に使用した ShellExecute
'             API 関数の戻り値をそのまま返却します。詳細は下記を参照してください。
'             http://msdn.microsoft.com/ja-jp/library/cc422072.aspx
'---------------------------------------------------------------------------------------
'
Public Function PrintPdf( _
        ByRef FilePath As String, _
        Optional ByRef PrinterName As String, _
        Optional ByRef DriverName As String, _
        Optional ByRef PortName As String, _
        Optional ByVal ShowPrintSettings As Boolean) As Double
On Error GoTo eh

    Dim App         As String
    Dim Command     As String
    Dim hInst       As Long         ' Instance handle from Shell function.
    Dim hWndApp     As Long         ' Window handle from GetWinHandle.

    ' Adobe Reader のパスを取得
    App = AdobeReaderPath

    ' ファイルが存在するかチェック
    IsFileExists FilePath

    ' 印刷設定を表示する場合
    If ShowPrintSettings Then
        ' コマンドを生成
        Command = "'<App>' /s /p '<File>'"      'Print with dialog
        GoSub ParseCommand

        ' PDF を印刷
        hInst = Shell(Command, vbHide)
        PrintPdf = hInst

        ' 印刷後終了
        ' -- ダイアログを表示した場合は終了が難しいので割愛
        'hWndApp = GetWinHandle(hInst)
        'If hWndApp <> 0& Then
        '    PostMessage hWndApp, WM_CLOSE, 0&, 0&
        'End If

    ' 印刷設定を表示しない場合
    Else
        ' PDF を印刷
        ' -- Shell 関数や CreateProcess を使ってみたが、どうしても一瞬
        '    Adobe Reader のウィンドウが表示される。また、印刷後も
        '    ウィンドウが残るので、終了させる必要がある。
        '    そのため、ウィンドウが表示されない ShellExecute の方を使う。

        If PrinterName = vbNullString Then
            PrintPdf = ShellExecute(Application.hWnd, "print", FilePath)
        Else
            ' コマンドを生成
            Command = "'<PrinterName>' '<DriverName>' '<PortName>'"  'PrintTo
            GoSub ParseCommand
            PrintPdf = ShellExecute(Application.hWnd, "printto", FilePath, Command)
        End If
    End If

    Exit Function

eh:
    If Err.Number = 16 Then
        ' Shell 実行時に '式が複雑すぎます' エラー。
        ' 起きたり起きなかったり。原因不明だが、動作に支障はないので無視。
    Else
        ' 呼び出し元に通知するため、改めて実行時エラーを発生させる。
        Dim num  As Long:   num = Err.Number
        Dim desc As String: desc = Err.Description
        On Error GoTo 0
        Err.Raise num, "PrintPdf", desc
    End If
    Exit Function

ParseCommand:
    Command = Replace(Command, "'", """")
    Command = Replace(Command, "<App>", App)
    Command = Replace(Command, "<File>", FilePath)
    Command = Replace(Command, "<PrinterName>", PrinterName)
    Command = Replace(Command, "<DriverName>", DriverName)
    Command = Replace(Command, "<PortName>", PortName)
    Return

End Function



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' 以下、サブルーチン
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Private Function AdobeReaderPath() As String
    
    Const SUB_KEY = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\AcroRd32.exe"
    Dim sPath As String

    ' Adobe Reader のパスを取得
    sPath = RegGetValue(HKEY_LOCAL_MACHINE, SUB_KEY, "", REG_SZ, "")
    ' 前後を二重引用符で括られていた場合に備えて、引用符を削除
    sPath = Replace(sPath, """", vbNullString)
    If sPath = vbNullString Then
        Err.Raise 5, "OpenPdf", "Adobe Reader が見つかりません。"
    Else
        AdobeReaderPath = sPath
    End If

End Function

' ファイルが見つからない場合は実行時エラー発生
Private Sub IsFileExists(ByRef strFilePath As String)

    If strFilePath <> vbNullString Then
        If Dir$(strFilePath) <> vbNullString Then
            Exit Sub
        End If
    End If

    Err.Raise 53    ' File not found

End Sub

'Shell Lightweight Utility APIs - HEROPA's HomePage
'http://www31.ocn.ne.jp/~heropa/vb123.htm#SHGetValue

'
' レジストリの値を取得する。
'
Private Function RegGetValue(lnghInKey As hKeyConstants, _
                            ByVal strSubKey As String, _
                            ByVal strValName As String, _
                            lngType As RegTypeConstants, _
                            ByVal varDefault As Variant) As Variant
    ' lngInKey   : キー
    ' strSubKey  : サブキー
    ' strValName : 値
    ' lngType    : データタイプ
    ' lngDefault : デフォルトの値
    ' 戻り値     : 対応する値
    Dim varRetVal           As Variant
    Dim lnghSubKey          As Long
    Dim lngBuffer           As Long
    Dim strBuffer           As String
    Dim lngResult           As Long
    ' デフォルトの値を代入。
    varRetVal = varDefault
    Select Case lngType
        Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
            ' 何か値を入れておく。
            lngBuffer = 0
            lngResult = SHGetValue(lnghInKey, _
                                   strSubKey, _
                                   strValName, _
                                   REG_DWORD, _
                                   lngBuffer, _
                                   Len(lngBuffer))
            If lngResult = ERROR_SUCCESS Then
                varRetVal = lngBuffer
            End If
        Case REG_SZ
            ' バッファを確保する。
            strBuffer = String(256, vbNullChar)
            lngResult = SHGetValue(lnghInKey, _
                                   strSubKey, _
                                   strValName, _
                                   REG_SZ, _
                                   ByVal strBuffer, _
                                   Len(strBuffer))
            If lngResult = ERROR_SUCCESS Then
                varRetVal = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
            End If
    End Select
    RegGetValue = varRetVal
End Function


' 参照
'インスタンス ハンドルからウィンドウのハンドルを検索する方法
'http://support.microsoft.com/kb/242308/ja
Private Function ProcIDFromWnd(ByVal hWnd As Long) As Long
   Dim idProc As Long
   
   ' Get PID for this HWnd
   GetWindowThreadProcessId hWnd, idProc
   
   ' Return PID
   ProcIDFromWnd = idProc
End Function
      
Private Function GetWinHandle(hInstance As Long) As Long
    Dim hWnd        As Long
    Dim Length      As Long
    Dim sClassName  As String * 100

    ' Grab the first window handle that Windows finds:
    hWnd = FindWindow(vbNullString, vbNullString)

    ' Loop until you find a match or there are no more window handles:
    Do Until hWnd = 0&
        ' Check if no parent for this window
        If GetParent(hWnd) = 0& Then
            ' Check for PID match
            If hInstance = ProcIDFromWnd(hWnd) Then
                ' Check for class name match
                Length = GetClassName(hWnd, sClassName, 100&)     ' ウィンドウクラス
                If Left(sClassName, Length) = "AcrobatSDIWindow" Then
                    ' Return found handle
                    GetWinHandle = hWnd
                    ' Exit search loop
                    Exit Do
                End If
            End If
        End If

        ' Get the next window handle
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
    Loop
End Function

' ************ UTF-8 変換関連 ************
'based on:
'保存形式をUTF-8にしたい
'http://rararahp.cool.ne.jp/cgi-bin/lng/vb/vblng.cgi?print+200508/05080003.txt
Private Function UrlEncodeUTF8(ByRef strInput As String) As String

    Dim s           As String
    Dim p           As Long
    Dim buff()      As Byte
    Dim i           As Integer
    Dim fPercentize As Boolean

    buff = EncodeUTF8(strInput)
    s = Space$((UBound(buff) + 1) * 3)
    p = 1
    For i = LBound(buff) To UBound(buff)
        If buff(i) < 128 Then
            Select Case buff(i)
                Case 32, 34: fPercentize = True     ' スペースと二重引用符
                Case Else:   fPercentize = False
            End Select
        Else
            fPercentize = True
        End If
        If fPercentize Then
            Mid(s, p) = "%":           p = p + 1
            Mid(s, p) = Hex(buff(i)):  p = p + 2
        Else
            Mid(s, p) = Chr$(buff(i)): p = p + 1
        End If
    Next

    UrlEncodeUTF8 = Left$(s, p - 1)

End Function

Private Function EncodeUTF8(ByVal strInput As String) As Byte()

    Dim lngLength     As Long    ' 変換対象の文字数
    Dim lngSize       As Long    ' 変換後UTF8文字列バイト数
    Dim bytUTF8Buff() As Byte    ' 変換後UTF8文字列バッファ
    Dim lngBuffSize   As Long    ' 文字列バッファ領域数

    ' 変換対象文字数を取得
    lngLength = Len(strInput)
    If lngLength = 0 Then Exit Function

    ' 文字列バッファ領域を設定
    lngBuffSize = lngLength * 3

    ' 変換後文字列バッファ領域の確保
    ReDim bytUTF8Buff(lngBuffSize - 1)
    ' Unicode文字列→UTF8文字列変換
    lngSize = WideCharToMultiByte( _
                CP_UTF8, _
                0&, _
                StrPtr(strInput), _
                lngLength, _
                bytUTF8Buff(LBound(bytUTF8Buff)), _
                lngBuffSize, _
                0&, _
                0&)

    ' 変換失敗の場合は終了
    If lngSize = 0 Then Exit Function

    ' 不要な領域を開放
    ReDim Preserve bytUTF8Buff(lngSize - 1)

    EncodeUTF8 = bytUTF8Buff

End Function

' Based on:
' Visual Basic でファイルのバージョンを取得
' http://aircross.hp.infoseek.co.jp/vb_ver.htm
'
'   バージョン情報を取得
'
'       FullPath   バージョンを取得するファイルのフルパス
'       Major      メジャー リリース番号  格納先
'       Minor      マイナー リリース番号  格納先
'       RevisionH  リビジョン番号         格納先
'       RevisionL  リビジョン番号         格納先
'
'       戻り値      True:成功   False:失敗
'
Private Function GetVersion( _
        ByVal FullPath As String, _
        Optional ByRef Major As Long, _
        Optional ByRef Minor As Long, _
        Optional ByRef RevisionH As Long, _
        Optional ByRef RevisionL As Long _
    ) As Boolean

    GetVersion = False

    Dim ret         As Boolean
    Dim nLen        As Long
    Dim nHandle     As Long

    '   バージョン情報が取得できるかチェック
    Dim nVerInfoSize    As Long
    nVerInfoSize = GetFileVersionInfoSize(FullPath, 0&)
    If nVerInfoSize < 1 Then Exit Function

    '   バージョン情報を取得
    Dim cVerInfo()  As Byte
    ReDim cVerInfo(nVerInfoSize) As Byte
    ret = GetFileVersionInfo(FullPath, 0&, nVerInfoSize, cVerInfo(0))
    If ret = False Then Exit Function

    Dim vf  As VS_FIXEDFILEINFO
    ret = VerQueryValue(cVerInfo(0), "\", nHandle, nLen)
    CopyMemory vf.dwSignature, ByVal nHandle, nLen

    'File Version を
    '   Major, Minor, Revision に編集
    '(Product Version なら dwProductVersionMS と dwProductVersionLS を使う)
    CopyMemory Major, ByVal VarPtr(vf.dwFileVersionMS) + 2, 2
    CopyMemory Minor, vf.dwFileVersionMS, 2
    CopyMemory RevisionH, ByVal VarPtr(vf.dwFileVersionLS) + 2, 2
    CopyMemory RevisionL, vf.dwFileVersionLS, 2

    '** 参考:
    ' API でメモリ操作せず、VBA のみで求める場合は以下のようになる
'    If vf.dwFileVersionMS < 0 Then
'        Major = &HFFFF& - (Not vf.dwFileVersionMS) \ &H10000
'        Minor = &HFFFF& - Not vf.dwFileVersionMS
'    Else
'        Major = vf.dwFileVersionMS \ &H10000
'        Minor = vf.dwFileVersionMS Mod &H10000
'    End If
'    If vf.dwFileVersionMS < 0 Then
'        RevisionH = &HFFFF& - (Not vf.dwFileVersionLS) \ &H10000
'        RevisionL = &HFFFF& - Not vf.dwFileVersionLS
'    Else
'        RevisionH = vf.dwFileVersionLS \ &H10000
'        RevisionL = vf.dwFileVersionLS Mod &H10000
'    End If

    GetVersion = True

End Function

