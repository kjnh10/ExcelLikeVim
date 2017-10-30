Attribute VB_Name = "copyRange"

Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

'**
' コピーアドレスの取得
'**
Public Function GetCopiedRange(SheetName As String) As Range
  Dim i As Long
  Dim Format As Long
  Dim hMem As Long
  Dim p As Long
  Dim Data() As Byte
  Dim Size As Long
  Dim Address As String

  Call OpenClipboard(0)
  hMem = GetClipboardData(RegisterClipboardFormat("Link"))
  If hMem = 0 Then
    Call CloseClipboard
    Exit Function
  End If

  Size = GlobalSize(hMem)
  p = GlobalLock(hMem)
  ReDim Data(0 To Size - 1)
  Call MoveMemory(VarPtr(Data(0)), p, Size)
  Call GlobalUnlock(hMem)

  Call CloseClipboard

  For i = 0 To Size - 1
    If Data(i) = 0 Then
      Data(i) = Asc(" ")
    End If
  Next i

  Dim buf As String
  buf = Trim(AnsiToUnicode(Data()))
  Address = Right(buf, InStr(StrReverse(buf), " ") - 1)

  Set GetCopiedRange = Range(Application.ConvertFormula(Address, xlR1C1, xlA1))
End Function

'**
' Unicode変換
'**
Private Function AnsiToUnicode(ByRef Ansi() As Byte) As String
  On Error GoTo ErrHandler
  Dim Size   As Long
  Dim Buf    As String
  Dim BufLen As Long
  Dim RtnLen As Long

  Size = UBound(Ansi) + 1
  BufLen = Size * 2 + 10
  Buf = String$(BufLen, vbNullChar)
  RtnLen = MultiByteToWideChar(0, 0, Ansi(0), Size, StrPtr(Buf), BufLen)
  If RtnLen > 0 Then
    AnsiToUnicode = Left$(Buf, RtnLen)
  End If
ErrHandler:
End Function
