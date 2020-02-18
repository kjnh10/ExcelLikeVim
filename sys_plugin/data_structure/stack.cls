VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute appEvent.VB_VarHelpID = -1

Option Explicit
Private ary() As Variant   '配列をスタックとして利用
Private size As Long       'スタックの現在のサイズ


Private Sub Class_Initialize()
  'コンストラクタ

  ReDim ary(0)
  size = 0

End Sub


Public Sub push(v As Variant)
  'スタックの一番上にデータを追加する
  size = UBound(ary) + 1
  ReDim Preserve ary(size)
  Dim vType As Long
  vType = VarType(v)

  Select Case vType
    'オブジェクト、データアクセスオブジェクト、ユーザー定義型は
    'Set構文を利用する
    Case vbObject
      Set ary(size) = v
    Case vbDataObject
      Set ary(size) = v
    Case vbUserDefinedType
      Set ary(size) = v
    Case Else
      ary(size) = v
  End Select
End Sub


Public Function pop() As Variant
  'スタックの一番上のデータを取り出す
  If size = 0 Then
    'スタックにデータが存在しないときは、エラー番号1001をRaiseする
    Err.Raise 1001, "Stack", "スタックにデータが存在しません"
  End If

  Dim vType As Long
  vType = VarType(ary(size))

  Select Case vType
    'オブジェクト、データアクセスオブジェクト、ユーザー定義型は
    'Set構文を利用する
    Case vbObject
      Set pop = ary(size)
    Case vbDataObject
      Set pop = ary(size)
    Case vbUserDefinedType
      Set pop = ary(size)
    Case Else
      pop = ary(size)
  End Select

  size = size - 1
  ReDim Preserve ary(size)

End Function


Public Function count() As Long
  'スタックのデータ数を返す
  count = size

End Function

Public Function getContents() As Variant
  'スタックの内容を配列として返す
  'インデックス1が一番下のデータ
  getContents = ary

End Function


Option Explicit
Sub test_Stack()
  Dim s As Stack
  Set s = New Stack

  s.push 1
  Dim c As Collection
  Set c = New Collection
  c.Add "valueテスト", "keyテスト"
  s.push c
  s.push "a"

  MsgBox "スタックのデータ数は" & s.count & "個です。"

  Dim ary() As Variant
  ary = s.getContents

  MsgBox "スタックから取り出したデータは" & s.pop & "です。"
  Dim returnC As Collection
  Set returnC = s.pop
  MsgBox "スタックから取り出したデータは" & returnC("keyテスト") & "です。"
  MsgBox "スタックから取り出したデータは" & s.pop & "です。"
  MsgBox "スタックから取り出したデータは" & s.pop & "です。"

  Set s = Nothing

End Sub