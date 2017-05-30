VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UniteInterface 
   Caption         =   "Unite"
   ClientHeight    =   7185
   ClientLeft      =   48
   ClientTop       =   376
   ClientWidth     =   15328
   OleObjectBlob   =   "UniteInterface.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UniteInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-----------------Below this line , Write down code-----------------------------------
Private Sub UserForm_Initialize() '{{{
        ListBox1.MultiSelect = fmMultiSelectMulti
        ' ListBox1.MultiSelect = fmMultiSelectExtended
        ' ListBox1.ListStyle = fmListStyleOption
        Me.ListBox1.Clear
        For Each Buf In UniteCandidatesList
                Me.ListBox1.AddItem Split(Buf, ":::")(0)
        Next Buf

        TextBox1.SetFocus
End Sub '}}}

Private Sub TextBox1_Change() '{{{
        Set RE = CreateObject("VBScript.RegExp") 'https://msdn.microsoft.com/ja-jp/library/cc392437.aspx
        RE.IgnoreCase = True
        patternlist = Split(Replace(Me.TextBox1, "　", " "), " ")
        'リストボックスの内容を初期化
        Me.ListBox1.Clear
        'GatherCandidateで集めたリストをパターンマッチング
        For Each Buf In UniteCandidatesList
                hit = True
                Buf = Split(Buf, ":::")(0)
                'patternに対してテストを繰り返す｡
                For i = 0 To UBound(patternlist)
                        RE.pattern = patternlist(i)
                        'migemo version. too late
                        'Dim buf2 As String: buf2 = patternlist(i)
                        'RE.pattern = migemize(buf2) 'migemo version
                        If Not RE.test(Buf) Then
                                hit = False
                        End If
                Next i
                If hit Then
                        Me.ListBox1.AddItem Buf
                End If
        Next Buf
End Sub '}}}

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal shift As Integer) '{{{
        If KeyCode = 27 Then 'ESC時の挙動
                ' If TextBox1.Text = "" Or Me.ListBox1.ListCount = 0 Then
                If Me.ListBox1.ListCount = 0 Then
                        ' If TextBox1.Text = "" Then
                        Unload Me
                Else
                        ListBox1.SetFocus
                        Me.ListBox1.ListIndex = 0
                End If
        End If

        If KeyCode = 40 Then '↓時の挙動
                ListBox1.SetFocus
                Me.ListBox1.ListIndex = 0
        End If
End Sub '}}}

Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal shift As Integer) '{{{
        ' If KeyCode = 13 And Not ListBox1.Text = "" Then 'Enter時の挙動
        If KeyCode = 13 Then 'Enter時の挙動
                Me.Hide
                With Me.ListBox1
                        selectCount = 0
                        Dim selected As String
                        For i = 0 To .ListCount - 1
                                If .selected(i) = True Then
                                        selected = selected & .List(i) & vbCrlF
                                        selectCount = selectCount + 1
                                End If
                        Next i

                        If selectCount > 0 Then
                                selected = Left(selected, Len(selected) - 2) '末尾のvbCrLfを削除
                        Else
                                selected = .List(.ListIndex)
                        End If

                        Unload Me
                        Call Application.run("defaultAction_" & unite_source, selected)
                End With
        End If

        'http://www.accessclub.jp/samplefile/help/help_154_1.htm keycode
        If KeyCode = 27 Then 'ESC時の挙動
                Unload Me
        End If

        If KeyCode = vbKeyA Then 'a
                Me.TextBox1.SetFocus
        End If

        If KeyCode = vbKeyI Then 'i
                Me.TextBox1.SetFocus
        End If

        If KeyCode = 191 Then '/
                Me.TextBox1.SetFocus
        End If

        If KeyCode = vbKeyK Then 'k
                sendkeys "{UP}"
        End If

        If KeyCode = vbKeyJ Then 'j
                sendkeys "{DOWN}"
        End If

        If KeyCode = vbKeyF Then 'i
                sendkeys " "
                sendkeys "{DOWN}"
        End If

        If KeyCode = vbKeyY Then 'y
                Me.Hide
                With New MSForms.DataObject
                        .SetText Me.ListBox1.List(Me.ListBox1.ListIndex) '変数の値をDataObjectに格納する
                        .PutInClipboard   'DataObjectのデータをクリップボードに格納する
                End With
                Unload Me
        End If

        If KeyCode = vbKeyTab Or KeyCode = 186 Then 'tab or colon: commnad box
                With Me.ListBox1
                        selectCount = 0
                        For i = 0 To .ListCount - 1
                                If .selected(i) = True Then
                                        unite_argument = unite_argument & .List(i) & vbCrlF
                                        selectCount = selectCount + 1
                                End If
                        Next i

                        If selectCount > 0 Then
                                unite_argument = Left(unite_argument, Len(unite_argument) - 2) '末尾のvbCrLfを削除
                        Else
                                unite_argument = .List(.ListIndex)
                        End If

                        Set UniteCandidatesList = Application.run("GatherCandidates_command")
                        unite_source = "command_parent"
                        Unload Me
                        Set commandForm = New UniteInterface
                        commandForm.Show
                End With
        End If

End Sub '}}}

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '{{{
        Me.Hide
        With Me.ListBox1
                selectCount = 0
                Dim selected As String
                For i = 0 To .ListCount - 1
                        If .selected(i) = True Then
                                selected = selected & .List(i) & vbCrlF
                                selectCount = selectCount + 1
                        End If
                Next i

                If selectCount > 0 Then
                        selected = Left(selected, Len(selected) - 2) '末尾のvbCrLfを削除
                Else
                        selected = .List(.ListIndex)
                End If

                Call Application.run("defaultAction_" & unite_source, selected)
        End With
        Unload Me
End Sub '}}}

