Attribute VB_Name = "keystrokeAsseser"

#If VBA7 And Win64 Then
  Private Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Integer
  Private Declare PtrSafe Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As LongLong
  Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As LongLong 'パフォーマンス計測のため｡
#Else
  Private Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Long
  Private Declare PtrSafe Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
  Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long 'パフォーマンス計測のため｡
#End If

Private Const timeoutLen As Single = 1000 'wait time for hitting next
Private keyStroke As String
Private isNewStroke As Boolean

Private keyMapDic As Object 'Collection of vim_mode_mapping_dictionary
Private visualMap As Object 
Private lin_visualMap As Object 
Private keybinde As String
Private modeOfVim As String
Private s As Double 'for storing time from when previousley pressing a key

Public Sub init()'{{{
  isNewStroke = True
  Set keyMapDic = CreateObject("Scripting.Dictionary")
  Call SetModeOfVim("normal")
End Sub'}}}

Public Sub SetModeOfVim(modeName)'{{{
  modeOfVim = modeName
End Sub'}}}

'----------- Application layer mapping-----------------------
Public Sub AllKeyToAssesKeyFunc()'{{{
    Application.OnKey "a", "AssesKey"
    Application.OnKey "b", "AssesKey"
    Application.OnKey "c", "AssesKey"
    Application.OnKey "d", "AssesKey"
    Application.OnKey "e", "AssesKey"
    Application.OnKey "f", "AssesKey"
    Application.OnKey "g", "AssesKey"
    Application.OnKey "h", "AssesKey"
    Application.OnKey "i", "AssesKey"
    Application.OnKey "j", "AssesKey"
    Application.OnKey "k", "AssesKey"
    Application.OnKey "l", "AssesKey"
    Application.OnKey "m", "AssesKey"
    Application.OnKey "n", "AssesKey"
    Application.OnKey "o", "AssesKey"
    Application.OnKey "p", "AssesKey"
    Application.OnKey "q", "AssesKey"
    Application.OnKey "r", "AssesKey"
    Application.OnKey "s", "AssesKey"
    Application.OnKey "t", "AssesKey"
    Application.OnKey "u", "AssesKey"
    Application.OnKey "v", "AssesKey"
    Application.OnKey "w", "AssesKey"
    Application.OnKey "x", "AssesKey"
    Application.OnKey "y", "AssesKey"
    Application.OnKey "z", "AssesKey"
    
    Application.OnKey "0", "AssesKey"
    Application.OnKey "1", "AssesKey"
    Application.OnKey "2", "AssesKey"
    Application.OnKey "3", "AssesKey"
    Application.OnKey "4", "AssesKey"
    Application.OnKey "5", "AssesKey"
    Application.OnKey "6", "AssesKey"
    Application.OnKey "7", "AssesKey"
    Application.OnKey "8", "AssesKey"
    Application.OnKey "9", "AssesKey"
    
    Application.OnKey "-", "AssesKey"
    Application.OnKey "{^}", "AssesKey"
    Application.OnKey "@", "AssesKey"
    Application.OnKey "{[}", "AssesKey"
    Application.OnKey ";", "AssesKey"
    Application.OnKey ":", "AssesKey"
    Application.OnKey "{]}", "AssesKey"
    Application.OnKey ",", "AssesKey"
    Application.OnKey ".", "AssesKey"
    Application.OnKey "/", "AssesKey"
    Application.OnKey "=", "AssesKey"
    Application.OnKey "{+}", "AssesKey"
    Application.OnKey ">", "AssesKey"
    Application.OnKey "<", "AssesKey"
    Application.OnKey "?", "AssesKey"
    Application.OnKey "|", "AssesKey"
    Application.OnKey "'", "AssesKey"
    Application.OnKey "*", "AssesKey"
    Application.OnKey "{{}", "AssesKey"
    Application.OnKey "{}}", "AssesKey"
    Application.OnKey "{(}", "AssesKey"
    Application.OnKey "{)}", "AssesKey"
    Application.OnKey "!", "AssesKey"
    Application.OnKey "#", "AssesKey"

    Application.OnKey "+{a}", "AssesKey"
    Application.OnKey "+{b}", "AssesKey"
    Application.OnKey "+{c}", "AssesKey"
    Application.OnKey "+{d}", "AssesKey"
    Application.OnKey "+{e}", "AssesKey"
    Application.OnKey "+{f}", "AssesKey"
    Application.OnKey "+{g}", "AssesKey"
    Application.OnKey "+{h}", "AssesKey"
    Application.OnKey "+{i}", "AssesKey"
    Application.OnKey "+{j}", "AssesKey"
    Application.OnKey "+{k}", "AssesKey"
    Application.OnKey "+{l}", "AssesKey"
    Application.OnKey "+{m}", "AssesKey"
    Application.OnKey "+{n}", "AssesKey"
    Application.OnKey "+{o}", "AssesKey"
    Application.OnKey "+{p}", "AssesKey"
    Application.OnKey "+{q}", "AssesKey"
    Application.OnKey "+{r}", "AssesKey"
    Application.OnKey "+{s}", "AssesKey"
    Application.OnKey "+{t}", "AssesKey"
    Application.OnKey "+{u}", "AssesKey"
    Application.OnKey "+{v}", "AssesKey"
    Application.OnKey "+{w}", "AssesKey"
    Application.OnKey "+{x}", "AssesKey"
    Application.OnKey "+{y}", "AssesKey"
    Application.OnKey "+{z}", "AssesKey"
    Application.OnKey "+0", "AssesKey"
    Application.OnKey "+1", "AssesKey"
    Application.OnKey "+2", "AssesKey"
    Application.OnKey "+3", "AssesKey"
    Application.OnKey "+4", "AssesKey"
    Application.OnKey "+5", "AssesKey"
    Application.OnKey "+6", "AssesKey"
    Application.OnKey "+7", "AssesKey"
    Application.OnKey "+8", "AssesKey"
    Application.OnKey "+9", "AssesKey"

    Application.OnKey "^{a}", "AssesKey"
    Application.OnKey "^{b}", "AssesKey"
    Application.OnKey "^{c}", "AssesKey"
    Application.OnKey "^{d}", "AssesKey"
    Application.OnKey "^{e}", "AssesKey"
    Application.OnKey "^{f}", "AssesKey"
    Application.OnKey "^{g}", "AssesKey"
    Application.OnKey "^{h}", "AssesKey"
    Application.OnKey "^{i}", "AssesKey"
    Application.OnKey "^{j}", "AssesKey"
    Application.OnKey "^{k}", "AssesKey"
    Application.OnKey "^{l}", "AssesKey"
    Application.OnKey "^{m}", "AssesKey"
    Application.OnKey "^{n}", "AssesKey"
    Application.OnKey "^{o}", "AssesKey"
    Application.OnKey "^{p}", "AssesKey"
    Application.OnKey "^{q}", "AssesKey"
    Application.OnKey "^{r}", "AssesKey"
    Application.OnKey "^{s}", "AssesKey"
    Application.OnKey "^{t}", "AssesKey"
    Application.OnKey "^{u}", "AssesKey"
    Application.OnKey "^{v}", "AssesKey"
    Application.OnKey "^{w}", "AssesKey"
    Application.OnKey "^{x}", "AssesKey"
    Application.OnKey "^{y}", "AssesKey"
    Application.OnKey "^{z}", "AssesKey"
    Application.OnKey "^0", "AssesKey"
    Application.OnKey "^1", "AssesKey"
    Application.OnKey "^2", "AssesKey"
    Application.OnKey "^3", "AssesKey"
    Application.OnKey "^4", "AssesKey"
    Application.OnKey "^5", "AssesKey"
    Application.OnKey "^6", "AssesKey"
    Application.OnKey "^7", "AssesKey"
    Application.OnKey "^8", "AssesKey"
    Application.OnKey "^9", "AssesKey"

    Application.OnKey "{F1}", "AssesKey"
    Application.OnKey "{F2}", "AssesKey"
    Application.OnKey "{F3}", "AssesKey"
    Application.OnKey "{F4}", "AssesKey"
    Application.OnKey "{F5}", "AssesKey"
    Application.OnKey "{F6}", "AssesKey"
    Application.OnKey "{F7}", "AssesKey"
    Application.OnKey "{F8}", "AssesKey"
    Application.OnKey "{F9}", "AssesKey"
    Application.OnKey "{F10}", "AssesKey"
    Application.OnKey "{F11}", "AssesKey"
    Application.OnKey "{F12}", "AssesKey"
    Application.OnKey "{F13}", "AssesKey"
    Application.OnKey "{F14}", "AssesKey"
    Application.OnKey "{F15}", "AssesKey"
    Application.OnKey "{F16}", "AssesKey"
    Application.OnKey "{ESC}", "AssesKey"
End Sub'}}}

Public Sub AllKeyAssign_reset()'{{{
    Application.OnKey "a"
    Application.OnKey "b"
    Application.OnKey "c"
    Application.OnKey "d"
    Application.OnKey "e"
    Application.OnKey "f"
    Application.OnKey "g"
    Application.OnKey "h"
    Application.OnKey "i"
    Application.OnKey "j"
    Application.OnKey "k"
    Application.OnKey "l"
    Application.OnKey "m"
    Application.OnKey "n"
    Application.OnKey "o"
    Application.OnKey "p"
    Application.OnKey "q"
    Application.OnKey "r"
    Application.OnKey "s"
    Application.OnKey "t"
    Application.OnKey "u"
    Application.OnKey "v"
    Application.OnKey "w"
    Application.OnKey "x"
    Application.OnKey "y"
    Application.OnKey "z"
    
    Application.OnKey "0"
    Application.OnKey "1"
    Application.OnKey "2"
    Application.OnKey "3"
    Application.OnKey "4"
    Application.OnKey "5"
    Application.OnKey "6"
    Application.OnKey "7"
    Application.OnKey "8"
    Application.OnKey "9"
    
    Application.OnKey "="
    Application.OnKey "-"
    Application.OnKey "{^}"
    Application.OnKey "?"
    Application.OnKey "@"
    Application.OnKey "{[}"
    Application.OnKey ";"
    Application.OnKey ":"
    Application.OnKey "{]}"
    Application.OnKey "."
    
    Application.OnKey "+a"
    Application.OnKey "+b"
    Application.OnKey "+c"
    Application.OnKey "+d"
    Application.OnKey "+e"
    Application.OnKey "+f"
    Application.OnKey "+g"
    Application.OnKey "+h"
    Application.OnKey "+i"
    Application.OnKey "+j"
    Application.OnKey "+k"
    Application.OnKey "+l"
    Application.OnKey "+m"
    Application.OnKey "+n"
    Application.OnKey "+o"
    Application.OnKey "+p"
    Application.OnKey "+q"
    Application.OnKey "+r"
    Application.OnKey "+s"
    Application.OnKey "+t"
    Application.OnKey "+u"
    Application.OnKey "+v"
    Application.OnKey "+w"
    Application.OnKey "+x"
    Application.OnKey "+y"
    Application.OnKey "+z"
    
    Application.OnKey "+0"
    Application.OnKey "+1"
    Application.OnKey "+2"
    Application.OnKey "+3"
    Application.OnKey "+4"
    Application.OnKey "+5"
    Application.OnKey "+6"
    Application.OnKey "+7"
    Application.OnKey "+8"
    Application.OnKey "+9"
    
    Application.OnKey "+-"
    Application.OnKey "+{^}"
    Application.OnKey "+?"
    Application.OnKey "+@"
    Application.OnKey "+{[}"
    Application.OnKey "+;"
    Application.OnKey "+:"
    Application.OnKey "+{]}"
    Application.OnKey "<"
    Application.OnKey "+."
    Application.OnKey "+/"
    Application.OnKey "_"
    
    'Ctrl
    Application.OnKey "^a"
    Application.OnKey "^b"
    Application.OnKey "^c"
    Application.OnKey "^d"
    Application.OnKey "^e"
    Application.OnKey "^f"
    Application.OnKey "^g"
    Application.OnKey "^h"
    Application.OnKey "^i"
    Application.OnKey "^j"
    Application.OnKey "^k"
    Application.OnKey "^l"
    Application.OnKey "^m"
    Application.OnKey "^n"
    Application.OnKey "^o"
    Application.OnKey "^p"
    Application.OnKey "^q"
    Application.OnKey "^r"
    Application.OnKey "^s"
    Application.OnKey "^t"
    Application.OnKey "^u"
    Application.OnKey "^v"
    Application.OnKey "^w"
    Application.OnKey "^x"
    Application.OnKey "^y"
    Application.OnKey "^z"
    
    Application.OnKey "^0"
    Application.OnKey "^1"
    Application.OnKey "^2"
    Application.OnKey "^3"
    Application.OnKey "^4"
    Application.OnKey "^5"
    Application.OnKey "^6"
    Application.OnKey "^7"
    Application.OnKey "^8"
    Application.OnKey "^9"
    
    Application.OnKey "^-"
    Application.OnKey "^{^}"
    Application.OnKey "^?"
    Application.OnKey "^@"
    Application.OnKey "^{[}"
    Application.OnKey "^;"
    Application.OnKey "^:"
    Application.OnKey "^{]}"
    Application.OnKey "^."
    
    Application.OnKey "^+a"
    Application.OnKey "^+b"
    Application.OnKey "^+c"
    Application.OnKey "^+d"
    Application.OnKey "^+e"
    Application.OnKey "^+f"
    Application.OnKey "^+g"
    Application.OnKey "^+h"
    Application.OnKey "^+i"
    Application.OnKey "^+j"
    Application.OnKey "^+k"
    Application.OnKey "^+l"
    Application.OnKey "^+m"
    Application.OnKey "^+n"
    Application.OnKey "^+o"
    Application.OnKey "^+p"
    Application.OnKey "^+q"
    Application.OnKey "^+r"
    Application.OnKey "^+s"
    Application.OnKey "^+t"
    Application.OnKey "^+u"
    Application.OnKey "^+v"
    Application.OnKey "^+w"
    Application.OnKey "^+x"
    Application.OnKey "^+y"
    Application.OnKey "^+z"
    
    Application.OnKey "^+0"
    Application.OnKey "^+1"
    Application.OnKey "^+2"
    Application.OnKey "^+3"
    Application.OnKey "^+4"
    Application.OnKey "^+5"
    Application.OnKey "^+6"
    Application.OnKey "^+7"
    Application.OnKey "^+8"
    Application.OnKey "^+9"
    
    Application.OnKey "^+-"
    Application.OnKey "^+{^}"
    Application.OnKey "^+?"
    Application.OnKey "^+@"
    Application.OnKey "^+{[}"
    Application.OnKey "^+;"
    Application.OnKey "^+:"
    Application.OnKey "^+{]}"
    Application.OnKey "^<"
    Application.OnKey "^+."
    Application.OnKey "^+/"
    Application.OnKey "^_"

    Application.OnKey "{F1}"
    Application.OnKey "{F2}"
    ' Application.OnKey "{F3}"
    Application.OnKey "{F4}"
    Application.OnKey "{F5}"
    Application.OnKey "{F6}"
    Application.OnKey "{F7}"
    Application.OnKey "{F8}"
    Application.OnKey "{F9}"
    Application.OnKey "{F10}"
    Application.OnKey "{F11}"
    Application.OnKey "{F12}"
    Application.OnKey "{F13}"
    Application.OnKey "{F14}"
    Application.OnKey "{F15}"
    Application.OnKey "{F16}"
End Sub '}}}

'----------- mapping def function -----------------------
Public Sub nmap(key, func, optional context = "default")'{{{
  if not keyMapDic.exists(context) then
    CreateMap(context)
  end if
  keyMapDic(context)("normal")(key) = func
End Sub'}}}

Public Sub vmap(key, func, optional context = "default")'{{{
  if not keyMapDic.exists(context) then
    CreateMap(context)
  end if
  keyMapDic(context)("visual")(key) = func
End Sub'}}}

Public Sub lvmap(key, func, optional context = "default")'{{{
  if not keyMapDic.exists(context) then
    CreateMap(context)
  end if
  keyMapDic(context)("line_visual")(key) = func
End Sub'}}}

Private Sub CreateMap(context)'{{{
  Dim tmp As Object
  Set tmp = CreateObject("Scripting.Dictionary")
  Set normalMap = CreateObject("Scripting.Dictionary")
  Set visualMap = CreateObject("Scripting.Dictionary")
  Set lin_visualMap = CreateObject("Scripting.Dictionary")
  tmp.Add "normal", normalMap
  tmp.Add "visual", visualMap
  tmp.Add "line_visual", lin_visualMap
  keyMapDic.Add context, tmp
End Sub'}}}

'----------- executer-----------------------
Private Sub AssesKey(optional context As String = "default")'{{{
  'この関数はキーによって呼び出され,実行すべき処理を判定します｡
  'TODO strokkeyはkeystringの順番によって連続で解釈出来るものとそうでないものが出来てしまっている。
  'TODO contextの処理

  Application.EnableCancelKey = xlDisabled 'for Esc Command. Without this, cannot catch ESC key.
  '
  If keyMapDic is Nothing Then 
    Application.Run("keystrokeAsseser.init")
    Application.Run("configure.init")
    On Error GoTo except
      Application.Run("user_configure.init")
    except:
      If Err.Number <> 0 Then
        Debug.print Err.Description
      End If
  End If

  s = GetTickCount '0 milisecond

  'Get put key
  If isNewStroke = True Then
    keyStroke = ""
    newkey = GetKeyString '新規の場合は､GetKeyboardStateを使う。こちらの関数でないと､のどかのmodifierkeyの影響を受けてしまう｡
  Else
    newkey = GetKeyStringAsync 'GetKeyboardStateを使うと前のキーの情報が残ってしまっている事があるためこちらを使う｡
  End If

  If newkey = "" Then 'When Application.OnKey Works, but GetKeyString does not work.'{{{
    MsgBox "couldn't get newkey"
    isNewStroke = True
    Exit Sub
  End If'}}}

  keyStroke = keyStroke + newkey

  'keyStrokeを評価
  ' Debug.print KeyStroke & "を評価します"
  candidate = NumberOfHits(keyStroke, context, modeOfVim)
  If candidate = 0 Then 'keyStrokeにヒットが0件
    ' Debug.Print "ヒットが0件"
    isNewStroke = True
    Exit Sub
  ElseIf candidate = 1 And keyMapDic(context)(modeOfVim).Exists(keyStroke) Then '候補が一意かつヒットしている時
    ' Debug.print "候補が一意かつヒットしている時のAssesKeyの(関数呼び出しまでの)実行時間は" & GetTickCount - st & "ミリセカンド"
    ' Debug.Print keyMapDic(context)(modeOfVim)(keyStroke) & "をkeystrokeから呼び出し"
    Call ExeStringPro(keyMapDic(context)(modeOfVim).Item(keyStroke))
    isNewStroke = True
    ' Debug.Print "poformanace time is " & GetTickCount - s
    Exit Sub
  Else ' Debug.print "候補が複数の時"
    isNewStroke = False
    e = GetTickCount

    'wait next input key.
    Do until e-s > timeoutLen
      key = GetKeyStringAsync '(* GetKeyStringAsync returns "", when nothing is being pressed)
      if key = "" Then '次のキーが押される前に､前のキーが離された場合｡
'        Debug.print "最初のキーが離れました｡"
        Exit Do
      End if

      if key <> "" And key <> newkey Then '前のキーが離される前に次のキーが押された場合
        ' Debug.print "最初のキーが離されないままに､別のキー"& key &"が連続で押されました"
        'AssesKeyCore(key) これを実行せずともApplication.onkeyによって次のAssesKeyが呼ばれる｡
        Exit Sub
      End if
      e = GetTickCount
    Loop

    '最初のキーが離れてからの監視体制
    Do until e-s > timeoutLen
      key = GetKeyStringAsync
      if key <> "" Then
        Exit Sub 
      End if
      e = GetTickCount
    Loop

    ' Debug.print "have waited for timeoutlen:" & timeoutlen & ", so will execute the stroke:" & KeyStroke
    Application.Run keyMapDic(context)(modeOfVim)(keyStroke)
    isNewStroke = True
  End If
End Sub
'}}}

'-----------supplimental functions-----------------------
Private Function GetKeyStringAsync()'{{{
  '関数実行時点で押されているキーを判別して返します｡
  'shift'{{{
  shift = False
    If GetAsyncKeyState(16) <> 0 Then shift = True '}}} 'なぜか<0だと検知しない｡

  'control'{{{
  control = False
    If GetAsyncKeyState(17) <> 0 Then control = True'}}}

  'mainkey'{{{
  mainkey = ""
  'alphabet'{{{
    If GetAsyncKeyState(65) < 0 Then mainkey = "a"
    If GetAsyncKeyState(66) < 0 Then mainkey = "b"
    If GetAsyncKeyState(67) < 0 Then mainkey = "c"
    If GetAsyncKeyState(68) < 0 Then mainkey = "d"
    If GetAsyncKeyState(69) < 0 Then mainkey = "e"
    If GetAsyncKeyState(70) < 0 Then mainkey = "f"
    If GetAsyncKeyState(71) < 0 Then mainkey = "g"
    If GetAsyncKeyState(72) < 0 Then mainkey = "h"
    If GetAsyncKeyState(73) < 0 Then mainkey = "i"
    If GetAsyncKeyState(74) < 0 Then mainkey = "j"
    If GetAsyncKeyState(75) < 0 Then mainkey = "k"
    If GetAsyncKeyState(76) < 0 Then mainkey = "l"
    If GetAsyncKeyState(77) < 0 Then mainkey = "m"
    If GetAsyncKeyState(78) < 0 Then mainkey = "n"
    If GetAsyncKeyState(79) < 0 Then mainkey = "o" 'なぜか
    If GetAsyncKeyState(80) < 0 Then mainkey = "p"
    If GetAsyncKeyState(81) < 0 Then mainkey = "q"
    If GetAsyncKeyState(82) < 0 Then mainkey = "r"
    If GetAsyncKeyState(83) < 0 Then mainkey = "s"
    If GetAsyncKeyState(84) < 0 Then mainkey = "t"
    If GetAsyncKeyState(85) < 0 Then mainkey = "u" 'なぜか
    If GetAsyncKeyState(86) < 0 Then mainkey = "v"
    If GetAsyncKeyState(87) < 0 Then mainkey = "w"
    If GetAsyncKeyState(88) < 0 Then mainkey = "x"
    If GetAsyncKeyState(89) < 0 Then mainkey = "y"
    If GetAsyncKeyState(90) < 0 Then mainkey = "z"'}}}
  'number'{{{
    If GetAsyncKeyState(48) < 0 Then mainkey = "0"
    If GetAsyncKeyState(49) < 0 Then mainkey = "1"
    If GetAsyncKeyState(50) < 0 Then mainkey = "2"
    If GetAsyncKeyState(51) < 0 Then mainkey = "3"
    If GetAsyncKeyState(52) < 0 Then mainkey = "4"
    If GetAsyncKeyState(53) < 0 Then mainkey = "5"
    If GetAsyncKeyState(54) < 0 Then mainkey = "6"
    If GetAsyncKeyState(55) < 0 Then mainkey = "7"
    If GetAsyncKeyState(56) < 0 Then mainkey = "8"
    If GetAsyncKeyState(57) < 0 Then mainkey = "9"'}}}
  'symbol'{{{
  If GetAsyncKeyState(186) < 0 Then mainkey = ":"
    If GetAsyncKeyState(187) < 0 Then mainkey = ";"
    If GetAsyncKeyState(188) < 0 Then mainkey = ","
    If GetAsyncKeyState(189) < 0 Then mainkey = "-"
    If GetAsyncKeyState(190) < 0 Then mainkey = "."
    If GetAsyncKeyState(191) < 0 Then mainkey = "/"
    If GetAsyncKeyState(192) < 0 Then mainkey = "@"
    If GetAsyncKeyState(219) < 0 Then mainkey = "["
    If GetAsyncKeyState(220) < 0 Then mainkey = "\"
    If GetAsyncKeyState(221) < 0 Then mainkey = "]"
    If GetAsyncKeyState(222) < 0 Then mainkey = "^"'}}}
  'others'{{{
  If GetAsyncKeyState(23) < 0 Then mainkey = "<END>"
  If GetAsyncKeyState(vbKeyEscape) < 0 Then mainkey = "<ESC>"
    If GetAsyncKeyState(24) < 0 Then mainkey = "<HOME>"'}}}
  'Function key'{{{
    If GetAsyncKeyState(112) < 0 Then mainkey = "F1"
    If GetAsyncKeyState(113) < 0 Then mainkey = "F2"
    If GetAsyncKeyState(114) < 0 Then mainkey = "F3"
    If GetAsyncKeyState(115) < 0 Then mainkey = "F4"
    If GetAsyncKeyState(116) < 0 Then mainkey = "F5"
    If GetAsyncKeyState(117) < 0 Then mainkey = "F6"
    If GetAsyncKeyState(118) < 0 Then mainkey = "F7"
    If GetAsyncKeyState(119) < 0 Then mainkey = "F8"
    If GetAsyncKeyState(120) < 0 Then mainkey = "F9"
    If GetAsyncKeyState(121) < 0 Then mainkey = "F10"
    'If GetAsyncKeyState(122) < 0 Then mainkey = "F11" 'なぜかF11が発動する事があるので､上書かれるように上に←VBE起動キーがF11
    If GetAsyncKeyState(123) < 0 Then mainkey = "F12"
    If GetAsyncKeyState(124) < 0 Then mainkey = "F13"
    If GetAsyncKeyState(125) < 0 Then mainkey = "F14"
    If GetAsyncKeyState(126) < 0 Then mainkey = "F15"
    If GetAsyncKeyState(127) < 0 Then mainkey = "F16"
'}}}'}}}

  '返り値をセット'{{{
  GetkeyStringAsync = ""
         'Debug.print "mainkey" & mainkey
  If shift Then
    GetKeyStringAsync = UCase(mainkey)
  ElseIf control Then
    GetKeyStringAsync = "<c-" & mainkey & ">"
  Else
    GetKeyStringAsync = mainkey
  End If'}}}
'  'Debug.print "GetKeyStringの実行時間は" & GetTickCount - s & "ミリセカンド"
End Function'}}}

Private Function GetKeyString()'{{{
'nodokaでmodifierkeyなどになっているキーは､Asyncでは取得出来ないためこちらで取得
  'keyboardの状態をstateにセット'{{{
  Dim state(255) As Byte
  Call GetKeyboardState(state(0))
  'http://www.yoshidastyle.net/2007/10/windowswin32api.html
  ' For i = 0 to 255
  '   if state(i) <> 0 And state(i) <> 1 Then
'  '     Debug.Print "仮想キーコード" & i "の状態は" & state(i)
  '   End If
  ' Next i'}}}

  'shiftキーの判定'{{{
  Dim shift As boolean
  shift = False
  shift = state(16) >= 128'}}}

  'controlキーの判定'{{{
  Dim control As boolean
  control = False
  control = state(17) >= 128'}}}

  'mainkeyキーの取得'{{{
  Dim mainkey As String : mainkey = ""
  'mainkey
  If shift Then
    'number
    If state(49) >= 128 Then mainkey = "!"
    If state(50) >= 128 Then mainkey = """
    If state(51) >= 128 Then mainkey = "#"
    If state(52) >= 128 Then mainkey = "$"
    If state(53) >= 128 Then mainkey = "%"
    If state(54) >= 128 Then mainkey = "&"
    If state(55) >= 128 Then mainkey = "'"
    If state(56) >= 128 Then mainkey = "("
    If state(57) >= 128 Then mainkey = ")"
    'alphabet
    If state(65) >= 128 Then mainkey = "A"
    If state(66) >= 128 Then mainkey = "B"
    If state(67) >= 128 Then mainkey = "C"
    If state(68) >= 128 Then mainkey = "D"
    If state(69) >= 128 Then mainkey = "E"
    If state(70) >= 128 Then mainkey = "F"
    If state(71) >= 128 Then mainkey = "G"
    If state(72) >= 128 Then mainkey = "H"
    If state(73) >= 128 Then mainkey = "I"
    If state(74) >= 128 Then mainkey = "J"
    If state(75) >= 128 Then mainkey = "K"
    If state(76) >= 128 Then mainkey = "L"
    If state(77) >= 128 Then mainkey = "M"
    If state(78) >= 128 Then mainkey = "N"
    If state(79) >= 128 Then mainkey = "O"
    If state(80) >= 128 Then mainkey = "P"
    If state(81) >= 128 Then mainkey = "Q"
    If state(82) >= 128 Then mainkey = "R"
    If state(83) >= 128 Then mainkey = "S"
    If state(84) >= 128 Then mainkey = "T"
    If state(85) >= 128 Then mainkey = "U"
    If state(86) >= 128 Then mainkey = "V"
    If state(87) >= 128 Then mainkey = "W"
    If state(88) >= 128 Then mainkey = "X"
    If state(89) >= 128 Then mainkey = "Y"
    If state(90) >= 128 Then mainkey = "Z"
    'symbol
    If state(186) >= 128 Then mainkey = "*"
    If state(187) >= 128 Then mainkey = "+"
    If state(188) >= 128 Then mainkey = "<<"
    If state(189) >= 128 Then mainkey = "="
    If state(190) >= 128 Then mainkey = ">"
    If state(191) >= 128 Then mainkey = "?"
    If state(192) >= 128 Then mainkey = "`"
    If state(219) >= 128 Then mainkey = "{"
    If state(220) >= 128 Then mainkey = "|"
    If state(221) >= 128 Then mainkey = "}"
    If state(222) >= 128 Then mainkey = "~"
  Else
    If state(48) >= 128 Then mainkey = "0"
    If state(49) >= 128 Then mainkey = "1"
    If state(50) >= 128 Then mainkey = "2"
    If state(51) >= 128 Then mainkey = "3"
    If state(52) >= 128 Then mainkey = "4"
    If state(53) >= 128 Then mainkey = "5"
    If state(54) >= 128 Then mainkey = "6"
    If state(55) >= 128 Then mainkey = "7"
    If state(56) >= 128 Then mainkey = "8"
    If state(57) >= 128 Then mainkey = "9"
    'alphabet
    If state(86) >= 128 Then mainkey = "v" 'visual_mode直後からの移動キーをスムーズにするため先頭に｡
    If state(65) >= 128 Then mainkey = "a"
    If state(66) >= 128 Then mainkey = "b"
    If state(67) >= 128 Then mainkey = "c"
    If state(68) >= 128 Then mainkey = "d"
    If state(69) >= 128 Then mainkey = "e"
    If state(70) >= 128 Then mainkey = "f"
    If state(71) >= 128 Then mainkey = "g"
    If state(72) >= 128 Then mainkey = "h"
    If state(73) >= 128 Then mainkey = "i"
    If state(74) >= 128 Then mainkey = "j"
    If state(75) >= 128 Then mainkey = "k"
    If state(76) >= 128 Then mainkey = "l"
    If state(77) >= 128 Then mainkey = "m"
    If state(78) >= 128 Then mainkey = "n"
    If state(79) >= 128 Then mainkey = "o"
    If state(80) >= 128 Then mainkey = "p"
    If state(81) >= 128 Then mainkey = "q"
    If state(82) >= 128 Then mainkey = "r"
    If state(83) >= 128 Then mainkey = "s"
    If state(84) >= 128 Then mainkey = "t"
    If state(85) >= 128 Then mainkey = "u"
    If state(87) >= 128 Then mainkey = "w"
    If state(88) >= 128 Then mainkey = "x"
    If state(89) >= 128 Then mainkey = "y"
    If state(90) >= 128 Then mainkey = "z"
    'symbol
    If state(186) >= 128 Then mainkey = ":"
    If state(187) >= 128 Then mainkey = ";"
    If state(188) >= 128 Then mainkey = ","
    If state(189) >= 128 Then mainkey = "-"
    If state(190) >= 128 Then mainkey = "."
    If state(191) >= 128 Then mainkey = "/"
    If state(192) >= 128 Then mainkey = "@"
    If state(219) >= 128 Then mainkey = "["
    If state(220) >= 128 Then mainkey = "\"
    If state(221) >= 128 Then mainkey = "]"
    If state(222) >= 128 Then mainkey = "^"
    'others
    If state(23) >= 128 Then mainkey = "<END>"
    If state(24) >= 128 Then mainkey = "<HOME>"
    If state(vbKeyEscape) >= 128 Then mainkey = "<ESC>"
  End If

  'Function key'{{{
    If state(112) >= 128 Then mainkey = "F1"
    If state(113) >= 128 Then mainkey = "F2"
    If state(114) >= 128 Then mainkey = "F3"
    If state(115) >= 128 Then mainkey = "F4"
    If state(116) >= 128 Then mainkey = "F5"
    If state(117) >= 128 Then mainkey = "F6"
    If state(118) >= 128 Then mainkey = "F7"
    If state(119) >= 128 Then mainkey = "F8"
    If state(120) >= 128 Then mainkey = "F9"
    If state(121) >= 128 Then mainkey = "F10"
    'If state(122) >= 128 Then mainkey = "F11" 'なぜかF11が発動する事があるので､上書かれるように上に←VBE起動キーがF11
    If state(123) >= 128 Then mainkey = "F12"
    If state(124) >= 128 Then mainkey = "F13"
    If state(125) >= 128 Then mainkey = "F14"
    If state(126) >= 128 Then mainkey = "F15"
    If state(127) >= 128 Then mainkey = "F16"
'}}}
'}}}

  '返り値にセット'{{{
  If control Then
    GetKeyString = "<c-" & mainkey & ">"
  Else
    GetKeyString = mainkey
  End If'}}}

End Function'}}}

Private Function NumberOfHits(stroke As String, context, modeOfVim) As Long'{{{
  'keyMapDicの中で､keystrokeに前方一致する項目の数を返す
  s = GetTickCount

  c = 0
  keyList = keyMapDic(context)(modeOfVim).Keys
  For i = 0 To UBound(keyList)
    If InStr(keyList(i), stroke) = 1 Then
      c = c + 1
    End If
  Next i
  NumberOfHits = c

  '  ' Debug.print "NumberOfHitsの実行時間は" & GetTickCount - s & "ミリセカンド"
End Function'}}}

