Attribute VB_Name = "ktHolidayName"

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/　CopyRight(C) K.Tsunoda(AddinBox) 2001 All Rights Reserved.
'_/　( http://www.h3.dion.ne.jp/~sakatsu/index.htm )
'_/
'_/　　この祝日マクロは『kt関数アドイン』で使用しているものです。
'_/　　このロジックは、レスポンスを第一義として、可能な限り少ない
'_/　  【条件判定の実行】で結果を出せるように設計してあります。
'_/　　この関数では、２０１６年施行の改正祝日法(山の日)までを
'_/　  サポートしています。
'_/
'_/　(*1)このマクロを引用するに当たっては、必ずこのコメントも
'_/　　　一緒に引用する事とします。
'_/　(*2)他サイト上で本マクロを直接引用する事は、ご遠慮願います。
'_/　　　【 http://www.h3.dion.ne.jp/~sakatsu/holiday_logic.htm 】
'_/　　　へのリンクによる紹介で対応して下さい。
'_/　(*3)[ktHolidayName]という関数名そのものは、各自の環境に
'_/　　　おける命名規則に沿って変更しても構いません。
'_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Public Function GetNextBusinessDay(d As Date) As Date '{{{
        Do
                d = DateAdd("d", 1, d)
        Loop While IsHoliday(d)
        GetNextBusinessDay = d
End Function '}}}

Public Function GetPreviousBusinessDay(d As Date) As Date '{{{
        Do
                d = DateAdd("d", -1, d)
        Loop While IsHoliday(d)
        GetPreviousBusinessDay = d
End Function '}}}

Public Function IsHoliday(Optional d As Date = "0") As Boolean '{{{
        If d = "0" Then d = Date

        If holidayname(d) <> "" Or (Weekday(d) = 1 Or Weekday(d) = 7) Then
                IsHoliday = True
        Else
                IsHoliday = False
        End If
End Function '}}}

Public Function holidayname(ByVal 日付 As Date) As String '{{{
Dim dtm日付 As Date
Dim str祝日名 As String
Const cst振替休日施行日 As Date = "1973/4/12"

'時刻/時刻誤差の削除(Now関数などへの対応)
    dtm日付 = DateSerial(Year(日付), Month(日付), Day(日付))
    'シリアル値は[±0.5秒]の誤差範囲で認識されます。2002/6/21はシリアル値で
    '[37428.0]ですが､これに[-0.5秒]の誤差が入れば[37427.9999942130]となり､
    'Int関数で整数部分を取り出せば[37427]で前日日付になってしまいます。
    '※ 但し､引数に指定する値が必ず【手入力した日付】等で､時刻や時刻誤差を
    '　 考慮しなくても良いならば､このステップは不要です。引数[日付]をそのまま
    '　 使用しても問題ありません(ほとんどの利用形態ではこちらでしょうが‥‥)。

    str祝日名 = prv祝日(dtm日付)
    If (str祝日名 = "") Then
        If (Weekday(dtm日付) = vbMonday) Then
            ' 月曜以外は振替休日判定不要
            ' 5/6(火,水)の判定は[prv祝日]で処理済
            ' 5/6(月)はここで判定する
            If (dtm日付 >= cst振替休日施行日) Then
                str祝日名 = prv祝日(dtm日付 - 1)
                If (str祝日名 <> "") Then
                    holidayname = "振替休日"
                Else
                    holidayname = ""
                End If
            Else
                holidayname = ""
            End If
        Else
            holidayname = ""
        End If
    Else
        holidayname = str祝日名
    End If
End Function
'}}}

'========================================================================
Private Function prv祝日(ByVal 日付 As Date) As String
  Dim int年 As Integer
  Dim int月 As Integer
  Dim int日 As Integer
  Dim int秋分日 As Integer
  Dim str第N曜日 As String
  ' 時刻データ(小数部)は取り除いてあるので、下記の日付との比較はＯＫ
  Const cst祝日法施行 As Date = "1948/7/20"
  Const cst昭和天皇の大喪の礼 As Date = "1989/2/24"
  Const cst明仁親王の結婚の儀 As Date = "1959/4/10"
  Const cst徳仁親王の結婚の儀 As Date = "1993/6/9"
  Const cst即位礼正殿の儀 As Date = "1990/11/12"

  int年 = Year(日付)
  int月 = Month(日付)
  int日 = Day(日付)

  prv祝日 = ""
  If (日付 < cst祝日法施行) Then
    Exit Function    ' 祝日法施行以前
  End If

  Select Case int月
    Case 1
      If (int日 = 1) Then
        prv祝日 = "元日"
      Else
        If (int年 >= 2000) Then
          str第N曜日 = (((int日 - 1) \ 7) + 1) & Weekday(日付)
          If (str第N曜日 = "22") Then  'Monday:2
            prv祝日 = "成人の日"
          End If
        Else
          If (int日 = 15) Then
            prv祝日 = "成人の日"
          End If
        End If
      End If
    Case 2
      If (int日 = 11) Then
        If (int年 >= 1967) Then
          prv祝日 = "建国記念の日"
        End If
      ElseIf (日付 = cst昭和天皇の大喪の礼) Then
        prv祝日 = "昭和天皇の大喪の礼"
      End If
    Case 3
      If (int日 = prv春分日(int年)) Then  ' 1948～2150以外は[99]
        prv祝日 = "春分の日"            ' が返るので､必ず≠になる
      End If
    Case 4
      If (int日 = 29) Then
        If (int年 >= 2007) Then
          prv祝日 = "昭和の日"
        ElseIf (int年 >= 1989) Then
          prv祝日 = "みどりの日"
        Else
          prv祝日 = "天皇誕生日"
        End If
      ElseIf (日付 = cst明仁親王の結婚の儀) Then
        prv祝日 = "皇太子明仁親王の結婚の儀"
      End If
    Case 5
      If (int日 = 3) Then
        prv祝日 = "憲法記念日"
      ElseIf (int日 = 4) Then
        If (int年 >= 2007) Then
          prv祝日 = "みどりの日"
        ElseIf (int年 >= 1986) Then
          ' 5/4が日曜日は『只の日曜』､月曜日は『憲法記念日の振替休日』(～2006年)
          If (Weekday(日付) > vbMonday) Then
            prv祝日 = "国民の休日"
          End If
        End If
      ElseIf (int日 = 5) Then
        prv祝日 = "こどもの日"
      ElseIf (int日 = 6) Then
        If (int年 >= 2007) Then
          Select Case Weekday(日付)
            Case vbTuesday, vbWednesday
              prv祝日 = "振替休日"    ' [5/3,5/4が日曜]ケースのみ、ここで判定
          End Select
        End If
      End If
    Case 6
      If (日付 = cst徳仁親王の結婚の儀) Then
        prv祝日 = "皇太子徳仁親王の結婚の儀"
      End If
    Case 7
      If (int年 >= 2003) Then
        str第N曜日 = (((int日 - 1) \ 7) + 1) & Weekday(日付)
        If (str第N曜日 = "32") Then  'Monday:2
          prv祝日 = "海の日"
        End If
      ElseIf (int年 >= 1996) Then
        If (int日 = 20) Then
          prv祝日 = "海の日"
        End If
      End If
    Case 8
      If (int日 = 11) Then
        If (int年 >= 2016) Then
          prv祝日 = "山の日"
        End If
      End If
    Case 9
      '第３月曜日(15～21)と秋分日(22～24)が重なる事はない
      int秋分日 = prv秋分日(int年)
      If (int日 = int秋分日) Then  ' 1948～2150以外は[99]
        prv祝日 = "秋分の日"      ' が返るので､必ず≠になる
      Else
        If (int年 >= 2003) Then
          str第N曜日 = (((int日 - 1) \ 7) + 1) & Weekday(日付)
          If (str第N曜日 = "32") Then  'Monday:2
            prv祝日 = "敬老の日"
          ElseIf (Weekday(日付) = vbTuesday) Then
            If (int日 = (int秋分日 - 1)) Then
              prv祝日 = "国民の休日"
            End If
          End If
        ElseIf (int年 >= 1966) Then
          If (int日 = 15) Then
            prv祝日 = "敬老の日"
          End If
        End If
      End If
    Case 10
      If (int年 >= 2000) Then
        str第N曜日 = (((int日 - 1) \ 7) + 1) & Weekday(日付)
        If (str第N曜日 = "22") Then  'Monday:2
          prv祝日 = "体育の日"
        End If
      ElseIf (int年 >= 1966) Then
        If (int日 = 10) Then
          prv祝日 = "体育の日"
        End If
      End If
    Case 11
      If (int日 = 3) Then
        prv祝日 = "文化の日"
      ElseIf (int日 = 23) Then
        prv祝日 = "勤労感謝の日"
      ElseIf (日付 = cst即位礼正殿の儀) Then
        prv祝日 = "即位礼正殿の儀"
      End If
    Case 12
      If (int日 = 23) Then
        If (int年 >= 1989) Then
          prv祝日 = "天皇誕生日"
        End If
      End If
  End Select
End Function

'======================================================================
'　春分/秋分日の略算式は
'　　『海上保安庁水路部 暦計算研究会編 新こよみ便利帳』
'　で紹介されている式です。
Private Function prv春分日(ByVal 年 As Integer) As Integer
  If (年 <= 1947) Then
    prv春分日 = 99        '祝日法施行前
  ElseIf (年 <= 1979) Then
    '(年 - 1983)がマイナスになるので『Fix関数』にする
    prv春分日 = Fix(20.8357 + (0.242194 * (年 - 1980)) - Fix((年 - 1983) / 4))
  ElseIf (年 <= 2099) Then
    prv春分日 = Fix(20.8431 + (0.242194 * (年 - 1980)) - Fix((年 - 1980) / 4))
  ElseIf (年 <= 2150) Then
    prv春分日 = Fix(21.851 + (0.242194 * (年 - 1980)) - Fix((年 - 1980) / 4))
  Else
    prv春分日 = 99        '2151年以降は略算式が無いので不明
  End If
End Function

'========================================================================
Private Function prv秋分日(ByVal 年 As Integer) As Integer
  If (年 <= 1947) Then
    prv秋分日 = 99        '祝日法施行前
  ElseIf (年 <= 1979) Then
    '(年 - 1983)がマイナスになるので『Fix関数』にする
    prv秋分日 = Fix(23.2588 + (0.242194 * (年 - 1980)) - Fix((年 - 1983) / 4))
  ElseIf (年 <= 2099) Then
    prv秋分日 = Fix(23.2488 + (0.242194 * (年 - 1980)) - Fix((年 - 1980) / 4))
  ElseIf (年 <= 2150) Then
    prv秋分日 = Fix(24.2488 + (0.242194 * (年 - 1980)) - Fix((年 - 1980) / 4))
  Else
    prv秋分日 = 99        '2151年以降は略算式が無いので不明
  End If
End Function

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/　CopyRight(C) K.Tsunoda(AddinBox) 2001 All Rights Reserved.
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

