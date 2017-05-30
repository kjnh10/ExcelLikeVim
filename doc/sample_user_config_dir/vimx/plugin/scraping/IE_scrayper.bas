Attribute VB_Name = "IE_scrayper"
#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

'----------------------------------------------------------------
'①指定URLを表示するサブルーチン「ieView」
Sub ieView(objIE As InternetExplorer, _
           urlName As String, _
           Optional viewFlg As Boolean = True)

  'IE(InternetExplorer)のオブジェクトを作成する
  Set objIE = CreateObject("InternetExplorer.Application")

  'IE(InternetExplorer)を表示・非表示
  objIE.Visible = viewFlg

  '指定したURLのページを表示する
  objIE.navigate urlName
 
 'IEが完全表示されるまで待機
 Call ieCheck(objIE)

End Sub


'----------------------------------------------------------------
'②Webページ完全読込待機処理サブルーチン「ieCheck」
Sub ieCheck(objIE As InternetExplorer)

  Dim timeOut As Date

  timeOut = Now + TimeSerial(0, 0, 20)

  Do While objIE.Busy = True Or objIE.readyState <> 4
    DoEvents
    Sleep 1
    If Now > timeOut Then
      objIE.Refresh
      timeOut = Now + TimeSerial(0, 0, 20)
    End If
  Loop

  timeOut = Now + TimeSerial(0, 0, 20)

  Do While objIE.Document.readyState <> "complete"
    DoEvents
    Sleep 1
    If Now > timeOut Then
      objIE.Refresh
      timeOut = Now + TimeSerial(0, 0, 20)
    End If
   Loop

End Sub


'----------------------------------------------------------------
'▼サブルーチンを利用して複数サイトをIEで起動させるマクロ
Sub IEsample()

  Dim objIE  As InternetExplorer
  Dim objIE2  As InternetExplorer

  '本サイトをIEで起動
  Call ieView(objIE, "http://www.vba-ie.net/")

  'yahooサイトをIEで起動
  Call ieView(objIE2, "http://www.yahoo.co.jp/")

End Sub

