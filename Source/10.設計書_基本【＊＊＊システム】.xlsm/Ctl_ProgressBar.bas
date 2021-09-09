Attribute VB_Name = "Ctl_ProgressBar"
'***********************************************************************************************************************************************
' * プログレスバー表示制御
' *
' * http://www.ne.jp/asahi/fuji/lake/excel/progress_01.html
'***********************************************************************************************************************************************
'Option Explicit

Public mypbProgCnt As Long       'Progress カウンター変数
Public mypbSCount As Long        '処理回数

Dim myJobCnt As Long             '現在進行中の回数
Dim myBarSize As Long            'プログレスバーサイズ


'***********************************************************************************************************************************************
' * プログレスバー表示開始
' *
'***********************************************************************************************************************************************
Public Sub ShowStart()
    Dim myMsg1 As String
    
    myMsg1 = " 処理中"

    'ダイアログへ表示
    With FProgress
      .StartUpPosition = 0
      .Top = Application.Top + (ActiveWindow.Width / 4)
      .Left = Application.Left + (ActiveWindow.Height / 2)
      .Caption = myMsg1
      
      'プログレスバーの枠の部分
      With .Label1
        .BorderStyle = fmBorderStyleSingle       '枠あり
        .SpecialEffect = fmSpecialEffectSunken
        .Height = 15
        .Left = 12
        .Width = 250
        .Top = 30
      End With
  
      'プログレスバーのバーの部分
      With .Label2
        .BackColor = RGB(90, 248, 82)
'        .BorderStyle = fmBorderStyleSingle       '枠あり
        .SpecialEffect = fmSpecialEffectRaised
        .Height = 13
        .Left = 13
        .Width = 200
        .Top = 31
      End With

      '進捗状況表示の部分 ( % )
      With .Label3
        .TextAlign = fmTextAlignCenter
        .BackStyle = 0
        .Height = 14
        .Left = 12
        .Width = 250
        .Top = 32
        .Font.Size = 10
        .Font.Bold = False
      End With
      
      
      'メッセージ表示の部分
      With .Label4
        '.TextAlign = fmTextAlignCenter
        '.SpecialEffect = fmSpecialEffectEtched   '枠が沈む
        '.SpecialEffect = fmSpecialEffectRaised   '浮き上がる
        '.SpecialEffect = fmSpecialEffectBump
        .Caption = myMsg1
        .Height = 14
        .Left = 12
        .Width = 250
        .Top = 9
        .Font.Size = 9
        .Font.Bold = False
      End With

      myBarSize = .Label3.Width
    End With
  
    FProgress.SHOW vbModeless
End Sub


'***********************************************************************************************************************************************
' * プログレスバー表示更新
' *
'***********************************************************************************************************************************************
Public Sub ShowCount(ProgressBarTitle As String, mypbProgCnt As Long, mypbSCount As Long, myMsg1 As String)
  Dim myMsg2 As String
  
  myJobCnt = mypbProgCnt / mypbSCount * 100
  myMsg2 = Int(myJobCnt) & " %"
  
  With FProgress
    .Caption = ProgressBarTitle
    .Label2.Width = myBarSize * myJobCnt / 100
    .Label3.Caption = myMsg2
    .Label4.Caption = myMsg1
  End With
  
  DoEvents
End Sub


'***********************************************************************************************************************************************
' * プログレスバー表示終了
' *
'***********************************************************************************************************************************************
Public Sub ShowEnd()
  
  'ダイアログを閉じる
  Unload FProgress
  
End Sub
