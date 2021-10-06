Attribute VB_Name = "Ctl_ProgressBar"

'**************************************************************************************************
' * プログレスバー表示制御
' *
' * http://www.ne.jp/asahi/fuji/lake/excel/progress_01.html
'**************************************************************************************************
'Option Explicit

Public PrgP_Barsize As Long
Public PrgC_Barsize As Long


Public PrgP_Meg As String
Public PrgC_Meg As String
Public PrgMeg   As String






'==================================================================================================
Public Function showStart()
  
  With Frm_Progress
    .StartUpPosition = 1
'    .Top = Application.Top + 10
'    .Left = Application.Left + 120
    .Caption = thisAppName
    
    
    '親進捗バー------------------------------------------------------------------------------------
    'プログレスバーの枠の部分
    With .PrgP_Area
      .BorderStyle = fmBorderStyleSingle
      .SpecialEffect = fmSpecialEffectSunken
      .Caption = ""
      .Top = 15
      .Left = 12
      .Height = 15
      .Width = 250
    End With

    'プログレスバーのバーの部分
    With .PrgP_Bar
      .BackColor = RGB(90, 248, 82)
      .SpecialEffect = fmSpecialEffectRaised
      .Caption = ""
      .Top = 16
      .Left = 13
      .Height = 13
      .Width = 0
    End With

    '進捗状況表示の部分 ( % )
    With .PrgP_Progress
      .TextAlign = fmTextAlignCenter
      .Caption = "0%"
      .BackStyle = 0
      .Caption = ""
      .Top = 17
      .Left = 12
      .Height = 14
      .Width = 250
      .Font.Size = 10
      .Font.Bold = False
    End With
    
    '子進捗バー------------------------------------------------------------------------------------
    'プログレスバーの枠の部分
    With .PrgC_Area
      .BorderStyle = fmBorderStyleSingle
      .SpecialEffect = fmSpecialEffectSunken
      .Caption = ""
      .Top = 45
      .Left = 12
      .Height = 15
      .Width = 250
    End With

    'プログレスバーのバーの部分
    With .PrgC_Bar
      .BackColor = RGB(90, 248, 82)
      .SpecialEffect = fmSpecialEffectRaised
      .Caption = ""
      .Top = 46
      .Left = 13
      .Height = 13
      .Width = 0
    End With

    '進捗状況表示の部分 ( % )
    With .PrgC_Progress
      .TextAlign = fmTextAlignCenter
      .Caption = "0%"
      .BackStyle = 0
      .Top = 47
      .Left = 12
      .Height = 14
      .Width = 250
      .Font.Size = 10
      .Font.Bold = False
    End With
    
    
    'メッセージ表示の部分
    With .Prg_Message
      .Caption = "待機中"
      .Top = 70
      .Left = 12
      .Height = 30
      .Width = 270
      .Font.Size = 9
      .Font.Bold = False
    End With

    PrgP_Barsize = .PrgP_Area.Width
    PrgC_Barsize = .PrgP_Area.Width
  End With

  Frm_Progress.Show vbModeless
End Function


'プログレスバー表示更新
'==================================================================================================
Public Function showCount( _
                            Prg_Title As String _
                          , PrgC_Cnt As Long, PrgC_Max As Long _
                          , PrgMeg As String _
                          , Optional flg As Boolean = False _
                        )

  Call showBar(Prg_Title, 2, 2, PrgC_Cnt, PrgC_Max, PrgMeg)
                
End Function
'==================================================================================================
Public Function showBar( _
                            Prg_Title As String _
                          , ByVal L_PrgP_Cnt As Long, PrgP_Max As Long _
                          , ByVal L_PrgC_Cnt As Long, PrgC_Max As Long _
                          , PrgMeg As String _
                        )
                        
                        
                        
  Dim myMsg2 As String
  Dim PrgP_Prg As Long, PrgC_Prg As Long
  
'  L_PrgP_Cnt = L_PrgP_Cnt - 1
  L_PrgC_Cnt = L_PrgC_Cnt
  
  If L_PrgP_Cnt <= 0 And L_PrgC_Cnt > 0 Then
    PrgP_Prg = 0
    PrgP_Meg = L_PrgP_Cnt & "/" & PrgP_Max & " (" & PrgP_Prg & "%)"
    
    PrgC_Prg = Int((L_PrgC_Cnt) / PrgC_Max * 100)
    PrgC_Meg = L_PrgC_Cnt & "/" & PrgC_Max & " (" & PrgC_Prg & "%)"
  
  
  
  ElseIf L_PrgP_Cnt > 0 And L_PrgC_Cnt > 0 Then
    PrgP_Prg = Int((L_PrgP_Cnt) / PrgP_Max * 100)
    PrgP_Meg = L_PrgP_Cnt & "/" & PrgP_Max & " (" & PrgP_Prg & "%)"
    
    PrgC_Prg = Int((L_PrgC_Cnt) / PrgC_Max * 100)
    PrgC_Meg = L_PrgC_Cnt & "/" & PrgC_Max & " (" & PrgC_Prg & "%)"
  
  ElseIf L_PrgP_Cnt = 0 Then
    PrgP_Prg = 0
    PrgP_Meg = L_PrgP_Cnt & "/" & PrgP_Max & " (" & PrgP_Prg & "%)"
    
  ElseIf L_PrgC_Cnt = 0 Then
    PrgC_Prg = 0
    PrgC_Meg = L_PrgC_Cnt & "/" & PrgC_Max & " (" & PrgC_Prg & "%)"
  
  End If
  
  With Frm_Progress
    .Caption = Prg_Title
    .PrgP_Bar.Width = Int(PrgP_Barsize * PrgP_Prg / 100)
    .PrgP_Progress.Caption = PrgP_Meg
    
    .PrgC_Bar.Width = Int(PrgC_Barsize * PrgC_Prg / 100)
    .PrgC_Progress.Caption = PrgC_Meg
    
    If PrgMeg = "" Then
      .Prg_Message.Caption = "処理中…"
    Else
      .Prg_Message.Caption = PrgMeg
    End If
  End With
  DoEvents
  
  
  
End Function


'**************************************************************************************************
' * プログレスバー表示終了
' *
'**************************************************************************************************
Public Function showEnd()
  
  'ダイアログを閉じる
  Unload Frm_Progress
  
End Function








