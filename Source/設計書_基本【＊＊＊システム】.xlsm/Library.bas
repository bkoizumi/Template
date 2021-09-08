Attribute VB_Name = "Library"

'***********************************************************************************************************************************************
' * 画面描写制御開始
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Public Function StartScript()

  'アクティブセルの取得
   SelectionCell = Selection.Address
  
  ' マクロ動作でシートやウィンドウが切り替わるのを見せないようにします
  Application.ScreenUpdating = False
  
  ' マクロ動作自体で別のイベントが生成されるのを抑制する
  Application.EnableEvents = False
  
  ' マクロ動作でセルItemNameなどが変わる時自動計算が処理を遅くするのを避ける
  Application.Calculation = xlCalculationManual
  
  ' マクロ動作中に一切のキーやマウス操作を制限する
'  Application.Interactive = False
  
  ' マクロ動作中はマウスカーソルを「砂時計」にする
'  Application.Cursor = xlWait
  
  ' 確認メッセージを出さない
  Application.DisplayAlerts = False

End Function

'***********************************************************************************************************************************************
' * 画面描写制御終了
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Public Function EndScript()

  ' マクロ動作でシートやウィンドウが切り替わるのを見せないようにします
  Application.ScreenUpdating = True
  
  ' マクロ動作自体で別のイベントが生成されるのを抑制する
  Application.EnableEvents = True
  
  ' マクロ動作でセルItemNameなどが変わる時自動計算が処理を遅くするのを避ける
  Application.Calculation = xlCalculationAutomatic
  
  ' マクロ動作中に一切のキーやマウス操作を制限する
'  Application.Interactive = True
  
  ' マクロ動作終了後はマウスカーソルを「デフォルト」にもどす
  Application.Cursor = xlDefault
  
  ' マクロ動作終了後はステータスバーを「デフォルト」にもどす
  Application.StatusBar = False

  ' 確認メッセージを出さない
  Application.DisplayAlerts = True
  
  ' 強制的に再計算させる
  'Application.CalculateFull
  
'  アクティブセルの選択
'  If SelectionCell <> "" Then
'    Range(SelectionCell).Select
'  End If

End Function
'**************************************************************************************************
' * デバッグ用画面表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showDebugForm(ByVal meg1 As String, Optional meg2 As String)
  Dim runTime As Date
  Dim StartUpPosition As Long

'  On Error GoTo catchError

  runTime = Format(Now(), "yyyy/mm/dd hh:nn:ss")

  meg1 = Replace(meg1, vbNewLine, " ")
  
  Select Case setVal("debugMode")
    Case "develop"
      Debug.Print runTime & vbTab & meg1
    Case Else
      Exit Function
  End Select
  
  DoEvents
  Exit Function

'エラー発生時=====================================================================================
catchError:
  Exit Function
End Function

'**************************************************************************************************
' * 処理情報通知
' *
' * Worksheets("Notice").Visible = True
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showNotice(Code As Long, Optional process As String, Optional runEndflg As Boolean)
  Dim Message As String
  Dim runTime As Date
  Dim endLine As Long

  On Error GoTo catchError


  runTime = Format(Now(), "yyyy/mm/dd hh:nn:ss")

  endLine = BK_sheetNotice.Cells(Rows.count, 1).End(xlUp).Row
  Message = Application.WorksheetFunction.VLookup(Code, BK_sheetNotice.Range("A2:B" & endLine), 2, False)
  Message = Replace(Message, "%%", process)
  If process = "" Then
    Message = Replace(Message, "<>", process)
  End If
  If runEndflg = True Then
    Message = Message & vbNewLine & "処理を中止します"
  End If

  If StopTime <> 0 Then
    Message = Message & vbNewLine & "<処理時間：" & StopTime & ">"
  End If

  If Message <> "" Then
    Message = Replace(Message, "<BR>", vbNewLine)
  End If

  If setVal("debugMode") = "speak" Or setVal("debugMode") = "develop" Or setVal("debugMode") = "all" Then
    Application.Speech.Speak Text:=Message, SpeakAsync:=True, SpeakXML:=True
  End If

  Select Case Code
    Case 0 To 399
      Call MsgBox(Message, vbInformation, thisAppName)

    Case 400 To 499
      Call MsgBox(Message, vbCritical, thisAppName)

    Case 500 To 599
      Call MsgBox(Message, vbExclamation, thisAppName)

    Case 999

    Case Else
      Call MsgBox(Message, vbCritical, thisAppName)
  End Select
  
  Message = Replace(Message, vbNewLine & "処理を中止します", "。処理を中止します")
  Message = "[" & Code & "]" & Message
  
  '画面描写制御終了処理
  If runEndflg = True Then
    Call EndScript
    Call Ctl_ProgressBar.ShowEnd
    End
  Else
    Call Library.showDebugForm(Message)
  End If

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call MsgBox(Message, vbCritical, thisAppName)

End Function
