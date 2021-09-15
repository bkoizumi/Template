Attribute VB_Name = "Ctl_Option"
Option Explicit

'**************************************************************************************************
' * Ctl_Option
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function ClearAll()
  Dim line As Long, endLine As Long
  Dim tempSheet As Object
  
  '処理開始--------------------------------------
  'On Error GoTo catchError

  FuncName = "Ctl_Option.ClearAll"
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
  End If
  
  Call Library.showDebugForm(FuncName & "==============================================")
  '----------------------------------------------


  For Each tempSheet In Sheets
    Select Case tempSheet.Name
      Case "設定-MySQL", "設定-ACC", "設定", "Notice", "DataType", "コピー用", "表紙", "TBLリスト", "変更履歴", "ER図"
      Case Else
        Call Library.showDebugForm("シート削除：" & tempSheet.Name)
        Worksheets(tempSheet.Name).Delete
    End Select
  Next

  '処理終了--------------------------------------
  Call Library.showDebugForm("=================================================================")
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.usetting
  End If
  '----------------------------------------------
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function


