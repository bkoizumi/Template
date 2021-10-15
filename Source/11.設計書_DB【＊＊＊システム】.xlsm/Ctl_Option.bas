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

  Const funcName As String = "Ctl_Option.ClearAll"
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
  End If
  PrgP_Max = 2
  PrgP_Cnt = 1
  
  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------

  line = 1
  endLine = Sheets.count
  For Each tempSheet In Sheets
    Select Case tempSheet.Name
      Case "表紙", "変更履歴", "1.エンティティ", "2.ER図", "5.容量計算", "空白"
      Case Else
        If tempSheet.Name Like "<*>" Then
        Else
          Call Library.showDebugForm("シート削除：" & tempSheet.Name)
          Worksheets(tempSheet.Name).Delete
        End If
    End Select
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, line, endLine, "データクリア")
    line = line + 1
  Next

  '処理終了--------------------------------------
  Call Library.showDebugForm("EndFun  ", funcName, "info")
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.unsetting(True)
  End If
  '----------------------------------------------
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function



'==================================================================================================
'オプション表示
Function showOption()

  '処理開始--------------------------------------
  'On Error GoTo catchError

  Const funcName As String = "Ctl_Option.ClearAll"
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
  End If
  
  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------
  
  With Frm_setOption
    .StartUpPosition = 1
    .Caption = thisAppName & " [" & thisAppVersion & "]"
    .Show
  End With

  '処理終了--------------------------------------
  Call Library.showDebugForm("EndFun  ", funcName, "info")
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
    Call Library.endScript
    Call init.unsetting(True)
  End If
  '----------------------------------------------
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function


'==================================================================================================
'オプション設定
Function setOption()
  '処理開始--------------------------------------
  'On Error GoTo catchError

  Const funcName As String = "Ctl_Option.ClearAll"
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
  End If
  
  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------














  '処理終了--------------------------------------
  Call Library.showDebugForm("EndFun  ", funcName, "info")
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.unsetting(True)
  End If
  '----------------------------------------------
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function


