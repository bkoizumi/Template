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
  
  '�����J�n--------------------------------------
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
      Case "�\��", "�ύX����", "1.�G���e�B�e�B", "2.ER�}", "5.�e�ʌv�Z", "��"
      Case Else
        If tempSheet.Name Like "<*>" Then
        Else
          Call Library.showDebugForm("�V�[�g�폜�F" & tempSheet.Name)
          Worksheets(tempSheet.Name).Delete
        End If
    End Select
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, line, endLine, "�f�[�^�N���A")
    line = line + 1
  Next

  '�����I��--------------------------------------
  Call Library.showDebugForm("EndFun  ", funcName, "info")
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.unsetting(True)
  End If
  '----------------------------------------------
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function



'==================================================================================================
'�I�v�V�����\��
Function showOption()

  '�����J�n--------------------------------------
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

  '�����I��--------------------------------------
  Call Library.showDebugForm("EndFun  ", funcName, "info")
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
    Call Library.endScript
    Call init.unsetting(True)
  End If
  '----------------------------------------------
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function


'==================================================================================================
'�I�v�V�����ݒ�
Function setOption()
  '�����J�n--------------------------------------
  'On Error GoTo catchError

  Const funcName As String = "Ctl_Option.ClearAll"
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
  End If
  
  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------














  '�����I��--------------------------------------
  Call Library.showDebugForm("EndFun  ", funcName, "info")
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.unsetting(True)
  End If
  '----------------------------------------------
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function


