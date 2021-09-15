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
      Case "�ݒ�-MySQL", "�ݒ�-ACC", "�ݒ�", "Notice", "DataType", "�R�s�[�p", "�\��", "TBL���X�g", "�ύX����", "ER�}"
      Case Else
        Call Library.showDebugForm("�V�[�g�폜�F" & tempSheet.Name)
        Worksheets(tempSheet.Name).Delete
    End Select
  Next

  '�����I��--------------------------------------
  Call Library.showDebugForm("=================================================================")
  If runFlg = False Then
    Application.Goto Reference:=Range("A1"), Scroll:=True
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.usetting
  End If
  '----------------------------------------------
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "�F" & Err.Description, True)
End Function


