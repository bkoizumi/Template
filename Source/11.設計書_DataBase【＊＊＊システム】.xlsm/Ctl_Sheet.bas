Attribute VB_Name = "Ctl_Sheet"
'**************************************************************************************************
' * �V�[�g����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'�V�[�g�ǉ�
Function showAddSheetOption()
  Dim topPosition As Long, leftPosition As Long
'  On Error GoTo catchError
  
  Call init.Setting
  topPosition = setVal("Frm_NewSheetTop")
  leftPosition = setVal("Frm_NewSheetLeft")
  
  With Frm_NewSheet
    If topPosition = 0 Then
      .StartUpPosition = 2
    Else
      .StartUpPosition = 0
      .Top = topPosition
      .Left = leftPosition
    End If
    '.Show vbModeless
    .Show
  End With

  Exit Function

'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function

'==================================================================================================
'�V�[�g�ǉ�
Function addSheet()
  Dim newSheetName As String
  
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  FuncName = "Ctl_Sheet.addSheet"

  If runFlg = False Then
    Call Library.startScript
  End If
  Call init.Setting
  '----------------------------------------------
  
  newSheetName = CopySheet.Range("H5")
  
  CopySheet.Copy After:=Worksheets(Worksheets.count)
  ActiveWorkbook.Sheets(Worksheets.count).Tab.ColorIndex = -4142
  ActiveWorkbook.Sheets(Worksheets.count).Name = newSheetName
  Sheets(newSheetName).Select
  
  CopySheet.Range("D5") = ""
  CopySheet.Range("H5") = ""
  CopySheet.Range("D6") = ""
  CopySheet.Range("B2") = ""
  
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
  End If
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Description, True)
End Function


