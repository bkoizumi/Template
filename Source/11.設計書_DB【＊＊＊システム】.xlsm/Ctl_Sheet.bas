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
  
  With Frm_addSheet
    .StartUpPosition = 1
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
'  Dim newSheetName As String
'
'  '�����J�n--------------------------------------
'  'On Error GoTo catchError
'  Const funcName As String = "Ctl_Sheet.addSheet"
'
'  If runFlg = False Then
'    Call Library.startScript
'  End If
'  Call init.Setting
'  '----------------------------------------------
'
'  newSheetName = sheetCopyTable.Range("F8")
'
'  sheetCopyTable.copy After:=Worksheets(Worksheets.count)
'  ActiveWorkbook.Sheets(Worksheets.count).Tab.ColorIndex = -4142
'  ActiveWorkbook.Sheets(Worksheets.count).Name = newSheetName
'  Sheets(newSheetName).Select
'  Application.GoTo Reference:=Range("A1"), Scroll:=True
'  Range("B" & 9).Select
'
'
'  sheetCopyTable.Range("F8") = ""
'  sheetCopyTable.Range("F9") = ""
'  sheetCopyTable.Range("F11") = ""
'  sheetCopyTable.Range("F10") = ""
'
'
'  '�����I��--------------------------------------
'  If runFlg = False Then
'    Call Library.endScript
'  End If
'  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Description, True)
End Function


