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
  
  With Frm_NewSheet
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
  Dim newSheetName As String
  
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_Sheet.addSheet"

  If runFlg = False Then
    Call Library.startScript
  End If
  Call init.Setting
  '----------------------------------------------
  
  newSheetName = sheetCopy.Range(setVal("Cell_logicalTableName"))
  
  sheetCopy.copy After:=Worksheets(Worksheets.count)
  ActiveWorkbook.Sheets(Worksheets.count).Tab.ColorIndex = -4142
  ActiveWorkbook.Sheets(Worksheets.count).Name = newSheetName
  Sheets(newSheetName).Select
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Range(setVal("Cell_logicalName") & 9).Select
  
  
  sheetCopy.Range(setVal("Cell_logicalTableName")) = ""
  sheetCopy.Range(setVal("Cell_physicalTableName")) = ""
  sheetCopy.Range(setVal("Cell_tableNote")) = ""
  sheetCopy.Range(setVal("Cell_TableType")) = ""
  
  
  '�����I��--------------------------------------
  If runFlg = False Then
    Call Library.endScript
  End If
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Description, True)
End Function


