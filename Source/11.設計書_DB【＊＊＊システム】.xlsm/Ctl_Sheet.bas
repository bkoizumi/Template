Attribute VB_Name = "Ctl_Sheet"
'**************************************************************************************************
' * シート処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'シート追加
Function showAddSheetOption()
  Dim topPosition As Long, leftPosition As Long
'  On Error GoTo catchError
  
  Call init.Setting
  
  With Frm_addSheet
    .StartUpPosition = 1
    .Show
  End With

  Exit Function

'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function

'==================================================================================================
'シート追加
Function addSheet()
'  Dim newSheetName As String
'
'  '処理開始--------------------------------------
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
'  '処理終了--------------------------------------
'  If runFlg = False Then
'    Call Library.endScript
'  End If
'  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Description, True)
End Function


