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

'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function

'==================================================================================================
'シート追加
Function addSheet()
  Dim newSheetName As String
  
  '処理開始--------------------------------------
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
  
  
  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.endScript
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Description, True)
End Function


