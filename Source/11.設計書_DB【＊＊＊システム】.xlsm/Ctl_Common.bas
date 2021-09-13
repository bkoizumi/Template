Attribute VB_Name = "Ctl_Common"
Option Explicit

'**************************************************************************************************
' * 共通処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function ClearData()
  Dim line As Long, endLine As Long

  '処理開始----------------------------------------------------------------------------------------
  'On Error GoTo catchError
  '初期値設定----------------
  FuncName = "Ctl_Common.ClearData"
  '--------------------------
  
  Call init.Setting
  Call Library.showDebugForm(FuncName & "=============================================")
  
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  line = startLine
  
  Do Until Range("A" & line) = "End"
    If Range("A" & line) = "" Then
      Rows(line & ":" & line).Delete Shift:=xlUp
      line = line - 1
    ElseIf Range("A" & line) = "Column" Then
      On Error Resume Next
      Range("B" & line & ":AA" & line).SpecialCells(xlCellTypeConstants, 23).ClearContents
      On Error GoTo catchError
    
    End If
'    Call Ctl_ProgressBar.showBar(thisAppName, 1, PrgP_Max, line, endLine, "データクリア")
    line = line + 1
  Loop
  Columns("L:R").EntireColumn.Hidden = True
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function


'==================================================================================================
Function addRow(line As Long)

  If line >= startLine + 2 Then
    Rows(line & ":" & line).Copy
    Rows(line & ":" & line).Insert Shift:=xlDown
    Range("A" & line) = ""
    Application.CutCopyMode = False
  End If
End Function


'==================================================================================================
Function chkIndexRow()
  Dim line As Long, endLine As Long
  Dim IndexRow As Long
  
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  For line = startLine To endLine
    If Range("A" & line) = "IndexStart" Then
      IndexRow = line
      Exit For
    End If
  Next
  Call Library.showDebugForm("IndexRow：" & IndexRow)
  chkIndexRow = IndexRow
End Function


'==================================================================================================
Function addSheet(newSheetName As String)
  
  On Error GoTo catchError
  
  If Library.chkSheetExists(newSheetName) = False Then
    sheetCopy.Copy After:=Worksheets(Worksheets.count)
    ActiveWorkbook.Sheets(Worksheets.count).Name = newSheetName
  End If
  Sheets(newSheetName).Select
    
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, Err.Description, True)
End Function


'==================================================================================================
Function makeTblList()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim sheetList As Object
  Dim targetSheet   As Worksheet
  Dim sheetName As String
  
'  On Error GoTo catchError

  Call init.Setting
  sheetTblList.Select
  endLine = sheetTblList.Cells(Rows.count, 3).End(xlUp).Row + 1
  Range("C6:U" & endLine).ClearContents
  
  With Range("B6:U" & endLine).Interior
    .Pattern = xlPatternNone
    .Color = xlNone
  End With
  
      
  line = 6
  For Each sheetList In Sheets
    sheetName = sheetList.Name
    
    
    Select Case sheetName
      Case "設定", "Notice", "DataType", "コピー用", "表紙", "TBLリスト", "変更履歴", "ER図"
      Case Else
    
        sheetTblList.Range("B" & line).FormulaR1C1 = "=ROW()-5"
        sheetTblList.Range("C" & line) = Sheets(sheetName).Range("B2")
  '      sheetTblList.Range("E" & line) = Sheets(sheetName).Range("D5")
  '      sheetTblList.Range("H" & line) = Sheets(sheetName).Range("H5")
        sheetTblList.Range("Q" & line) = Sheets(sheetName).Range("D6")
      
        '論理テーブル名
        If Sheets(sheetName).Range("D5") <> "" Then
          With sheetTblList.Range("E" & line)
            .Value = Sheets(sheetName).Range("D5")
            .Select
            .Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=sheetName & "!" & "A9"
            .Font.Color = RGB(0, 0, 0)
            .Font.Underline = False
            .Font.Size = 10
            .Font.Name = "メイリオ"
          End With
        End If
        
         '物理テーブル名
        With sheetTblList.Range("H" & line)
          .Value = Sheets(sheetName).Range("H5")
          .Select
          .Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=sheetName & "!" & "A9"
          .Font.Color = RGB(0, 0, 0)
          .Font.Underline = False
          .Font.Size = 10
          .Font.Name = "メイリオ"
        End With
        
        ' シート色と同じ色をセルに設定
        If Sheets(sheetName).Tab.Color Then
          With sheetTblList.Range("B" & line & ":U" & line).Interior
            .Pattern = xlPatternNone
            .Color = Sheets(sheetName).Tab.Color
          End With
        End If
        
        
        line = line + 1
    End Select
  Next
  
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
    
    
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, Err.Description, True)
End Function

