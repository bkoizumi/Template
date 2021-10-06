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

  '処理開始--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_Common.ClearData"
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
  End If
  '--------------------------
  
  Call init.Setting
  Call Library.showDebugForm(funcName & "=============================================")
  
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  line = startLine
  
  Do Until Range("A" & line) = "End"
    If Range("A" & line) = "" Then
      Rows(line & ":" & line).Delete Shift:=xlUp
      line = line - 1
    ElseIf Range("A" & line) = "Column" Then
      On Error Resume Next
      Range("B" & line & ":AZ" & line).SpecialCells(xlCellTypeConstants, 23).ClearContents
      Rows(line & ":" & line).RowHeight = setVal("defaultRowHeight")
      On Error GoTo catchError
    End If
    DoEvents
    line = line + 1
  Loop
  Columns("H:H").EntireColumn.Hidden = True
  Columns("M:S").EntireColumn.Hidden = True
  
  '処理終了--------------------------------------
  Call Library.showDebugForm("=================================================================")
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.unsetting(True)
  End If
  '----------------------------------------------
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function


'==================================================================================================
Function addRow(line As Long)

  If line >= 47 Then
    Rows(line & ":" & line).copy
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
  Call Library.showDebugForm("IndexRow", IndexRow)
  chkIndexRow = IndexRow
End Function


'==================================================================================================
Function addSheet(newSheetName As String)
  
  On Error GoTo catchError
  
  If Library.chkSheetExists(newSheetName) = False Then
    sheetCopy.copy After:=Worksheets(Worksheets.count)
    ThisWorkbook.Sheets(Worksheets.count).Name = newSheetName
  End If
  Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, 1, 2, "シート追加")
  
  Sheets(newSheetName).Select
  Range(setVal("Cell_UpdateBy")) = Application.UserName
  Range(setVal("Cell_UpdateAt")) = Format(Date, "yyyy/mm/dd")
'  Range("W9:AA48").Merge True
  
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
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_Common.makeTblList"

  'runFlg = True
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
    Call Library.showDebugForm("runFlg", runFlg)
  End If
  Call Library.showDebugForm(funcName & "===========================================")
  '----------------------------------------------
  
  sheetTblList.Select
  endLine = sheetTblList.Cells(Rows.count, 3).End(xlUp).Row + 1
  Range("C6:I" & endLine).ClearContents
  
  With Range("B6:U" & endLine).Interior
    .Pattern = xlPatternNone
    .Color = xlNone
  End With
  
  line = 6
  For Each sheetList In Sheets
    sheetName = sheetList.Name
    
    Select Case sheetName
      Case "設定-MySQL", "設定-ACC", "Notice", "DataType", "コピー用", "表紙", "TBLリスト", "変更履歴", "ER図"
      Case Else
        Call Library.showDebugForm("sheetName", sheetName)
    
        sheetTblList.Range("B" & line).FormulaR1C1 = "=ROW()-5"
        sheetTblList.Range("C" & line) = Sheets(sheetName).Range("C2")
        sheetTblList.Range(setVal("Cell_dateType") & line) = Sheets(sheetName).Range("D6")
      
        '論理テーブル名
        If Sheets(sheetName).Range("D5") <> "" Then
          With sheetTblList.Range(setVal("Cell_logicalName") & line)
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
        With sheetTblList.Range("E" & line)
          .Value = Sheets(sheetName).Range(setVal("Cell_physicalTableName"))
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
  'Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.unsetting(True)
  End If
  '----------------------------------------------
    
    
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, Err.Description, True)
End Function

'==================================================================================================
Function IsTable(tableName As String) As Boolean
  Dim rslFlg As Boolean
  
  Select Case setVal("DBMS")
    Case "MSAccess"
      rslFlg = Ctl_Access.IsTable(Range(setVal("Cell_physicalTableName")))
      
    Case "MySQL"
      rslFlg = Ctl_MySQL.IsTable(Range(setVal("Cell_physicalTableName")))
      
    Case "PostgreSQL"
      
    Case "SQLServer"
      
    Case Else
  End Select
  
  If rslFlg = False Then
    Range("B5") = "newTable"
  Else
    Range("B5") = "exist"
  End If
  
  IsTable = rslFlg
End Function



'==================================================================================================
Function insertRow()
  Dim line As Long, endLine As Long

  '処理開始--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_Common.insertRow"
  
  Call init.Setting
  Call Library.showDebugForm(funcName & "=============================================")
  '--------------------------
  Set targetSheet = ActiveSheet
  line = ActiveCell.Row
  
  Rows(line & ":" & line).Insert Shift:=xlDown
  
  sheetCopy.Rows("46:46").copy
  targetSheet.Range("A" & line).Select
  ActiveSheet.Paste
  targetSheet.Range("B" & line) = "insert"
  targetSheet.Range(setVal("Cell_logicalName") & line).Select
  
  '処理終了--------------------------------------
  Call Library.showDebugForm("=================================================================")
  '----------------------------------------------
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function


'==================================================================================================
Function deleteRow()
  Dim line As Long, endLine As Long

  '処理開始--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_Common.deleteRow"
  
  Call init.Setting
  Call Library.showDebugForm(funcName & "=============================================")
  '--------------------------
  Set targetSheet = ActiveSheet
  line = ActiveCell.Row
  
  targetSheet.Range("C" & line & ":Z" & line).Style = "不要"
  targetSheet.Range("B" & line) = "delete"
  
  
  
  
  '処理終了--------------------------------------
  Call Library.showDebugForm("=================================================================")
  '----------------------------------------------
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function


'==================================================================================================
Function 右クリックメニュー(Target As Range, Cancel As Boolean)
  Dim menu01 As CommandBarControl
  Dim cmdBra As CommandBarControl
  
  Call init.Setting
  
  '標準状態にリセット
  Application.CommandBars("Cell").Reset

  '全てのメニューを一旦削除
  For Each cmdBra In Application.CommandBars("Cell").Controls
    cmdBra.Visible = False
  Next
  
  With CommandBars("Cell")
    With .Controls.Add()
      .BeginGroup = True
      .Caption = "行の挿入"
      .FaceId = 296
      .OnAction = "Ctl_Common.insertRow"
    End With
    With .Controls.Add()
      .Caption = "行の削除"
      .FaceId = 293
      .OnAction = "Ctl_Common.deleteRow"
    End With
  End With

  Application.CommandBars("Cell").ShowPopup
  Application.CommandBars("Cell").Reset
  Cancel = True
End Function


'==================================================================================================
Function chkEditRow(targetRange As Range, changeVal As String)
  Dim line As Long, endLine As Long
  
  '処理開始--------------------------------------
  'On Error GoTo catchError

  
  Const funcName As String = "Ctl_Option.chkEditRow"
  Call Library.showDebugForm(funcName & "==========================================")
  '--------------------------------------

  Call Library.showDebugForm("targetRange.Value ：" & targetRange.Value)
  Call Library.showDebugForm("targetRange.Column：" & targetRange.Column)
  Call Library.showDebugForm("targetRange.Row   ：" & targetRange.Row)
    
  If targetRange.Column = 5 Then
    Select Case Range("B" & targetRange.Row)
      Case "", "edit"
        Range("B" & targetRange.Row) = "rename:" & oldCellVal
      Case "insert"
      Case "delete"
      Case Else
        If Range("B" & targetRange.Row) Like "rename:*" Then
          '元に戻したとき
          If targetRange.Value = Replace(Range("B" & targetRange.Row), "rename:", "") Then
            Range("B" & targetRange.Row) = ""
          End If
        End If
    End Select
  Else
    Select Case Range("B" & targetRange.Row)
      Case ""
        Range("B" & targetRange.Row) = changeVal
      Case "edit"
      
      Case "insert"
      Case "delete"
        
      Case Else
    End Select
  
  End If

  '更新情報を設定
  If Range(setVal("Cell_CreateAt")) <> Format(Date, "yyyy/mm/dd") Then
    Range(setVal("Cell_UpdateBy")) = Application.UserName
    Range(setVal("Cell_UpdateAt")) = Format(Date, "yyyy/mm/dd")
  End If



  '処理終了--------------------------------------
  Call Library.showDebugForm("=================================================================")

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function
