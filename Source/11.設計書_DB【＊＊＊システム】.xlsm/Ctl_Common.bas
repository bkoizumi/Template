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
  Call Library.showDebugForm("StartFun", funcName, "info")
  
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  line = startLine
  Call Ctl_Common.chkRowStartLine
  
  On Error Resume Next
  Range("B" & startLine & ":BB" & setLine("columnEnd")).SpecialCells(xlCellTypeConstants, 23).ClearContents
  Rows(startLine & ":" & setLine("columnEnd")).RowHeight = setVal("defaultRowHeight")
  
  Range("B" & setLine("indexStart") & ":BB" & setLine("indexEnd")).SpecialCells(xlCellTypeConstants, 23).ClearContents
  Rows(setLine("indexStart") & ":" & setLine("indexEnd")).RowHeight = setVal("defaultRowHeight")
  
  
  '処理終了--------------------------------------
  Call Library.showDebugForm("EndFun  ", funcName, "info")
  If runFlg = False Then
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
Function addRow(ByVal line As Long)
  
  If Range("BF" & line + 1) = "ColumEnd" Then
    Rows(line & ":" & line).copy
    Rows(line & ":" & line).Insert Shift:=xlDown
    Application.CutCopyMode = False
    
    Range("BF" & line) = "addLine"
    Call Library.showDebugForm("Ctl_Common.addRow", "true", "info")
    Call Ctl_Common.chkRowStartLine
  
  ElseIf Range("BF" & line + 2) = "IndexEnd" Then
    Rows(line + 1 & ":" & line + 1).copy
    Rows(line + 1 & ":" & line + 1).Insert Shift:=xlDown
    Application.CutCopyMode = False
    
    
    Range("BF" & line + 1) = "addLine"
    Call Library.showDebugForm("Ctl_Common.addRow", "true", "info")
    Call Ctl_Common.chkRowStartLine
  End If
End Function


'==================================================================================================
Function chkRowStartLine()
  Dim line As Long, endLine As Long
  Dim IndexRow As Long
  
  If setLine Is Nothing Then
    Call init.Setting
  End If
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  For line = startLine To endLine
    If Range("BF" & line) = "ColumEnd" Then
      setLine("columnEnd") = line
      
    ElseIf Range("BF" & line) = "IndexStart" Then
      setLine("indexStart") = line
    
    ElseIf Range("BF" & line) = "IndexEnd" Then
      setLine("indexEnd") = line
      Exit For
      
    ElseIf Range("BF" & line) = "TriggerStart" Then
        setLine("triggerStart") = line + 1
    End If
    DoEvents
  Next
  
  Call Library.showDebugForm("columnEnd   ", setLine("columnEnd"), "info")
  Call Library.showDebugForm("indexStart  ", setLine("indexStart"), "info")
  Call Library.showDebugForm("indexEnd    ", setLine("indexEnd"), "info")
  Call Library.showDebugForm("triggerStart", setLine("triggerStart"), "info")

End Function


'==================================================================================================
Function addSheet(newSheetName As String)
  
  Const funcName As String = "Ctl_Common.addSheet"
  On Error GoTo catchError
  
  If newSheetName Like "3.*" Then
  
  Else
    newSheetName = "3." & newSheetName
  End If
  
  If Library.chkSheetExists(newSheetName) = False Then
    sheetCopyTable.copy before:=Worksheets("5.容量計算")
    
    If LenB(StrConv(newSheetName, vbFromUnicode)) > 25 Then
      ActiveSheet.Name = Library.cutRight(newSheetName, LenB(StrConv(newSheetName, vbFromUnicode)) - 25) & "..."
    Else
      ActiveSheet.Name = newSheetName
    End If
  End If
  Call Library.showDebugForm(funcName, ActiveSheet.Name)
  
  Sheets(ActiveSheet.Name).Select
  ActiveSheet.Tab.ColorIndex = -4142
  
  Range("AO3") = ActiveSheet.Name
  
  Range("AO1") = Application.UserName
  Range("AX1") = Format(Date, "yyyy/mm/dd")

  '書式設定----------------------------------
  '初期値のリスト化
  With Range("AB16:AE31").Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=defVal_" & setVal("DBMS")
  End With
  
  Range("AF16:AO31").NumberFormatLocal = """YES"""
  
  '備考の結合
  Range("AP16:BB31").Merge True
  
  addSheet = ActiveSheet.Name
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function

'==================================================================================================
Function chkTableName2SheetName(tableName As String)
  Dim SheetName As String
  
  Const funcName As String = "Ctl_Common.chkTableName2SheetName"
  On Error GoTo catchError
  
  If LenB(StrConv(tableName, vbFromUnicode)) > 25 Then
    SheetName = "3." & Library.cutRight(tableName, LenB(StrConv(tableName, vbFromUnicode)) - 25) & "..."
  Else
    SheetName = "3." & tableName
  End If
  If Library.chkSheetExists(tableName) = False Then
    SheetName = Ctl_Common.addSheet(SheetName)
  End If
  
  Worksheets(SheetName).Select
  chkTableName2SheetName = SheetName
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function


'==================================================================================================
Function makeTblList()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim sheetList As Object
  Dim targetSheet   As Worksheet
  Dim SheetName As String
  
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
  Call Library.showDebugForm("StartFun", funcName, "info")
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
    SheetName = sheetList.Name
    
    Select Case SheetName
      Case "表紙", "変更履歴", "1.エンティティ", "2.ER図", "5.容量計算", "空白"
      Case Else
        If SheetName.Name Like "<*>" Then
        Else
          Call Library.showDebugForm("sheetName", SheetName)
      
          sheetTblList.Range("B" & line).FormulaR1C1 = "=ROW()-5"
          
          '分類
          sheetTblList.Range("C" & line) = Sheets(SheetName).Range("C2")
          
          '説明
          sheetTblList.Range("V" & line) = Sheets(SheetName).Range("E6")
        
          '論理テーブル名
          If Sheets(SheetName).Range("F8") <> "" Then
            With sheetTblList.Range("B" & line)
              .Value = Sheets(SheetName).Range("F8")
              .Select
              .Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=SheetName & "!" & "A9"
              .Font.Color = RGB(0, 0, 0)
              .Font.Underline = False
              .Font.Size = 9
              .Font.Name = "メイリオ"
            End With
          End If
          
           '物理テーブル名
          With sheetTblList.Range("E" & line)
            .Value = Sheets(SheetName).Range("W5") & "." & Sheets(SheetName).Range("F9")
            .Select
            .Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=SheetName & "!" & "A9"
            .Font.Color = RGB(0, 0, 0)
            .Font.Underline = False
            .Font.Size = 9
            .Font.Name = "メイリオ"
          End With
          
          ' シート色と同じ色をセルに設定
          If Sheets(SheetName).Tab.Color Then
            With sheetTblList.Range("B" & line & ":U" & line).Interior
              .Pattern = xlPatternNone
              .Color = Sheets(SheetName).Tab.Color
            End With
          End If
        
        
        line = line + 1
      End If
    End Select
  Next
  
  
  '処理終了--------------------------------------
  'Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("EndFun  ", funcName, "info")
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
      rslFlg = Ctl_Access.IsTable(Range("F9"))
      
    Case "MySQL"
      rslFlg = Ctl_MySQL.IsTable(Range("F9"))
      
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
  Call Library.showDebugForm("StartFun", funcName, "info")
  '--------------------------
  Set targetSheet = ActiveSheet
  line = ActiveCell.Row
  
  Rows(line & ":" & line).Insert Shift:=xlDown
  
  sheetCopyTable.Rows("46:46").copy
  targetSheet.Range("A" & line).Select
  ActiveSheet.Paste
  targetSheet.Range("B" & line) = "insert"
  targetSheet.Range("B" & line).Select
  
  '処理終了--------------------------------------
  Call Library.showDebugForm("EndFun  ", funcName, "info")
  '----------------------------------------------
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function


'==================================================================================================
Function deleteRow()
  Dim line As Long, endLine As Long

  '処理開始--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_Common.deleteRow"
  
  Call init.Setting
  Call Library.showDebugForm("StartFun", funcName, "info")
  '--------------------------
  Set targetSheet = ActiveSheet
  line = ActiveCell.Row
  
  targetSheet.Range("C" & line & ":Z" & line).Style = "不要"
  targetSheet.Range("B" & line) = "delete"
  
  
  
  
  '処理終了--------------------------------------
  Call Library.showDebugForm("EndFun  ", funcName, "info")
  '----------------------------------------------
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
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
  Call Library.showDebugForm("StartFun", funcName, "info")
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
  Call Library.showDebugForm("EndFun  ", funcName, "info")

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function
