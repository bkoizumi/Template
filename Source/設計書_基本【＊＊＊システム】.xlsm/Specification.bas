Attribute VB_Name = "Specification"
Option Explicit


'**************************************************************************************************
' * 目次生成
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function makeTOC()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim TitleName As String, FunctionName As String, tocTitle As String
  Dim PageCnt As Long, RowCnt As Long
  Dim tocLineCnt As Long, startLine As Long
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  FuncName = "Specification.makeTOC"

  Call Library.StartScript
  Call init.setting(True)
  Call Ctl_ProgressBar.ShowStart
  Call Library.showDebugForm(FuncName & "============================================")
  sheetMain.Select
  '----------------------------------------------
  PageCnt = 1
  RowCnt = 4
  tocLineCnt = 1
  colLine = setVal("Toc1StartColLine")
  startLine = 0
  
  Range("B4:AU42").Cells.Clear
  endLine = sheetMain.Cells(Rows.count, 50).End(xlUp).Row
  
  For line = 44 To endLine Step setVal("PageLine")
    If tocLineCnt Mod 2 <> 0 And PageCnt = setVal("TocMaxCnt") * tocLineCnt Then
      
      colLine = setVal("Toc2StartColLine")
      If tocLineCnt = 1 Then
        RowCnt = 4
      Else
        RowCnt = setVal("PageLine") * WorksheetFunction.RoundDown(tocLineCnt / 2, 0) + 4
      End If
      tocLineCnt = tocLineCnt + 1
      
    ElseIf tocLineCnt Mod 2 = 0 And PageCnt = setVal("TocMaxCnt") * tocLineCnt Then
      colLine = setVal("Toc1StartColLine")
      RowCnt = setVal("PageLine") * WorksheetFunction.RoundDown(tocLineCnt / 2, 0) + 4
      tocLineCnt = tocLineCnt + 1
      
      'タイトルが目次でなければページ追加
      If Cells(RowCnt - 2, colLine + 2) Like "[目次,もくじ]*" Then
        Range(Cells(RowCnt, colLine), Cells(RowCnt + 38, 47)).Cells.Clear
      Else
        Call addPage(Cells(RowCnt - 3, 1).Address)
      End If
    Else
      RowCnt = RowCnt + 1
    End If
    
    If Cells(line + 1, 4) Like "[目次,もくじ]*" Then
    Else
      If startLine = 0 Then
        startLine = line
      End If
    End If
    Cells(line, 50) = "P." & PageCnt + 1
    PageCnt = PageCnt + 1
    
    Call Ctl_ProgressBar.ShowCount("ページ構成確認", line, endLine, tocTitle)
  Next
  
  
  
  
  '----------------------------------------------
  PageCnt = 1
  RowCnt = 4
  tocLineCnt = 1
  colLine = setVal("Toc1StartColLine")
  
  For line = startLine To endLine Step setVal("PageLine")
    TitleName = Cells(line + 1, 4)
    FunctionName = Cells(line + 1, 19)
    
    If FunctionName <> "" Then
      tocTitle = TitleName & " - " & FunctionName
    Else
      tocTitle = TitleName
    End If
    
    'ページ番号書き込み
    'Application.Goto Reference:=Cells(RowCnt, colLine), Scroll:=True
    Range(Cells(RowCnt, colLine), Cells(RowCnt, colLine + 1)).Select
    With Range(Cells(RowCnt, colLine), Cells(RowCnt, colLine + 1))
      .MergeCells = True
      .NumberFormatLocal = "@"
      .Value = PageCnt & "."
      .Font.Name = "メイリオ"
      .Font.Size = 9
      .HorizontalAlignment = xlRight
      .VerticalAlignment = xlCenter
    End With
    
    'Application.Goto Cells(RowCnt, colLine + 2), Scroll:=True
    Range(Cells(RowCnt, colLine + 2), Cells(RowCnt, colLine + 2 + setVal("TocTitle"))).Select
    With Range(Cells(RowCnt, colLine + 2), Cells(RowCnt, colLine + 2 + setVal("TocTitle")))
      .Merge
      .NumberFormatLocal = "@"
      .Value = tocTitle
      .Select
      .Hyperlinks.add Anchor:=Selection, Address:="", SubAddress:="#" & "A" & line
      .Font.ColorIndex = 1
      .Font.Underline = xlUnderlineStyleNone
      .Font.Name = "メイリオ"
      .Font.Size = 9
      .HorizontalAlignment = xlGeneral
      .VerticalAlignment = xlCenter
    End With
    
    
    If tocLineCnt Mod 2 <> 0 And PageCnt = setVal("TocMaxCnt") * tocLineCnt Then
      Call 中央線設定(tocLineCnt)
      
      colLine = setVal("Toc2StartColLine")
      If tocLineCnt = 1 Then
        RowCnt = 4
      Else
        RowCnt = setVal("PageLine") * WorksheetFunction.RoundDown(tocLineCnt / 2, 0) + 4
      End If
      tocLineCnt = tocLineCnt + 1
      
    ElseIf tocLineCnt Mod 2 = 0 And PageCnt = setVal("TocMaxCnt") * tocLineCnt Then
      colLine = setVal("Toc1StartColLine")
      RowCnt = setVal("PageLine") * WorksheetFunction.RoundDown(tocLineCnt / 2, 0) + 4
      tocLineCnt = tocLineCnt + 1
      
      'タイトルが目次でなければページ追加
      If Cells(RowCnt - 2, colLine + 2) Like "[目次,もくじ]*" Then
      Else
        Call addPage(Cells(RowCnt - 3, 1).Address)
      End If
      
      Range(Cells(RowCnt, colLine + 2), Cells(RowCnt, colLine + 2 + setVal("TocTitle"))).Select
    Else
      RowCnt = RowCnt + 1
    End If
    PageCnt = PageCnt + 1
    
    Call Ctl_ProgressBar.ShowCount("目次生成中", line, endLine, "P." & PageCnt & " " & tocTitle)
  Next

  '処理終了--------------------------------------
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  Call Ctl_ProgressBar.ShowEnd
  Call Library.EndScript
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function

'==================================================================================================
Function addPage(Optional targetCell As String)
  Dim line As Long, endLine As Long
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  FuncName = "Specification.addPage"

'  Call Library.StartScript
  Call init.setting
'  Call Ctl_ProgressBar.ShowStart
  Call Library.showDebugForm(FuncName & "============================================")
  '----------------------------------------------
  
  If targetCell = "" Then
    sheetCopy.Range("A1:AZ43").Copy
    endLine = Cells(Rows.count, 50).End(xlUp).Row + 1
    sheetMain.Range("A" & endLine).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Else
    sheetCopy.Range("A44:AZ86").Copy
    sheetMain.Range(targetCell).Insert Shift:=xlDown
  End If
  Application.CutCopyMode = False


  '処理終了--------------------------------------
  If targetCell = "" Then
    Application.Goto Reference:=Range("A" & endLine), Scroll:=True
    Range("D" & endLine + 1).Select
  Else
    Cells(Range(targetCell).Row + 1, 4) = "目次" & WorksheetFunction.RoundDown(Range(targetCell).Row / 44, 0) + 1
    'Cells(Range(targetCell).Row + 1, 4) = "目次"
    
    Application.Goto Reference:=Range(targetCell), Scroll:=True
  End If
  Call Library.showDebugForm("=================================================================")
'  Call Ctl_ProgressBar.ShowEnd
'  Call Library.EndScript
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function


'==================================================================================================
Function 中央線設定(tocPageCnt As Long)
  Dim line As Long
  
  If tocPageCnt = 1 Then
    Range("X4:Y42").Select
  Else
    line = 43 * WorksheetFunction.RoundDown(tocPageCnt / 2, 0) + 4
    Range("X" & line & ":Y" & line + 38).Select
  End If
  
  Selection.Borders(xlDiagonalDown).LineStyle = xlNone
  Selection.Borders(xlDiagonalUp).LineStyle = xlNone
  Selection.Borders(xlEdgeLeft).LineStyle = xlNone
  Selection.Borders(xlEdgeTop).LineStyle = xlNone
  Selection.Borders(xlEdgeBottom).LineStyle = xlNone
  Selection.Borders(xlEdgeRight).LineStyle = xlNone
  With Selection.Borders(xlInsideVertical)
    .LineStyle = xlDouble
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThick
  End With
  Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Function


'==================================================================================================
Sub リセット()
  Dim endLine As Long, line As Long
  Dim count As Long
  
  Call Library.StartScript
  Call init.setting(True)

  Call Library.showDebugForm(FuncName & "============================================")
  sheetMain.Select
  '----------------------------------------------
  sheetCopy.Range("A44:AZ86").Copy
  
  Range("A1:C43").Select
  ActiveSheet.Paste
  Range("A44:C44").Select
  ActiveSheet.Paste
  Range("A87:C87").Select
  ActiveSheet.Paste
  Range("A130:C130").Select
  ActiveSheet.Paste
  Range("A173:C173").Select
  ActiveSheet.Paste
  Range("A216:C216").Select
  ActiveSheet.Paste
  Range("A259:C259").Select
  ActiveSheet.Paste
  
  count = 1
  endLine = sheetMain.Cells(Rows.count, 50).End(xlUp).Row
  For line = 44 To endLine Step setVal("PageLine")
    Range("D" & line + 1) = "タイトル_" & count
    Range("S" & line + 1) = "機能_" & count
    
    count = count + 1

  Next
  Range("D2") = "目次"
  Range("S2") = ""
  
  
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.EndScript
End Sub
