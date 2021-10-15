Attribute VB_Name = "common"
Sub Macro4()
    
  Call Library.startScript
    Sheets("Master").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=31
    Sheets(Array("Master", "MC_案件管理表", "MC_積上管理表", "SI_チーム別マスタ", "SI_プロジェクト一覧", _
        "SI_プロジェクト採算見込", "SI_前週案件一覧", "異動社員情報", "過年度データ", "活動費_月別会議費", "活動費_月別交際費_個顧別", _
        "活動費_月別交際費_個人別", "活動費_月別交通費", "活動費_月別旅費出張費", "活動費_顧客データ", "活動費_精算データ", "活動費_同席者データ", _
        "活動費_不正検出_タクシー_回数", "活動費_不正検出_タクシー_金額", "活動費_不正検出_交際費_回数", "活動費_不正検出_交際費_金額", _
        "活動費_不正検出_交通費_金額", "活動費_不正検出_国内諸経費", "活動費_不正検出_旅費出張費_回数", "管理会計コード")).Select
    Sheets("Master").Activate
    Sheets(Array("実績", "社員情報", "社内経費試算表", "全社員情報", "全社組織コード", "組織マスタ", "直接間接比率", "年計データ", _
        "要員数")).Select Replace:=False
    Sheets(Array("Master", "MC_案件管理表", "MC_積上管理表", "SI_チーム別マスタ", "SI_プロジェクト一覧", _
        "SI_プロジェクト採算見込", "SI_前週案件一覧", "異動社員情報", "過年度データ", "活動費_月別会議費", "活動費_月別交際費_個顧別", _
        "活動費_月別交際費_個人別", "活動費_月別交通費", "活動費_月別旅費出張費", "活動費_顧客データ", "活動費_精算データ", "活動費_同席者データ", _
        "活動費_不正検出_タクシー_回数", "活動費_不正検出_タクシー_金額", "活動費_不正検出_交際費_回数", "活動費_不正検出_交際費_金額", _
        "活動費_不正検出_交通費_金額", "活動費_不正検出_国内諸経費", "活動費_不正検出_旅費出張費_回数", "要員数")).Select
    Sheets("要員数").Activate
    Sheets(Array("管理会計コード", "実績", "社員情報", "社内経費試算表", "全社員情報", "全社組織コード", "組織マスタ", "直接間接比率" _
        , "年計データ")).Select Replace:=False
    ActiveWindow.SelectedSheets.Delete

  Call Library.endScript
End Sub


'**************************************************************************************************
' * 共通処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function シートクリア()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
'  On Error GoTo catchError

  endLine = Cells(Rows.count, 2).End(xlUp).Row
  
  On Error Resume Next
  Rows(startLine & ":" & endLine).SpecialCells(xlCellTypeConstants, 23).ClearContents
  

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, Err.Description, True)
End Function


'==================================================================================================
Function シート追加(newSheetName As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
'  On Error GoTo catchError

    sheetCopyTable.copy After:=Worksheets(Worksheets.count)
    ActiveWorkbook.Sheets(Worksheets.count).Tab.ColorIndex = -4142
    ActiveWorkbook.Sheets(Worksheets.count).Name = newSheetName
    Sheets(newSheetName).Select
    
    
    
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, Err.Description, True)
End Function


'==================================================================================================
Function TBLリストシート生成()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim sheetList As Object
  Dim targetSheet   As Worksheet
  Dim SheetName As String
  Dim result As Boolean
  
'  On Error GoTo catchError

'処理開始--------------------------------------
  Call init.Setting
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  
  sheetTblList.Select
  endLine = sheetTblList.Cells(Rows.count, 3).End(xlUp).Row + 1
  Range("B6:U" & endLine).ClearContents
  
  With Range("B6:U" & endLine).Interior
    .Pattern = xlPatternNone
    .Color = xlNone
  End With
  
      
  line = 6
  For Each sheetList In ActiveWorkbook.Sheets
    SheetName = sheetList.Name
    result = Library.chkExcludeSheet(SheetName)
    
    If result = True Then
    Else
      sheetTblList.Range("B" & line).FormulaR1C1 = "=ROW()-5"
      sheetTblList.Range("C" & line) = Sheets(SheetName).Range("B2")
'      sheetTblList.Range("E" & line) = Sheets(sheetName).Range("D5")
'      sheetTblList.Range("H" & line) = Sheets(sheetName).Range("H5")
      sheetTblList.Range("Q" & line) = Sheets(SheetName).Range("D6")
    
      '論理テーブル名
      If Sheets(SheetName).Range("D5") <> "" Then
        With sheetTblList.Range("E" & line)
          .Value = Sheets(SheetName).Range("D5")
          .Select
          .Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=SheetName & "!" & "A9"
          .Font.Color = RGB(0, 0, 0)
          .Font.Underline = False
          .Font.Size = 10
          .Font.Name = "メイリオ"
        End With
      End If
      
       '物理テーブル名
      With sheetTblList.Range("H" & line)
        .Value = Sheets(SheetName).Range("H5")
        .Select
        .Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=SheetName & "!" & "A9"
        .Font.Color = RGB(0, 0, 0)
        .Font.Underline = False
        .Font.Size = 10
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
  Next
  
  
  '処理終了--------------------------------------
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
    
    
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, Err.Description, True)
End Function


