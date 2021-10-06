Attribute VB_Name = "Ctl_ErImg"
'Option Explicit

Dim ER_imgCnt   As Long


'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function showUserForm()
  Dim line As Long, endLine As Long
  Dim objShp As Shape
  Dim fileAllCnt As Long, fileCnt As Long, fileName As String
  Dim tableListCnt As Integer, arrCnt As Integer
  
  '処理開始--------------------------------------
'  On Error GoTo catchError
  Const funcName As String = "Ctl_ErImg.showUserForm"
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
  End If
  Call Library.showDebugForm(funcName & "===========================================")
  '----------------------------------------------
  sheetTmp.Select
  Call Library.delSheetData
  
  sheetERImage.Select

  Set tableList = Nothing
  Set tableList = CreateObject("Scripting.Dictionary")
  
  Ctl_MySQL.dbOpen
  Call Ctl_MySQL.getDatabaseInfo(True)
  
  endLine = sheetTmp.Cells(Rows.count, 1).End(xlUp).Row
  With Frm_TableList
    .StartUpPosition = 1
    With .ListBox1
      .ColumnHeads = True
      .ColumnCount = 4
      .ColumnWidths = "20;150;150;120"
      .RowSource = "Tmp!A2:D" & endLine
    End With
    .Show
  End With
    
  Set tableList = Nothing

  sheetERImage.Select
  '処理終了--------------------------------------
  Ctl_MySQL.dbClose
  Application.GoTo Reference:=Range("A1"), Scroll:=True
  Call Ctl_ProgressBar.showEnd
  Call Library.showDebugForm("=================================================================")
  If runFlg = False Then
    Call Library.endScript
    Call init.unsetting
  End If
  '----------------------------------------------
  
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function



'==================================================================================================
Function deleteImages(targetShapeName As String)
  Dim line As Long, endLine As Long
  Dim objShp As Shape
  Dim fileAllCnt As Long, fileCnt As Long, fileName As String
  
  '処理開始--------------------------------------
'  On Error GoTo catchError
  Const funcName As String = "Ctl_ErImg.deleteImages"
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
    Call Library.showDebugForm(funcName & "=============================================")
  End If
  
  '----------------------------------------------
  sheetERImage.Select

  fileCnt = 1
  fileAllCnt = ActiveSheet.Shapes.count
  For Each objShp In ActiveSheet.Shapes
    'If objShp.Name Like "ERImg-*" Then
    If objShp.Name = "ERImg-" & targetShapeName Then
      fileName = objShp.Name
      Call Library.showDebugForm("delete", objShp.Name)
      objShp.Delete
    Else
      Call Library.showDebugForm("対象外", objShp.Name)
    End If
    DoEvents
    Call Ctl_ProgressBar.showBar(thisAppName, 1, 2, fileCnt, fileAllCnt, fileName)
    fileCnt = fileCnt + 1
  Next
  
  '処理終了--------------------------------------
  Application.GoTo Reference:=Range("A1"), Scroll:=True
  If runFlg = False Then
    Call Library.showDebugForm("=================================================================")
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.unsetting
  End If
  '----------------------------------------------
  
  Exit Function

'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function


'==================================================================================================
Function makeTable(tableName As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_ErImg.make"

  'runFlg = True
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
    Call Library.showDebugForm("runFlg", runFlg)
    PrgP_Cnt = 0
  End If
  Call Library.showDebugForm(funcName & "==================================================")
  '----------------------------------------------
  Call Library.showDebugForm("tableName", tableName)
  Call Library.showDebugForm("PrgP_Cnt", PrgP_Cnt)
  
  sheetSetting.Activate
  sheetSetting.Shapes.Range(Array("ERImg")).Select
  
  'テーブル名設定
  sheetSetting.Shapes.Range(Array("TableName")).Select
  Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = tableName

  'カラム名をリセット
  sheetSetting.Shapes.Range(Array("ColumnList")).Select
  Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = ""


  '処理終了--------------------------------------
'  Application.GoTo Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function


'==================================================================================================
Function makeColumnList(physicalName As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim columnName As String
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_ErImg.makeColumnList"

  'runFlg = True
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
    Call Library.showDebugForm("runFlg", runFlg)
    PrgP_Cnt = 0
    Call Library.showDebugForm(funcName & "=========================================")
  End If
  '----------------------------------------------
 
  
  Call Ctl_MySQL.getColumnInfo(physicalName, True)
  
  sheetSetting.Activate
  sheetSetting.Shapes.Range(Array("ERImg")).Select
  sheetSetting.Shapes.Range(Array("ColumnList")).Select
  
  For line = 0 To UBound(lValues)
    Call Library.showDebugForm("lValues", lValues(line, 0))
      If setVal("useLogicalName") = True Then
        columnName = lValues(line, 0)
      Else
        columnName = lValues(line, 1)
      End If
      
      If setVal("useLogicalName") = True Then
        columnName = lValues(line, 0)
      Else
        columnName = lValues(line, 1)
      End If
  
    If Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "" Then
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = columnName
    Else
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text & vbNewLine & columnName
    End If

'    If line = 9 And UBound(lValues) >= 10 And setVal("useImage") = True Then
'      Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text & vbNewLine & "　　　　　　　　　　　　　<続きあり>"
'      Exit For
'    End If
  
  Next
  

  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.showDebugForm("=================================================================")
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function


'==================================================================================================
Function copy(tableName As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim rowLine As Long
  
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_ErImg.copy"

  'runFlg = True
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
    Call Library.showDebugForm("runFlg", runFlg)
    PrgP_Cnt = 0
  End If
'  Call Library.showDebugForm(funcName & "==================================================")
  Call Library.showDebugForm("ER_imgCnt", ER_imgCnt)
  '----------------------------------------------
'  colLine = 2 + 10 * ER_imgCnt
'
'  If colLine >= 44 Then
'    rowLine = (5 + 13) * Int(ER_imgCnt / 5)
'    colLine = 2
'  Else
'    rowLine = 5
'  End If
'  Call Library.showDebugForm("rowLine", rowLine)
'  Call Library.showDebugForm("colLine", colLine)
  
  Call Library.showDebugForm("useImage", setVal("useImage"))
    
  sheetSetting.Shapes.Range(Array("ERImg")).Select
  Selection.copy
  sheetERImage.Select
  Cells(6, 3).Select
  If setVal("useImage") = True Then
    sheetERImage.Pictures.Paste.Select
    With Selection.ShapeRange.line
      .Visible = msoTrue
      .ForeColor.ObjectThemeColor = msoThemeColorBackground1
      .ForeColor.TintAndShade = 0
      .ForeColor.Brightness = -0.5
      .Weight = 1.5
      .Transparency = 0
    End With
  
  Else
    ActiveSheet.Paste
  End If
  Selection.Name = "ERImg-" & tableName
  ER_imgCnt = ER_imgCnt + 1
  
  

  '処理終了--------------------------------------
'  Call Library.showDebugForm("=================================================================")
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.unsetting
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function

