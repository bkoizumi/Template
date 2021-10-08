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
  
  
  With Frm_TableList
    .StartUpPosition = 1
    .Show
  End With
    
  Set tableList = Nothing

  sheetERImage.Select
  ActiveWindow.Zoom = 100
  
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
    If objShp.Name = "ERImg-" & targetShapeName Then
      fileName = objShp.Name
      Call Library.showDebugForm("delete", objShp.Name)
      objShp.Delete
    ElseIf objShp.Name Like "ERImg_Line*" Then
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
        columnName = lValues(line, 2) & lValues(line, 0)
      Else
        columnName = lValues(line, 2) & lValues(line, 1)
      End If
      
      If setVal("useLogicalName") = True Then
        columnName = lValues(line, 2) & lValues(line, 0)
      Else
        columnName = lValues(line, 2) & lValues(line, 1)
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
    Call Library.waitTime(100)
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


'==================================================================================================
Function ConnectLine(lineType As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim rowLine As Long
  Dim slctImg, slctImgs(), slctImgCnt As Integer
  Dim counter As String
  Dim startCell As String, endCell As String
  Dim ERImg_LineSName As String, ERImg_LineEName As String
  
  Dim startImg As String, endImg As String
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_ErImg.conLine"

  'runFlg = True
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
'    Call Ctl_ProgressBar.showStart
'    Call Library.showDebugForm("runFlg", runFlg)
'    PrgP_Cnt = 0
  End If
'  Call Library.showDebugForm(funcName & "==================================================")
  '----------------------------------------------
  Call Library.showDebugForm("lineType", lineType)
  
'  ReDim slctImgs(Selection.count - 1)
'  slctImgCnt = 0
'  For Each slctImg In Selection
'    Call Library.showDebugForm("slctImg", slctImg.Name)
'    slctImgs(slctImgCnt) = slctImg.Name
'    If slctImgCnt = 0 Then
'      startCell = ActiveSheet.Shapes(slctImg.Name).TopLeftCell.Offset(, -1).Address
'    Else
'      endCell = ActiveSheet.Shapes(slctImg.Name).TopLeftCell.Offset(, -1).Address
'    End If
'
'    slctImgCnt = slctImgCnt + 1
'  Next
'  Call Library.showDebugForm("startCell", startCell)
'  Call Library.showDebugForm("endCell  ", endCell)
  
  If typeName(ActiveCell) = "Range" Then
    startCell = ActiveCell.Address
    endCell = ActiveCell.Offset(, 3).Address
  Else
    startCell = "C4"
    endCell = "F4"
  End If
  counter = ActiveSheet.Shapes.count + 1
  
  Select Case lineType
    Case "ERLine1"
      startImg = "ERImg_1"
      endImg = "ERImg_1"
    
    Case "ERLine2"
      startImg = "ERImg_1"
      endImg = "ERImg_N"
    
    Case "ERLine3"
      startImg = "ERImg_1"
      endImg = "ERImg_0"
    
    Case "ERLine4"
      startImg = "ERImg_1"
      endImg = "ERImg_1N"
    
    Case "ERLine5"
      startImg = "ERImg_1N"
      endImg = "ERImg_1N"
    
    Case "ERLine6"
      startImg = "ERImg_01"
      endImg = "ERImg_1N"
    
    Case Else
  End Select
  
  
  
  sheetSetting.Select
  sheetSetting.Shapes.Range(Array(startImg)).Select
  Selection.copy

  sheetERImage.Select
  Range(startCell).Select
  sheetERImage.Pictures.Paste.Select
  Call Library.waitTime(100)
  ERImg_LineSName = "ERImg_LineS_" & counter
  Selection.Name = ERImg_LineSName
  
  sheetSetting.Select
  sheetSetting.Shapes.Range(Array(endImg)).Select
  Call Library.waitTime(100)
  Selection.copy
  
  sheetERImage.Select
  Range(endCell).Select
  sheetERImage.Pictures.Paste.Select
  Call Library.waitTime(100)
  ERImg_LineEName = "ERImg_LineE_" & counter
  Selection.Name = ERImg_LineEName
  Selection.ShapeRange.Flip msoFlipHorizontal

  ActiveSheet.Shapes.AddConnector(msoConnectorElbow, 251, 92, 408, 128).Select
  Selection.Name = "ERImg_Line_" & counter
  
  Selection.ShapeRange.ConnectorFormat.BeginConnect ActiveSheet.Shapes(ERImg_LineSName), 4
  Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet.Shapes(ERImg_LineEName), 4
  
'  For slctImgCnt = 0 To UBound(slctImgs)
'    slctImg = slctImgs(slctImgCnt)
'    Call Library.showDebugForm("slctImg", slctImg)
'    If slctImgCnt = 0 Then
'      ActiveSheet.Shapes.Range(Array("ERImg_LineS_" & counter, slctImg)).Select
'    ElseIf slctImgCnt = 1 Then
'      ActiveSheet.Shapes.Range(Array("ERImg_LineE_" & counter, slctImg)).Select
'    End If
'    Selection.ShapeRange.Group.Select
'    Selection.Name = "ERImg_" & Selection.Name
'  Next
  
  

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


