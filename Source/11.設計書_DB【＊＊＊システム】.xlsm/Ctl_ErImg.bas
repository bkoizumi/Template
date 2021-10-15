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
  Dim line As Long, endLine As Long, tmpLine As Long
  Dim tempSheet As Object
  
  '処理開始--------------------------------------
'  On Error GoTo catchError
  Const funcName As String = "Ctl_ErImg.showUserForm"
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
  End If
  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------
  sheetTmp.Select
  Call Library.delSheetData(2)
  
  sheetERImage.Select
  tmpLine = 2
  For Each tempSheet In Sheets
    If tempSheet.Name Like "3.*" Then
      Set targetSheet = Sheets(tempSheet.Name)
      sheetTmp.Range("J" & tmpLine) = tempSheet.Name
      sheetTmp.Range("K" & tmpLine) = targetSheet.Range("F7")
      sheetTmp.Range("L" & tmpLine) = targetSheet.Range("F8")
      sheetTmp.Range("M" & tmpLine) = targetSheet.Range("F9")
      Set targetSheet = Nothing
      tmpLine = tmpLine + 1
    End If
  Next
  
  With Frm_TableList
    .StartUpPosition = 1
    .Show
  End With
    
  sheetERImage.Select
'  ActiveWindow.Zoom = 70
  
  '処理終了--------------------------------------
  Ctl_MySQL.dbClose
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Ctl_ProgressBar.showEnd
  Call Library.showDebugForm("EndFun  ", funcName, "info")
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
Function getSheetData()
  Dim line As Long, endLine As Long
  Dim fileAllCnt As Long, fileCnt As Long, fileName As String
  Dim tableListCnt As Integer, arrCnt As Integer
  
  Dim isShapesImg As Boolean
  
  
  '処理開始--------------------------------------
'  On Error GoTo catchError
  Const funcName As String = "Ctl_ErImg.getSheetData"
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
  End If
  PrgP_Max = Worksheets.count
  PrgP_Cnt = 1
  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------
  sheetTmp.Select
  Call Library.delSheetData(2)
  
  sheetERImage.Select
  isShapesImg = False
  
  'シェイプにないシート情報を表示
  For Each tempSheet In Sheets
    If tempSheet.Name Like "3.*" Then
      For Each objShp In ActiveSheet.Shapes
        If objShp.Name = "ERImg-" & tempSheet.Name Then
          fileName = objShp.Name
          Call Library.showDebugForm("ER図", objShp.Name)
          isShapesImg = True
          Exit For
        End If
      Next
  
      If isShapesImg = False Then
        Call Library.showDebugForm("tempSheet.Name", tempSheet.Name)
        Call Ctl_ErImg.getColumnInfo(tempSheet.Name)
        Call Ctl_ErImg.makeERImage
        Call Ctl_ErImg.copy
        
        sheetTmp.Select
        Call Library.delSheetData(2)
      
      End If
    
    End If
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, 1, 2, "ER図生成")
    PrgP_Cnt = PrgP_Cnt + 1
  Next
  
  sheetERImage.Select
  '処理終了--------------------------------------
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Ctl_ProgressBar.showEnd
  Call Library.showDebugForm("EndFun  ", funcName, "info")
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
Function deleteImages()
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
    Call Library.showDebugForm("StartFun", funcName, "info")
  End If
  
  '----------------------------------------------
  sheetERImage.Select

  fileCnt = 1
  fileAllCnt = ActiveSheet.Shapes.count
  For Each objShp In ActiveSheet.Shapes
    If objShp.Name Like "ERImg-*" Then
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
  Application.Goto Reference:=Range("A1"), Scroll:=True
  If runFlg = False Then
    Call Library.showDebugForm("EndFun  ", funcName, "info")
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
Function getColumnInfo(SheetName As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim tmpLine As Long
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_ErImg.getColumnInfo"

  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------
  Call Library.showDebugForm("SheetName", SheetName)
  
  Set targetSheet = Sheets(SheetName)
  Call Library.showDebugForm("TableName", targetSheet.Range("F9"))
  targetSheet.Activate
  Call Ctl_Common.chkRowStartLine
  
  'カラム情報取得--------------------------------
  tmpLine = 2
  sheetTmp.Range("A2") = targetSheet.Range("F9")
  sheetTmp.Range("A3") = targetSheet.Range("F8")
  sheetTmp.Range("A4") = targetSheet.Range("G10")
  
  
  For line = startLine To setLine("columnEnd")
    sheetTmp.Range("B" & tmpLine) = targetSheet.Range("L" & line)
    sheetTmp.Range("C" & tmpLine) = targetSheet.Range("B" & line)
    sheetTmp.Range("D" & tmpLine) = targetSheet.Range("V" & line)
    sheetTmp.Range("E" & tmpLine) = targetSheet.Range("AF" & line).Value
    sheetTmp.Range("F" & tmpLine) = targetSheet.Range("AH" & line)
    sheetTmp.Range("G" & tmpLine) = targetSheet.Range("AJ" & line)
    sheetTmp.Range("H" & tmpLine) = targetSheet.Range("AL" & line)
    
    tmpLine = tmpLine + 1
  Next
  
  '処理終了--------------------------------------
'  Application.GoTo Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("EndFun  ", funcName, "info")
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
Function makeERImage()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim columnName As String
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_ErImg.makeERImage"
  '----------------------------------------------
  
  sheetTmp.Select
  endLine = Cells(Rows.count, 2).End(xlUp).Row
  
  sheetCopyLine.Activate
  sheetCopyLine.Shapes.Range(Array("ERImg")).Select
  
  'テーブル名を設定
  sheetCopyLine.Shapes.Range(Array("TableName")).Select
  
  If setVal("useLogicalName") = True Then
    If sheetTmp.Range("A3") = "" Then
      Selection.ShapeRange(1).TextFrame2.TextRange.Characters.text = sheetTmp.Range("A2")
    Else
      Selection.ShapeRange(1).TextFrame2.TextRange.Characters.text = sheetTmp.Range("A3")
    End If
  Else
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.text = sheetTmp.Range("A2")
  End If
  'カラム名を設定
  sheetCopyLine.Shapes.Range(Array("ColumnList")).Select
  With Selection.ShapeRange
    If sheetTmp.Range("A4") = "マスターテーブル" Then
      .Fill.ForeColor.RGB = RGB(242, 220, 219)
    ElseIf sheetTmp.Range("A4") = "トランザクションテーブル" Or sheetTmp.Range("A3") = "ワークテーブル" Then
      .Fill.ForeColor.RGB = RGB(215, 228, 189)
    End If
  End With
  
  For line = 2 To endLine
    
    If setVal("useLogicalName") = True Then
      If sheetTmp.Range("C" & line).text = "" Then
        columnName = sheetTmp.Range("B" & line).text
      Else
        columnName = sheetTmp.Range("C" & line).text
      End If
    Else
      columnName = sheetTmp.Range("C" & line).text
    End If
    
    'PK設定
    If sheetTmp.Range("E" & line) <> "" Then
      columnName = setVal("Char_PK") & columnName
    
    ElseIf sheetTmp.Range("H" & line) <> "" Then
      columnName = setVal("Char_NotNull") & columnName
    
    Else
      columnName = "　" & columnName
    End If
    
    '外部キー設定
    If sheetTmp.Range("G" & line) <> "" Then
      columnName = columnName & " [FK]"
    End If
    
    
    If Selection.ShapeRange(1).TextFrame2.TextRange.Characters.text = "" Then
      Selection.ShapeRange(1).TextFrame2.TextRange.Characters.text = columnName
    Else
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters.text = Selection.ShapeRange(1).TextFrame2.TextRange.Characters.text & vbNewLine & columnName
    End If
    
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, line, endLine, "ER図生成")
  Next
  

  '処理終了--------------------------------------
  If runFlg = False Then
    Call Library.showDebugForm("EndFun  ", funcName, "info")
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
Function copy()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim rowLine As Long
  Dim objShp
  Dim objShpCnt As Long
  Dim objShpCntR As Long, objShpCntC As Long
  
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_ErImg.copy"

  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------
  Call Library.showDebugForm("useImage", setVal("useImage"))
    
  sheetCopyLine.Shapes.Range(Array("ERImg")).Select
  Selection.copy
  sheetERImage.Select
  
'  objShpCnt = 0
'  objShpCntR = 20
'  objShpCntC = 2
'  For Each objShp In ActiveSheet.Shapes
'    If objShp.Name Like "ERImg-*" Then
'      Call Library.showDebugForm("ER図", objShp.Name)
'      objShpCnt = objShpCnt + 1
'    End If
'  Next
'  If objShpCnt = 0 Then
'    objShpCntR = 20
'    objShpCntC = 2
'  ElseIf objShpCnt <= 3 Then
'    objShpCntR = 20
'    objShpCntC = (12 * objShpCnt) + objShpCntC
'
'  ElseIf objShpCnt = 4 Then
'    objShpCntR = 40
'    objShpCntC = 5
'    objShpCntC = 2
'  ElseIf objShpCnt <= 6 Then
'    objShpCntR = 40
'    objShpCntC = (12 * objShpCnt) + objShpCntC
'
'  End If
'  Cells(objShpCntR, objShpCntC).Select
  Range("B20").Select
  
  
  Call Library.waitTime(500)
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
  Selection.Name = "ERImg-" & sheetTmp.Range("A2")
  
  'ER図をクリア----------------------------------
  sheetCopyLine.Activate
  sheetCopyLine.Shapes.Range(Array("ERImg")).Select
  
  'テーブル名を設定
  sheetCopyLine.Shapes.Range(Array("TableName")).Select
  Selection.ShapeRange(1).TextFrame2.TextRange.Characters.text = "テーブル名"
  
  'カラム名を設定
  sheetCopyLine.Shapes.Range(Array("ColumnList")).Select
  Selection.ShapeRange(1).TextFrame2.TextRange.Characters.text = ""
  

  '処理終了--------------------------------------
'  Call Library.showDebugForm("=================================================================",,"info")
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
  Dim ShapesInfo(), ShapeInfo
  Dim connectTypeS As Integer, connectTypeE As Integer
  
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
  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------
  Call Library.showDebugForm("lineType", lineType)
  
  If typeName(Selection) = "Range" Then
    startCell = ActiveCell.Address
    endCell = ActiveCell.Offset(, 3).Address
    
  ElseIf typeName(Selection) = "DrawingObjects" Then
    ReDim slctImgs(Selection.count - 1)
    ReDim ShapesInfo(Selection.count - 1, 5)
    
    slctImgCnt = 0
    For Each slctImg In Selection
      Call Library.showDebugForm("slctImg", slctImg.Name)
      slctImgs(slctImgCnt) = slctImg.Name
      
      ShapesInfo(slctImgCnt, 0) = slctImg.Name
      ShapesInfo(slctImgCnt, 1) = ActiveSheet.Shapes(slctImg.Name).Top
      ShapesInfo(slctImgCnt, 2) = ActiveSheet.Shapes(slctImg.Name).Left
      ShapesInfo(slctImgCnt, 3) = ActiveSheet.Shapes(slctImg.Name).Width
      ShapesInfo(slctImgCnt, 4) = ActiveSheet.Shapes(slctImg.Name).TopLeftCell.Offset(3, -1).Address(False, False)
      ShapesInfo(slctImgCnt, 5) = ActiveSheet.Shapes(slctImg.Name).TopLeftCell.Offset(3, 9).Address(False, False)
      
      slctImgCnt = slctImgCnt + 1
    Next
  End If
  Call Library.showDebugForm("slctImgCnt   ", slctImgCnt - 1, "info")
  
  For ShapeInfo = LBound(ShapesInfo, 1) To UBound(ShapesInfo, 1)
    Call Library.showDebugForm("Name        ", ShapesInfo(ShapeInfo, 0), "info")
    Call Library.showDebugForm("Top         ", ShapesInfo(ShapeInfo, 1), "info")
    Call Library.showDebugForm("Left        ", ShapesInfo(ShapeInfo, 2), "info")
    Call Library.showDebugForm("Width       ", ShapesInfo(ShapeInfo, 3), "info")
    Call Library.showDebugForm("TopLeftCell1", ShapesInfo(ShapeInfo, 4), "info")
    Call Library.showDebugForm("TopLeftCell2", ShapesInfo(ShapeInfo, 5), "info")
  Next
  If slctImgCnt = 2 Then
    If ShapesInfo(0, 2) < ShapesInfo(1, 2) Then
      Call Library.showDebugForm("左位置", "後のほうが右にある", "info")
      startCell = ShapesInfo(0, 5)
      endCell = ShapesInfo(1, 4)
      connectTypeS = 4
      connectTypeE = 4
    Else
      Call Library.showDebugForm("左位置", "後のほうが左にある", "info")
      startCell = ShapesInfo(0, 4)
      endCell = ShapesInfo(1, 5)
      connectTypeS = 2
      connectTypeE = 2
    
    End If
  End If
  Call Library.showDebugForm("connectTypeS", connectTypeS, "info")
  Call Library.showDebugForm("connectTypeE", connectTypeE, "info")
  
  Call Library.showDebugForm("startCell", startCell)
  Call Library.showDebugForm("endCell  ", endCell)
  
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
  
  
  
  sheetCopyLine.Select
  sheetCopyLine.Shapes.Range(Array(startImg)).Select
  Selection.copy

  sheetERImage.Select
  Range(startCell).Select
  sheetERImage.Pictures.Paste.Select
  Call Library.waitTime(100)
  ERImg_LineSName = "ERImg_LineS_" & counter
  Selection.Name = ERImg_LineSName
  If connectTypeS = 2 Then
    Selection.ShapeRange.Flip msoFlipHorizontal
    connectTypeS = 4
  End If
  
  sheetCopyLine.Select
  sheetCopyLine.Shapes.Range(Array(endImg)).Select
  Call Library.waitTime(100)
  Selection.copy
  
  sheetERImage.Select
  Range(endCell).Select
  sheetERImage.Pictures.Paste.Select
  Call Library.waitTime(100)
  ERImg_LineEName = "ERImg_LineE_" & counter
  Selection.Name = ERImg_LineEName
  If connectTypeE = 2 Then
    connectTypeE = 4
  Else
    Selection.ShapeRange.Flip msoFlipHorizontal
    connectTypeE = 4
  End If

  ActiveSheet.Shapes.AddConnector(msoConnectorElbow, 251, 92, 408, 128).Select
  Selection.Name = "ERImg_Line_" & counter
  
  Selection.ShapeRange.ConnectorFormat.BeginConnect ActiveSheet.Shapes(ERImg_LineSName), connectTypeS
  Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet.Shapes(ERImg_LineEName), connectTypeE
  
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
'  Application.Goto Reference:=Range("A1"), Scroll:=True
'  Call Library.showDebugForm("=================================================================",,"info")
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


