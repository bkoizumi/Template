Attribute VB_Name = "Ctl_Ribbon"
#If VBA7 And Win64 Then
  Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cbLen As LongPtr)
#Else
  Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cbLen As Long)
#End If


'**************************************************************************************************
' * リボンメニュー初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'読み込み時処理
Function onLoad(ribbon As IRibbonUI)
  Call init.Setting(True)
  
  Set ribbonUI = ribbon
  
  Call Library.setRegistry("Main", "DB_ribbonUI", CStr(ObjPtr(ribbonUI)))
  
  ribbonUI.ActivateTab ("DBTab")
  ribbonUI.Invalidate
  
End Function


'==================================================================================================
'更新
Function Refresh()
  Call init.Setting
  
  #If VBA7 And Win64 Then
    Set ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "DB_ribbonUI")))
  #Else
    Set ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "DB_ribbonUI")))
  #End If
  
  ribbonUI.ActivateTab ("DBTab")
  ribbonUI.Invalidate
End Function
  
  
'==================================================================================================
'シート一覧メニュー
Function getSheetsList(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, subMenu As Object
  Dim sheetName As Worksheet
  
  Call init.Setting
   
  If ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "DB_ribbonUI")))
    #Else
      Set ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "DB_ribbonUI")))
    #End If
  End If
  
  Set DOMDoc = CreateObject("Msxml2.DOMDocument")
  Set Menu = DOMDoc.createElement("menu")

  Menu.SetAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.SetAttribute "itemSize", "normal"

  For Each sheetName In ActiveWorkbook.Sheets
    Select Case sheetName.Name
      Case "設定", "Notice", "DataType", "コピー用"
      Case Else
        Set Button = DOMDoc.createElement("button")
        With Button
          sheetNameID = sheetName.Name
          .SetAttribute "id", encode(sheetName.Name)
          .SetAttribute "label", sheetName.Name
        
        If Sheets(sheetName.Name).Visible = True Then
          .SetAttribute "imageMso", "HeaderFooterSheetNameInsert"
        ElseIf Sheets(sheetName.Name).Visible <> True Then
          .SetAttribute "imageMso", "SheetProtect"
        
        End If
        If ActiveWorkbook.ActiveSheet.Name = sheetName.Name Then
          .SetAttribute "imageMso", "ExcelSpreadsheetInsert"
        End If
          .SetAttribute "onAction", "Ctl_Ribbon.selectActiveSheet"
        End With
        Menu.AppendChild Button
        Set Button = Nothing
    
    End Select
  Next
  DOMDoc.AppendChild Menu
  
  'Call Library.showDebugForm(DOMDoc.XML)
  
  returnedVal = DOMDoc.XML
  Set Menu = Nothing
  Set DOMDoc = Nothing
  
   ribbonUI.Invalidate
End Function

'--------------------------------------------------------------------------------------------------
Function dMenuRefresh(control As IRibbonControl)
  
  If ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "DB_ribbonUI")))
    #Else
      Set ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "DB_ribbonUI")))
    #End If
  End If
  ribbonUI.Invalidate
End Function


'--------------------------------------------------------------------------------------------------
Function selectActiveSheet(control As IRibbonControl)
  Dim sheetNameID As String
  Dim sheetCount As Integer
  Dim sheetName As Worksheet
  
  Call Library.startScript
  sheetNameID = decode(control.ID)
  
  If Sheets(sheetNameID).Visible <> True Then
    Sheets(sheetNameID).Visible = True
  End If
  
  sheetCount = 1
  For Each sheetName In ActiveWorkbook.Sheets
    If Sheets(sheetName.Name).Visible = True And sheetName.Name = sheetNameID Then
      Exit For
    Else
      sheetCount = sheetCount + 1
    End If
  Next
  ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
  ActiveWindow.ScrollWorkbookTabs Sheets:=sheetCount
  Sheets(sheetNameID).Select
  
  Application.Goto Reference:=Range("A1"), Scroll:=True
  
  Call Library.endScript
End Function

'--------------------------------------------------------------------------------------------------
Function encode(strVal As String)

  strVal = Replace(strVal, "(", "bk-1-lib")
  strVal = Replace(strVal, ")", "bk-2-lib")
  strVal = Replace(strVal, " ", "bk-3-lib")
  strVal = Replace(strVal, "　", "bk-4-lib")
  strVal = Replace(strVal, "【", "bk-5-lib")
  strVal = Replace(strVal, "】", "bk-6-lib")
  
  strVal = "bk-0-lib" & strVal
  encode = strVal
End Function

'--------------------------------------------------------------------------------------------------
Function decode(strVal As String)

  strVal = Replace(strVal, "bk-0-lib", "")
  strVal = Replace(strVal, "bk-1-lib", "(")
  strVal = Replace(strVal, "bk-2-lib", ")")
  strVal = Replace(strVal, "bk-3-lib", " ")
  strVal = Replace(strVal, "bk-4-lib", "　")
  strVal = Replace(strVal, "bk-5-lib", "【")
  strVal = Replace(strVal, "bk-6-lib", "】")
  
  decode = strVal
End Function


'--------------------------------------------------------------------------------------------------
#If VBA7 And Win64 Then
Private Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
  Dim p As LongPtr
#Else
Private Function GetRibbon(ByVal lRibbonPointer As Long) As Object
  Dim p As Long
#End If
  Dim ribbonObj As Object
  
  MoveMemory ribbonObj, lRibbonPointer, LenB(lRibbonPointer)
  Set GetRibbon = ribbonObj
  p = 0: MoveMemory ribbonObj, p, LenB(p) '後始末
End Function





' お気に入りメニュー作成---------------------------------------------------------------------------
Function FavoriteMenu(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, subMenu As Object
  Dim regLists As Variant, i As Long
  Dim line As Long, endLine As Long
  Dim objFSO As New FileSystemObject
   
  Call init.Setting
   
  If ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("Main", "DB_ribbonUI")))
    #Else
      Set ribbonUI = GetRibbon(CLng(Library.getRegistry("Main", "DB_ribbonUI")))
    #End If
  End If
  
  Set DOMDoc = CreateObject("Msxml2.DOMDocument")
  Set Menu = DOMDoc.createElement("menu") ' menuの作成

  Menu.SetAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.SetAttribute "itemSize", "normal"

  endLine = sheetFavorite.Cells(Rows.count, 1).End(xlUp).Row
  For line = 2 To endLine
    Set Button = DOMDoc.createElement("button")
    With Button
      .SetAttribute "id", "Favorite_" & line
      .SetAttribute "label", objFSO.GetFileName(sheetFavorite.Range("A" & line))
      .SetAttribute "imageMso", "Favorites"
      .SetAttribute "onAction", "OpenFavoriteList"
    End With
    Menu.AppendChild Button
    Set Button = Nothing
  
  Next
  DOMDoc.AppendChild Menu
  returnedVal = DOMDoc.XML
'  Call Library.showDebugForm(DOMDoc.XML)
  
  Set Menu = Nothing
  Set DOMDoc = Nothing
End Function


'--------------------------------------------------------------------------------------------------
Function OpenFavoriteList(control As IRibbonControl)
  Dim fileNamePath As String
  Dim line As Long
  
  line = Replace(control.ID, "Favorite_", "")
  fileNamePath = sheetFavorite.Range("A" & line)
  
  If Library.chkFileExists(fileNamePath) Then
    Workbooks.Open fileName:=fileNamePath
  End If
  Application.Goto Reference:=Range("A1"), Scroll:=True
End Function





'Label 設定----------------------------------------------------------------------------------------
Public Sub getLabel(control As IRibbonControl, ByRef setRibbonVal)
  
  Call init.Setting
  setRibbonVal = Replace(ribbonVal("Lbl_" & control.ID), "<BR>", vbNewLine)
End Sub


'Action 設定---------------------------------------------------------------------------------------
Sub getAction(control As IRibbonControl)
  Dim setRibbonVal As Variant
  
  Call init.Setting
  setRibbonVal = ribbonVal("Act_" & control.ID)
  
  If setRibbonVal Like "*Ctl_Ribbon*" Then
    Call Application.run(setRibbonVal, control)
  
  ElseIf setRibbonVal = "" Then
    Call Library.showDebugForm("Act_" & control.ID)
  Else
    Call Application.run(setRibbonVal)
  End If
End Sub


'Supertip 設定-------------------------------------------------------------------------------------
Public Sub getSupertip(control As IRibbonControl, ByRef setRibbonVal)
  Call init.Setting
  setRibbonVal = ribbonVal("Sup_" & control.ID)
End Sub


'Description 設定----------------------------------------------------------------------------------
Public Sub getDescription(control As IRibbonControl, ByRef setRibbonVal)
  Call init.Setting
  setRibbonVal = Replace(ribbonVal("Dec_" & control.ID), "<BR>", vbNewLine)

End Sub

'getImageMso 設定----------------------------------------------------------------------------------
Public Sub getImage(control As IRibbonControl, ByRef image)
  Call init.Setting
  image = ribbonVal("Img_" & control.ID)
End Sub


'size 設定-----------------------------------------------------------------------------------------
Public Sub getSize(control As IRibbonControl, ByRef setRibbonVal)
  Dim getVal As String
  
  Call init.Setting
  setRibbonVal = ribbonVal("Siz_" & control.ID)
  Select Case setRibbonVal
    Case "large"
      setRibbonVal = 1
    Case "normal"
      setRibbonVal = 0
    Case Else
      setRibbonVal = 0
  End Select
End Sub

'--------------------------------------------------------------------------------------------------
'有効/無効切り替え
Function getEnabled(control As IRibbonControl, ByRef returnedVal)
  Dim wb As Workbook
  Call init.Setting
  
  If Workbooks.count = 0 Then
    returnedVal = False
  ElseIf setVal("debugMode") = "develop" Then
    returnedVal = True
  Else
    returnedVal = False
  End If
  
End Function


'--------------------------------------------------------------------------------------------------
Sub getVisible(control As IRibbonControl, ByRef returnedVal)
  Call init.Setting
  returnedVal = Library.getRegistry("CustomRibbon")
End Sub


'--------------------------------------------------------------------------------------------------
Function RefreshRibbon()
  #If VBA7 And Win64 Then
    Set ribbonUI = GetRibbon(CLngPtr(Library.getRegistry("ribbonUI")))
  #Else
    Set ribbonUI = GetRibbon(CLng(Library.getRegistry("ribbonUI")))
  #End If
  ribbonUI.Invalidate

End Function

'中央揃え------------------------------------------------------------------------------------------
Function setCenter(control As IRibbonControl)
  If typeName(Selection) = "Range" Then
    Selection.HorizontalAlignment = xlCenterAcrossSelection
  End If
End Function


'**************************************************************************************************
' * リボンメニュー[オプション]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function Optionshow(control As IRibbonControl)
  Call Ctl_Option.showOption
End Function

'==================================================================================================
Function ClearAll(control As IRibbonControl)
  
  '処理開始--------------------------------------
  FuncName = "Ctl_Ribbon.ClearAll"
  Call Library.startScript
  Call init.Setting
  Call Ctl_ProgressBar.showStart
  runFlg = True
  Call Library.showDebugForm(FuncName & "==============================================")
  '----------------------------------------------
  
  Call Ctl_Option.ClearAll
  
  '処理終了--------------------------------------
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
  Call init.usetting
  '----------------------------------------------
End Function

'**************************************************************************************************
' * リボンメニュー[共通]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function addSheet(control As IRibbonControl)
  Call Ctl_Sheet.showAddSheetOption
  Call Ctl_Sheet.addSheet
End Function




'**************************************************************************************************
' * リボンメニュー[DB操作]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function getDatabaseInfo(control As IRibbonControl)
  
  '処理開始--------------------------------------
  FuncName = "Ctl_Ribbon.getDatabaseInfo"
  Call Library.startScript
  Call init.Setting
  Call Ctl_ProgressBar.showStart
  runFlg = True
  Call Library.showDebugForm(FuncName & "=========================================")
  '----------------------------------------------

  Select Case setVal("DBMS")
    Case "MSAccess"
      'Call Ctl_Access.getDatabaseInfo
      
    Case "MySQL"
      Call Ctl_MySQL.dbOpen
      Call Ctl_MySQL.getDatabaseInfo
      Call Ctl_MySQL.dbClose
      
    Case "PostgreSQL"
      'Call Ctl_Access.getDatabaseInfo
      
    Case "SQLServer"
      'Call Ctl_SQLServer.getDatabaseInfo
      
    Case Else
  End Select
  
  '処理終了--------------------------------------
'  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
  Call init.usetting
  '----------------------------------------------
End Function

'==================================================================================================
Function getTableInfo(control As IRibbonControl)
  
  '処理開始--------------------------------------
  FuncName = "Ctl_Ribbon.getTableInfo"
  Call Library.startScript
  Call init.Setting
  Call Ctl_ProgressBar.showStart
  runFlg = True
  Call Library.showDebugForm(FuncName & "==========================================")
  '----------------------------------------------
  
  Select Case setVal("DBMS")
    Case "MSAccess"
      Call Ctl_Access.getTableInfo
      
    Case "MySQL"
      Call Ctl_MySQL.dbOpen
      Call Ctl_MySQL.getTableInfo
      Call Ctl_MySQL.dbClose
      
    Case "PostgreSQL"
      Ctl_Access.getTableInfo
      
    Case "SQLServer"
      Call Ctl_SQLServer.getTableInfo
      
    Case Else
  End Select
  
  '処理終了--------------------------------------
  
'  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
  Call init.usetting
  '----------------------------------------------
  
End Function


'==================================================================================================
Function CreateTableInfo(control As IRibbonControl)
  
  '処理開始--------------------------------------
  FuncName = "Ctl_Ribbon.CreateTableInfo"
  Call Library.startScript
  Call init.Setting
  Call Ctl_ProgressBar.showStart
  runFlg = True
  Call Library.showDebugForm(FuncName & "=======================================")
  '----------------------------------------------
  
  Select Case setVal("DBMS")
    Case "MSAccess"
      Call Ctl_Access.CreateTable
      
    Case "MySQL"
      Call Ctl_MySQL.dbOpen
      Call Ctl_MySQL.CreateTable
      Call Ctl_MySQL.dbClose
      
    Case "PostgreSQL"
'      Ctl_Access.getTableInfo
      
    Case "SQLServer"
'      Call Ctl_SQLServer.CreateTable
      
    Case Else
  End Select
  
  '処理終了--------------------------------------
  
'  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
  Call init.usetting
  '----------------------------------------------
  
End Function
