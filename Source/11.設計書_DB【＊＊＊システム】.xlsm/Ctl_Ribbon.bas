Attribute VB_Name = "Ctl_Ribbon"
#If VBA7 And Win64 Then
  Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cbLen As LongPtr)
#Else
  Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cbLen As Long)
#End If


'**************************************************************************************************
' * ���{�����j���[�����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'�ǂݍ��ݎ�����
Function onLoad(ribbon As IRibbonUI)
  Call init.Setting(True)
  
  Set ribbonUI = ribbon
  
  Call Library.setRegistry("Main", "DB_ribbonUI", CStr(ObjPtr(ribbonUI)))
  
  ribbonUI.ActivateTab ("DBTab")
  ribbonUI.Invalidate
  
End Function


'==================================================================================================
'�X�V
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
'�V�[�g�ꗗ���j���[
Function getSheetsList(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, subMenu As Object
  Dim SheetName As Worksheet
  
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

  For Each SheetName In ActiveWorkbook.Sheets
    If SheetName.Name Like "<*>" Then
    Else
      Set Button = DOMDoc.createElement("button")
      With Button
        sheetNameID = SheetName.Name
        .SetAttribute "id", encode(SheetName.Name)
        .SetAttribute "label", SheetName.Name
      
      If Sheets(SheetName.Name).Visible = True Then
        .SetAttribute "imageMso", "HeaderFooterSheetNameInsert"
      ElseIf Sheets(SheetName.Name).Visible <> True Then
        .SetAttribute "imageMso", "SheetProtect"
      
      End If
      If ActiveWorkbook.ActiveSheet.Name = SheetName.Name Then
        .SetAttribute "imageMso", "ExcelSpreadsheetInsert"
      End If
        .SetAttribute "onAction", "Ctl_Ribbon.selectActiveSheet"
      End With
      Menu.AppendChild Button
      Set Button = Nothing
    End If
  Next
  DOMDoc.AppendChild Menu
  
'  Call Library.showDebugForm(DOMDoc.XML)
  
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
  Dim SheetName As Worksheet
  
  Call Library.startScript
  sheetNameID = decode(control.ID)
  
  If Sheets(sheetNameID).Visible <> True Then
    Sheets(sheetNameID).Visible = True
  End If
  
  sheetCount = 1
  For Each SheetName In ActiveWorkbook.Sheets
    If Sheets(SheetName.Name).Visible = True And SheetName.Name = sheetNameID Then
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
  strVal = Replace(strVal, "�@", "bk-4-lib")
  strVal = Replace(strVal, "�y", "bk-5-lib")
  strVal = Replace(strVal, "�z", "bk-6-lib")
  
  strVal = "bk-0-lib" & strVal
  encode = strVal
End Function

'--------------------------------------------------------------------------------------------------
Function decode(strVal As String)

  strVal = Replace(strVal, "bk-0-lib", "")
  strVal = Replace(strVal, "bk-1-lib", "(")
  strVal = Replace(strVal, "bk-2-lib", ")")
  strVal = Replace(strVal, "bk-3-lib", " ")
  strVal = Replace(strVal, "bk-4-lib", "�@")
  strVal = Replace(strVal, "bk-5-lib", "�y")
  strVal = Replace(strVal, "bk-6-lib", "�z")
  
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
  p = 0: MoveMemory ribbonObj, p, LenB(p) '��n��
End Function



'==================================================================================================
'�J���p���j���[
Function setDeveloperMenu(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, subMenu As Object
  Dim menuList(3, 2) As String
  Dim i As Long
  
  menuList(0, 0) = "initialization"
  menuList(0, 1) = "������"
  menuList(0, 2) = "AccessRefreshAllLists"
  
  menuList(1, 0) = "Unprotect"
  menuList(1, 1) = "�ی����"
  menuList(1, 2) = "SheetProtect"
  
  menuList(2, 0) = "endScript"
  menuList(2, 1) = "��ʍX�V"
  menuList(2, 2) = "MacroPlay"
  
  menuList(3, 0) = "reset"
  menuList(3, 1) = "��蒼��"
  menuList(3, 2) = "MacroPlay"
  
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

  For i = 0 To UBound(menuList, 1)
    Set Button = DOMDoc.createElement("button")
    With Button
      .SetAttribute "id", menuList(i, 0)
      .SetAttribute "label", menuList(i, 1)
      .SetAttribute "imageMso", menuList(i, 2)
      .SetAttribute "onAction", "Ctl_Ribbon.runDeveloperMenu"
    End With
    
    Menu.AppendChild Button
    Set Button = Nothing
  Next
  
  
  DOMDoc.AppendChild Menu
  
'  Call Library.showDebugForm(DOMDoc.XML)
  
  returnedVal = DOMDoc.XML
  Set Menu = Nothing
  Set DOMDoc = Nothing
  
   ribbonUI.Invalidate
End Function

'--------------------------------------------------------------------------------------------------
Function runDeveloperMenu(control As IRibbonControl)
  
  Select Case control.ID
    Case "initialization"
      'Call Ctl_DeveloperMenu.initialization
    Case "Unprotect"
      'Call Ctl_DeveloperMenu.Unprotect
    Case "endScript"
      'Call Ctl_DeveloperMenu.endScript
    Case "reset"
      Call Ctl_DeveloperMenu.Reset
    Case Else
    
  End Select

End Function


'**************************************************************************************************
' * ���{�����j���[[�I�v�V����]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function showOption(control As IRibbonControl)
  
  '�����J�n--------------------------------------
  Const funcName As String = "Ctl_Ribbon.ClearAll"
  Call Library.startScript
  Call init.Setting
  runFlg = True
  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------
  Call Ctl_Option.showOption

  '�����I��--------------------------------------
  Call Library.endScript
  Call init.unsetting(True)
  '----------------------------------------------

End Function

'==================================================================================================
Function ClearAll(control As IRibbonControl)
  
  '�����J�n--------------------------------------
  Const funcName As String = "Ctl_Ribbon.ClearAll"
  Call Library.startScript
  Call init.Setting
  Call Ctl_ProgressBar.showStart
  runFlg = True
  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------
  
  Call Ctl_Option.ClearAll
  sheetCopyTable.Select
  
  '�����I��--------------------------------------
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
  Call init.unsetting(True)
  '----------------------------------------------
End Function

'**************************************************************************************************
' * ���{�����j���[[����]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'�V�[�g�ǉ�
Function addSheet(control As IRibbonControl)
  Call Ctl_Sheet.showAddSheetOption
  Call Ctl_Common.addSheet(sheetCopyTable.Range("F8").Value)
  
  sheetCopyTable.Range("F8") = ""
  sheetCopyTable.Range("F9") = ""
  sheetCopyTable.Range("F10") = ""
  sheetCopyTable.Range("F11") = ""
  sheetCopyTable.Range("AO3") = ""
  sheetCopyTable.Range("AO1") = ""
  sheetCopyTable.Range("AX1") = ""
  
  
End Function

'==================================================================================================
'�e�[�u�����X�g�X�V
Function makeTblList(control As IRibbonControl)
  Call Ctl_Common.makeTblList
End Function

'==================================================================================================
'ER�}�X�V
Function makeERImage(control As IRibbonControl)
  '�����J�n--------------------------------------
  Const funcName As String = "Ctl_Ribbon.getDatabaseInfo"
  Call Library.startScript
  Call init.Setting
  runFlg = True
  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------

  Call Ctl_ErImg.showUserForm
  
  
  '�����I��--------------------------------------
  Call Library.endScript
  Call init.unsetting(True)
  '----------------------------------------------

End Function

'==================================================================================================
'ER�}�p�R�l�N�^�[����
Function makeER_ConnectLine(control As IRibbonControl)
  '�����J�n--------------------------------------
  Const funcName As String = "Ctl_Ribbon.getDatabaseInfo"
  Call Library.startScript
  Call init.Setting
  runFlg = True
  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------
  
  
  Call Ctl_ErImg.ConnectLine(CStr(control.ID))
  
  Application.Goto Reference:=Range("A1"), Scroll:=True
  '�����I��--------------------------------------
  Call Library.endScript
  Call init.unsetting(True)
  '----------------------------------------------

End Function



'**************************************************************************************************
' * ���{�����j���[[DB����]
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'�ꊇ�擾
Function getDatabaseInfo(control As IRibbonControl)
  
  '�����J�n--------------------------------------
  Const funcName As String = "Ctl_Ribbon.getDatabaseInfo"
  Call Library.startScript
  Call init.Setting
  Call Ctl_ProgressBar.showStart
  runFlg = True
  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------

  Select Case setVal("DBMS")
    Case "MSAccess"
      'Call Ctl_Access.getDatabaseInfo
      
    Case "MySQL"
      Call Ctl_MySQL.dbOpen
      Call Ctl_MySQL.getDatabaseInfo
      Call Ctl_MySQL.dbClose
      
    Case "PostgreSQL"
      
    Case "SQLServer"
      Call Ctl_SQLServer.dbOpen
      Call Ctl_SQLServer.getDatabaseInfo
      Call Ctl_SQLServer.dbClose
      
    Case Else
  End Select
  
  '�����I��--------------------------------------
'  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
  Call init.unsetting(True)
  '----------------------------------------------
End Function

'==================================================================================================
'�A�N�e�B�u�V�[�g�̃e�[�u�����̂ݎ擾
Function getTableInfo(control As IRibbonControl)
  
  '�����J�n--------------------------------------
  Const funcName As String = "Ctl_Ribbon.getTableInfo"
  Call Library.startScript
  Call init.Setting
  Call Ctl_ProgressBar.showStart
  runFlg = True
  Call Library.showDebugForm("StartFun", funcName, "info")
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
      Call Ctl_SQLServer.dbOpen
      Call Ctl_SQLServer.getTableInfo
      Call Ctl_SQLServer.dbClose
      
    Case Else
  End Select
  
  '�����I��--------------------------------------
  
'  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
  Call init.unsetting(True)
  '----------------------------------------------
  
End Function


'==================================================================================================
Function CreateTableInfo(control As IRibbonControl)
  
  '�����J�n--------------------------------------
  Const funcName As String = "Ctl_Ribbon.CreateTableInfo"
  Call Library.startScript
  Call init.Setting
  Call Ctl_ProgressBar.showStart
  runFlg = True
  Call Library.showDebugForm("StartFun", funcName, "info")
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
      Call Ctl_SQLServer.dbOpen
      Call Ctl_SQLServer.CreateTable
      Call Ctl_SQLServer.dbClose

    Case Else
  End Select
  
  '�����I��--------------------------------------
  
'  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
  Call init.unsetting(True)
  '----------------------------------------------
  
End Function


'==================================================================================================
Function makeDDL(control As IRibbonControl)
  
  '�����J�n--------------------------------------
  Const funcName As String = "Ctl_Ribbon.CreateTableInfo"
  Call Library.startScript
  Call init.Setting
  Call Ctl_ProgressBar.showStart
  runFlg = True
  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------
  
  Select Case setVal("DBMS")
    Case "MSAccess"
    
    Case "MySQL"
      Call Ctl_MySQL.makeDDL
      
    Case "PostgreSQL"
      
    Case "SQLServer"
      Call Ctl_SQLServer.makeDDL
      
    Case Else
  End Select
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
  Call init.unsetting(True)
  '----------------------------------------------
  
End Function
