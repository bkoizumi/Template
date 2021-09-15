Attribute VB_Name = "init"
'���[�N�u�b�N�p�ϐ�------------------------------
Public ThisBook   As Workbook
Public targetBook As Workbook


'���[�N�V�[�g�p�ϐ�------------------------------
Public targetSheet    As Worksheet

Public sheetSetting   As Worksheet
Public sheetNotice    As Worksheet
Public sheetDataType  As Worksheet
Public sheetCopy      As Worksheet
Public sheetTblList   As Worksheet



'�O���[�o���ϐ�----------------------------------
Public Const thisAppName = "DB�݌v��"
Public Const thisAppVersion = "V1.0-beta.1"

Public ConnectServer      As String
Public Const startLine    As Long = 9
Public isDBOpen           As Boolean
Public runFlg             As Boolean

Public PrgP_Max           As Long
Public PrgP_Cnt           As Long

Public FuncName As String

Public accFileName        As String
Public accFileDir         As String
Public ArryTypeName(205)  As String
Public oldCellVal         As String

'���W�X�g���o�^�p�T�u�L�[
'Public Const RegistryKey  As String = "BK_Documents"


'�ݒ�l�ێ�
Public setVal         As Object


'�t�@�C���֘A
Public logFile      As String

'�������Ԍv���p
Public StartTime          As Date
Public StopTime           As Date



'���{���֘A--------------------------------------
Public ribbonUI       As Office.IRibbonUI
Public ribbonVal      As Object


'**************************************************************************************************
' * �ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function usetting(Optional flg As Boolean = True)

  Set ThisBook = Nothing
  
  Set sheetSetting = Nothing
  Set sheetNotice = Nothing
  Set sheetDataType = Nothing
  Set sheetCopy = Nothing
  Set sheetTblList = Nothing
  
  Set setVal = Nothing
  
  PrgP_Max = 0
  PrgP_Cnt = 0
  logFile = ""
  
  If flg = True Then
    runFlg = False
  End If
End Function


'==================================================================================================
Function Setting(Optional reCheckFlg As Boolean)
  Dim line As Long, endLine As Long
'  On Error GoTo catchError
'  ThisWorkbook.Save

  If logFile = "" Or reCheckFlg = True Then
    Call usetting(False)
  Else
    Exit Function
  End If

  '�u�b�N�̐ݒ�
  Set ThisBook = ThisWorkbook
  
  '���[�N�V�[�g���̐ݒ�
  'Set sheetSetting = ThisBook.Worksheets("�ݒ�-MySQL")
  Set sheetSetting = ThisBook.Worksheets("�ݒ�-ACC")
  
  
  Set sheetNotice = ThisBook.Worksheets("Notice")
  Set sheetDataType = ThisBook.Worksheets("DataType")
  
  Set sheetCopy = ThisBook.Worksheets("�R�s�[�p")
  Set sheetTblList = ThisBook.Worksheets("TBL���X�g")
  
  logFile = ThisWorkbook.Path & "\ExcelMacro.log"
        
  '�ݒ�l�ǂݍ���----------------------------------------------------------------------------------
  Set setVal = Nothing
  Set setVal = CreateObject("Scripting.Dictionary")
  
  For line = 5 To sheetSetting.Cells(Rows.count, 1).End(xlUp).Row
    If sheetSetting.Range("A" & line) <> "" Then
      setVal.Add sheetSetting.Range("A" & line).Text, sheetSetting.Range("B" & line).Text
    End If
  Next

  Select Case setVal("DBMS")
    Case "MSAccess"
      accFileName = Library.getFileInfo(setVal("DBServer"), , "fileName")
      accFileDir = Library.getFileInfo(setVal("DBServer"), , "CurrentDir")
      ConnectServer = "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                      "Data Source=" & setVal("DBServer") & ";" & _
                      "Jet OLEDB:Database Password=" & setVal("passwd") & ";"
                     
      Range("DBName") = accFileName
      
      endLine = sheetSetting.Cells(Rows.count, 12).End(xlUp).Row
      For line = 5 To endLine
        ArryTypeName(sheetSetting.Range("L" & line)) = sheetSetting.Range("M" & line)
      Next
    Case "MySQL"
      ConnectServer = "Driver={" & setVal("ODBCDriver") & "};" & _
                      " Server=" & setVal("DBServer") & ";" & _
                      " Port=" & setVal("Port") & ";" & _
                      " Database=" & setVal("DBName") & ";" & _
                      " User=" & setVal("userID") & ";" & _
                      " Password=" & setVal("passwd") & ";" & _
                      " Charset=sjis;"
    
    Case "PostgreSQL"
      ConnectServer = ""
      
    Case "SQLServer"
      ConnectServer = "Provider=SQLOLEDB;" & _
                      "Data Source=" & setVal("DBServer") & ";" & _
                      "Initial Catalog=" & setVal("DBName") & ";" & _
                      "Trusted_Connection=Yes"
  
  End Select
  
  
  
  Call ���O��`
  Exit Function
  
'�G���[������--------------------------------------------------------------------------------------
catchError:
  
End Function



'**************************************************************************************************
' * ���O��`
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ���O��`()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim Name As Object
  
'  On Error GoTo catchError
   
  For Each Name In Names
    If Name.Visible = False Then
      Name.Visible = True
    End If
    If Not Name.Name Like "[Print_]*" Then
      If Name.Name Like "_xlfn*" Then
'        MsgBox "�}�N���ł͍폜�ł��Ȃ����O�̒�`������܂�" & vbNewLine & Name.Name, vbExclamation
      Else
        Name.Delete
      End If
    End If
  Next
  
  'VBA�p�̐ݒ�
  For line = 3 To sheetSetting.Cells(Rows.count, 1).End(xlUp).Row
    If sheetSetting.Range("A" & line) <> "" Then
      sheetSetting.Range("B" & line).Name = sheetSetting.Range("A" & line)
    End If
  Next
  
  'Book�p�̐ݒ�
  For colLine = 7 To 10
    endLine = sheetSetting.Cells(Rows.count, colLine).End(xlUp).Row
    sheetSetting.Range(sheetSetting.Cells(5, colLine), sheetSetting.Cells(endLine, colLine)).Name = sheetSetting.Cells(4, colLine)
  Next
  
  'DataType�p�̐ݒ�
  For colLine = 1 To 15 Step 3
    endLine = sheetDataType.Cells(Rows.count, colLine).End(xlUp).Row
    sheetDataType.Range(sheetDataType.Cells(3, colLine), sheetDataType.Cells(endLine, colLine)).Name = sheetDataType.Cells(1, colLine)
  Next
  
  
  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
  
End Function


'**************************************************************************************************
' * �V�[�g�̕\��/��\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function �V�[�g��\��()

  If setVal("debugMode") <> "develop" Then
    Worksheets("�ݒ�").Visible = xlSheetVeryHidden
    Worksheets("Notice").Visible = xlSheetVeryHidden
    Worksheets("DataType").Visible = xlSheetVeryHidden
  End If
  
  Worksheets("TBL���X�g").Select
End Function


'==================================================================================================
Function �V�[�g�\��()
  
  Worksheets("�ݒ�").Visible = True
  Worksheets("Notice").Visible = True
  Worksheets("DataType").Visible = True
  
  Worksheets("TBL���X�g").Select
  
End Function


'==================================================================================================
Function �V�[�g�ی�()
  Dim sheetName As String
  Dim tempSheet As Object

  Call init.Setting
  Call Library.showDebugForm("sheetProtect--------------------------")
  For Each tempSheet In Sheets
    sheetName = tempSheet.Name
    If Not (sheetName Like "[�ݒ�,Notice,DataType]*") Then
      Call Library.showDebugForm("  " & sheetName)
      
      DoEvents
      ThisWorkbook.Worksheets(sheetName).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, passWord:=thisAppPasswd
      ThisWorkbook.Worksheets(sheetName).EnableSelection = xlNoRestrictions
    End If
  Next
  Call Library.showDebugForm("--------------------------------------")
End Function

'==================================================================================================
Function �V�[�g�ی����()
  Dim sheetName As String
  Dim tempSheet As Object

  Call init.Setting
  Call Library.showDebugForm("sheetUnprotect--------------------------")
  For Each tempSheet In Sheets
    sheetName = tempSheet.Name
    If Not (sheetName Like "[�ݒ�,Notice,DataType]*") Then
      Call Library.showDebugForm("  " & sheetName)
      
      DoEvents
      ThisWorkbook.Worksheets(sheetName).Unprotect passWord:=thisAppPasswd
    End If
  Next
  Call Library.showDebugForm("----------------------------------------")
End Function





