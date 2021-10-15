VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_setOption 
   Caption         =   "��{�ݒ�"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10065
   OleObjectBlob   =   "Frm_setOption.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Frm_setOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit














'**************************************************************************************************
' * �����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  Dim ListIndex As Integer
  Dim line As Long, endLine As Long, i As Long
  Dim objShp As Shape
  
  Call init.Setting
  
  'DBList�擾
  endLine = sheetSetting.Cells(Rows.count, 8).End(xlUp).Row
  For line = 5 To endLine
    DBMS.AddItem sheetSetting.Range("H" & line).text
  Next
  DBMS.ListIndex = 0
  
  
  CustomerName.Value = setVal("CustomerName")
  ProjectName.Value = setVal("ProjectName")
  systemName.Value = setVal("systemName")
  subSystemName.Value = setVal("subSystemName")
  CreateBy.Value = setVal("CreateBy")
  CreateAt.Value = setVal("CreateAt")
  outputDir.Value = setVal("outputDir")
  
  DBMS.Value = setVal("DBMS")
  DBServer.Value = setVal("DBServer")
  DBName.Value = setVal("DBName")
  Port.Value = setVal("Port")
  Instance.Value = setVal("Instance")
  Schema.Value = setVal("Schema")
  UserId.Value = setVal("userID")
  passwd.Value = setVal("passwd")
  
  
  conMessage = ""
End Sub


'**************************************************************************************************
' * �C�x���g����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Private Sub DBName_Change()
  
  Select Case DBMS.Value
    Case "MySQL"
      Schema.Value = DBName.Value
    Case Else
  End Select

End Sub


'**************************************************************************************************
' * �{�^������������
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Private Sub Btn_outputDir_Click()
  Dim targetDir As String
  
  targetDir = Library.getDirPath(outputDir.Value)
  If targetDir <> "" Then
    outputDir.Value = targetDir
  End If
End Sub

'==================================================================================================
'�ڑ��e�X�g
Private Sub Btn_ConnectTEST_Click()
  Dim line As Long, endLine As Long
  Dim ErrMessage As String
  
  Select Case DBMS.Value
    '----------------------------------------------------------------------------------------------
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
    
    '----------------------------------------------------------------------------------------------
    Case "MySQL"
      ConnectServer = "Driver={MySQL ODBC 8.0 Unicode Driver};" & _
                      " Server=" & DBServer.Value & ";" & _
                      " Port=" & Port.Value & ";" & _
                      " Database=" & DBName.Value & ";" & _
                      " User=" & UserId.Value & ";" & _
                      " Password=" & passwd.Value & ";" & _
                      " Charset=sjis;"
      
      Call Ctl_MySQL.dbOpen(False, ErrMessage)
      If isDBOpen = False Then
        conMessage = "�ڑ��Ɏ��s���܂���"
      Else
        conMessage = "�ڑ�����!!"
        Call Ctl_MySQL.dbClose
      End If
      
    '----------------------------------------------------------------------------------------------
    Case "PostgreSQL"
      ConnectServer = ""
    
    '----------------------------------------------------------------------------------------------
    Case "SQLServer"
      ConnectServer = "Provider=SQLOLEDB;" & _
                      "Data Source=" & DBServer.Value & ";" & _
                      "Initial Catalog=" & DBName.Value & ";" & _
                      "Trusted_Connection=Yes"
  
      Call Ctl_SQLServer.dbOpen(False, ErrMessage)
      If isDBOpen = False Then
        conMessage = "�ڑ��Ɏ��s���܂���" & vbNewLine & ErrMessage
      Else
        conMessage = "�ڑ�����!!"
        Call Ctl_SQLServer.dbClose
      End If
    End Select
End Sub


'==================================================================================================
'�L�����Z������
Private Sub Cancel_Click()
  Unload Me
  
End Sub


'==================================================================================================
' ���s
Private Sub Submit_Click()
  Dim i As Integer
  
  Call init.Setting
  
  Select Case DBMS.Value
    Case "MySQL"
      Schema.Value = DBName.Value
    Case Else
  End Select
  
  Call Library.setValandRange("CustomerName", CustomerName.Value)
  Call Library.setValandRange("systemName", systemName.Value)
  Call Library.setValandRange("subSystemName", subSystemName.Value)
  Call Library.setValandRange("CreateBy", CreateBy.Value)
  Call Library.setValandRange("CreateAt", CreateAt.Value)
  Call Library.setValandRange("outputDir", outputDir.Value)

  Call Library.setValandRange("DBMS", DBMS.Value)
  Call Library.setValandRange("DBServer", DBServer.Value)
  Call Library.setValandRange("DBName", DBName.Value)
  Call Library.setValandRange("Port", Port.Value)
  Call Library.setValandRange("Instance", Instance.Value)
  Call Library.setValandRange("Schema", Schema.Value)
  Call Library.setValandRange("UserId", UserId.Value)
  Call Library.setValandRange("passwd", passwd.Value)


  Unload Me
End Sub

  



