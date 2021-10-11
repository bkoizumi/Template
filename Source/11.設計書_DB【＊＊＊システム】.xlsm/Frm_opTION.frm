VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Option 
   Caption         =   "UserForm1"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10065
   OleObjectBlob   =   "Frm_opTION.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_opTION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit









'**************************************************************************************************
' * 初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  Dim ListIndex As Integer
  Dim line As Long, endLine As Long, i As Long
  Dim objShp As Shape
  
  Call init.Setting
  
  'DBList取得
  endLine = sheetSetting.Cells(Rows.count, 8).End(xlUp).Row
  For line = 5 To endLine
    DBMS.AddItem sheetSetting.Range("H" & line).Text
  Next
  DBMS.ListIndex = 0
  
  
  CustomerName.Value = setVal("CustomerName")
  ProjectName.Value = setVal("ProjectName")
  systemName.Value = setVal("systemName")
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
' * ボタン押下時処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Private Sub outputDDL_Click()
  Dim targetDir As String
  
  targetDir = Library.getDirPath(outputDir.Value)
  If targetDir <> "" Then
    outputDir.Value = targetDir
  End If
End Sub

'==================================================================================================
Private Sub ConnectTEST_Click()
  Dim line As Long, endLine As Long

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
      
      Call Ctl_MySQL.dbOpen(False)
      If isDBOpen = False Then
        conMessage = "接続に失敗しました"
      Else
        conMessage = "接続成功!!"
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
  
  End Select
  
  
  


End Sub


'==================================================================================================
'キャンセル処理
Private Sub Cancel_Click()
  Unload Me
  
End Sub


'==================================================================================================
' 実行
Private Sub Submit_Click()
  Dim i As Integer
  
  Call init.Setting
  
'  Call Library.setValandRange("useLogicalName", CStr(useLogicalName.Value))
'  Call Library.setValandRange("usePhysicalName", CStr(usePhysicalName.Value))
'  Call Library.setValandRange("useImage", CStr(useImage.Value))
'
  Unload Me
End Sub

  



