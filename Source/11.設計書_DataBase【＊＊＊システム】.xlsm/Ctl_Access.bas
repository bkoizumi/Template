Attribute VB_Name = "Ctl_Access"
Option Explicit

Dim dbCon       As ADODB.Connection
Dim DBRecordset As ADODB.Recordset
Dim QueryString As String


'**************************************************************************************************
' * MS Access
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function dbOpen()
  On Error GoTo catchError
  
  If isDBOpen = True Then
    Call Library.showDebugForm("Database is already opened")
    Exit Function
  End If
  
  Set dbCon = New ADODB.Connection
  dbCon.Open ConnectServer
  
  isDBOpen = True
  Call Library.showDebugForm("isDBOpen�F" & isDBOpen)
  
  Exit Function
  
'�G���[������--------------------------------------------------------------------------------------
catchError:
  isDBOpen = False
  Call Library.showNotice(400, Err.Description, True)
End Function

'==================================================================================================
Function dbClose()
  On Error GoTo catchError
  
  If dbCon Is Nothing Then
    Call Library.showDebugForm("Database is already closed")
  Else
    dbCon.Close
    isDBOpen = False
    
    Set dbCon = Nothing
    Call Library.showDebugForm("isDBOpen�F" & isDBOpen)
  End If
  
  Exit Function

'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm("isDBOpen�F" & isDBOpen)
  Call Library.showNotice(400, Err.Number & "�F" & Err.Description, True)
End Function


'==================================================================================================
'�e�[�u�����擾
Function getTableInfo()
  Dim line As Long, endLine As Long
  Dim cat As ADOX.Catalog
  Dim tbl As ADOX.Table
  Dim qry1 As ADOX.View
  Dim qry2 As ADOX.Procedure
  Dim constr As String
  Dim DBFile As String
  

  '�����J�n----------------------------------------------------------------------------------------
  'On Error GoTo catchError
  '�����l�ݒ�----
  runFlg = True
  FuncName = "Ctl_Access.getTableInfo"

  Call Library.startScript
  Call init.Setting
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm(FuncName & "==========================================")
  'Call Ctl_Access.dbOpen
  
  '----------------------------------------------
  Set cat = New ADOX.Catalog
  cat.ActiveConnection = ConnectServer
  For Each tbl In cat.Tables
    Select Case tbl.Type
      Case "TABLE"
        Call Ctl_Common.addSheet(tbl.Name)
        Range("B2") = "�}�X�^�[�e�[�u��"
        
      Case "VIEW"
        Call Ctl_Common.addSheet(tbl.Name)
        Range("B2") = "�N�G���r���["

      
      '�����N�e�[�u���iODBC�ȊO�j
      Case "LINK"
        Call Ctl_Common.addSheet(tbl.Name)
        Range("B2") = "�����N�e�[�u��"
        
      'Access �V�X�e���e�[�u���AMicrosoft jet �V�X�e���e�[�u��
      Case "ACCESS TABLE", "SYSTEM TABLE"
        GoTo Lbl_nextfor
        
      '�����N�e�[�u���iODBC�j
      Case "PASS-THROUGH"
        Call Library.showDebugForm("tbl.Name�F" & tbl.Name)
        GoTo Lbl_nextfor
    End Select
    
    Range("D5") = ""
    Range("H5") = tbl.Name
    Call Ctl_Access.getColumnInfo

Lbl_nextfor:
  Next tbl
  Call Ctl_Common.makeTblList
  
  '�����I��----------------------------------------------------------------------------------------
'  Call Ctl_Access.dbClose
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  If runFlg = True Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.usetting
  End If
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "�F" & Err.Description, True)
End Function


'==================================================================================================
'�J�������擾
Function getColumnInfo()
  Dim line As Long, endLine As Long
  Dim tableName As String
  Dim columnCnt As Long
  Dim ClmRecordset As ADODB.Recordset

  Dim ColumnNames() As Variant
  Dim indexCount As Integer
  
  Dim Fields As ADODB.Field
  
  '�����J�n----------------------------------------------------------------------------------------
  'On Error GoTo catchError
  
  '�����l�ݒ�----------------
  FuncName = "Ctl_Access.getColumnInfo"
  PrgP_Max = 2
  '--------------------------
  Call Library.startScript
  Call init.Setting
  Call Ctl_ProgressBar.showStart
  Call Library.showDebugForm(FuncName & "==========================================")
  Call Ctl_Access.dbOpen
  '----------------------------------------------
  
  Call Ctl_Common.ClearData
  
  tableName = Range("H5")
  '�J�������--------------------------------------------------------------------------------------
  QueryString = "SELECT * FROM " & tableName
  
  Set ClmRecordset = dbCon.Execute(QueryString)
  
  line = startLine
  columnCnt = 1
  For Each Fields In ClmRecordset.Fields
    Range("E" & line) = Fields.Name
    Range("G" & line) = ArryTypeName(Fields.Type)
    Range("H" & line) = Fields.DefinedSize
    
    Call Ctl_ProgressBar.showBar(thisAppName, 1, PrgP_Max, columnCnt, ClmRecordset.Fields.count, "�J�������擾�F" & Fields.Name)

    line = line + 1
    columnCnt = columnCnt + 1
    Call Ctl_Common.addRow(line)

  Next
  '�����I��----------------------------------------------------------------------------------------
  Call Ctl_Access.dbClose
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.usetting
  End If
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "�F" & Err.Description, True)
End Function
