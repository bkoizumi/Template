Attribute VB_Name = "Ctl_Access"
Option Explicit

Dim dbCon       As ADODB.Connection
Dim DBRecordset As ADODB.Recordset
Dim queryString As String


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
  Call Library.showDebugForm("ConnectServer�F" & ConnectServer)
  
  Set dbCon = New ADODB.Connection
  dbCon.Open ConnectServer
  
  isDBOpen = True
  Call Library.showDebugForm("isDBOpen�F" & isDBOpen)
  
  Exit Function
  
'�G���[������--------------------------------------------------------------------------------------
catchError:
  isDBOpen = False
  Call Library.showNotice(500, Err.Description, True)
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
  Call Library.showNotice(501, Err.Description, True)
End Function


'==================================================================================================
'�e�[�u�����擾
Function getTableInfo()
  Dim line As Long, endLine As Long
  Dim cat As ADOX.Catalog
  Dim tbl As ADOX.Table
  Dim tableCnt As Long

  '�����J�n--------------------------------------
  'On Error GoTo catchError
  ----
  Const funcName As String = "Ctl_Access.getTableInfo"
  PrgP_Max = 4
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
  End If
  Call Library.showDebugForm(funcName & "==========================================")
  'Call Ctl_Access.dbOpen
  '----------------------------------------------
  Set cat = New ADOX.Catalog
  cat.ActiveConnection = ConnectServer
  For Each tbl In cat.Tables
    Select Case tbl.Type
      Case "TABLE"
        Call Ctl_Common.addSheet(tbl.Name)
        Range(setVal("Cell_TableType")) = "�}�X�^�[�e�[�u��"
        
      Case "VIEW"
        Call Ctl_Common.addSheet(tbl.Name)
        Range(setVal("Cell_TableType")) = "�N�G���r���["
      
      Case "LINK", "PASS-THROUGH"
        Call Ctl_Common.addSheet(tbl.Name)
        Range(setVal("Cell_TableType")) = "�����N�e�[�u��"
        
      '�V�X�e���e�[�u��
      Case "ACCESS TABLE", "SYSTEM TABLE"
        GoTo Lbl_nextfor
    End Select
    
    Range(setVal("Cell_logicalTableName")) = ""
    Range(setVal("Cell_physicalTableName")) = tbl.Name
    Call Ctl_Access.getColumnInfo

Lbl_nextfor:
  tableCnt = tableCnt + 1
  Call Ctl_ProgressBar.showBar(thisAppName, 1, PrgP_Max, tableCnt, cat.Tables.count, "�e�[�u�����擾�F" & tbl.Name)
  Next tbl
  Call Ctl_Common.makeTblList
  
  '�����I��--------------------------------------
'  Call Ctl_Access.dbClose
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.unsetting(True)
  End If
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
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
  
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  
  
  Const funcName As String = "Ctl_Access.getColumnInfo"
  If PrgP_Max = 0 Then
    PrgP_Max = 2
  End If
  '--------------------------
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
  End If
  
  Call Library.showDebugForm(funcName & "=========================================")
  Call Ctl_Access.dbOpen
  '----------------------------------------------
  Set targetSheet = ActiveSheet
  Select Case targetSheet.Name
    Case "�ݒ�-MySQL", "�ݒ�-ACC", "Notice", "DataType", "�R�s�[�p", "�\��", "TBL���X�g", "�ύX����", "ER�}"
    Exit Function
  End Select
  Call Ctl_Common.ClearData
  
  tableName = targetSheet.Range(setVal("Cell_physicalTableName"))
  '�J�������--------------------------------------------------------------------------------------
  queryString = "SELECT * FROM " & tableName
  
  Set ClmRecordset = dbCon.Execute(queryString)
  
  line = startLine
  columnCnt = 1
  For Each Fields In ClmRecordset.Fields
    targetSheet.Range("B" & line) = ""
    targetSheet.Range(setVal("Cell_physicalName") & line) = Fields.Name
    
    If ArryTypeName(Fields.Type) Like "ad*" Then
      targetSheet.Range(setVal("Cell_dateType") & line) = Fields.Type & "," & ArryTypeName(Fields.Type)
    Else
      targetSheet.Range(setVal("Cell_dateType") & line) = ArryTypeName(Fields.Type)
    End If
    
    Select Case Range(setVal("Cell_dateType") & line)
      Case "MEMO", "DATE", "CURRENCY", "INT", "YESNO", "LONGBINARY"
      Case Else
        targetSheet.Range(setVal("Cell_digits") & line) = Fields.DefinedSize
    End Select
    

    
    
    Call Ctl_ProgressBar.showBar(thisAppName, 1, PrgP_Max, columnCnt, ClmRecordset.Fields.count, "�J�������擾�F" & Fields.Name)

    line = line + 1
    columnCnt = columnCnt + 1
    Call Ctl_Common.addRow(line)

  Next
  '�����I��--------------------------------------
  Call Ctl_Access.dbClose
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.unsetting(True)
  End If
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function




'==================================================================================================
'DDL�쐬
Function makeDDL()
  Dim line As Long, endLine As Long
  Dim ColumnString As String
  
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_Access.makeDDL"
  PrgP_Max = 4
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
  End If
  Call Library.showDebugForm(funcName & "==========================================")
  'Call Ctl_Access.dbOpen
  '----------------------------------------------
  endLine = Cells(Rows.count, 5).End(xlUp).Row
  
  queryString = "CREATE TABLE " & Range(setVal("Cell_physicalTableName")) & "("
  For line = startLine To endLine
    If Range(setVal("Cell_logicalName") & line) <> "" Then
      If ColumnString = "" Then
        ColumnString = Range(setVal("Cell_logicalName") & line) & " " & Range(setVal("Cell_physicalName") & line)
      Else
        ColumnString = ColumnString & ",  " & Range(setVal("Cell_logicalName") & line) & " " & Range(setVal("Cell_physicalName") & line)
      End If
      
      If Range(setVal("Cell_dateType") & line) <> "" Then
        ColumnString = ColumnString & " (" & Range(setVal("Cell_dateType") & line) & ")" & vbNewLine
      Else
        ColumnString = ColumnString & vbNewLine
      End If
      
      
    Else
      Exit For
    End If
  Next
  queryString = queryString & vbNewLine & ColumnString & ")"
  
  Debug.Print queryString
  
'  Call Ctl_ProgressBar.showBar(thisAppName, 1, PrgP_Max, tableCnt, cat.Tables.count, "�e�[�u�����擾�F" & tbl.Name)
  
  
  '�����I��--------------------------------------
'  Call Ctl_Access.dbClose
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.unsetting(True)
  End If
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function


'==================================================================================================
'���ڎ��s
Function CreateTable()
  Dim line As Long, endLine As Long
  Dim tableName As String
  Dim ColumnString As String
  Dim oldColumnName As String
  
  '�����J�n--------------------------------------
  'On Error GoTo catchError

  Const funcName As String = "Ctl_Access.CreateTable"
  PrgP_Max = 4
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
  End If
  Call Library.showDebugForm(funcName & "===========================================")
  Call Ctl_Access.dbOpen
  '----------------------------------------------
  endLine = Cells(Rows.count, 3).End(xlUp).Row
  
  tableName = Range(setVal("Cell_physicalTableName"))
  
  If Ctl_Access.IsTable(tableName) = True Then
    '�e�[�u�������݂���ꍇ----------------------
    For line = startLine To endLine
      If Range("B" & line) = "edit" Then
        '�f�[�^�^�ύX------------------------------
        queryString = "ALTER TABLE [" & Range(setVal("Cell_physicalTableName")) & "] ALTER COLUMN [" & Range(setVal("Cell_physicalName") & line) & "] " & Range(setVal("Cell_dateType") & line)
        If Range(setVal("Cell_digits") & line) <> "" Then
          queryString = queryString & " (" & Range(setVal("Cell_digits") & line) & ");"
        Else
          queryString = queryString & ";" & vbNewLine
        End If
        Call Library.showDebugForm("QueryString", queryString)
        Call Ctl_Access.runQuery(queryString)
        Range("B" & line) = ""
        
      '�J�������ύX[�ǉ��ˍ폜]------------------
      ElseIf Range("B" & line) Like "rename:*" Then
'        oldColumnName = Replace(Range("B" & line), "rename:", "")
'
'        queryString = "ALTER TABLE [" & Range(setVal("Cell_physicalTableName")) & "] ADD COLUMN [" & Range(setVal("Cell_physicalName") & line) & "] " & Range(setVal("Cell_dateType") & line)
'        If Range(setVal("Cell_digits") & line) <> "" Then
'          queryString = queryString & " (" & Range(setVal("Cell_digits") & line) & ");"
'        Else
'          queryString = queryString & ";" & vbNewLine
'        End If
'        Call Library.showDebugForm("QueryString", queryString)
'        Call Ctl_Access.runQuery(queryString)
'
'        queryString = "ALTER TABLE [" & Range(setVal("Cell_physicalTableName")) & "] DROP COLUMN [" & oldColumnName & "];"
'        Call Library.showDebugForm("QueryString", queryString)
'        Call Ctl_Access.runQuery(queryString)
      
      '�J�����폜--------------------------------
      ElseIf Range("B" & line) = "delete" Then
        queryString = "ALTER TABLE [" & Range(setVal("Cell_physicalTableName")) & "] DROP COLUMN [" & Range(setVal("Cell_physicalName") & line) & "];"
        Call Library.showDebugForm("QueryString", queryString)
        Call Ctl_Access.runQuery(queryString)
        Rows(line & ":" & line).Delete Shift:=xlUp
        line = line - 1
        
      End If
      
      If Range("B" & line) <> "" Then
        Call Ctl_ProgressBar.showBar(thisAppName, 1, PrgP_Max, line, endLine, "�J�������ύX�F" & Range(setVal("Cell_physicalName") & line))
      End If
    Next
    
    
  Else
    queryString = "CREATE TABLE " & Range(setVal("Cell_physicalTableName")) & "("
    For line = startLine To endLine
      If Range(setVal("Cell_logicalName") & line) <> "" Then
        If ColumnString = "" Then
          ColumnString = "[" & Range(setVal("Cell_physicalName") & line) & "] " & Range(setVal("Cell_dateType") & line)
        Else
          ColumnString = ColumnString & ",  [" & Range(setVal("Cell_physicalName") & line) & "] " & Range(setVal("Cell_dateType") & line)
        End If
        
        If Range(setVal("Cell_digits") & line) <> "" Then
          ColumnString = ColumnString & " (" & Range(setVal("Cell_digits") & line) & ")" & vbNewLine
        Else
          ColumnString = ColumnString & vbNewLine
        End If
      Else
        Exit For
      End If
    Next
    queryString = queryString & vbNewLine & ColumnString & ")"
    Call Library.showDebugForm("QueryString", queryString)
    Call Ctl_Access.runQuery(queryString)
    Range("B5") = "exist"
  End If
  
  

  
  
  '�����I��--------------------------------------
  Call Ctl_Access.dbClose
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.unsetting(True)
  End If
  '----------------------------------------------

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function


'==================================================================================================
Function IsTable(tableName As String) As Boolean
  Dim cat As ADOX.Catalog
  Dim tbl As ADOX.Table
  Dim rslFlg As Boolean

  rslFlg = False
  Set cat = New ADOX.Catalog
  cat.ActiveConnection = ConnectServer
  For Each tbl In cat.Tables
    If tbl.Name = tableName Then
      rslFlg = True
      Exit For
    End If
  Next
  IsTable = rslFlg
End Function

'==================================================================================================
Function runQuery(queryString As String)
  Dim oCn As ADODB.Connection
  Dim oRs As ADODB.Recordset

  On Error GoTo catchError
  
  Set oCn = CreateObject("ADODB.Connection")
  Set oRs = CreateObject("ADODB.Recordset")
  
  oCn.Open ConnectServer
  oRs.Open queryString, oCn
  
  oCn.Close
  Set oRs = Nothing
  Set oCn = Nothing
  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  oCn.Close
  Set oRs = Nothing
  Set oCn = Nothing
  
  If Err.Number = -2147217900 Then
    Call Library.showNotice(502, funcName & " �\���G���[" & vbNewLine & queryString, True)
  Else
    Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
  End If
End Function
