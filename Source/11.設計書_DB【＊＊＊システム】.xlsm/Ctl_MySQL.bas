Attribute VB_Name = "Ctl_MySQL"
Option Explicit

Dim dbCon       As ADODB.Connection
Dim DBRecordset As ADODB.Recordset
Dim queryString As String

'**************************************************************************************************
' * MySQL
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function dbOpen(Optional NoticeFlg As Boolean = True)
  Const funcName As String = "Ctl_MySQL.dbOpen"
  
  On Error GoTo catchError
  
  
  
  If isDBOpen = True Then
    Call Library.showDebugForm("Database is already opened")
    Exit Function
  End If
  If setVal("debugMode") = "develop" Then
    Call Library.showDebugForm("ConnectServer", ConnectServer)
  End If
  
  Set dbCon = New ADODB.Connection
  dbCon.Open ConnectServer
  dbCon.CursorLocation = 3
  
  isDBOpen = True
  Call Library.showDebugForm("isDBOpen", isDBOpen)
  
  Exit Function
  
'�G���[������--------------------------------------------------------------------------------------
catchError:
  isDBOpen = False
  Call Library.showDebugForm("isDBOpen", isDBOpen)
  If NoticeFlg = True Then
    
    Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
  End If
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
    Call Library.showDebugForm("isDBOpen", isDBOpen)
  End If
  
  Exit Function

'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm("isDBOpen", isDBOpen)
  Call Library.showNotice(501, Err.Description, True)
End Function


'==================================================================================================
Function IsTable(tableName As String) As Boolean
  Dim TblRecordset As ADODB.Recordset
  Dim rslFlg As Boolean

  rslFlg = False

  Call Library.showDebugForm("TableName", tableName)
  Call Ctl_MySQL.dbOpen
  
  '�e�[�u�����--------------------------------------------------------------------------------------
  queryString = " SELECT" & _
                "   TABLE_NAME as TableName, TABLE_COMMENT as Comments" & _
                " from" & _
                "   information_schema.TABLES" & _
                " WHERE" & _
                "   TABLE_SCHEMA = DATABASE()" & _
                "   and TABLE_NAME='" & tableName & "'"
      
  Call Library.showDebugForm("QueryString", queryString)
  
  Set TblRecordset = New ADODB.Recordset
  TblRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly
  If TblRecordset.RecordCount = 1 Then
    rslFlg = True
  End If
  
  Set TblRecordset = Nothing
  Set targetSheet = Nothing
  Call Ctl_MySQL.dbClose
  
  IsTable = rslFlg
End Function


'==================================================================================================
Function runQuery(queryString As String)
  Dim runRecordset As ADODB.Recordset
  
  Const funcName As String = "Ctl_MySQL.runQuery"

  On Error GoTo catchError
  
  Set runRecordset = dbCon.Execute(queryString)
  Set runRecordset = Nothing
  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Dim errId, errMeg, meg
  errId = Err.Number
  errMeg = Err.Description
  
  Call Library.showDebugForm(funcName, errId & "�F" & errMeg)
  If errId = -2147217900 Then
    meg = " �\���G���[" & vbNewLine
    meg = meg & "-------------------------------------------------" & vbNewLine
    meg = meg & errMeg & vbNewLine
    meg = meg & "-------------------------------------------------" & vbNewLine & vbNewLine
    meg = meg & queryString & vbNewLine

    Call Library.showNotice(502, funcName & meg, True)
  
  Else
    Call Library.showNotice(400, funcName & " [" & errId & "]" & errMeg, True)
  End If
End Function


'==================================================================================================
'DB���擾
Function getDatabaseInfo(Optional ErImgflg As Boolean = False)
  Dim line As Long, endLine As Long
  Dim TblRecordset As ADODB.Recordset
  Dim tableName   As String
  Dim lValues(2) As Variant

  Dim physicalTableName As String, logicalTableName As String
  Dim tableNote As String, TableCretateAt As String

  '�����J�n--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_MySQL.getDatabaseInfo"
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
    Call Library.showDebugForm("runFlg", runFlg)
  End If
  Call Library.showDebugForm(funcName & "==========================================")
  '----------------------------------------------
  'ER�}�����p�ݒ�
  If ErImgflg = True Then
    line = 1
    sheetTmp.Range("A" & line) = "#"
    sheetTmp.Range("B" & line) = "�����e�[�u����"
    sheetTmp.Range("C" & line) = "�_���e�[�u����"
    sheetTmp.Range("D" & line) = "�쐬��"
    line = line + 1
  End If
  
  '�e�[�u�����--------------------------------------------------------------------------------------
  queryString = " SELECT" & _
                "   TABLE_NAME as TableName, TABLE_COMMENT as Comments, CREATE_TIME AS CREATETIME" & _
                " FROM" & _
                "   information_schema.TABLES" & _
                " WHERE" & _
                "   TABLE_SCHEMA = DATABASE();"
      
  Call Library.showDebugForm("QueryString", queryString)
  
  Set TblRecordset = New ADODB.Recordset
  TblRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly
  PrgP_Max = TblRecordset.RecordCount
  
  Do Until TblRecordset.EOF
    physicalTableName = TblRecordset.Fields("TableName").Value
    Call Library.showDebugForm("physicalTableName", physicalTableName)
    
    If TblRecordset.Fields("Comments").Value Like "*" & vbTab & "*" Then
      logicalTableName = Split(TblRecordset.Fields("Comments").Value, vbTab)(0)
      tableNote = Replace(Split(TblRecordset.Fields("Comments").Value, vbTab)(1), "\n", vbNewLine)
    Else
      tableNote = TblRecordset.Fields("Comments")
    End If
    TableCretateAt = TblRecordset.Fields("CREATETIME")
    
    PrgP_Cnt = TblRecordset.AbsolutePosition
    Call Ctl_ProgressBar.showBar(thisAppName, TblRecordset.AbsolutePosition, TblRecordset.RecordCount, 1, 2, "�J�������擾")
    
  If ErImgflg = False Then
    '�V�[�g�ǉ�
    If Library.chkSheetExists(physicalTableName) = True Then
      Set targetSheet = ThisWorkbook.Worksheets(physicalTableName)
    Else
      Call Ctl_Common.addSheet(physicalTableName)
      Set targetSheet = ActiveSheet
    End If
    
    targetSheet.Select
    targetSheet.Range("B5") = "exist"
    If targetSheet.Range(setVal("Cell_TableType")) = "" Then
      targetSheet.Range(setVal("Cell_TableType")) = "�}�X�^�[�e�[�u��"
    End If
    targetSheet.Range(setVal("Cell_physicalTableName")) = physicalTableName
    
    targetSheet.Range(setVal("Cell_logicalTableName")) = logicalTableName
    targetSheet.Range(setVal("Cell_physicalTableName")) = physicalTableName
    targetSheet.Range(setVal("Cell_tableNote")) = tableNote

    '�J�������擾
    Call Ctl_MySQL.getColumnInfo(physicalTableName)
    
    Else
      'ER�}�������̏���--------------------------
      sheetTmp.Range("A" & line) = TblRecordset.AbsolutePosition
      sheetTmp.Range("B" & line) = physicalTableName
      sheetTmp.Range("C" & line) = logicalTableName
      sheetTmp.Range("D" & line) = TableCretateAt
      line = line + 1
  End If
  
  TblRecordset.MoveNext
  Loop
  Set TblRecordset = Nothing
  Set targetSheet = Nothing

  '�e�[�u�����X�g����
  If ErImgflg = False Then
    Call Ctl_Common.makeTblList
  End If
  
  '�����I��--------------------------------------
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  '----------------------------------------------
  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function


'==================================================================================================
'�e�[�u�����擾
Function getTableInfo()
  Dim line As Long, endLine As Long
  Dim TblRecordset As ADODB.Recordset
  Dim tableName   As String
  
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_MySQL.getTableInfo"
  PrgP_Max = 4
  PrgP_Cnt = 2
  Call Library.showDebugForm(funcName & "==========================================")
  '----------------------------------------------
  tableName = ActiveSheet.Range(setVal("Cell_physicalTableName"))
  Call Library.showDebugForm("TableName", tableName)

  '�e�[�u�����--------------------------------------------------------------------------------------
  queryString = " SELECT" & _
                "   TABLE_NAME as TableName, TABLE_COMMENT as Comments" & _
                " from" & _
                "   information_schema.TABLES" & _
                " WHERE" & _
                "   TABLE_SCHEMA = DATABASE()" & _
                "   and TABLE_NAME='" & tableName & "'"
      
  Call Library.showDebugForm("QueryString", queryString)
  
  Set TblRecordset = New ADODB.Recordset
  TblRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly
  PrgP_Max = TblRecordset.RecordCount
    
  If Library.chkSheetExists(TblRecordset.Fields("TableName").Value) = True Then
    Set targetSheet = ThisWorkbook.Worksheets(TblRecordset.Fields("TableName").Value)
  Else
    Call Ctl_Common.addSheet(TblRecordset.Fields("TableName").Value)
    Set targetSheet = ActiveSheet
  End If
  
  targetSheet.Select
  targetSheet.Range(setVal("Cell_TableType")) = "�}�X�^�[�e�[�u��"
  targetSheet.Range(setVal("Cell_physicalTableName")) = TblRecordset.Fields("TableName")
  
  If TblRecordset.Fields("Comments").Value Like "*" & vbTab & "*" Then
    targetSheet.Range(setVal("Cell_logicalTableName")) = Split(TblRecordset.Fields("Comments").Value, vbTab)(0)
    targetSheet.Range(setVal("Cell_tableNote")) = Replace(Split(TblRecordset.Fields("Comments").Value, vbTab)(1), "\n", vbNewLine)
  Else
    targetSheet.Range(setVal("Cell_tableNote")) = TblRecordset.Fields("Comments")
  End If
  
  PrgP_Cnt = TblRecordset.AbsolutePosition
  Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, TblRecordset.AbsolutePosition, TblRecordset.RecordCount, "�J�������擾")
  
  '�J�������擾
  Call Ctl_MySQL.getColumnInfo(tableName)
  
  Set TblRecordset = Nothing
  Set targetSheet = Nothing


  '�����I��--------------------------------------
'  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  '----------------------------------------------
  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function


'==================================================================================================
'�J�������擾
Function getColumnInfo(tableName As String, Optional ErImgflg As Boolean = False)
  Dim line As Long, endLine As Long
  Dim ClmRecordset As ADODB.Recordset
  'Dim tableName As String
  Dim LFCount As Long, IndexLine As Long
  Dim ColCnt As Long
  Dim searchColCell As Range
  Dim IndexName As String, IndexColName As String
  Dim ER_LogicalName As String, ER_PhysicalName As String
  Dim searchWord As Range
  
  
  'ER�}�������̏����p
  Dim logicalName As String, physicalName As String, PK As String, Note As String
  
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_MySQL.getColumnInfo"
  If PrgP_Max = 0 Then
    PrgP_Max = 2
    PrgP_Cnt = 1
  End If
  Call Library.showDebugForm(funcName & "==========================================")
  Call Library.showDebugForm("runFlg", runFlg)
  If ErImgflg = False Then
    Call Ctl_Common.ClearData
  End If
  '----------------------------------------------
  If targetSheet Is Nothing Then
    Set targetSheet = ActiveSheet
  End If
  targetSheet.Select
  
'  tableName = targetSheet.Range(setVal("Cell_physicalTableName"))
  line = startLine
  ColCnt = 1
  
  '�J�������--------------------------------------------------------------------------------------
  queryString = " SELECT " & _
                "   COLUMN_NAME                            AS ColumName " & _
                "   , DATA_TYPE                            AS DataType " & _
                "   , IFNULL(CHARACTER_MAXIMUM_LENGTH, '') AS Length    " & _
                "   , COLUMN_KEY                           AS PrimaryKey " & _
                "   , IS_NULLABLE                          AS Nullable " & _
                "   , COLUMN_DEFAULT                       AS ColumnDefault " & _
                "   , COLUMN_COMMENT                       AS Comments " & _
                "   , EXTRA                                AS EXTRA " & _
                " FROM" & _
                "   information_schema.Columns c " & _
                " WHERE" & _
                "   c.table_schema = '" & setVal("DBName") & "' " & _
                "   AND c.table_name   = '" & tableName & "' " & _
                " ORDER BY" & _
                "   ordinal_position;"
      
  Call Library.showDebugForm("QueryString", queryString)
  
  Set ClmRecordset = New ADODB.Recordset
  ClmRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly
  If ErImgflg = True Then
    ReDim lValues(Int(ClmRecordset.RecordCount - 1), 2)
  End If
  
  Do Until ClmRecordset.EOF
      If ClmRecordset.Fields("Comments").Value Like "*" & vbTab & "*" Then
        logicalName = Split(ClmRecordset.Fields("Comments").Value, vbTab)(0)
        Note = Replace(Split(ClmRecordset.Fields("Comments").Value, vbTab)(1), "\n", vbNewLine)
      Else
        logicalName = ClmRecordset.Fields("Comments").Value
        Note = ""
      End If
      physicalName = ClmRecordset.Fields("ColumName").Value
      
      If ClmRecordset.Fields("PrimaryKey").Value = "PRI" Then
        PK = "��"
      Else
        PK = "�@"
      End If
      
    If ErImgflg = False Then
      'ER�}�������̏����łȂ�--------------------
      targetSheet.Range("C" & line) = ColCnt
      targetSheet.Range(setVal("Cell_logicalName") & line) = logicalName
      targetSheet.Range(setVal("Cell_physicalName") & line) = physicalName
      targetSheet.Range(setVal("Cell_Note") & line) = Note
      targetSheet.Range(setVal("Cell_dateType") & line) = ClmRecordset.Fields("DataType").Value
      targetSheet.Range(setVal("Cell_digits") & line) = ClmRecordset.Fields("Length").Value
      
      targetSheet.Range(setVal("Cell_PK") & line) = PK
      If ClmRecordset.Fields("Nullable").Value = "NO" Then
        targetSheet.Range(setVal("Cell_Null") & line) = 1
      End If
      
      '�����l
      targetSheet.Range(setVal("Cell_Default") & line) = ClmRecordset.Fields("ColumnDefault").Value
      If ClmRecordset.Fields("EXTRA").Value <> "" Then
        targetSheet.Range(setVal("Cell_Default") & line) = targetSheet.Range(setVal("Cell_Default") & line) & Replace(ClmRecordset.Fields("EXTRA").Value, "DEFAULT_GENERATED", "")
      End If
      
      If ClmRecordset.Fields("Comments").Value Like "*" & vbTab & "*" Then
        targetSheet.Range(setVal("Cell_logicalName") & line) = Split(ClmRecordset.Fields("Comments").Value, vbTab)(0)
        targetSheet.Range(setVal("Cell_Note") & line) = Replace(Split(ClmRecordset.Fields("Comments").Value, vbTab)(1), "\n", vbNewLine)
      Else
        targetSheet.Range(setVal("Cell_logicalName") & line) = ClmRecordset.Fields("Comments").Value
      End If
      
      If ClmRecordset.Fields("PrimaryKey").Value = "PRI" Then
        targetSheet.Range(setVal("Cell_PK") & line) = Application.WorksheetFunction.Max(targetSheet.Range(setVal("Cell_PK") & startLine & ":" & setVal("Cell_PK") & line)) + 1

      End If
      If ClmRecordset.Fields("Nullable").Value = "NO" Then
        targetSheet.Range(setVal("Cell_Null") & line) = 1
      End If
      
      '�s�̍�������
      LFCount = UBound(Split(targetSheet.Range(setVal("Cell_Note") & line).Value, vbNewLine)) + 2
      If LFCount > 0 Then
        targetSheet.Rows(line & ":" & line).RowHeight = setVal("defaultRowHeight") * LFCount
      End If
      
      
      Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, ClmRecordset.AbsolutePosition, ClmRecordset.RecordCount, "�J�������擾")
    
    Else
      'ER�}�������̏���--------------------------
      lValues(Int(ClmRecordset.AbsolutePosition - 1), 0) = physicalName
      lValues(Int(ClmRecordset.AbsolutePosition - 1), 1) = logicalName
      lValues(Int(ClmRecordset.AbsolutePosition - 1), 2) = PK
      
      'Erase lValues
    End If
    
    ClmRecordset.MoveNext
    
    line = line + 1
    ColCnt = ColCnt + 1
    Call Ctl_Common.addRow(line)
  Loop
  Set ClmRecordset = Nothing
  
  
  
  
  '�C���f�b�N�X���擾----------------------------------------------------------------------------
  ColCnt = 1
  If PrgP_Max = 0 Then
    PrgP_Cnt = 2
  End If
  If ErImgflg = False Then
    Call Ctl_Common.chkRowStartLine
    IndexLine = setLine("indexStart")
    
    queryString = "SHOW INDEX FROM " & tableName & ";"
    Call Library.showDebugForm("QueryString", queryString)
    
    Set ClmRecordset = New ADODB.Recordset
    ClmRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly
    
    Do Until ClmRecordset.EOF
      IndexName = ClmRecordset.Fields("Key_name").Value
      Call Library.showDebugForm("IndexName", IndexName)
      
      Set searchWord = targetSheet.Range("D" & setLine("indexStart") & ":D" & setLine("indexStart") + 10).Find(What:=IndexName, LookAt:=xlWhole, SearchOrder:=xlByRows)
      If searchWord Is Nothing And IndexName = "PRIMARY" Then
        line = IndexLine
        
      ElseIf searchWord Is Nothing Then
        line = IndexLine + 1
        ColCnt = ColCnt + 1
      Else
        line = searchWord.Row
      End If
      
      targetSheet.Range("D" & line) = IndexName
      targetSheet.Range("F" & line) = ClmRecordset.Fields("Index_type").Value
      If targetSheet.Range("G" & line) = "" Then
        targetSheet.Range("G" & line) = ClmRecordset.Fields("Column_name").Value
      Else
        targetSheet.Range("G" & line) = targetSheet.Range("G" & line) & ", " & ClmRecordset.Fields("Column_name").Value
      End If
      
      If ClmRecordset.Fields("Non_unique").Value = 0 Then
        targetSheet.Range("E" & line) = "UNIQUE"
      Else
        targetSheet.Range("E" & line) = "NONUNIQUE"
      End If
      
      '�J�������̃Z��������
      Set searchColCell = Columns("E:E").Find(What:=ClmRecordset.Fields("Column_name").Value)
      If ColCnt <= 10 Then
        Cells(searchColCell.Row, 7 + ColCnt).Select
        Cells(searchColCell.Row, 7 + ColCnt) = ClmRecordset.Fields("Seq_in_index").Value
        
        IndexColName = Library.getColumnName(7 + ColCnt)
      End If
      
      Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, ClmRecordset.AbsolutePosition, ClmRecordset.RecordCount, "�C���f�b�N�X���擾")
      Set searchWord = Nothing
      Set searchColCell = Nothing
  
      ClmRecordset.MoveNext
    Loop
  
    If Range("B5") = "" Then
      Range("B5") = "exist"
    End If
  End If

  
  '�����I��--------------------------------------
  Application.Goto Reference:=Range("A50"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
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

  Const funcName As String = "Ctl_MySQL.CreateTable"
  PrgP_Max = 4
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
    Call Ctl_MySQL.dbOpen
  End If
  queryString = ""
  Call Library.showDebugForm(funcName & "============================================")
  '----------------------------------------------
  endLine = Cells(Rows.count, 3).End(xlUp).Row
  
  tableName = Range(setVal("Cell_physicalTableName"))
  
  If Range("B5") = "" Then
  
  '�V�K�쐬----------------------------------
  ElseIf Range("B5") = "newTable" Then
    queryString = Ctl_MySQL.makeDDL(False)
      
  
  ElseIf Range("B5") = "exist" Then
    '�����e�[�u���̕ύX����------------------------
    Call Library.showDebugForm("�����e�[�u���̕ύX", tableName)
    
    For line = startLine To endLine
      If Range("B" & line) = "edit" Then
        '�f�[�^�^�ύX------------------------------
        queryString = "ALTER TABLE " & Range(setVal("Cell_physicalTableName")) & " MODIFY COLUMN " & Range(setVal("Cell_physicalName") & line) & " " & Range(setVal("Cell_dateType") & line)
        If Range(setVal("Cell_digits") & line) <> "" Then
          queryString = queryString & " (" & Range(setVal("Cell_digits") & line) & ")"
        End If
        
        'NotNull����-----------------------------
        If Range(setVal("Cell_Null") & line) = 1 Then
          queryString = queryString & " NOT NULL"
        End If
        
        '�����l�ݒ�------------------------------
        If Range(setVal("Cell_Default") & line) <> "" Then
          queryString = queryString
        End If
        
        '�R�����g--------------------------------
        If Range(setVal("Cell_Note") & line) <> "" Then
          queryString = queryString & " Comment '" & Range(setVal("Cell_logicalName") & line) & "<|>" & _
              Replace(Range(setVal("Cell_Note") & line), vbNewLine, "<BR>") & "'"
        Else
          queryString = queryString & " Comment '" & Range(setVal("Cell_logicalName") & line) & "'"

        End If
      
      '�J�������ύX[�ǉ��ˍ폜]------------------
      ElseIf Range("B" & line) Like "rename:*" Then
      
      End If
      
      If queryString <> "" Then
        Call Library.showDebugForm("QueryString", queryString)
        Call Ctl_MySQL.runQuery(queryString)
        queryString = ""
        Range("B" & line) = ""
      End If
    Next
  End If
  
  If queryString <> "" Then
    Call Library.showDebugForm("QueryString", queryString)
    Call Ctl_MySQL.runQuery(queryString)
    queryString = ""
    'Range("B" & line) = ""
  End If
  
  '�����I��--------------------------------------
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  If runFlg = False Then
    Call Ctl_MySQL.dbClose
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
'DDL����
Function makeDDL(Optional outpuFileFlg As Boolean = True)
  Dim line As Long, endLine As Long
  Dim idxLine As Long, idxEndLine As Long
  Dim tableName As String
  Dim queryColumn As String, queryColumnTmp As String
  Dim colMaxLen As Long
  Dim strHeader As String, strCopyright As String
  
  '�����J�n--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_MySQL.getColumnInfo"
  PrgP_Max = 4
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
  End If
  Call Library.showDebugForm(funcName & "=========================================")
  Call Library.showDebugForm("runFlg", runFlg)
  '----------------------------------------------
  If targetSheet Is Nothing Then
    Set targetSheet = ActiveSheet
  End If
  targetSheet.Select
  
  tableName = targetSheet.Range(setVal("Cell_physicalTableName"))
  endLine = Cells(Rows.count, 3).End(xlUp).Row
  endLine = targetSheet.Range(setVal("Cell_physicalName") & startLine).End(xlDown).Row
  
  '�J�������̍ő啶�������擾
  colMaxLen = WorksheetFunction.Max(targetSheet.Range(targetSheet.Range(setVal("Cell_colMaxLen") & startLine).Address & ":" & targetSheet.Range(setVal("Cell_colMaxLen") & endLine).Address))
  
  '�w�b�_�[���(Copyright)-------------------------------------------------------------------------
  strHeader = "/* -------------------------------------------------------------------------------" & vbNewLine & _
              "TABLE NAME �F" & targetSheet.Range(setVal("Cell_logicalTableName")) & " [" & tableName & "]" & vbNewLine & _
              "CREATE BY  �F" & targetSheet.Range("U2") & vbNewLine & _
              "CREATE DATA�F" & Format(Now(), "yyyy/mm/dd hh:nn:ss") & vbNewLine & _
              "" & vbNewLine & _
               thisAppName & " [" & thisAppVersion & "]             Copyright (c) 2021 B.Koizumi" & vbNewLine & _
              "------------------------------------------------------------------------------- */" & vbNewLine & vbNewLine
  
  
  '�J�������--------------------------------------------------------------------------------------
  If outpuFileFlg = True Then
    queryString = "DROP TABLE IF EXISTS " & tableName & ";" & vbNewLine & vbNewLine
  End If
  
  queryString = queryString & "CREATE TABLE " & tableName & " ("
  For line = startLine To endLine
    If line = startLine Then
      queryColumn = "   "
    Else
      queryColumn = "  ,"
    End If
    
    '�J������
    queryColumn = queryColumn & Library.convFixedLength("" & targetSheet.Range(setVal("Cell_physicalName") & line), colMaxLen + 4, " ")
    
    '�f�[�^�^
    queryColumnTmp = targetSheet.Range(setVal("Cell_dateType") & line)
    
    '����
    If targetSheet.Range(setVal("Cell_digits") & line) <> "" Then
      queryColumnTmp = queryColumnTmp & "(" & targetSheet.Range(setVal("Cell_digits") & line).Value & ")"
    End If
    queryColumn = queryColumn & Library.convFixedLength(queryColumnTmp, 20, " ")
    
    'NULL����
    If targetSheet.Range(setVal("Cell_Null") & line) <> "" Then
      queryColumn = queryColumn & "     " & targetSheet.Range(setVal("Cell_Null") & line).Text
    End If
    
    '�����l
    If targetSheet.Range(setVal("Cell_Default") & line) <> "" Then
      If targetSheet.Range(setVal("Cell_Default") & line) = "AUTO_INCREMENT" Then
        queryColumn = queryColumn & " " & targetSheet.Range(setVal("Cell_Default") & line)
      
      Else
        queryColumn = queryColumn & " DEFAULT " & targetSheet.Range(setVal("Cell_Default") & line)
      End If
    ElseIf targetSheet.Range(setVal("Cell_Null") & line) = "" Then
      queryColumn = queryColumn & " DEFAULT NULL"
    
    End If
   
    '���l
    If targetSheet.Range(setVal("Cell_Note") & line) <> "" Then
      queryColumn = queryColumn & " COMMENT '" & targetSheet.Range(setVal("Cell_logicalName") & line) & vbTab
      queryColumn = queryColumn & Replace(targetSheet.Range(setVal("Cell_Note") & line), vbLf, "\n") & "'"
    Else
      queryColumn = queryColumn & " COMMENT '" & targetSheet.Range(setVal("Cell_logicalName") & line) & "'"
    End If
    
    queryString = queryString & vbNewLine & queryColumn
    
  Next
  
  
  '�C���f�b�N�X���------------------------------
  queryString = queryString & vbNewLine & vbNewLine & "-- �C���f�b�N�X���------------------------------"
  
  Call Ctl_Common.chkRowStartLine
  idxLine = setLine("indexStart")
  idxEndLine = targetSheet.Range(setVal("Cell_logicalName") & idxLine).End(xlDown).Row
  
  For line = idxLine To idxEndLine
    queryColumnTmp = targetSheet.Range(setVal("Cell_digits") & line)
    queryColumnTmp = Replace(queryColumnTmp, ", ", ", ")
  
    If targetSheet.Range("C" & line) = "PK" Then
      queryString = queryString & vbNewLine & "  ,PRIMARY KEY (" & queryColumnTmp & ")"
    
    ElseIf targetSheet.Range("C" & line) = "#" Then
      Exit For
      
    ElseIf targetSheet.Range("C" & line) <> "" Then
      queryString = queryString & vbNewLine & "  ,        KEY " & targetSheet.Range(setVal("Cell_logicalName") & line) & " (" & queryColumnTmp & ")"
    End If
  Next
  
  
  
  queryColumnTmp = targetSheet.Range(setVal("Cell_logicalTableName")) & vbTab & targetSheet.Range(setVal("Cell_tableNote"))
  queryString = queryString & vbNewLine & ")" & vbNewLine & "ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci COMMENT='" & queryColumnTmp & "'"
  
  'Copyright���---------------------------------
'  strCopyright = vbNewLine & vbNewLine & vbNewLine & _
'              "/* -------------------------------------------------------------------------------" & vbNewLine & _
'              "" & vbNewLine & _
'              "" & vbNewLine & _
'               thisAppName & " [" & thisAppVersion & "]             Copyright (c) 2021 B.Koizumi" & vbNewLine & _
'              "------------------------------------------------------------------------------- */"
  
  
  If outpuFileFlg = True Then
    queryString = strHeader & queryString
    Call Library.outputText(queryString, setVal("outputDir") & "\CREATE_TABLE_" & targetSheet.Range(setVal("Cell_physicalTableName")) & ".sql")
  Else
    makeDDL = queryString
  End If
  
  '�����I��--------------------------------------
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
'�C���f�b�N�X���ݒ�
Function setIndexInfo(Optional Target As Range)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim aryColumn() As String
  Dim maxIndexNo As Long
  Dim i As Integer
  
  Const funcName As String = "Ctl_MySQL.setIndexInfo"
  
'  On Error GoTo catchError
  Call Library.startScript
  Call Library.showDebugForm(funcName & "===========================================")
  
  
  Call Ctl_Common.chkRowStartLine
  maxIndexNo = Application.WorksheetFunction.Max(Range(Cells(startLine, Target.Column), Cells(setLine("columnEnd"), Target.Column)))
  Call Library.showDebugForm("maxIndexNo", maxIndexNo)
  ReDim aryColumn(maxIndexNo)
  
  For line = startLine To CLng(setLine("columnEnd"))
    If Cells(line, Target.Column) <> "" Then
      aryColumn(Cells(line, Target.Column)) = Range(setVal("Cell_physicalName") & line)
      
      Call Library.showDebugForm("Key   ", Cells(line, Target.Column))
      Call Library.showDebugForm("Val   ", Range(setVal("Cell_physicalName") & line))
    End If
    DoEvents
  Next
  
  Select Case True
    Case Target.Column = Library.getColumnNo(setVal("Cell_PK"))
      colLine = setLine("indexStart")
      Range("C" & colLine) = "PK"
      Range("E" & colLine) = "UNIQUE"
      Range("F" & colLine) = "BTREE"
      Range(setVal("Cell_Null") & Target.Row) = 1
      
      Range("D" & colLine) = "PRIMARY"
    
    Case Target.Column = Library.getColumnNo(setVal("Cell_Idx01"))
      colLine = 1
      Range("C" & colLine) = "1"
    
    Case Target.Column = Library.getColumnNo(setVal("Cell_Idx02"))
      colLine = 2
    
    Case Target.Column = Library.getColumnNo(setVal("Cell_Idx03"))
      colLine = 3
    
    Case Target.Column = Library.getColumnNo(setVal("Cell_Idx04"))
      colLine = 4
    
    Case Target.Column = Library.getColumnNo(setVal("Cell_Idx05"))
      colLine = 5
    
    Case Target.Column = Library.getColumnNo(setVal("Cell_Idx06"))
      colLine = 6
    
    Case Target.Column = Library.getColumnNo(setVal("Cell_Idx07"))
      colLine = 7
    
    Case Target.Column = Library.getColumnNo(setVal("Cell_Idx08"))
      colLine = 8
    
    Case Target.Column = Library.getColumnNo(setVal("Cell_Idx09"))
      colLine = 9
    
    Case Target.Column = Library.getColumnNo(setVal("Cell_Idx10"))
      colLine = 10
    Case Else
  End Select

  Range("D" & setLine("indexStart") + colLine) = "Idx_" & Range(setVal("Cell_physicalTableName")) & "_" & Format(colLine, "00")
  Range("E" & setLine("indexStart") + colLine) = "NONUNIQUE"
  Range("F" & setLine("indexStart") + colLine) = "BTREE"



  If UBound(aryColumn) = 0 Then
    Range("C" & colLine & ":X" & colLine).ClearContents
    
  Else
    For i = 1 To UBound(aryColumn)
      If i = 1 Then
        Range("G" & setLine("indexStart") + colLine) = aryColumn(i)
      Else
        Range("G" & setLine("indexStart") + colLine) = Range("G" & setLine("indexStart") + colLine) & "," & aryColumn(i)
      End If
      DoEvents
    Next
  End If
  
  
  Call Library.showDebugForm("=================================================================")

  Call Library.endScript
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function
