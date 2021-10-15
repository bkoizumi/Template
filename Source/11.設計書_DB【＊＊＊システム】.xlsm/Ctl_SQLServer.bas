Attribute VB_Name = "Ctl_SQLServer"
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
Function dbOpen(Optional NoticeFlg As Boolean = True, Optional ErrMessage As String)
  Const funcName As String = "Ctl_SQLServer.dbOpen"
  
  On Error GoTo catchError
  
  If isDBOpen = True Then
    Call Library.showDebugForm("Database is already opened")
    Exit Function
  End If
  If setVal("LogLevel") = "develop" Then
    Call Library.showDebugForm("ConnectServer", ConnectServer)
  End If
  
  Set dbCon = New ADODB.Connection
  dbCon.Open ConnectServer
  dbCon.CursorLocation = 3
  
  isDBOpen = True
  Call Library.showDebugForm("isDBOpen", isDBOpen)
  
  Exit Function
  
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  isDBOpen = False
  Call Library.showDebugForm("isDBOpen", isDBOpen)
  If NoticeFlg = True Then
    Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
  Else
    dbOpen = ErrMessage
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

'エラー発生時--------------------------------------------------------------------------------------
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
  Call Ctl_SQLServer.dbOpen
  
  'テーブル情報--------------------------------------------------------------------------------------
  queryString = " SELECT" & _
                "   TABLE_NAME as TableName, TABLE_COMMENT as Comments" & _
                " from" & _
                "   information_schema.TABLES" & _
                " WHERE" & _
                "   TABLE_SCHEMA = DATABASE()" & _
                "   and TABLE_NAME='" & tableName & "'"
      
  Call Library.showDebugForm("QueryString", queryString, , "notice")
  
  Set TblRecordset = New ADODB.Recordset
  TblRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly
  If TblRecordset.RecordCount = 1 Then
    rslFlg = True
  End If
  
  Set TblRecordset = Nothing
  Set targetSheet = Nothing
  Call Ctl_SQLServer.dbClose
  
  IsTable = rslFlg
End Function


'==================================================================================================
Function runQuery(queryString As String)
  Dim runRecordset As ADODB.Recordset
  
  Const funcName As String = "Ctl_SQLServer.runQuery"

  On Error GoTo catchError
  
  Set runRecordset = dbCon.Execute(queryString)
  Set runRecordset = Nothing
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Dim errId, errMeg, meg
  errId = Err.Number
  errMeg = Err.Description
  
  Call Library.showDebugForm(funcName, errId & "：" & errMeg)
  If errId = -2147217900 Then
    meg = " 構文エラー" & vbNewLine
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
'DB情報取得
Function getDatabaseInfo(Optional ErImgflg As Boolean = False)
  Dim line As Long, endLine As Long
  Dim TblRecordset As ADODB.Recordset
  Dim tableName   As String
  Dim lValues(2) As Variant

  Dim physicalTableName As String, logicalTableName As String
  Dim tableNote As String, TableCretateAt As String

  '処理開始--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_SQLServer.getDatabaseInfo"
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
    Call Library.showDebugForm("runFlg", runFlg)
  End If
  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------
  'ER図生成用設定
  If ErImgflg = True Then
    line = 1
    sheetTmp.Range("A" & line) = "#"
    sheetTmp.Range("B" & line) = "物理テーブル名"
    sheetTmp.Range("C" & line) = "論理テーブル名"
    sheetTmp.Range("D" & line) = "作成日"
    line = line + 1
  End If
  
  'テーブル情報--------------------------------------------------------------------------------------
  queryString = " SELECT" & _
                "   sys.tables.name AS tableName, sys.tables.create_date AS CreateTime,sys.schemas.name as SchemaName, ep.value as Comments" & _
                " FROM" & _
                "   sys.tables INNER JOIN sys.schemas ON sys.tables.schema_id = schemas.schema_id" & _
                "   LEFT JOIN sys.extended_properties AS ep ON sys.tables.object_id =ep.major_id AND ep.minor_id=0" & _
                " ORDER BY schemas.schema_id"
      
  Call Library.showDebugForm("QueryString", queryString, , "notice")
  
  Set TblRecordset = New ADODB.Recordset
  TblRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly
  PrgP_Max = TblRecordset.RecordCount
  
  Do Until TblRecordset.EOF
    physicalTableName = TblRecordset.Fields("tableName").Value
    Call Library.showDebugForm("physicalTableName", physicalTableName)
    
    If TblRecordset.Fields("Comments").Value Like "*" & vbTab & "*" Then
      logicalTableName = Split(TblRecordset.Fields("Comments").Value, vbTab)(0)
      tableNote = Replace(Split(TblRecordset.Fields("Comments").Value, vbTab)(1), "\n", vbNewLine)
    ElseIf TblRecordset.Fields("Comments").Value <> "" Then
      tableNote = TblRecordset.Fields("Comments")
    Else
      tableNote = ""
    End If
    TableCretateAt = TblRecordset.Fields("CreateTime")
    
    PrgP_Cnt = TblRecordset.AbsolutePosition
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, TblRecordset.AbsolutePosition, TblRecordset.RecordCount, "テーブル情報取得[" & physicalTableName & "]")
    
  If ErImgflg = False Then
    'シート追加
    If Library.chkSheetExists(physicalTableName) = True Then
      Set targetSheet = ThisWorkbook.Worksheets(physicalTableName)
    Else
      Call Ctl_Common.addSheet(physicalTableName)
      Set targetSheet = ActiveSheet
    End If
    
    targetSheet.Select
    targetSheet.Range("B5") = "exist"
    If targetSheet.Range("F10") = "" Then
      targetSheet.Range("F10") = "マスターテーブル"
    End If
    targetSheet.Range("F9") = physicalTableName
    
    targetSheet.Range("F8") = logicalTableName
    targetSheet.Range("F9") = physicalTableName
    targetSheet.Range("F11") = tableNote
    
    targetSheet.Range("W5") = TblRecordset.Fields("SchemaName")

    'カラム情報取得
    Call Ctl_SQLServer.getColumnInfo(physicalTableName)
    
    Else
      'ER図生成時の処理--------------------------
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

  'テーブルリスト生成
  If ErImgflg = False Then
    Call Ctl_Common.makeTblList
  End If
  
  '処理終了--------------------------------------
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("EndFun  ", funcName, "info")
  '----------------------------------------------
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function


'==================================================================================================
'テーブル情報取得
Function getTableInfo()
  Dim line As Long, endLine As Long
  Dim TblRecordset As ADODB.Recordset
  Dim tableName   As String
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_SQLServer.getTableInfo"
  PrgP_Max = 4
  PrgP_Cnt = 2
  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------
  tableName = ActiveSheet.Range("F9")
  Call Library.showDebugForm("TableName", tableName)

  'テーブル情報--------------------------------------------------------------------------------------
  queryString = " SELECT" & _
                "   tbl.name AS tableName   ,ep.value as Comments" & _
                " FROM" & _
                "   sys.tables AS tbl   LEFT JOIN sys.extended_properties AS ep ON tbl.object_id =ep.major_id AND ep.minor_id=0" & _
                " WHERE" & _
                "   tbl.name='" & tableName & "'"
                
  Call Library.showDebugForm("QueryString", queryString, , "notice")
  
  Set TblRecordset = New ADODB.Recordset
  TblRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly
  PrgP_Max = TblRecordset.RecordCount
    
  If Library.chkSheetExists(TblRecordset.Fields("tableName").Value) = True Then
    Set targetSheet = ThisWorkbook.Worksheets(TblRecordset.Fields("tableName").Value)
  Else
    Call Ctl_Common.addSheet(TblRecordset.Fields("tableName").Value)
    Set targetSheet = ActiveSheet
  End If
  
  targetSheet.Select
  targetSheet.Range("F10") = "マスターテーブル"
  targetSheet.Range("F9") = TblRecordset.Fields("tableName").Value
  
  If TblRecordset.Fields("Comments").Value Like "*" & vbTab & "*" Then
    targetSheet.Range("F8") = Split(TblRecordset.Fields("Comments").Value, vbTab)(0)
    targetSheet.Range("F11") = Replace(Split(TblRecordset.Fields("Comments").Value, vbTab)(1), "\n", vbNewLine)
  Else
    targetSheet.Range("F11") = TblRecordset.Fields("Comments").Value
  End If
  
  PrgP_Cnt = TblRecordset.AbsolutePosition
  Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, TblRecordset.AbsolutePosition, TblRecordset.RecordCount, "カラム情報取得")
  
  'カラム情報取得
  Call Ctl_SQLServer.getColumnInfo(tableName)
  
  Set TblRecordset = Nothing
  Set targetSheet = Nothing


  '処理終了--------------------------------------
'  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("EndFun  ", funcName, "info")
  '----------------------------------------------
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function


'==================================================================================================
'カラム情報取得
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
  
  
  'ER図生成時の処理用
  Dim logicalName As String, physicalName As String, PK As String, Note As String
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_SQLServer.getColumnInfo"
  If PrgP_Max = 0 Then
    PrgP_Max = 2
    PrgP_Cnt = 1
  End If
  Call Library.showDebugForm("StartFun", funcName, "info")
  Call Library.showDebugForm("runFlg", runFlg)
  If ErImgflg = False Then
    Call Ctl_Common.ClearData
  End If
  '----------------------------------------------
  If targetSheet Is Nothing Then
    Set targetSheet = ActiveSheet
  End If
  targetSheet.Select
  
'  tableName = targetSheet.Range("F9")
  line = startLine
  ColCnt = 1
  
  'カラム情報--------------------------------------------------------------------------------------
  queryString = " SELECT" & _
" tables.name AS table_name,schemas.name AS schema_name,table_descriptions.value AS table_description,columns.column_id,columns.name AS ColumName,types.name AS DataType," & _
" CASE WHEN types.name in ('varchar', 'nvarchar', 'varbinary') and columns.max_length = -1 THEN 'MAX' WHEN types.name in ('decimal', 'numeric') THEN CONVERT(NVARCHAR(10), columns.precision) + ', ' + CONVERT(NVARCHAR(10), columns.scale) WHEN types.name in ('binary', 'char', 'varbinary', 'varchar') THEN CONVERT(NVARCHAR(10),columns.max_length) WHEN types.name in ('nchar', 'nvarchar') THEN CONVERT(NVARCHAR(10),(columns.max_length / 2)) WHEN types.name in ('datetime2', 'datetimeoffset', 'time') THEN CONVERT(NVARCHAR(10),columns.scale) ELSE '' END AS Length," & _
" CASE WHEN columns.is_nullable = 1 THEN 'Y' ELSE 'N' END AS Nullable," & _
" 'Identity(' + CONVERT(NVARCHAR(10), identity_columns.seed_value) + ', ' + CONVERT(NVARCHAR(10), identity_columns.increment_value) + ')' AS identity_set," & _
" primary_keys.key_ordinal as PrimaryKey," & _
" CASE WHEN left(default_constraints.definition, 2) = '((' AND RIGHT(default_constraints.definition, 2) = '))' THEN SUBSTRING(default_constraints.definition, 3, LEN(default_constraints.definition) - 4) WHEN left(default_constraints.definition, 1) = '(' AND RIGHT(default_constraints.definition, 1) = ')' THEN SUBSTRING(default_constraints.definition, 2, LEN(default_constraints.definition) - 2) ELSE NULL END AS ColumnDefault," & _
" column_descriptions.value AS Comments" & _
" FROM" & _
"   sys.tables INNER JOIN sys.schemas ON tables.schema_id = schemas.schema_id AND schemas.name = '" & targetSheet.Range("W5") & "' AND tables.name = '" & tableName & "'" & _
" LEFT OUTER JOIN sys.extended_properties AS table_descriptions ON table_descriptions.class = 1 AND tables.object_id = table_descriptions.major_id AND table_descriptions.minor_id = 0" & _
" INNER JOIN sys.columns ON sys.tables.object_id = sys.columns.object_id" & _
" INNER JOIN sys.types ON columns.user_type_id = types.user_type_id" & _
" LEFT OUTER JOIN sys.identity_columns ON columns.object_id = identity_columns.object_id AND columns.column_id = identity_columns.column_id" & _
" LEFT OUTER JOIN sys.default_constraints ON columns.default_object_id = default_constraints.object_id" & _
" LEFT OUTER JOIN sys.extended_properties AS column_descriptions ON column_descriptions.class = 1 AND columns.object_id = column_descriptions.major_id AND columns.column_id = column_descriptions.minor_id" & _
" LEFT OUTER JOIN (SELECT index_columns.object_id,index_columns.column_id,index_columns.key_ordinal FROM sys.index_columns" & _
"   INNER JOIN sys.key_constraints ON key_constraints.type = 'PK' AND index_columns.object_id = key_constraints.parent_object_id AND index_columns.index_id = key_constraints.unique_index_id ) AS primary_keys" & _
"   ON columns.object_id = primary_keys.object_id AND columns.column_id = primary_keys.column_id" & _
" ORDER BY schema_name, table_name, column_id"
      
  Call Library.showDebugForm("QueryString", queryString, , "notice")
  
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
        If IsNull(ClmRecordset.Fields("Comments").Value) Then
          logicalName = ""
          Note = ""
        Else
          logicalName = ""
          Note = ClmRecordset.Fields("Comments").Value
        End If
      End If
      physicalName = ClmRecordset.Fields("ColumName").Value
      
      If ClmRecordset.Fields("PrimaryKey").Value = "PRI" Then
        PK = "◆"
      Else
        PK = "　"
      End If
      
    If ErImgflg = False Then
      'ER図生成時の処理でない--------------------
      targetSheet.Range("C" & line) = ColCnt
      targetSheet.Range("B" & line) = logicalName
      targetSheet.Range("L" & line) = physicalName
      targetSheet.Range("AP" & line) = Note
      targetSheet.Range("V" & line) = ClmRecordset.Fields("DataType").Value
      targetSheet.Range(setVal("Cell_digits") & line) = ClmRecordset.Fields("Length").Value
      
      targetSheet.Range("AF" & line) = PK
      If ClmRecordset.Fields("Nullable").Value = "NO" Then
        targetSheet.Range("AL" & line) = 1
      End If
      
      '初期値
      If ClmRecordset.Fields("ColumnDefault").Value <> "" Then
        targetSheet.Range("AB" & line) = Replace(ClmRecordset.Fields("ColumnDefault").Value, "'", "")
      End If
      
      If ClmRecordset.Fields("PrimaryKey").Value = "1" Then
        targetSheet.Range("AF" & line) = Application.WorksheetFunction.Max(targetSheet.Range("AF" & startLine & ":" & "AF" & line)) + 1
      End If
      
      'NotNull制約
      If ClmRecordset.Fields("Nullable").Value = "N" Then
        targetSheet.Range("AL" & line) = 1
      End If
      
      '行の高さ調整
      LFCount = UBound(Split(targetSheet.Range("AP" & line).Value, vbNewLine)) + 1
      If LFCount > 0 Then
        targetSheet.Rows(line & ":" & line).RowHeight = setVal("defaultRowHeight") * LFCount
      End If
      
      
      Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, ClmRecordset.AbsolutePosition, ClmRecordset.RecordCount, "テーブル情報取得[" & tableName & "]")
    
    Else
      'ER図生成時の処理--------------------------
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
  
  
  
  'インデックス情報取得----------------------------------------------------------------------------
  ColCnt = 1
  If PrgP_Max = 0 Then
    PrgP_Cnt = 2
  End If
  If ErImgflg = False Then
    Call Ctl_Common.chkRowStartLine
    IndexLine = setLine("indexStart")
    
    Range("AF" & startLine & ":" & setVal("Cell_Idx10") & setLine("indexStart") - 4).Select
    Range("AF" & startLine & ":" & setVal("Cell_Idx10") & setLine("indexStart") - 4).ClearContents
    
    
    
    queryString = "SELECT distinct" & _
                  "    SYS.INDEXES.INDEX_ID  AS ID" & _
                  "  , SYS.INDEX_COLUMNS.COLUMN_ID AS IndexID" & _
                  "  , SYS.INDEXES.NAME      AS indexName" & _
                  "  , SYS.INDEXES.TYPE_DESC AS indexType" & _
                  "  , SYS.OBJECTS.NAME      AS tableName" & _
                  "  , SYS.COLUMNS.NAME      AS columnName" & _
                  " FROM" & _
                  "  SYS.INDEXES" & _
                  "  INNER JOIN SYS.INDEX_COLUMNS" & _
                  "    ON SYS.INDEXES.OBJECT_ID = SYS.INDEX_COLUMNS.OBJECT_ID" & _
                  "  INNER JOIN SYS.COLUMNS" & _
                  "    ON SYS.COLUMNS.COLUMN_ID = SYS.INDEX_COLUMNS.COLUMN_ID" & _
                  "    AND SYS.COLUMNS.OBJECT_ID = SYS.INDEX_COLUMNS.OBJECT_ID" & _
                  "  INNER JOIN SYS.OBJECTS" & _
                  "    ON SYS.INDEXES.OBJECT_ID = SYS.OBJECTS.OBJECT_ID" & _
                  " WHERE" & _
                  "  SYS.OBJECTS.TYPE = 'U'" & _
                  "  AND SYS.INDEXES.TYPE_DESC != 'HEAP'" & _
                  "  AND SYS.OBJECTS.NAME = '" & tableName & "'"
                
    Call Library.showDebugForm("QueryString", queryString, , "notice")
    
    Set ClmRecordset = New ADODB.Recordset
    ClmRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly
    
    Do Until ClmRecordset.EOF
      IndexName = ClmRecordset.Fields("IndexName").Value
      Call Library.showDebugForm("IndexName", IndexName)
      
      Set searchWord = targetSheet.Range("D" & setLine("indexStart") & ":D" & setLine("indexStart") + 10).Find(What:=IndexName, LookAt:=xlWhole, SearchOrder:=xlByRows)
      If searchWord Is Nothing And IndexName Like "PK_*" Then
        line = IndexLine
        
      ElseIf searchWord Is Nothing Then
        line = IndexLine + 1
        ColCnt = ColCnt + 1
      Else
        line = searchWord.Row
      End If
      
      targetSheet.Range("D" & line) = IndexName
      targetSheet.Range("E" & line) = ClmRecordset.Fields("indexType").Value
      If targetSheet.Range("G" & line) = "" Then
        targetSheet.Range("G" & line) = ClmRecordset.Fields("columnName").Value
      Else
        targetSheet.Range("G" & line) = targetSheet.Range("G" & line) & ", " & ClmRecordset.Fields("columnName").Value
      End If
      
'      If ClmRecordset.Fields("indexType").Value = 0 Then
'        targetSheet.Range("E" & line) = "UNIQUE"
'      Else
'        targetSheet.Range("E" & line) = "NONUNIQUE"
'      End If
      
      'カラム名のセルを検索
      Set searchColCell = Columns("E:E").Find(What:=ClmRecordset.Fields("columnName").Value)
      If searchColCell Is Nothing Then
        Set searchColCell = Range("E" & startLine)
      End If
      If ColCnt <= 10 Then
        Range(Cells(startLine, 7 + ColCnt), Cells(setLine("indexStart") - 4, 7 + ColCnt)).Select
        Cells(searchColCell.Row, 7 + ColCnt) = Application.WorksheetFunction.Max(Range(Cells(startLine, 7 + ColCnt), Cells(setLine("indexStart") - 4, 7 + ColCnt))) + 1
        
        IndexColName = Library.getColumnName(7 + ColCnt)
      End If
      
      Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, ClmRecordset.AbsolutePosition, ClmRecordset.RecordCount, "インデックス情報取得")
      Set searchWord = Nothing
      Set searchColCell = Nothing
  
      ClmRecordset.MoveNext
    Loop
  
    If Range("B5") = "" Then
      Range("B5") = "exist"
    End If
  End If

  
  '処理終了--------------------------------------
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("EndFun  ", funcName, "info")
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function

'==================================================================================================
'直接実行
Function CreateTable()
  Dim line As Long, endLine As Long
  Dim tableName As String
  Dim ColumnString As String
  Dim oldColumnName As String
  Dim queryString As String, tmpQueryString, megQueryString As String
  '処理開始--------------------------------------
  'On Error GoTo catchError

  Const funcName As String = "Ctl_SQLServer.CreateTable"
  PrgP_Max = 4
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
    Call Ctl_SQLServer.dbOpen
  End If
  queryString = ""
  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------
  endLine = Cells(Rows.count, 3).End(xlUp).Row
  
  tableName = Range("F9")
  
  If Range("B5") = "" Then
  
  '新規作成----------------------------------
  ElseIf Range("B5") = "newTable" Then
    queryString = Ctl_SQLServer.makeDDL(False)
    megQueryString = queryString
      
  
  ElseIf Range("B5") = "exist" Then
    '既存テーブルの変更処理------------------------
    Call Library.showDebugForm("既存テーブルの変更", tableName)
    
    For line = startLine To endLine
      If Range("B" & line) = "edit" Then
        'データ型変更------------------------------
        queryString = "ALTER TABLE " & Range("F9") & " MODIFY COLUMN " & Range("L" & line) & " " & Range("V" & line)
        If Range(setVal("Cell_digits") & line) <> "" Then
          queryString = queryString & " (" & Range(setVal("Cell_digits") & line) & ")"
        End If
        
        'NotNull制約-----------------------------
        If Range("AL" & line) = 1 Then
          queryString = queryString & " NOT NULL"
        End If
        
        '初期値設定------------------------------
        If Range("AB" & line) <> "" Then
          queryString = queryString
        End If
        
        'コメント--------------------------------
        If Range("AP" & line) <> "" Then
          queryString = queryString & " Comment '" & Range("B" & line) & "<|>" & _
              Replace(Range("AP" & line), vbNewLine, "<BR>") & "'"
        Else
          queryString = queryString & " Comment '" & Range("B" & line) & "'"

        End If
      
      'カラム名変更[追加⇒削除]------------------
      ElseIf Range("B" & line) Like "rename:*" Then
      
      End If
      
      If queryString <> "" Then
        Call Ctl_SQLServer.runQuery(queryString)
        queryString = ""
        Range("B" & line) = ""
      End If
    Next
  End If
  
  If queryString <> "" Then
    Call Library.showDebugForm("QueryString", queryString, , "notice")
    For Each tmpQueryString In Split(queryString, "GO /* EndOfQuery */")
      Call Library.showDebugForm("QueryString", tmpQueryString)
      If tmpQueryString <> "" Then
        Call Ctl_SQLServer.runQuery(CStr(tmpQueryString))
      End If
    Next
    
    queryString = ""
    'Range("B" & line) = ""
  End If
  
  '処理終了--------------------------------------
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("EndFun  ", funcName, "info")
  If runFlg = False Then
    Call Ctl_SQLServer.dbClose
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.unsetting(True)
  End If
  '----------------------------------------------
  Call Library.showNotice(210, megQueryString)
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function



'==================================================================================================
'DDL生成
Function makeDDL(Optional outpuFileFlg As Boolean = True)
  Dim line As Long, endLine As Long
  Dim idxLine As Long, idxEndLine As Long
  Dim tableName As String
  Dim queryColumn As String, queryColumnTmp As String, queryNoteColumn As String
  Dim colMaxLen As Long
  Dim strHeader As String, strCopyright As String, searchWord As String
  Dim columnName  As Variant, defaultVal As Variant
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_SQLServer.getColumnInfo"
  PrgP_Max = 4
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
  End If
  Call Library.showDebugForm("StartFun", funcName, "info")
  Call Library.showDebugForm("runFlg", runFlg)
  '----------------------------------------------
  If targetSheet Is Nothing Then
    Set targetSheet = ActiveSheet
  End If
  targetSheet.Select
  
  tableName = targetSheet.Range("F9")
  endLine = Cells(Rows.count, 3).End(xlUp).Row
  endLine = targetSheet.Range("L" & startLine).End(xlDown).Row
  
  'カラム名の最大文字数を取得
  colMaxLen = WorksheetFunction.Max(targetSheet.Range(targetSheet.Range(setVal("Cell_colMaxLen") & startLine).Address & ":" & targetSheet.Range(setVal("Cell_colMaxLen") & endLine).Address)) + 2
  
  'ヘッダー情報(Copyright)-------------------------------------------------------------------------
  strHeader = "/* -------------------------------------------------------------------------------" & vbNewLine & _
              "TABLE NAME ：" & setVal("Schema") & "." & targetSheet.Range("F8") & " [" & tableName & "]" & vbNewLine & _
              "CREATE BY  ：" & targetSheet.Range("U2") & vbNewLine & _
              "CREATE DATA：" & Format(Now(), "yyyy/mm/dd hh:nn:ss") & vbNewLine & _
              "" & vbNewLine & _
               thisAppName & " [" & thisAppVersion & "]             Copyright (c) 2021 B.Koizumi" & vbNewLine & _
              "------------------------------------------------------------------------------- */" & vbNewLine & vbNewLine
  
  
  'カラム情報--------------------------------------------------------------------------------------
  queryString = "USE [prfSummary]" & vbNewLine & _
                "GO /* EndOfQuery */" & vbNewLine & vbNewLine & _
                "SET ANSI_NULLS ON" & vbNewLine & _
                "GO /* EndOfQuery */" & vbNewLine & vbNewLine & _
                "SET QUOTED_IDENTIFIER ON" & vbNewLine & _
                "GO /* EndOfQuery */" & vbNewLine & vbNewLine & _
                "DROP TABLE IF EXISTS [" & setVal("Schema") & "].[" & tableName & "]" & vbNewLine & _
                "GO /* EndOfQuery */" & vbNewLine & vbNewLine
  queryString = queryString & "CREATE TABLE [" & setVal("Schema") & "].[" & tableName & "]("
  For line = startLine To endLine
    If line = startLine Then
      queryColumn = "   "
    Else
      queryColumn = "  ,"
    End If
    
    'カラム名
    queryColumn = queryColumn & Library.convFixedLength("[" & targetSheet.Range("L" & line) & "]", colMaxLen + 4, " ")
    
    'データ型
    queryColumnTmp = "[" & targetSheet.Range("V" & line) & "]"
    
    '桁数
    If targetSheet.Range(setVal("Cell_digits") & line) <> "" Then
      queryColumnTmp = queryColumnTmp & "(" & targetSheet.Range(setVal("Cell_digits") & line).Value & ")"
    End If
    queryColumn = queryColumn & Library.convFixedLength(queryColumnTmp, 20, " ")
    
    'NULL制約
    If targetSheet.Range("AL" & line) <> "" Then
      queryColumn = queryColumn & "     " & targetSheet.Range("AL" & line).text
    End If

    queryString = queryString & vbNewLine & queryColumn
  Next
  queryString = queryString & vbNewLine
  
  
  'プライマリーキー-----------------------------
  Call Ctl_Common.chkRowStartLine
  idxLine = setLine("indexStart")
  idxEndLine = targetSheet.Range("B" & idxLine).End(xlDown).Row
  
  queryString = queryString & vbNewLine & vbNewLine & "/* プライマリーキー-------------------------------- */"
  If targetSheet.Range("C" & idxLine) = "PK" Then
    queryString = queryString & vbNewLine & "  ,CONSTRAINT [" & targetSheet.Range("D" & idxLine) & "] PRIMARY KEY " & targetSheet.Range("E" & idxLine) & "(" & vbNewLine
    queryColumnTmp = ""
    For Each columnName In Split(targetSheet.Range("G" & idxLine), ", ")
      If queryColumnTmp = "" Then
        queryColumnTmp = "     " & Library.convFixedLength("[" & columnName & "]", colMaxLen + 4, " ") & " ASC" & vbNewLine
      Else
        queryColumnTmp = queryColumnTmp & "    ," & Library.convFixedLength("[" & columnName & "]", colMaxLen + 4, " ") & " ASC" & vbNewLine
      End If
    Next
    queryString = queryString & queryColumnTmp & "  )" & vbNewLine
    queryString = queryString & "  WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]" & vbNewLine
    queryString = queryString & ") ON [PRIMARY]" & vbNewLine
    queryString = queryString & "GO /* EndOfQuery */"
  End If
  
  '初期値----------------------------------------
  queryString = queryString & vbNewLine & vbNewLine & "/* 初期値------------------------------------------ */"
  
  For line = startLine To endLine
    queryColumnTmp = ""
    If targetSheet.Range("AB" & line) <> "" Then
      defaultVal = targetSheet.Range("AB" & line)
      Call Library.showDebugForm("defaultVal", defaultVal & "[" & typeName(defaultVal) & "]")
      
      If typeName(defaultVal) = "String" Then
        On Error Resume Next
        searchWord = Application.WorksheetFunction.VLookup(defaultVal, Range("defaultVal"), 1, False)
        On Error GoTo 0
        If searchWord = "" Then
          defaultVal = "'" & defaultVal & "'"
        End If
      End If
      Call Library.showDebugForm("defaultVal", defaultVal)
      
      queryColumnTmp = "ALTER TABLE [" & setVal("Schema") & "].[" & targetSheet.Range("F9") & "]" & _
                       " ADD CONSTRAINT [DF_" & targetSheet.Range("F9") & "_" & targetSheet.Range("L" & line) & "]  DEFAULT" & _
                       " (" & defaultVal & ") FOR [" & targetSheet.Range("L" & line) & "]"
      
      queryString = queryString & vbNewLine & queryColumnTmp
    End If
  Next
  queryString = queryString & vbNewLine & "GO /* EndOfQuery */"
  
  
  'コメント設定----------------------------------
  queryString = queryString & vbNewLine & vbNewLine & "/* 論理テーブル名---------------------------------- */"
  queryNoteColumn = "EXECUTE sp_addextendedproperty" & _
                    " @name = N'MS_Description'," & _
                    " @value = N'" & targetSheet.Range("F8") & "'," & _
                    " @level0type = N'SCHEMA'," & _
                    " @level0name = N'dbo'," & _
                    " @level1type = N'TABLE'," & _
                    " @level1name = N'" & targetSheet.Range("F9") & "';"

  queryString = queryString & vbNewLine & queryNoteColumn
  queryString = queryString & vbNewLine & "GO /* EndOfQuery */"
  
  '論理名、コメント設定--------------------------
  queryString = queryString & vbNewLine & vbNewLine & "/* 論理名、コメント設定---------------------------- */"
  For line = startLine To endLine
      queryNoteColumn = "EXECUTE sp_addextendedproperty @name = N'MS_Description',"

      If targetSheet.Range("AP" & line) <> "" Then
        queryNoteColumn = queryNoteColumn & "@value = N'" & targetSheet.Range("B" & line) & vbTab
        queryNoteColumn = queryNoteColumn & Replace(targetSheet.Range("AP" & line), vbCrLf, "\n") & "',"
      Else
        queryNoteColumn = queryNoteColumn & "@value = N'" & targetSheet.Range("B" & line) & "',"
      End If
      queryNoteColumn = queryNoteColumn & "@level0type = N'SCHEMA'," & _
                    " @level0name = N'" & setVal("Schema") & "'," & _
                    " @level1type = N'TABLE'," & _
                    " @level1name = N'" & tableName & "'," & _
                    " @level2type = N'COLUMN'," & _
                    " @level2name = N'" & targetSheet.Range("L" & line) & "';"
    
    queryString = queryString & vbNewLine & queryNoteColumn
  Next
  queryString = queryString & vbNewLine & "GO /* EndOfQuery */"
    
    
    
  'インデックス情報------------------------------
  queryString = queryString & vbNewLine & vbNewLine & "/* インデックス情報-------------------------------- */"
  
  idxLine = setLine("indexStart") + 1
  idxEndLine = setLine("indexStart") + 10
  
  For line = idxLine To idxEndLine
    If targetSheet.Range("D" & line) <> "" Then
      queryString = queryString & vbNewLine & "DROP TABLE IF EXISTS [" & setVal("Schema") & "].[" & targetSheet.Range("D" & line) & "]" & vbNewLine
      queryString = queryString & "GO /* EndOfQuery */" & vbNewLine & vbNewLine
    
      If targetSheet.Range("E" & line) = "NONCLUSTERED" Then
        queryString = queryString & "CREATE NONCLUSTERED INDEX [" & targetSheet.Range("D" & line) & "] ON [dbo].[" & targetSheet.Range("F9") & "] ("
        queryColumnTmp = ""
        For Each columnName In Split(targetSheet.Range("G" & idxLine), ", ")
          If queryColumnTmp = "" Then
            queryColumnTmp = "     " & Library.convFixedLength("[" & columnName & "]", colMaxLen + 4, " ") & " ASC" & vbNewLine
          Else
            queryColumnTmp = queryColumnTmp & "    ," & Library.convFixedLength("[" & columnName & "]", colMaxLen + 4, " ") & " ASC" & vbNewLine
          End If
        Next
        queryString = queryString & vbNewLine & queryColumnTmp & "  )" & vbNewLine
        queryString = queryString & "  WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]" & vbNewLine
        queryString = queryString & "GO /* EndOfQuery */"
      
      
      ElseIf targetSheet.Range("E" & line) = "CLUSTERED" Then
        queryString = queryString & "CREATE CLUSTERED INDEX [" & targetSheet.Range("D" & line) & "] ON [dbo].[" & targetSheet.Range("F9") & "]"
      End If
    End If
  Next
     
    
  'Copyright情報---------------------------------
'  strCopyright = vbNewLine & vbNewLine & vbNewLine & _
'              "/* -------------------------------------------------------------------------------" & vbNewLine & _
'              "" & vbNewLine & _
'              "" & vbNewLine & _
'               thisAppName & " [" & thisAppVersion & "]             Copyright (c) 2021 B.Koizumi" & vbNewLine & _
'              "------------------------------------------------------------------------------- */"
  
  
  If outpuFileFlg = True Then
    queryString = strHeader & queryString
    queryString = Replace(queryString, "GO /* EndOfQuery */", "GO")
    
    Call Library.outputText(queryString, setVal("outputDir") & "\CREATE_TABLE_" & targetSheet.Range("F9") & ".sql")
  Else
    makeDDL = queryString
  End If
  
  '処理終了--------------------------------------
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("EndFun  ", funcName, "info")
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.unsetting(True)
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function


'==================================================================================================
'インデックス情報設定
Function setIndexInfo(Optional Target As Range)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim aryColumn() As String
  Dim maxIndexNo As Long
  Dim i As Integer
  
  Const funcName As String = "Ctl_SQLServer.setIndexInfo"
  
'  On Error GoTo catchError
  Call Library.startScript
  Call Library.showDebugForm("StartFun", funcName, "info")
  
  
  Call Ctl_Common.chkRowStartLine
  maxIndexNo = Application.WorksheetFunction.Max(Range(Cells(startLine, Target.Column), Cells(setLine("columnEnd"), Target.Column)))
  Call Library.showDebugForm("maxIndexNo", maxIndexNo)
  ReDim aryColumn(maxIndexNo)
  
  For line = startLine To CLng(setLine("columnEnd"))
    If Cells(line, Target.Column) <> "" Then
      aryColumn(Cells(line, Target.Column)) = Range("L" & line)
      
      Call Library.showDebugForm("Key   ", Cells(line, Target.Column))
      Call Library.showDebugForm("Val   ", Range("L" & line))
    End If
    DoEvents
  Next
  
  Select Case True
    Case Target.Column = Library.getColumnNo("AF")
      colLine = setLine("indexStart")
      Range("C" & colLine) = "PK"
      Range("E" & colLine) = "UNIQUE"
      Range("F" & colLine) = "BTREE"
      Range("AL" & Target.Row) = 1
      
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

  Range("D" & setLine("indexStart") + colLine) = "Idx_" & Range("F9") & "_" & Format(colLine, "00")
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
  
  
  Call Library.showDebugForm("EndFun  ", funcName, "info")

  Call Library.endScript
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, funcName & " [" & Err.Number & "]" & Err.Description, True)
End Function


