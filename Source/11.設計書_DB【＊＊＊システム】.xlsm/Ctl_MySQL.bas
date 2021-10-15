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
Function dbOpen(Optional NoticeFlg As Boolean = True, Optional ErrMessage As String)
  Const funcName As String = "Ctl_MySQL.dbOpen"
  
  On Error GoTo catchError
  
  If isDBOpen = True Then
    Call Library.showDebugForm("Database is already opened", , "notice")
    Exit Function
  End If
  If setVal("LogLevel") = "develop" Then
    Call Library.showDebugForm("ConnectServer", ConnectServer)
  End If
  
  Set dbCon = New ADODB.Connection
  dbCon.Open ConnectServer
  dbCon.CursorLocation = 3
  
  isDBOpen = True
  Call Library.showDebugForm("isDBOpen", isDBOpen, "info")
  
  Exit Function
  
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  isDBOpen = False
  Call Library.showDebugForm("isDBOpen", isDBOpen, "info")
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
    Call Library.showDebugForm("Database is already closed", , "notice")
  Else
    dbCon.Close
    isDBOpen = False
    
    Set dbCon = Nothing
    Call Library.showDebugForm("isDBOpen", isDBOpen, "info")
  End If
  
  Exit Function

'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm("isDBOpen", isDBOpen, "info")
  Call Library.showNotice(501, Err.Description, True)
End Function


'==================================================================================================
Function IsTable(tableName As String) As Boolean
  Dim TblRecordset As ADODB.Recordset
  Dim rslFlg As Boolean

  rslFlg = False

  Call Library.showDebugForm("TableName", tableName)
  Call Ctl_MySQL.dbOpen
  
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
  Const funcName As String = "Ctl_MySQL.getDatabaseInfo"
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
                "   TABLE_NAME as TableName, TABLE_SCHEMA AS SchemaName, TABLE_COMMENT as Comments, CREATE_TIME AS CreateTime" & _
                " FROM" & _
                "   information_schema.TABLES" & _
                " WHERE" & _
                "   TABLE_SCHEMA = DATABASE();"
      
  Call Library.showDebugForm("QueryString", queryString, "notice")
  
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
    TableCretateAt = TblRecordset.Fields("CreateTime")
    
    PrgP_Cnt = TblRecordset.AbsolutePosition
    Call Ctl_ProgressBar.showBar(thisAppName, TblRecordset.AbsolutePosition, TblRecordset.RecordCount, 1, 2, "カラム情報取得")
    
  If ErImgflg = False Then
    'シート追加
    Call Ctl_Common.chkTableName2SheetName(physicalTableName)
    Set targetSheet = ActiveSheet
    
    targetSheet.Select
    targetSheet.Range("B5") = "exist"
    If targetSheet.Range("G10") = "" Then
      If physicalTableName Like "*マスタ" Or logicalTableName Like "m_*" Then
        targetSheet.Range("G10") = "マスターテーブル"
      
      ElseIf physicalTableName Like "*ワーク" Or logicalTableName Like "w_*" Then
        targetSheet.Range("G10") = "ワークテーブル"
      
      ElseIf physicalTableName Like "*[!_]*" Then
        targetSheet.Range("G10") = "マスターテーブル"
        
      Else
        targetSheet.Range("G10") = "トランザクションテーブル"
      
      End If
    End If
    
    targetSheet.Range("F7") = TblRecordset.Fields("SchemaName")
    targetSheet.Range("F8") = logicalTableName
    targetSheet.Range("F9") = physicalTableName
    targetSheet.Range("F11") = tableNote

    'カラム情報取得
    Call Ctl_MySQL.getColumnInfo(physicalTableName)
    
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
'    Call Ctl_Common.makeTblList
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
  Dim physicalTableName As String, logicalTableName As String
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_MySQL.getTableInfo"
  PrgP_Max = 3
  PrgP_Cnt = 1
  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------
  If Not ActiveSheet.Name Like "3.*" Then
    Call Library.showNotice(410, , True)
  End If
  
  tableName = ActiveSheet.Range("F9")
  Call Library.showDebugForm("TableName", tableName, "info")

  'テーブル情報--------------------------------------------------------------------------------------
  queryString = " SELECT" & _
                "   TABLE_NAME as TableName, TABLE_COMMENT as Comments" & _
                " FROM" & _
                "   information_schema.TABLES" & _
                " WHERE" & _
                "   TABLE_SCHEMA = DATABASE()" & _
                "   AND TABLE_NAME='" & tableName & "'"
      
  Call Library.showDebugForm("QueryString", queryString, "notice")
  
  Set TblRecordset = New ADODB.Recordset
  TblRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly
    
  If TblRecordset.RecordCount = 0 Then
    Call Library.showNotice(510, , True)
  End If
  physicalTableName = TblRecordset.Fields("TableName").Value
  If TblRecordset.Fields("Comments").Value Like "*" & vbTab & "*" Then
    logicalTableName = Split(TblRecordset.Fields("Comments").Value, vbTab)(0)
  End If
  
  Set targetSheet = ThisWorkbook.Worksheets(Ctl_Common.chkTableName2SheetName(TblRecordset.Fields("TableName").Value))
  
  targetSheet.Select
  If targetSheet.Range("G10") = "" Then
    If physicalTableName Like "*マスタ" Or logicalTableName Like "m_*" Then
      targetSheet.Range("G10") = "マスターテーブル"
    ElseIf physicalTableName Like "*ワーク" Or logicalTableName Like "w_*" Then
      targetSheet.Range("G10") = "ワークテーブル"
    ElseIf physicalTableName Like "*[!_]*" Then
      targetSheet.Range("G10") = "マスターテーブル"
    Else
      targetSheet.Range("G10") = "トランザクションテーブル"
    End If
  End If
  
  targetSheet.Range("F9") = TblRecordset.Fields("TableName")
  
  If TblRecordset.Fields("Comments").Value Like "*" & vbTab & "*" Then
    targetSheet.Range("F8") = Split(TblRecordset.Fields("Comments").Value, vbTab)(0)
    targetSheet.Range("F11") = Replace(Split(TblRecordset.Fields("Comments").Value, vbTab)(1), "\n", vbNewLine)
  Else
    targetSheet.Range("F8") = TblRecordset.Fields("Comments")
  End If
  
  'カラム情報取得
  Call Ctl_MySQL.getColumnInfo(tableName)
  
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
Function getColumnInfo(tableName As String)
  Dim line As Long, endLine As Long
  Dim ClmRecordset As ADODB.Recordset
  'Dim tableName As String
  Dim LFCount As Long, IndexLine As Long
  Dim ColCnt As Long
  Dim IndexName As String, IndexColName As String
  Dim ER_LogicalName As String, ER_PhysicalName As String
  Dim searchWord As Range
  Dim searchColCell As Range
  
  Dim logicalName As String, physicalName As String, PK As String, Note As String
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_MySQL.getColumnInfo"
  If PrgP_Max = 0 Then
    PrgP_Max = 2
  End If
  PrgP_Cnt = 1
  
  Call Library.showDebugForm("StartFun", funcName, "info")
  Call Library.showDebugForm("runFlg", runFlg, "info")
  Call Ctl_Common.ClearData
  '----------------------------------------------
  If targetSheet Is Nothing Then
    Set targetSheet = ActiveSheet
  End If
  targetSheet.Select
  
  line = startLine
  ColCnt = 1
  
  'カラム情報--------------------------------------------------------------------------------------
  queryString = " SELECT " & _
                "   COLUMN_NAME                            AS ColumName " & _
                "   , DATA_TYPE                            AS DataType " & _
                "   , IFNULL(CHARACTER_MAXIMUM_LENGTH, '') AS Length    " & _
                "   , COLUMN_KEY                           AS PrimaryKey " & _
                "   , IS_NULLABLE                          AS Nullable " & _
                "   , COLUMN_DEFAULT                       AS ColumnDefault " & _
                "   , COLUMN_COMMENT                       AS Comments " & _
                "   , EXTRA                                AS EXTRA " & _
                "   , COLUMN_TYPE                          AS ColumType " & _
                " FROM" & _
                "   information_schema.Columns c " & _
                " WHERE" & _
                "   c.table_schema = '" & setVal("DBName") & "' " & _
                "   AND c.table_name   = '" & tableName & "' " & _
                " ORDER BY" & _
                "   ordinal_position;"
      
  Call Library.showDebugForm("カラム情報", queryString, "notice")
  
  Set ClmRecordset = New ADODB.Recordset
  ClmRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly
  
  Do Until ClmRecordset.EOF
    If ClmRecordset.Fields("Comments").Value Like "*" & vbTab & "*" Then
      logicalName = Split(ClmRecordset.Fields("Comments").Value, vbTab)(0)
      Note = Replace(Split(ClmRecordset.Fields("Comments").Value, vbTab)(1), "\n", vbNewLine)
    Else
      logicalName = ClmRecordset.Fields("Comments").Value
      Note = ""
    End If
    physicalName = ClmRecordset.Fields("ColumName").Value
    
    targetSheet.Range("B" & line) = logicalName
    targetSheet.Range("L" & line) = physicalName
    targetSheet.Range("AP" & line) = Note
    targetSheet.Range("V" & line) = ClmRecordset.Fields("ColumType").Value
    
    
    If ClmRecordset.Fields("Nullable").Value = "NO" Then
      targetSheet.Range("AL" & line) = 1
    End If
    
    '初期値
    targetSheet.Range("AB" & line) = ClmRecordset.Fields("ColumnDefault").Value
    If ClmRecordset.Fields("EXTRA").Value <> "" Then
      targetSheet.Range("AB" & line) = targetSheet.Range("AB" & line) & Replace(ClmRecordset.Fields("EXTRA").Value, "DEFAULT_GENERATED", "")
    End If
    
    If ClmRecordset.Fields("Comments").Value Like "*" & vbTab & "*" Then
      targetSheet.Range("B" & line) = Split(ClmRecordset.Fields("Comments").Value, vbTab)(0)
      targetSheet.Range("AP" & line) = Replace(Split(ClmRecordset.Fields("Comments").Value, vbTab)(1), "\n", vbNewLine)
    Else
      targetSheet.Range("B" & line) = ClmRecordset.Fields("Comments").Value
    End If
    
    If ClmRecordset.Fields("PrimaryKey").Value = "PRI" Then
      targetSheet.Range("AF" & line) = Application.WorksheetFunction.Max(targetSheet.Range("AF" & startLine & ":" & "AF" & line)) + 1

'    ElseIf ClmRecordset.Fields("PrimaryKey").Value = "MUL" Then
'      targetSheet.Range("AJ" & line) = Application.WorksheetFunction.Max(targetSheet.Range("AJ" & startLine & ":" & "AJ" & line)) + 1

    End If
    If ClmRecordset.Fields("Nullable").Value = "NO" Then
      targetSheet.Range("AL" & line) = 1
    End If
    
    '行の高さ調整
    LFCount = UBound(Split(targetSheet.Range("AP" & line).Value, vbNewLine)) + 2
    If LFCount > 0 Then
      targetSheet.Rows(line & ":" & line).RowHeight = setVal("defaultRowHeight") * LFCount
    End If
    
    '書式設定----------------------------------
    '初期値のリスト化
    With Range("AB" & line & ":AE" & line).Validation
      .Delete
      .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=defVal_MySQL"
    End With
    
    Range("AF" & line & ":AO" & line).NumberFormatLocal = """YES"""
    
    '備考の結合
    Range("AP" & line & ":BB" & line).Merge True
    
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, ClmRecordset.AbsolutePosition, ClmRecordset.RecordCount, "カラム情報取得")
    
    ClmRecordset.MoveNext
    
    line = line + 1
    ColCnt = ColCnt + 1
    Call Ctl_Common.addRow(line)
  Loop
  Set ClmRecordset = Nothing
  
  
  'インデックス情報--------------------------------------------------------------------------------
  ColCnt = 0
  PrgP_Cnt = 2
  
  Call Ctl_Common.chkRowStartLine
  
  queryString = "SHOW INDEX FROM " & tableName & ";"
  Call Library.showDebugForm("インデックス情報", queryString, "notice")
  
  Set ClmRecordset = New ADODB.Recordset
  ClmRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly
  
  Do Until ClmRecordset.EOF
    IndexName = ClmRecordset.Fields("Key_name").Value
    Call Library.showDebugForm("IndexName", IndexName, "info")
    
    Set searchWord = targetSheet.Range("B" & setLine("indexStart") & ":K" & setLine("indexEnd")).Find(What:=IndexName, LookAt:=xlWhole)
    If searchWord Is Nothing Then
      line = Cells(Rows.count, 2).End(xlUp).Row + 1
      ColCnt = ColCnt + 1
    Else
      line = searchWord.Row
    End If
    Call Library.showDebugForm("line", line, "info")
    
    targetSheet.Range("A" & line).FormulaR1C1 = "=ROW()- " & setLine("indexStart") - 1
    targetSheet.Range("B" & line) = IndexName
    targetSheet.Range("BJ" & line) = ClmRecordset.Fields("Index_type").Value
    
    If targetSheet.Range("L" & line) = "" Then
      targetSheet.Range("L" & line) = ClmRecordset.Fields("Column_name").Value
    Else
      targetSheet.Range("L" & line) = targetSheet.Range("L" & line) & ", " & ClmRecordset.Fields("Column_name").Value
    End If
    
    If ClmRecordset.Fields("Non_unique").Value = 0 Then
      targetSheet.Range("BI" & line) = "UNIQUE"
    Else
      targetSheet.Range("BI" & line) = "NONUNIQUE"
    End If
    
    targetSheet.Range("BJ" & line) = ClmRecordset.Fields("Index_type").Value
    
    'カラム名のセルを検索
    Set searchColCell = Columns("L:U").Find(What:=ClmRecordset.Fields("Column_name").Value)
    Range("BI" & searchColCell.Row) = 1
    
    Set searchWord = Nothing
    Set searchColCell = Nothing
    
    Call Ctl_Common.addRow(line)
    
    ClmRecordset.MoveNext
    
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, ClmRecordset.AbsolutePosition, ClmRecordset.RecordCount, "インデックス情報取得")
  Loop
  
  
  '外部キー情報------------------------------------------------------------------------------------
  ColCnt = 0
  PrgP_Cnt = 3
  Dim fKeyColName As String
  
  Call Ctl_Common.chkRowStartLine
  
  queryString = "SELECT * FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE  TABLE_NAME = '" & tableName & "' and REFERENCED_TABLE_NAME is not null"
  Call Library.showDebugForm("外部キー情報", queryString, "notice")
  
  Set ClmRecordset = New ADODB.Recordset
  ClmRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly
  
  Do Until ClmRecordset.EOF
    fKeyColName = ClmRecordset.Fields("COLUMN_NAME").Value
    Call Library.showDebugForm("fKeyColName", fKeyColName, "info")
    
    Set searchColCell = targetSheet.Range("L" & startLine & ":U" & setLine("columnEnd")).Find(What:=fKeyColName, LookAt:=xlWhole)
    If searchColCell Is Nothing Then
      GoTo Lbl_nextRecode
    Else
      line = searchColCell.Row
    End If
    
    targetSheet.Range("AJ" & line) = 1
    targetSheet.Range("BJ" & line) = ClmRecordset.Fields("REFERENCED_TABLE_NAME").Value & "." & ClmRecordset.Fields("REFERENCED_COLUMN_NAME").Value
    
    'インデックス情報に追記----------------------
    endLine = Cells(Rows.count, 2).End(xlUp).Row + 1
    If endLine = setLine("indexEnd") Then
      Call Ctl_Common.addRow(line)
    End If
    
    targetSheet.Range("B" & endLine) = ClmRecordset.Fields("CONSTRAINT_NAME").Value
    targetSheet.Range("L" & endLine) = ClmRecordset.Fields("COLUMN_NAME").Value
    targetSheet.Range("BJ" & endLine) = "FOREIGN KEY"
    
Lbl_nextRecode:
    Set searchColCell = Nothing
    ClmRecordset.MoveNext
    
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, ClmRecordset.AbsolutePosition, ClmRecordset.RecordCount, "インデックス情報取得")
  Loop
  
  
'    If Range("B5") = "" Then
'      Range("B5") = "exist"
'    End If


  
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
  
  '処理開始--------------------------------------
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
  Call Library.showDebugForm("StartFun", funcName, "info")
  '----------------------------------------------
  endLine = Cells(Rows.count, 3).End(xlUp).Row
  
  tableName = Range("F9")
  
  If Range("B5") = "" Then
  
  '新規作成----------------------------------
  ElseIf Range("B5") = "newTable" Then
    queryString = Ctl_MySQL.makeDDL(False)
      
  
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
        Call Library.showDebugForm("QueryString", queryString, , "notice")
        Call Ctl_MySQL.runQuery(queryString)
        queryString = ""
        Range("B" & line) = ""
      End If
    Next
  End If
  
  If queryString <> "" Then
    Call Library.showDebugForm("QueryString", queryString, , "notice")
    Call Ctl_MySQL.runQuery(queryString)
    queryString = ""
    'Range("B" & line) = ""
  End If
  
  '処理終了--------------------------------------
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("EndFun  ", funcName, "info")
  If runFlg = False Then
    Call Ctl_MySQL.dbClose
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
'DDL生成
Function makeDDL(Optional outpuFileFlg As Boolean = True)
  Dim line As Long, endLine As Long
  Dim idxLine As Long, idxEndLine As Long
  Dim tableName As String
  Dim queryColumn As String, queryColumnTmp As String
  Dim colMaxLen As Long
  Dim strHeader As String, strCopyright As String
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  Const funcName As String = "Ctl_MySQL.getColumnInfo"
  PrgP_Max = 4
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
  End If
  Call Library.showDebugForm("StartFun", funcName, "info")
  Call Library.showDebugForm("runFlg", runFlg, , "info")
  '----------------------------------------------
  If targetSheet Is Nothing Then
    Set targetSheet = ActiveSheet
  End If
  targetSheet.Select
  
  tableName = targetSheet.Range("F9")
  endLine = Cells(Rows.count, 3).End(xlUp).Row
  endLine = targetSheet.Range("L" & startLine).End(xlDown).Row
  
  'カラム名の最大文字数を取得
  colMaxLen = WorksheetFunction.Max(targetSheet.Range(targetSheet.Range("BH" & startLine).Address & ":" & targetSheet.Range("BH" & endLine).Address))
  
  'ヘッダー情報(Copyright)-------------------------------------------------------------------------
  strHeader = "/* ----------------------------------------------------------------------------------------------------------------------------" & vbNewLine & _
              "TABLE NAME ：" & targetSheet.Range("F8") & " [" & tableName & "]" & vbNewLine & _
              "CREATE BY  ：" & targetSheet.Range("U2") & vbNewLine & _
              "CREATE DATA：" & Format(Now(), "yyyy/mm/dd hh:nn:ss") & vbNewLine & _
              "" & vbNewLine & _
               thisAppName & " [" & thisAppVersion & "]                                                          Copyright (c) 2021 B.Koizumi" & vbNewLine & _
              "---------------------------------------------------------------------------------------------------------------------------- */" & vbNewLine & vbNewLine
  
  
  'カラム情報--------------------------------------------------------------------------------------
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
    
    'カラム名
    queryColumn = queryColumn & Library.convFixedLength("" & targetSheet.Range("L" & line), colMaxLen + 4, " ")
    
    'データ型
    queryColumnTmp = targetSheet.Range("V" & line)
    
    'NULL制約
    If targetSheet.Range("AL" & line) <> "" Then
      queryColumn = queryColumn & "     " & "Not NULL"
    End If
    
    '初期値
    If targetSheet.Range("AB" & line) <> "" Then
      If targetSheet.Range("AB" & line) = "AUTO_INCREMENT" Then
        queryColumn = queryColumn & " " & targetSheet.Range("AB" & line)
      
      Else
        queryColumn = queryColumn & " DEFAULT " & targetSheet.Range("AB" & line)
      End If
    ElseIf targetSheet.Range("AL" & line) = "" Then
      queryColumn = queryColumn & " DEFAULT NULL"
    
    End If
    
    Call Library.showDebugForm("queryColumn", queryColumn)
    queryColumn = Library.convFixedLength(queryColumn, 100, " ")
    Call Library.showDebugForm("queryColumn", queryColumn)
    
    
    '備考
    If targetSheet.Range("AP" & line) <> "" Then
      queryColumn = queryColumn & " COMMENT '" & targetSheet.Range("B" & line) & vbTab
      queryColumn = queryColumn & Replace(targetSheet.Range("AP" & line), vbLf, "\n") & "'"
    Else
      queryColumn = queryColumn & " COMMENT '" & targetSheet.Range("B" & line) & "'"
    End If
    
    queryString = queryString & vbNewLine & queryColumn
    
  Next
  
  
  'インデックス情報------------------------------
  queryString = queryString & vbNewLine & vbNewLine & "-- インデックス情報------------------------------"
  
  Call Ctl_Common.chkRowStartLine
  idxLine = setLine("indexStart")
  idxEndLine = targetSheet.Range("B" & idxLine).End(xlDown).Row
  
  For line = idxLine To idxEndLine
    queryColumnTmp = targetSheet.Range("L" & line)
    queryColumnTmp = Replace(queryColumnTmp, ", ", ", ")
  
    If targetSheet.Range("C" & line) = "PK" Then
      queryString = queryString & vbNewLine & "  ,PRIMARY KEY (" & queryColumnTmp & ")"
    
    ElseIf targetSheet.Range("C" & line) = "#" Then
      Exit For
      
    ElseIf targetSheet.Range("C" & line) <> "" Then
      queryString = queryString & vbNewLine & "  ,        KEY " & targetSheet.Range("B" & line) & " (" & queryColumnTmp & ")"
    End If
  Next
  
  
  
  queryColumnTmp = targetSheet.Range("F8") & vbTab & targetSheet.Range("F11")
  queryString = queryString & vbNewLine & ")" & vbNewLine & "ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci COMMENT='" & queryColumnTmp & "'"
  
  'Copyright情報---------------------------------
'  strCopyright = vbNewLine & vbNewLine & vbNewLine & _
'              "/* -------------------------------------------------------------------------------" & vbNewLine & _
'              "" & vbNewLine & _
'              "" & vbNewLine & _
'               thisAppName & " [" & thisAppVersion & "]             Copyright (c) 2021 B.Koizumi" & vbNewLine & _
'              "------------------------------------------------------------------------------- */"
  
  
  If outpuFileFlg = True Then
    queryString = strHeader & queryString
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
  Set targetSheet = Nothing
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
  
  Const funcName As String = "Ctl_MySQL.setIndexInfo"
  
'  On Error GoTo catchError
  Call Library.startScript
  Call Library.showDebugForm("StartFun", funcName, "info")
  
  
  Call Ctl_Common.chkRowStartLine
  maxIndexNo = Application.WorksheetFunction.Max(Range(Cells(startLine, Target.Column), Cells(setLine("columnEnd"), Target.Column)))
  Call Library.showDebugForm("maxIndexNo", maxIndexNo, "info")
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
