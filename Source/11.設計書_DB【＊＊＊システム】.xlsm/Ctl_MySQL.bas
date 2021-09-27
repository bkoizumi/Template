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
Function dbOpen()
  On Error GoTo catchError
  
  If isDBOpen = True Then
    Call Library.showDebugForm("Database is already opened")
    Exit Function
  End If
  Call Library.showDebugForm("ConnectServer：" & ConnectServer)
  
  Set dbCon = New ADODB.Connection
  dbCon.Open ConnectServer
  dbCon.CursorLocation = 3
  
  isDBOpen = True
  Call Library.showDebugForm("isDBOpen：" & isDBOpen)
  
  Exit Function
  
'エラー発生時--------------------------------------------------------------------------------------
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
    Call Library.showDebugForm("isDBOpen：" & isDBOpen)
  End If
  
  Exit Function

'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm("isDBOpen：" & isDBOpen)
  Call Library.showNotice(501, Err.Description, True)
End Function


'==================================================================================================
Function IsTable(tableName As String) As Boolean
  Dim TblRecordset As ADODB.Recordset
  Dim rslFlg As Boolean

  rslFlg = False

  Call Library.showDebugForm("TableName：" & tableName)
  Call Ctl_MySQL.dbOpen
  
  'テーブル情報--------------------------------------------------------------------------------------
  queryString = " SELECT" & _
                "   TABLE_NAME as TableName, TABLE_COMMENT as Comments" & _
                " from" & _
                "   information_schema.TABLES" & _
                " WHERE" & _
                "   TABLE_SCHEMA = DATABASE()" & _
                "   and TABLE_NAME='" & tableName & "'"
      
  Call Library.showDebugForm("QueryString：" & queryString)
  
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

  On Error GoTo catchError
  
  Set runRecordset = dbCon.Execute(queryString)
  Set runRecordset = Nothing
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  If Err.Number = -2147217900 Then
    Call Library.showNotice(502, FuncName & " 構文エラー" & vbNewLine & queryString, True)
  Else
    Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "：" & Err.Description, True)
  End If
End Function


'==================================================================================================
'DB情報取得
Function getDatabaseInfo()
  Dim line As Long, endLine As Long
  Dim TblRecordset As ADODB.Recordset
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  FuncName = "Ctl_MySQL.getDatabaseInfo"
  Call Library.showDebugForm(FuncName & "==========================================")
  '----------------------------------------------

  'テーブル情報--------------------------------------------------------------------------------------
  queryString = " SELECT" & _
                "   TABLE_NAME as TableName, TABLE_COMMENT as Comments" & _
                " from" & _
                "   information_schema.TABLES" & _
                " WHERE" & _
                " TABLE_SCHEMA = DATABASE();"
      
  Call Library.showDebugForm("QueryString：" & queryString)
  
  Set TblRecordset = New ADODB.Recordset
  TblRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly
  PrgP_Max = TblRecordset.RecordCount
  
  Do Until TblRecordset.EOF
    Call Library.showDebugForm("TableName：" & TblRecordset.Fields("TableName").Value)
    
    If Library.chkSheetExists(TblRecordset.Fields("TableName").Value) = True Then
      Set targetSheet = ThisWorkbook.Worksheets(TblRecordset.Fields("TableName").Value)
    Else
      Call Ctl_Common.addSheet(TblRecordset.Fields("TableName").Value)
      Set targetSheet = ActiveSheet
    End If
    
    targetSheet.Select
    targetSheet.Range(setVal("Cell_TableType")) = "マスターテーブル"
    targetSheet.Range(setVal("Cell_physicalTableName")) = TblRecordset.Fields("TableName")
    
    If TblRecordset.Fields("Comments").Value Like "*<|>*" Then
      targetSheet.Range(setVal("Cell_logicalTableName")) = Split(TblRecordset.Fields("Comments").Value, "<|>")(0)
      targetSheet.Range(setVal("Cell_tableNote")) = Replace(Split(TblRecordset.Fields("Comments").Value, "<|>")(1), "<BR>", vbNewLine)
    Else
      targetSheet.Range(setVal("Cell_tableNote")) = TblRecordset.Fields("Comments")
    End If
    
    PrgP_Cnt = TblRecordset.AbsolutePosition
    Call Ctl_ProgressBar.showBar(thisAppName, TblRecordset.AbsolutePosition, TblRecordset.RecordCount, 1, 2, "カラム情報取得")
    TblRecordset.MoveNext
    
    'カラム情報取得
    Call Ctl_MySQL.getColumnInfo
    
  Loop
  Set TblRecordset = Nothing
  Set targetSheet = Nothing


  '処理終了--------------------------------------
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  '----------------------------------------------
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function


'==================================================================================================
'テーブル情報取得
Function getTableInfo()
  Dim line As Long, endLine As Long
  Dim TblRecordset As ADODB.Recordset
  Dim tableName As String
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  FuncName = "Ctl_MySQL.getTableInfo"
  Call Library.showDebugForm(FuncName & "==========================================")
  '----------------------------------------------
  tableName = ActiveSheet.Range(setVal("Cell_physicalTableName"))
  Call Library.showDebugForm("TableName：" & tableName)

  'テーブル情報--------------------------------------------------------------------------------------
  queryString = " SELECT" & _
                "   TABLE_NAME as TableName, TABLE_COMMENT as Comments" & _
                " from" & _
                "   information_schema.TABLES" & _
                " WHERE" & _
                "   TABLE_SCHEMA = DATABASE()" & _
                "   and TABLE_NAME='" & tableName & "'"
      
  Call Library.showDebugForm("QueryString：" & queryString)
  
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
  targetSheet.Range(setVal("Cell_TableType")) = "マスターテーブル"
  targetSheet.Range(setVal("Cell_physicalTableName")) = TblRecordset.Fields("TableName")
  
  If TblRecordset.Fields("Comments").Value Like "*<|>*" Then
    targetSheet.Range(setVal("Cell_logicalTableName")) = Split(TblRecordset.Fields("Comments").Value, "<|>")(0)
    targetSheet.Range(setVal("Cell_tableNote")) = Replace(Split(TblRecordset.Fields("Comments").Value, "<|>")(1), "<BR>", vbNewLine)
  Else
    targetSheet.Range(setVal("Cell_tableNote")) = TblRecordset.Fields("Comments")
  End If
  
  PrgP_Cnt = TblRecordset.AbsolutePosition
  Call Ctl_ProgressBar.showBar(thisAppName, TblRecordset.AbsolutePosition, TblRecordset.RecordCount, 1, 2, "カラム情報取得")
  
  'カラム情報取得
  Call Ctl_MySQL.getColumnInfo
    
  Set TblRecordset = Nothing
  Set targetSheet = Nothing


  '処理終了--------------------------------------
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  '----------------------------------------------
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function


'==================================================================================================
'カラム情報取得
Function getColumnInfo()
  Dim line As Long, endLine As Long
  Dim ClmRecordset As ADODB.Recordset
  Dim tableName As String
  Dim LFCount As Long, IndexLine As Long
  Dim ColCnt As Long
  Dim searchColCell As Range
  Dim IndexName As String, oldIndexName As String, IndexColName As String
  '処理開始--------------------------------------
  'On Error GoTo catchError
  FuncName = "Ctl_MySQL.getColumnInfo"
  If PrgP_Max = 0 Then
    PrgP_Max = 2
    PrgP_Cnt = 1
  End If
  Call Library.showDebugForm(FuncName & "=========================================")
  Call Library.showDebugForm("runFlg：" & runFlg)
  Call Ctl_Common.ClearData
  '----------------------------------------------
  If targetSheet Is Nothing Then
    Set targetSheet = ActiveSheet
  End If
  targetSheet.Select
  
  tableName = targetSheet.Range(setVal("Cell_physicalTableName"))
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
                " FROM" & _
                "   information_schema.Columns c " & _
                " WHERE" & _
                "   c.table_schema = '" & setVal("DBName") & "' " & _
                "   AND c.table_name   = '" & tableName & "' " & _
                " ORDER BY" & _
                "   ordinal_position;"
      
  Call Library.showDebugForm("QueryString：" & queryString)
  
  Set ClmRecordset = New ADODB.Recordset
  ClmRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly
  
  
  Do Until ClmRecordset.EOF
    targetSheet.Range("C" & line) = ColCnt
    If ClmRecordset.Fields("Comments").Value Like "*<|>*" Then
      targetSheet.Range(setVal("Cell_logicalName") & line) = Split(ClmRecordset.Fields("Comments").Value, "<|>")(0)
      targetSheet.Range(setVal("Cell_Note") & line) = Replace(Split(ClmRecordset.Fields("Comments").Value, "<|>")(1), "<BR>", vbNewLine)
    Else
      targetSheet.Range(setVal("Cell_logicalName") & line) = ClmRecordset.Fields("Comments").Value
    End If
    targetSheet.Range(setVal("Cell_physicalName") & line) = ClmRecordset.Fields("ColumName").Value
    targetSheet.Range(setVal("Cell_dateType") & line) = ClmRecordset.Fields("DataType").Value
    targetSheet.Range(setVal("Cell_digits") & line) = ClmRecordset.Fields("Length").Value
    If ClmRecordset.Fields("PrimaryKey").Value = "PRI" Then
      targetSheet.Range(setVal("Cell_PK") & line) = 1
    End If
    If ClmRecordset.Fields("Nullable").Value = "NO" Then
      targetSheet.Range(setVal("Cell_Null") & line) = 1
    End If
    targetSheet.Range(setVal("Cell_Default") & line) = ClmRecordset.Fields("ColumnDefault").Value
    
    
    '行の高さ調整
    LFCount = UBound(Split(targetSheet.Range(setVal("Cell_Note") & line).Value, vbNewLine)) + 1
    If LFCount > 0 Then
      targetSheet.Rows(line & ":" & line).RowHeight = 18 * LFCount
    End If
    
    
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, ClmRecordset.AbsolutePosition, ClmRecordset.RecordCount, "カラム情報取得")
    ClmRecordset.MoveNext
    
    line = line + 1
    ColCnt = ColCnt + 1
    Call Ctl_Common.addRow(line)
  Loop
  Set ClmRecordset = Nothing
  
  'インデックス情報取得----------------------------------------------------------------------------
  If PrgP_Max = 0 Then
    PrgP_Cnt = 2
  End If
  
  IndexLine = Ctl_Common.chkIndexRow
  queryString = "SHOW INDEX FROM " & tableName & ";"
  Call Library.showDebugForm("QueryString：" & queryString)
  
  Set ClmRecordset = New ADODB.Recordset
  ClmRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly
  
  line = IndexLine
  ColCnt = -1
  Do Until ClmRecordset.EOF
    IndexName = ClmRecordset.Fields("Key_name").Value
    If oldIndexName <> IndexName Then
      line = line + 1
      ColCnt = ColCnt + 1
      Call Ctl_Common.addRow(line + 1)
    End If
    
    If IndexName = "PRIMARY" Then
      targetSheet.Range("C" & line) = "PK"
    Else
      targetSheet.Range("C" & line) = ColCnt
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
    
    Set searchColCell = Columns("E:E").Find(What:=ClmRecordset.Fields("Column_name").Value)
    If ColCnt <= 10 Then
      Cells(searchColCell.Row, 9 + ColCnt) = ClmRecordset.Fields("Seq_in_index").Value
      
      IndexColName = Library.getColumnName(9 + ColCnt)
      Columns(IndexColName & ":" & IndexColName).EntireColumn.Hidden = False
    End If
    
    Call Ctl_ProgressBar.showBar(thisAppName, PrgP_Cnt, PrgP_Max, ClmRecordset.AbsolutePosition, ClmRecordset.RecordCount, "インデックス情報取得")
    Set searchColCell = Nothing
    oldIndexName = IndexName


    ClmRecordset.MoveNext
  Loop

  If Range("B5") = "" Then
    Range("B5") = "exist"
  End If
  
  '処理終了--------------------------------------
  Application.Goto Reference:=Range("A41"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "：" & Err.Description, True)
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

  FuncName = "Ctl_MySQL.CreateTable"
  PrgP_Max = 4
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
    Call Ctl_MySQL.dbOpen
  End If
  queryString = ""
  Call Library.showDebugForm(FuncName & "============================================")
  '----------------------------------------------
  endLine = Cells(Rows.count, 3).End(xlUp).Row
  
  tableName = Range(setVal("Cell_physicalTableName"))
  
  If Range("B5") = "" Then
  
  ElseIf Range("B5") = "exist" Then
    '既存テーブルの変更処理------------------------
    Call Library.showDebugForm("既存テーブルの変更", tableName)
    
    For line = startLine To endLine
      If Range("B" & line) = "edit" Then
        'データ型変更------------------------------
        queryString = "ALTER TABLE " & Range(setVal("Cell_physicalTableName")) & " MODIFY COLUMN " & Range(setVal("Cell_physicalName") & line) & " " & Range(setVal("Cell_dateType") & line)
        If Range(setVal("Cell_digits") & line) <> "" Then
          queryString = queryString & " (" & Range(setVal("Cell_digits") & line) & ")"
        End If
        
        'NotNull制約-----------------------------
        If Range(setVal("Cell_Null") & line) = 1 Then
          queryString = queryString & " NOT NULL"
        End If
        
        '初期値設定------------------------------
        If Range(setVal("Cell_Default") & line) <> "" Then
          queryString = queryString & ""
        End If
        
        'コメント--------------------------------
        If Range(setVal("Cell_Note") & line) <> "" Then
          queryString = queryString & " Comment '" & Range(setVal("Cell_logicalName") & line) & "<|>" & _
              Replace(Range(setVal("Cell_Note") & line), vbNewLine, "<BR>") & "'"
        Else
          queryString = queryString & " Comment '" & Range(setVal("Cell_logicalName") & line) & "'"

        End If
      
      'カラム名変更[追加⇒削除]------------------
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
  
  
  '処理終了--------------------------------------
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  If runFlg = False Then
    Call Ctl_MySQL.dbClose
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.usetting
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function



'==================================================================================================
'DDL生成
Function makeDDL()
  Dim line As Long, endLine As Long
  Dim idxLine As Long, idxEndLine As Long
  Dim tableName As String
  Dim queryColumn As String, queryColumnTmp As String
  Dim colMaxLen As Long
  Dim Copyright As String
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  FuncName = "Ctl_MySQL.getColumnInfo"
  PrgP_Max = 4
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
  End If
  Call Library.showDebugForm(FuncName & "=========================================")
  Call Library.showDebugForm("runFlg", runFlg)
  '----------------------------------------------
  If targetSheet Is Nothing Then
    Set targetSheet = ActiveSheet
  End If
  targetSheet.Select
  
  tableName = targetSheet.Range(setVal("Cell_physicalTableName"))
  endLine = Cells(Rows.count, 3).End(xlUp).Row
  endLine = targetSheet.Range(setVal("Cell_physicalName") & startLine).End(xlDown).Row
  
  'カラム名の最大文字数を取得
  colMaxLen = WorksheetFunction.Max(targetSheet.Range(targetSheet.Range(setVal("Cell_colMaxLen") & startLine).Address & ":" & targetSheet.Range(setVal("Cell_colMaxLen") & endLine).Address))
  
  'ヘッダー情報(Copyright)-------------------------------------------------------------------------
  Copyright = "/* -------------------------------------------------------------------------------" & vbNewLine & _
              "Tabel Name ：" & targetSheet.Range(setVal("Cell_logicalTableName")) & " [" & tableName & "]" & vbNewLine & _
              "" & vbNewLine & _
              "CREATE Date：" & Format(Now(), "yyyy/mm/dd hh:nn:ss") & vbNewLine & _
               thisAppName & " [" & thisAppVersion & "]             Copyright (c) 2021 B.Koizumi" & vbNewLine & _
              "------------------------------------------------------------------------------- */" & vbNewLine & vbNewLine & vbNewLine
  
  
  'カラム情報--------------------------------------------------------------------------------------
  queryString = " CREATE TABLE `" & tableName & "` ("
  For line = startLine To endLine
    If line = startLine Then
      queryColumn = "  "
    Else
      queryColumn = "  ,"
    End If
    
    'カラム名
    queryColumn = queryColumn & Library.convFixedLength("`" & targetSheet.Range(setVal("Cell_physicalName") & line) & "`", colMaxLen + 4, " ")
    
    'データ型
    queryColumnTmp = targetSheet.Range(setVal("Cell_dateType") & line)
    
    '桁数
    If targetSheet.Range(setVal("Cell_digits") & line) <> "" Then
      queryColumnTmp = queryColumnTmp & "(" & targetSheet.Range(setVal("Cell_digits") & line).Value & ")"
    End If
    queryColumn = queryColumn & Library.convFixedLength(queryColumnTmp, 20, " ") & ""
    
    'NULL制約
    If targetSheet.Range(setVal("Cell_Null") & line) <> "" Then
      queryColumn = queryColumn & "     " & targetSheet.Range(setVal("Cell_Null") & line).Text & ""
    Else
      queryColumn = queryColumn & " DEFAULT NULL"
    End If
    
    If targetSheet.Range(setVal("Cell_Note") & line) <> "" Then
      queryColumn = queryColumn & " COMMENT '" & targetSheet.Range(setVal("Cell_logicalName") & line) & "<|>"
      queryColumn = queryColumn & Replace(targetSheet.Range(setVal("Cell_Note") & line), vbNewLine, "<BR>") & "'"
    Else
      queryColumn = queryColumn & " COMMENT '" & targetSheet.Range(setVal("Cell_logicalName") & line) & "'"
    End If
    
    queryString = queryString & vbNewLine & queryColumn
    
  Next
  
  
  'インデックス情報------------------------------
  queryString = queryString & vbNewLine & vbNewLine & "-- インデックス情報------------------------------"
  
  idxLine = Ctl_Common.chkIndexRow + 1
  idxEndLine = targetSheet.Range(setVal("Cell_logicalName") & idxLine).End(xlDown).Row
  
  For line = idxLine To idxEndLine
    queryColumnTmp = targetSheet.Range(setVal("Cell_digits") & line)
    queryColumnTmp = Replace(queryColumnTmp, ", ", "`, `")
  
    If targetSheet.Range("C" & line) = "PK" Then
      queryString = queryString & vbNewLine & "  ,PRIMARY KEY (`" & queryColumnTmp & "`)"
    Else
      queryString = queryString & vbNewLine & "  ,        KEY `" & targetSheet.Range(setVal("Cell_logicalName") & line) & "` (`" & queryColumnTmp & "`)"
    End If
  Next
  
  
  
  queryColumnTmp = targetSheet.Range(setVal("Cell_logicalTableName")) & "<|>" & targetSheet.Range(setVal("Cell_tableNote"))
  queryString = queryString & vbNewLine & ") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci COMMENT='" & queryColumnTmp & "'"
  
  queryString = Copyright & queryString
  Call Library.outputText(queryString, ThisWorkbook.Path & "\CREATE_TABLE_" & targetSheet.Range(setVal("Cell_physicalTableName")) & ".sql")
  
  
  Debug.Print queryString
  'Call Library.showDebugForm("QueryString：" & queryString)
  
  
  '処理終了--------------------------------------
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.usetting
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function

