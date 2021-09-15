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
'テーブル情報取得
Function getTableInfo()
  Dim line As Long, endLine As Long
  Dim TblRecordset As ADODB.Recordset
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  FuncName = "Ctl_MySQL.getTableInfo"
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



  '処理終了--------------------------------------
  Application.Goto Reference:=Range("A41"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function

