Attribute VB_Name = "Ctl_SQLServer"
Dim dbCon       As ADODB.Connection
Dim DBRecordset As ADODB.Recordset
Dim queryString As String


'**************************************************************************************************
' * SQLServer
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
  dbCon.ConnectionString = ConnectServer
  dbCon.Open
  
  isDBOpen = True
  Call Library.showDebugForm("isDBOpen：" & isDBOpen)
  
  Exit Function
  
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  If dbCon Is Nothing Then
  Else
    dbCon.Close
    Set DBRecordset = Nothing
  End If
  isOpen = False
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
    Call Library.showDebugForm("isDBOpen：" & isDBOpen)
  End If
  
  Exit Function

'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm("isDBOpen：" & isDBOpen)
  Call Library.showNotice(501, Err.Description, True)
End Function


'==================================================================================================
Function getTableInfo()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim tableName As String
  
  Dim TblRecordset As ADODB.Recordset

'  On Error GoTo catchError

  Call init.Setting
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  Call dbOpen
  
  runFlg = True

  queryString = "SELECT * FROM sys.objects where type='U' order by name"
  
  Set TblRecordset = New ADODB.Recordset
  TblRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly

  Do Until TblRecordset.EOF
    tableName = TblRecordset.Fields("name").Value
    
'    Call Ctl_ProgressBar.showCount("テーブル情報取得", TblRecordset.AbsolutePosition, TblRecordset.RecordCount, tableName)
    L_PrgP_Cnt = TblRecordset.AbsolutePosition
    PrgP_Max = TblRecordset.RecordCount
    Call Ctl_ProgressBar.showBar("情報取得", L_PrgP_Cnt, PrgP_Max, 1, 1, tableName)
    
    If Library.chkSheetExists(tableName) Then
      Sheets(tableName).Select
    Else
      CopySheet.Range("H5") = tableName
      Call Ctl_Sheet.addSheet
    End If
    
    Range("T2") = TblRecordset.Fields("create_date").Value
    Range("T3") = TblRecordset.Fields("modify_date").Value
    
    Call getColumnInfo
    TblRecordset.MoveNext
  Loop
  Set TblRecordset = Nothing
  
  Call dbClose
  
  runFlg = False
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  runFlg = False
  Call Library.showNotice(400, Err.Description, True)
End Function





'==================================================================================================
Function getColumnInfo()
  Dim line As Long, endLine As Long
  Dim columnName As String
  Dim columnType As String
  Dim columnDigit As String
  Dim columnDefValue As String
  Dim columnNotNull As Integer
  Dim columnComment As String
  Dim ClmRecordset As ADODB.Recordset
  
  Dim ColumnNames() As Variant
  Dim indexCount As Long
  

'  On Error GoTo catchError
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
    Call dbOpen
  End If
  
  '既存データクリア------------------------------
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  
  Rows("9:" & endLine).EntireRow.Hidden = False
  Range("A9:V" & endLine).ClearContents
  
  If Range("B2") = "" Then
    Range("B2") = "マスターテーブル"
  End If
  
  tableName = Range("H5")
  queryString = "SELECT * FROM INFORMATION_SCHEMA.Columns where TABLE_NAME='" & tableName & "' order by ORDINAL_POSITION"
  Call Library.showDebugForm(queryString)
  
  Set ClmRecordset = New ADODB.Recordset
  ClmRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly

  'テーブル情報
'  Range("H5") = ClmRecordset.Fields("TABLE_NAME").Value
  Range("T5") = ClmRecordset.Fields("TABLE_SCHEMA").Value
  
  
  'カラム情報--------------------------------------------------------------------------------------
  line = startLine
  Do Until ClmRecordset.EOF
    Call Ctl_ProgressBar.showBar("情報取得", L_PrgP_Cnt, PrgP_Max, ClmRecordset.AbsolutePosition, ClmRecordset.RecordCount, ClmRecordset.Fields("COLUMN_NAME").Value)
    If line >= 108 Then
      Rows(line & ":" & line).Select
      Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End If
    Range("B" & line) = line - 8
    Range("E" & line) = ClmRecordset.Fields("COLUMN_NAME").Value
    Range("G" & line) = ClmRecordset.Fields("DATA_TYPE").Value
    
    If Range("G" & line) <> "ntext" Then
      Range("H" & line) = ClmRecordset.Fields("CHARACTER_MAXIMUM_LENGTH").Value
    End If
    
    If ClmRecordset.Fields("IS_NULLABLE").Value = "NO" Then
      Range("U" & line) = 1
    End If
    Range("V" & line) = ClmRecordset.Fields("COLUMN_DEFAULT").Value
    Range("V" & line) = Replace(Range("V" & line), "((", "")
    Range("V" & line) = Replace(Range("V" & line), "))", "")

    ClmRecordset.MoveNext
    line = line + 1
  Loop
  Set ClmRecordset = Nothing
  Range("A" & line - 1) = "RowEnd"
  
  '行の非表示
  If line < 90 Then
    Rows(line & ":107").EntireRow.Hidden = True
  Else
    line = line + 4
  End If
  Call Library.showDebugForm("line：" & line)

  
  
  'プライマリーキー情報--------------------------------------------------------------------------------
  indexCount = 1
  Set ClmRecordset = New ADODB.Recordset
  queryString = "SELECT" & _
                " TBLS.NAME              AS TABLE_NAME" & _
                " , KEY_CONST.NAME       AS CONSTRAINT_NAME" & _
                " , KEY_CONST.TYPE_DESC  AS TYPE_DESC" & _
                " , IDX_COLS.KEY_ORDINAL AS KEY_ORDINAL" & _
                " , COLS.COLUMN_ID       AS COLUMN_ID" & _
                "  FROM" & _
                "    SYS.TABLES AS TBLS" & _
                "    INNER JOIN SYS.KEY_CONSTRAINTS AS KEY_CONST" & _
                "      ON TBLS.OBJECT_ID = KEY_CONST.PARENT_OBJECT_ID" & _
                "      AND KEY_CONST.TYPE = 'PK'  AND TBLS.NAME = '" & tableName & "'" & _
                "    INNER JOIN SYS.INDEX_COLUMNS AS IDX_COLS" & _
                "      ON KEY_CONST.PARENT_OBJECT_ID = IDX_COLS.OBJECT_ID" & _
                "      AND KEY_CONST.UNIQUE_INDEX_ID = IDX_COLS.INDEX_ID" & _
                "    INNER JOIN SYS.COLUMNS AS COLS" & _
                "      ON IDX_COLS.OBJECT_ID = COLS.OBJECT_ID" & _
                "      AND IDX_COLS.COLUMN_ID = COLS.COLUMN_ID"
  
  Call Library.showDebugForm(queryString)
  
  Set ClmRecordset = New ADODB.Recordset
  ClmRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly
  Do Until ClmRecordset.EOF
    Range("J" & ClmRecordset.Fields("COLUMN_ID").Value + 8) = indexCount
    
    Range("C" & line) = ClmRecordset.Fields("CONSTRAINT_NAME").Value
    Range(setVal("Cell_dateType") & line) = ClmRecordset.Fields("TYPE_DESC").Value
    
    indexCount = indexCount + 1
    ClmRecordset.MoveNext
  Loop
  Set ClmRecordset = Nothing
  
  
  'インデックス情報--------------------------------------------------------------------------------
  indexCount = 0
  oldID = 1
  Set ClmRecordset = New ADODB.Recordset
  queryString = "SELECT distinct" & _
                "    SYS.INDEXES.INDEX_ID  AS ID" & _
                "  , SYS.INDEX_COLUMNS.COLUMN_ID AS Index_ID" & _
                "  , SYS.INDEXES.NAME      AS INDEX_NAME" & _
                "  , SYS.INDEXES.TYPE_DESC AS INDEX_TYPE" & _
                "  , SYS.OBJECTS.NAME      AS TABLE_NAME" & _
                "  , SYS.COLUMNS.NAME      AS COLUMN_NAME" & _
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
  Call Library.showDebugForm(queryString)
  
  ColumnNames = Array("", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T")
  
  Set ClmRecordset = New ADODB.Recordset
  ClmRecordset.Open queryString, dbCon, adOpenKeyset, adLockReadOnly
  Do Until ClmRecordset.EOF
    If oldID = ClmRecordset.Fields("ID").Value Then
      indexCount = indexCount + 1
    Else
      indexCount = 1
      Range("C" & ClmRecordset.Fields("ID").Value + 110) = ClmRecordset.Fields("INDEX_NAME").Value
      
      Range(setVal("Cell_dateType") & ClmRecordset.Fields("ID").Value + 110) = ClmRecordset.Fields("INDEX_TYPE").Value
      
    End If
    If ClmRecordset.Fields("ID").Value <= 10 Then
      columnName = ColumnNames(ClmRecordset.Fields("ID").Value)
      Range(columnName & ClmRecordset.Fields("Index_ID").Value + 8) = indexCount
    End If
    
    
    line = line + 1
    oldID = ClmRecordset.Fields("ID").Value
    ClmRecordset.MoveNext
  Loop
  Set ClmRecordset = Nothing
  
  Application.Goto Reference:=Range("A1"), Scroll:=True
  If runFlg = False Then
    Call dbClose
    
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
  End If
  
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, Err.Description, True)
End Function
