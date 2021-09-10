Attribute VB_Name = "Aold_MakeDML"
'Public SaveDirPath As String
'Public ODBCDriver As String
'Public DBMS As String
'Public DBServer As String
'Public DBTableSpace As String
'Public DBName As String
'Public DBInstance As String
'Public DBScheme As String
'Public DBPort As String
'
'Public LoginID As String
'Public LoginPW As String
'Public FlameWorkName As String
'Public SetDisplyAlertFlg As Boolean
'Public SetDisplyProgressBarFlg As Boolean
'Public SetSelectTargetRows As String
'
''Public InputTableName As String
'Public InputTableID As String
'Public BeforeCloseFlg As Boolean
'
'Public DebugFlg As Boolean
'Public ConnectionString As String
'Public LineBreakCode As String
'Public LineSeparator As Integer
'Public CharacterSet As String
'Public DBMode As String
'
''***********************************************************************************************************************************************
'' * DB情報取得
'' *
'' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
''***********************************************************************************************************************************************
'Function MakeDML_Init()
'
'  DebugFlg = True
'
'  ODBCDriver = Worksheets("設定").Range("B3").Value
'  DBMS = Application.WorksheetFunction.VLookup(ODBCDriver, Range("ODBCDriverList"), 2, False)
'
'  SetDisplyAlertFlg = True
'  BeforeCloseFlg = False
'
'
'  DBPort = Worksheets("設定").Range("B9").Value
'
'  LoginID = Worksheets("設定").Range("B10").Value
'  LoginPW = Worksheets("設定").Range("B11").Value
'
'  Select Case DBMS
'    Case "PostgreSQL"
'      DBServer = Worksheets("設定").Range("B4").Value
'      DBName = Worksheets("設定").Range("B5").Value
'      DBInstance = Worksheets("設定").Range("B6").Value
'      DBScheme = Worksheets("設定").Range("B7").Value
'
'      ConnectionString = ""
'
'    Case "MySQL"
'      DBServer = Worksheets("設定").Range("B4").Value
'      DBName = Worksheets("設定").Range("B5").Value
'      DBInstance = Worksheets("設定").Range("B6").Value
'      DBScheme = Worksheets("設定").Range("B7").Value
'
'      ConnectionString = "Driver={" & ODBCDriver & "}; Server=" & DBServer & ";Port=" & _
'                          DBPort & ";Database=" & DBName & ";UID=" & LoginID & ";PWD=" & LoginPW & ""
'
'    Case "Oracle"
'      DBServer = Worksheets("設定").Range("B4").Value
'      DBName = Worksheets("設定").Range("B5").Value
'
'      DBTableSpace = Worksheets("設定").Range("B6").Value
'      DBInstance = Worksheets("設定").Range("B7").Value
'      DBScheme = Worksheets("設定").Range("B8").Value
'
'
'      ConnectionString = "Driver={" & ODBCDriver & "};DBQ=" & DBName & ";UID=" & LoginID & ";PWD=" & LoginPW & ""
'
'    Case "SQLServer"
'      ConnectionString = "Provider=SQLOLEDB.1;Data Source=" & DBServer & ";Initial Catalog=" & _
'                          DBName & ";User ID=" & LoginID & ";Password=" & LoginPW & ";"
'
'  End Select
'
'
'  SaveDirPath = Worksheets("設定").Range("B12")
'  DBMode = Worksheets("設定").Range("B13")
'
'  CharacterSet = Worksheets("設定").Range("B13").Value
'
'  Select Case Worksheets("設定").Range("B14").Value
'  Case "CRLF"
'    LineSeparator = -1
'    LineBreakCode = vbCrLf
'  Case "LF"
'    LineSeparator = 10
'    LineBreakCode = vbLf
'  Case "CR"
'    LineSeparator = 13
'    LineBreakCode = vbCr
'  End Select
'
''  LineBreakCode = vbLf
''  CharacterSet = "UTF-8"
'
'
'End Function
'
'
'
''***********************************************************************************************************************************************
'' * SQL生成用シート追加
'' *
'' * Bunpei.Koizumi<bunbun0716@gmail.com>
''***********************************************************************************************************************************************
'Function MakeDML_AddSheet()
'
'  With MakeDMLForm
'    .StartUpPosition = 0
'      .Top = Application.Top + (ActiveWindow.Width / 4)
'      .Left = Application.Left + (ActiveWindow.Height / 2)
'  End With
'  MakeDMLForm.Show
'
'  Call MakeDML_Init
'
'
'End Function
'
''***********************************************************************************************************************************************
'' * カラム情報再取得
'' *
'' * Bunpei.Koizumi<bunbun0716@gmail.com>
''***********************************************************************************************************************************************
'Function MakeDML_reGetColumn()
'
'  Dim QueryString As String
'  Dim SelectString As String
'  Dim RunQueryString As String
'
'  Dim tableName As String
'  Dim NowLine As Long
'  Dim DelLine As Long
'
'  Dim dbCon As ADODB.Connection
'  Dim DBRecordset As ADODB.Recordset
'
'  Dim columnName As String
'  Dim dataType As String
'  Dim maxLength As String
'  Dim PrimaryKeyIndex As String
'  Dim RefTableName As String
'  Dim RefColumnName As String
'  Dim isNullable As String
'  Dim is_identity As String
'  Dim EPvalue As String
'  Dim Comments As String
'  Dim RowCount As Long
'
''  On Error GoTo GetColumnList_Error
'
'  Call MakeDML_Init
'
'  tableName = Range("B1")
'  NowLine = 1
'
'  ' プログレスバーの表示開始
'  ProgressBar_ProgShowStart
'
'
'  'ADODB.Connection生成し、DBに接続
'  ProgressBar_ProgShowCount "処理中", 5, 100, "DBに接続"
'
'  Set dbCon = New ADODB.Connection
'  dbCon.Open ConnectionString
'  ProgressBar_ProgShowCount "処理中", 10, 100, "DBに接続"
'
'  'SQL文 ----------------------------------------------------------
'  SelectString = ""
'  QueryString = ""
'  Select Case DBMS
'    Case "PostgreSQL"
'
'
'    Case "MySQL"
'      SelectString = "select "
'      SelectString = SelectString & "COLUMN_NAME as ColumName,DATA_TYPE as DataType,COLUMN_KEY as PrimaryKey,ifnull(CHARACTER_MAXIMUM_LENGTH, '') as Length,"
'      SelectString = SelectString & "IS_NULLABLE as Nullable,COLUMN_DEFAULT as ColumnDefault,COLUMN_COMMENT as Comments,c.extra "
'      QueryString = QueryString & "from"
'      QueryString = QueryString & " information_schema.Columns c "
'      QueryString = QueryString & "where"
'      QueryString = QueryString & " c.table_schema = '" & DBName & "' "
'      QueryString = QueryString & " and"
'      QueryString = QueryString & " c.table_name   = '" & tableName & "' "
'      QueryString = QueryString & "Order by"
'      QueryString = QueryString & " ordinal_position;"
'
'    Case "Oracle"
'      SelectString = "select " & vbCrLf
'      SelectString = SelectString & "    UTC.column_name                                as ColumName," & vbCrLf
'      SelectString = SelectString & "    UTC.data_type                                  as DataType," & vbCrLf
'      SelectString = SelectString & "    NVL(UTC.DATA_PRECISION, CHAR_COL_DECL_LENGTH)  as Length," & vbCrLf
'      SelectString = SelectString & "    UCCPkey.position                               as PrimaryKey," & vbCrLf
'      SelectString = SelectString & "    case" & vbCrLf
'      SelectString = SelectString & "      when UTC.nullable ='Y' then 0" & vbCrLf
'      SelectString = SelectString & "      when UTC.nullable ='N' then 1" & vbCrLf
'      SelectString = SelectString & "    end                                            as Nullable," & vbCrLf
'      SelectString = SelectString & "    UCC.COMMENTS                                   as Comments," & vbCrLf
'      SelectString = SelectString & "    ''                                             as ColumnDefault" & vbCrLf
'      QueryString = QueryString & "  from" & vbCrLf
'      QueryString = QueryString & "    USER_TAB_COLUMNS UTC left join USER_COL_COMMENTS UCC on UTC.table_name = UCC.table_name and UTC.column_name = UCC.column_name" & vbCrLf
'      QueryString = QueryString & "    left join USER_CONS_COLUMNS UCCPkey on UTC.table_name = UCCPkey.table_name and UTC.column_name = UCCPkey.column_name and UCCPkey.position is not null" & vbCrLf
'      QueryString = QueryString & "  where UTC.table_name='" & tableName & "'" & vbCrLf
'      QueryString = QueryString & "  order by UTC.column_id" & vbCrLf
'
'    Case "SQLServer"
'      QueryString = "select table_name TableName,'' Comments from USER_TABLES;"
'  End Select
'
'  ProgressBar_ProgShowCount "処理中", 15, 100, "DBに接続"
'
'  Set DBRecordset = New ADODB.Recordset
'
'  'カラム数取得
'  RunQueryString = "select count(*) as count " & QueryString
'  Set DBRecordset = New ADODB.Recordset
'  DBRecordset.Open RunQueryString, dbCon, adOpenKeyset, adLockReadOnly
'  Do Until DBRecordset.EOF
'    RowCount = CLng(DBRecordset.Fields("count").Value)
'
'    '次のレコード
'    DBRecordsetCount = DBRecordsetCount + 1
'    DBRecordset.MoveNext
'  Loop
'  Set DBRecordset = Nothing
'  ProgressBar_ProgShowCount "処理中", 50, 100, "カラム数取得"
'
'  'カラム情報取得
'  RunQueryString = SelectString & QueryString
'
'  Set DBRecordset = New ADODB.Recordset
'  DBRecordset.Open RunQueryString, dbCon, adOpenKeyset, adLockReadOnly
'  ProgressBar_ProgShowCount "処理中", 100, 100, "カラム情報取得"
'
'  Do Until DBRecordset.EOF
'    ' プログレスバーのカウント変更（現在のカウント、全カウント数、メッセージ）
'    ProgressBar_ProgShowCount "処理中", NowLine, RowCount, "カラム情報取得：" & columnName
'
'    'コメント(項目名として利用)
'    If IsNull(DBRecordset.Fields("Comments").Value) Then
'      Comments = ""
'    Else
'      Comments = DBRecordset.Fields("Comments").Value
'    End If
'
'    'カラム名
'    columnName = DBRecordset.Fields("ColumName").Value
'
'    '型
'    dataType = DBRecordset.Fields("DataType").Value
'    Select Case dataType
'      Case "TIMESTAMP(6)"
'        dataType = "TIMESTAMP"
'    End Select
'
'    '桁数
'    If IsNull(DBRecordset.Fields("Length").Value) Then
'      maxLength = ""
'    Else
'      maxLength = DBRecordset.Fields("Length").Value
'    End If
'    Select Case dataType
'      Case "numeric", "decimal"
'        maxLength = Precision & ", " & ScaleString
'
'      Case "datetime2", "datetime", "tinyint", "bit", "varbinary", "xml", "image", "money", "text", "TIMESTAMP"
'        maxLength = ""
'    End Select
'
'    'プライマリキー
'    If IsNull(DBRecordset.Fields("PrimaryKey").Value) Then
'      PrimaryKeyIndex = False
'    Else
'      PrimaryKeyIndex = DBRecordset.Fields("PrimaryKey").Value
'      If PrimaryKeyIndex = "PRI" Then
'        PrimaryKeyIndex = True
'      ElseIf PrimaryKeyIndex = "UNI" Then
'        PrimaryKeyIndex = True
'      ElseIf PrimaryKeyIndex = "MUL" Then
'        PrimaryKeyIndex = True
'      End If
'    End If
'
'    'NotNULL制約
'    If IsNull(DBRecordset.Fields("Nullable").Value) Then
'      isNullable = False
'    Else
'      isNullable = DBRecordset.Fields("Nullable").Value
'    End If
'
'    If isNullable = "1" Then
'      isNullable = True
'    Else
'      isNullable = False
'    End If
'
'
'    '初期値
'    If IsNull(DBRecordset.Fields("ColumnDefault").Value) Then
'      ColumnDefault = ""
'    Else
'      'Range("P" & NowLine).NumberFormatLocal = "@"
'      ColumnDefault = DBRecordset.Fields("ColumnDefault").Value
'      ColumnDefault = Replace(ColumnDefault, "(", "")
'      ColumnDefault = Replace(ColumnDefault, ")", "")
'      ColumnDefault = Replace(ColumnDefault, "'", "")
'      ColumnDefault = Replace(ColumnDefault, "'", "")
'
'      Select Case ColumnDefault
'        Case "getdate"
'          ColumnDefault = "getdate()"
'
'      End Select
'    End If
'
'    'カラム設定
'    Select Case Range("A1")
'      Case "INSERT"
'        Sheets("Insert文").Range("B3:B105").Copy
'
'        Cells(3, NowLine).Select
'        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
'          SkipBlanks:=False, Transpose:=False
'        'ActiveSheet.Paste
'        'Range(Cells(3, NowLine), Cells(5, NowLine)).Select
'        'Selection.ClearContents
'
'        Cells(3, NowLine) = dataType
'        Cells(4, NowLine) = columnName
'        Cells(5, NowLine) = Comments
'
'
'        If Not IsNull(PrimaryKeyIndex) Then
'          Cells(4, NowLine).Font.Bold = True
'        Else
'          Cells(4, NowLine).Font.Bold = False
'        End If
'
'        If isNullable Then
'          Cells(4, NowLine).Font.Color = RGB(255, 0, 0)
'        Else
'          Cells(4, NowLine).Font.ColorIndex = xlAutomatic
'        End If
'
'
'      Case "UPDATE"
'        Sheets("Update文").Range("B3:B107").Copy
'
'        Cells(3, NowLine).Select
'        ActiveSheet.Paste
'        Range(Cells(3, NowLine), Cells(5, NowLine)).Select
'        Selection.ClearContents
'
'        Cells(3, NowLine) = dataType
'        Cells(6, NowLine) = columnName
'        Cells(7, NowLine) = Comments
'
'
'        If Not IsNull(PrimaryKeyIndex) Then
'          Cells(6, NowLine).Font.Bold = True
'        Else
'          Cells(6, NowLine).Font.Bold = False
'        End If
'
'        If isNullable Then
'          Cells(6, NowLine).Font.Color = RGB(255, 0, 0)
'        Else
'          Cells(6, NowLine).Font.ColorIndex = xlAutomatic
'        End If
'
'      Case "DELETE"
'        Sheets("Delete文").Range("B3:B104").Copy
'
'        Cells(3, NowLine).Select
'        ActiveSheet.Paste
'        Range(Cells(3, NowLine), Cells(5, NowLine)).Select
'        Selection.ClearContents
'
'        Cells(3, NowLine) = dataType
'        Cells(5, NowLine) = columnName
'        Cells(6, NowLine) = Comments
'
'
'        If PrimaryKeyIndex Then
'          Cells(6, NowLine).Font.Bold = True
'        Else
'          Cells(6, NowLine).Font.Bold = False
'        End If
'
'        If isNullable Then
'          Cells(6, NowLine).Font.Color = RGB(255, 0, 0)
'        Else
'          Cells(6, NowLine).Font.ColorIndex = xlAutomatic
'        End If
'
'      End Select
'
'    '次のレコード
'    NowLine = NowLine + 1
'    DBRecordset.MoveNext
'
'  Loop
'
'  Cells.Select
'  ProgressBar_ProgShowClose
'  Selection.ColumnWidth = 20
'  Range("A1").Select
'
'End Function
'
''***********************************************************************************************************************************************
'' * SQL生成用SQL生成
'' *
'' * Bunpei.Koizumi<bunbun0716@gmail.com>
''***********************************************************************************************************************************************
'Sub MakeDML_MakeSQL()
'
'  Call MakeDML_Init
'
'  If Range("A1") = "INSERT" Then
'    MakeDML_Insert ("One")
'  ElseIf Range("A1") = "UPDATE" Then
'    MakeDML_Update ("One")
'  ElseIf Range("A1") = "DELETE" Then
'    MakeDML_Delete ("One")
'  End If
'
'  Library_EndScript
'
'End Sub
'
''***********************************************************************************************************************************************
'' * 全SQL生成用SQL生成
'' *
'' * Bunpei.Koizumi<bunbun0716@gmail.com>
''***********************************************************************************************************************************************
'Sub MakeDML_MakeAllSheetSQL()
'
'  Dim tempSheet As Object
'  Dim tempSheetName As String
'  Call MakeDML_Init
'
'  For Each tempSheet In Sheets
'    tempSheetName = tempSheet.Name
'    If Library_CheckExcludeSheet(tempSheetName, 5) Then
'
'      Sheets(tempSheetName).Select
'      Range("A1").Select
'
'      If Range("A1") = "INSERT" Then
'        MakeDML_Insert ("ALL")
'      ElseIf Range("A1") = "UPDATE" Then
'        MakeDML_Update ("ALL")
'      ElseIf Range("A1") = "DELETE" Then
'        MakeDML_Delete ("ALL")
'      End If
'    End If
'  Next
'
'  Shell "Explorer.exe " & SaveDirPath, vbNormalFocus
'End Sub
'
''***********************************************************************************************************************************************
'' * Insert文作成
'' *
'' * Bunpei.Koizumi<bunbun0716@gmail.com>
''***********************************************************************************************************************************************
'Sub MakeDML_Insert(ExecType As String)
'
'  Dim Query As String
'  Dim line As Long
'  Dim colLine As Long
'  Dim endLine As Long
'  Dim endColLine As Long
'  Dim ObjFileSys As Object
'  Dim ObjFile As Object
'
'  Dim DirName As String
'  Dim Today As String
'  Dim SQLTable As String
'  Dim SheetValue As String
'
'  Dim QueryColumn As String
'  Dim QueryValues As String
'  Dim QuartString As String
'  Dim FileCnt As Integer
'
'  Dim ObjADODB As Object
'  Dim FoundCell As Range
'
'  On Error GoTo Error
'
'  MakeDML_Init
'
'  ' ファイルの保存ディレクトリの指定
'  If (Worksheets("設定").Range("B12") = "") Then
'    SaveDirPath = Library_GetDirPath(ActiveWorkbook.Path)
'    If (SaveDirPath = "") Then
'      Exit Sub
'    End If
'  End If
'  Worksheets("設定").Range("B12") = SaveDirPath
'
'  SQLTable = Range("B1")
'  Today = Format(Now, "yyyymmdd")
'  FileCnt = 0
'
'
'  ' プログレスバーの表示開始
'  ProgressBar_ProgShowStart
'
'
'  Set ObjADODB = CreateObject("ADODB.Stream")
'
'  'オブジェクトに保存するデータの種類を文字列型に指定する
'  ObjADODB.Type = 2
'
'  '文字列型のオブジェクトの文字コードを指定する(改行コード：CR：13　CRLF：-1  LF：10)
'  ObjADODB.Charset = CharacterSet
'  ObjADODB.LineSeparator = LineSeparator
'
'  'オブジェクトのインスタンスを作成
'  ObjADODB.Open
'
'
'  endLine = Cells(Rows.count, 1).End(xlUp).Row
'  endColLine = Cells(4, Columns.count).End(xlToLeft).Column
'
'  ' カラム名取得
'  For colLine = 1 To endColLine
'    If colLine = endColLine Then
'      QueryColumn = QueryColumn & Cells(4, colLine).Value
'    Else
'      QueryColumn = QueryColumn & Cells(4, colLine).Value & ","
'    End If
'  Next
'
''  If EndLine > 1000 Then
''    EndLine = 100
''  End If
'
'
'  ' 実際のデータ取得
'  For line = 6 To endLine
'
'    ' プログレスバーのカウント変更（現在のカウント、全カウント数、メッセージ）
'    ProgressBar_ProgShowCount "処理中", line, endLine, SaveDirPath & "生成"
'
'    QueryValues = ""
'
'    For colLine = 1 To endColLine
'
'      ' 文字列型は、シングルクオートで囲む
'      Set FoundCell = Range("StringData").CurrentRegion.Find(What:=Cells(3, colLine).Value)
'      If FoundCell Is Nothing Then
'        QuartString = ""
'        If Cells(line, colLine).Value <> "" Then
'          SheetValue = (Cells(line, colLine).Value)
'        Else
'          SheetValue = "Null"
'        End If
'      Else
'        QuartString = "'"
'        SheetValue = (Cells(line, colLine).Value)
'        SheetValue = Replace(SheetValue, "'", "''")
'      End If
'
'      If Cells(line, colLine).Value = "" Then
'        QuartString = ""
'        SheetValue = "Null"
'      End If
'
'
'      If colLine = endColLine Then
'        QueryValues = QueryValues & QuartString & SheetValue & QuartString
'      Else
'        QueryValues = QueryValues & QuartString & SheetValue & QuartString & ","
'      End If
'    Next
'
'    Query = "insert into " & SQLTable & "(" & QueryColumn & ") Values(" & QueryValues & ");"
'
'    '編集した1項目分を出力
'    If line = 6 Then
'      ObjADODB.WriteText Range("A2"), 1
'    End If
'    ObjADODB.WriteText Query, 1
'
'    If (line - 5) Mod 5000 = 0 Or line = endLine Then
'      FileCnt = FileCnt + 1
'      fileName = Range("A1") & "_" & Range("B1") & "_" & Format(FileCnt, "00") & ".sql"
'      fileName = ActiveSheet.Name & "_" & Format(FileCnt, "00") & ".sql"
'      filePath = SaveDirPath & "\" & fileName
'
'      'UTF-8のBOM削除
'      If CharacterSet = "UTF-8" Then
'        ObjADODB.Position = 0
'        ObjADODB.Type = adTypeBinary
'        ObjADODB.Position = 3
'        byteData = ObjADODB.Read
'        ObjADODB.Close
'        ObjADODB.Open
'        ObjADODB.Write byteData
'      End If
'
'      'オブジェクトの内容をファイルに保存
'      ObjADODB.SaveToFile (filePath), 2
'
'      'オブジェクトを閉じる
'      ObjADODB.Close
'
'      'メモリからオブジェクトを削除する
'      Set ObjADODB = Nothing
'
'      If line <> endLine Then
'        Set ObjADODB = CreateObject("ADODB.Stream")
'
'        'オブジェクトに保存するデータの種類を文字列型に指定する
'        ObjADODB.Type = 2
'
'        '文字列型のオブジェクトの文字コードを指定する(改行コード：CR：13　CRLF：-1  LF：10)
'        ObjADODB.Charset = CharacterSet
'        ObjADODB.LineSeparator = LineSeparator
'
'
'
'        'オブジェクトのインスタンスを作成
'        ObjADODB.Open
'      End If
'    End If
'
'
'
'
'  Next
'
''  'オブジェクトの内容をファイルに保存
''  ObjADODB.SaveToFile (FilePath), 2
''
''  'オブジェクトを閉じる
''  ObjADODB.Close
''
''  'メモリからオブジェクトを削除する
''  Set ObjADODB = Nothing
'
'
'  ' プログレスバーの表示終了処理
'  ProgressBar_ProgShowClose
'
'  ' 画面描写制御終了
'  Library_EndScript
'
'  If ExecType <> "ALL" Then
'    Call Shell("Explorer.exe /select, " & filePath, vbNormalFocus)
'  End If
'  Exit Sub
'
'Error:
'  Set rs = Nothing
'  Set con = Nothing
'  MsgBox (Err.Description)
'
'  ' プログレスバーの表示終了処理
'  ProgressBar_ProgShowClose
'
'  ' 画面描写制御終了
'  Library_EndScript
'
'
'End Sub
'
''***********************************************************************************************************************************************
'' * Update文作成
'' *
'' * Bunpei.Koizumi<bunbun0716@gmail.com>
''***********************************************************************************************************************************************
'Sub MakeDML_Update(ExecType As String)
'
'  Dim line As Integer
'  Dim colLine As Long
'  Dim endLine As Long
'  Dim endColLine As Long
'  Dim ObjFileSys As Object
'  Dim ObjFile As Object
'  Dim filePath As String
'  Dim DirName As String
'  Dim Today As String
'  Dim SQLTable As String
'  Dim SheetValue As String
'
'  Dim QueryColumn As String
'  Dim QueryValues As String
'  Dim QuartString As String
'  Dim WhereString As String
'
'  Dim ObjADODB As Object
'  Dim FoundCell As Range
'
'  ' ファイルの保存ディレクトリの指定
'  If (SaveDirPath = "") Then
'    SaveDirPath = Library_GetDirPath(ActiveWorkbook.Path)
'    If (SaveDirPath = "") Then
'      Exit Sub
'    End If
'  End If
'  Worksheets("設定").Range("B12") = SaveDirPath
'
'  SQLTable = Range("B1")
'  Today = Format(Now, "yyyymmdd")
'
'  fileName = Range("A1") & "_" & ActiveSheet.Name & "_" & Today & ".sql"
''  FileName = ActiveSheet.Name & ".sql
'
'  filePath = SaveDirPath & "\" & fileName
'
'  ' プログレスバーの表示開始
'  ProgressBar_ProgShowStart
'  '  画面描写制御開始
'  Library_StartScript
'  On Error GoTo Error
'
'  Set ObjADODB = CreateObject("ADODB.Stream")
'
'  'オブジェクトに保存するデータの種類を文字列型に指定する
'  ObjADODB.Type = 2
'
'  '文字列型のオブジェクトの文字コードを指定する(改行コード：CR：13　CRLF：-1  LF：10)
'  ObjADODB.Charset = "UTF-8"
'  ObjADODB.LineSeparator = 10
'
'  'オブジェクトのインスタンスを作成
'  ObjADODB.Open
'
'  ' 行数取得
'  endLine = Cells(Rows.count, 1).End(xlUp).Row
'
'  ' 列数取得(基準は、4列目)
'  endColLine = Cells(4, Columns.count).End(xlToLeft).Column
'
'  '8行目から行方向へのループ
'  For line = 8 To endLine
'    QueryValues = ""
'    WhereString = ""
'
'    ' 列方向のループ
'    For colLine = 1 To endColLine
'
'      ' 文字列型は、シングルクオートで囲む
'      Set FoundCell = Range("StringData").CurrentRegion.Find(What:=Cells(3, colLine).Value)
'      If FoundCell Is Nothing Then
'        QuartString = ""
'        SheetValue = Trim(Cells(line, colLine).Value)
'      Else
'        QuartString = "'"
'        SheetValue = Trim(Cells(line, colLine).Value)
'        SheetValue = Replace(SheetValue, "'", "''")
'      End If
'
'
'      ' update指定を抜き出し
'      If Cells(4, colLine).Value = "update" Then
'        SheetValue = Cells(line, colLine).Value
'        SheetValue = Replace(SheetValue, "'", "''")
'
'        If QueryValues = "" Then
'          QueryValues = Cells(6, colLine).Value & " =" & QuartString & SheetValue & QuartString
'        Else
'          QueryValues = QueryValues & ", " & Cells(6, colLine).Value & " =" & QuartString & SheetValue & QuartString
'        End If
'      End If
'
'      ' where指定を抜き出し
'      If Cells(5, colLine).Value = "where" Then
'        SheetValue = Cells(line, colLine).Value
'        SheetValue = Replace(SheetValue, "'", "''")
'
'        If WhereString = "" Then
'          WhereString = "WHERE " & Cells(6, colLine).Value & " =" & QuartString & SheetValue & QuartString
'        Else
'          WhereString = WhereString & " AND " & Cells(6, colLine).Value & " =" & QuartString & SheetValue & QuartString
'        End If
'      End If
'    Next
'
'    ' SQL結合
'    Query = "update " & SQLTable & " set " & QueryValues & " " & WhereString & ";"
'    '編集した1項目分を出力
'    ObjADODB.WriteText Query, 1
'
'  Next
'
'  'オブジェクトの内容をファイルに保存
'  ObjADODB.SaveToFile (filePath), 2
'
'  'オブジェクトを閉じる
'  ObjADODB.Close
'
'  'メモリからオブジェクトを削除する
'  Set ObjADODB = Nothing
'
'  ' プログレスバーの表示終了処理
'  ProgressBar_ProgShowClose
'
'  ' 画面描写制御終了
'  Library_EndScript
'
'  If ExecType <> "ALL" Then
'    Call Shell("Explorer.exe /select, " & filePath, vbNormalFocus)
'  End If
'
'  Exit Sub
'
'Error:
'  Set rs = Nothing
'  Set con = Nothing
'  MsgBox (Err.Description)
'
'  ' プログレスバーの表示終了処理
'  ProgressBar_ProgShowClose
'
'  ' 画面描写制御終了
'  Library_EndScript
'End Sub
'
'
''***********************************************************************************************************************************************
'' * Delete文作成
'' *
'' * Bunpei.Koizumi<bunbun0716@gmail.com>
''***********************************************************************************************************************************************
'Sub MakeDML_Delete(ExecType As String)
'
'  Dim line As Integer
'  Dim colLine As Long
'  Dim endLine As Long
'  Dim endColLine As Long
'  Dim ObjFileSys As Object
'  Dim ObjFile As Object
'  Dim filePath As String
'  Dim DirName As String
'  Dim Today As String
'  Dim SQLTable As String
'  Dim SheetValue As String
'
'  Dim QueryColumn As String
'  Dim QueryValues As String
'  Dim QuartString As String
'  Dim WhereString As String
'
'  Dim ObjADODB As Object
'  Dim FoundCell As Range
'
'  ' ファイルの保存ディレクトリの指定
''  If (SaveDirPath = "") Then
''    SaveDirPath = Library_GetDirPath(ThisWorkbook.Path & "/..")
''  End If
'
'  SaveDirPath = Library_GetDirPath(ActiveWorkbook.Path & "/../")
'  If (SaveDirPath = "") Then
'      Exit Sub
'  End If
'
'  SQLTable = Range("B1")
'  Today = Format(Now, "yyyymmdd")
'
'  fileName = Range("A1") & "_" & Range("B1") & "_" & Today & ".sql"
''  FileName = ActiveSheet.Name & ".sql"
'  filePath = SaveDirPath & "\" & fileName
'
'  ' プログレスバーの表示開始
'  ProgressBar_ProgShowStart
'  '  画面描写制御開始
'  Library_StartScript
'  On Error GoTo Error
'
'  Set ObjADODB = CreateObject("ADODB.Stream")
'
'  'オブジェクトに保存するデータの種類を文字列型に指定する
'  ObjADODB.Type = 2
'
'  '文字列型のオブジェクトの文字コードを指定する(改行コード：CR：13　CRLF：-1  LF：10)
'  ObjADODB.Charset = "UTF-8"
'  ObjADODB.LineSeparator = 10
'
'  'オブジェクトのインスタンスを作成
'  ObjADODB.Open
'
'  ' 行数取得
'  endLine = Cells(Rows.count, 1).End(xlUp).Row
'
'  ' 列数取得(基準は、4列目)
'  endColLine = Cells(4, Columns.count).End(xlToLeft).Column
'
'  '8行目から行方向へのループ
'  For line = 8 To endLine
'    QueryValues = ""
'    WhereString = ""
'
'    ' 列方向のループ
'    For colLine = 1 To endColLine
'
'      ' 文字列型は、シングルクオートで囲む
'      Set FoundCell = Range("StringData").CurrentRegion.Find(What:=Cells(3, colLine).Value)
'      If FoundCell Is Nothing Then
'        QuartString = ""
'        SheetValue = Trim(Cells(line, colLine).Value)
'      Else
'        QuartString = "'"
'        SheetValue = Trim(Cells(line, colLine).Value)
'        SheetValue = Replace(SheetValue, "'", "''")
'      End If
'
'      ' where指定を抜き出し
'      If Cells(4, colLine).Value = "where" Then
'        SheetValue = Cells(line, colLine).Value
'        SheetValue = Replace(SheetValue, "'", "''")
'
'        If WhereString = "" Then
'          WhereString = "WHERE " & Cells(6, colLine).Value & " =" & QuartString & SheetValue & QuartString
'        Else
'          WhereString = WhereString & " AND " & Cells(6, colLine).Value & " =" & QuartString & SheetValue & QuartString
'        End If
'      End If
'    Next
'
'    ' SQL結合
'    Query = "DELETE FROM" & SQLTable & WhereString & ";"
'    '編集した1項目分を出力
'    ObjADODB.WriteText Query, 1
'
'  Next
'
'  'オブジェクトの内容をファイルに保存
'  ObjADODB.SaveToFile (filePath), 2
'
'  'オブジェクトを閉じる
'  ObjADODB.Close
'
'  'メモリからオブジェクトを削除する
'  Set ObjADODB = Nothing
'
'  ' プログレスバーの表示終了処理
'  ProgressBar_ProgShowClose
'
'  ' 画面描写制御終了
'  Library_EndScript
'
'  If ExecType <> "ALL" Then
'    Call Shell("Explorer.exe , " & filePath, vbNormalFocus)
'  End If
'
'  Exit Sub
'
'Error:
'  Set rs = Nothing
'  Set con = Nothing
'  MsgBox (Err.Description)
'
'  ' プログレスバーの表示終了処理
'  ProgressBar_ProgShowClose
'
'  ' 画面描写制御終了
'  Library_EndScript
'End Sub
'
'
''***********************************************************************************************************************************************
'' * カラム情報取得
'' *
'' * Bunpei.Koizumi<bunbun0716@gmail.com>
''***********************************************************************************************************************************************
'Sub MakeDML_SetColumnForPostgreSQL()
'
'  Dim tableName As String
'  Dim CMDType As String
'
'  Dim rowNo As Integer
'
'  '  画面描写制御開始
'  Library_StartScript
'
'  Call MakeDML_Init
'
'  ' プログレスバーの表示開始
'  ProgressBar_ProgShowStart
'
'
'  tableName = Range("B1").Value
'  CMDType = Range("A1").Value
'
'  '接続文字列
'  Select Case BDMS
'    Case "PostgreSQL-64Bit"
'      ConnectionString = "Driver={PostgreSQL Unicode(X64)}; SERVER=" & BDServer & ";PORT=5432;DATABASE=" & DBName & ";UID=" & LoginID & ";PWD=" & LoginPW & ";"
'
'    Case "PostgreSQL-32Bit"
'      ConnectionString = "Driver={PostgreSQL Unicode}; SERVER=" & BDServer & ";PORT=5432;DATABASE=" & DBName & ";UID=" & LoginID & ";PWD=" & LoginPW & ";"
'
'    Case Else
'      ConnectionString = "Driver={PostgreSQL Unicode(X64)}; SERVER=" & BDServer & ";PORT=5432;DATABASE=" & DBName & ";UID=" & LoginID & ";PWD=" & LoginPW & ";"
'
'  End Select
'
'
'  'ADODB.Connection生成
'  Set con = New ADODB.Connection
'
'  On Error GoTo Err
'
'  'DBに接続
'  con.Open ConnectionString
'
'  'SQL文 ----------------------------------------------------------
'  sqlStr = "select column_name,data_type,is_nullable from information_schema.columns where table_catalog='" & DBName & "' and table_name='" & tableName & "' order by ordinal_position;"
'
'  'SQL文実行
'  Set rs = con.Execute(sqlStr)
'
'  rowNo = 1
'
'  'RecordSetの終了まで
'  Do While rs.EOF = False
'
'    ' プログレスバーのカウント変更（現在のカウント、全カウント数、メッセージ）
'    ProgressBar_ProgShowCount "処理中", rowNo, 200, rs.Fields("column_name")
'
'    Select Case rs.Fields("data_type")
'    Case "timestamp without time zone"
'      Cells(3, rowNo).Value = "timestamp"
'
'    Case "character"
'      Cells(3, rowNo).Value = "text"
'
'    Case "character varying"
'      Cells(3, rowNo).Value = "text"
'
'
'
'
'    Case Else
'       Cells(3, rowNo).Value = rs.Fields("data_type")
'  End Select
'
'
'
'    If CMDType = "INSERT" Then
'      Cells(4, rowNo).Value = rs.Fields("column_name")
'
'      If rs.Fields("is_nullable") = "NO" Then
'        Cells(4, rowNo).Select
'        Selection.Font.Bold = True
'      End If
'      Range(Cells(3, rowNo), Cells(5, rowNo)).Select
'      SetStyle_SetColumnForPostgreSQL_1
'      Range(Cells(6, rowNo), Cells(20, rowNo)).Select
'      SetStyle_SetColumnForPostgreSQL_2
'
'      '中央寄せ
'      Range(Cells(3, rowNo), Cells(5, rowNo)).Select
'      With Selection
'          .HorizontalAlignment = xlCenter
'          .VerticalAlignment = xlCenter
'          .WrapText = False
'          .Orientation = 0
'          .AddIndent = False
'          .IndentLevel = 0
'          .ShrinkToFit = False
'          .ReadingOrder = xlContext
'          .MergeCells = False
'      End With
'
'
'
'
'    Else
'      Cells(6, rowNo).Value = rs.Fields("column_name")
'
'      If rs.Fields("is_nullable") = "NO" Then
'        Cells(6, rowNo).Select
'        Selection.Font.Bold = True
'      End If
'
'      '背景色設定
'      Range(Cells(3, rowNo), Cells(7, rowNo)).Select
'      SetStyle_SetColumnForPostgreSQL_1
'
'      '入力部分の罫線設定
'      Range(Cells(8, rowNo), Cells(50, rowNo)).Select
'      SetStyle_SetColumnForPostgreSQL_2
'
'      '入力規則設定
'      Cells(4, rowNo).Select
'      With Selection.Validation
'          .delete
'          .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'          xlBetween, Formula1:="update"
'          .IgnoreBlank = True
'          .InCellDropdown = True
'          .InputTitle = ""
'          .ErrorTitle = ""
'          .InputMessage = ""
'          .ErrorMessage = ""
'          .IMEMode = xlIMEModeNoControl
'          .ShowInput = True
'          .ShowError = True
'      End With
'      Cells(5, rowNo).Select
'      With Selection.Validation
'          .delete
'          .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'          xlBetween, Formula1:="where"
'          .IgnoreBlank = True
'          .InCellDropdown = True
'          .InputTitle = ""
'          .ErrorTitle = ""
'          .InputMessage = ""
'          .ErrorMessage = ""
'          .IMEMode = xlIMEModeNoControl
'          .ShowInput = True
'          .ShowError = True
'      End With
'
'      '中央寄せ
'      Range(Cells(3, rowNo), Cells(7, rowNo)).Select
'      With Selection
'          .HorizontalAlignment = xlCenter
'          .VerticalAlignment = xlCenter
'          .WrapText = False
'          .Orientation = 0
'          .AddIndent = False
'          .IndentLevel = 0
'          .ShrinkToFit = False
'          .ReadingOrder = xlContext
'          .MergeCells = False
'      End With
'    End If
'
'    rowNo = rowNo + 1
'    '次のレコード
'    rs.MoveNext
'  Loop
'
'
'  'クローズ
'  con.Close
'  Set rs = Nothing
'  Set con = Nothing
'
'  ' プログレスバーの表示終了処理
'
'  ProgressBar_ProgShowClose
'
'  ' 画面描写制御終了
'  Library_EndScript
'
'  Exit Sub
'
'Err:
'  Set rs = Nothing
'  Set con = Nothing
'  MsgBox (Err.Description)
'
'  ' プログレスバーの表示終了処理
'
'  ProgressBar_ProgShowClose
'
'  ' 画面描写制御終了
'  Library_EndScript
'
'End Sub
