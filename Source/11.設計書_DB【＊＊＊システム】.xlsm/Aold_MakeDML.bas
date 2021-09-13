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
'' * DB���擾
'' *
'' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
''***********************************************************************************************************************************************
'Function MakeDML_Init()
'
'  DebugFlg = True
'
'  ODBCDriver = Worksheets("�ݒ�").Range("B3").Value
'  DBMS = Application.WorksheetFunction.VLookup(ODBCDriver, Range("ODBCDriverList"), 2, False)
'
'  SetDisplyAlertFlg = True
'  BeforeCloseFlg = False
'
'
'  DBPort = Worksheets("�ݒ�").Range("B9").Value
'
'  LoginID = Worksheets("�ݒ�").Range("B10").Value
'  LoginPW = Worksheets("�ݒ�").Range("B11").Value
'
'  Select Case DBMS
'    Case "PostgreSQL"
'      DBServer = Worksheets("�ݒ�").Range("B4").Value
'      DBName = Worksheets("�ݒ�").Range("B5").Value
'      DBInstance = Worksheets("�ݒ�").Range("B6").Value
'      DBScheme = Worksheets("�ݒ�").Range("B7").Value
'
'      ConnectionString = ""
'
'    Case "MySQL"
'      DBServer = Worksheets("�ݒ�").Range("B4").Value
'      DBName = Worksheets("�ݒ�").Range("B5").Value
'      DBInstance = Worksheets("�ݒ�").Range("B6").Value
'      DBScheme = Worksheets("�ݒ�").Range("B7").Value
'
'      ConnectionString = "Driver={" & ODBCDriver & "}; Server=" & DBServer & ";Port=" & _
'                          DBPort & ";Database=" & DBName & ";UID=" & LoginID & ";PWD=" & LoginPW & ""
'
'    Case "Oracle"
'      DBServer = Worksheets("�ݒ�").Range("B4").Value
'      DBName = Worksheets("�ݒ�").Range("B5").Value
'
'      DBTableSpace = Worksheets("�ݒ�").Range("B6").Value
'      DBInstance = Worksheets("�ݒ�").Range("B7").Value
'      DBScheme = Worksheets("�ݒ�").Range("B8").Value
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
'  SaveDirPath = Worksheets("�ݒ�").Range("B12")
'  DBMode = Worksheets("�ݒ�").Range("B13")
'
'  CharacterSet = Worksheets("�ݒ�").Range("B13").Value
'
'  Select Case Worksheets("�ݒ�").Range("B14").Value
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
'' * SQL�����p�V�[�g�ǉ�
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
'' * �J�������Ď擾
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
'  ' �v���O���X�o�[�̕\���J�n
'  ProgressBar_ProgShowStart
'
'
'  'ADODB.Connection�������ADB�ɐڑ�
'  ProgressBar_ProgShowCount "������", 5, 100, "DB�ɐڑ�"
'
'  Set dbCon = New ADODB.Connection
'  dbCon.Open ConnectionString
'  ProgressBar_ProgShowCount "������", 10, 100, "DB�ɐڑ�"
'
'  'SQL�� ----------------------------------------------------------
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
'  ProgressBar_ProgShowCount "������", 15, 100, "DB�ɐڑ�"
'
'  Set DBRecordset = New ADODB.Recordset
'
'  '�J�������擾
'  RunQueryString = "select count(*) as count " & QueryString
'  Set DBRecordset = New ADODB.Recordset
'  DBRecordset.Open RunQueryString, dbCon, adOpenKeyset, adLockReadOnly
'  Do Until DBRecordset.EOF
'    RowCount = CLng(DBRecordset.Fields("count").Value)
'
'    '���̃��R�[�h
'    DBRecordsetCount = DBRecordsetCount + 1
'    DBRecordset.MoveNext
'  Loop
'  Set DBRecordset = Nothing
'  ProgressBar_ProgShowCount "������", 50, 100, "�J�������擾"
'
'  '�J�������擾
'  RunQueryString = SelectString & QueryString
'
'  Set DBRecordset = New ADODB.Recordset
'  DBRecordset.Open RunQueryString, dbCon, adOpenKeyset, adLockReadOnly
'  ProgressBar_ProgShowCount "������", 100, 100, "�J�������擾"
'
'  Do Until DBRecordset.EOF
'    ' �v���O���X�o�[�̃J�E���g�ύX�i���݂̃J�E���g�A�S�J�E���g���A���b�Z�[�W�j
'    ProgressBar_ProgShowCount "������", NowLine, RowCount, "�J�������擾�F" & columnName
'
'    '�R�����g(���ږ��Ƃ��ė��p)
'    If IsNull(DBRecordset.Fields("Comments").Value) Then
'      Comments = ""
'    Else
'      Comments = DBRecordset.Fields("Comments").Value
'    End If
'
'    '�J������
'    columnName = DBRecordset.Fields("ColumName").Value
'
'    '�^
'    dataType = DBRecordset.Fields("DataType").Value
'    Select Case dataType
'      Case "TIMESTAMP(6)"
'        dataType = "TIMESTAMP"
'    End Select
'
'    '����
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
'    '�v���C�}���L�[
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
'    'NotNULL����
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
'    '�����l
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
'    '�J�����ݒ�
'    Select Case Range("A1")
'      Case "INSERT"
'        Sheets("Insert��").Range("B3:B105").Copy
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
'        Sheets("Update��").Range("B3:B107").Copy
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
'        Sheets("Delete��").Range("B3:B104").Copy
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
'    '���̃��R�[�h
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
'' * SQL�����pSQL����
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
'' * �SSQL�����pSQL����
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
'' * Insert���쐬
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
'  ' �t�@�C���̕ۑ��f�B���N�g���̎w��
'  If (Worksheets("�ݒ�").Range("B12") = "") Then
'    SaveDirPath = Library_GetDirPath(ActiveWorkbook.Path)
'    If (SaveDirPath = "") Then
'      Exit Sub
'    End If
'  End If
'  Worksheets("�ݒ�").Range("B12") = SaveDirPath
'
'  SQLTable = Range("B1")
'  Today = Format(Now, "yyyymmdd")
'  FileCnt = 0
'
'
'  ' �v���O���X�o�[�̕\���J�n
'  ProgressBar_ProgShowStart
'
'
'  Set ObjADODB = CreateObject("ADODB.Stream")
'
'  '�I�u�W�F�N�g�ɕۑ�����f�[�^�̎�ނ𕶎���^�Ɏw�肷��
'  ObjADODB.Type = 2
'
'  '������^�̃I�u�W�F�N�g�̕����R�[�h���w�肷��(���s�R�[�h�FCR�F13�@CRLF�F-1  LF�F10)
'  ObjADODB.Charset = CharacterSet
'  ObjADODB.LineSeparator = LineSeparator
'
'  '�I�u�W�F�N�g�̃C���X�^���X���쐬
'  ObjADODB.Open
'
'
'  endLine = Cells(Rows.count, 1).End(xlUp).Row
'  endColLine = Cells(4, Columns.count).End(xlToLeft).Column
'
'  ' �J�������擾
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
'  ' ���ۂ̃f�[�^�擾
'  For line = 6 To endLine
'
'    ' �v���O���X�o�[�̃J�E���g�ύX�i���݂̃J�E���g�A�S�J�E���g���A���b�Z�[�W�j
'    ProgressBar_ProgShowCount "������", line, endLine, SaveDirPath & "����"
'
'    QueryValues = ""
'
'    For colLine = 1 To endColLine
'
'      ' ������^�́A�V���O���N�I�[�g�ň͂�
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
'    '�ҏW����1���ڕ����o��
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
'      'UTF-8��BOM�폜
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
'      '�I�u�W�F�N�g�̓��e���t�@�C���ɕۑ�
'      ObjADODB.SaveToFile (filePath), 2
'
'      '�I�u�W�F�N�g�����
'      ObjADODB.Close
'
'      '����������I�u�W�F�N�g���폜����
'      Set ObjADODB = Nothing
'
'      If line <> endLine Then
'        Set ObjADODB = CreateObject("ADODB.Stream")
'
'        '�I�u�W�F�N�g�ɕۑ�����f�[�^�̎�ނ𕶎���^�Ɏw�肷��
'        ObjADODB.Type = 2
'
'        '������^�̃I�u�W�F�N�g�̕����R�[�h���w�肷��(���s�R�[�h�FCR�F13�@CRLF�F-1  LF�F10)
'        ObjADODB.Charset = CharacterSet
'        ObjADODB.LineSeparator = LineSeparator
'
'
'
'        '�I�u�W�F�N�g�̃C���X�^���X���쐬
'        ObjADODB.Open
'      End If
'    End If
'
'
'
'
'  Next
'
''  '�I�u�W�F�N�g�̓��e���t�@�C���ɕۑ�
''  ObjADODB.SaveToFile (FilePath), 2
''
''  '�I�u�W�F�N�g�����
''  ObjADODB.Close
''
''  '����������I�u�W�F�N�g���폜����
''  Set ObjADODB = Nothing
'
'
'  ' �v���O���X�o�[�̕\���I������
'  ProgressBar_ProgShowClose
'
'  ' ��ʕ`�ʐ���I��
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
'  ' �v���O���X�o�[�̕\���I������
'  ProgressBar_ProgShowClose
'
'  ' ��ʕ`�ʐ���I��
'  Library_EndScript
'
'
'End Sub
'
''***********************************************************************************************************************************************
'' * Update���쐬
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
'  ' �t�@�C���̕ۑ��f�B���N�g���̎w��
'  If (SaveDirPath = "") Then
'    SaveDirPath = Library_GetDirPath(ActiveWorkbook.Path)
'    If (SaveDirPath = "") Then
'      Exit Sub
'    End If
'  End If
'  Worksheets("�ݒ�").Range("B12") = SaveDirPath
'
'  SQLTable = Range("B1")
'  Today = Format(Now, "yyyymmdd")
'
'  fileName = Range("A1") & "_" & ActiveSheet.Name & "_" & Today & ".sql"
''  FileName = ActiveSheet.Name & ".sql
'
'  filePath = SaveDirPath & "\" & fileName
'
'  ' �v���O���X�o�[�̕\���J�n
'  ProgressBar_ProgShowStart
'  '  ��ʕ`�ʐ���J�n
'  Library_StartScript
'  On Error GoTo Error
'
'  Set ObjADODB = CreateObject("ADODB.Stream")
'
'  '�I�u�W�F�N�g�ɕۑ�����f�[�^�̎�ނ𕶎���^�Ɏw�肷��
'  ObjADODB.Type = 2
'
'  '������^�̃I�u�W�F�N�g�̕����R�[�h���w�肷��(���s�R�[�h�FCR�F13�@CRLF�F-1  LF�F10)
'  ObjADODB.Charset = "UTF-8"
'  ObjADODB.LineSeparator = 10
'
'  '�I�u�W�F�N�g�̃C���X�^���X���쐬
'  ObjADODB.Open
'
'  ' �s���擾
'  endLine = Cells(Rows.count, 1).End(xlUp).Row
'
'  ' �񐔎擾(��́A4���)
'  endColLine = Cells(4, Columns.count).End(xlToLeft).Column
'
'  '8�s�ڂ���s�����ւ̃��[�v
'  For line = 8 To endLine
'    QueryValues = ""
'    WhereString = ""
'
'    ' ������̃��[�v
'    For colLine = 1 To endColLine
'
'      ' ������^�́A�V���O���N�I�[�g�ň͂�
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
'      ' update�w��𔲂��o��
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
'      ' where�w��𔲂��o��
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
'    ' SQL����
'    Query = "update " & SQLTable & " set " & QueryValues & " " & WhereString & ";"
'    '�ҏW����1���ڕ����o��
'    ObjADODB.WriteText Query, 1
'
'  Next
'
'  '�I�u�W�F�N�g�̓��e���t�@�C���ɕۑ�
'  ObjADODB.SaveToFile (filePath), 2
'
'  '�I�u�W�F�N�g�����
'  ObjADODB.Close
'
'  '����������I�u�W�F�N�g���폜����
'  Set ObjADODB = Nothing
'
'  ' �v���O���X�o�[�̕\���I������
'  ProgressBar_ProgShowClose
'
'  ' ��ʕ`�ʐ���I��
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
'  ' �v���O���X�o�[�̕\���I������
'  ProgressBar_ProgShowClose
'
'  ' ��ʕ`�ʐ���I��
'  Library_EndScript
'End Sub
'
'
''***********************************************************************************************************************************************
'' * Delete���쐬
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
'  ' �t�@�C���̕ۑ��f�B���N�g���̎w��
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
'  ' �v���O���X�o�[�̕\���J�n
'  ProgressBar_ProgShowStart
'  '  ��ʕ`�ʐ���J�n
'  Library_StartScript
'  On Error GoTo Error
'
'  Set ObjADODB = CreateObject("ADODB.Stream")
'
'  '�I�u�W�F�N�g�ɕۑ�����f�[�^�̎�ނ𕶎���^�Ɏw�肷��
'  ObjADODB.Type = 2
'
'  '������^�̃I�u�W�F�N�g�̕����R�[�h���w�肷��(���s�R�[�h�FCR�F13�@CRLF�F-1  LF�F10)
'  ObjADODB.Charset = "UTF-8"
'  ObjADODB.LineSeparator = 10
'
'  '�I�u�W�F�N�g�̃C���X�^���X���쐬
'  ObjADODB.Open
'
'  ' �s���擾
'  endLine = Cells(Rows.count, 1).End(xlUp).Row
'
'  ' �񐔎擾(��́A4���)
'  endColLine = Cells(4, Columns.count).End(xlToLeft).Column
'
'  '8�s�ڂ���s�����ւ̃��[�v
'  For line = 8 To endLine
'    QueryValues = ""
'    WhereString = ""
'
'    ' ������̃��[�v
'    For colLine = 1 To endColLine
'
'      ' ������^�́A�V���O���N�I�[�g�ň͂�
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
'      ' where�w��𔲂��o��
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
'    ' SQL����
'    Query = "DELETE FROM" & SQLTable & WhereString & ";"
'    '�ҏW����1���ڕ����o��
'    ObjADODB.WriteText Query, 1
'
'  Next
'
'  '�I�u�W�F�N�g�̓��e���t�@�C���ɕۑ�
'  ObjADODB.SaveToFile (filePath), 2
'
'  '�I�u�W�F�N�g�����
'  ObjADODB.Close
'
'  '����������I�u�W�F�N�g���폜����
'  Set ObjADODB = Nothing
'
'  ' �v���O���X�o�[�̕\���I������
'  ProgressBar_ProgShowClose
'
'  ' ��ʕ`�ʐ���I��
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
'  ' �v���O���X�o�[�̕\���I������
'  ProgressBar_ProgShowClose
'
'  ' ��ʕ`�ʐ���I��
'  Library_EndScript
'End Sub
'
'
''***********************************************************************************************************************************************
'' * �J�������擾
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
'  '  ��ʕ`�ʐ���J�n
'  Library_StartScript
'
'  Call MakeDML_Init
'
'  ' �v���O���X�o�[�̕\���J�n
'  ProgressBar_ProgShowStart
'
'
'  tableName = Range("B1").Value
'  CMDType = Range("A1").Value
'
'  '�ڑ�������
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
'  'ADODB.Connection����
'  Set con = New ADODB.Connection
'
'  On Error GoTo Err
'
'  'DB�ɐڑ�
'  con.Open ConnectionString
'
'  'SQL�� ----------------------------------------------------------
'  sqlStr = "select column_name,data_type,is_nullable from information_schema.columns where table_catalog='" & DBName & "' and table_name='" & tableName & "' order by ordinal_position;"
'
'  'SQL�����s
'  Set rs = con.Execute(sqlStr)
'
'  rowNo = 1
'
'  'RecordSet�̏I���܂�
'  Do While rs.EOF = False
'
'    ' �v���O���X�o�[�̃J�E���g�ύX�i���݂̃J�E���g�A�S�J�E���g���A���b�Z�[�W�j
'    ProgressBar_ProgShowCount "������", rowNo, 200, rs.Fields("column_name")
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
'      '������
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
'      '�w�i�F�ݒ�
'      Range(Cells(3, rowNo), Cells(7, rowNo)).Select
'      SetStyle_SetColumnForPostgreSQL_1
'
'      '���͕����̌r���ݒ�
'      Range(Cells(8, rowNo), Cells(50, rowNo)).Select
'      SetStyle_SetColumnForPostgreSQL_2
'
'      '���͋K���ݒ�
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
'      '������
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
'    '���̃��R�[�h
'    rs.MoveNext
'  Loop
'
'
'  '�N���[�Y
'  con.Close
'  Set rs = Nothing
'  Set con = Nothing
'
'  ' �v���O���X�o�[�̕\���I������
'
'  ProgressBar_ProgShowClose
'
'  ' ��ʕ`�ʐ���I��
'  Library_EndScript
'
'  Exit Sub
'
'Err:
'  Set rs = Nothing
'  Set con = Nothing
'  MsgBox (Err.Description)
'
'  ' �v���O���X�o�[�̕\���I������
'
'  ProgressBar_ProgShowClose
'
'  ' ��ʕ`�ʐ���I��
'  Library_EndScript
'
'End Sub
