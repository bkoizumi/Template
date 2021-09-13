Attribute VB_Name = "Aold_DataBase"
Public SaveDirPath As String
Public ODBCDriver As String
Public DBMS As String
Public DBServer As String
Public DBTableSpace As String
Public DBName As String
Public DBInstance As String
Public DBScheme As String
Public DBPort As String

Public LoginID As String
Public LoginPW As String
Public FlameWorkName As String
Public SetDisplyAlertFlg As Boolean
Public SetDisplyProgressBarFlg As Boolean
Public SetSelectTargetRows As String

Public InputTableName As String
Public InputTableIDa As String
Public BeforeCloseFlg As Boolean

Public DebugFlg As Boolean
Public ConnectionString As String
Public LineBreakCode As String
Public LineSeparator As Integer
Public CharacterSet As String
Public DBMode As String


Public DBRecordsetCount As Integer
Public dbCon As ADODB.Connection
'Public DBRecordset As ADODB.Recordset



'***************************************************************************************************************************************************
' * DB���擾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***************************************************************************************************************************************************
Function DataBase_Init()

  DebugFlg = True

  ODBCDriver = Worksheets("�ݒ�").Range("B3").Value
  DBMS = Application.WorksheetFunction.VLookup(ODBCDriver, Range("ODBCDriverList"), 2, False)

  SetDisplyAlertFlg = True
  BeforeCloseFlg = False


  DBPort = Worksheets("�ݒ�").Range("B9").Value

  LoginID = Worksheets("�ݒ�").Range("B10").Value
  LoginPW = Worksheets("�ݒ�").Range("B11").Value

  Select Case DBMS
    Case "PostgreSQL"
      DBServer = Worksheets("�ݒ�").Range("B4").Value
      DBName = Worksheets("�ݒ�").Range("B5").Value
      DBInstance = Worksheets("�ݒ�").Range("B6").Value
      DBScheme = Worksheets("�ݒ�").Range("B7").Value

      ConnectionString = ""

    Case "MySQL"
      DBServer = Worksheets("�ݒ�").Range("B4").Value
      DBName = Worksheets("�ݒ�").Range("B5").Value
      DBInstance = Worksheets("�ݒ�").Range("B6").Value
      DBScheme = Worksheets("�ݒ�").Range("B7").Value

      ConnectionString = "Driver={" & ODBCDriver & "}; Server=" & DBServer & ";Port=" & DBPort & _
                          ";Option=131072;Stmt=SET CHARACTER SET SJIS;Database=" & DBName & ";Uid=" & LoginID & ";Pwd=" & LoginPW & ""


    Case "Oracle"
      DBServer = Worksheets("�ݒ�").Range("B4").Value
      DBName = Worksheets("�ݒ�").Range("B5").Value

      DBTableSpace = Worksheets("�ݒ�").Range("B6").Value
      DBInstance = Worksheets("�ݒ�").Range("B7").Value
      DBScheme = Worksheets("�ݒ�").Range("B8").Value


      ConnectionString = "Driver={" & ODBCDriver & "};DBQ=" & DBName & ";UID=" & LoginID & ";PWD=" & LoginPW & ""

    Case "SQLServer"
      ConnectionString = "Provider=SQLOLEDB.1;Data Source=" & DBServer & ";Initial Catalog=" & _
                          DBName & ";User ID=" & LoginID & ";Password=" & LoginPW & ";"

  End Select


  SaveDirPath = Worksheets("�ݒ�").Range("B12")

  CharacterSet = Worksheets("�ݒ�").Range("B13").Value

  Select Case Worksheets("�ݒ�").Range("B14").Value
  Case "CRLF"
    LineSeparator = -1
    LineBreakCode = vbCrLf
  Case "LF"
    LineSeparator = 10
    LineBreakCode = vbLf
  Case "CR"
    LineSeparator = 13
    LineBreakCode = vbCr
  End Select

'  LineBreakCode = vbLf
'  CharacterSet = "UTF-8"


End Function

'***************************************************************************************************************************************************
' * �ݒ�V�[�g�Đݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***************************************************************************************************************************************************
Function DataBase_Reset()

  On Error Resume Next

  Worksheets("�ݒ�").Select

  ' �ݒ�ς̖��O���폜
  Dim nm As Name
  For Each nm In ActiveWorkbook.Names
    nm.Delete
  Next nm

  SaveDirPath = ""

  ActiveWorkbook.Names.Add Name:="DBMS", RefersTo:=Range("B3")
  ActiveWorkbook.Names.Add Name:="DB��", RefersTo:=Range("B5")
  ActiveWorkbook.Names.Add Name:="�e�[�u���X�y�[�X", RefersTo:=Range("B6")

  ActiveWorkbook.Names.Add Name:="�C���X�^���X", RefersTo:=Range("B7")
  ActiveWorkbook.Names.Add Name:="�X�L�[�}", RefersTo:=Range("B8")

  ActiveWorkbook.Names.Add Name:="ODBCDriver", RefersTo:=Range("K3:K" & Cells(Rows.count, 11).End(xlUp).Row)
  ActiveWorkbook.Names.Add Name:="ODBCDriverList", RefersTo:=Range("K3:L" & Cells(Rows.count, 11).End(xlUp).Row)



  ActiveWorkbook.Names.Add Name:="�ڋq��", RefersTo:=Range("E3")
  ActiveWorkbook.Names.Add Name:="�쐬��", RefersTo:=Range("E4")
  ActiveWorkbook.Names.Add Name:="�쐬��", RefersTo:=Range("E5")
  ActiveWorkbook.Names.Add Name:="�X�V��", RefersTo:=Range("E6")
  ActiveWorkbook.Names.Add Name:="�X�V��", RefersTo:=Range("E7")
  ActiveWorkbook.Names.Add Name:="�v���W�F�N�g��", RefersTo:=Range("E8")
  ActiveWorkbook.Names.Add Name:="�V�X�e����", RefersTo:=Range("E9")
  ActiveWorkbook.Names.Add Name:="�\���^�C�g������", RefersTo:=Range("E10")
  ActiveWorkbook.Names.Add Name:="�e�[�u�����", RefersTo:=Range("G3:G" & Cells(Rows.count, 7).End(xlUp).Row)

  Worksheets("DataType").Select
  ActiveWorkbook.Names.Add Name:="PostgreSQL", RefersTo:=Range("A3:A" & Cells(Rows.count, 1).End(xlUp).Row)
  ActiveWorkbook.Names.Add Name:="MySQL", RefersTo:=Range("E3:E" & Cells(Rows.count, 5).End(xlUp).Row)
  ActiveWorkbook.Names.Add Name:="Oracle", RefersTo:=Range("I3:I" & Cells(Rows.count, 9).End(xlUp).Row)
  ActiveWorkbook.Names.Add Name:="SQLServer", RefersTo:=Range("M3:M" & Cells(Rows.count, 13).End(xlUp).Row)

  Worksheets("�ύX����").Select
  ActiveWorkbook.Names.Add Name:="�����", RefersTo:=Range("C6:C100")
  ActiveWorkbook.Names.Add Name:="�����", RefersTo:=Range("B6:B100")

  '�g���ʕ\���p�̖��̐ݒ�
  Dim sheetName As String
  Dim endLine As Integer

  For Each objSheet In ActiveWorkbook.Sheets
    sheetName = objSheet.Name

    If Library_CheckExcludeSheet(sheetName, 9) Then
      endLine = Worksheets(sheetName).Cells(Rows.count, 2).End(xlUp).Row

      '�Z�b�g�X�e�[�g�����g
'      ActiveWorkbook.Worksheets(SheetName).Names.Add Name:="SetStatement", RefersToR1C1:=Worksheets(SheetName).Range("D7")
'      ActiveWorkbook.Worksheets(SheetName).Names("SetStatement").Comment = "�g���ʕ\���p�̖��̐ݒ�"
'
'      '�g���K�[
'      ActiveWorkbook.Worksheets(SheetName).Names.Add Name:="Trigger1", RefersToR1C1:=Worksheets(SheetName).Range("H" & Endline - 3)
'      ActiveWorkbook.Worksheets(SheetName).Names("Trigger1").Comment = "�g���ʕ\���p�̖��̐ݒ�"
'
'      ActiveWorkbook.Worksheets(SheetName).Names.Add Name:="Trigger2", RefersToR1C1:=Worksheets(SheetName).Range("H" & Endline - 2)
'      ActiveWorkbook.Worksheets(SheetName).Names("Trigger2").Comment = "�g���ʕ\���p�̖��̐ݒ�"
'
'      ActiveWorkbook.Worksheets(SheetName).Names.Add Name:="Trigger3", RefersToR1C1:=Worksheets(SheetName).Range("H" & Endline - 1)
'      ActiveWorkbook.Worksheets(SheetName).Names("Trigger3").Comment = "�g���ʕ\���p�̖��̐ݒ�"
'
'      ActiveWorkbook.Worksheets(SheetName).Names.Add Name:="Trigger4", RefersToR1C1:=Worksheets(SheetName).Range("H" & Endline)
'      ActiveWorkbook.Worksheets(SheetName).Names("Trigger4").Comment = "�g���ʕ\���p�̖��̐ݒ�"

      ActiveWorkbook.Worksheets(sheetName).Select
      ActiveWindow.DisplayGridlines = False
      ActiveWindow.FreezePanes = True

      'TBL���X�g�\�� �{�^���ݒ�
      ActiveWorkbook.ActiveSheet.Shapes.Range(Array("Button 1")).Select
      Selection.OnAction = ActiveWorkbook.Name & "!DisplayTableList"

      ActiveWorkbook.ActiveSheet.Select
      If sheetName <> Worksheets(sheetName).Range("H5") Then
        Worksheets(sheetName).Name = Worksheets(sheetName).Range("H5")
      End If

      Range("A9").Select
      Library_UnsetLineColor ("B9:U" & Cells(Rows.count, 2).End(xlUp).Row)
      Call Library_SetLineColor("B9:U" & Cells(Rows.count, 2).End(xlUp).Row, False, RGB(255, 255, 155))
    End If
    ActiveWindow.Zoom = 90
    Range("A9").Select
  Next

  Worksheets("�ݒ�").Select

End Function


'***************************************************************************************************************************************************
' * �S�V�[�g��SQL���ꊇ�쐬
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***************************************************************************************************************************************************
Sub DataBase_MakeAllSheetSQL()

  Dim sheetName As String


  For Each objSheet In ActiveWorkbook.Sheets

    sheetName = objSheet.Name

    If Library_CheckExcludeSheet(sheetName, 9) Then
      Worksheets(sheetName).Select
      Range("C9").Select
      SetDisplyAlertFlg = False
      DataBase_MakeSQL (False)
    End If
  Next

  MsgBox ("�X�N���v�g�t�@�C���̍쐬���������܂����B" & LineBreakCode & SaveDirPath)

End Sub


'***********************************************************************************************************************************************
' * DB�݌v���p�V�[�g�ǉ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function DataBase_AddSheet()

'  On Error GoTo DataBase_AddSheetError:

  '  ���͗p�{�b�N�X�̕\��
  With TableInfoBox
    .StartUpPosition = 0
    .Top = Application.Top + (ActiveWindow.Width / 4)
    .Left = Application.Left + (ActiveWindow.Height / 2)
  End With
  TableInfoBox.Show

  If InputTableName <> "" And InputTableIDa <> "" Then
    Library_StartScript
    Sheets("�R�s�[�p").Copy After:=Worksheets(Worksheets.count)
    ActiveWorkbook.Sheets(Worksheets.count).Tab.ColorIndex = -4142
    ActiveWorkbook.Sheets(Worksheets.count).Name = InputTableIDa

    Range("D5").Value = InputTableName
    Range("H5").Value = InputTableIDa

'    Call DataBase_Reset
    Sheets(InputTableIDa).Select

    Library_UnsetLineColor ("B9:U" & Cells(Rows.count, 2).End(xlUp).Row)
    Call Library_SetLineColor("B9:U" & Cells(Rows.count, 2).End(xlUp).Row, False, RGB(255, 255, 155))

    Library_EndScript
  End If

  Range("C9").Select
  Exit Function

'---------------------------------------------------------------------------------------
'�G���[�������̏���
'---------------------------------------------------------------------------------------
DataBase_AddSheetError:

  Call Library_ErrorHandle(Err.Number, Err.Description)


End Function

'***************************************************************************************************************************************************
' * DB�݌v���pSQL����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***************************************************************************************************************************************************
Function DataBase_MakeSQL(SetDisplyAlertFlg As Boolean)

  Dim ObjADODB As Object

  Dim WriteBuff As String
  Dim WriteBuffTmp As String
  Dim Space As String
  Dim DefaultRowLine As Integer
  Dim rowLine As Integer

  Dim DB_ColumnName As String
  Dim DB_ColumnType As String
  Dim DB_ColumnDigit As String
  Dim DB_ColumnDefValue As String
  Dim DB_NotNull As Integer

  Dim objPrimaryKey As New Dictionary
  Dim objIndex1 As New Dictionary
  Dim objIndex2 As New Dictionary
  Dim objIndex3 As New Dictionary
  Dim objIndex4 As New Dictionary

  Dim DB_ColumnNameLength As Long
  Dim DB_ColumnDefValueLength As Long

  Dim arryobjIndex1() As String
  Dim arryobjIndex2() As String
  Dim arryobjIndex3() As String
  Dim arryobjIndex4() As String
  Dim WriteBuffIndex As String
  Dim WriteBuffTrigger As String
  Dim endLine As Integer

  Dim tableID As String
  Dim tableName As String
  Dim columnComment As String

  Dim Author As String

  On Error GoTo Oracle_MakeDDLError:

  DataBase_Init

  ' �t�@�C���̕ۑ��f�B���N�g���̎w��
  If (SaveDirPath = "") Then
    SaveDirPath = Library_GetDirPath(ActiveWorkbook.Path)
  End If
  If (SaveDirPath = "") Then
      Exit Function
  End If
  Worksheets("�ݒ�").Range("B12") = SaveDirPath

  If Range("B2").Value = "�r���[" Then
    Exit Function
  End If


  '�����ݒ�
  Space = "                                                                        "
  DataBase_Init

  Set ObjADODB = CreateObject("ADODB.Stream")
  DefaultRowLine = 9
  rowLine = DefaultRowLine

  ' �e�[�u��ID�擾
  tableID = Range("H5").Value
  tableName = Range("D5").Value
  columnComment = ""

  '�I�u�W�F�N�g�ɕۑ�����f�[�^�̎�ނ𕶎���^�Ɏw�肷��
  ObjADODB.Type = 2

  '������^�̃I�u�W�F�N�g�̕����R�[�h���w�肷��
  ObjADODB.Charset = CharacterSet        '"UTF-8"
  ObjADODB.LineSeparator = LineSeparator       '-1: CRLF, 10: LF, 13: CR

  '�I�u�W�F�N�g�̃C���X�^���X���쐬
  ObjADODB.Open

  If Range("Q3").Value <> "" Then
   Author = Range("Q3").Value
  Else
    Author = Range("Q2").Value
  End If

  ObjADODB.WriteText "-- ****************************************************************", 1
  ObjADODB.WriteText "-- * @Author      : " & Author, 1
  ObjADODB.WriteText "-- * @Create Date : " & Range("T2").Value, 1
  ObjADODB.WriteText "-- * @Edit   Date : " & Range("T3").Value, 1
  ObjADODB.WriteText "-- * @Description : " + tableName + "[" + tableID + "]", 1
  ObjADODB.WriteText "-- * @version     : $Id: $", 1
  ObjADODB.WriteText "-- ****************************************************************", 1

  Select Case DBMS
    Case "PostgreSQL"

    Case "MySQL"

    Case "Oracle"
'      ObjADODB.WriteText "DECLARE", 1
'      ObjADODB.WriteText "l_exists INTEGER;", 1
'      ObjADODB.WriteText "BEGIN", 1
'      ObjADODB.WriteText "  SELECT COUNT(*) INTO l_exists FROM USER_TABLES where table_name= '" & TableID + "' AND ROWNUM = 1;", 1
'      ObjADODB.WriteText "  If l_exists = 1 Then", 1
'      ObjADODB.WriteText "    DROP TABLE " & TableID + " CASCADE CONSTRAINTS PURGE;", 1
'      ObjADODB.WriteText "  END IF;", 1
'      ObjADODB.WriteText "END;", 1


      ObjADODB.WriteText LineBreakCode & "-- DROP TABLE " & tableID + " CASCADE CONSTRAINTS PURGE;" & LineBreakCode, 1

    Case "SQLServer"
      '�e�[�u����`�o��
      ObjADODB.WriteText "USE [" & DBName & "]", 1
      ObjADODB.WriteText "GO", 1
      ObjADODB.WriteText "", 1

      '�e�[�u�������݂��Ă���΍폜
      ObjADODB.WriteText "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[" + DBScheme & "].[" & tableID + "]') AND type in (N'U'))", 1
      ObjADODB.WriteText "DROP TABLE [" + DBScheme & "].[" & tableID + "]", 1
      ObjADODB.WriteText "GO", 1

      ObjADODB.WriteText "SET ANSI_NULLS ON", 1
      ObjADODB.WriteText "GO", 1
      ObjADODB.WriteText "SET QUOTED_IDENTIFIER ON", 1
      ObjADODB.WriteText "GO", 1

  End Select

  ObjADODB.WriteText "CREATE TABLE " & tableID + " (", 1

  '�J�������̍ő啶��������
  Dim MaxColumnLength As Long
  MaxColumnLength = 0
  Do While (Range("E" & rowLine).Value <> "")
    If Range("E" & rowLine) <> "" And MaxColumnLength < Len(Range("E" & rowLine)) Then
      MaxColumnLength = Len(Range("E" & rowLine))
    End If
    rowLine = rowLine + 1
  Loop
  rowLine = DefaultRowLine

  Do While (Range("E" & rowLine).Value <> "")
    If Range("G" & rowLine) <> "" Then
      '�J�������擾
      DB_ColumnName = Library_GetFixlng(Range("E" & rowLine) & Space, MaxColumnLength + 4)

      '�f�[�^�^���擾
      DB_ColumnType = Range("G" & rowLine)

      '�����擾
      DB_ColumnDigit = Range("H" & rowLine)

      'NOT NULL�l�擾
      DB_NotNull = Range("O" & rowLine)

      '�����l�擾
      DB_ColumnDefValue = Range("P" & rowLine)

      '�f�[�^�^�Z�b�g
      Select Case DB_ColumnType
        Case "IDENTITY"
          WriteBuffTmp = "int IDENTITY"
        Case Else
          WriteBuffTmp = DB_ColumnType
      End Select

      '�����̐ݒ�L���𔻒�
      If DB_ColumnDigit <> "" Then
        WriteBuffTmp = WriteBuffTmp + "(" + CStr(DB_ColumnDigit) + ")"
      End If
      WriteBuff = "  " & DB_ColumnName & Library_GetFixlng(WriteBuffTmp & Space, 17)


      'NOT NULL�̐ݒ�L���𔻒�---------------------------------------------------------------------------------------------------------------
      If DB_NotNull = 1 Then
        WriteBuff = WriteBuff + "  NOT NULL"
      Else
        WriteBuff = WriteBuff + "          "
      End If

      ' �����l�ݒ�----------------------------------------------------------------------------------------------------------------------------
      If DB_ColumnDefValue <> "" Then
        WriteBuff = WriteBuff + "  DEFAULT " + DB_ColumnDefValue
      End If


      '���ږ����R�����g�Őݒ�
      If Range("C" & rowLine) <> "" Then
        WriteBuff = Library_GetFixlng(WriteBuff & Space, 100)
        WriteBuff = WriteBuff + "COMMENT '" + Range("C" & rowLine) & "'"
      End If

      If Range("E" & rowLine + 1) <> "" Then
        WriteBuff = WriteBuff + ","
      End If

      'PRIMARY KEY INDEX�w��̍��ڐݒ�---------------------------------------------------------------------------------------------------------
      If Range("J" & rowLine) <> "" Then
        objPrimaryKey.Add Range("J" & rowLine), Range("E" & rowLine)
      End If

      If Range("K" & rowLine) <> "" Then
        objIndex1.Add Range("K" & rowLine), Range("E" & rowLine)
      End If

      If Range("L" & rowLine) <> "" Then
        objIndex2.Add Range("L" & rowLine), Range("E" & rowLine)
      End If

      If Range("M" & rowLine) <> "" Then
        objIndex3.Add Range("M" & rowLine), Range("E" & rowLine)
      End If

      If Range("N" & rowLine) <> "" Then
        objIndex4.Add Range("N" & rowLine), Range("E" & rowLine)
      End If

      '�ҏW����1���ڕ����o��
      ObjADODB.WriteText WriteBuff, 1
    End If

    '�J�����̃R�����g
    If Range("C" & rowLine) <> "" Then
      Select Case DBMS
        Case "PostgreSQL"
          columnComment = columnComment & "COMMENT ON COLUMN " & tableID & "." & DB_ColumnName & " IS '" & Range("C" & rowLine) & "';" & LineBreakCode
        Case "MySQL"

        Case "Oracle"
          columnComment = columnComment & "COMMENT ON COLUMN " & tableID & "." & DB_ColumnName & " IS '" & Range("C" & rowLine) & "';" & LineBreakCode

        Case "SQLServer"
          columnComment = columnComment & "IF NOT EXISTS (SELECT * FROM ::fn_listextendedproperty(N'MS_Description' , N'SCHEMA',N'" & DBScheme & "', N'TABLE',N'" & tableID & "', N'COLUMN',N'" & DB_ColumnName & "'))" & LineBreakCode
          columnComment = columnComment & "EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'" & Range("C" & rowLine) & "' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'" & tableID & "', @level2type=N'COLUMN',@level2name=N'" & DB_ColumnName & "'" & LineBreakCode
          columnComment = columnComment & "GO" & LineBreakCode

      End Select
    End If
    rowLine = rowLine + 1
  Loop

  'PRIMARY KEY�̐ݒ�--------------------------------------------------------------------------------------------------------------------------
  If objPrimaryKey.count > 0 Then
    Dim arryPrimaryKey() As String
    Dim IndexTableSpace As String
    Dim IndexName As String

    ReDim arryPrimaryKey(1 To objPrimaryKey.count)
    endLine = Cells(Rows.count, 2).End(xlUp).Row - 10

    If Range("C" & endLine) = "" Then
      IndexName = "PK_" & tableID
    Else
      IndexName = Range("C" & endLine)
    End If

    If Range("H" & endLine) = "" Then
      IndexTableSpace = Range("T5")
    Else
      IndexTableSpace = Range("H" & endLine)
    End If

    For Each Var In objPrimaryKey
      arryPrimaryKey(Var) = objPrimaryKey.Item(Var)
    Next Var

    For i = 1 To UBound(arryPrimaryKey)
      Select Case DBMS
        Case "PostgreSQL"

        Case "MySQL"

        Case "Oracle"
          If i = UBound(arryPrimaryKey) Then
            PrimaryKeyNames = PrimaryKeyNames & arryPrimaryKey(i)
          Else
            PrimaryKeyNames = PrimaryKeyNames & arryPrimaryKey(i) & ","
          End If

        Case "SQLServer"
          If i = UBound(arryPrimaryKey) Then
            PrimaryKeyNames = PrimaryKeyNames & "    [" & arryPrimaryKey(i) & "] ASC"
          Else
            PrimaryKeyNames = PrimaryKeyNames & "    [" & arryPrimaryKey(i) & "] ASC," & LineBreakCode
          End If

      End Select
    Next i

      Select Case DBMS
        Case "PostgreSQL"
            ObjADODB.WriteText ");" & LineBreakCode, 1

        Case "MySQL"
            ObjADODB.WriteText ")" & LineBreakCode, 1

        Case "Oracle"
          WriteBuff = LineBreakCode & "  CONSTRAINT " & IndexName & " PRIMARY KEY ("
          ObjADODB.WriteText WriteBuff, 1
          ObjADODB.WriteText "    " & PrimaryKeyNames, 1

          If IndexTableSpace <> "" Then
            ObjADODB.WriteText "  ) USING INDEX TABLESPACE " & IndexTableSpace, 1
          Else
            ObjADODB.WriteText "  ) ", 1
          End If

        Case "SQLServer"
          WriteBuff = LineBreakCode & "  CONSTRAINT [PK_" & tableID & "] PRIMARY KEY CLUSTERED ("
          ObjADODB.WriteText WriteBuff, 1
          ObjADODB.WriteText PrimaryKeyNames, 1

          ObjADODB.WriteText "  )" & LineBreakCode & "  WITH (", 1
          ObjADODB.WriteText "    PAD_INDEX                 = OFF,", 1
          ObjADODB.WriteText "    STATISTICS_NORECOMPUTE    = OFF,", 1
          ObjADODB.WriteText "    IGNORE_DUP_KEY            = OFF,", 1
          ObjADODB.WriteText "    ALLOW_ROW_LOCKS           = ON,", 1
          ObjADODB.WriteText "    ALLOW_PAGE_LOCKS          = ON", 1

          ObjADODB.WriteText "  ) ON [PRIMARY]", 1
          ObjADODB.WriteText ") ON [PRIMARY]", 1
          ObjADODB.WriteText LineBreakCode & "GO" & LineBreakCode & LineBreakCode, 1
      End Select
  End If

  ' Table�̏I���
  Select Case DBMS
    Case "PostgreSQL"

    Case "MySQL"

    Case "Oracle"
      If Range("T5") = "" Then
        ObjADODB.WriteText ");" & LineBreakCode & LineBreakCode, 1
      Else
        ObjADODB.WriteText ") TABLESPACE " & Range("T5") & ";" & LineBreakCode & LineBreakCode, 1
      End If

    Case "SQLServer"
      ObjADODB.WriteText ")", 1
      ObjADODB.WriteText LineBreakCode & "GO" & LineBreakCode & LineBreakCode, 1
  End Select

  'INDEX1�̒�`���o��-----------------------------------------------------------------------------------------------------------------------
  endLine = Cells(Rows.count, 2).End(xlUp).Row - 9

  If Range("C" & endLine) <> "" Then
    ObjADODB.WriteText "-- *******************************************************************", 1

    ReDim arryobjIndex1(1 To objIndex1.count)
    WriteBuffIndex = ""

    For Each Var In objIndex1
      arryobjIndex1(Var) = objIndex1.Item(Var)
    Next Var
    For i = 1 To UBound(arryobjIndex1)
      If i = UBound(arryobjIndex1) Then
        WriteBuffIndex = WriteBuffIndex & arryobjIndex1(i)
      Else
        WriteBuffIndex = WriteBuffIndex & arryobjIndex1(i) & ","
      End If
    Next i

    Select Case DBMS
      Case "PostgreSQL"

      Case "MySQL"

      Case "Oracle"
        WriteBuff = "CREATE INDEX " + Range("C" & endLine) + " ON " & Range("H" & endLine) & "." & tableID & " ("
        ObjADODB.WriteText WriteBuff, 1
        ObjADODB.WriteText WriteBuffIndex, 1

        If Range("H" & endLine) = "" Then
          ObjADODB.WriteText ") TABLESPACE " & Range("T5") & LineBreakCode, 1
        Else
          ObjADODB.WriteText ") TABLESPACE " & Range("H" & endLine) & LineBreakCode, 1
        End If

      Case "SQLServer"
        WriteBuff = "CREATE " & Range("E" & endLine) & " INDEX [" + Range("C" & endLine) + "] ON [" & Range("Q6") & "].[" & tableID & "] ("
        ObjADODB.WriteText WriteBuff, 1
        ObjADODB.WriteText WriteBuffIndex, 1
        ObjADODB.WriteText ")" & LineBreakCode & "GO", 1
    End Select
  End If

  'INDEX2�̒�`���o��-----------------------------------------------------------------------------------------------------------------------
  endLine = Cells(Rows.count, 2).End(xlUp).Row - 8
  If Range("C" & endLine) <> "" Then
    ObjADODB.WriteText LineBreakCode & LineBreakCode, 1
    ObjADODB.WriteText "-- *******************************************************************", 1
    ReDim arryobjIndex2(1 To objIndex2.count)
    WriteBuffIndex = ""

    For Each Var In objIndex2
      arryobjIndex2(Var) = objIndex2.Item(Var)
    Next Var
    For i = 1 To UBound(arryobjIndex2)
      If i = UBound(arryobjIndex2) Then
        WriteBuffIndex = WriteBuffIndex & arryobjIndex2(i)
      Else
        WriteBuffIndex = WriteBuffIndex & arryobjIndex2(i) & " ," & LineBreakCode
      End If
    Next i

    Select Case DBMS
      Case "PostgreSQL"

      Case "MySQL"

      Case "Oracle"
      WriteBuff = "CREATE INDEX " + Range("C" & endLine) + " ON " & Range("H" & endLine) & "." & tableID & " ("
      ObjADODB.WriteText WriteBuff, 1
      ObjADODB.WriteText WriteBuffIndex, 1
      ObjADODB.WriteText ");" & LineBreakCode, 1

      Case "SQLServer"
      WriteBuff = "CREATE " & Range("E" & endLine) & " INDEX [" + Range("C" & endLine) + "] ON [" & Range("Q6") & "].[" & tableID & "] ("
      ObjADODB.WriteText WriteBuff, 1
      ObjADODB.WriteText WriteBuffIndex, 1
      ObjADODB.WriteText ")" & LineBreakCode & "GO", 1
    End Select
  End If

  'INDEX3�̒�`���o��-----------------------------------------------------------------------------------------------------------------------
  endLine = Cells(Rows.count, 2).End(xlUp).Row - 7
  If Range("C" & endLine) <> "" Then
    ObjADODB.WriteText LineBreakCode & LineBreakCode, 1
    ObjADODB.WriteText "-- *******************************************************************", 1
    ReDim arryobjIndex3(1 To objIndex3.count)
    WriteBuffIndex = ""

    For Each Var In objIndex3
      arryobjIndex3(Var) = objIndex3.Item(Var)
    Next Var
    For i = 1 To UBound(arryobjIndex3)
      If i = UBound(arryobjIndex3) Then
        WriteBuffIndex = WriteBuffIndex & arryobjIndex3(i)
      Else
        WriteBuffIndex = WriteBuffIndex & arryobjIndex3(i) & "," & LineBreakCode
      End If
    Next i

    Select Case DBMS
      Case "PostgreSQL"

      Case "MySQL"

      Case "Oracle"
      WriteBuff = "CREATE INDEX " + Range("C" & endLine) + " ON " & Range("H" & endLine) & "." & tableID & " ("
      ObjADODB.WriteText WriteBuff, 1
      ObjADODB.WriteText WriteBuffIndex, 1
      ObjADODB.WriteText ");" & LineBreakCode, 1

      Case "SQLServer"
      WriteBuff = "CREATE " & Range("E" & endLine) & " INDEX [" + Range("C" & endLine) + "] ON [" & Range("Q6") & "].[" & tableID & "] ("
      ObjADODB.WriteText WriteBuff, 1
      ObjADODB.WriteText WriteBuffIndex, 1
      ObjADODB.WriteText ")" & LineBreakCode & "GO", 1
    End Select
  End If

  'INDEX4�̒�`���o��-----------------------------------------------------------------------------------------------------------------------
  endLine = Cells(Rows.count, 2).End(xlUp).Row - 6
  If Range("C" & endLine) <> "" Then
    ObjADODB.WriteText LineBreakCode & LineBreakCode, 1
    ObjADODB.WriteText "-- *******************************************************************", 1
    ReDim arryobjIndex4(1 To objIndex4.count)
    WriteBuffIndex = ""

    For Each Var In objIndex4
      arryobjIndex4(Var) = objIndex4.Item(Var)
    Next Var
    For i = 1 To UBound(arryobjIndex4)
      If i = UBound(arryobjIndex4) Then
        WriteBuffIndex = WriteBuffIndex & arryobjIndex4(i)
      Else
        WriteBuffIndex = WriteBuffIndex & arryobjIndex4(i) & "," & LineBreakCode
      End If
    Next i

    Select Case DBMS
      Case "PostgreSQL"

      Case "MySQL"

      Case "Oracle"
      WriteBuff = "CREATE INDEX " + Range("C" & endLine) + " ON " & Range("H" & endLine) & "." & tableID & " ("
      ObjADODB.WriteText WriteBuff, 1
      ObjADODB.WriteText WriteBuffIndex, 1
      ObjADODB.WriteText ");" & LineBreakCode, 1

      Case "SQLServer"
      WriteBuff = "CREATE " & Range("E" & endLine) & " INDEX [" + Range("C" & endLine) + "] ON [" & Range("Q6") & "].[" & tableID & "] ("
      ObjADODB.WriteText WriteBuff, 1
      ObjADODB.WriteText WriteBuffIndex, 1
      ObjADODB.WriteText ")" & LineBreakCode & "GO", 1
    End Select
  End If



  'Trigger1�̒�`���o��-----------------------------------------------------------------------------------------------------------------------
  endLine = Cells(Rows.count, 2).End(xlUp).Row - 3
  If Range("C" & endLine) <> "" Then
    ObjADODB.WriteText LineBreakCode & LineBreakCode, 1
    ObjADODB.WriteText "-- *******************************************************************", 1
    WriteBuffTrigger = ""

    WriteBuffTrigger = "CREATE TRIGGER [" & Range("Q6") & "].[" + Range("C" & endLine) + "] ON [" & Range("Q6") & "].[" & tableID & "]" & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "FOR " & Range("E" & endLine) & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "NOT FOR REPLICATION" & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "AS" & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "BEGIN" & LineBreakCode & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "SET NOCOUNT ON" & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & Range("H" & endLine) & LineBreakCode & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "END"

    ObjADODB.WriteText WriteBuffTrigger, 1
    ObjADODB.WriteText LineBreakCode & "GO", 1
  End If


  'Trigger2�̒�`���o��-----------------------------------------------------------------------------------------------------------------------
  endLine = Cells(Rows.count, 2).End(xlUp).Row - 2
  If Range("C" & endLine) <> "" Then
    ObjADODB.WriteText LineBreakCode & LineBreakCode, 1
    ObjADODB.WriteText "-- *******************************************************************", 1
    WriteBuffTrigger = ""

    WriteBuffTrigger = "CREATE TRIGGER [" & Range("Q6") & "].[" + Range("C" & endLine) + "] ON [" & Range("Q6") & "].[" & tableID & "]" & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "FOR " & Range("E" & endLine) & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "NOT FOR REPLICATION" & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "AS" & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "BEGIN" & LineBreakCode & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "SET NOCOUNT ON" & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & Range("H" & endLine) & LineBreakCode & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "END"

    ObjADODB.WriteText WriteBuffTrigger, 1
    ObjADODB.WriteText LineBreakCode & "GO", 1
  End If

  'Trigger3�̒�`���o��-----------------------------------------------------------------------------------------------------------------------
  endLine = Cells(Rows.count, 2).End(xlUp).Row - 1
  If Range("C" & endLine) <> "" Then
    ObjADODB.WriteText LineBreakCode & LineBreakCode, 1
    ObjADODB.WriteText "-- *******************************************************************", 1
    WriteBuffTrigger = ""

    WriteBuffTrigger = "CREATE TRIGGER [" & Range("Q6") & "].[" + Range("C" & endLine) + "] ON [" & Range("Q6") & "].[" & tableID & "]" & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "FOR " & Range("E" & endLine) & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "NOT FOR REPLICATION" & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "AS" & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "BEGIN" & LineBreakCode & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "SET NOCOUNT ON" & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & Range("H" & endLine) & LineBreakCode & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "END"

    ObjADODB.WriteText WriteBuffTrigger, 1
    ObjADODB.WriteText LineBreakCode & "GO", 1
  End If


  'Trigger4�̒�`���o��-----------------------------------------------------------------------------------------------------------------------
  endLine = Cells(Rows.count, 2).End(xlUp).Row
  If Range("C" & endLine) <> "" Then
    ObjADODB.WriteText LineBreakCode & LineBreakCode, 1
    ObjADODB.WriteText "-- *******************************************************************", 1
    WriteBuffTrigger = ""

    WriteBuffTrigger = "CREATE TRIGGER [" & Range("Q6") & "].[" + Range("C" & endLine) + "] ON [" & Range("Q6") & "].[" & tableID & "]" & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "FOR " & Range("E" & endLine) & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "NOT FOR REPLICATION" & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "AS" & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "BEGIN" & LineBreakCode & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "SET NOCOUNT ON" & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & Range("H" & endLine) & LineBreakCode & LineBreakCode
    WriteBuffTrigger = WriteBuffTrigger & "END"

    ObjADODB.WriteText WriteBuffTrigger, 1
    ObjADODB.WriteText LineBreakCode & "GO", 1
  End If


  Select Case DBMS
    Case "PostgreSQL"
      ObjADODB.WriteText "COMMENT ON TABLE  " & Library_GetFixlng(tableID & Space, MaxColumnLength + Len(tableID) + 5) & " IS '" & tableName & "';", 1

    Case "MySQL"
      ObjADODB.WriteText "COMMENT ='" & tableName & "';", 1

    Case "Oracle"
      ObjADODB.WriteText "COMMENT ON TABLE  " & Library_GetFixlng(tableID & Space, MaxColumnLength + Len(tableID) + 5) & " IS '" & tableName & "';", 1

    Case "SQLServer"
      WriteBuff = "CREATE " & Range("E" & endLine) & " INDEX [" + Range("C" & endLine) + "] ON [" & Range("Q6") & "].[" & tableID & "] ("
      ObjADODB.WriteText WriteBuff, 1
      ObjADODB.WriteText WriteBuffIndex, 1
      ObjADODB.WriteText ")" & LineBreakCode & "GO", 1
  End Select











  ObjADODB.WriteText columnComment, 1

  'UTF-8��BOM�폜
  If CharacterSet = "UTF-8" Then
    ObjADODB.Position = 0
    ObjADODB.Type = adTypeBinary
    ObjADODB.Position = 3
    byteData = ObjADODB.Read
    ObjADODB.Close
    ObjADODB.Open
    ObjADODB.Write byteData
  End If

  '�I�u�W�F�N�g�̓��e���t�@�C���ɕۑ�
  ObjADODB.SaveToFile (SaveDirPath & "\CREATE_TABLE_" & tableID & ".sql"), 2

  '�I�u�W�F�N�g�����
  ObjADODB.Close

  '����������I�u�W�F�N�g���폜����
  Set ObjADODB = Nothing

  If SetDisplyAlertFlg = True Then
    MsgBox (Range("D5") + "�p�X�N���v�g�t�@�C���̍쐬���������܂����B" & LineBreakCode & SaveDirPath)

    Call Shell("Explorer.exe /select, " & SaveDirPath & "\CREATE_TABLE_" & tableID & ".sql", vbNormalFocus)
  End If
  Exit Function

Oracle_MakeDDLError:

  Call Library_ErrorHandle(Err.Number, Err.Description & LineBreakCode & _
        "SQL�����Ɏ��s���܂���" & tableID & "  " & rowLine - DefaultRowLine + 1 & "�s��")

End Function


'***************************************************************************************************************************************************
' * DB�e�[�u�����擾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***************************************************************************************************************************************************
Function DataBase_GetTableList()

  Dim QueryString As String

  Dim tableName As String
  Dim Comment As Variant
  Dim newSheetName As String

'  Dim DBRecordsetCount As Integer
'  Dim DBCon As ADODB.Connection
  Dim DBRecordset As ADODB.Recordset

  On Error GoTo GetTableList_Error
  ProgressBar_ProgShowStart

  DataBase_Init
  DBRecordsetCount = 1

  'ADODB.Connection�������ADB�ɐڑ�
  Set dbCon = New ADODB.Connection
  dbCon.Open ConnectionString

  'SQL�� -------------------------------------------------------------------------------------------------------------------------------------
  Select Case DBMS
    Case "PostgreSQL"
'      PostgreSQL_MakeDDL

    Case "MySQL"
      QueryString = "SELECT TABLE_NAME as TableName, TABLE_COMMENT as Comments,'' as  TableSpaceName from information_schema.TABLES WHERE TABLE_SCHEMA = DATABASE();"

    Case "Oracle"
      If DBTableSpace = "" Then
        QueryString = "select" & LineBreakCode
        QueryString = QueryString & "  UT.table_name TableName," & LineBreakCode
        QueryString = QueryString & "  UTC.Comments," & LineBreakCode
        QueryString = QueryString & "  UT.tablespace_name TableSpaceName" & LineBreakCode
        QueryString = QueryString & "from USER_TABLES UT left join USER_TAB_COMMENTS UTC on UT.table_name =UTC.table_name" & LineBreakCode
        QueryString = QueryString & "where UT.tablespace_name is not null " & LineBreakCode
      Else
        QueryString = "select" & LineBreakCode
        QueryString = QueryString & "  UT.table_name TableName," & LineBreakCode
        QueryString = QueryString & "  UTC.Comments," & LineBreakCode
        QueryString = QueryString & "  UT.tablespace_name TableSpaceName" & LineBreakCode
        QueryString = QueryString & "from USER_TABLES UT left join USER_TAB_COMMENTS UTC on UT.table_name =UTC.table_name" & LineBreakCode
        QueryString = QueryString & "where UT.tablespace_name='" & DBTableSpace & "';" & LineBreakCode
      End If

      QueryString = QueryString & " order by UT.table_name" & LineBreakCode

    Case "SQLServer"
      QueryString = "select table_name TableName,'' Comments from USER_TABLES;"
  End Select

  '�ʂɃe�[�u���ꗗ���擾�������Ƃ��p
  If LocalQueryString <> "" Then
    QueryString = LocalQueryString
  End If

  Set DBRecordset = New ADODB.Recordset
  DBRecordset.Open QueryString, dbCon, adOpenKeyset, adLockReadOnly

  Sheets("TBL���X�g").Range("W2").Value = "�e�[�u�����擾SQL"
  Sheets("TBL���X�g").Range("X2").Value = QueryString

  Do Until DBRecordset.EOF

    tableName = DBRecordset.Fields("TableName").Value

    newSheetName = tableName
    newSheetName = Left(newSheetName, 30)

    If IsNull(DBRecordset.Fields("TableSpaceName").Value) Or DBRecordset.Fields("TableSpaceName").Value = "" Then
      TableSpaceName = ""
    Else
      TableSpaceName = DBRecordset.Fields("TableSpaceName").Value
    End If

    If Library_CheckExcludeSheet(tableName, 9) = False Then
      GoTo Continue
    End If

    If Library_ChkSheetName(newSheetName) = False Then
      Sheets("�R�s�[�p").Copy After:=Worksheets(Worksheets.count)
      ActiveWorkbook.Sheets(Worksheets.count).Tab.ColorIndex = -4142
      ActiveWorkbook.Sheets(Worksheets.count).Name = newSheetName
    Else
'      GoTo Continue
    End If


    Sheets(newSheetName).Select
    Range("D5").Value = Comment
    Range("H5").Value = tableName
    Range("T5").Value = tableName

    '�e�[�u����ʐݒ�
    If InStr(Comment, "���[�N") Or InStr(tableName, "WRK") Then
      Range("B2").Value = "���[�N�e�[�u��"

    ElseIf InStr(Comment, "�g����") Or InStr(tableName, "TRN") Then
      Range("B2").Value = "�g�����U�N�V�����e�[�u��"
    Else
      Range("B2").Value = "�}�X�^�[�e�[�u��"
    End If

    '�J�������擾
    DataBase_GetColumn (False)

Continue:
    '���̃��R�[�h
    DBRecordsetCount = DBRecordsetCount + 1
    DBRecordset.MoveNext

  Loop

  Set DBRecordset = Nothing

  'SQL�� -------------------------------------------------------------------------------------------------------------------------------------
'  QueryString = "SELECT * FROM sysobjects WHERE xtype = 'V ' order by xtype, name;"
'
'  Set DBRecordset = New ADODB.Recordset
'  DBRecordset.Open QueryString, DBCon, adOpenKeyset, adLockReadOnly
'
'  Do Until DBRecordset.EOF
'
'    TableName = DBRecordset.Fields("Name").Value
'    Comment = DBRecordset.Fields("Comment").Value
'
'    If Library_ChkSheetName(Left(TableName, 30)) = False Then
'      Sheets("�R�s�[�p").Copy After:=Worksheets(Worksheets.Count)
'      ActiveWorkbook.Sheets(Worksheets.Count).Tab.ColorIndex = -4142
'      ActiveWorkbook.Sheets(Worksheets.Count).Name = Left(TableName, 30)
'      Range("D5").Value = Comment
'      Range("H5").Value = TableName
'
'      Call DataBase_Reset
'      Sheets(Left(TableName, 30)).Select
'
'      '�J�������擾
'      Range("B2").Value = "�r���["
'      Oracle_GetViewColumnList
'
'    End If
'    '���̃��R�[�h
'    DBRecordsetCount = DBRecordsetCount + 1
'    DBRecordset.MoveNext
'  Loop

  'DB�N���[�Y
  dbCon.Close
  Set DBRecordset = Nothing


  '�s������-----------------------------------------------------------------------------------------------------------------------------------
  Dim endLine As Integer
  ProgressBar_ProgShowCount "�e�[�u�����X�g�������E�E�E", 5, 100, "�e�[�u�����X�g������"

  DataBase_MakeTableList
  DataBase_Reset

  ProgressBar_ProgShowClose

  Exit Function

'---------------------------------------------------------------------------------------------------------------------------------------------
'�G���[�������̏���
'---------------------------------------------------------------------------------------------------------------------------------------------
GetTableList_Error:
  Call Library_ErrorHandle(Err.Number, Err.Description)
  ProgressBar_ProgShowClose

  If (dbCon.State And adStateOpen) = adStateOpen Then
    'DB�N���[�Y
    dbCon.Close
  End If
  ConnectionString = ""
  Set DBRecordset = Nothing
  Set GetTableList_Result = Nothing
  Set GetTableList_Con = Nothing

End Function


'***************************************************************************************************************************************************
' * DB�J�������擾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***************************************************************************************************************************************************
Function DataBase_GetColumn(SetDisplyProgressBarFlg As Boolean)

  Dim QueryString As String
  Dim SelectString As String
  Dim RunQueryString As String

  Dim tableName As String
  Dim NowLine As Long
  Dim DelLine As Long

  Dim DBConGetColumn As ADODB.Connection
  Dim DBRecordset As ADODB.Recordset
'  Dim DBRecordsetCount As Integer

  Dim columnName As String
  Dim dataType As String
  Dim maxLength As String
  Dim PrimaryKeyIndex As String
  Dim RefTableName As String
  Dim RefColumnName As String
  Dim isNullable As String
  Dim is_identity As String
  Dim EPvalue As String
  Dim Comments As String
  Dim RowCount As Long
  Dim ScaleLength As String

'  On Error GoTo GetColumnList_Error

  DataBase_Init

  tableName = Range("H5")
  NowLine = 9
  DBRecordsetCount = 1

  ' �v���O���X�o�[�̕\���J�n
  If SetDisplyProgressBarFlg Then
    ProgressBar_ProgShowStart
  End If


  If DBConGetColumn Is Nothing Then
    'ADODB.Connection�������ADB�ɐڑ�
    ProgressBar_ProgShowCount tableName & " �ڑ����E�E�E", 5, 100, "DB�ɐڑ�"

    Set DBConGetColumn = New ADODB.Connection
    DBConGetColumn.Open ConnectionString
    ProgressBar_ProgShowCount tableName & " �ڑ����E�E�E", 50, 100, "DB�ɐڑ�"
  End If

  'SQL�� -------------------------------------------------------------------------------------------------------------------------------------
  Select Case DBMS
    Case "PostgreSQL"
'      PostgreSQL_MakeDDL

    Case "MySQL"
      QueryString = "SELECT TABLE_NAME as TableName, TABLE_COMMENT as Comments,'' as  TableSpaceName from "
      QueryString = QueryString & " information_schema.TABLES WHERE TABLE_NAME='" & tableName & "'"

    Case "Oracle"
        QueryString = "select" & LineBreakCode
        QueryString = QueryString & "  UT.table_name TableName," & LineBreakCode
        QueryString = QueryString & "  UTC.Comments," & LineBreakCode
        QueryString = QueryString & "  UT.tablespace_name TableSpaceName" & LineBreakCode
        QueryString = QueryString & "from USER_TABLES UT left join USER_TAB_COMMENTS UTC on UT.table_name =UTC.table_name" & LineBreakCode
        QueryString = QueryString & "where UT.table_name='" & tableName & "'" & LineBreakCode

    Case "SQLServer"
      QueryString = "select table_name TableName,'' Comments,'' as  TableSpaceName from USER_TABLES;"
  End Select

  '�e�[�u�����擾
  Range("W2").Value = "�e�[�u�����擾SQL"
  Range("X2").Value = QueryString
  Range("X2").WrapText = False

  Set DBRecordset = New ADODB.Recordset
  DBRecordset.Open QueryString, DBConGetColumn, adOpenKeyset, adLockReadOnly

  Do Until DBRecordset.EOF

    tableName = DBRecordset.Fields("TableName").Value

    If IsNull(DBRecordset.Fields("Comments").Value) Or DBRecordset.Fields("Comments").Value = "" Then
      Comments = ""
    Else
      Comments = DBRecordset.Fields("Comments").Value
    End If

    If IsNull(DBRecordset.Fields("TableSpaceName").Value) Or DBRecordset.Fields("TableSpaceName").Value = "" Then
      TableSpaceName = ""
    Else
      TableSpaceName = DBRecordset.Fields("TableSpaceName").Value
    End If

    If Comments <> "" Then
      Range("D5").Value = Comments
    End If
    Range("T5").Value = TableSpaceName

    '�e�[�u����ʐݒ�
    If InStr(Comment, "���[�N") Or InStr(tableName, "WRK") Then
      Range("B2").Value = "���[�N�e�[�u��"

    ElseIf InStr(Comment, "�g����") Or InStr(tableName, "TRN") Then
      Range("B2").Value = "�g�����U�N�V�����e�[�u��"
    Else
      Range("B2").Value = "�}�X�^�[�e�[�u��"
    End If

    '���̃��R�[�h
    DBRecordsetCount = DBRecordsetCount + 1
    DBRecordset.MoveNext

  Loop
  Set DBRecordset = Nothing

  'SQL�� -------------------------------------------------------------------------------------------------------------------------------------
  SelectString = ""
  QueryString = ""
  Select Case DBMS
    Case "PostgreSQL"
'      PostgreSQL_MakeDDL

    Case "MySQL"
      SelectString = "SELECT " & LineBreakCode
      SelectString = SelectString & "      COLUMN_NAME                          AS ColumName " & LineBreakCode
      SelectString = SelectString & "    , DATA_TYPE                            AS DataType " & LineBreakCode
      SelectString = SelectString & "    , IFNULL(CHARACTER_MAXIMUM_LENGTH, '') AS Length    " & LineBreakCode
      SelectString = SelectString & "    , ''                                   AS ScaleLength " & LineBreakCode
      SelectString = SelectString & "    , COLUMN_KEY                           AS PrimaryKey " & LineBreakCode
      SelectString = SelectString & "    , IS_NULLABLE                          AS Nullable " & LineBreakCode
      SelectString = SelectString & "    , COLUMN_DEFAULT                       AS ColumnDefault " & LineBreakCode
      SelectString = SelectString & "    , COLUMN_COMMENT                       AS Comments " & LineBreakCode
      QueryString = QueryString & "FROM"
      QueryString = QueryString & " information_schema.Columns c "
      QueryString = QueryString & "WHERE"
      QueryString = QueryString & "     c.table_schema = '" & DBName & "' "
      QueryString = QueryString & " AND c.table_name   = '" & tableName & "' "
      QueryString = QueryString & "ORDER BY"
      QueryString = QueryString & " ordinal_position;"

    Case "Oracle"
      SelectString = "select " & LineBreakCode
      SelectString = SelectString & "    UTC.column_name                                as ColumName," & LineBreakCode
      SelectString = SelectString & "    UTC.data_type                                  as DataType," & LineBreakCode
      SelectString = SelectString & "    NVL(UTC.DATA_PRECISION, CHAR_COL_DECL_LENGTH)  as Length," & LineBreakCode
      SelectString = SelectString & "    UTC.data_scale                                 as ScaleLength," & LineBreakCode
      SelectString = SelectString & "    UCCPkey.position                               as PrimaryKey," & LineBreakCode
      SelectString = SelectString & "    case" & LineBreakCode
      SelectString = SelectString & "      when UTC.nullable ='Y' then 0" & LineBreakCode
      SelectString = SelectString & "      when UTC.nullable ='N' then 1" & LineBreakCode
      SelectString = SelectString & "    end                                            as Nullable," & LineBreakCode
      SelectString = SelectString & "    UCC.COMMENTS                                   as Comments," & LineBreakCode
      SelectString = SelectString & "    UTC.data_default                               as ColumnDefault" & LineBreakCode
      QueryString = QueryString & "  FROM" & LineBreakCode
      QueryString = QueryString & "    USER_TAB_COLUMNS UTC left join USER_COL_COMMENTS UCC on UTC.table_name = UCC.table_name and UTC.column_name = UCC.column_name" & LineBreakCode
      QueryString = QueryString & "    left join USER_CONS_COLUMNS UCCPkey on UTC.table_name = UCCPkey.table_name and UTC.column_name = UCCPkey.column_name and UCCPkey.position is not null" & LineBreakCode
      QueryString = QueryString & "  WHERE UTC.table_name='" & tableName & "'" & LineBreakCode
      QueryString = QueryString & "  ORDER BY UTC.column_id" & LineBreakCode

    Case "SQLServer"
      QueryString = "select table_name TableName,'' Comments from USER_TABLES;"
  End Select

  ProgressBar_ProgShowCount tableName & " �ڑ����E�E�E", 75, 100, "DB�ɐڑ�"

  Set DBRecordset = New ADODB.Recordset

  '�J�������擾
  RunQueryString = "select count(*) as count " & LineBreakCode & QueryString

  Set DBRecordset = New ADODB.Recordset
  DBRecordset.Open RunQueryString, DBConGetColumn, adOpenKeyset, adLockReadOnly
  Do Until DBRecordset.EOF
    RowCount = CLng(DBRecordset.Fields("count").Value)

    '���̃��R�[�h
    DBRecordsetCount = DBRecordsetCount + 1
    DBRecordset.MoveNext
  Loop
  Set DBRecordset = Nothing
  ProgressBar_ProgShowCount tableName & " �ڑ����E�E�E", 90, 100, tableName & " �J�������擾"

  '�J�������擾
  RunQueryString = SelectString & QueryString
  Range("W3").Value = "�J�������擾SQL"
  Range("X3").Value = RunQueryString
  Range("X3").WrapText = False

  Set DBRecordset = New ADODB.Recordset
  DBRecordset.Open RunQueryString, DBConGetColumn, adOpenKeyset, adLockReadOnly
  ProgressBar_ProgShowCount tableName & " �ڑ����E�E�E", 100, 100, tableName & " �J�������擾"

  Do Until DBRecordset.EOF
    If (NowLine >= 109 And Range("B" & NowLine) <> NowLine - 8) Then
      ActiveSheet.Tab.Color = RGB(255, 183, 183)
      Range("D6") = "�s���s�� �J�������F" & RowCount

      If SetDisplyProgressBarFlg Then
        ProgressBar_ProgShowClose
      End If

      Exit Function
    Else
      Range("D6") = ""
      If ActiveSheet.Tab.ColorIndex = 22 Then
        ActiveSheet.Tab.ColorIndex = -4142
      End If
    End If
    ' �v���O���X�o�[�̃J�E���g�ύX�i���݂̃J�E���g�A�S�J�E���g���A���b�Z�[�W�j
    ProgressBar_ProgShowCount tableName & " �ڑ����E�E�E", NowLine - 9, RowCount, " �J�������擾�F" & columnName

    '�����ݒ�l���N���A
    If DBMode <> "Diff" Then
      Range("E" & NowLine & ":P" & NowLine).Select
      Selection.ClearContents
    End If
    '�R�����g(���ږ��Ƃ��ė��p)
    If IsNull(DBRecordset.Fields("Comments").Value) Then
      Comments = ""
    Else
      Comments = DBRecordset.Fields("Comments").Value
    End If
    If Range("C" & NowLine).Value = "" Then
      Range("C" & NowLine).Value = Comments
    End If

    '�J������
    columnName = DBRecordset.Fields("ColumName").Value
    If DBMode = "Diff" Then
      If Range("E" & NowLine).Value <> columnName Then
        If Range("E" & NowLine).Value = "" Then
          Range("E" & NowLine).Style = "�J�����ǉ�"
        Else
          Range("E" & NowLine).AddComment
          Range("E" & NowLine).Comment.Visible = False
          Range("E" & NowLine).Comment.Text Text:=Range("E" & NowLine).Value
          Range("E" & NowLine).Style = "�J�����ύX"
        End If
      End If
    End If
    Range("E" & NowLine).Value = columnName

    '�^
    dataType = DBRecordset.Fields("DataType").Value
    Select Case dataType
      Case "TIMESTAMP(6)"
        dataType = "TIMESTAMP"
    End Select

    If DBMode = "Diff" Then
      If Range("G" & NowLine).Value <> dataType Then
        If Range("G" & NowLine).Value = "" Then
          Range("G" & NowLine).Style = "�J�����ǉ�"
        Else
          Range("G" & NowLine).AddComment
          Range("G" & NowLine).Comment.Visible = False
          Range("G" & NowLine).Comment.Text Text:=Range("G" & NowLine).Value
          Range("G" & NowLine).Style = "�J�����ύX"
        End If
      End If
    End If
    Range("G" & NowLine).Value = dataType

    '����(���x)
    If IsNull(DBRecordset.Fields("ScaleLength").Value) Then
      ScaleLength = ""
    Else
      ScaleLength = DBRecordset.Fields("ScaleLength").Value
    End If

    If IsNull(DBRecordset.Fields("Length").Value) Then
      maxLength = ""
    Else
      maxLength = DBRecordset.Fields("Length").Value
    End If
    Select Case dataType
      Case "numeric", "decimal", "NUMBER"
        maxLength = maxLength & "," & ScaleLength

      Case "int"
'        If MaxLength <> "" Then
'          Range("H" & NowLine).Value = MaxLength
'        End If

      Case "datetime2", "datetime", "tinyint", "bit", "varbinary", "xml", "image", "money", "text", "TIMESTAMP"
        maxLength = ""
'      Case Else
'        Range("H" & NowLine).Value = MaxLength
    End Select
    If DBMode = "Diff" Then
      If Range("H" & NowLine).Value <> maxLength Then
        If Range("H" & NowLine).Value = "" Then
          Range("H" & NowLine).Style = "�J�����ǉ�"
        Else
          Range("H" & NowLine).AddComment
          Range("H" & NowLine).Comment.Visible = False
          Range("H" & NowLine).Comment.Text Text:=" " & Range("H" & NowLine).Value
          Range("H" & NowLine).Style = "�J�����ύX"
        End If
      End If
    End If
    Range("H" & NowLine).Value = maxLength


    '�v���C�}���L�[
    If IsNull(DBRecordset.Fields("PrimaryKey").Value) Then
      PrimaryKeyIndex = ""
    Else
      PrimaryKeyIndex = DBRecordset.Fields("PrimaryKey").Value
      If PrimaryKeyIndex = "PRI" Then
        PrimaryKeyIndex = 1
      ElseIf PrimaryKeyIndex = "UNI" Then
        PrimaryKeyIndex = 2
      ElseIf PrimaryKeyIndex = "MUL" Then
        PrimaryKeyIndex = 3
      End If
    End If
    Range("J" & NowLine).Value = PrimaryKeyIndex

    'NotNULL����
    If IsNull(DBRecordset.Fields("Nullable").Value) Then
      isNullable = ""
    Else
      isNullable = DBRecordset.Fields("Nullable").Value
    End If
    If isNullable = "1" Or isNullable = "No" Or isNullable = "NO" Then
      isNullable = 1

    Else
      isNullable = ""
    End If

    If DBMode = "Diff" Then
      If Range("O" & NowLine).Value <> isNullable Then
        If Range("O" & NowLine).Value = "" Then
          Range("O" & NowLine).Style = "�J�����ǉ�"
        Else
          Range("O" & NowLine).AddComment
          Range("O" & NowLine).Comment.Visible = False
          Range("O" & NowLine).Comment.Text Text:=" " & Range("O" & NowLine).Value
          Range("O" & NowLine).Style = "�J�����ύX"
        End If
      End If
    End If
    Range("O" & NowLine).Value = isNullable

    '�����l
    If IsNull(DBRecordset.Fields("ColumnDefault").Value) Then
      ColumnDefault = ""
    Else
      'Range("P" & NowLine).NumberFormatLocal = "@"
      ColumnDefault = DBRecordset.Fields("ColumnDefault").Value
'      ColumnDefault = Replace(ColumnDefault, "(", "")
'      ColumnDefault = Replace(ColumnDefault, ")", "")
'      ColumnDefault = Replace(ColumnDefault, "'", "")
'      ColumnDefault = Replace(ColumnDefault, "'", "")

      Select Case ColumnDefault
        Case "getdate"
          ColumnDefault = "getdate()"
      End Select

    If DBMode = "Diff" Then
      If Range("P" & NowLine).Value <> ColumnDefault Then
        If Range("P" & NowLine).Value = "" Then
          Range("P" & NowLine).Style = "�J�����ǉ�"
        Else
          Range("P" & NowLine).AddComment
          Range("P" & NowLine).Comment.Visible = False
          Range("P" & NowLine).Comment.Text Text:=" " & Range("P" & NowLine).Value
          Range("P" & NowLine).Style = "�J�����ύX"
        End If
      End If
    End If
    Range("P" & NowLine) = "'" & ColumnDefault

    End If

    '���̃��R�[�h
    NowLine = NowLine + 1
    DBRecordset.MoveNext

  Loop

  '�ŏI�s���}�[�N
  Range("A" & NowLine).Value = "Column"


'-------------------------------------------------------------------------------------------------------------------------------------------
'�C���f�b�N�X���擾
'-------------------------------------------------------------------------------------------------------------------------------------------
  Dim IndexColumnName As String
  Dim IndexName As String
  Dim rowLine As Long
  Dim IndexTableSpace As String
  Dim indexCount As Long
  Dim IndexUniqueness As Long
'  Dim IndexCount As Long


  '�C���f�b�N�X��ݒ肷��s�����擾
  For i = 20 To 1000
    If Range("A" & i) = "Index" Then
      NowLine = i
      Exit For
    End If
  Next

  indexCount = 0
  Set DBRecordset = Nothing

  Select Case DBMS
    Case "PostgreSQL"
      QueryString = "EXEC sp_MShelpindex " & tableName

    Case "MySQL"
      QueryString = "SHOW INDEX FROM " & tableName & ";"

    Case "Oracle"
      QueryString = "SELECT" & LineBreakCode
      QueryString = QueryString + "UIC.index_name IndexName" & LineBreakCode
      QueryString = QueryString + "  , UIC.column_name ColumnName" & LineBreakCode
      QueryString = QueryString + "  , UIC.Column_Position ColumnPosition" & LineBreakCode
      QueryString = QueryString + "  , UI.tablespace_name  TableSpace" & LineBreakCode
      QueryString = QueryString + "  , case" & LineBreakCode
      QueryString = QueryString + "      when UI.uniqueness ='UNIQUE' then 0" & LineBreakCode
      QueryString = QueryString + "      when UI.uniqueness ='NONUNIQUE' then 1" & LineBreakCode
      QueryString = QueryString + "    end as Uniqueness" & LineBreakCode
      QueryString = QueryString + "FROM" & LineBreakCode
      QueryString = QueryString + "  USER_IND_COLUMNS  UIC left join USER_INDEXES UI on UIC.table_name=UI.table_name and UIC.index_name=UI.index_name" & LineBreakCode
      QueryString = QueryString + "where" & LineBreakCode
      QueryString = QueryString + "  UIC.table_name = '" & tableName & "'" & LineBreakCode
      QueryString = QueryString + "Order by" & LineBreakCode
      QueryString = QueryString + "  uniqueness ASC" & LineBreakCode
      QueryString = QueryString + "  , UIC.index_name ASC" & LineBreakCode
      QueryString = QueryString + "  , UIC.column_position ASC"

    Case "SQLServer"
      QueryString = "EXEC sp_MShelpindex " & tableName
  End Select

  '�C���f�b�N�X���擾
  Range("W5").Value = "�C���f�b�N�X���擾SQL"
  Range("X5").Value = QueryString
  Range("X5").WrapText = False

  Set DBRecordset = New ADODB.Recordset
  DBRecordset.Open QueryString, DBConGetColumn, adOpenKeyset, adLockReadOnly

  Do Until DBRecordset.EOF

    ProgressBar_ProgShowCount tableName & " �ڑ����E�E�E", indexCount, 5, " �C���f�b�N�X���擾�F"

    If DBMS = "SQLServer" Then
      DBRecordset.MoveNext
    Else
      Select Case DBMS
        Case "PostgreSQL"

        Case "MySQL"
          IndexName = DBRecordset.Fields("Key_name").Value
          IndexColumnName = DBRecordset.Fields("Column_name").Value
          IndexTableSpace = ""
          IndexUniqueness = DBRecordset.Fields("Non_unique").Value
          indexCount = DBRecordset.Fields("Seq_in_index").Value

        Case "Oracle"
          IndexName = DBRecordset.Fields("IndexName").Value
          IndexColumnName = DBRecordset.Fields("Column_name").Value
          IndexTableSpace = DBRecordset.Fields("TableSpace").Value
          IndexUniqueness = DBRecordset.Fields("Non_unique").Value

        Case "SQLServer"
      End Select


      rowLine = ActiveSheet.Cells.Find(IndexColumnName, LookAt:=xlWhole).Row

      If Range("C" & NowLine).Value <> IndexName Then
        indexCount = indexCount + 1
        If indexCount > 5 Then
          ActiveSheet.Tab.Color = RGB(255, 183, 183)
          Range("D6") = Range("D6") & "�C���f�b�N�X�s���s��"

          If SetDisplyProgressBarFlg Then
            ProgressBar_ProgShowClose
          End If
          Exit Function
        End If

        NowLine = NowLine + 1
        Range("C" & NowLine).Value = IndexName
        Range("H" & NowLine).Value = IndexTableSpace
        Select Case IndexUniqueness
          Case 0
            Range("E" & NowLine).Value = "UNIQUE"
          Case 1
            Range("E" & NowLine).Value = "NONUNIQUE"
          End Select
      End If


      Select Case indexCount
        Case 1
          Range("j" & rowLine).Value = indexCount
        Case 2
          Range("k" & rowLine).Value = indexCount
        Case 3
          Range("L" & rowLine).Value = indexCount
        Case 4
          Range("M" & rowLine).Value = indexCount
        Case 5
          Range("N" & rowLine).Value = indexCount
      End Select


      '���̃��R�[�h
      DBRecordset.MoveNext
    End If
  Loop
  ProgressBar_ProgShowCount tableName & " �ڑ����E�E�E", 5, 5, " �C���f�b�N�X���擾�F"


'-------------------------------------------------------------------------------------------------------------------------------------------
'�g���K�[���擾
'-------------------------------------------------------------------------------------------------------------------------------------------
'  NowLine = 117
'  Set DBRecordset = Nothing
'  QueryString = "SELECT "
'  QueryString = QueryString + "triggers.name,modules.definition "
'  QueryString = QueryString + "FROM sys.triggers triggers "
'  QueryString = QueryString + "INNER JOIN sys.objects objects ON triggers.object_id = objects.object_id "
'  QueryString = QueryString + "INNER JOIN sys.tables as t ON triggers.parent_id = t.object_id  "
'  QueryString = QueryString + "INNER JOIN sys.schemas schemas ON objects.schema_id = schemas.schema_id "
'  QueryString = QueryString + "INNER JOIN sys.sql_modules modules ON objects.object_id = modules.object_id "
'  QueryString = QueryString + "where t.name='EmployeeMaster' "
'
'  Set DBRecordset = New ADODB.Recordset
'  DBRecordset.Open QueryString, DBCon, adOpenKeyset, adLockReadOnly
'
'
'  Do Until DBRecordset.EOF
'    TriggersName = DBRecordset.Fields("name").Value
'    TriggersDefinition = DBRecordset.Fields("definition").Value
'
'    Range("C" & NowLine).Value = TriggersName
'    Range("E" & NowLine).Value = "INSERT, UPDATE"
'    Range("H" & NowLine).Value = TriggersDefinition
'
'    '���̃��R�[�h
'    NowLine = NowLine + 1
'    DBRecordset.MoveNext
'  Loop

  ProgressBar_ProgShowCount "�������E�E�E", 1, 5, " DB�ؒf��"

  'DB�N���[�Y
  If SetDisplyProgressBarFlg = True Then
    DBConGetColumn.Close
    Set DBRecordset = Nothing
    Set GetTableListCon = Nothing
  End If

  '�s�v�s�̍폜
  ProgressBar_ProgShowCount "�������E�E�E", 2, 5, "�s������"
  DataBase_SetSheetStyle

  ProgressBar_ProgShowCount "�������E�E�E", 3, 5, "����͈͐ݒ�"
  DataBase_SetPrintArea

  Range("A9").Select


  ' �v���O���X�o�[�̕\���I������
  If SetDisplyProgressBarFlg Then
    ProgressBar_ProgShowClose
  End If

  Select Case DBMS
    Case "PostgreSQL"

    Case "MySQL"
'      MsgBox "�v���C�}���L�[�̐ݒ肪�ł��Ă��܂���"
    Case "Oracle"

    Case "SQLServer"
  End Select
  Range("C9").Select

  Exit Function
'-------------------------------------------------------------------------------------------------------------------------------------------
'�G���[�������̏���
'-------------------------------------------------------------------------------------------------------------------------------------------
GetColumnList_Error:

'  If (Err.Number = 3265) Or (Err.Number = 3709) Then
'  Else
    Call Library_ErrorHandle(Err.Number, Err.Description)
'  End If

  'DB�N���[�Y
  If (dbCon.State And adStateOpen) = adStateOpen Then
    dbCon.Close
  End If

  Set DBRecordset = Nothing
  Set GetTableListCon = Nothing

  ' �v���O���X�o�[�̕\���I������
  ProgressBar_ProgShowClose

  Range("C9").Select
End Function



'***************************************************************************************************************************************************
' * �e�[�u�����X�g����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***************************************************************************************************************************************************
Function DataBase_MakeTableList()

  Dim objSheet As Object
  Dim intLoop As Integer
  Dim Name As String
  Dim DB_table_name As String
  Dim DB_table_id As String
  Dim DB_table_id_column As Integer
  Dim DB_table_comment As String
  Dim DB_table_Kind As String

  Dim endLine As Integer

  Sheets("TBL���X�g").Select
  endLine = Cells(Rows.count, 3).End(xlUp).Row + 1


  ' ���݂̃A�N�e�B�u�Z���̍s�ԍ����i�[
  Range("C6").Select
  intLoop = ActiveCell.Row

  ' ���ݐݒ肳��Ă���l���폜
  Range("C6:P" & endLine).Select
  Selection.ClearContents
  Range("C6").Select

  For Each objSheet In ActiveWorkbook.Sheets
    ' ���ʔ����֐��ďo��
    Name = objSheet.Name
    endLine = Cells(Rows.count, 3).End(xlUp).Row + 1

    If Library_CheckExcludeSheet(Name, 9) = True Then

      ' �ꗗ�\���쐬
      Set sh_ini = Worksheets(Name)
      ' �e�[�u�����̎擾
      DB_table_Kind = sh_ini.Range("B2").Value
      DB_table_name = sh_ini.Range("D5").Value
      DB_table_id = sh_ini.Range("H5").Value
      DB_table_comment = sh_ini.Range("D6").Value
      DBTableSpace = sh_ini.Range("T5").Value

      ' ����
      Range("C" & endLine).Value = DB_table_Kind

      If DB_table_name = "" Then
        DB_table_name = " "
      End If

      '�e�[�u������
      With Range("E" & endLine)
        .Value = DB_table_name
        .Select
        .Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=objSheet.Name & "!" & "A9"
        .Font.Color = RGB(0, 0, 0)
        .Font.Underline = False
        .Font.Size = 10
        .Font.Name = "���C���I"
      End With


       '�e�[�u��ID
      With Range("H" & endLine)
        .Value = DB_table_id
        .Select
        .Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=objSheet.Name & "!" & "A9"
        .Font.Color = RGB(0, 0, 0)
        .Font.Underline = False
        .Font.Size = 10
        .Font.Name = "���C���I"
      End With

      '����
      Range("Q" & endLine).Value = DB_table_comment

      ' �Z���̔w�i�F����
      With Range("B" & endLine & ":U" & endLine).Interior
        .Pattern = xlPatternNone
        .Color = xlNone
      End With

      ' �V�[�g�F�Ɠ����F���Z���ɐݒ�
      If sh_ini.Tab.Color Then
        With Range("B" & endLine & ":U" & endLine).Interior
          .Pattern = xlPatternNone
          .Color = sh_ini.Tab.Color
        End With
      End If
    End If

  Next
  Worksheets("�ύX����").Select
  endLine = Cells(Rows.count, 4).End(xlUp).Row
  Range("�X�V��") = Range("D" & endLine)

  endLine = Cells(Rows.count, 3).End(xlUp).Row
  Range("�X�V��") = Range("C" & endLine)


  Worksheets("TBL���X�g").Select
  Range("B6").Select

End Function

'***************************************************************************************************************************************************
' * ����͈͐ݒ�
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***************************************************************************************************************************************************
Function DataBase_SetPrintArea()

  On Error GoTo ErrHand

  Dim endLine As Integer
  Dim PageCnt As Integer
  Dim OnePageRow As Integer
  Dim RowCnt As Integer
  Dim ThisActiveSheetName As String
  Dim WindowZoomLevel As Integer

  WindowZoomLevel = ActiveWindow.Zoom

  ThisActiveSheetName = ActiveSheet.Name

  endLine = ActiveSheet.Cells(Rows.count, 2).End(xlUp).Row
  OnePageRow = 30
  PageCnt = 1

  ' ======================= �����J�n ======================
  '���y�[�W�v���r���[
  ActiveWindow.View = xlPageBreakPreview

  '���ׂẲ��y�[�W������
  ActiveSheet.ResetAllPageBreaks

  '����͈͂��N���A����
  ActiveSheet.PageSetup.PrintArea = ""

  '����͈͂̏ڍאݒ�
  With ActiveSheet.PageSetup
    .CenterFooter = "&P / &N"
    .PrintTitleRows = "$2:$8"                 '�s�^�C�g��
    .PrintArea = "$B$2:$U$" & endLine
    .BlackAndWhite = False                    '������� True:����  False:���Ȃ�
    .Zoom = False                             '�g��E�k�������w�肵�Ȃ�
    .FitToPagesTall = False                   '�c�����͎w�肵�Ȃ�
    .FitToPagesWide = 1                       '������1�y�[�W�ň��

    .TopMargin = Application.CentimetersToPoints(1.2)       '��]��
    .BottomMargin = Application.CentimetersToPoints(1)    '���]��
    .LeftMargin = Application.CentimetersToPoints(1)        '���]��
    .RightMargin = Application.CentimetersToPoints(1)       '�E�]��
    .HeaderMargin = Application.CentimetersToPoints(0.8)    '�w�b�_�[�]��
    .FooterMargin = Application.CentimetersToPoints(0.5)    '�t�b�^�[�]��
  End With

  '�W����ʂɖ߂�
  ActiveWindow.View = xlNormalView
  ActiveWindow.Zoom = WindowZoomLevel

Exit Function

ErrHand:
  ActiveWindow.View = xlNormalView
  ActiveWindow.Zoom = WindowZoomLevel

  Call Library_ErrorHandle(Err.Number, Err.Description)
End Function


'***************************************************************************************************************************************************
' * �J�����G���A�̕s�v�s�폜
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***************************************************************************************************************************************************
Sub DataBase_SetSheetStyle()

  Dim ColumnEndLine As Integer
  Dim IndexLine As Integer
  Dim i As Integer
  Dim DeleteFlg As Boolean

  DeleteFlg = False

  Library_UnsetLineColor ("B9:U" & Cells(Rows.count, 2).End(xlUp).Row)

  Columns("A:A").ClearContents

  '�J�����ݒ�s����ݒ�
  For i = 9 To 1000
    If Range("E" & i) = "" Then
      Range("A" & i) = "Column"
      ColumnEndLine = i

      Exit For
     End If
  Next
  For i = 9 To 1000
    If Range("C" & i) = "index" Or Range("C" & i) = "�C���f�b�N�X��" Then
      Range("A" & i) = "Index"
      IndexLine = i - 3
      Exit For
    End If
  Next
  For i = 9 To 1000
    If Range("C" & i) = "trigger" Or Range("C" & i) = "�g���K�[��" Then
      Range("A" & i) = "Trigger"
      Exit For
    End If
  Next

  '1�y�[�W
  If (IndexLine < 29) Then
    DeleteFlg = False

  ElseIf 9 <= ColumnEndLine And ColumnEndLine < 27 Then
    Rows("27:" & IndexLine).Select
    DeleteFlg = True

  '2�y�[�W
  ElseIf 28 <= ColumnEndLine And ColumnEndLine < 60 Then
    Rows("60:" & IndexLine).Select
    DeleteFlg = True

  '3�y�[�W
  ElseIf 63 <= ColumnEndLine And ColumnEndLine < 92 Then
    Rows("92:" & IndexLine).Select
    DeleteFlg = True

  '4�y�[�W
  ElseIf 115 <= ColumnEndLine And ColumnEndLine < 126 Then
    Rows("126:" & IndexLine).Select
    DeleteFlg = True

  Else
    DeleteFlg = False
  End If

  If DeleteFlg Then
    Selection.Delete Shift:=xlUp
  End If

  Call Library_SetLineColor("B9:U" & Cells(Rows.count, 2).End(xlUp).Row, False, RGB(255, 255, 155))
  Range("A9").Select
End Sub

'***************************************************************************************************************************************************
' * �J�����̌���
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***************************************************************************************************************************************************
Function DataBase_SetCellsStyle()

  Dim ColumnEndLine As Integer
  Dim IndexLine As Integer
  Dim i As Integer
  Dim DeleteFlg As Boolean
  Dim ColumnValue As String

  DeleteFlg = False
  '�J�����ݒ�s�����擾
  For i = 9 To 1000
    If Range("A" & i) = "Column" Then
      ColumnEndLine = i

    ElseIf Range("A" & i) = "Index" Then
      IndexLine = i - 3
      Exit For
    End If

    ColumnValue = Range("Q" & i)
    ColumnValue = Replace(ColumnValue, ":", "�F")
    ColumnValue = Replace(ColumnValue, ",", "�A")

    Range("Q" & i) = ColumnValue
  Next

  Range("C8:D" & IndexLine).Select
  Selection.Merge True

  Range("E8:F" & IndexLine).Select
  Selection.Merge True

  Range("Q8:U" & IndexLine).Select
  Selection.Merge True

  Library_UnsetLineColor ("B9:U" & Cells(Rows.count, 2).End(xlUp).Row)
  Call Library_SetLineColor("B9:U" & Cells(Rows.count, 2).End(xlUp).Row, False, RGB(255, 255, 155))
  Range("B9").Select
End Function

' *********************************************************************
' * �ڑ��m�F
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
' *********************************************************************
Function DataBase_Connection()

  Dim tableName As String
  Dim Comment As Variant
  Dim newSheetName As String

  Dim DBRecordsetCount As Integer
  Dim dbCon As ADODB.Connection
  Dim DBRecordset As ADODB.Recordset



  On Error GoTo DBConnection_Error

  DataBase_Init
  DBRecordsetCount = 1

  'ADODB.Connection�������ADB�ɐڑ�
  Set dbCon = New ADODB.Connection
  dbCon.Open ConnectionString

  'SQL�� -------------------------------------------------------------------------------------------------------------------------------------
  Select Case DBMS
    Case "PostgreSQL"
      PostgreSQL_MakeDDL

    Case "MySQL"
      QueryString = "SELECT TABLE_NAME as TableName, TABLE_COMMENT as Comments,'' as  TableSpaceName from information_schema.TABLES WHERE TABLE_SCHEMA = DATABASE();"

    Case "Oracle"
      If DBTableSpace = "" Then
        QueryString = "select" & LineBreakCode
        QueryString = QueryString & "  UT.table_name TableName," & LineBreakCode
        QueryString = QueryString & "  UTC.Comments," & LineBreakCode
        QueryString = QueryString & "  UT.tablespace_name TableSpaceName" & LineBreakCode
        QueryString = QueryString & "from USER_TABLES UT left join USER_TAB_COMMENTS UTC on UT.table_name =UTC.table_name" & LineBreakCode
        QueryString = QueryString & "where UT.tablespace_name is not null " & LineBreakCode
      Else
        QueryString = "select" & LineBreakCode
        QueryString = QueryString & "  UT.table_name TableName," & LineBreakCode
        QueryString = QueryString & "  UTC.Comments," & LineBreakCode
        QueryString = QueryString & "  UT.tablespace_name TableSpaceName" & LineBreakCode
        QueryString = QueryString & "from USER_TABLES UT left join USER_TAB_COMMENTS UTC on UT.table_name =UTC.table_name" & LineBreakCode
        QueryString = QueryString & "where UT.tablespace_name='" & DBTableSpace & "';" & LineBreakCode
      End If

      QueryString = QueryString & " order by UT.table_name" & LineBreakCode

    Case "SQLServer"
      QueryString = "select table_name TableName,'' Comments from USER_TABLES;"
  End Select


  Set DBRecordset = New ADODB.Recordset
  DBRecordset.Open QueryString, dbCon, adOpenKeyset, adLockReadOnly


  Set DBRecordset = Nothing
  'DB�N���[�Y
  dbCon.Close
  Set DBRecordset = Nothing

  Exit Function

'---------------------------------------------------------------------------------------------------------------------------------------------
'�G���[�������̏���
'---------------------------------------------------------------------------------------------------------------------------------------------
DBConnection_Error:
  Call Library_ErrorHandle(Err.Number, Err.Description)

  If (dbCon.State And adStateOpen) = adStateOpen Then
    'DB�N���[�Y
    dbCon.Close
  End If
  ConnectionString = ""
  Set DBRecordset = Nothing
  Set GetTableList_Result = Nothing
  Set GetTableList_Con = Nothing

End Function
