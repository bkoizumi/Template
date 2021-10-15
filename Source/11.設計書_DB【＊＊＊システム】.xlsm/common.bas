Attribute VB_Name = "common"
Sub Macro4()
    
  Call Library.startScript
    Sheets("Master").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=31
    Sheets(Array("Master", "MC_�Č��Ǘ��\", "MC_�Ϗ�Ǘ��\", "SI_�`�[���ʃ}�X�^", "SI_�v���W�F�N�g�ꗗ", _
        "SI_�v���W�F�N�g�̎Z����", "SI_�O�T�Č��ꗗ", "�ٓ��Ј����", "�ߔN�x�f�[�^", "������_���ʉ�c��", "������_���ʌ��۔�_�ڕ�", _
        "������_���ʌ��۔�_�l��", "������_���ʌ�ʔ�", "������_���ʗ���o����", "������_�ڋq�f�[�^", "������_���Z�f�[�^", "������_���Ȏ҃f�[�^", _
        "������_�s�����o_�^�N�V�[_��", "������_�s�����o_�^�N�V�[_���z", "������_�s�����o_���۔�_��", "������_�s�����o_���۔�_���z", _
        "������_�s�����o_��ʔ�_���z", "������_�s�����o_�������o��", "������_�s�����o_����o����_��", "�Ǘ���v�R�[�h")).Select
    Sheets("Master").Activate
    Sheets(Array("����", "�Ј����", "�Г��o��Z�\", "�S�Ј����", "�S�Бg�D�R�[�h", "�g�D�}�X�^", "���ڊԐڔ䗦", "�N�v�f�[�^", _
        "�v����")).Select Replace:=False
    Sheets(Array("Master", "MC_�Č��Ǘ��\", "MC_�Ϗ�Ǘ��\", "SI_�`�[���ʃ}�X�^", "SI_�v���W�F�N�g�ꗗ", _
        "SI_�v���W�F�N�g�̎Z����", "SI_�O�T�Č��ꗗ", "�ٓ��Ј����", "�ߔN�x�f�[�^", "������_���ʉ�c��", "������_���ʌ��۔�_�ڕ�", _
        "������_���ʌ��۔�_�l��", "������_���ʌ�ʔ�", "������_���ʗ���o����", "������_�ڋq�f�[�^", "������_���Z�f�[�^", "������_���Ȏ҃f�[�^", _
        "������_�s�����o_�^�N�V�[_��", "������_�s�����o_�^�N�V�[_���z", "������_�s�����o_���۔�_��", "������_�s�����o_���۔�_���z", _
        "������_�s�����o_��ʔ�_���z", "������_�s�����o_�������o��", "������_�s�����o_����o����_��", "�v����")).Select
    Sheets("�v����").Activate
    Sheets(Array("�Ǘ���v�R�[�h", "����", "�Ј����", "�Г��o��Z�\", "�S�Ј����", "�S�Бg�D�R�[�h", "�g�D�}�X�^", "���ڊԐڔ䗦" _
        , "�N�v�f�[�^")).Select Replace:=False
    ActiveWindow.SelectedSheets.Delete

  Call Library.endScript
End Sub


'**************************************************************************************************
' * ���ʏ���
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function �V�[�g�N���A()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
'  On Error GoTo catchError

  endLine = Cells(Rows.count, 2).End(xlUp).Row
  
  On Error Resume Next
  Rows(startLine & ":" & endLine).SpecialCells(xlCellTypeConstants, 23).ClearContents
  

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, Err.Description, True)
End Function


'==================================================================================================
Function �V�[�g�ǉ�(newSheetName As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
'  On Error GoTo catchError

    sheetCopyTable.copy After:=Worksheets(Worksheets.count)
    ActiveWorkbook.Sheets(Worksheets.count).Tab.ColorIndex = -4142
    ActiveWorkbook.Sheets(Worksheets.count).Name = newSheetName
    Sheets(newSheetName).Select
    
    
    
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, Err.Description, True)
End Function


'==================================================================================================
Function TBL���X�g�V�[�g����()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim sheetList As Object
  Dim targetSheet   As Worksheet
  Dim SheetName As String
  Dim result As Boolean
  
'  On Error GoTo catchError

'�����J�n--------------------------------------
  Call init.Setting
  Call Library.startScript
  Call Ctl_ProgressBar.showStart
  
  sheetTblList.Select
  endLine = sheetTblList.Cells(Rows.count, 3).End(xlUp).Row + 1
  Range("B6:U" & endLine).ClearContents
  
  With Range("B6:U" & endLine).Interior
    .Pattern = xlPatternNone
    .Color = xlNone
  End With
  
      
  line = 6
  For Each sheetList In ActiveWorkbook.Sheets
    SheetName = sheetList.Name
    result = Library.chkExcludeSheet(SheetName)
    
    If result = True Then
    Else
      sheetTblList.Range("B" & line).FormulaR1C1 = "=ROW()-5"
      sheetTblList.Range("C" & line) = Sheets(SheetName).Range("B2")
'      sheetTblList.Range("E" & line) = Sheets(sheetName).Range("D5")
'      sheetTblList.Range("H" & line) = Sheets(sheetName).Range("H5")
      sheetTblList.Range("Q" & line) = Sheets(SheetName).Range("D6")
    
      '�_���e�[�u����
      If Sheets(SheetName).Range("D5") <> "" Then
        With sheetTblList.Range("E" & line)
          .Value = Sheets(SheetName).Range("D5")
          .Select
          .Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=SheetName & "!" & "A9"
          .Font.Color = RGB(0, 0, 0)
          .Font.Underline = False
          .Font.Size = 10
          .Font.Name = "���C���I"
        End With
      End If
      
       '�����e�[�u����
      With sheetTblList.Range("H" & line)
        .Value = Sheets(SheetName).Range("H5")
        .Select
        .Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=SheetName & "!" & "A9"
        .Font.Color = RGB(0, 0, 0)
        .Font.Underline = False
        .Font.Size = 10
        .Font.Name = "���C���I"
      End With
      
      ' �V�[�g�F�Ɠ����F���Z���ɐݒ�
      If Sheets(SheetName).Tab.Color Then
        With sheetTblList.Range("B" & line & ":U" & line).Interior
          .Pattern = xlPatternNone
          .Color = Sheets(SheetName).Tab.Color
        End With
      End If
      
      
      line = line + 1
    End If
  Next
  
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
    
    
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, Err.Description, True)
End Function


