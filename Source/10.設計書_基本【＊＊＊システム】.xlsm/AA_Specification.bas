Attribute VB_Name = "AA_Specification"
Public W_Page1NoCol As Integer
Public W_Page2NoCol As Integer
Public OnePageRow As Integer
Public Page1Area As String
Public Page2Area As String
Public Page1StartArea As Integer
Public Page1CenterArea As String
Public Page2CenterArea As String


Public InputName As String
Public InputType As String
Public InputDataType As String
Public InputNameTag As String
Public InputLimit_tmp As String
Public InputRequired As String
Public InputTestString As String
Public InputLimit As Variant
Public URL As String
Public InputLimitMin As Long
Public InputLimitMax As Long
Public Title As String
Public InputNo As Long


'***********************************************************************************************************************************************
' * �݌v���p���ݒ�
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Sub init()

  ' 1�s�ڂ̃y�[�W���������݈ʒu
  W_Page1NoCol = Worksheets("�ݒ�").Range("B3")
  
  ' 2�s�ڂ̃y�[�W���������݈ʒu
  W_Page2NoCol = Worksheets("�ݒ�").Range("B4")

  ' 1�y�[�W�̍s��
  OnePageRow = Worksheets("�ݒ�").Range("B5")

  ' 1�y�[�W�ڂ̖ڎ��J�n�ʒu
  Page1StartArea = Worksheets("�ݒ�").Range("B6")
  
  ' 1�y�[�W�ڂ̖ڎ��\���ʒu
  Page1Area = Worksheets("�ݒ�").Range("B7")

  ' 1�y�[�W�ڂ̖ڎ������ʒu
  Page1CenterArea = Worksheets("�ݒ�").Range("B8")
  
  ' 2�y�[�W�ڂ̖ڎ��\���ʒu
  Page2Area = Worksheets("�ݒ�").Range("B9")

  ' 2�y�[�W�ڂ̖ڎ������ʒu
  Page2CenterArea = Worksheets("�ݒ�").Range("B10")

End Sub

'***********************************************************************************************************************************************
' * �݌v���p�ڎ��쐬
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Sub MakeMenu()

  Dim PageLine As Long
  Dim TitleName As String
  Dim FunctionName As String
  Dim PageCnt As Long
  Dim TitleCnt As Long
  Dim EndBookRowLine As Long
  Dim RowCnt As Long
  Dim W_PageNoCol As Long
  
  Dim ThisActiveSheetName As String
  ThisActiveSheetName = ActiveSheet.Name
  
  Call Specification.init
  
  ' �ŏI�s�擾
  EndBookRowLine = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row
  
  ' ���ݐݒ肳��Ă���ڎ����폜
  Call Specification.DeleteMenu(1)

  '---------------------------------------------------------------------------------------
  ' �ڎ��������C������
  '---------------------------------------------------------------------------------------
  PageLine = Page1StartArea
  PageCnt = 1
  W_PageNoCol = W_Page1NoCol
  
  ' �v���O���X�o�[�̕\���J�n
  ProgressBar_ProgShowStart
 
  For RowCnt = 44 To EndBookRowLine Step OnePageRow
  
    ' �^�C�g���擾
    TitleName = Cells(RowCnt + 1, 4)
    
    ' �@�\�擾
    FunctionName = Cells(RowCnt + 1, 19)
    
    ' �y�[�W�ԍ���������
    With Cells(PageLine, W_PageNoCol)
      .Value = PageCnt
      .Font.Name = "Meiryo UI"
      .Font.Size = 9
      .HorizontalAlignment = xlGeneral
      .VerticalAlignment = xlCenter
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = True
      .ReadingOrder = xlContext
      .MergeCells = False
      .ShrinkToFit = True
'      .NumberFormatLocal = "@"
    End With
        
    If FunctionName <> "" Then
      TitleName = TitleName & " - " & FunctionName
    End If
    
    ' �v���O���X�o�[�̃J�E���g�ύX�i���݂̃J�E���g�A�S�J�E���g���A���b�Z�[�W�j
    ProgressBar_ProgShowCount "�ڎ�������", RowCnt, EndBookRowLine, "P." & PageCnt & " " & TitleName
    
    ' �^�C�g��(�����N�t)��������
    With Cells(PageLine, W_PageNoCol + 1)
      .Value = TitleName
      .Select
      .Hyperlinks.add Anchor:=Selection, Address:="", SubAddress:="#" & "A" & RowCnt
      .Font.ColorIndex = 1
      .Font.Underline = xlUnderlineStyleNone
      .Font.Name = "Meiryo UI"
      .Font.Size = 9
      .HorizontalAlignment = xlGeneral
      .VerticalAlignment = xlCenter
      .ShrinkToFit = True
    End With
    
    ' �Z���̌���
    Range("E" & PageLine & ":V" & PageLine).Select
    Selection.Merge
    Range("AA" & PageLine & ":AR" & PageLine).Select
    Selection.Merge
    
    ' �e�y�[�W�Ƀy�[�W�ԍ���������
    Range("AW" & RowCnt & ":AX" & RowCnt + 1).Select
    Selection.Merge
    
    Range("AW" & RowCnt).Value = "P." & PageCnt
    
    ' �ڎ��ւ̃����N�ǉ�
    If RowCnt > 2 Then
      Range("AW" & RowCnt - 1 & ":AX" & RowCnt - 1).Select
      Selection.Merge
      Range("AW" & RowCnt - 1).Value = "=HYPERLINK(""#$A$1"",""�ڎ���"")"
    End If
    
    PageLine = PageLine + 1
    PageCnt = PageCnt + 1
  
    ' ======================= ���� ======================
    ' 1�y�[�W�ڂ�2���
    If PageCnt = OnePageRow - 4 Then
      W_PageNoCol = W_Page2NoCol
      PageLine = OnePageRow + 5
      Call Specification.AddLine(1)
      
    ' 2�y�[�W�ڂ�1��
    ElseIf PageCnt = (OnePageRow - 5) * 2 + 1 Then
    
      If Range("D88") <> "�ڎ�" And Range("D88") <> "������" Then
        If MsgBox("�ڎ���2�y�[�W�ڂɑ}������܂�" & vbLf & " 2�y�[�W�̏���OK�H", vbYesNo, "2�y�[�W�̏���OK�H") = vbNo Then
          Call Library.EndScript
          MsgBox "2�y�[�W�ڂ̃^�C�g����ڎ��ɐݒ肵�Ă�������" & vbLf & "�����𒆒f���܂��B"
          
          ' �v���O���X�o�[�̕\���I������
          ProgressBar_ProgShowClose
          
          Exit Sub
        End If
      Else
        ' ======================= 2�y�[�W�ڎ�����======================
        W_PageNoCol = W_Page1NoCol
        PageLine = OnePageRow * 2 + 5
        Call Specification.DeleteMenu(2)
        Call Specification.AddLine(2)

      End If
    
    ' 2�y�[�W�ڂ�2��
    ElseIf PageCnt = (OnePageRow - 5) * 3 + 1 Then
      W_PageNoCol = W_Page2NoCol
      PageLine = OnePageRow * 2 + 5
    End If
  Next

  ' �v���O���X�o�[�̕\���I������
  ProgressBar_ProgShowClose

  ' ����̈�ݒ�
  'Call Specification.SetPrintArea


End Sub
'***********************************************************************************************************************************************
' * �݌v���p�ڎ��폜
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function DeleteMenu(Page As Integer)

  If Page = 1 Then
    Range(Page1Area).Select
  ElseIf Page = 2 Then
    Range(Page2Area).Select
  End If
  
  Selection.Clear
  Application.CutCopyMode = False
End Function


'***********************************************************************************************************************************************
' * �݌v���p�r���ݒ�
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function AddLine(Page As Integer)
  
  If Page = 1 Then
    Range(Page1CenterArea).Select
  ElseIf Page = 2 Then
    Range(Page2CenterArea).Select
  End If
  
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Function

'***********************************************************************************************************************************************
' * �݌v���p�y�[�W�ǉ�
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function addPage()

  On Error GoTo ErrHand
  
  Dim PageLine As Integer
  Dim TitleName As String
  Dim FunctionName As String
  Dim PageCnt As Integer
  Dim TitleCnt As Integer
  Dim EndBookRowLine As Integer
  Dim RowCnt As Integer
  Dim W_PageNoCol As Integer
  Dim ThisActiveSheetName As String
  
  Call Specification.init
  
  ThisActiveSheetName = ActiveSheet.Name
  EndBookRowLine = Sheets(ThisActiveSheetName).Cells(Rows.count, 1).End(xlUp).Row + OnePageRow - 1
  

  Sheets("Sheet1").Select
  Range("A1:AW43").Select
  Selection.Copy

  Sheets(ThisActiveSheetName).Select
  Range("A" & EndBookRowLine).Select
  ActiveSheet.Paste

  Application.CutCopyMode = False

  ' �O�y�[�W�̃^�C�g���ݒ�
  ActiveSheet.Range("D" & EndBookRowLine + 1).Value = Range("D" & EndBookRowLine - OnePageRow + 1).Value
  
  ActiveSheet.Range("A" & EndBookRowLine).Select
  With ActiveWindow
    .ScrollRow = EndBookRowLine
    .ScrollColumn = 1
  End With
Exit Function

ErrHand:
  Call Library.EndScript
  Resume Next
End Function


'***********************************************************************************************************************************************
' * �݌v���p����͈͐ݒ�
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Sub SetPrintArea()

  On Error GoTo ErrHand
  
  Dim EndBookRowLine As Long
  Dim PageCnt As Long
  Dim W_PageNoCol As Long
  Dim RowCnt As Long
  Dim ThisActiveSheetName As String
  
  Call Specification.init
  
  ThisActiveSheetName = ActiveSheet.Name
  
  EndBookRowLine = ActiveSheet.Cells(Rows.count, 49).End(xlUp).Row
  W_PageNoCol = OnePageRow
  PageCnt = 1
  
  ActiveSheet.PageSetup.PrintArea = "A1:AU" & EndBookRowLine
  
  '���y�[�W�v���r���[
  ActiveWindow.View = xlPageBreakPreview
  
  ' �v���O���X�o�[�̕\���J�n
  ProgressBar_ProgShowStart
  
  For RowCnt = 1 To EndBookRowLine Step OnePageRow

    ' �v���O���X�o�[�̃J�E���g�ύX�i���݂̃J�E���g�A�S�J�E���g���A���b�Z�[�W�j
    ProgressBar_ProgShowCount "����͈͐ݒ�", RowCnt, EndBookRowLine, "P." & PageCnt
    
    Set Sheets(ThisActiveSheetName).HPageBreaks(PageCnt).Location = Range("A" & W_PageNoCol + 1)
    W_PageNoCol = W_PageNoCol + OnePageRow
    PageCnt = PageCnt + 1
  Next RowCnt
  
  ActiveWindow.View = xlNormalView

  ' �v���O���X�o�[�̕\���I������
  ProgressBar_ProgShowClose

Exit Sub

ErrHand:
  ActiveWindow.View = xlNormalView
  
  ' �v���O���X�o�[�̕\���I������
  ProgressBar_ProgShowClose

  ' ��ʕ`�ʐ���I��
  Call Library.EndScript
End Sub


'***********************************************************************************************************************************************
' * �݌v���pSelenium�݌v���`�F�b�N
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function Selenium_Check()

  Dim EndBookRowLine As Long
  Dim RowCnt As Long
  Dim TitleName As String
  
  
  Call Specification.init
  
  ' �ŏI�s�擾
  EndBookRowLine = ActiveSheet.Cells(Rows.count, 49).End(xlUp).Row
  For RowCnt = 1 To EndBookRowLine Step OnePageRow
    ' �@�\���擾
    If Cells(RowCnt + 1, 19) = "���͍��ڐ���" Then
      Title = Cells(RowCnt + 1, 4)
      Call Specification.Selenium_Get(RowCnt + 1)
    End If
  Next

End Function


'***********************************************************************************************************************************************
' * �݌v���pSelenium���擾
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function Selenium_Get(RowCnt As Long)

  Dim line As Long

  For line = RowCnt To RowCnt + OnePageRow Step 1
    If Range("B" & line) = "No." Then
      URL = Range("B" & line - 1)
      Call Specification.Selenium_Make(line + 1)
    End If
  Next
End Function


'***********************************************************************************************************************************************
' * �݌v���pSelenium�e�X�g�P�[�X����
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function Selenium_Make(RowCnt As Long)

  Dim line As Long

  For line = RowCnt To RowCnt + OnePageRow Step 1
    If Range("J" & line) = "�e�L�X�g�G���A" Or Range("J" & line) = "�e�L�X�g�{�b�N�X" Then
      InputName = Range("C" & line)
      InputType = Range("J" & line)
      InputDataType = Range("O" & line)
      InputNameTag = Range("S" & line)
      InputLimit_tmp = Range("W" & line)
      InputRequired = Range("Z" & line)
      InputNo = Range("B" & line)

      '���͌����̍ŏ�/�ő���擾
      If InStr(InputLimit_tmp, "�`") <> 0 Then
        InputLimit = Split(InputLimit_tmp, "�`")
        InputLimitMin = CLng(InputLimit(0))
        InputLimitMax = CLng(InputLimit(1))
      ElseIf InputLimit_tmp <> "" Then
        InputLimitMin = InputLimit_tmp
        InputLimitMax = 0
      End If
      
    '�e�X�g���ڍ쐬-----------------------------------------------------------------------------------------
      Call Specification.Selenium_Makehtml("���p����-�ŏ������ȉ�")
      Call Specification.Selenium_Makehtml("���p����-�ŏ���")
      Call Specification.Selenium_Makehtml("���p����")
      Call Specification.Selenium_Makehtml("���p����-�ő包")
      Call Specification.Selenium_Makehtml("���p����-�ő包���ȏ�")
      
      Call Specification.Selenium_Makehtml("���p�p������-�ŏ������ȉ�")
      Call Specification.Selenium_Makehtml("���p�p������-�ŏ���")
      Call Specification.Selenium_Makehtml("���p�p������")
      Call Specification.Selenium_Makehtml("���p�p������-�ő包")
      Call Specification.Selenium_Makehtml("���p�p������-�ő包���ȏ�")
      
      Call Specification.Selenium_Makehtml("���p�p�啶��-�ŏ������ȉ�")
      Call Specification.Selenium_Makehtml("���p�p�啶��-�ŏ���")
      Call Specification.Selenium_Makehtml("���p�p�啶��")
      Call Specification.Selenium_Makehtml("���p�p�啶��-�ő包")
      Call Specification.Selenium_Makehtml("���p�p�啶��-�ő包���ȏ�")
      
      Call Specification.Selenium_Makehtml("���p�p����-�ŏ������ȉ�")
      Call Specification.Selenium_Makehtml("���p�p����-�ŏ���")
      Call Specification.Selenium_Makehtml("���p�p����")
      Call Specification.Selenium_Makehtml("���p�p����-�ő包")
      Call Specification.Selenium_Makehtml("���p�p����-�ő包���ȏ�")
      
      Call Specification.Selenium_Makehtml("���p�p����-�ŏ������ȉ�")
      Call Specification.Selenium_Makehtml("���p�p����-�ŏ���")
      Call Specification.Selenium_Makehtml("���p�p����")
      Call Specification.Selenium_Makehtml("���p�p����-�ő包")
      Call Specification.Selenium_Makehtml("���p�p����-�ő包���ȏ�")
      
      Call Specification.Selenium_Makehtml("���p�L��-�ŏ������ȉ�")
      Call Specification.Selenium_Makehtml("���p�L��-�ŏ���")
      Call Specification.Selenium_Makehtml("���p�L��")
      Call Specification.Selenium_Makehtml("���p�L��-�ő包")
      Call Specification.Selenium_Makehtml("���p�L��-�ő包���ȏ�")
      
      Call Specification.Selenium_Makehtml("���p����-�ŏ������ȉ�")
      Call Specification.Selenium_Makehtml("���p����-�ŏ���")
      Call Specification.Selenium_Makehtml("���p����")
      Call Specification.Selenium_Makehtml("���p����-�ő包")
      Call Specification.Selenium_Makehtml("���p����-�ő包���ȏ�")
 
      Call Specification.Selenium_Makehtml("�S�p����-�ŏ������ȉ�")
      Call Specification.Selenium_Makehtml("�S�p����-�ŏ���")
      Call Specification.Selenium_Makehtml("�S�p����")
      Call Specification.Selenium_Makehtml("�S�p����-�ő包")
      Call Specification.Selenium_Makehtml("�S�p����-�ő包���ȏ�")
      
      Call Specification.Selenium_Makehtml("�S�p�p������-�ŏ������ȉ�")
      Call Specification.Selenium_Makehtml("�S�p�p������-�ŏ���")
      Call Specification.Selenium_Makehtml("�S�p�p������")
      Call Specification.Selenium_Makehtml("�S�p�p������-�ő包")
      Call Specification.Selenium_Makehtml("�S�p�p������-�ő包���ȏ�")
      
      Call Specification.Selenium_Makehtml("�S�p�p�啶��-�ŏ������ȉ�")
      Call Specification.Selenium_Makehtml("�S�p�p�啶��-�ŏ���")
      Call Specification.Selenium_Makehtml("�S�p�p�啶��")
      Call Specification.Selenium_Makehtml("�S�p�p�啶��-�ő包")
      Call Specification.Selenium_Makehtml("�S�p�p�啶��-�ő包���ȏ�")
      
      Call Specification.Selenium_Makehtml("�S�p�p����-�ŏ������ȉ�")
      Call Specification.Selenium_Makehtml("�S�p�p����-�ŏ���")
      Call Specification.Selenium_Makehtml("�S�p�p����")
      Call Specification.Selenium_Makehtml("�S�p�p����-�ő包")
      Call Specification.Selenium_Makehtml("�S�p�p����-�ő包���ȏ�")
      
      Call Specification.Selenium_Makehtml("�S�p�p����-�ŏ������ȉ�")
      Call Specification.Selenium_Makehtml("�S�p�p����-�ŏ���")
      Call Specification.Selenium_Makehtml("�S�p�p����")
      Call Specification.Selenium_Makehtml("�S�p�p����-�ő包")
      Call Specification.Selenium_Makehtml("�S�p�p����-�ő包���ȏ�")
      
      Call Specification.Selenium_Makehtml("�S�p�L��-�ŏ������ȉ�")
      Call Specification.Selenium_Makehtml("�S�p�L��-�ŏ���")
      Call Specification.Selenium_Makehtml("�S�p�L��")
      Call Specification.Selenium_Makehtml("�S�p�L��-�ő包")
      Call Specification.Selenium_Makehtml("�S�p�L��-�ő包���ȏ�")
      
      Call Specification.Selenium_Makehtml("�S�p�Ђ炪��-�ŏ������ȉ�")
      Call Specification.Selenium_Makehtml("�S�p�Ђ炪��-�ŏ���")
      Call Specification.Selenium_Makehtml("�S�p�Ђ炪��")
      Call Specification.Selenium_Makehtml("�S�p�Ђ炪��-�ő包")
      Call Specification.Selenium_Makehtml("�S�p�Ђ炪��-�ő包���ȏ�")
      
      Call Specification.Selenium_Makehtml("�S�p�J�^�J�i-�ŏ������ȉ�")
      Call Specification.Selenium_Makehtml("�S�p�J�^�J�i-�ŏ���")
      Call Specification.Selenium_Makehtml("�S�p�J�^�J�i")
      Call Specification.Selenium_Makehtml("�S�p�J�^�J�i-�ő包")
      Call Specification.Selenium_Makehtml("�S�p�J�^�J�i-�ő包���ȏ�")
      
      Call Specification.Selenium_Makehtml("��p����-�ŏ������ȉ�")
      Call Specification.Selenium_Makehtml("��p����-�ŏ���")
      Call Specification.Selenium_Makehtml("��p����")
      Call Specification.Selenium_Makehtml("��p����-�ő包")
      Call Specification.Selenium_Makehtml("��p����-�ő包���ȏ�")
      
      Call Specification.Selenium_Makehtml("���p�J�^�J�i-�ŏ������ȉ�")
      Call Specification.Selenium_Makehtml("���p�J�^�J�i-�ŏ���")
      Call Specification.Selenium_Makehtml("���p�J�^�J�i")
      Call Specification.Selenium_Makehtml("���p�J�^�J�i-�ő包")
      Call Specification.Selenium_Makehtml("���p�J�^�J�i-�ő包���ȏ�")

      Call Specification.Selenium_Makehtml("�S�p����-�ŏ������ȉ�")
      Call Specification.Selenium_Makehtml("�S�p����-�ŏ���")
      Call Specification.Selenium_Makehtml("�S�p����")
      Call Specification.Selenium_Makehtml("�S�p����-�ő包")
      Call Specification.Selenium_Makehtml("�S�p����-�ő包���ȏ�")
      
      Call Specification.Selenium_Makehtml("�@��ˑ�����")
        
      If InputDataType = "���t" Then
        Call Specification.Selenium_Makehtml("���t����01")
        Call Specification.Selenium_Makehtml("���t�ُ�01")
        Call Specification.Selenium_Makehtml("���t���ُ�01")
        Call Specification.Selenium_Makehtml("���t���ُ�01")
        
      
      ElseIf InputDataType = "email" Then
        Call Specification.Selenium_Makehtml("���[���A�h���X����01")
        Call Specification.Selenium_Makehtml("���[���A�h���X����02")
        Call Specification.Selenium_Makehtml("���[���A�h���X����03")
        Call Specification.Selenium_Makehtml("���[���A�h���X����04")
        Call Specification.Selenium_Makehtml("���[���A�h���X���[�J�����ُ�")
        Call Specification.Selenium_Makehtml("���[���A�h���X�ُ�")
      End If
      
    ElseIf Range("J" & line) = "�o�^/�����{�^��" Then
      Call Specification.Selenium_MakehtmlFooter(line)
      Call Specification.Selenium_MakeIndex
    End If
  Next
End Function


'***********************************************************************************************************************************************
' * �݌v���pSelenium
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function Selenium_Makehtml(MakeType As String)

  Dim htmlTag As String
  Dim L_InputLimit As Long



  htmlTag = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbLf & _
              "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">" & vbLf & _
              "<html xmlns=""http://www.w3.org/1999/xhtml"" xml:lang=""ja"" lang=""ja"">" & vbLf & _
              "<head profile=""http://selenium-ide.openqa.org/profiles/test-case"">" & vbLf & _
              "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"" />" & vbLf & _
              "<link rel=""selenium.base"" href=""" & Range("BaseURL") & """ />" & vbLf & _
              "<title>" & InputName & "</title>" & vbLf & _
              "</head>" & vbLf & _
              "<body>" & vbLf & _
              "<table cellpadding='1' cellspacing='1' border='1'>" & vbLf & _
              "<thead>" & vbLf & _
              "<tr><td rowspan='1' colspan='3'></td></tr>" & vbLf & _
              "</thead><tbody>" & vbLf
  
    htmlTag = htmlTag & "<!--��" & InputName & " " & MakeType & "-->" & vbLf
    htmlTag = htmlTag & "<tr>" & vbLf
    htmlTag = htmlTag & "  <td>open</td>" & vbLf
    htmlTag = htmlTag & "  <td>" & URL & "</td>" & vbLf
    htmlTag = htmlTag & "  <td></td>" & vbLf
    htmlTag = htmlTag & "</tr>" & vbLf

    Select Case MakeType
'=====================================================================================================================================
      Case "���p����-�ŏ������ȉ�", "���p����-�ŏ���", "���p����", "���p����-�ő包", "���p����-�ő包���ȏ�"
        InputTestString = HalfWidthDigit
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "���p�p������-�ŏ������ȉ�", "���p�p������-�ŏ���", "���p�p������", "���p�p������-�ő包", "���p�p������-�ő包���ȏ�"
        InputTestString = StrConv(HalfWidthCharacters, vbLowerCase)
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "���p�p�啶��-�ŏ������ȉ�", "���p�p�啶��-�ŏ���", "���p�p�啶��", "���p�p�啶��-�ő包", "���p�p�啶��-�ő包���ȏ�"
        InputTestString = HalfWidthCharacters
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "���p�p����-�ŏ������ȉ�", "���p�p����-�ŏ���", "���p�p����", "���p�p����-�ő包", "���p�p����-�ő包���ȏ�"
        InputTestString = HalfWidthCharacters & StrConv(HalfWidthCharacters, vbLowerCase) & HalfWidthDigit
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "���p�L��-�ŏ������ȉ�", "���p�L��-�ŏ�����", "���p�L��", "���p�L��-�ő包", "���p�L��-�ő包���ȏ�"
        InputTestString = SymbolCharacters
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "���p����-�ŏ������ȉ�", "���p����-�ŏ�����", "���p����", "���p����-�ő包", "���p����-�ő包���ȏ�"
        InputTestString = StrConv(HalfWidthCharacters, vbLowerCase) & _
                          HalfWidthCharacters & _
                          HalfWidthDigit & _
                          SymbolCharacters
'=====================================================================================================================================
      Case "�S�p����-�ŏ������ȉ�", "�S�p����-�ŏ���", "�S�p����", "�S�p����-�ő包", "�S�p����-�ő包���ȏ�"
        InputTestString = StrConv(HalfWidthDigit, vbWide)
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "�S�p�p������-�ŏ������ȉ�", "�S�p�p������-�ŏ���", "�S�p�p������", "�S�p�p������-�ő包", "�S�p�p������-�ő包���ȏ�"
        InputTestString = StrConv(StrConv(HalfWidthCharacters, vbLowerCase), vbWide)
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "�S�p�p�啶��-�ŏ������ȉ�", "�S�p�p�啶��-�ŏ���", "�S�p�p�啶��", "�S�p�p�啶��-�ő包", "�S�p�p�啶��-�ő包���ȏ�"
        InputTestString = StrConv(HalfWidthCharacters, vbWide)
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "�S�p�p����-�ŏ������ȉ�", "�S�p�p����-�ŏ���", "�S�p�p����", "�S�p�p����-�ő包", "�S�p�p����-�ő包���ȏ�"
        InputTestString = StrConv(StrConv(HalfWidthCharacters, vbLowerCase), vbWide) & _
                          StrConv(HalfWidthCharacters, vbWide)
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "�S�p�p����-�ŏ������ȉ�", "�S�p�p����-�ŏ���", "�S�p�p����", "�S�p�p����-�ő包", "�S�p�p����-�ő包���ȏ�"
        InputTestString = StrConv(HalfWidthDigit, vbWide) & _
                          StrConv(StrConv(HalfWidthCharacters, vbLowerCase), vbWide) & _
                          StrConv(HalfWidthCharacters, vbWide)
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "�S�p�L��-�ŏ������ȉ�", "�S�p�L��-�ŏ�����", "�S�p�L��", "�S�p�L��-�ő包", "�S�p�L��-�ő包���ȏ�"
        InputTestString = StrConv(SymbolCharacters, vbWide)

'=====================================================================================================================================
      Case "�S�p�Ђ炪��-�ŏ������ȉ�", "�S�p�Ђ炪��-�ŏ�����", "�S�p�Ђ炪��", "�S�p�Ђ炪��-�ő包", "�S�p�Ђ炪��-�ő包���ȏ�"
        InputTestString = JapaneseCharacters
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "�S�p�J�^�J�i-�ŏ������ȉ�", "�S�p�J�^�J�i-�ŏ�����", "�S�p�J�^�J�i", "�S�p�J�^�J�i-�ő包", "�S�p�J�^�J�i-�ő包���ȏ�"
        InputTestString = StrConv(JapaneseCharacters, vbKatakana)
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "���p�J�^�J�i-�ŏ������ȉ�", "���p�J�^�J�i-�ŏ�����", "���p�J�^�J�i", "���p�J�^�J�i-�ő包", "���p�J�^�J�i-�ő包���ȏ�"
        InputTestString = StrConv(StrConv(JapaneseCharacters, vbKatakana), vbNarrow)
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "��p����-�ŏ������ȉ�", "��p����-�ŏ�����", "��p����", "��p����-�ő包", "��p����-�ő包���ȏ�"
        InputTestString = JapaneseCharactersCommonUse
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "�S�p����-�ŏ������ȉ�", "�S�p����-�ŏ�����", "�S�p����", "�S�p����-�ő包", "�S�p����-�ő包���ȏ�"
        InputTestString = StrConv(HalfWidthDigit, vbWide) & _
                          StrConv(StrConv(HalfWidthCharacters, vbLowerCase), vbWide) & _
                          StrConv(HalfWidthCharacters, vbWide) & _
                          StrConv(SymbolCharacters, vbWide) & _
                          JapaneseCharacters & _
                          StrConv(JapaneseCharacters, vbKatakana) & _
                          StrConv(StrConv(JapaneseCharacters, vbKatakana), vbNarrow) & _
                          JapaneseCharactersCommonUse
                          
'=====================================================================================================================================
      Case "���t����01"
        InputTestString = "2016/01/01"
      
      Case "���t-�ُ�01"
        InputTestString = "2016/0101/"
      
      Case "���t���ُ�01"
        InputTestString = "2016/15/01"
      
      Case "���t���ُ�01"
        InputTestString = "2016/01/55"

'=====================================================================================================================================
      Case "���[���A�h���X-����01"
        InputTestString = "vb.project@vb-project.com"
        
      Case "���[���A�h���X-����02"
        InputTestString = "user+mailbox/department=shipping@vb-project.com"
        
      Case "���[���A�h���X-����03"
        InputTestString = """Joe.\\Blow""@vb-project.com"
        
      Case "���[���A�h���X-����04"
        InputTestString = "1234567890123456789012345678901234567890123456789012345678901234@abcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvyzab.co.jp"
        
      Case "���[���A�h���X-���[�J�����ُ�"
        InputTestString = "vb..project@vb-project.com"
        
      Case "���[���A�h���X�ُ�"
        InputTestString = "1234567890123456789012345678901234567890123456789012345678901234@abcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvyzabcdefghijklmnopqrstuvyz.co.jp"

'=====================================================================================================================================
      Case Else
        InputTestString = StrConv(HalfWidthCharacters, vbLowerCase) & _
                          HalfWidthCharacters & _
                          HalfWidthDigit & _
                          SymbolCharacters & _
                          StrConv(HalfWidthDigit, vbWide) & _
                          StrConv(StrConv(HalfWidthCharacters, vbLowerCase), vbWide) & _
                          StrConv(HalfWidthCharacters, vbWide) & _
                          StrConv(SymbolCharacters, vbWide) & _
                          MachineDependentCharacters
    End Select
    
    '�ő啶�����̐ݒ�
    If InStr(MakeType, "-") <> 0 Then
      MakeType_tmp = Split(MakeType, "-")
      Select Case MakeType_tmp(1)
        Case "�ŏ������ȉ�"
          L_InputLimit = InputLimitMin - 1
        
        Case "�ŏ�����"
          L_InputLimit = InputLimitMin
          
        Case "�ő包��"
          If InputLimitMax = 0 Then
            L_InputLimit = 0
          Else
            L_InputLimit = InputLimitMax
          End If
        Case "�ő包���ȏ�"
          If InputLimitMax = 0 Then
            L_InputLimit = 0
          Else
            L_InputLimit = InputLimitMax + 1
          End If
        Case Else
          L_InputLimit = InputLimitMax
      End Select
    Else
      If InputLimitMax = 0 Then
        L_InputLimit = 0
      Else
        L_InputLimit = InputLimitMin + 1
      End If
    End If
    
    If L_InputLimit = 0 Then
      Exit Function
    End If
    
    
    '���͕����������_���ɐݒ�
    InputTestString = call Library.Randomize(InputTestString, L_InputLimit)
    
    htmlTag = htmlTag & "<!-- " & MakeType & "-->" & vbLf
    htmlTag = htmlTag & "<tr>" & vbLf
    htmlTag = htmlTag & "  <td>type</td>" & vbLf
    htmlTag = htmlTag & "  <td>" & InputNameTag & "</td>" & vbLf
    htmlTag = htmlTag & "  <td>" & InputTestString & "</td>" & vbLf
    htmlTag = htmlTag & "</tr>" & vbLf

    '=======================================================================================
    '�f�B���N�g���쐬
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    SeleniumFolder = ActiveWorkbook.path & "\" & Title
    If objFSO.FolderExists(folderspec:=SeleniumFolder) = False Then
      objFSO.CreateFolder SeleniumFolder
    End If
    Set objFSO = Nothing


    ' ADODB����
    Set ObjADODB_TestCase = CreateObject("ADODB.Stream")
  
    '�I�u�W�F�N�g�ɕۑ�����f�[�^�̎�ނ𕶎���^�Ɏw�肷��
    ObjADODB_TestCase.Type = 2
    
    '������^�̃I�u�W�F�N�g�̕����R�[�h���w�肷��
    ObjADODB_TestCase.Charset = "UTF-8"
    ObjADODB_TestCase.LineSeparator = 10
  
    '�I�u�W�F�N�g�̃C���X�^���X���쐬
    ObjADODB_TestCase.Open
    

    ' �e�X�g�P�[�X�ۑ�
    ObjADODB_TestCase.WriteText htmlTag, 1
    ObjADODB_TestCase.SaveToFile (SeleniumFolder & "\" & InputNo & "_" & InputName & "_" & MakeType & ".html"), 2
   
    '�I�u�W�F�N�g�����
    ObjADODB_TestCase.Close
    Set ObjADODB_TestCase = Nothing
End Function


Function Selenium_MakehtmlFooter(ByVal line As Long)

  Dim htmlTag As String
  '=======================================================================================
  '�f�B���N�g���쐬
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  
  captureFolder = ActiveWorkbook.path & "\" & Title
  If objFSO.FolderExists(folderspec:=captureFolder) = False Then
    objFSO.CreateFolder captureFolder
  End If
  captureFolder = ActiveWorkbook.path & "\" & Title & "\capture"
  If objFSO.FolderExists(folderspec:=captureFolder) = False Then
    objFSO.CreateFolder captureFolder
  End If
  
  
  Set objFSO = Nothing
  
  Dim buf As String
  buf = Dir(ActiveWorkbook.path & "\" & Title & "\*.html")
  
  Do While Len(buf) > 0
    If LCase(buf) Like "*.html" Then
      
      htmlTag = "<tr>" & vbLf
      htmlTag = htmlTag & "  <td>captureEntirePageScreenshot</td>" & vbLf
      htmlTag = htmlTag & "  <td>" & ActiveWorkbook.path & "\" & Title & "\capture\" & buf & "01.png</td>" & vbLf
      htmlTag = htmlTag & "  <td>background=#FFFFFF</td>" & vbLf
      htmlTag = htmlTag & "</tr>" & vbLf
      htmlTag = htmlTag & "<tr>" & vbLf
      
      If (Range("AF" & line) <> "") Then
        htmlTag = htmlTag & "  <td>runScriptAndWait</td>" & vbLf
        htmlTag = htmlTag & "  <td>" & Range("AF" & line) & "</td>" & vbLf
      Else
        htmlTag = htmlTag & "  <td>clickAndWait</td>" & vbLf
        htmlTag = htmlTag & "  <td>" & Range("S" & line) & "</td>" & vbLf
      End If
      htmlTag = htmlTag & "  <td></td>" & vbLf
      htmlTag = htmlTag & "</tr>" & vbLf
      htmlTag = htmlTag & "<tr>" & vbLf
      htmlTag = htmlTag & "  <td>captureEntirePageScreenshot</td>" & vbLf
      htmlTag = htmlTag & "  <td>" & ActiveWorkbook.path & "\" & Title & "\capture\" & buf & "02.png</td>" & vbLf
      htmlTag = htmlTag & "  <td>background=#FFFFFF</td>" & vbLf
      htmlTag = htmlTag & "</tr>" & vbLf
      htmlTag = htmlTag & "</tbody></table></body></html>" & vbLf
  
  
    ' ADODB����
    Set ObjADODB_TestCase = CreateObject("ADODB.Stream")
  
    '�I�u�W�F�N�g�ɕۑ�����f�[�^�̎�ނ𕶎���^�Ɏw�肷��
    ObjADODB_TestCase.Type = 2
    
    '������^�̃I�u�W�F�N�g�̕����R�[�h���w�肷��
    ObjADODB_TestCase.Charset = "UTF-8"
    ObjADODB_TestCase.LineSeparator = 10
  
    '�I�u�W�F�N�g�̃C���X�^���X���쐬
    ObjADODB_TestCase.Open
    ObjADODB_TestCase.LoadFromFile (ActiveWorkbook.path & "\" & Title & "\" & buf)
    ObjADODB_TestCase.Position = ObjADODB_TestCase.Size '�|�C���^���I�[��


    ' �e�X�g�P�[�X�ۑ�
    ObjADODB_TestCase.WriteText htmlTag, 1
    ObjADODB_TestCase.SaveToFile (ActiveWorkbook.path & "\" & Title & "\" & buf), 2
   
    '�I�u�W�F�N�g�����
    ObjADODB_TestCase.Close
    Set ObjADODB_TestCase = Nothing
  
    End If
    buf = Dir()
  Loop

End Function


Function Selenium_MakeIndex()

  Dim htmlTag As String


  Dim buf As String
  htmlTag = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbLf & _
              "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">" & vbLf & _
              "<html xmlns=""http://www.w3.org/1999/xhtml"" xml:lang=""ja"" lang=""ja"">" & vbLf & _
              "<head profile=""http://selenium-ide.openqa.org/profiles/test-case"">" & vbLf & _
              "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"" />" & vbLf & _
              "<link rel=""selenium.base"" href=""" & baseURL & """ />" & vbLf & _
              "<title>" & TitleName & "</title>" & vbLf & _
              "</head>" & vbLf & _
              "<body>" & vbLf & _
              "<table id='suiteTable' cellpadding='1' cellspacing='1' border='1' class='selenium'><tbody>" & vbLf
              
              
  buf = Dir(ActiveWorkbook.path & "\" & Title & "\*.html")
  
  Do While Len(buf) > 0
    If LCase(buf) Like "*.html" Then
      If buf <> "00_index.html" Then
        htmlTag = htmlTag & "<tr><td><a href='" & buf & "'>" & buf & "</a></td></tr>" & vbLf
      End If
    End If
    buf = Dir()
  Loop
  htmlTag = htmlTag & "</tbody></table></body></html>" & vbLf
  
    ' ADODB����
    Set ObjADODB_TestCase = CreateObject("ADODB.Stream")
  
    '�I�u�W�F�N�g�ɕۑ�����f�[�^�̎�ނ𕶎���^�Ɏw�肷��
    ObjADODB_TestCase.Type = 2
    
    '������^�̃I�u�W�F�N�g�̕����R�[�h���w�肷��
    ObjADODB_TestCase.Charset = "UTF-8"
    ObjADODB_TestCase.LineSeparator = 10
  
    '�I�u�W�F�N�g�̃C���X�^���X���쐬
    ObjADODB_TestCase.Open
    

    ' �e�X�g�P�[�X�ۑ�
    ObjADODB_TestCase.WriteText htmlTag, 1
    ObjADODB_TestCase.SaveToFile (ActiveWorkbook.path & "\" & Title & "\00_index.html"), 2
   
    '�I�u�W�F�N�g�����
    ObjADODB_TestCase.Close
    Set ObjADODB_TestCase = Nothing
End Function


