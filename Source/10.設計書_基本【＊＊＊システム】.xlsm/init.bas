Attribute VB_Name = "init"
Option Explicit


'���[�N�u�b�N�p�ϐ�------------------------------
Public ThisBook       As Workbook
Public targetBook     As Workbook


'���[�N�V�[�g�p�ϐ�------------------------------
Public targetsheet    As Worksheet

Public sheetSetting   As Worksheet
Public sheetNotice    As Worksheet
Public sheetCopy      As Worksheet
Public sheetMain      As Worksheet

'�O���[�o���ϐ�----------------------------------
Public Const thisAppName    As String = "�݌v��"
Public Const thisAppVersion As String = "V1.0-beta.1"
Public FuncName             As String
Public logFile              As String

'�ݒ�l�ێ�
Public setVal          As Object



'**************************************************************************************************
' * �ݒ����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function usetting()

  Set ThisBook = Nothing
  
  '���[�N�V�[�g���̐ݒ�
  Set sheetSetting = Nothing
  Set sheetNotice = Nothing
  Set sheetCopy = Nothing
  Set sheetMain = Nothing

  '�ݒ�l�ǂݍ���
  Set setVal = Nothing
End Function


'**************************************************************************************************
' * �ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setting(Optional reCheckFlg As Boolean)
  Dim line As Long, endLine As Long
  
  On Error GoTo catchError
  ThisWorkbook.Save

  If ThisBook Is Nothing Or reCheckFlg = True Then
    Call usetting
  Else
    Exit Function
  End If

  '�u�b�N�̐ݒ�
  Set ThisBook = ThisWorkbook
  
  '���[�N�V�[�g���̐ݒ�
  Set sheetSetting = ThisBook.Worksheets("�ݒ�")
  Set sheetNotice = ThisBook.Worksheets("Notice")
  Set sheetCopy = ThisBook.Worksheets("Copy")
  Set sheetMain = ThisBook.Worksheets("�݌v��")
 
  
        
  '�ݒ�l�ǂݍ���----------------------------------------------------------------------------------
  Set setVal = Nothing
  Set setVal = CreateObject("Scripting.Dictionary")
  
  For line = 5 To sheetSetting.Cells(Rows.count, 1).End(xlUp).Row
    If sheetSetting.Range("A" & line) <> "" Then
      setVal.add sheetSetting.Range("A" & line).Text, sheetSetting.Range("B" & line).Text
    End If
  Next
  
  logFile = ThisWorkbook.Path & "\ExcelMacro.log"
  
  Call ���O��`
  
  Exit Function
  
'�G���[������=====================================================================================
catchError:
  
End Function


'**************************************************************************************************
' * ���O��`
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ���O��`()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim Name As Object
  
'  On Error GoTo catchError

  '���O�̒�`���폜
  For Each Name In Names
    If Name.Visible = False Then
      Name.Visible = True
    End If
    If Not Name.Name Like "[Print_,Slc,Pvt,Tbl,����]*" Then
      Name.delete
    End If
  Next
  
  'VBA�p�̐ݒ�
  For line = 5 To sheetSetting.Cells(Rows.count, 1).End(xlUp).Row
    If sheetSetting.Range("A" & line) <> "" Then
      sheetSetting.Range("B" & line).Name = sheetSetting.Range("A" & line)
    End If
  Next
  
  'Book�p�̐ݒ�
  For line = 5 To sheetSetting.Cells(Rows.count, 4).End(xlUp).Row
    If sheetSetting.Range("D" & line) <> "" Then
      sheetSetting.Range("E" & line).Name = sheetSetting.Range("D" & line)
    End If
  Next
  

  Exit Function
'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
  
End Function

