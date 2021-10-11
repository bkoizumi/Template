VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_NewSheet 
   Caption         =   "�e�[�u�����ǉ�"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6570
   OleObjectBlob   =   "Frm_NewSheet.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Frm_NewSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************************************************************************************
' * �����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  Dim ListIndex As Integer
  Dim line As Long, endLine As Long
  
  Call init.Setting
  endLine = sheetSetting.Cells(Rows.count, 7).End(xlUp).Row
  
  For line = 5 To endLine
    DBType.AddItem sheetSetting.Range("G" & line).Text
  Next
  DBType.ListIndex = 0

End Sub

'**************************************************************************************************
' * �{�^������������
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'�L�����Z������
Private Sub Cancel_Click()
  Unload Me
  
  Call Library.endScript
  End
End Sub


'==================================================================================================
' ���s
Private Sub Submit_Click()
  Dim execDay As Date

  sheetCopy.Range(setVal("Cell_logicalTableName")) = TableName01.Text
  sheetCopy.Range(setVal("Cell_physicalTableName")) = TableName02.Text
  sheetCopy.Range(setVal("Cell_tableNote")) = Comment.Text
  sheetCopy.Range(setVal("Cell_TableType")) = DBType.Value
  
  Unload Me
End Sub

  

