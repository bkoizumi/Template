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
  endLine = SettingSheet.Cells(Rows.count, 7).End(xlUp).Row
  
  With Me
    For line = 3 To endLine
      .DBType.AddItem SettingSheet.Range("G" & line).Text
    Next
  End With

End Sub

'**************************************************************************************************
' * �{�^������������
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'�L�����Z������
Private Sub Cancel_Click()
  Range("Frm_NewSheetTop") = Me.Top
  Range("Frm_NewSheetLeft") = Me.Left

  Unload Me
  
  Call Library.endScript
  End
End Sub


'==================================================================================================
' ���s
Private Sub Submit_Click()
  Dim execDay As Date
  
  Range("Frm_NewSheetTop") = Me.Top
  Range("Frm_NewSheetLeft") = Me.Left

  CopySheet.Range("D5") = Me.TableName01.Text
  CopySheet.Range("H5") = Me.TableName02.Text
  CopySheet.Range("D6") = Me.Comment.Text
  CopySheet.Range("B2") = Me.DBType.Value
  
  Unload Me
End Sub

  

