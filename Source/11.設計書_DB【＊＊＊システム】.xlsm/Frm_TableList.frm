VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_TableList 
   Caption         =   "�e�[�u�����"
   ClientHeight    =   3615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360
   OleObjectBlob   =   "Frm_TableList.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Frm_TableList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'**************************************************************************************************
' * �����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  Dim ListIndex As Integer
  Dim line As Long, endLine As Long
  
  usePhysicalName.Value = True
  useImage.Value = True
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
  
End Sub


'==================================================================================================
' ���s
Private Sub Submit_Click()
  Dim i As Integer
  
  Call init.Setting
  useLogicalName = useLogicalName.Value
  usePhysicalName = usePhysicalName.Value
  
  setVal.Add "useImage", CStr(useImage.Value)
  
  With ListView1
  For i = 1 To .ListItems.count
    If .ListItems(i).Selected Then
      Call Library.showDebugForm("���X�g", .ListItems(i).Text)
      Call Library.showDebugForm("���X�g", .ListItems(i).SubItems(1))
      
      Call Ctl_ErImg.deleteImages(.ListItems(i).Text)
      If useLogicalName.Value = True Then
        Call Ctl_ErImg.makeTable(.ListItems(i).Text)
      Else
        Call Ctl_ErImg.makeTable(.ListItems(i).SubItems(1))
      End If
      
      Call Ctl_ErImg.makeColumnList(.ListItems(i).Text)
      Call Ctl_ErImg.copy(.ListItems(i).Text)
    End If
  Next i
End With
  
  
  
  
  Unload Me
End Sub

  


