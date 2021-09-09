VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FProgress 
   Caption         =   "èàóùíÜ"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5565
   OleObjectBlob   =   "FProgress.frx":0000
   StartUpPosition =   3  'Windows ÇÃä˘íËíl
End
Attribute VB_Name = "FProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

