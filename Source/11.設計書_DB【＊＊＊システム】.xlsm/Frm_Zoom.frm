VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Zoom 
   Caption         =   "ï“èW"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7050
   OleObjectBlob   =   "Frm_Zoom.frx":0000
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
End
Attribute VB_Name = "Frm_Zoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'==================================================================================================
Private Sub CancelButton_Click()
  
'  Call Ctl_DefaultVal.delVal("ZoomIn")
  Unload Me
End Sub

'==================================================================================================
Private Sub OK_Button_Click()
  
  Call Ctl_Zoom.ZoomOut(TextBox, Frm_Zoom.Label1.Caption)
  Unload Me
End Sub



'==================================================================================================
Private Sub TextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
'  Call Ctl_DefaultVal.setVal("reSetZoomIn", TextBox.Text)
End Sub


