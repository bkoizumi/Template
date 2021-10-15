VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_log 
   Caption         =   "デバッグ情報"
   ClientHeight    =   3165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12495
   OleObjectBlob   =   "Frm_log.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'**************************************************************************************************
' * 初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()

'  Select Case setVal("LogLevel")
'    Case "warning"
'      LV_warning = True
'    Case "notice"
'      LV_notice = True
'    Case "info"
'      LV_info = True
'    Case "debug"
'      LV_debug = True
'  End Select


End Sub


'**************************************************************************************************
' * ボタン押下時処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Private Sub LV_warning_Click()

End Sub


'==================================================================================================
'キャンセル処理
Private Sub CancelButton_Click()

  Unload Me
End Sub



