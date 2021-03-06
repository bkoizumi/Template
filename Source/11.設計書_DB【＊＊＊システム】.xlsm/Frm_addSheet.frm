VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_addSheet 
   Caption         =   "テーブル情報追加"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6570
   OleObjectBlob   =   "Frm_addSheet.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_addSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************************************************************************************
' * 初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  Dim ListIndex As Integer
  Dim line As Long, endLine As Long
  
  Call init.Setting
  endLine = sheetSetting.Cells(Rows.count, 7).End(xlUp).Row
  
  For line = 5 To endLine
    DBType.AddItem sheetSetting.Range("G" & line).text
  Next
  DBType.ListIndex = 0

End Sub

'**************************************************************************************************
' * ボタン押下時処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
'キャンセル処理
Private Sub Cancel_Click()
  Unload Me
  
  Call Library.endScript
  End
End Sub


'==================================================================================================
' 実行
Private Sub Submit_Click()
  Dim execDay As Date

  sheetCopyTable.Range("F8") = TableName01.text
  sheetCopyTable.Range("F9") = TableName02.text
  sheetCopyTable.Range("F10") = DBType.Value
  sheetCopyTable.Range("F11") = Comment.text
  
  Unload Me
End Sub

  

