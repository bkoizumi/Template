VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_TableList 
   Caption         =   "テーブル情報"
   ClientHeight    =   3615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360
   OleObjectBlob   =   "Frm_TableList.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_TableList"
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
  Dim ListIndex As Integer
  Dim line As Long, endLine As Long, i As Long
  Dim objShp As Shape
  
  If setVal("usePhysicalName") = True Then
    usePhysicalName.Value = True
  Else
    useLogicalName.Value = True
  End If
  useImage.Value = setVal("useImage")
  
  With ListBox1
    .ColumnHeads = True
    .ColumnCount = 4
    .ColumnWidths = "0,80;150;150"
    .RowSource = sheetTmp.Range("J2:M" & sheetTmp.Cells(Rows.count, 10).End(xlUp).Row).Address(External:=True)
    
'    For i = 0 To .ListCount - 1
'      For Each objShp In ActiveSheet.Shapes
'        Call Library.showDebugForm("objShp.Name", objShp.Name, "info")
'        Call Library.showDebugForm(".list(" & i & ", 1)", .list(i, 1), "info")
'        Call Library.showDebugForm(".list(" & i & ", 0)", .list(i, 0), "info")
'
'        If objShp.Name = "ERImg-" & .list(i, 1) Then
'          .Selected(i) = True
'          Exit For
'        End If
'      Next
'    Next
  End With

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
  
End Sub


'==================================================================================================
' 実行
Private Sub Submit_Click()
  Dim i As Integer
  
  Call init.Setting
  useLogicalName = useLogicalName.Value
  usePhysicalName = usePhysicalName.Value
  
  Call Library.setValandRange("useLogicalName", CStr(useLogicalName.Value))
  Call Library.setValandRange("usePhysicalName", CStr(usePhysicalName.Value))
  Call Library.setValandRange("useImage", CStr(useImage.Value))
  
  With ListBox1
  For i = 0 To .ListCount - 1
    If .Selected(i) = True Then
      Call Library.showDebugForm("リスト0", .list(i, 0))
      Call Library.showDebugForm("リスト1", .list(i, 1))
      Call Library.showDebugForm("リスト2", .list(i, 2))
      Call Library.showDebugForm("リスト3", .list(i, 3))
      
      Call Ctl_ErImg.getColumnInfo(.list(i, 0))
      Call Ctl_ErImg.makeERImage
      Call Ctl_ErImg.copy
    End If
  Next i
End With
  
  
  
  
  Unload Me
End Sub

  


