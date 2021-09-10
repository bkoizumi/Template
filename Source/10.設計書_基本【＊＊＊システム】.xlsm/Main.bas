Attribute VB_Name = "Main"
' *********************************************************************
' * ページ追加
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
' *********************************************************************
Sub addPage(control As IRibbonControl)
  Call Library.startScript
  
  If ActiveCell.Value = "システム" Then
    Call Specification.addPage(ActiveCell.Address)
  Else
    Call Specification.addPage
  End If
  
  Call Library.endScript
  
  ThisWorkbook.Activate
End Sub

' *********************************************************************
' * 目次生成
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
' *********************************************************************
Sub MakeMenu(control As IRibbonControl)
  Call Library.startScript
  Call Specification.makeTOC
  Call Library.endScript
  
  ThisWorkbook.Activate
End Sub



' *********************************************************************
' * 印刷範囲設定
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
' *********************************************************************
Sub SetPrintArea(control As IRibbonControl)
  Call Library.startScript
  Call Ctl_ProgressBar.ShowStart
  
  Call Specification.SetPrintArea
  
  Call Ctl_ProgressBar.ShowEnd
  Call Library.endScript
End Sub

  
'==================================================================================================
'タイトル一覧
Function getMenuList(control As IRibbonControl, ByRef returnedVal)
  Dim DOMDoc As Object, Menu As Object, Button As Object, FunctionMenu As Object
  Dim sheetName As Worksheet
  Dim MenuSepa, sheetNameID
  Dim line As Long, endLine As Long
  Dim TitleName As String, FunctionName As String, tocTitle As String
  
  '処理開始--------------------------------------
  'On Error GoTo catchError
  FuncName = "Main.getMenuList"

  Call Library.startScript
  Call init.setting(True)
  'Call Library.showDebugForm(FuncName & "============================================")
  sheetMain.Select
  '----------------------------------------------
  endLine = sheetMain.Cells(Rows.count, 50).End(xlUp).Row
  
  
  Set DOMDoc = CreateObject("Msxml2.DOMDocument")
  Set Menu = DOMDoc.createElement("menu")

  Menu.SetAttribute "xmlns", "http://schemas.microsoft.com/office/2009/07/customui"
  Menu.SetAttribute "itemSize", "normal"

  For line = 44 To endLine Step setVal("PageLine")
    TitleName = Cells(line + 1, 4)
    FunctionName = Cells(line + 1, 19)
    
    If FunctionName <> "" Then
      tocTitle = TitleName & " - " & FunctionName
    Else
      tocTitle = TitleName
    End If
    
    If tocTitle Like "[目次,もくじ]*" Then
    Else
      Set Button = DOMDoc.createElement("button")
      With Button
        .SetAttribute "id", "CellID_" & line
        .SetAttribute "label", TitleName
        .SetAttribute "onAction", "Main.selectActiveCell"
      End With
      Menu.AppendChild Button
      Set Button = Nothing
    End If
  Next
  DOMDoc.AppendChild Menu
  'Debug.Print DOMDoc.XML
  
  returnedVal = DOMDoc.XML
  Set Menu = Nothing
  Set DOMDoc = Nothing


  '処理終了--------------------------------------
  Call Library.endScript
  '----------------------------------------------
  
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  
End Function

'==================================================================================================
Function selectActiveCell(control As IRibbonControl)
  Dim sheetNameID As Integer
  Dim sheetCount As Integer
  Dim sheetName As Worksheet
  Dim line
  
  Call Library.startScript
  line = Replace(control.id, "CellID_", "")
  
  Application.Goto Reference:=Range("A" & line), Scroll:=True
  Call Library.endScript
End Function
