Attribute VB_Name = "init"
Option Explicit


'ワークブック用変数------------------------------
Public ThisBook       As Workbook
Public targetBook     As Workbook


'ワークシート用変数------------------------------
Public targetsheet    As Worksheet

Public sheetSetting   As Worksheet
Public sheetNotice    As Worksheet
Public sheetCopy      As Worksheet
Public sheetMain      As Worksheet

'グローバル変数----------------------------------
Public Const thisAppName    As String = "設計書"
Public Const thisAppVersion As String = "V1.0-beta.1"
Public FuncName             As String
Public logFile              As String

'設定値保持
Public setVal          As Object



'**************************************************************************************************
' * 設定解除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function usetting()

  Set ThisBook = Nothing
  
  'ワークシート名の設定
  Set sheetSetting = Nothing
  Set sheetNotice = Nothing
  Set sheetCopy = Nothing
  Set sheetMain = Nothing

  '設定値読み込み
  Set setVal = Nothing
End Function


'**************************************************************************************************
' * 設定
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

  'ブックの設定
  Set ThisBook = ThisWorkbook
  
  'ワークシート名の設定
  Set sheetSetting = ThisBook.Worksheets("設定")
  Set sheetNotice = ThisBook.Worksheets("Notice")
  Set sheetCopy = ThisBook.Worksheets("Copy")
  Set sheetMain = ThisBook.Worksheets("設計書")
 
  
        
  '設定値読み込み----------------------------------------------------------------------------------
  Set setVal = Nothing
  Set setVal = CreateObject("Scripting.Dictionary")
  
  For line = 5 To sheetSetting.Cells(Rows.count, 1).End(xlUp).Row
    If sheetSetting.Range("A" & line) <> "" Then
      setVal.add sheetSetting.Range("A" & line).Text, sheetSetting.Range("B" & line).Text
    End If
  Next
  
  logFile = ThisWorkbook.Path & "\ExcelMacro.log"
  
  Call 名前定義
  
  Exit Function
  
'エラー発生時=====================================================================================
catchError:
  
End Function


'**************************************************************************************************
' * 名前定義
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 名前定義()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim Name As Object
  
'  On Error GoTo catchError

  '名前の定義を削除
  For Each Name In Names
    If Name.Visible = False Then
      Name.Visible = True
    End If
    If Not Name.Name Like "[Print_,Slc,Pvt,Tbl,改訂]*" Then
      Name.delete
    End If
  Next
  
  'VBA用の設定
  For line = 5 To sheetSetting.Cells(Rows.count, 1).End(xlUp).Row
    If sheetSetting.Range("A" & line) <> "" Then
      sheetSetting.Range("B" & line).Name = sheetSetting.Range("A" & line)
    End If
  Next
  
  'Book用の設定
  For line = 5 To sheetSetting.Cells(Rows.count, 4).End(xlUp).Row
    If sheetSetting.Range("D" & line) <> "" Then
      sheetSetting.Range("E" & line).Name = sheetSetting.Range("D" & line)
    End If
  Next
  

  Exit Function
'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
  
End Function

