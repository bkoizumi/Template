Attribute VB_Name = "init"
'ワークブック用変数------------------------------
Public ThisBook   As Workbook
Public targetBook As Workbook


'ワークシート用変数------------------------------
Public targetSheet    As Worksheet

Public sheetSetting   As Worksheet
Public sheetNotice    As Worksheet
Public sheetDataType  As Worksheet
Public sheetCopy      As Worksheet
Public sheetTblList   As Worksheet



'グローバル変数----------------------------------
Public Const thisAppName = "DB設計書"
Public Const thisAppVersion = "V1.0-beta.1"

Public ConnectServer      As String
Public Const startLine    As Long = 9
Public isDBOpen           As Boolean
Public runFlg             As Boolean

Public PrgP_Max           As Long
Public PrgP_Cnt           As Long

Public FuncName As String

Public accFileName        As String
Public accFileDir         As String
Public ArryTypeName(205)  As String
Public oldCellVal         As String

'レジストリ登録用サブキー
'Public Const RegistryKey  As String = "BK_Documents"


'設定値保持
Public setVal         As Object


'ファイル関連
Public logFile      As String

'処理時間計測用
Public StartTime          As Date
Public StopTime           As Date



'リボン関連--------------------------------------
Public ribbonUI       As Office.IRibbonUI
Public ribbonVal      As Object


'**************************************************************************************************
' * 設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function usetting(Optional flg As Boolean = True)

  Set ThisBook = Nothing
  
  Set sheetSetting = Nothing
  Set sheetNotice = Nothing
  Set sheetDataType = Nothing
  Set sheetCopy = Nothing
  Set sheetTblList = Nothing
  
  Set setVal = Nothing
  
  PrgP_Max = 0
  PrgP_Cnt = 0
  logFile = ""
  
  If flg = True Then
    runFlg = False
  End If
End Function


'==================================================================================================
Function Setting(Optional reCheckFlg As Boolean)
  Dim line As Long, endLine As Long
'  On Error GoTo catchError
'  ThisWorkbook.Save

  If logFile = "" Or reCheckFlg = True Then
    Call usetting(False)
  Else
    Exit Function
  End If

  'ブックの設定
  Set ThisBook = ThisWorkbook
  
  'ワークシート名の設定
  'Set sheetSetting = ThisBook.Worksheets("設定-MySQL")
  Set sheetSetting = ThisBook.Worksheets("設定-ACC")
  
  
  Set sheetNotice = ThisBook.Worksheets("Notice")
  Set sheetDataType = ThisBook.Worksheets("DataType")
  
  Set sheetCopy = ThisBook.Worksheets("コピー用")
  Set sheetTblList = ThisBook.Worksheets("TBLリスト")
  
  logFile = ThisWorkbook.Path & "\ExcelMacro.log"
        
  '設定値読み込み----------------------------------------------------------------------------------
  Set setVal = Nothing
  Set setVal = CreateObject("Scripting.Dictionary")
  
  For line = 5 To sheetSetting.Cells(Rows.count, 1).End(xlUp).Row
    If sheetSetting.Range("A" & line) <> "" Then
      setVal.Add sheetSetting.Range("A" & line).Text, sheetSetting.Range("B" & line).Text
    End If
  Next

  Select Case setVal("DBMS")
    Case "MSAccess"
      accFileName = Library.getFileInfo(setVal("DBServer"), , "fileName")
      accFileDir = Library.getFileInfo(setVal("DBServer"), , "CurrentDir")
      ConnectServer = "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                      "Data Source=" & setVal("DBServer") & ";" & _
                      "Jet OLEDB:Database Password=" & setVal("passwd") & ";"
                     
      Range("DBName") = accFileName
      
      endLine = sheetSetting.Cells(Rows.count, 12).End(xlUp).Row
      For line = 5 To endLine
        ArryTypeName(sheetSetting.Range("L" & line)) = sheetSetting.Range("M" & line)
      Next
    Case "MySQL"
      ConnectServer = "Driver={" & setVal("ODBCDriver") & "};" & _
                      " Server=" & setVal("DBServer") & ";" & _
                      " Port=" & setVal("Port") & ";" & _
                      " Database=" & setVal("DBName") & ";" & _
                      " User=" & setVal("userID") & ";" & _
                      " Password=" & setVal("passwd") & ";" & _
                      " Charset=sjis;"
    
    Case "PostgreSQL"
      ConnectServer = ""
      
    Case "SQLServer"
      ConnectServer = "Provider=SQLOLEDB;" & _
                      "Data Source=" & setVal("DBServer") & ";" & _
                      "Initial Catalog=" & setVal("DBName") & ";" & _
                      "Trusted_Connection=Yes"
  
  End Select
  
  
  
  Call 名前定義
  Exit Function
  
'エラー発生時--------------------------------------------------------------------------------------
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
   
  For Each Name In Names
    If Name.Visible = False Then
      Name.Visible = True
    End If
    If Not Name.Name Like "[Print_]*" Then
      If Name.Name Like "_xlfn*" Then
'        MsgBox "マクロでは削除できない名前の定義があります" & vbNewLine & Name.Name, vbExclamation
      Else
        Name.Delete
      End If
    End If
  Next
  
  'VBA用の設定
  For line = 3 To sheetSetting.Cells(Rows.count, 1).End(xlUp).Row
    If sheetSetting.Range("A" & line) <> "" Then
      sheetSetting.Range("B" & line).Name = sheetSetting.Range("A" & line)
    End If
  Next
  
  'Book用の設定
  For colLine = 7 To 10
    endLine = sheetSetting.Cells(Rows.count, colLine).End(xlUp).Row
    sheetSetting.Range(sheetSetting.Cells(5, colLine), sheetSetting.Cells(endLine, colLine)).Name = sheetSetting.Cells(4, colLine)
  Next
  
  'DataType用の設定
  For colLine = 1 To 15 Step 3
    endLine = sheetDataType.Cells(Rows.count, colLine).End(xlUp).Row
    sheetDataType.Range(sheetDataType.Cells(3, colLine), sheetDataType.Cells(endLine, colLine)).Name = sheetDataType.Cells(1, colLine)
  Next
  
  
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
  
End Function


'**************************************************************************************************
' * シートの表示/非表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function シート非表示()

  If setVal("debugMode") <> "develop" Then
    Worksheets("設定").Visible = xlSheetVeryHidden
    Worksheets("Notice").Visible = xlSheetVeryHidden
    Worksheets("DataType").Visible = xlSheetVeryHidden
  End If
  
  Worksheets("TBLリスト").Select
End Function


'==================================================================================================
Function シート表示()
  
  Worksheets("設定").Visible = True
  Worksheets("Notice").Visible = True
  Worksheets("DataType").Visible = True
  
  Worksheets("TBLリスト").Select
  
End Function


'==================================================================================================
Function シート保護()
  Dim sheetName As String
  Dim tempSheet As Object

  Call init.Setting
  Call Library.showDebugForm("sheetProtect--------------------------")
  For Each tempSheet In Sheets
    sheetName = tempSheet.Name
    If Not (sheetName Like "[設定,Notice,DataType]*") Then
      Call Library.showDebugForm("  " & sheetName)
      
      DoEvents
      ThisWorkbook.Worksheets(sheetName).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, passWord:=thisAppPasswd
      ThisWorkbook.Worksheets(sheetName).EnableSelection = xlNoRestrictions
    End If
  Next
  Call Library.showDebugForm("--------------------------------------")
End Function

'==================================================================================================
Function シート保護解除()
  Dim sheetName As String
  Dim tempSheet As Object

  Call init.Setting
  Call Library.showDebugForm("sheetUnprotect--------------------------")
  For Each tempSheet In Sheets
    sheetName = tempSheet.Name
    If Not (sheetName Like "[設定,Notice,DataType]*") Then
      Call Library.showDebugForm("  " & sheetName)
      
      DoEvents
      ThisWorkbook.Worksheets(sheetName).Unprotect passWord:=thisAppPasswd
    End If
  Next
  Call Library.showDebugForm("----------------------------------------")
End Function





