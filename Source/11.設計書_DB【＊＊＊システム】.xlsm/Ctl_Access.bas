Attribute VB_Name = "Ctl_Access"
Option Explicit

Dim dbCon       As ADODB.Connection
Dim DBRecordset As ADODB.Recordset
Dim QueryString As String


'**************************************************************************************************
' * MS Access
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function dbOpen()
  On Error GoTo catchError
  
  If isDBOpen = True Then
    Call Library.showDebugForm("Database is already opened")
    Exit Function
  End If
  Call Library.showDebugForm("ConnectServer：" & ConnectServer)
  
  Set dbCon = New ADODB.Connection
  dbCon.Open ConnectServer
  
  isDBOpen = True
  Call Library.showDebugForm("isDBOpen：" & isDBOpen)
  
  Exit Function
  
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  isDBOpen = False
  Call Library.showNotice(500, Err.Description, True)
End Function

'==================================================================================================
Function dbClose()
  On Error GoTo catchError
  
  If dbCon Is Nothing Then
    Call Library.showDebugForm("Database is already closed")
  Else
    dbCon.Close
    isDBOpen = False
    
    Set dbCon = Nothing
    Call Library.showDebugForm("isDBOpen：" & isDBOpen)
  End If
  
  Exit Function

'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showDebugForm("isDBOpen：" & isDBOpen)
  Call Library.showNotice(501, Err.Description, True)
End Function


'==================================================================================================
'テーブル情報取得
Function getTableInfo()
  Dim line As Long, endLine As Long
  Dim cat As ADOX.Catalog
  Dim tbl As ADOX.Table
  Dim tableCnt As Long

  '処理開始--------------------------------------
  'On Error GoTo catchError
  '初期値設定----
  FuncName = "Ctl_Access.getTableInfo"
  PrgP_Max = 4
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
  End If
  Call Library.showDebugForm(FuncName & "==========================================")
  'Call Ctl_Access.dbOpen
  '----------------------------------------------
  Set cat = New ADOX.Catalog
  cat.ActiveConnection = ConnectServer
  For Each tbl In cat.Tables
    Select Case tbl.Type
      Case "TABLE"
        Call Ctl_Common.addSheet(tbl.Name)
        Range("B2") = "マスターテーブル"
        
      Case "VIEW"
        Call Ctl_Common.addSheet(tbl.Name)
        Range("B2") = "クエリビュー"
      
      Case "LINK", "PASS-THROUGH"
        Call Ctl_Common.addSheet(tbl.Name)
        Range("B2") = "リンクテーブル"
        
      'システムテーブル
      Case "ACCESS TABLE", "SYSTEM TABLE"
        GoTo Lbl_nextfor
    End Select
    
    Range("D5") = ""
    Range("F5") = tbl.Name
    Call Ctl_Access.getColumnInfo

Lbl_nextfor:
  tableCnt = tableCnt + 1
  Call Ctl_ProgressBar.showBar(thisAppName, 1, PrgP_Max, tableCnt, cat.Tables.count, "テーブル情報取得：" & tbl.Name)
  Next tbl
  Call Ctl_Common.makeTblList
  
  '処理終了--------------------------------------
'  Call Ctl_Access.dbClose
  Application.Goto Reference:=Range("A1"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.usetting
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function


'==================================================================================================
'カラム情報取得
Function getColumnInfo()
  Dim line As Long, endLine As Long
  Dim tableName As String
  Dim columnCnt As Long
  Dim ClmRecordset As ADODB.Recordset

  Dim ColumnNames() As Variant
  Dim indexCount As Integer
  
  Dim Fields As ADODB.Field
  
  '処理開始----------------------------------------------------------------------------------------
  'On Error GoTo catchError
  
  '初期値設定----------------
  FuncName = "Ctl_Access.getColumnInfo"
  If PrgP_Max = 0 Then
    PrgP_Max = 2
  End If
  '--------------------------
  If runFlg = False Then
    Call Library.startScript
    Call init.Setting
    Call Ctl_ProgressBar.showStart
  End If
  
  Call Library.showDebugForm(FuncName & "==========================================")
  Call Ctl_Access.dbOpen
  '----------------------------------------------
  
  Call Ctl_Common.ClearData
  
  tableName = Range("H5")
  'カラム情報--------------------------------------------------------------------------------------
  QueryString = "SELECT * FROM " & tableName
  
  Set ClmRecordset = dbCon.Execute(QueryString)
  
  line = startLine
  columnCnt = 1
  For Each Fields In ClmRecordset.Fields
    Range("F" & line) = Fields.Name
    Range("E" & line) = ArryTypeName(Fields.Type)
    Range("F" & line) = Fields.DefinedSize
    
    Call Ctl_ProgressBar.showBar(thisAppName, 1, PrgP_Max, columnCnt, ClmRecordset.Fields.count, "カラム情報取得：" & Fields.Name)

    line = line + 1
    columnCnt = columnCnt + 1
    Call Ctl_Common.addRow(line)

  Next
  '処理終了----------------------------------------------------------------------------------------
  Call Ctl_Access.dbClose
  Application.Goto Reference:=Range("B8"), Scroll:=True
  Call Library.showDebugForm("=================================================================")
  If runFlg = False Then
    Call Ctl_ProgressBar.showEnd
    Call Library.endScript
    Call init.usetting
  End If
  '----------------------------------------------

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "：" & Err.Description, True)
End Function
