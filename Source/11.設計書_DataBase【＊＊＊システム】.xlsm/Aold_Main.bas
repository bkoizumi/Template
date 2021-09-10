Attribute VB_Name = "Aold_Main"
'' *********************************************************************
'' * 新規シート追加(リボンメニューからの呼び出し)
'' *
'' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'' *********************************************************************
'Sub addSheet(ByVal control As IRibbonControl)
'  Library_StartScript
'  DataBase_AddSheet
'  Library_EndScript
'
'  ThisWorkbook.Activate
'  VBA.AppActivate Excel.Application.Caption
'End Sub
'
'' *********************************************************************
'' * SQL作成(リボンメニューからの呼び出し)
'' *
'' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'' *********************************************************************
'Sub MakeSQL(ByVal control As IRibbonControl)
'  Library_StartScript
'  DataBase_MakeSQL (1)
'  Library_EndScript
'  ThisWorkbook.Activate
'  VBA.AppActivate Excel.Application.Caption
'End Sub
'
'' *********************************************************************
'' * 全シート分のSQL作成(リボンメニューからの呼び出し)
'' *
'' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'' *********************************************************************
'Sub MakeAllSheetSQL(ByVal control As IRibbonControl)
'  Library_StartScript
'  DataBase_MakeAllSheetSQL
'  Library_EndScript
'  ThisWorkbook.Activate
'  VBA.AppActivate Excel.Application.Caption
'End Sub
'
'' *********************************************************************
'' * Flamework用SQL作成(リボンメニューからの呼び出し)
'' *
'' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'' *********************************************************************
'Sub MakeFlameworkSQL(ByVal control As IRibbonControl)
'  Library_StartScript
'  DataBase_MakeFlameworkSQL (1)
'  Library_EndScript
'  ThisWorkbook.Activate
'  VBA.AppActivate Excel.Application.Caption
'End Sub
'
'' *********************************************************************
'' * 全シート分のSFlamework用QL作成(リボンメニューからの呼び出し)
'' *
'' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'' *********************************************************************
'Sub MakeAllFlameworkSQL(ByVal control As IRibbonControl)
'  Library_StartScript
'  DataBase_MakeAllSheetFlameworkSQL
'  Library_EndScript
'  ThisWorkbook.Activate
'  VBA.AppActivate Excel.Application.Caption
'End Sub
'
'' *********************************************************************
'' * 印刷範囲設定(リボンメニューからの呼び出し)
'' *
'' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'' *********************************************************************
'Sub SetPrintArea(ByVal control As IRibbonControl)
'  Library_StartScript
'  DataBase_SetPrintArea
'  Library_EndScript
'  ThisWorkbook.Activate
'  VBA.AppActivate Excel.Application.Caption
'End Sub
'
'' *********************************************************************
'' * サーバーからテーブル情報取得(リボンメニューからの呼び出し)
'' *
'' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'' *********************************************************************
'Sub GetTableList(ByVal control As IRibbonControl)
'  Library_StartScript
'  DataBase_GetTableList
'  Library_EndScript
'  ThisWorkbook.Activate
'  VBA.AppActivate Excel.Application.Caption
'End Sub
'
'' *********************************************************************
'' * サーバーからカラム情報取得(リボンメニューからの呼び出し)
'' *
'' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'' *********************************************************************
'Sub GetColumnList(ByVal control As IRibbonControl)
'  Library_StartScript
'  DataBase_GetColumn (True)
'  Library_EndScript
'  ThisWorkbook.Activate
'  VBA.AppActivate Excel.Application.Caption
'End Sub
'
'' *********************************************************************
'' * テーブル一覧生成(リボンメニューからの呼び出し)
'' *
'' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'' *********************************************************************
'Sub MakeTableList(ByVal control As IRibbonControl)
'
'  Library_StartScript
'  DataBase_MakeTableList
'  Library_EndScript
'
'  ThisWorkbook.Activate
'  VBA.AppActivate Excel.Application.Caption
'End Sub
'
'' *********************************************************************
'' * セル設定(リボンメニューからの呼び出し)
'' *
'' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'' *********************************************************************
'Sub Cell_StyleSetting(ByVal control As IRibbonControl)
'
'  Library_StartScript
'  DataBase_SetCellsStyle
'  Library_EndScript
'
'  ThisWorkbook.Activate
'  VBA.AppActivate Excel.Application.Caption
'End Sub
'
'
'' *********************************************************************
'' * シート設定(リボンメニューからの呼び出し)
'' *
'' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'' *********************************************************************
'Sub Sheet_StyleSetting(ByVal control As IRibbonControl)
'
'  DataBase_SetSheetStyle
'
'  ThisWorkbook.Activate
'  VBA.AppActivate Excel.Application.Caption
'End Sub
'
'' *********************************************************************
'' * 再設定ボタン
'' *
'' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'' *********************************************************************
'Sub ResetButton()
'  Library_StartScript
'  DataBase_Reset
'  Library_EndScript
'
'  ThisWorkbook.Activate
'  VBA.AppActivate Excel.Application.Caption
'End Sub
'
'' *********************************************************************
'' * シート取得ボタン
'' *
'' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'' *********************************************************************
'Sub GetSheetListButton()
'
'  Library_StartScript
'  Call Library_GetSheetList("I")
'  Library_EndScript
'
'  ThisWorkbook.Activate
'  VBA.AppActivate Excel.Application.Caption
'End Sub
'
'' *********************************************************************
'' * シート「テーブルリスト」をアクティブにする
'' *
'' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'' *********************************************************************
'Sub DisplayTableList()
'    Sheets("TBLリスト").Select
'End Sub
'
'' *********************************************************************
'' * カラム情報再取得
'' *
'' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'' *********************************************************************
'Sub GetAllSheetColumnList()
'
'  Dim sheetName As String
'  Library_StartScript
'
'  For Each objSheet In ActiveWorkbook.Sheets
'
'    sheetName = objSheet.Name
'
'    If Library_CheckExcludeSheet(sheetName, 9) Then
'      Worksheets(sheetName).Select
'      Range("C9").Select
'      DataBase_GetColumn (True)
'    End If
'  Next
'
'  Library_EndScript
'  ThisWorkbook.Activate
'  VBA.AppActivate Excel.Application.Caption
'End Sub
'
