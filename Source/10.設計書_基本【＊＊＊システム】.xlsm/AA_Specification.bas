Attribute VB_Name = "AA_Specification"
Public W_Page1NoCol As Integer
Public W_Page2NoCol As Integer
Public OnePageRow As Integer
Public Page1Area As String
Public Page2Area As String
Public Page1StartArea As Integer
Public Page1CenterArea As String
Public Page2CenterArea As String


Public InputName As String
Public InputType As String
Public InputDataType As String
Public InputNameTag As String
Public InputLimit_tmp As String
Public InputRequired As String
Public InputTestString As String
Public InputLimit As Variant
Public URL As String
Public InputLimitMin As Long
Public InputLimitMax As Long
Public Title As String
Public InputNo As Long


'***********************************************************************************************************************************************
' * 設計書用環境設定
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Sub init()

  ' 1行目のページ数書き込み位置
  W_Page1NoCol = Worksheets("設定").Range("B3")
  
  ' 2行目のページ数書き込み位置
  W_Page2NoCol = Worksheets("設定").Range("B4")

  ' 1ページの行数
  OnePageRow = Worksheets("設定").Range("B5")

  ' 1ページ目の目次開始位置
  Page1StartArea = Worksheets("設定").Range("B6")
  
  ' 1ページ目の目次表示位置
  Page1Area = Worksheets("設定").Range("B7")

  ' 1ページ目の目次分割位置
  Page1CenterArea = Worksheets("設定").Range("B8")
  
  ' 2ページ目の目次表示位置
  Page2Area = Worksheets("設定").Range("B9")

  ' 2ページ目の目次分割位置
  Page2CenterArea = Worksheets("設定").Range("B10")

End Sub

'***********************************************************************************************************************************************
' * 設計書用目次作成
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Sub MakeMenu()

  Dim PageLine As Long
  Dim TitleName As String
  Dim FunctionName As String
  Dim PageCnt As Long
  Dim TitleCnt As Long
  Dim EndBookRowLine As Long
  Dim RowCnt As Long
  Dim W_PageNoCol As Long
  
  Dim ThisActiveSheetName As String
  ThisActiveSheetName = ActiveSheet.Name
  
  Call Specification.init
  
  ' 最終行取得
  EndBookRowLine = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row
  
  ' 現在設定されている目次を削除
  Call Specification.DeleteMenu(1)

  '---------------------------------------------------------------------------------------
  ' 目次生成メイン処理
  '---------------------------------------------------------------------------------------
  PageLine = Page1StartArea
  PageCnt = 1
  W_PageNoCol = W_Page1NoCol
  
  ' プログレスバーの表示開始
  ProgressBar_ProgShowStart
 
  For RowCnt = 44 To EndBookRowLine Step OnePageRow
  
    ' タイトル取得
    TitleName = Cells(RowCnt + 1, 4)
    
    ' 機能取得
    FunctionName = Cells(RowCnt + 1, 19)
    
    ' ページ番号書き込み
    With Cells(PageLine, W_PageNoCol)
      .Value = PageCnt
      .Font.Name = "Meiryo UI"
      .Font.Size = 9
      .HorizontalAlignment = xlGeneral
      .VerticalAlignment = xlCenter
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = True
      .ReadingOrder = xlContext
      .MergeCells = False
      .ShrinkToFit = True
'      .NumberFormatLocal = "@"
    End With
        
    If FunctionName <> "" Then
      TitleName = TitleName & " - " & FunctionName
    End If
    
    ' プログレスバーのカウント変更（現在のカウント、全カウント数、メッセージ）
    ProgressBar_ProgShowCount "目次生成中", RowCnt, EndBookRowLine, "P." & PageCnt & " " & TitleName
    
    ' タイトル(リンク付)書き込み
    With Cells(PageLine, W_PageNoCol + 1)
      .Value = TitleName
      .Select
      .Hyperlinks.add Anchor:=Selection, Address:="", SubAddress:="#" & "A" & RowCnt
      .Font.ColorIndex = 1
      .Font.Underline = xlUnderlineStyleNone
      .Font.Name = "Meiryo UI"
      .Font.Size = 9
      .HorizontalAlignment = xlGeneral
      .VerticalAlignment = xlCenter
      .ShrinkToFit = True
    End With
    
    ' セルの結合
    Range("E" & PageLine & ":V" & PageLine).Select
    Selection.Merge
    Range("AA" & PageLine & ":AR" & PageLine).Select
    Selection.Merge
    
    ' 各ページにページ番号書き込み
    Range("AW" & RowCnt & ":AX" & RowCnt + 1).Select
    Selection.Merge
    
    Range("AW" & RowCnt).Value = "P." & PageCnt
    
    ' 目次へのリンク追加
    If RowCnt > 2 Then
      Range("AW" & RowCnt - 1 & ":AX" & RowCnt - 1).Select
      Selection.Merge
      Range("AW" & RowCnt - 1).Value = "=HYPERLINK(""#$A$1"",""目次へ"")"
    End If
    
    PageLine = PageLine + 1
    PageCnt = PageCnt + 1
  
    ' ======================= 制御 ======================
    ' 1ページ目の2列目
    If PageCnt = OnePageRow - 4 Then
      W_PageNoCol = W_Page2NoCol
      PageLine = OnePageRow + 5
      Call Specification.AddLine(1)
      
    ' 2ページ目の1列名
    ElseIf PageCnt = (OnePageRow - 5) * 2 + 1 Then
    
      If Range("D88") <> "目次" And Range("D88") <> "もくじ" Then
        If MsgBox("目次が2ページ目に挿入されます" & vbLf & " 2ページの準備OK？", vbYesNo, "2ページの準備OK？") = vbNo Then
          Call Library.EndScript
          MsgBox "2ページ目のタイトルを目次に設定してください" & vbLf & "処理を中断します。"
          
          ' プログレスバーの表示終了処理
          ProgressBar_ProgShowClose
          
          Exit Sub
        End If
      Else
        ' ======================= 2ページ目次生成======================
        W_PageNoCol = W_Page1NoCol
        PageLine = OnePageRow * 2 + 5
        Call Specification.DeleteMenu(2)
        Call Specification.AddLine(2)

      End If
    
    ' 2ページ目の2列名
    ElseIf PageCnt = (OnePageRow - 5) * 3 + 1 Then
      W_PageNoCol = W_Page2NoCol
      PageLine = OnePageRow * 2 + 5
    End If
  Next

  ' プログレスバーの表示終了処理
  ProgressBar_ProgShowClose

  ' 印刷領域設定
  'Call Specification.SetPrintArea


End Sub
'***********************************************************************************************************************************************
' * 設計書用目次削除
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function DeleteMenu(Page As Integer)

  If Page = 1 Then
    Range(Page1Area).Select
  ElseIf Page = 2 Then
    Range(Page2Area).Select
  End If
  
  Selection.Clear
  Application.CutCopyMode = False
End Function


'***********************************************************************************************************************************************
' * 設計書用罫線設定
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function AddLine(Page As Integer)
  
  If Page = 1 Then
    Range(Page1CenterArea).Select
  ElseIf Page = 2 Then
    Range(Page2CenterArea).Select
  End If
  
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Function

'***********************************************************************************************************************************************
' * 設計書用ページ追加
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function addPage()

  On Error GoTo ErrHand
  
  Dim PageLine As Integer
  Dim TitleName As String
  Dim FunctionName As String
  Dim PageCnt As Integer
  Dim TitleCnt As Integer
  Dim EndBookRowLine As Integer
  Dim RowCnt As Integer
  Dim W_PageNoCol As Integer
  Dim ThisActiveSheetName As String
  
  Call Specification.init
  
  ThisActiveSheetName = ActiveSheet.Name
  EndBookRowLine = Sheets(ThisActiveSheetName).Cells(Rows.count, 1).End(xlUp).Row + OnePageRow - 1
  

  Sheets("Sheet1").Select
  Range("A1:AW43").Select
  Selection.Copy

  Sheets(ThisActiveSheetName).Select
  Range("A" & EndBookRowLine).Select
  ActiveSheet.Paste

  Application.CutCopyMode = False

  ' 前ページのタイトル設定
  ActiveSheet.Range("D" & EndBookRowLine + 1).Value = Range("D" & EndBookRowLine - OnePageRow + 1).Value
  
  ActiveSheet.Range("A" & EndBookRowLine).Select
  With ActiveWindow
    .ScrollRow = EndBookRowLine
    .ScrollColumn = 1
  End With
Exit Function

ErrHand:
  Call Library.EndScript
  Resume Next
End Function


'***********************************************************************************************************************************************
' * 設計書用印刷範囲設定
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Sub SetPrintArea()

  On Error GoTo ErrHand
  
  Dim EndBookRowLine As Long
  Dim PageCnt As Long
  Dim W_PageNoCol As Long
  Dim RowCnt As Long
  Dim ThisActiveSheetName As String
  
  Call Specification.init
  
  ThisActiveSheetName = ActiveSheet.Name
  
  EndBookRowLine = ActiveSheet.Cells(Rows.count, 49).End(xlUp).Row
  W_PageNoCol = OnePageRow
  PageCnt = 1
  
  ActiveSheet.PageSetup.PrintArea = "A1:AU" & EndBookRowLine
  
  '改ページプレビュー
  ActiveWindow.View = xlPageBreakPreview
  
  ' プログレスバーの表示開始
  ProgressBar_ProgShowStart
  
  For RowCnt = 1 To EndBookRowLine Step OnePageRow

    ' プログレスバーのカウント変更（現在のカウント、全カウント数、メッセージ）
    ProgressBar_ProgShowCount "印刷範囲設定", RowCnt, EndBookRowLine, "P." & PageCnt
    
    Set Sheets(ThisActiveSheetName).HPageBreaks(PageCnt).Location = Range("A" & W_PageNoCol + 1)
    W_PageNoCol = W_PageNoCol + OnePageRow
    PageCnt = PageCnt + 1
  Next RowCnt
  
  ActiveWindow.View = xlNormalView

  ' プログレスバーの表示終了処理
  ProgressBar_ProgShowClose

Exit Sub

ErrHand:
  ActiveWindow.View = xlNormalView
  
  ' プログレスバーの表示終了処理
  ProgressBar_ProgShowClose

  ' 画面描写制御終了
  Call Library.EndScript
End Sub


'***********************************************************************************************************************************************
' * 設計書用Selenium設計欄チェック
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function Selenium_Check()

  Dim EndBookRowLine As Long
  Dim RowCnt As Long
  Dim TitleName As String
  
  
  Call Specification.init
  
  ' 最終行取得
  EndBookRowLine = ActiveSheet.Cells(Rows.count, 49).End(xlUp).Row
  For RowCnt = 1 To EndBookRowLine Step OnePageRow
    ' 機能名取得
    If Cells(RowCnt + 1, 19) = "入力項目制限" Then
      Title = Cells(RowCnt + 1, 4)
      Call Specification.Selenium_Get(RowCnt + 1)
    End If
  Next

End Function


'***********************************************************************************************************************************************
' * 設計書用Selenium欄取得
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function Selenium_Get(RowCnt As Long)

  Dim line As Long

  For line = RowCnt To RowCnt + OnePageRow Step 1
    If Range("B" & line) = "No." Then
      URL = Range("B" & line - 1)
      Call Specification.Selenium_Make(line + 1)
    End If
  Next
End Function


'***********************************************************************************************************************************************
' * 設計書用Seleniumテストケース生成
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function Selenium_Make(RowCnt As Long)

  Dim line As Long

  For line = RowCnt To RowCnt + OnePageRow Step 1
    If Range("J" & line) = "テキストエリア" Or Range("J" & line) = "テキストボックス" Then
      InputName = Range("C" & line)
      InputType = Range("J" & line)
      InputDataType = Range("O" & line)
      InputNameTag = Range("S" & line)
      InputLimit_tmp = Range("W" & line)
      InputRequired = Range("Z" & line)
      InputNo = Range("B" & line)

      '入力桁数の最小/最大を取得
      If InStr(InputLimit_tmp, "～") <> 0 Then
        InputLimit = Split(InputLimit_tmp, "～")
        InputLimitMin = CLng(InputLimit(0))
        InputLimitMax = CLng(InputLimit(1))
      ElseIf InputLimit_tmp <> "" Then
        InputLimitMin = InputLimit_tmp
        InputLimitMax = 0
      End If
      
    'テスト項目作成-----------------------------------------------------------------------------------------
      Call Specification.Selenium_Makehtml("半角数字-最小桁数以下")
      Call Specification.Selenium_Makehtml("半角数字-最小桁")
      Call Specification.Selenium_Makehtml("半角数字")
      Call Specification.Selenium_Makehtml("半角数字-最大桁")
      Call Specification.Selenium_Makehtml("半角数字-最大桁数以上")
      
      Call Specification.Selenium_Makehtml("半角英小文字-最小桁数以下")
      Call Specification.Selenium_Makehtml("半角英小文字-最小桁")
      Call Specification.Selenium_Makehtml("半角英小文字")
      Call Specification.Selenium_Makehtml("半角英小文字-最大桁")
      Call Specification.Selenium_Makehtml("半角英小文字-最大桁数以上")
      
      Call Specification.Selenium_Makehtml("半角英大文字-最小桁数以下")
      Call Specification.Selenium_Makehtml("半角英大文字-最小桁")
      Call Specification.Selenium_Makehtml("半角英大文字")
      Call Specification.Selenium_Makehtml("半角英大文字-最大桁")
      Call Specification.Selenium_Makehtml("半角英大文字-最大桁数以上")
      
      Call Specification.Selenium_Makehtml("半角英文字-最小桁数以下")
      Call Specification.Selenium_Makehtml("半角英文字-最小桁")
      Call Specification.Selenium_Makehtml("半角英文字")
      Call Specification.Selenium_Makehtml("半角英文字-最大桁")
      Call Specification.Selenium_Makehtml("半角英文字-最大桁数以上")
      
      Call Specification.Selenium_Makehtml("半角英数字-最小桁数以下")
      Call Specification.Selenium_Makehtml("半角英数字-最小桁")
      Call Specification.Selenium_Makehtml("半角英数字")
      Call Specification.Selenium_Makehtml("半角英数字-最大桁")
      Call Specification.Selenium_Makehtml("半角英数字-最大桁数以上")
      
      Call Specification.Selenium_Makehtml("半角記号-最小桁数以下")
      Call Specification.Selenium_Makehtml("半角記号-最小桁")
      Call Specification.Selenium_Makehtml("半角記号")
      Call Specification.Selenium_Makehtml("半角記号-最大桁")
      Call Specification.Selenium_Makehtml("半角記号-最大桁数以上")
      
      Call Specification.Selenium_Makehtml("半角文字-最小桁数以下")
      Call Specification.Selenium_Makehtml("半角文字-最小桁")
      Call Specification.Selenium_Makehtml("半角文字")
      Call Specification.Selenium_Makehtml("半角文字-最大桁")
      Call Specification.Selenium_Makehtml("半角文字-最大桁数以上")
 
      Call Specification.Selenium_Makehtml("全角数字-最小桁数以下")
      Call Specification.Selenium_Makehtml("全角数字-最小桁")
      Call Specification.Selenium_Makehtml("全角数字")
      Call Specification.Selenium_Makehtml("全角数字-最大桁")
      Call Specification.Selenium_Makehtml("全角数字-最大桁数以上")
      
      Call Specification.Selenium_Makehtml("全角英小文字-最小桁数以下")
      Call Specification.Selenium_Makehtml("全角英小文字-最小桁")
      Call Specification.Selenium_Makehtml("全角英小文字")
      Call Specification.Selenium_Makehtml("全角英小文字-最大桁")
      Call Specification.Selenium_Makehtml("全角英小文字-最大桁数以上")
      
      Call Specification.Selenium_Makehtml("全角英大文字-最小桁数以下")
      Call Specification.Selenium_Makehtml("全角英大文字-最小桁")
      Call Specification.Selenium_Makehtml("全角英大文字")
      Call Specification.Selenium_Makehtml("全角英大文字-最大桁")
      Call Specification.Selenium_Makehtml("全角英大文字-最大桁数以上")
      
      Call Specification.Selenium_Makehtml("全角英文字-最小桁数以下")
      Call Specification.Selenium_Makehtml("全角英文字-最小桁")
      Call Specification.Selenium_Makehtml("全角英文字")
      Call Specification.Selenium_Makehtml("全角英文字-最大桁")
      Call Specification.Selenium_Makehtml("全角英文字-最大桁数以上")
      
      Call Specification.Selenium_Makehtml("全角英数字-最小桁数以下")
      Call Specification.Selenium_Makehtml("全角英数字-最小桁")
      Call Specification.Selenium_Makehtml("全角英数字")
      Call Specification.Selenium_Makehtml("全角英数字-最大桁")
      Call Specification.Selenium_Makehtml("全角英数字-最大桁数以上")
      
      Call Specification.Selenium_Makehtml("全角記号-最小桁数以下")
      Call Specification.Selenium_Makehtml("全角記号-最小桁")
      Call Specification.Selenium_Makehtml("全角記号")
      Call Specification.Selenium_Makehtml("全角記号-最大桁")
      Call Specification.Selenium_Makehtml("全角記号-最大桁数以上")
      
      Call Specification.Selenium_Makehtml("全角ひらがな-最小桁数以下")
      Call Specification.Selenium_Makehtml("全角ひらがな-最小桁")
      Call Specification.Selenium_Makehtml("全角ひらがな")
      Call Specification.Selenium_Makehtml("全角ひらがな-最大桁")
      Call Specification.Selenium_Makehtml("全角ひらがな-最大桁数以上")
      
      Call Specification.Selenium_Makehtml("全角カタカナ-最小桁数以下")
      Call Specification.Selenium_Makehtml("全角カタカナ-最小桁")
      Call Specification.Selenium_Makehtml("全角カタカナ")
      Call Specification.Selenium_Makehtml("全角カタカナ-最大桁")
      Call Specification.Selenium_Makehtml("全角カタカナ-最大桁数以上")
      
      Call Specification.Selenium_Makehtml("常用漢字-最小桁数以下")
      Call Specification.Selenium_Makehtml("常用漢字-最小桁")
      Call Specification.Selenium_Makehtml("常用漢字")
      Call Specification.Selenium_Makehtml("常用漢字-最大桁")
      Call Specification.Selenium_Makehtml("常用漢字-最大桁数以上")
      
      Call Specification.Selenium_Makehtml("半角カタカナ-最小桁数以下")
      Call Specification.Selenium_Makehtml("半角カタカナ-最小桁")
      Call Specification.Selenium_Makehtml("半角カタカナ")
      Call Specification.Selenium_Makehtml("半角カタカナ-最大桁")
      Call Specification.Selenium_Makehtml("半角カタカナ-最大桁数以上")

      Call Specification.Selenium_Makehtml("全角文字-最小桁数以下")
      Call Specification.Selenium_Makehtml("全角文字-最小桁")
      Call Specification.Selenium_Makehtml("全角文字")
      Call Specification.Selenium_Makehtml("全角文字-最大桁")
      Call Specification.Selenium_Makehtml("全角文字-最大桁数以上")
      
      Call Specification.Selenium_Makehtml("機種依存文字")
        
      If InputDataType = "日付" Then
        Call Specification.Selenium_Makehtml("日付正常01")
        Call Specification.Selenium_Makehtml("日付異常01")
        Call Specification.Selenium_Makehtml("日付月異常01")
        Call Specification.Selenium_Makehtml("日付日異常01")
        
      
      ElseIf InputDataType = "email" Then
        Call Specification.Selenium_Makehtml("メールアドレス正常01")
        Call Specification.Selenium_Makehtml("メールアドレス正常02")
        Call Specification.Selenium_Makehtml("メールアドレス正常03")
        Call Specification.Selenium_Makehtml("メールアドレス正常04")
        Call Specification.Selenium_Makehtml("メールアドレスローカル部異常")
        Call Specification.Selenium_Makehtml("メールアドレス異常")
      End If
      
    ElseIf Range("J" & line) = "登録/検索ボタン" Then
      Call Specification.Selenium_MakehtmlFooter(line)
      Call Specification.Selenium_MakeIndex
    End If
  Next
End Function


'***********************************************************************************************************************************************
' * 設計書用Selenium
' *
' * Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function Selenium_Makehtml(MakeType As String)

  Dim htmlTag As String
  Dim L_InputLimit As Long



  htmlTag = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbLf & _
              "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">" & vbLf & _
              "<html xmlns=""http://www.w3.org/1999/xhtml"" xml:lang=""ja"" lang=""ja"">" & vbLf & _
              "<head profile=""http://selenium-ide.openqa.org/profiles/test-case"">" & vbLf & _
              "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"" />" & vbLf & _
              "<link rel=""selenium.base"" href=""" & Range("BaseURL") & """ />" & vbLf & _
              "<title>" & InputName & "</title>" & vbLf & _
              "</head>" & vbLf & _
              "<body>" & vbLf & _
              "<table cellpadding='1' cellspacing='1' border='1'>" & vbLf & _
              "<thead>" & vbLf & _
              "<tr><td rowspan='1' colspan='3'></td></tr>" & vbLf & _
              "</thead><tbody>" & vbLf
  
    htmlTag = htmlTag & "<!--■" & InputName & " " & MakeType & "-->" & vbLf
    htmlTag = htmlTag & "<tr>" & vbLf
    htmlTag = htmlTag & "  <td>open</td>" & vbLf
    htmlTag = htmlTag & "  <td>" & URL & "</td>" & vbLf
    htmlTag = htmlTag & "  <td></td>" & vbLf
    htmlTag = htmlTag & "</tr>" & vbLf

    Select Case MakeType
'=====================================================================================================================================
      Case "半角数字-最小桁数以下", "半角数字-最小桁", "半角数字", "半角数字-最大桁", "半角数字-最大桁数以上"
        InputTestString = HalfWidthDigit
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "半角英小文字-最小桁数以下", "半角英小文字-最小桁", "半角英小文字", "半角英小文字-最大桁", "半角英小文字-最大桁数以上"
        InputTestString = StrConv(HalfWidthCharacters, vbLowerCase)
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "半角英大文字-最小桁数以下", "半角英大文字-最小桁", "半角英大文字", "半角英大文字-最大桁", "半角英大文字-最大桁数以上"
        InputTestString = HalfWidthCharacters
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "半角英数字-最小桁数以下", "半角英数字-最小桁", "半角英数字", "半角英数字-最大桁", "半角英数字-最大桁数以上"
        InputTestString = HalfWidthCharacters & StrConv(HalfWidthCharacters, vbLowerCase) & HalfWidthDigit
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "半角記号-最小桁数以下", "半角記号-最小桁数", "半角記号", "半角記号-最大桁", "半角記号-最大桁数以上"
        InputTestString = SymbolCharacters
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "半角文字-最小桁数以下", "半角文字-最小桁数", "半角文字", "半角文字-最大桁", "半角文字-最大桁数以上"
        InputTestString = StrConv(HalfWidthCharacters, vbLowerCase) & _
                          HalfWidthCharacters & _
                          HalfWidthDigit & _
                          SymbolCharacters
'=====================================================================================================================================
      Case "全角数字-最小桁数以下", "全角数字-最小桁", "全角数字", "全角数字-最大桁", "全角数字-最大桁数以上"
        InputTestString = StrConv(HalfWidthDigit, vbWide)
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "全角英小文字-最小桁数以下", "全角英小文字-最小桁", "全角英小文字", "全角英小文字-最大桁", "全角英小文字-最大桁数以上"
        InputTestString = StrConv(StrConv(HalfWidthCharacters, vbLowerCase), vbWide)
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "全角英大文字-最小桁数以下", "全角英大文字-最小桁", "全角英大文字", "全角英大文字-最大桁", "全角英大文字-最大桁数以上"
        InputTestString = StrConv(HalfWidthCharacters, vbWide)
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "全角英文字-最小桁数以下", "全角英文字-最小桁", "全角英文字", "全角英文字-最大桁", "全角英文字-最大桁数以上"
        InputTestString = StrConv(StrConv(HalfWidthCharacters, vbLowerCase), vbWide) & _
                          StrConv(HalfWidthCharacters, vbWide)
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "全角英数字-最小桁数以下", "全角英数字-最小桁", "全角英数字", "全角英数字-最大桁", "全角英数字-最大桁数以上"
        InputTestString = StrConv(HalfWidthDigit, vbWide) & _
                          StrConv(StrConv(HalfWidthCharacters, vbLowerCase), vbWide) & _
                          StrConv(HalfWidthCharacters, vbWide)
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "全角記号-最小桁数以下", "全角記号-最小桁数", "全角記号", "全角記号-最大桁", "全角記号-最大桁数以上"
        InputTestString = StrConv(SymbolCharacters, vbWide)

'=====================================================================================================================================
      Case "全角ひらがな-最小桁数以下", "全角ひらがな-最小桁数", "全角ひらがな", "全角ひらがな-最大桁", "全角ひらがな-最大桁数以上"
        InputTestString = JapaneseCharacters
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "全角カタカナ-最小桁数以下", "全角カタカナ-最小桁数", "全角カタカナ", "全角カタカナ-最大桁", "全角カタカナ-最大桁数以上"
        InputTestString = StrConv(JapaneseCharacters, vbKatakana)
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "半角カタカナ-最小桁数以下", "半角カタカナ-最小桁数", "半角カタカナ", "半角カタカナ-最大桁", "半角カタカナ-最大桁数以上"
        InputTestString = StrConv(StrConv(JapaneseCharacters, vbKatakana), vbNarrow)
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "常用漢字-最小桁数以下", "常用漢字-最小桁数", "常用漢字", "常用漢字-最大桁", "常用漢字-最大桁数以上"
        InputTestString = JapaneseCharactersCommonUse
'-------------------------------------------------------------------------------------------------------------------------------------
      Case "全角文字-最小桁数以下", "全角文字-最小桁数", "全角文字", "全角文字-最大桁", "全角文字-最大桁数以上"
        InputTestString = StrConv(HalfWidthDigit, vbWide) & _
                          StrConv(StrConv(HalfWidthCharacters, vbLowerCase), vbWide) & _
                          StrConv(HalfWidthCharacters, vbWide) & _
                          StrConv(SymbolCharacters, vbWide) & _
                          JapaneseCharacters & _
                          StrConv(JapaneseCharacters, vbKatakana) & _
                          StrConv(StrConv(JapaneseCharacters, vbKatakana), vbNarrow) & _
                          JapaneseCharactersCommonUse
                          
'=====================================================================================================================================
      Case "日付正常01"
        InputTestString = "2016/01/01"
      
      Case "日付-異常01"
        InputTestString = "2016/0101/"
      
      Case "日付月異常01"
        InputTestString = "2016/15/01"
      
      Case "日付日異常01"
        InputTestString = "2016/01/55"

'=====================================================================================================================================
      Case "メールアドレス-正常01"
        InputTestString = "vb.project@vb-project.com"
        
      Case "メールアドレス-正常02"
        InputTestString = "user+mailbox/department=shipping@vb-project.com"
        
      Case "メールアドレス-正常03"
        InputTestString = """Joe.\\Blow""@vb-project.com"
        
      Case "メールアドレス-正常04"
        InputTestString = "1234567890123456789012345678901234567890123456789012345678901234@abcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvyzab.co.jp"
        
      Case "メールアドレス-ローカル部異常"
        InputTestString = "vb..project@vb-project.com"
        
      Case "メールアドレス異常"
        InputTestString = "1234567890123456789012345678901234567890123456789012345678901234@abcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvyzabcdefghijklmnopqrstuvyz.co.jp"

'=====================================================================================================================================
      Case Else
        InputTestString = StrConv(HalfWidthCharacters, vbLowerCase) & _
                          HalfWidthCharacters & _
                          HalfWidthDigit & _
                          SymbolCharacters & _
                          StrConv(HalfWidthDigit, vbWide) & _
                          StrConv(StrConv(HalfWidthCharacters, vbLowerCase), vbWide) & _
                          StrConv(HalfWidthCharacters, vbWide) & _
                          StrConv(SymbolCharacters, vbWide) & _
                          MachineDependentCharacters
    End Select
    
    '最大文字数の設定
    If InStr(MakeType, "-") <> 0 Then
      MakeType_tmp = Split(MakeType, "-")
      Select Case MakeType_tmp(1)
        Case "最小桁数以下"
          L_InputLimit = InputLimitMin - 1
        
        Case "最小桁数"
          L_InputLimit = InputLimitMin
          
        Case "最大桁数"
          If InputLimitMax = 0 Then
            L_InputLimit = 0
          Else
            L_InputLimit = InputLimitMax
          End If
        Case "最大桁数以上"
          If InputLimitMax = 0 Then
            L_InputLimit = 0
          Else
            L_InputLimit = InputLimitMax + 1
          End If
        Case Else
          L_InputLimit = InputLimitMax
      End Select
    Else
      If InputLimitMax = 0 Then
        L_InputLimit = 0
      Else
        L_InputLimit = InputLimitMin + 1
      End If
    End If
    
    If L_InputLimit = 0 Then
      Exit Function
    End If
    
    
    '入力文字をランダムに設定
    InputTestString = call Library.Randomize(InputTestString, L_InputLimit)
    
    htmlTag = htmlTag & "<!-- " & MakeType & "-->" & vbLf
    htmlTag = htmlTag & "<tr>" & vbLf
    htmlTag = htmlTag & "  <td>type</td>" & vbLf
    htmlTag = htmlTag & "  <td>" & InputNameTag & "</td>" & vbLf
    htmlTag = htmlTag & "  <td>" & InputTestString & "</td>" & vbLf
    htmlTag = htmlTag & "</tr>" & vbLf

    '=======================================================================================
    'ディレクトリ作成
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    SeleniumFolder = ActiveWorkbook.path & "\" & Title
    If objFSO.FolderExists(folderspec:=SeleniumFolder) = False Then
      objFSO.CreateFolder SeleniumFolder
    End If
    Set objFSO = Nothing


    ' ADODB処理
    Set ObjADODB_TestCase = CreateObject("ADODB.Stream")
  
    'オブジェクトに保存するデータの種類を文字列型に指定する
    ObjADODB_TestCase.Type = 2
    
    '文字列型のオブジェクトの文字コードを指定する
    ObjADODB_TestCase.Charset = "UTF-8"
    ObjADODB_TestCase.LineSeparator = 10
  
    'オブジェクトのインスタンスを作成
    ObjADODB_TestCase.Open
    

    ' テストケース保存
    ObjADODB_TestCase.WriteText htmlTag, 1
    ObjADODB_TestCase.SaveToFile (SeleniumFolder & "\" & InputNo & "_" & InputName & "_" & MakeType & ".html"), 2
   
    'オブジェクトを閉じる
    ObjADODB_TestCase.Close
    Set ObjADODB_TestCase = Nothing
End Function


Function Selenium_MakehtmlFooter(ByVal line As Long)

  Dim htmlTag As String
  '=======================================================================================
  'ディレクトリ作成
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  
  captureFolder = ActiveWorkbook.path & "\" & Title
  If objFSO.FolderExists(folderspec:=captureFolder) = False Then
    objFSO.CreateFolder captureFolder
  End If
  captureFolder = ActiveWorkbook.path & "\" & Title & "\capture"
  If objFSO.FolderExists(folderspec:=captureFolder) = False Then
    objFSO.CreateFolder captureFolder
  End If
  
  
  Set objFSO = Nothing
  
  Dim buf As String
  buf = Dir(ActiveWorkbook.path & "\" & Title & "\*.html")
  
  Do While Len(buf) > 0
    If LCase(buf) Like "*.html" Then
      
      htmlTag = "<tr>" & vbLf
      htmlTag = htmlTag & "  <td>captureEntirePageScreenshot</td>" & vbLf
      htmlTag = htmlTag & "  <td>" & ActiveWorkbook.path & "\" & Title & "\capture\" & buf & "01.png</td>" & vbLf
      htmlTag = htmlTag & "  <td>background=#FFFFFF</td>" & vbLf
      htmlTag = htmlTag & "</tr>" & vbLf
      htmlTag = htmlTag & "<tr>" & vbLf
      
      If (Range("AF" & line) <> "") Then
        htmlTag = htmlTag & "  <td>runScriptAndWait</td>" & vbLf
        htmlTag = htmlTag & "  <td>" & Range("AF" & line) & "</td>" & vbLf
      Else
        htmlTag = htmlTag & "  <td>clickAndWait</td>" & vbLf
        htmlTag = htmlTag & "  <td>" & Range("S" & line) & "</td>" & vbLf
      End If
      htmlTag = htmlTag & "  <td></td>" & vbLf
      htmlTag = htmlTag & "</tr>" & vbLf
      htmlTag = htmlTag & "<tr>" & vbLf
      htmlTag = htmlTag & "  <td>captureEntirePageScreenshot</td>" & vbLf
      htmlTag = htmlTag & "  <td>" & ActiveWorkbook.path & "\" & Title & "\capture\" & buf & "02.png</td>" & vbLf
      htmlTag = htmlTag & "  <td>background=#FFFFFF</td>" & vbLf
      htmlTag = htmlTag & "</tr>" & vbLf
      htmlTag = htmlTag & "</tbody></table></body></html>" & vbLf
  
  
    ' ADODB処理
    Set ObjADODB_TestCase = CreateObject("ADODB.Stream")
  
    'オブジェクトに保存するデータの種類を文字列型に指定する
    ObjADODB_TestCase.Type = 2
    
    '文字列型のオブジェクトの文字コードを指定する
    ObjADODB_TestCase.Charset = "UTF-8"
    ObjADODB_TestCase.LineSeparator = 10
  
    'オブジェクトのインスタンスを作成
    ObjADODB_TestCase.Open
    ObjADODB_TestCase.LoadFromFile (ActiveWorkbook.path & "\" & Title & "\" & buf)
    ObjADODB_TestCase.Position = ObjADODB_TestCase.Size 'ポインタを終端へ


    ' テストケース保存
    ObjADODB_TestCase.WriteText htmlTag, 1
    ObjADODB_TestCase.SaveToFile (ActiveWorkbook.path & "\" & Title & "\" & buf), 2
   
    'オブジェクトを閉じる
    ObjADODB_TestCase.Close
    Set ObjADODB_TestCase = Nothing
  
    End If
    buf = Dir()
  Loop

End Function


Function Selenium_MakeIndex()

  Dim htmlTag As String


  Dim buf As String
  htmlTag = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbLf & _
              "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">" & vbLf & _
              "<html xmlns=""http://www.w3.org/1999/xhtml"" xml:lang=""ja"" lang=""ja"">" & vbLf & _
              "<head profile=""http://selenium-ide.openqa.org/profiles/test-case"">" & vbLf & _
              "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"" />" & vbLf & _
              "<link rel=""selenium.base"" href=""" & baseURL & """ />" & vbLf & _
              "<title>" & TitleName & "</title>" & vbLf & _
              "</head>" & vbLf & _
              "<body>" & vbLf & _
              "<table id='suiteTable' cellpadding='1' cellspacing='1' border='1' class='selenium'><tbody>" & vbLf
              
              
  buf = Dir(ActiveWorkbook.path & "\" & Title & "\*.html")
  
  Do While Len(buf) > 0
    If LCase(buf) Like "*.html" Then
      If buf <> "00_index.html" Then
        htmlTag = htmlTag & "<tr><td><a href='" & buf & "'>" & buf & "</a></td></tr>" & vbLf
      End If
    End If
    buf = Dir()
  Loop
  htmlTag = htmlTag & "</tbody></table></body></html>" & vbLf
  
    ' ADODB処理
    Set ObjADODB_TestCase = CreateObject("ADODB.Stream")
  
    'オブジェクトに保存するデータの種類を文字列型に指定する
    ObjADODB_TestCase.Type = 2
    
    '文字列型のオブジェクトの文字コードを指定する
    ObjADODB_TestCase.Charset = "UTF-8"
    ObjADODB_TestCase.LineSeparator = 10
  
    'オブジェクトのインスタンスを作成
    ObjADODB_TestCase.Open
    

    ' テストケース保存
    ObjADODB_TestCase.WriteText htmlTag, 1
    ObjADODB_TestCase.SaveToFile (ActiveWorkbook.path & "\" & Title & "\00_index.html"), 2
   
    'オブジェクトを閉じる
    ObjADODB_TestCase.Close
    Set ObjADODB_TestCase = Nothing
End Function


