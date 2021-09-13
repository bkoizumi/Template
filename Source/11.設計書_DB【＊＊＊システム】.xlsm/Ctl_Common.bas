Attribute VB_Name = "Ctl_Common"
Option Explicit

'**************************************************************************************************
' * ���ʏ���
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'==================================================================================================
Function ClearData()
  Dim line As Long, endLine As Long

  '�����J�n----------------------------------------------------------------------------------------
  'On Error GoTo catchError
  '�����l�ݒ�----------------
  FuncName = "Ctl_Common.ClearData"
  '--------------------------
  
  Call init.Setting
  Call Library.showDebugForm(FuncName & "=============================================")
  
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  line = startLine
  
  Do Until Range("A" & line) = "End"
    If Range("A" & line) = "" Then
      Rows(line & ":" & line).Delete Shift:=xlUp
      line = line - 1
    ElseIf Range("A" & line) = "Column" Then
      On Error Resume Next
      Range("B" & line & ":AA" & line).SpecialCells(xlCellTypeConstants, 23).ClearContents
      On Error GoTo catchError
    
    End If
'    Call Ctl_ProgressBar.showBar(thisAppName, 1, PrgP_Max, line, endLine, "�f�[�^�N���A")
    line = line + 1
  Loop
  Columns("L:R").EntireColumn.Hidden = True
  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, FuncName & vbNewLine & Err.Number & "�F" & Err.Description, True)
End Function


'==================================================================================================
Function addRow(line As Long)

  If line >= startLine + 2 Then
    Rows(line & ":" & line).Copy
    Rows(line & ":" & line).Insert Shift:=xlDown
    Range("A" & line) = ""
    Application.CutCopyMode = False
  End If
End Function


'==================================================================================================
Function chkIndexRow()
  Dim line As Long, endLine As Long
  Dim IndexRow As Long
  
  endLine = Cells(Rows.count, 1).End(xlUp).Row
  For line = startLine To endLine
    If Range("A" & line) = "IndexStart" Then
      IndexRow = line
      Exit For
    End If
  Next
  Call Library.showDebugForm("IndexRow�F" & IndexRow)
  chkIndexRow = IndexRow
End Function


'==================================================================================================
Function addSheet(newSheetName As String)
  
  On Error GoTo catchError
  
  If Library.chkSheetExists(newSheetName) = False Then
    sheetCopy.Copy After:=Worksheets(Worksheets.count)
    ActiveWorkbook.Sheets(Worksheets.count).Name = newSheetName
  End If
  Sheets(newSheetName).Select
    
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, Err.Description, True)
End Function


'==================================================================================================
Function makeTblList()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim sheetList As Object
  Dim targetSheet   As Worksheet
  Dim sheetName As String
  
'  On Error GoTo catchError

  Call init.Setting
  sheetTblList.Select
  endLine = sheetTblList.Cells(Rows.count, 3).End(xlUp).Row + 1
  Range("C6:U" & endLine).ClearContents
  
  With Range("B6:U" & endLine).Interior
    .Pattern = xlPatternNone
    .Color = xlNone
  End With
  
      
  line = 6
  For Each sheetList In Sheets
    sheetName = sheetList.Name
    
    
    Select Case sheetName
      Case "�ݒ�", "Notice", "DataType", "�R�s�[�p", "�\��", "TBL���X�g", "�ύX����", "ER�}"
      Case Else
    
        sheetTblList.Range("B" & line).FormulaR1C1 = "=ROW()-5"
        sheetTblList.Range("C" & line) = Sheets(sheetName).Range("B2")
  '      sheetTblList.Range("E" & line) = Sheets(sheetName).Range("D5")
  '      sheetTblList.Range("H" & line) = Sheets(sheetName).Range("H5")
        sheetTblList.Range("Q" & line) = Sheets(sheetName).Range("D6")
      
        '�_���e�[�u����
        If Sheets(sheetName).Range("D5") <> "" Then
          With sheetTblList.Range("E" & line)
            .Value = Sheets(sheetName).Range("D5")
            .Select
            .Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=sheetName & "!" & "A9"
            .Font.Color = RGB(0, 0, 0)
            .Font.Underline = False
            .Font.Size = 10
            .Font.Name = "���C���I"
          End With
        End If
        
         '�����e�[�u����
        With sheetTblList.Range("H" & line)
          .Value = Sheets(sheetName).Range("H5")
          .Select
          .Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=sheetName & "!" & "A9"
          .Font.Color = RGB(0, 0, 0)
          .Font.Underline = False
          .Font.Size = 10
          .Font.Name = "���C���I"
        End With
        
        ' �V�[�g�F�Ɠ����F���Z���ɐݒ�
        If Sheets(sheetName).Tab.Color Then
          With sheetTblList.Range("B" & line & ":U" & line).Interior
            .Pattern = xlPatternNone
            .Color = Sheets(sheetName).Tab.Color
          End With
        End If
        
        
        line = line + 1
    End Select
  Next
  
  
  '�����I��--------------------------------------
  Call Ctl_ProgressBar.showEnd
  Call Library.endScript
    
    
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, Err.Description, True)
End Function

