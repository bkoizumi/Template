Attribute VB_Name = "Ctl_Zoom"
Option Explicit

'**************************************************************************************************
' * �I���Z���̊g��\��/�I��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ZoomIn(Optional slctCellAddress As String)
  Dim cellVal As String
  Dim topPosition As Long, leftPosition As Long
  Dim cellWidth As Long
  
  If slctCellAddress = "" Then
    
  End If
  
  If ActiveCell.HasFormula = False Then
    cellVal = ActiveCell.text
  Else
    cellVal = ActiveCell.Formula
  End If
  Set targetBook = ActiveWorkbook
  Set targetSheet = ActiveSheet
  
  With Frm_Zoom
    .StartUpPosition = 1
    .TextBox = cellVal
    .TextBox.MultiLine = True
    .TextBox.MultiLine = True
    .TextBox.EnterKeyBehavior = True
    
    
    If ActiveCell.HasFormula = False Then
      .TextBox.IMEMode = fmIMEModeOff
    Else
      .TextBox.IMEMode = ActiveCell.Validation.IMEMode
    End If
    
    .Label1.Caption = "�I���Z���F" & ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    .Show vbModeless
  
  End With
End Function


'==================================================================================================
Function ZoomOut(text As String, SetTargetAddress As String)
  
  SetTargetAddress = Replace(SetTargetAddress, "�I���Z���F", "")
  
  targetBook.Activate
  targetSheet.Activate
  Range(SetTargetAddress).Value = text
  
  Call Library.endScript
End Function


