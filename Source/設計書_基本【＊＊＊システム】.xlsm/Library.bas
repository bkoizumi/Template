Attribute VB_Name = "Library"

'***********************************************************************************************************************************************
' * ��ʕ`�ʐ���J�n
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Public Function StartScript()

  '�A�N�e�B�u�Z���̎擾
   SelectionCell = Selection.Address
  
  ' �}�N������ŃV�[�g��E�B���h�E���؂�ւ��̂������Ȃ��悤�ɂ��܂�
  Application.ScreenUpdating = False
  
  ' �}�N�����쎩�̂ŕʂ̃C�x���g�����������̂�}������
  Application.EnableEvents = False
  
  ' �}�N������ŃZ��ItemName�Ȃǂ��ς�鎞�����v�Z��������x������̂������
  Application.Calculation = xlCalculationManual
  
  ' �}�N�����쒆�Ɉ�؂̃L�[��}�E�X����𐧌�����
'  Application.Interactive = False
  
  ' �}�N�����쒆�̓}�E�X�J�[�\�����u�����v�v�ɂ���
'  Application.Cursor = xlWait
  
  ' �m�F���b�Z�[�W���o���Ȃ�
  Application.DisplayAlerts = False

End Function

'***********************************************************************************************************************************************
' * ��ʕ`�ʐ���I��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Public Function EndScript()

  ' �}�N������ŃV�[�g��E�B���h�E���؂�ւ��̂������Ȃ��悤�ɂ��܂�
  Application.ScreenUpdating = True
  
  ' �}�N�����쎩�̂ŕʂ̃C�x���g�����������̂�}������
  Application.EnableEvents = True
  
  ' �}�N������ŃZ��ItemName�Ȃǂ��ς�鎞�����v�Z��������x������̂������
  Application.Calculation = xlCalculationAutomatic
  
  ' �}�N�����쒆�Ɉ�؂̃L�[��}�E�X����𐧌�����
'  Application.Interactive = True
  
  ' �}�N������I����̓}�E�X�J�[�\�����u�f�t�H���g�v�ɂ��ǂ�
  Application.Cursor = xlDefault
  
  ' �}�N������I����̓X�e�[�^�X�o�[���u�f�t�H���g�v�ɂ��ǂ�
  Application.StatusBar = False

  ' �m�F���b�Z�[�W���o���Ȃ�
  Application.DisplayAlerts = True
  
  ' �����I�ɍČv�Z������
  'Application.CalculateFull
  
'  �A�N�e�B�u�Z���̑I��
'  If SelectionCell <> "" Then
'    Range(SelectionCell).Select
'  End If

End Function
'**************************************************************************************************
' * �f�o�b�O�p��ʕ\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showDebugForm(ByVal meg1 As String, Optional meg2 As String)
  Dim runTime As Date
  Dim StartUpPosition As Long

'  On Error GoTo catchError

  runTime = Format(Now(), "yyyy/mm/dd hh:nn:ss")

  meg1 = Replace(meg1, vbNewLine, " ")
  
  Select Case setVal("debugMode")
    Case "develop"
      Debug.Print runTime & vbTab & meg1
    Case Else
      Exit Function
  End Select
  
  DoEvents
  Exit Function

'�G���[������=====================================================================================
catchError:
  Exit Function
End Function

'**************************************************************************************************
' * �������ʒm
' *
' * Worksheets("Notice").Visible = True
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showNotice(Code As Long, Optional process As String, Optional runEndflg As Boolean)
  Dim Message As String
  Dim runTime As Date
  Dim endLine As Long

  On Error GoTo catchError


  runTime = Format(Now(), "yyyy/mm/dd hh:nn:ss")

  endLine = BK_sheetNotice.Cells(Rows.count, 1).End(xlUp).Row
  Message = Application.WorksheetFunction.VLookup(Code, BK_sheetNotice.Range("A2:B" & endLine), 2, False)
  Message = Replace(Message, "%%", process)
  If process = "" Then
    Message = Replace(Message, "<>", process)
  End If
  If runEndflg = True Then
    Message = Message & vbNewLine & "�����𒆎~���܂�"
  End If

  If StopTime <> 0 Then
    Message = Message & vbNewLine & "<�������ԁF" & StopTime & ">"
  End If

  If Message <> "" Then
    Message = Replace(Message, "<BR>", vbNewLine)
  End If

  If setVal("debugMode") = "speak" Or setVal("debugMode") = "develop" Or setVal("debugMode") = "all" Then
    Application.Speech.Speak Text:=Message, SpeakAsync:=True, SpeakXML:=True
  End If

  Select Case Code
    Case 0 To 399
      Call MsgBox(Message, vbInformation, thisAppName)

    Case 400 To 499
      Call MsgBox(Message, vbCritical, thisAppName)

    Case 500 To 599
      Call MsgBox(Message, vbExclamation, thisAppName)

    Case 999

    Case Else
      Call MsgBox(Message, vbCritical, thisAppName)
  End Select
  
  Message = Replace(Message, vbNewLine & "�����𒆎~���܂�", "�B�����𒆎~���܂�")
  Message = "[" & Code & "]" & Message
  
  '��ʕ`�ʐ���I������
  If runEndflg = True Then
    Call EndScript
    Call Ctl_ProgressBar.ShowEnd
    End
  Else
    Call Library.showDebugForm(Message)
  End If

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call MsgBox(Message, vbCritical, thisAppName)

End Function
