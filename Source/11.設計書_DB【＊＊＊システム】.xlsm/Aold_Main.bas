Attribute VB_Name = "Aold_Main"
'' *********************************************************************
'' * �V�K�V�[�g�ǉ�(���{�����j���[����̌Ăяo��)
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
'' * SQL�쐬(���{�����j���[����̌Ăяo��)
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
'' * �S�V�[�g����SQL�쐬(���{�����j���[����̌Ăяo��)
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
'' * Flamework�pSQL�쐬(���{�����j���[����̌Ăяo��)
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
'' * �S�V�[�g����SFlamework�pQL�쐬(���{�����j���[����̌Ăяo��)
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
'' * ����͈͐ݒ�(���{�����j���[����̌Ăяo��)
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
'' * �T�[�o�[����e�[�u�����擾(���{�����j���[����̌Ăяo��)
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
'' * �T�[�o�[����J�������擾(���{�����j���[����̌Ăяo��)
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
'' * �e�[�u���ꗗ����(���{�����j���[����̌Ăяo��)
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
'' * �Z���ݒ�(���{�����j���[����̌Ăяo��)
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
'' * �V�[�g�ݒ�(���{�����j���[����̌Ăяo��)
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
'' * �Đݒ�{�^��
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
'' * �V�[�g�擾�{�^��
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
'' * �V�[�g�u�e�[�u�����X�g�v���A�N�e�B�u�ɂ���
'' *
'' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'' *********************************************************************
'Sub DisplayTableList()
'    Sheets("TBL���X�g").Select
'End Sub
'
'' *********************************************************************
'' * �J�������Ď擾
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
