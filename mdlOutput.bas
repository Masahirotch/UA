Attribute VB_Name = "mdlOutput"
Option Compare Database
Option Explicit

'#-----------------------------------------------------------------
'#�v���V�[�W�����@:funOutPutProc1
'#�����@�@�@�@�@�@:�Ƒ��J�[�h�\�����o��
'#�����@�@�@�@�@�@:pstrTemp: �e���v���[�g�t�@�C���p�X
'#�@�@�@�@�@�@�@�@:pstrOut : �o�͐�p�X
'#�߂�l�@�@�@�@�@:Ture:����/False:�ُ�
'#-----------------------------------------------------------------
Public Function funOutPutProc1( _
                    ByVal pstrTemp As String, _
                    ByVal pstrOut As String _
                ) As Boolean
Dim blnSyokai       As Boolean      '����t���O
Dim lngRow          As Long         '�s
Dim lngCopyRow      As Long         '�R�s�[�J�n�s
Dim strFileNm       As String       'Excel�t�@�C����
Dim strSyainNoN     As String       '�Ј��ԍ�(�V)
Dim strSyainNoO     As String       '�Ј��ԍ�(��)
Dim varRowData()    As Variant
Dim adoRs           As New ADODB.Recordset
Dim ExApp           As Excel.Application
Dim ExBook          As Excel.Workbook
Dim ExSheet         As Excel.Worksheet

On Error GoTo Err_Exit

    funOutPutProc1 = False

'------------------------------
'���Ƒ��J�[�h�f�[�^�擾
'------------------------------
    '�����R�[�h�Z�b�g�쐬
    Call adoRs.Open(funMakeSQL1(), goADOCon, adOpenStatic, adLockReadOnly)

    If (adoRs.EOF) Then
        '���f�[�^���Ȃ��ꍇ
        Call MsgBox(mdlCommon.funGetMsgW("010"), vbExclamation, Sys_Title)
        Call mdlCommon.subCloseRecordset(adoRs)
        Exit Function
    End If

'------------------------------
'��Excel�I�u�W�F�N�g�̏���
'------------------------------
    Set ExApp = CreateObject("Excel.Application")
    ExApp.Application.DisplayAlerts = False
    Call subExcelObjet(pstrTemp & EXCEL_TEMP_FILE1, ExApp, ExBook, ExSheet, True)

'------------------------------
'��Excel�o��
'------------------------------
    strSyainNoO = ""
    lngCopyRow = 50
    blnSyokai = True
    Erase varRowData:    ReDim varRowData(EXCEL_ROW_DATA1, EXCEL_COL_DATA1)

    Do Until adoRs.EOF
        '���Ј��ԍ�
        strSyainNoN = adoRs.Fields("EmployeeNo").Value

        If (strSyainNoO <> strSyainNoN) Then
            '���Ј��ԍ����ς������

            If (blnSyokai = False) Then
                '������ȊO�̏ꍇ�̏���

                '���f�[�^�ꊇ�o��
                Call subDataOut(ExSheet, varRowData, lngCopyRow - 22)

                '��Excel�ۑ�
                strFileNm = "�Ƒ��J�[�h�\����(" & strSyainNoO & ").xls"
                Call sudExcelSave(ExApp, ExSheet, pstrOut & strFileNm, strFileNm)

                '���I�u�W�F�N�g����
                Call subExcelObjet(pstrTemp & EXCEL_TEMP_FILE1, ExApp, ExBook, ExSheet, False)
            End If

            With ExSheet
                '���V�[�g��
                .Name = strSyainNoN
                '������
                .Cells(4, 2) = "�����@�F�@" & "" & adoRs.Fields("�������̖�").Value
                '���Ј��ԍ�
                .Cells(4, 5) = "�Ј��ԍ��@�F�@" & adoRs.Fields("EmployeeNo").Value
                '���Ј��ԍ�
                .Cells(5, 2) = "�����@�F�@" & adoRs.Fields("����").Value
            End With

            '���s�J�n�ʒu
            lngRow = 0
            lngCopyRow = 50
            strSyainNoO = strSyainNoN
            blnSyokai = False
            Erase varRowData:    ReDim varRowData(EXCEL_ROW_DATA1, EXCEL_COL_DATA1)
        End If

        If (lngRow >= 10) Then
            '�����׌������P�O���𒴂����ꍇ�A���s����
            
            '���f�[�^�ꊇ�o��
            Call subDataOut(ExSheet, varRowData, lngCopyRow - 22)

            '���Z���R�s�[
            With ExSheet
                .Activate
                .Select
                .Rows("1:49").Copy
                .Range("A" & CStr(lngCopyRow)).Select
                .Paste
            End With

            '���ϐ�������
            lngRow = 0
            lngCopyRow = lngCopyRow + 49
            Erase varRowData:    ReDim varRowData(EXCEL_ROW_DATA1, EXCEL_COL_DATA1)
        End If

        '���Ƒ�����
        varRowData(lngRow, 0) = "" & adoRs.Fields("FamilyNm").Value
        '������
        varRowData(lngRow, 2) = "" & adoRs.Fields("RelationShipNm").Value
        '���t���K�i
        varRowData(lngRow, 3) = "" & adoRs.Fields("Furigana").Value

        lngRow = lngRow + 1
        adoRs.MoveNext
    Loop

    '���f�[�^�ꊇ�o��
    Call subDataOut(ExSheet, varRowData, lngCopyRow - 22)

    '��Excel�ۑ�
    strFileNm = "�Ƒ��J�[�h�\����(" & strSyainNoO & ").xls"
    Call sudExcelSave(ExApp, ExSheet, pstrOut & strFileNm, strFileNm)

    funOutPutProc1 = True
    GoTo End_Proc

Err_Exit:
    Call MsgBox(mdlCommon.funGetMsgE("028", Err.Number, Err.Description), vbCritical, Sys_Title)
    Call mdlCommon.subCloseRecordset(adoRs)

End_Proc:
    ExBook.Close
    Set ExBook = Nothing
    Set ExSheet = Nothing

    ExApp.Application.DisplayAlerts = True
    ExApp.Quit

    Set ExApp = Nothing

End Function

'#-----------------------------------------------------------------
'#�v���V�[�W�����@:funMakeSQL1
'#�����@�@�@�@�@�@:�r�p�k���쐬(�Ƒ��J�[�h�\�����o��)
'#�����@�@�@�@�@�@:�Ȃ�
'#�߂�l�@�@�@�@�@:�r�p�k��
'#-----------------------------------------------------------------
Private Function funMakeSQL1() As String
Dim strSQL          As String       '�r�p�k��

    '���r�p�k�ҏW
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "  A.*, "
    strSQL = strSQL & "  B.RelationShipNm, "
    strSQL = strSQL & "  E.����, "
'2012/03/08 No8 Upd-Start
'    strSQL = strSQL & "  E.�������̖� "
    strSQL = strSQL & "  E.�Ɩ��� �������̖� "
'2012/03/08 No8 Upd-End
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "  T_Family A "
    strSQL = strSQL & "LEFT JOIN "
    strSQL = strSQL & " (SELECT * FROM M_RelationShip WHERE DelFlg = '" & CST_DEL_FLG & "') B "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  A.RelationShipCd = B.RelationShipCd "
    strSQL = strSQL & "LEFT JOIN "
    strSQL = strSQL & gstrLinkDB & "�Ј��Œ��� E "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  A.EmployeeNo = E.�Ј��ԍ� "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "  EXISTS ("
    strSQL = strSQL & "    SELECT "
    strSQL = strSQL & "      * "
    strSQL = strSQL & "    FROM ("
    strSQL = strSQL & "      SELECT "
    strSQL = strSQL & "        EmployeeNo, "
    strSQL = strSQL & "        FamilyNo, "
    strSQL = strSQL & "        MAX(SeqNo) SeqNo "
    strSQL = strSQL & "      FROM "
    strSQL = strSQL & "        T_Family "
    strSQL = strSQL & "      WHERE "
    strSQL = strSQL & "        DelFlg = '" & CST_DEL_FLG & "' "
    strSQL = strSQL & "      GROUP BY "
    strSQL = strSQL & "        EmployeeNo, "
    strSQL = strSQL & "        FamilyNo) E "
    strSQL = strSQL & "    WHERE "
    strSQL = strSQL & "      E.EmployeeNo = A.EmployeeNo AND "
    strSQL = strSQL & "      E.FamilyNo   = A.FamilyNo   AND "
    strSQL = strSQL & "      E.SeqNo      = A.SeqNo) AND "
    strSQL = strSQL & "  A.ApprovalYmd >= '" & gtApply(0).ApprovalYmdF & "' AND "
    strSQL = strSQL & "  A.ApprovalYmd <= '" & gtApply(0).ApprovalYmdT & "' AND "
'    strSQL = strSQL & "  A.Status       = '" & CST_STATUS_2 & "' AND "
    strSQL = strSQL & "  A.DelFlg       = '" & CST_DEL_FLG & "' "
    strSQL = strSQL & "ORDER BY "
    strSQL = strSQL & "  EmployeeNo, "
    strSQL = strSQL & "  FamilyNo "

    funMakeSQL1 = strSQL

End Function

'#-----------------------------------------------------------------
'#�v���V�[�W�����@:subExcelObjet
'#�����@�@�@�@�@�@:Excel�I�u�W�F�N�g�쐬
'#�����@�@�@�@�@�@:pstrTemp: �e���v���[�g�t�@�C���p�X
'#�@�@�@�@�@�@�@�@:pExApp  : Excel�A�v���P�[�V�����I�u�W�F�N�g
'#�@�@�@�@�@�@�@�@:pExBook : ���[�N�u�b�N�I�u�W�F�N�g
'#�@�@�@�@�@�@�@�@:pExSheet: ���[�N�V�[�g�I�u�W�F�N�g
'#�@�@�@�@�@�@�@�@:pblnFlg : True:�t�@�C���I�[�v��
'#�߂�l�@�@�@�@�@:�Ȃ�
'#-----------------------------------------------------------------
Private Sub subExcelObjet( _
                ByVal pstrTemp As String, _
                ByRef pExApp As Object, _
                ByRef pExBook As Object, _
                ByRef pExSheet As Object, _
                ByVal pblnFlg As Boolean _
            )
Dim objExBook       As Object

'-----------------------------------
'��Excel�I�u�W�F�N�g�̏���
'-----------------------------------
    If (pblnFlg) Then Set pExBook = pExApp.Application.Workbooks.Open(pstrTemp)
    pExBook.Activate
    pExBook.Sheets(Array("Temp")).Copy

    Set pExSheet = pExApp.Sheets(1)

End Sub

'#-----------------------------------------------------------------
'#�v���V�[�W�����@:subDataOut
'#�����@�@�@�@�@�@:Excel�Ƀf�[�^�o��
'#�����@�@�@�@�@�@:pExSheet: ���[�N�V�[�g�I�u�W�F�N�g
'#�@�@�@�@�@�@�@�@:pData() : �o�̓f�[�^
'#�߂�l�@�@�@�@�@:�Ȃ�
'#-----------------------------------------------------------------
Private Sub subDataOut( _
                ByRef pExSheet As Object, _
                ByRef pData() As Variant, _
                ByVal plngRow As Long _
            )

    With pExSheet
        .Range(.Cells(plngRow, 3), .Cells(plngRow + 9, 7)).Value = pData
    End With

End Sub

'#-----------------------------------------------------------------
'#�v���V�[�W�����@:sudExcelSave
'#�����@�@�@�@�@�@:Excel�f�[�^�ۑ�
'#�����@�@�@�@�@�@:pExApp  : Excel�A�v���P�[�V�����I�u�W�F�N�g
'#    �@�@�@�@�@�@:pExSheet  : ���[�N�V�[�g�I�u�W�F�N�g
'#�@�@�@�@�@�@�@�@:pstrPath  : �ۑ�����t�@�C���p�X
'#�@�@�@�@�@�@�@�@:pstrFileNm: �t�@�C����
'#�߂�l�@�@�@�@�@:�Ȃ�
'#-----------------------------------------------------------------
Private Sub sudExcelSave( _
                ByRef pExApp As Object, _
                ByRef pExSheet As Object, _
                ByVal pstrPath As String, _
                ByVal pstrFileNm As String _
            )
Dim ExBookWk        As Excel.Workbook

    With pExSheet
        '��Excel�ۑ�
        .Activate
        .Cells(1, 1).Select
'2012/11/02 Windows7�Ή� Upd-Start
'        .SaveAs Filename:=pstrPath, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
        If (pExApp.Version < Excel_Version_2007) Then
            .SaveAs Filename:=pstrPath, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
        Else
            .SaveAs Filename:=pstrPath, FileFormat:=Excel_FileFormat_xlExcel8, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
        End If
'2012/11/02 Windows7�Ή� Upd-End

        If (gtApply(0).PrintFlg = CST_CHK_ON) Then
            '�����
            .PrintOut
        End If
    End With


    '���ۑ������t�@�C�������
    For Each ExBookWk In pExApp.Workbooks
        If (ExBookWk.Name = pstrFileNm) Then
            ExBookWk.Close
            Exit For
        End If
    Next

End Sub

'#-----------------------------------------------------------------
'#�v���V�[�W�����@:funOutPutProc2
'#�����@�@�@�@�@�@:�Ƒ��J�[�h�\�����o��
'#�����@�@�@�@�@�@:pstrTemp: �e���v���[�g�t�@�C���p�X
'#�@�@�@�@�@�@�@�@:pstrOut : �o�͐�p�X
'#�߂�l�@�@�@�@�@:Ture:����/False:�ُ�
'#-----------------------------------------------------------------
Public Function funOutPutProc2( _
                    ByVal pstrTemp As String, _
                    ByVal pstrOut As String _
                ) As Boolean
Dim lngRow          As Long         '�s
Dim lngCount        As Long         '���[�N�J�E���^
Dim varRowData()    As Variant
Dim adoRs           As New ADODB.Recordset
Dim ExApp           As Excel.Application
Dim ExBook          As Excel.Workbook
Dim ExSheet         As Excel.Worksheet
Dim ExBookWk        As Excel.Workbook

On Error GoTo Err_Exit

    funOutPutProc2 = False

'------------------------------
'���Ƒ��J�[�h�f�[�^�擾
'------------------------------
    '�����R�[�h�Z�b�g�쐬
    Call adoRs.Open(funMakeSQL2(), goADOCon, adOpenStatic, adLockReadOnly)

    If (adoRs.EOF) Then
        '���f�[�^���Ȃ��ꍇ
        Call MsgBox(mdlCommon.funGetMsgW("010"), vbExclamation, Sys_Title)
        Call mdlCommon.subCloseRecordset(adoRs)
        Exit Function
    End If

'------------------------------
'��Excel�I�u�W�F�N�g�̏���
'------------------------------
    Set ExApp = CreateObject("Excel.Application")
    ExApp.Application.DisplayAlerts = False
    Call subExcelObjet(pstrTemp & EXCEL_TEMP_FILE2_1, ExApp, ExBook, ExSheet, True)

'------------------------------
'��Excel�o��
'------------------------------
    Erase varRowData:    ReDim varRowData(EXCEL_COL_DATA2)

    '���s�J�n�ʒu
    lngRow = EXCEL_ROW_START2:  lngCount = 0
    Erase gtFamily:

    Do Until adoRs.EOF
        If ("" & adoRs.Fields("IssueYmd").Value = "") Then
            lngCount = lngCount + 1
            ReDim Preserve gtFamily(1 To lngCount)
            '���Ј��ԍ�
            gtFamily(lngCount).EmployeeNo = adoRs.Fields("EmployeeNo").Value
            '���Ƒ��ԍ�
            gtFamily(lngCount).FamilyNo = adoRs.Fields("FamilyNo").Value
            '���}��
            gtFamily(lngCount).SeqNo = adoRs.Fields("SeqNo").Value
        End If

        '���Ј��ԍ�
        varRowData(0) = "'" & adoRs.Fields("EmployeeNo").Value
        '���Ј���
        varRowData(1) = "" & adoRs.Fields("����").Value
        '���t���K�i
        varRowData(2) = ""
        '�������R�[�h
        varRowData(3) = ""
        '������������
        varRowData(4) = "" & adoRs.Fields("DepartmentNm").Value
        '�����Ƒ���
        varRowData(5) = "" & adoRs.Fields("FamilyNm").Value
        '������
        varRowData(6) = "" & adoRs.Fields("RelationShipNm").Value
        '��(�ض��)
        varRowData(7) = "" & adoRs.Fields("Furigana").Value
'2013/01/28 Upd-Start
        varRowData(8) = "" & adoRs.Fields("PeriodValidNm").Value
'2013/01/28 Upd-End

        '���s���ꊇ�o��
        With ExSheet
            .Range(.Cells(lngRow, 1), .Cells(lngRow, EXCEL_COL_DATA2 + 1)).Value = varRowData
        End With

        lngRow = lngRow + 1
        adoRs.MoveNext
    Loop

    With ExSheet
        '��Excel�ۑ�
        .Activate
        .Cells(1, 1).Select
'2012/11/02 Windows7�Ή� Upd-Start
'        .SaveAs Filename:=pstrOut, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
        If (ExApp.Version < Excel_Version_2007) Then
            .SaveAs Filename:=pstrOut, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
        Else
            .SaveAs Filename:=pstrOut, FileFormat:=Excel_FileFormat_xlExcel8, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
        End If
'2012/11/02 Windows7�Ή� Upd-End
    End With

    '���ۑ������t�@�C�������
    For Each ExBookWk In ExApp.Workbooks
        If (ExBookWk.Name = EXCEL_TEMP_FILE2_2) Then
            ExBookWk.Close
            Exit For
        End If
    Next

    If (lngCount = 0) Then ReDim gtFamily(1)

    funOutPutProc2 = True
    GoTo End_Proc

Err_Exit:
    Call MsgBox(mdlCommon.funGetMsgE("031", Err.Number, Err.Description), vbCritical, Sys_Title)
    Call mdlCommon.subCloseRecordset(adoRs)

End_Proc:
    ExBook.Close
    Set ExBook = Nothing
    Set ExSheet = Nothing

    ExApp.Application.DisplayAlerts = True
    ExApp.Quit

    Set ExApp = Nothing

End Function

'#-----------------------------------------------------------------
'#�v���V�[�W�����@:funMakeSQL2
'#�����@�@�@�@�@�@:�r�p�k���쐬(���Ղ��Ɨp�f�[�^�o��)
'#�����@�@�@�@�@�@:�Ȃ�
'#�߂�l�@�@�@�@�@:�r�p�k��
'#-----------------------------------------------------------------
Private Function funMakeSQL2() As String
Dim strSQL          As String       '�r�p�k��

    '���r�p�k�ҏW
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "  A.*, "
    strSQL = strSQL & "  B.RelationShipNm, "
'2013/01/28 Upd-Start
    strSQL = strSQL & "  C.PeriodValidNm, "
'2013/01/28 Upd-End
    strSQL = strSQL & "  E.���� "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "  T_Family A "
    strSQL = strSQL & "LEFT JOIN "
    strSQL = strSQL & " (SELECT RelationShipCd, RelationShipNm FROM M_RelationShip WHERE DelFlg = '" & CST_DEL_FLG & "') B "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  A.RelationShipCd = B.RelationShipCd "
    strSQL = strSQL & "LEFT JOIN "
    strSQL = strSQL & gstrLinkDB & "�Ј��Œ��� E "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  A.EmployeeNo = E.�Ј��ԍ� "
'2013/01/28 Upd-Start
    strSQL = strSQL & "LEFT JOIN "
    strSQL = strSQL & "  (SELECT PeriodValidCd, PeriodValidNm FROM M_PeriodValid WHERE DelFlg = '" & CST_DEL_FLG & "') C "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  A.PeriodValidCd = C.PeriodValidCd "
'2013/01/28 Upd-End
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "  EXISTS ("
    strSQL = strSQL & "    SELECT "
    strSQL = strSQL & "      * "
    strSQL = strSQL & "    FROM ("
    strSQL = strSQL & "      SELECT "
    strSQL = strSQL & "        EmployeeNo, "
    strSQL = strSQL & "        FamilyNo, "
    strSQL = strSQL & "        MAX(SeqNo) SeqNo "
    strSQL = strSQL & "      FROM "
    strSQL = strSQL & "        T_Family "
    strSQL = strSQL & "      WHERE "
    strSQL = strSQL & "        DelFlg = '" & CST_DEL_FLG & "' "
    strSQL = strSQL & "      GROUP BY "
    strSQL = strSQL & "        EmployeeNo, "
    strSQL = strSQL & "        FamilyNo) F "
    strSQL = strSQL & "    WHERE "
    strSQL = strSQL & "      F.EmployeeNo = A.EmployeeNo AND "
    strSQL = strSQL & "      F.FamilyNo   = A.FamilyNo   AND "
    strSQL = strSQL & "      F.SeqNo      = A.SeqNo) AND "
    strSQL = strSQL & "  A.ApprovalYmd >= '" & gtApply(0).ApprovalYmdF & "' AND "
    strSQL = strSQL & "  A.ApprovalYmd <= '" & gtApply(0).ApprovalYmdT & "' AND "
'    strSQL = strSQL & "  A.Status       = '" & CST_STATUS_2 & "' AND "
    strSQL = strSQL & "  A.DelFlg       = '" & CST_DEL_FLG & "' "
    strSQL = strSQL & "ORDER BY "
    strSQL = strSQL & "  EmployeeNo, "
    strSQL = strSQL & "  FamilyNo "

    funMakeSQL2 = strSQL

End Function



