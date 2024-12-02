VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmOutput1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Database

'#-----------------------------------------------------------------
'#�v���V�[�W�����@:funCheckInput
'#�����@�@�@�@�@�@:���̓`�F�b�N
'#�����@�@�@�@�@�@:pstrPath: �e���v���[�g�p�X
'#�߂�l�@�@�@�@�@:True:����/False:�ُ�
'#-----------------------------------------------------------------
Private Function funCheckInput(ByRef pstrPath As String) As Boolean

    funCheckInput = False

    Erase gtApply:  ReDim gtApply(0)

    With gtApply(0)
'-----------------------------------
'�����F��(From)
'-----------------------------------
        .ApprovalYmdF = "" & txtApprovalYmdF.Value
        If (Len(.ApprovalYmdF) = 0) Then
            Call MsgBox(mdlCommon.funGetMsgW("015"), vbExclamation, Sys_Title)
            Exit Function
        End If

'-----------------------------------
'�����F��(To)
'-----------------------------------
        .ApprovalYmdT = "" & txtApprovalYmdT.Value
        If (Len(.ApprovalYmdT) = 0) Then .ApprovalYmdT = .ApprovalYmdF

'-----------------------------------
'���֘A�`�F�b�N
'-----------------------------------
        If (.ApprovalYmdF > .ApprovalYmdT) Then
            Call MsgBox(mdlCommon.funGetMsgW("016"), vbExclamation, Sys_Title)
            Exit Function
        End If

'-----------------------------------
'�����
'-----------------------------------
        .PrintFlg = chkPrint.Value
    End With

'-----------------------------------
'��EXCEL�e���v���[�g�p�X�`�F�b�N
'-----------------------------------
    pstrPath = CurrentProject.Path
    If (Right$(pstrPath, 1) <> "\") Then pstrPath = pstrPath & "\"
    pstrPath = pstrPath & EXCEL_TEMP_PATH
    If (Right$(pstrPath, 1) <> "\") Then pstrPath = pstrPath & "\"

    '��EXCEL�e���v���[�g���݃`�F�b�N
    If (Len(Dir(pstrPath & EXCEL_TEMP_FILE1)) = 0) Then
        '�����݂��Ȃ��ꍇ
        Call MsgBox(mdlCommon.funGetMsgW("018"), vbExclamation, Sys_Title)
        Exit Function
    End If

    funCheckInput = True

End Function

'#-----------------------------------------------------------------
'#�v���V�[�W�����@:cmdCancel_Click
'#�����@�@�@�@�@�@:�L�����Z���{�^���N���b�N��
'#-----------------------------------------------------------------
Private Sub cmdCancel_Click()

'-----------------------------------
'����ʂ����
'-----------------------------------
    Call DoCmd.Close(acForm, Me.Name)
    Call DoCmd.OpenForm("frmMenuOutput")

End Sub

'#-----------------------------------------------------------------
'#�v���V�[�W�����@:cmdOutput_Click
'#�����@�@�@�@�@�@:�o�̓{�^���N���b�N��
'#-----------------------------------------------------------------
Private Sub cmdOutput_Click()
Dim blnRtn          As Boolean      '�߂�l
Dim intRtn          As Integer      '�߂�l
Dim strTemp         As String       '�e���v���[�g�p�X
Dim strOut          As String       '�o�͐�p�X

'-----------------------------------
'�����̓`�F�b�N
'-----------------------------------
    blnRtn = funCheckInput(strTemp)
    If (blnRtn = False) Then Exit Sub

'-----------------------------------
'���m�F���b�Z�[�W
'-----------------------------------
    intRtn = MsgBox(mdlCommon.funGetMsgI("005"), vbInformation + vbYesNo, Sys_Title)
    If (intRtn = vbNo) Then Exit Sub

'-----------------------------------
'���t�H���_�Q�ƃ_�C�A���O�\��
'-----------------------------------
    strOut = mdlCommon.funFileSaveDialog1(mdlCommon.funGetMsgI("007"))
    If (Len(strOut) = 0) Then Exit Sub

    If (Right(strOut, 1) <> "\") Then strOut = strOut & "\"

    '���A�C�R���i�ҋ@�j
    Screen.MousePointer = CST_MP_H
DoEvents

'-----------------------------------
'���Ƒ��J�[�h�\�����o��
'-----------------------------------
    blnRtn = mdlOutput.funOutPutProc1(strTemp, strOut)
    If (blnRtn = False) Then GoTo End_Proc

    Call MsgBox(mdlCommon.funGetMsgI("006"), vbInformation, Sys_Title)

End_Proc:
    '���A�C�R���i�f�t�H���g�j
    Screen.MousePointer = CST_MP_D

End Sub

'#-----------------------------------------------------------------
'#�v���V�[�W�����@:Form_Load
'#�����@�@�@�@�@�@:��ʋN����
'#-----------------------------------------------------------------
Private Sub Form_Load()

'-----------------------------------
'����������
'-----------------------------------
    '����ʃT�C�Y�ݒ�
    Call mdlCommon.subGamenSizeCM(12, 9)

    chkPrint.Value = CST_CHK_OFF

End Sub

'#-----------------------------------------------------------------
'#�v���V�[�W�����@:txtApprovalYmdF_DblClick
'#�����@�@�@�@�@�@:�_�u���N���b�N��
'#-----------------------------------------------------------------
Private Sub txtApprovalYmdF_DblClick(Cancel As Integer)

'-----------------------------------
'���J�����_�[�\��
'-----------------------------------
    Call DoCmd.OpenForm("frmCalender", , , , , acDialog)

    If (Len(gstrCalender) > 0) Then
        txtApprovalYmdF.Value = gstrCalender
    End If

End Sub

'#-----------------------------------------------------------------
'#�v���V�[�W�����@:txtApprovalYmdT_DblClick
'#�����@�@�@�@�@�@:�_�u���N���b�N��
'#-----------------------------------------------------------------
Private Sub txtApprovalYmdT_DblClick(Cancel As Integer)

'-----------------------------------
'���J�����_�[�\��
'-----------------------------------
    Call DoCmd.OpenForm("frmCalender", , , , , acDialog)

    If (Len(gstrCalender) > 0) Then
        txtApprovalYmdT.Value = gstrCalender
    End If

End Sub