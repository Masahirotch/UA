VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenuOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Database

'#-----------------------------------------------------------------
'#�v���V�[�W�����@:cmdEnd_Click
'#�����@�@�@�@�@�@:�I���{�^���N���b�N��
'#-----------------------------------------------------------------
Private Sub cmdEnd_Click()

    Call DoCmd.Close(acForm, Me.Name)
    Call DoCmd.OpenForm("frmMenu")

End Sub

'#-----------------------------------------------------------------
'#�v���V�[�W�����@:cmdOutput1_Click
'#�����@�@�@�@�@�@:�Ƒ��J�[�h�\�����o�̓{�^���N���b�N��
'#-----------------------------------------------------------------
Private Sub cmdOutput1_Click()

    Call DoCmd.Close(acForm, Me.Name)
    Call DoCmd.OpenForm("frmOutput1")

End Sub

'#-----------------------------------------------------------------
'#�v���V�[�W�����@:cmdOutput1_Click
'#�����@�@�@�@�@�@:���Ղ��Ɨp�f�[�^�o�̓{�^���N���b�N��
'#-----------------------------------------------------------------
Private Sub cmdOutput2_Click()

    Call DoCmd.Close(acForm, Me.Name)
    Call DoCmd.OpenForm("frmOutput2")

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
    Call mdlCommon.subGamenSizeCM(12, 10)

    '���t�H�[���^�C�g��
    Me.Caption = "�Ƒ��J�[�h [" & gstrServer & " : " & gstrDB & "]"

End Sub