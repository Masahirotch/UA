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
'#プロシージャ名　:cmdEnd_Click
'#説明　　　　　　:終了ボタンクリック時
'#-----------------------------------------------------------------
Private Sub cmdEnd_Click()

    Call DoCmd.Close(acForm, Me.Name)
    Call DoCmd.OpenForm("frmMenu")

End Sub

'#-----------------------------------------------------------------
'#プロシージャ名　:cmdOutput1_Click
'#説明　　　　　　:家族カード申請書出力ボタンクリック時
'#-----------------------------------------------------------------
Private Sub cmdOutput1_Click()

    Call DoCmd.Close(acForm, Me.Name)
    Call DoCmd.OpenForm("frmOutput1")

End Sub

'#-----------------------------------------------------------------
'#プロシージャ名　:cmdOutput1_Click
'#説明　　　　　　:夢ぷりんと用データ出力ボタンクリック時
'#-----------------------------------------------------------------
Private Sub cmdOutput2_Click()

    Call DoCmd.Close(acForm, Me.Name)
    Call DoCmd.OpenForm("frmOutput2")

End Sub

'#-----------------------------------------------------------------
'#プロシージャ名　:Form_Load
'#説明　　　　　　:画面起動時
'#-----------------------------------------------------------------
Private Sub Form_Load()

'-----------------------------------
'■初期処理
'-----------------------------------
    '▼画面サイズ設定
    Call mdlCommon.subGamenSizeCM(12, 10)

    '▼フォームタイトル
    Me.Caption = "家族カード [" & gstrServer & " : " & gstrDB & "]"

End Sub
