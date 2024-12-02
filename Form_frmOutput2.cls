VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmOutput2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Database

'#-----------------------------------------------------------------
'#プロシージャ名　:funCheckInput
'#説明　　　　　　:入力チェック
'#引数　　　　　　:なし
'#戻り値　　　　　:True:正常/False:異常
'#-----------------------------------------------------------------
Private Function funCheckInput(ByRef pstrPath As String) As Boolean

    funCheckInput = False

    Erase gtApply:  ReDim gtApply(0)

    With gtApply(0)
'-----------------------------------
'■承認日(From)
'-----------------------------------
        .ApprovalYmdF = "" & txtApprovalYmdF.Value
        If (Len(.ApprovalYmdF) = 0) Then
            Call MsgBox(mdlCommon.funGetMsgW("015"), vbExclamation, Sys_Title)
            Exit Function
        End If

'-----------------------------------
'■承認日(To)
'-----------------------------------
        .ApprovalYmdT = "" & txtApprovalYmdT.Value
        If (Len(.ApprovalYmdT) = 0) Then .ApprovalYmdT = .ApprovalYmdF

'-----------------------------------
'■関連チェック
'-----------------------------------
        If (.ApprovalYmdF > .ApprovalYmdT) Then
            Call MsgBox(mdlCommon.funGetMsgW("016"), vbExclamation, Sys_Title)
            Exit Function
        End If
    End With

'-----------------------------------
'■EXCELテンプレートパスチェック
'-----------------------------------
    pstrPath = CurrentProject.Path
    If (Right$(pstrPath, 1) <> "\") Then pstrPath = pstrPath & "\"
    pstrPath = pstrPath & EXCEL_TEMP_PATH
    If (Right$(pstrPath, 1) <> "\") Then pstrPath = pstrPath & "\"

    '▼EXCELテンプレート存在チェック
    If (Len(Dir(pstrPath & EXCEL_TEMP_FILE2_1)) = 0) Then
        '▼存在しない場合
        Call MsgBox(mdlCommon.funGetMsgW("018"), vbExclamation, Sys_Title)
        Exit Function
    End If

    funCheckInput = True

End Function

'#-----------------------------------------------------------------
'#プロシージャ名　:cmdCancel_Click
'#説明　　　　　　:キャンセルボタンクリック時
'#-----------------------------------------------------------------
Private Sub cmdCancel_Click()

'-----------------------------------
'■画面を閉じる
'-----------------------------------
    Call DoCmd.Close(acForm, Me.Name)
    Call DoCmd.OpenForm("frmMenuOutput")

End Sub

'#-----------------------------------------------------------------
'#プロシージャ名　:cmdOutput_Click
'#説明　　　　　　:出力ボタンクリック時
'#-----------------------------------------------------------------
Private Sub cmdOutput_Click()
Dim blnRtn          As Boolean      '戻り値
Dim intRtn          As Integer      '戻り値
Dim strTemp         As String       'テンプレートパス
Dim strOut          As String       '出力先パス

'-----------------------------------
'■入力チェック
'-----------------------------------
    blnRtn = funCheckInput(strTemp)
    If (blnRtn = False) Then Exit Sub

'-----------------------------------
'■確認メッセージ
'-----------------------------------
    intRtn = MsgBox(mdlCommon.funGetMsgI("008"), vbInformation + vbYesNo, Sys_Title)
    If (intRtn = vbNo) Then Exit Sub

'-----------------------------------
'■フォルダ参照ダイアログ表示
'-----------------------------------
    strOut = mdlCommon.funFileSaveDialog2(EXCEL_TEMP_FILE2_2)
    If (Len(strOut) = 0) Then Exit Sub

    '▼アイコン（待機）
    Screen.MousePointer = CST_MP_H
DoEvents

'-----------------------------------
'■夢ぷりんと用データ出力
'-----------------------------------
    blnRtn = mdlOutput.funOutPutProc2(strTemp, strOut)
    If (blnRtn = False) Then GoTo End_Proc

    If (gtFamily(1).EmployeeNo = "") Then
    Else
        '▼トランザクション開始
        Call goADOCon.BeginTrans

'-----------------------------------
'■家族カード更新
'-----------------------------------
        blnRtn = mdlTFamily.funUpdFamily5()
        If (blnRtn = False) Then
            '▼ロールバック
            Call goADOCon.RollbackTrans
    
            GoTo End_Proc
        Else
            '▼コミット
            Call goADOCon.CommitTrans
        End If
    End If

    Call MsgBox(mdlCommon.funGetMsgI("009"), vbInformation, Sys_Title)

End_Proc:
    '▼アイコン（デフォルト）
    Screen.MousePointer = CST_MP_D

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
    Call mdlCommon.subGamenSizeCM(12, 9)

End Sub

'#-----------------------------------------------------------------
'#プロシージャ名　:txtApprovalYmdF_DblClick
'#説明　　　　　　:ダブルクリック時
'#-----------------------------------------------------------------
Private Sub txtApprovalYmdF_DblClick(Cancel As Integer)

'-----------------------------------
'■カレンダー表示
'-----------------------------------
    Call DoCmd.OpenForm("frmCalender", , , , , acDialog)

    If (Len(gstrCalender) > 0) Then
        txtApprovalYmdF.Value = gstrCalender
    End If

End Sub

'#-----------------------------------------------------------------
'#プロシージャ名　:txtApprovalYmdT_DblClick
'#説明　　　　　　:ダブルクリック時
'#-----------------------------------------------------------------
Private Sub txtApprovalYmdT_DblClick(Cancel As Integer)

'-----------------------------------
'■カレンダー表示
'-----------------------------------
    Call DoCmd.OpenForm("frmCalender", , , , , acDialog)

    If (Len(gstrCalender) > 0) Then
        txtApprovalYmdT.Value = gstrCalender
    End If

End Sub
