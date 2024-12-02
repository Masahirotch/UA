Attribute VB_Name = "mdlOutput"
Option Compare Database
Option Explicit

'#-----------------------------------------------------------------
'#プロシージャ名　:funOutPutProc1
'#説明　　　　　　:家族カード申請書出力
'#引数　　　　　　:pstrTemp: テンプレートファイルパス
'#　　　　　　　　:pstrOut : 出力先パス
'#戻り値　　　　　:Ture:正常/False:異常
'#-----------------------------------------------------------------
Public Function funOutPutProc1( _
                    ByVal pstrTemp As String, _
                    ByVal pstrOut As String _
                ) As Boolean
Dim blnSyokai       As Boolean      '初回フラグ
Dim lngRow          As Long         '行
Dim lngCopyRow      As Long         'コピー開始行
Dim strFileNm       As String       'Excelファイル名
Dim strSyainNoN     As String       '社員番号(新)
Dim strSyainNoO     As String       '社員番号(元)
Dim varRowData()    As Variant
Dim adoRs           As New ADODB.Recordset
Dim ExApp           As Excel.Application
Dim ExBook          As Excel.Workbook
Dim ExSheet         As Excel.Worksheet

On Error GoTo Err_Exit

    funOutPutProc1 = False

'------------------------------
'■家族カードデータ取得
'------------------------------
    '▼レコードセット作成
    Call adoRs.Open(funMakeSQL1(), goADOCon, adOpenStatic, adLockReadOnly)

    If (adoRs.EOF) Then
        '▼データがない場合
        Call MsgBox(mdlCommon.funGetMsgW("010"), vbExclamation, Sys_Title)
        Call mdlCommon.subCloseRecordset(adoRs)
        Exit Function
    End If

'------------------------------
'■Excelオブジェクトの準備
'------------------------------
    Set ExApp = CreateObject("Excel.Application")
    ExApp.Application.DisplayAlerts = False
    Call subExcelObjet(pstrTemp & EXCEL_TEMP_FILE1, ExApp, ExBook, ExSheet, True)

'------------------------------
'■Excel出力
'------------------------------
    strSyainNoO = ""
    lngCopyRow = 50
    blnSyokai = True
    Erase varRowData:    ReDim varRowData(EXCEL_ROW_DATA1, EXCEL_COL_DATA1)

    Do Until adoRs.EOF
        '▽社員番号
        strSyainNoN = adoRs.Fields("EmployeeNo").Value

        If (strSyainNoO <> strSyainNoN) Then
            '▼社員番号が変わった時

            If (blnSyokai = False) Then
                '▼初回以外の場合の処理

                '▼データ一括出力
                Call subDataOut(ExSheet, varRowData, lngCopyRow - 22)

                '▼Excel保存
                strFileNm = "家族カード申請書(" & strSyainNoO & ").xls"
                Call sudExcelSave(ExApp, ExSheet, pstrOut & strFileNm, strFileNm)

                '▼オブジェクト生成
                Call subExcelObjet(pstrTemp & EXCEL_TEMP_FILE1, ExApp, ExBook, ExSheet, False)
            End If

            With ExSheet
                '▽シート名
                .Name = strSyainNoN
                '▽所属
                .Cells(4, 2) = "所属　：　" & "" & adoRs.Fields("所属略称名").Value
                '▽社員番号
                .Cells(4, 5) = "社員番号　：　" & adoRs.Fields("EmployeeNo").Value
                '▽社員番号
                .Cells(5, 2) = "氏名　：　" & adoRs.Fields("氏名").Value
            End With

            '▽行開始位置
            lngRow = 0
            lngCopyRow = 50
            strSyainNoO = strSyainNoN
            blnSyokai = False
            Erase varRowData:    ReDim varRowData(EXCEL_ROW_DATA1, EXCEL_COL_DATA1)
        End If

        If (lngRow >= 10) Then
            '▼明細件数が１０件を超えた場合、改行処理
            
            '▼データ一括出力
            Call subDataOut(ExSheet, varRowData, lngCopyRow - 22)

            '▼セルコピー
            With ExSheet
                .Activate
                .Select
                .Rows("1:49").Copy
                .Range("A" & CStr(lngCopyRow)).Select
                .Paste
            End With

            '▼変数初期化
            lngRow = 0
            lngCopyRow = lngCopyRow + 49
            Erase varRowData:    ReDim varRowData(EXCEL_ROW_DATA1, EXCEL_COL_DATA1)
        End If

        '▽家族氏名
        varRowData(lngRow, 0) = "" & adoRs.Fields("FamilyNm").Value
        '▽続柄
        varRowData(lngRow, 2) = "" & adoRs.Fields("RelationShipNm").Value
        '▽フリガナ
        varRowData(lngRow, 3) = "" & adoRs.Fields("Furigana").Value

        lngRow = lngRow + 1
        adoRs.MoveNext
    Loop

    '▼データ一括出力
    Call subDataOut(ExSheet, varRowData, lngCopyRow - 22)

    '▼Excel保存
    strFileNm = "家族カード申請書(" & strSyainNoO & ").xls"
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
'#プロシージャ名　:funMakeSQL1
'#説明　　　　　　:ＳＱＬ文作成(家族カード申請書出力)
'#引数　　　　　　:なし
'#戻り値　　　　　:ＳＱＬ文
'#-----------------------------------------------------------------
Private Function funMakeSQL1() As String
Dim strSQL          As String       'ＳＱＬ文

    '▼ＳＱＬ編集
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "  A.*, "
    strSQL = strSQL & "  B.RelationShipNm, "
    strSQL = strSQL & "  E.氏名, "
'2012/03/08 No8 Upd-Start
'    strSQL = strSQL & "  E.所属略称名 "
    strSQL = strSQL & "  E.業務名 所属略称名 "
'2012/03/08 No8 Upd-End
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "  T_Family A "
    strSQL = strSQL & "LEFT JOIN "
    strSQL = strSQL & " (SELECT * FROM M_RelationShip WHERE DelFlg = '" & CST_DEL_FLG & "') B "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  A.RelationShipCd = B.RelationShipCd "
    strSQL = strSQL & "LEFT JOIN "
    strSQL = strSQL & gstrLinkDB & "社員固定情報 E "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  A.EmployeeNo = E.社員番号 "
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
'#プロシージャ名　:subExcelObjet
'#説明　　　　　　:Excelオブジェクト作成
'#引数　　　　　　:pstrTemp: テンプレートファイルパス
'#　　　　　　　　:pExApp  : Excelアプリケーションオブジェクト
'#　　　　　　　　:pExBook : ワークブックオブジェクト
'#　　　　　　　　:pExSheet: ワークシートオブジェクト
'#　　　　　　　　:pblnFlg : True:ファイルオープン
'#戻り値　　　　　:なし
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
'■Excelオブジェクトの準備
'-----------------------------------
    If (pblnFlg) Then Set pExBook = pExApp.Application.Workbooks.Open(pstrTemp)
    pExBook.Activate
    pExBook.Sheets(Array("Temp")).Copy

    Set pExSheet = pExApp.Sheets(1)

End Sub

'#-----------------------------------------------------------------
'#プロシージャ名　:subDataOut
'#説明　　　　　　:Excelにデータ出力
'#引数　　　　　　:pExSheet: ワークシートオブジェクト
'#　　　　　　　　:pData() : 出力データ
'#戻り値　　　　　:なし
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
'#プロシージャ名　:sudExcelSave
'#説明　　　　　　:Excelデータ保存
'#引数　　　　　　:pExApp  : Excelアプリケーションオブジェクト
'#    　　　　　　:pExSheet  : ワークシートオブジェクト
'#　　　　　　　　:pstrPath  : 保存するファイルパス
'#　　　　　　　　:pstrFileNm: ファイル名
'#戻り値　　　　　:なし
'#-----------------------------------------------------------------
Private Sub sudExcelSave( _
                ByRef pExApp As Object, _
                ByRef pExSheet As Object, _
                ByVal pstrPath As String, _
                ByVal pstrFileNm As String _
            )
Dim ExBookWk        As Excel.Workbook

    With pExSheet
        '▼Excel保存
        .Activate
        .Cells(1, 1).Select
'2012/11/02 Windows7対応 Upd-Start
'        .SaveAs Filename:=pstrPath, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
        If (pExApp.Version < Excel_Version_2007) Then
            .SaveAs Filename:=pstrPath, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
        Else
            .SaveAs Filename:=pstrPath, FileFormat:=Excel_FileFormat_xlExcel8, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
        End If
'2012/11/02 Windows7対応 Upd-End

        If (gtApply(0).PrintFlg = CST_CHK_ON) Then
            '▼印刷
            .PrintOut
        End If
    End With


    '▼保存したファイルを閉じる
    For Each ExBookWk In pExApp.Workbooks
        If (ExBookWk.Name = pstrFileNm) Then
            ExBookWk.Close
            Exit For
        End If
    Next

End Sub

'#-----------------------------------------------------------------
'#プロシージャ名　:funOutPutProc2
'#説明　　　　　　:家族カード申請書出力
'#引数　　　　　　:pstrTemp: テンプレートファイルパス
'#　　　　　　　　:pstrOut : 出力先パス
'#戻り値　　　　　:Ture:正常/False:異常
'#-----------------------------------------------------------------
Public Function funOutPutProc2( _
                    ByVal pstrTemp As String, _
                    ByVal pstrOut As String _
                ) As Boolean
Dim lngRow          As Long         '行
Dim lngCount        As Long         'ワークカウンタ
Dim varRowData()    As Variant
Dim adoRs           As New ADODB.Recordset
Dim ExApp           As Excel.Application
Dim ExBook          As Excel.Workbook
Dim ExSheet         As Excel.Worksheet
Dim ExBookWk        As Excel.Workbook

On Error GoTo Err_Exit

    funOutPutProc2 = False

'------------------------------
'■家族カードデータ取得
'------------------------------
    '▼レコードセット作成
    Call adoRs.Open(funMakeSQL2(), goADOCon, adOpenStatic, adLockReadOnly)

    If (adoRs.EOF) Then
        '▼データがない場合
        Call MsgBox(mdlCommon.funGetMsgW("010"), vbExclamation, Sys_Title)
        Call mdlCommon.subCloseRecordset(adoRs)
        Exit Function
    End If

'------------------------------
'■Excelオブジェクトの準備
'------------------------------
    Set ExApp = CreateObject("Excel.Application")
    ExApp.Application.DisplayAlerts = False
    Call subExcelObjet(pstrTemp & EXCEL_TEMP_FILE2_1, ExApp, ExBook, ExSheet, True)

'------------------------------
'■Excel出力
'------------------------------
    Erase varRowData:    ReDim varRowData(EXCEL_COL_DATA2)

    '▽行開始位置
    lngRow = EXCEL_ROW_START2:  lngCount = 0
    Erase gtFamily:

    Do Until adoRs.EOF
        If ("" & adoRs.Fields("IssueYmd").Value = "") Then
            lngCount = lngCount + 1
            ReDim Preserve gtFamily(1 To lngCount)
            '▽社員番号
            gtFamily(lngCount).EmployeeNo = adoRs.Fields("EmployeeNo").Value
            '▽家族番号
            gtFamily(lngCount).FamilyNo = adoRs.Fields("FamilyNo").Value
            '▽枝番
            gtFamily(lngCount).SeqNo = adoRs.Fields("SeqNo").Value
        End If

        '▽社員番号
        varRowData(0) = "'" & adoRs.Fields("EmployeeNo").Value
        '▽社員名
        varRowData(1) = "" & adoRs.Fields("氏名").Value
        '▽フリガナ
        varRowData(2) = ""
        '▽所属コード
        varRowData(3) = ""
        '▽所属名略称
        varRowData(4) = "" & adoRs.Fields("DepartmentNm").Value
        '▽ご家族名
        varRowData(5) = "" & adoRs.Fields("FamilyNm").Value
        '▽続柄
        varRowData(6) = "" & adoRs.Fields("RelationShipNm").Value
        '▽(ﾌﾘｶﾞﾅ)
        varRowData(7) = "" & adoRs.Fields("Furigana").Value
'2013/01/28 Upd-Start
        varRowData(8) = "" & adoRs.Fields("PeriodValidNm").Value
'2013/01/28 Upd-End

        '▼行情報一括出力
        With ExSheet
            .Range(.Cells(lngRow, 1), .Cells(lngRow, EXCEL_COL_DATA2 + 1)).Value = varRowData
        End With

        lngRow = lngRow + 1
        adoRs.MoveNext
    Loop

    With ExSheet
        '▼Excel保存
        .Activate
        .Cells(1, 1).Select
'2012/11/02 Windows7対応 Upd-Start
'        .SaveAs Filename:=pstrOut, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
        If (ExApp.Version < Excel_Version_2007) Then
            .SaveAs Filename:=pstrOut, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
        Else
            .SaveAs Filename:=pstrOut, FileFormat:=Excel_FileFormat_xlExcel8, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
        End If
'2012/11/02 Windows7対応 Upd-End
    End With

    '▼保存したファイルを閉じる
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
'#プロシージャ名　:funMakeSQL2
'#説明　　　　　　:ＳＱＬ文作成(夢ぷりんと用データ出力)
'#引数　　　　　　:なし
'#戻り値　　　　　:ＳＱＬ文
'#-----------------------------------------------------------------
Private Function funMakeSQL2() As String
Dim strSQL          As String       'ＳＱＬ文

    '▼ＳＱＬ編集
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "  A.*, "
    strSQL = strSQL & "  B.RelationShipNm, "
'2013/01/28 Upd-Start
    strSQL = strSQL & "  C.PeriodValidNm, "
'2013/01/28 Upd-End
    strSQL = strSQL & "  E.氏名 "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "  T_Family A "
    strSQL = strSQL & "LEFT JOIN "
    strSQL = strSQL & " (SELECT RelationShipCd, RelationShipNm FROM M_RelationShip WHERE DelFlg = '" & CST_DEL_FLG & "') B "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  A.RelationShipCd = B.RelationShipCd "
    strSQL = strSQL & "LEFT JOIN "
    strSQL = strSQL & gstrLinkDB & "社員固定情報 E "
    strSQL = strSQL & "ON "
    strSQL = strSQL & "  A.EmployeeNo = E.社員番号 "
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



