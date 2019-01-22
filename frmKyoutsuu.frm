VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKyoutsuu 
   Caption         =   "共通項目"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9630
   OleObjectBlob   =   "frmKyoutsuu.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmKyoutsuu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intConfirmMsg As Integer

'初期表示
Public Sub UserForm_Initialize()
    
    On Error GoTo Error
    
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    'フォームのタイトル設定
    If FleetTypeFlg = 1 Then
        frmKyoutsuu.Caption = frmKyoutsuu.Caption & "（フリート契約）"
    Else
        frmKyoutsuu.Caption = frmKyoutsuu.Caption & "（ノンフリート明細付契約）"
    End If
    
    '明細入力シート取得
    Dim wsMeisai As Worksheet
    If FleetTypeFlg = 1 Then
        Call subSetSheet(1, wsMeisai)       'シートオブジェクト(明細入力)
    Else
        Call subSetSheet(17, wsMeisai)       'シートオブジェクト(明細入力（ノンフリート）)（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
    End If
    
    'コード値シート取得
    Dim wsCode As Worksheet
    If FleetTypeFlg = 1 Then
        Call subSetSheet(4, wsCode)         'シートオブジェクト(別紙　コード値)
    Else
        Call subSetSheet(16, wsCode)        'シートオブジェクト(別紙　コード値（ノンフリート）)（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
    End If
    
    '共通項目シート取得
    Dim wsKyoutsuU As Worksheet
    Call subSetSheet(2, wsKyoutsuU)         'シートオブジェクト(別紙　共通項目)

    '総付保台数
    If blnFleetBtnFlg Then
        'デフォルトの行数(10行)を設定
        txtSouFuhoDaisu.Value = "10"
        blnFleetBtnFlg = False
    ElseIf blnNonFleetBtnFlg Then
        'デフォルトの行数(3行)を設定（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        txtSouFuhoDaisu.Value = "3"
        blnNonFleetBtnFlg = False
    Else
        '明細入力シートの総付保台数を設定
        Dim objSouhuho As Object
        Set objSouhuho = wsMeisai.OLEObjects("txtSouhuho").Object
        txtSouFuhoDaisu.Value = Left(objSouhuho.Value, Len(objSouhuho.Value) - 2)
    End If
            
    '受付区分
    With cmbUketsukekbn
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("B2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("B2").Column + 1).End(xlUp)).Value
        .ColumnWidths = "-1;0"
    End With
    
    '保険種類
    With cmbHokenSyurui
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("J2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("J2").Column + 1).End(xlUp)).Value
        .ColumnWidths = "-1;0"
    End With

    'フリート区分
    Dim intRows As Integer
    intRows = wsCode.Cells(wsCode.Rows.Count, wsCode.Range("N2").Column).End(xlUp).Row

    Dim strArray() As Variant

    If FleetTypeFlg = 1 Then    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        'フリート
        Dim intListNo As Integer
        Dim intRow As Integer
        Dim strListName As String
        Dim strNonFleetName As String
        ReDim strArray(intRows - 3, 1) As Variant

        strNonFleetName = "ノンフリート"
        intRow = 0

        For intListNo = 0 To intRows - 2
            strListName = wsCode.Cells(intListNo + 2, wsCode.Range("N2").Column).Value
            If strListName <> strNonFleetName Then
                strArray(intRow, 0) = strListName
                strArray(intRow, 1) = wsCode.Cells(intListNo + 2, wsCode.Range("N2").Column + 1).Value
                intRow = intRow + 1
            End If
        Next intListNo
    Else
        'ノンフリート
        ReDim strArray(0, 1) As Variant
        strArray(0, 0) = wsCode.Cells(3, wsCode.Range("N2").Column).Value
        strArray(0, 1) = wsCode.Cells(3, wsCode.Range("O2").Column).Value
    End If

    With cmbFreetkbn
        .AddItem ""
        .List = strArray
        .ColumnWidths = "-1;0"
    End With

    'フリート区分を非活性に変更　'2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg <> 1 Then
        cmbFreetkbn.ListIndex = 0
        cmbFreetkbn.Enabled = False
    End If

    '払込方法
    With cmbHaraiHouhou
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("AX2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("AX2").Column + 1).End(xlUp)).Value
        .ColumnWidths = "-1;0"
    End With

    '被保険者
    optKojin = True

    'ﾉﾝﾌﾘｰﾄ多数割引（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
    With cmbNonfleetTawari
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("AP2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("AP2").Column + 1).End(xlUp)).Value
        .ColumnWidths = "-1;0"
    End With

    '画面下部のフレーム
    If FleetTypeFlg <> 1 Then
        'ノンフリート（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        FrameFleet.Visible = False
        FrameNonFleet.Visible = True
    Else
        'フリート契約
        FrameFleet.Visible = True
        FrameNonFleet.Visible = False
    End If


    '保存情報がある場合反映する
    Dim strSaveContent As String
    If fncFormRef(1, strSaveContent) Then

    Else

        Dim varSaveContent As Variant
        varSaveContent = Split(strSaveContent, "/")

        Dim i As Integer

        '受付区分
        For i = 0 To cmbUketsukekbn.ListCount - 1
            If cmbUketsukekbn.List(i, 1) = varSaveContent(0) Then
                cmbUketsukekbn.ListIndex = i
                Exit For
            End If
        Next

        '被保険者
        If varSaveContent(1) = "1" Then
            optKojin = True
            optHoujin = False
        Else
            optKojin = False
            optHoujin = True
        End If

        '保険種類
        For i = 0 To cmbHokenSyurui.ListCount - 1
            If cmbHokenSyurui.List(i, 1) = varSaveContent(2) Then
                cmbHokenSyurui.ListIndex = i
                Exit For
            End If
        Next

        'フリート区分
        For i = 0 To cmbFreetkbn.ListCount - 1
            If cmbFreetkbn.List(i, 1) = varSaveContent(3) Then
                cmbFreetkbn.ListIndex = i
                Exit For
            End If
        Next

        '保険始期日
'        txtHokenStart_Nen = Format(Val(Left(varSaveContent(4), 4)) - 1988, "00")
        txtHokenStart_Nen = Format(Val(Left(varSaveContent(4), 4)) - 2000, "00") '新元号対応
        txtHokenStart_Tsuki = Mid(varSaveContent(4), 5, 2)
        txtHokenStart_Hi = Right(varSaveContent(4), 2)

        '払込方法
        For i = 0 To cmbHaraiHouhou.ListCount - 1
            If cmbHaraiHouhou.List(i, 1) = varSaveContent(7) Then
                cmbHaraiHouhou.ListIndex = i
                Exit For
            End If
        Next

        '優良割引
        txtYuuryowari = varSaveContent(8)

        '第一種デメ割増
        txtFirstDeme = varSaveContent(9)

        'フリート多数割引
        chkFreetTasuu = IIf(varSaveContent(10) = "2 ", True, False)

        'フリートコード
        txtFreetCode = varSaveContent(11)

        'ﾉﾝﾌﾘｰﾄ多数割引（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        For i = 0 To cmbNonfleetTawari.ListCount - 1
            If cmbNonfleetTawari.List(i, 1) = varSaveContent(12) Then
                cmbNonfleetTawari.ListIndex = i
                Exit For
            End If
        Next

        '団体割増引（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        txtDantaiWarimashibiki = varSaveContent(13)

    End If

    Set wsMeisai = Nothing
    Set wsCode = Nothing
    Set wsKyoutsuU = Nothing
    Set objSouhuho = Nothing

    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "UserForm_Initialize" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"

End Sub

'「次へ」ボタン押下
Private Sub btnNext_Click()
    Dim strErrMsg As String

    strErrMsg = ""

    On Error GoTo Error

    '入力チェック
    Call fncEntryCheckKyotsu(strErrMsg)

    'エラー判定
    If strErrMsg <> "" Then
        txtErrMsg = strErrMsg

        txtErrMsg.SetFocus
        txtErrMsg.SelStart = 0
        Exit Sub
    End If

    '入力情報取得
    Dim strSaveData As String
    Call subSaveData(strSaveData)

    '入力情報保存
    Dim blnResult As Boolean
    blnResult = fncFormSave(1, strSaveData)

    'シートの保護の解除
    Call subMeisaiUnProtect

    '明細入力画面ヘッダ設定
    Dim wsMeisai As Worksheet
    'フリート
    If FleetTypeFlg = 1 Then
        Call subSetSheet(1, wsMeisai)       'シートオブジェクト(明細入力)
    'ノンフリート
    Else
        Call subSetSheet(17, wsMeisai)      'シートオブジェクト(明細入力（ノンフリート）)　2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    End If

    '保険期間
'    wsMeisai.Range("B3") = "　保険期間　　：平成" & Format(Trim(txtHokenStart_Nen), "00") & "年" & Format(Trim(txtHokenStart_Tsuki), "00") & "月" & Format(Trim(txtHokenStart_Hi), "00") & "日から1年間"
    wsMeisai.Range("B3") = "　保険期間　　：20" & Format(Trim(txtHokenStart_Nen), "00") & "年" & Format(Trim(txtHokenStart_Tsuki), "00") & "月" & Format(Trim(txtHokenStart_Hi), "00") & "日から1年間"
    '受付区分
    wsMeisai.Range("E3") = "　受付区分　　：" & cmbUketsukekbn
    '被保険者
    wsMeisai.Range("G3") = "　被保険者　　　　　　：" & IIf(optKojin, "個人", "法人")
    '保険種類
    wsMeisai.Range("B4") = "　保険種類　　：" & cmbHokenSyurui
    'フリート区分
    wsMeisai.Range("E4") = "　フリート区分：" & cmbFreetkbn

    If FleetTypeFlg = 1 Then  'フリート
        '全車両一括付保特約
        wsMeisai.Range("G4") = "　全車両一括付保特約　：" & IIf(cmbFreetkbn = "全車両一括" Or cmbFreetkbn = "全車両連結合算", "有り", "無し")
    Else
        'ノンフリート多数割引（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        wsMeisai.Range("G4") = "　ノンフリート多数割引：" & cmbNonfleetTawari
    End If

    '払込方法
    wsMeisai.Range("B5") = "　払込方法　　：" & cmbHaraiHouhou

    If FleetTypeFlg = 1 Then  'フリート
        '優良割引
        wsMeisai.Range("E5") = "　優良割引　　：" & IIf(Trim(txtYuuryowari) = "", "", Trim(txtYuuryowari) & "%")
    Else                      'ノンフリート
        '団体割増引（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        wsMeisai.Range("E5") = "　団体割増引　：" & IIf(Trim(txtDantaiWarimashibiki) = "", "", Trim(txtDantaiWarimashibiki) & "%")
    End If

    If FleetTypeFlg = 1 Then  'フリート
        '第一種デメ割増
        wsMeisai.Range("G5") = "　第一種デメ割増  　　：" & IIf(Trim(txtFirstDeme) = "", "", Trim(txtFirstDeme) & "%")
        'ﾌﾘｰﾄ多数割引
        wsMeisai.Range("B6") = "　ﾌﾘｰﾄ多数割引：" & IIf(chkFreetTasuu, "有り", "無し")
        'ﾌﾘｰﾄｺｰﾄﾞ
        wsMeisai.Range("E6") = "　ﾌﾘｰﾄｺｰﾄﾞ　　：" & txtFreetCode
    Else
        wsMeisai.Range("G5") = "　"
        wsMeisai.Range("B6") = "　"
        wsMeisai.Range("E6") = "　"
    End If

    Call subBookUnProtect           'ブックの保護を解除
    Call subSheetVisible(True)      'シート・ブックの表示
    Call subBookProtect             'ブックの保護

    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        ThisWorkbook.Worksheets("明細入力").Activate
    Else
        ThisWorkbook.Worksheets("明細入力（ノンフリート）").Activate
    End If

    Dim intSouhuho As Integer
    Dim objSouhuho As Object
    Set objSouhuho = wsMeisai.OLEObjects("txtSouhuho").Object
    intSouhuho = Val(Left(objSouhuho.Value, Len(objSouhuho.Value) - 2))

    '総付保台数分、明細行を追加・削除
    If Val(Trim(txtSouFuhoDaisu)) > intSouhuho Then
        Call subMeisaiAdd(Val(Trim(txtSouFuhoDaisu)) - intSouhuho, "2")
    ElseIf Val(Trim(txtSouFuhoDaisu)) < Val(intSouhuho) Then
        Call subMeisaiDel(intSouhuho - Val(Trim(txtSouFuhoDaisu)), "2")
    End If

    '明細入力画面のエラー用リスト初期化
    wsMeisai.OLEObjects("txtErrMsg").Object.Value = ""
    wsMeisai.OLEObjects("txtErrMsg").Activate
    wsMeisai.Range("A1").Activate

    Set wsMeisai = Nothing
    Set objSouhuho = Nothing

    Call subMeisaiProtect       'シートの保護

    '20190110対応
    MeisaiBackFlg = 1

    Unload Me

    On Error GoTo 0

    Exit Sub
Error:
    MsgBox "エラー番号:" & Err.Number & vbCrLf & _
           "エラーの種類:" & Err.Description, vbExclamation, "btnNext_Click"

End Sub


'入力情報取得
Private Sub subSaveData(ByRef strSaveData As String)

'    '証券番号
'    strSaveData = strSaveData & Trim(txtShokenNo) & "/"

    '受付区分
    If cmbUketsukekbn.ListIndex > -1 Then
        strSaveData = strSaveData & cmbUketsukekbn.List(cmbUketsukekbn.ListIndex, 1) & "/"
    Else
        strSaveData = strSaveData & "" & "/"
    End If
    '被保険者_個人法人区分
    strSaveData = strSaveData & IIf(optKojin, "1", "2") & "/"
    '保険種類
    If cmbHokenSyurui.ListIndex > -1 Then
        strSaveData = strSaveData & cmbHokenSyurui.List(cmbHokenSyurui.ListIndex, 1) & "/"
    Else
        strSaveData = strSaveData & "" & "/"
    End If
    'フリート・ノンフリート区分
    If cmbFreetkbn.ListIndex > -1 Then
        strSaveData = strSaveData & cmbFreetkbn.List(cmbFreetkbn.ListIndex, 1) & "/"
    Else
        strSaveData = strSaveData & "" & "/"
    End If
    '保険始期日年月日
'    strSaveData = strSaveData & Format(Val(Trim(txtHokenStart_Nen)) + 1988, "0000") & Format(Trim(txtHokenStart_Tsuki), "00") & Format(Trim(txtHokenStart_Hi), "00") & "/"
    strSaveData = strSaveData & Format(Val(Trim(txtHokenStart_Nen)) + 2000, "0000") & Format(Trim(txtHokenStart_Tsuki), "00") & Format(Trim(txtHokenStart_Hi), "00") & "/"
    strSaveData = strSaveData & "1" & "/"
    strSaveData = strSaveData & "0" & "/"
    If cmbHaraiHouhou.ListIndex > -1 Then
        strSaveData = strSaveData & cmbHaraiHouhou.List(cmbHaraiHouhou.ListIndex, 1) & "/"
    Else
        strSaveData = strSaveData & "" & "/"
    End If
    '優良割引
    strSaveData = strSaveData & Trim(txtYuuryowari) & "/"
    '第一種デメ割増
    strSaveData = strSaveData & Trim(txtFirstDeme) & "/"
    'フリート多数割引
    strSaveData = strSaveData & IIf(chkFreetTasuu, "2 ", "") & "/"
    'フリートコード
    strSaveData = strSaveData & Trim(txtFreetCode) & "/"

    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    'ノンフリート多数割引
    If cmbNonfleetTawari.ListIndex > -1 Then
        strSaveData = strSaveData & cmbNonfleetTawari.List(cmbNonfleetTawari.ListIndex, 1) & "/"
    Else
        strSaveData = strSaveData & "" & "/"
    End If
    '団体割増引
    strSaveData = strSaveData & Trim(txtDantaiWarimashibiki) & "/"

End Sub


'「戻る」ボタン押下
Private Sub BtnBack_Click()
    Dim intMsgBox As Integer

    On Error GoTo Error

    intMsgBox = MsgBox("入力内容を削除してTOP画面に遷移します。" & vbCrLf & "よろしいですか?", vbYesNo, "確認ダイアログ")

    If intMsgBox = 6 Then

        '明細行の初期化
        Call subSaveDel

        Unload Me
        frmTop.Show vbModeless

    End If

    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "btnBack_Click" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"

End Sub


'「×」ボタン押下
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
        intConfirmMsg = MsgBox("ツールを終了します。" & vbCrLf & "よろしいですか?" & vbCrLf & "※入力内容は保存されません。", vbYesNo, "確認ダイアログ")
        If intConfirmMsg = 6 Then
            Cancel = False
            Call subAppClose
        Else
            Cancel = True
        End If
    End If

End Sub

Private Function fncEntryCheckKyotsu(ByRef strErrContent As String)
'関数名：fncEntryErrCheckKyotsu
'内容　：関連チェックのエラーチェックを行いエラーがある場合は,エラー内容を返す。
'引数　：strErrContent = :エラー内容

    Dim strErrChkMsg As String
    Dim strErrKoumoku As String
    Dim strHokenStartErrMsg As String
    Dim strHokenStart As String

    strHokenStartErrMsg = ""
    strErrKoumoku = ""
    strErrChkMsg = ""

    '・個別チェック

    '総付保台数
    strErrKoumoku = "・総付保台数" & vbCrLf
    strErrChkMsg = fncNeedCheck(Trim(txtSouFuhoDaisu))                      '必須チェック

    If strErrChkMsg = "" Then
        strErrChkMsg = fncNumCheck(Trim(txtSouFuhoDaisu))                   '数字チェック
        If strErrChkMsg = "" Then

            '数値チェック
            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
'            strErrChkMsg = fncNumRangeCheck(Trim(txtSouFuhoDaisu), 1, 999)
            If FleetTypeFlg = 1 Then
                strErrChkMsg = fncNumRangeCheck(Trim(txtSouFuhoDaisu), 1, 999)
            Else
                strErrChkMsg = fncNonfleetSoufuhodaisuCheck(Trim(txtSouFuhoDaisu), 9)
                If strErrChkMsg = "" Then
                    strErrChkMsg = fncNumRangeCheck(Trim(txtSouFuhoDaisu), 1, 9)
                End If
            End If
        End If
    End If
    If strErrChkMsg = "" Then
    Else
        strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
    End If

    '受付区分
    strErrChkMsg = ""

    strErrKoumoku = "・受付区分" & vbCrLf
    strErrChkMsg = fncNeedCheck(cmbUketsukekbn)                             '必須チェック

    If strErrChkMsg = "" Then
    Else
        strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
    End If

    '被保険者
    strErrChkMsg = ""

    strErrKoumoku = "・被保険者" & vbCrLf
    strErrChkMsg = fncNeedOptCheck(optKojin, optHoujin)                     '必須チェック

    If strErrChkMsg = "" Then
    Else
        strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
    End If

    '保険種類
    strErrChkMsg = ""

    strErrKoumoku = "・保険種類" & vbCrLf
    strErrChkMsg = fncNeedCheck(cmbHokenSyurui)                             '必須チェック

    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If strErrChkMsg = "" Then
        strErrKoumoku = "○保険種類エラー" & vbCrLf
        strErrChkMsg = fncHokenSyuruiCheck(cmbHokenSyurui.Value, optHoujin.Value) '関連チェック
    End If

    If strErrChkMsg = "" Then
    Else
        strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
    End If

    'フリート区分
    strErrChkMsg = ""

    strErrKoumoku = "・フリート区分" & vbCrLf
    strErrChkMsg = fncNeedCheck(cmbFreetkbn)                                '必須チェック

    If strErrChkMsg = "" Then
    Else
        strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
    End If

    '保険始期日_年
    strErrChkMsg = ""

    strErrKoumoku = "・保険始期日_年" & vbCrLf
    strErrChkMsg = fncNeedCheck(Trim(txtHokenStart_Nen))                      '必須チェック

    If strErrChkMsg = "" Then
        strErrChkMsg = fncNumCheck(Trim(txtHokenStart_Nen))                   '数字チェック
        If strErrChkMsg = "" Then
            strErrChkMsg = fncNumRangeCheck(Trim(txtHokenStart_Nen), 1, 99)   '数値チェック
        End If
    End If
    If strErrChkMsg = "" Then
    Else
        strHokenStartErrMsg = strHokenStartErrMsg & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
    End If

    '保険始期日_月
    strErrChkMsg = ""

    strErrKoumoku = "・保険始期日_月" & vbCrLf
    strErrChkMsg = fncNeedCheck(Trim(txtHokenStart_Tsuki))                      '必須チェック

    If strErrChkMsg = "" Then
        strErrChkMsg = fncNumCheck(Trim(txtHokenStart_Tsuki))                   '数字チェック
        If strErrChkMsg = "" Then
            strErrChkMsg = fncNumRangeCheck(Trim(txtHokenStart_Tsuki), 1, 12)   '数値チェック
        End If
    End If
    If strErrChkMsg = "" Then
    Else
        strHokenStartErrMsg = strHokenStartErrMsg & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
    End If

    '保険始期日_日
    strErrChkMsg = ""

    strErrKoumoku = "・保険始期日_日" & vbCrLf
    strErrChkMsg = fncNeedCheck(Trim(txtHokenStart_Hi))                    '必須チェック

    If strErrChkMsg = "" Then
        strErrChkMsg = fncNumCheck(Trim(txtHokenStart_Hi))                 '数字チェック
        If strErrChkMsg = "" Then
            strErrChkMsg = fncNumRangeCheck(Trim(txtHokenStart_Hi), 1, 31) '数値チェック
        End If
    End If
    If strErrChkMsg = "" Then
    Else
        strHokenStartErrMsg = strHokenStartErrMsg & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
    End If
    strErrContent = strErrContent & strHokenStartErrMsg

    '保険始期日_年月日
    strErrChkMsg = ""
    strErrKoumoku = "・保険始期日" & vbCrLf
    '新元号対応↓
'    If strHokenStartErrMsg = "" Then
'        strHokenStart = Format(Trim(txtHokenStart_Nen) + 1988, "0000") & "/" & _
'                        Format(Trim(txtHokenStart_Tsuki), "00") & "/" & _
'                        Format(Trim(txtHokenStart_Hi), "00")
'        strErrChkMsg = fncDateCheck(strHokenStart)                              '日付チェック
'
'        If strErrChkMsg = "" Then
'            strErrChkMsg = fncShikiCheck(strHokenStart)                              '保険始期チェック
'        End If
'        If strErrChkMsg = "" Then
'        Else
'            strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
'        End If
'
'    End If
    If strHokenStartErrMsg = "" Then
        strHokenStart = Format(Trim(txtHokenStart_Nen) + 2000, "0000") & "/" & _
                        Format(Trim(txtHokenStart_Tsuki), "00") & "/" & _
                        Format(Trim(txtHokenStart_Hi), "00")
        strErrChkMsg = fncDateCheck(strHokenStart)                              '日付チェック

        If strErrChkMsg = "" Then
            strErrChkMsg = fncShikiCheck(strHokenStart)                              '保険始期チェック
        End If
        If strErrChkMsg = "" Then
        Else
            strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
        End If

    End If
    '新元号対応↑
    '払込方法
    strErrChkMsg = ""

    strErrKoumoku = "・払込方法" & vbCrLf
    strErrChkMsg = fncNeedCheck(cmbHaraiHouhou)                             '必須チェック

    If strErrChkMsg = "" Then
    Else
        strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
    End If

    '優良割引
    If FleetTypeFlg = 1 Then '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        If Trim(txtYuuryowari) = "" Then                                    'ブランクの場合、何もしない
        Else
            strErrChkMsg = ""

            strErrKoumoku = "・優良割引" & vbCrLf
            strErrChkMsg = fncNumCheck(Trim(txtYuuryowari))                 '数字チェック
            If strErrChkMsg = "" Then
                strErrChkMsg = fncNumRangeCheck(Trim(txtYuuryowari), 1, 99) '数値チェック
            End If
            If strErrChkMsg = "" Then
            Else
                strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
            End If
        End If
    End If

    '第一種デメ割増
    If FleetTypeFlg = 1 Then '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        If Trim(txtFirstDeme) = "" Then                                    'ブランクの場合、何もしない
        Else
            strErrChkMsg = ""

            strErrKoumoku = "・第一種デメ割増" & vbCrLf
            strErrChkMsg = fncNumCheck(Trim(txtFirstDeme))                      '数字チェック
            If strErrChkMsg = "" Then
                strErrChkMsg = fncNumRangeCheck(Trim(txtFirstDeme), 1, 100)     '数値チェック
            End If
            If strErrChkMsg = "" Then
            Else
                strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
            End If
        End If
    End If

    '優良割引、第一種デメ割増
    If FleetTypeFlg = 1 Then '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        strErrChkMsg = ""

        strErrKoumoku = "○割引エラー" & vbCrLf
        strErrChkMsg = fncWariCheck(Trim(txtYuuryowari), Trim(txtFirstDeme))                    '関連チェック

        If strErrChkMsg = "" Then
        Else
            strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
        End If
    End If

    'フリート多数割引
    If FleetTypeFlg = 1 Then '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        strErrChkMsg = ""

        strErrKoumoku = "○フリート多数割引エラー" & vbCrLf
        strErrChkMsg = fncFleetTasuuCheck(cmbFreetkbn, chkFreetTasuu)                           '関連チェック

        If strErrChkMsg = "" Then
        Else
            strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
        End If
    End If

    'フリートコード
    If FleetTypeFlg = 1 Then '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        strErrChkMsg = ""

        strErrKoumoku = "・フリートコード" & vbCrLf
        strErrChkMsg = fncNeedCheck(Trim(txtFreetCode))                     '必須チェック

        If strErrChkMsg = "" Then
            strErrChkMsg = fncNumCheck(Trim(txtFreetCode))                  '数字チェック
            If strErrChkMsg = "" Then
                strErrChkMsg = fncKetaCheck(Trim(txtFreetCode), 5, "=") '桁数(同じ)チェック
            End If
        End If
        If strErrChkMsg = "" Then
        Else
            strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
        End If
    End If

    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    'ﾉﾝﾌﾘｰﾄ多数割引
    If FleetTypeFlg = 2 Then
        strErrChkMsg = ""

        strErrKoumoku = "○ﾉﾝﾌﾘｰﾄ多数割引" & vbCrLf
        strErrChkMsg = fncNonfleetTawariCheck(cmbNonfleetTawari.Value, txtSouFuhoDaisu.Value) 'ﾉﾝﾌﾘｰﾄ多数割引エラーチェック

        If strErrChkMsg = "" Then
        Else
            strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
        End If
    End If

    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    '団体割増引
    If FleetTypeFlg = 2 Then
        strErrChkMsg = ""

        If txtDantaiWarimashibiki.Value = "" Then
        Else
            strErrKoumoku = "・団体割増引" & vbCrLf

            strErrChkMsg = strErrChkMsg & fncDecimalCheck(txtDantaiWarimashibiki.Value) '数字チェック(マイナス、少数点入り)
            If strErrChkMsg = "" Then
                strErrChkMsg = strErrChkMsg & fncCommaCheck(txtDantaiWarimashibiki.Value) 'カンマチェック
                If strErrChkMsg = "" Then
                    strErrChkMsg = strErrChkMsg & fncNumRangeCheck(Val(Trim(txtDantaiWarimashibiki.Value)), -50, 10) '数値チェック
                End If
            End If
        End If

        If strErrChkMsg = "" Then
        Else
            strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
        End If
    End If


    ''エラー判定
    If strErrContent <> "" Then
        strErrContent = Left(strErrContent, Len(strErrContent) - 2)
        strErrContent = strErrContent & String(62, "-") & vbCrLf & "チェック完了" & " [ " & Format(Time, "HH:MM:SS") & " ]"
    End If

End Function

Private Function fncNeedOptCheck(ByVal blnKojin As Boolean, ByVal blnHojin As Boolean) As String
'関数名：fncNeedOptCheck
'内容　：必須チェック
'引数　：
'        blnValue       = 入力内容
    
    fncNeedOptCheck = ""

    If blnKojin = False And blnHojin = False Then
         fncNeedOptCheck = " 必須入力項目です。入力してください。"
    End If
    
End Function

Private Function fncWariCheck(ByVal strYuryoWari As String, ByVal strDemeWari As String) As String
'関数名：fncWariCheck
'内容　：関連チェック
'引数　：
'        blnValue       = 入力内容
    
    fncWariCheck = ""

    If strYuryoWari <> "" And strDemeWari <> "" Then
         fncWariCheck = " フリート優良割引と第一種デメ割増は同時に入力できません。"
    End If

    If strYuryoWari = "" And strDemeWari = "" Then
         fncWariCheck = " フリート優良割引、第一種デメ割増のいずれかを入力してください。"
    End If
    
End Function

Private Function fncFleetTasuuCheck(ByVal strFleetKbn As String, ByVal blnTasuuWari As Boolean) As String
'関数名：fncFleetTasuuCheck
'内容　：関連チェック
'引数　：
'        blnValue       = 入力内容

    fncFleetTasuuCheck = ""

    If strFleetKbn = "全車両一括" Or strFleetKbn = "全車両連結合算" Then
        If blnTasuuWari = False Then
            fncFleetTasuuCheck = " 全車両一括または全車両連結の場合、フリート多数割引が適用できます。"
        End If
    End If
    
End Function


