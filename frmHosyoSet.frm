VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHosyoSet 
   Caption         =   "補償内容セット（一括）"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14910
   OleObjectBlob   =   "frmHosyoSet.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmHosyoSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intConfirmMsg As Integer


Private Sub Frame89_Click()

End Sub

'初期表示
Private Sub UserForm_Initialize()
    Dim strSaveContent  As String

    strSaveContent = ""
    
    On Error GoTo Error
    
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    'フォームのタイトル設定
    If FleetTypeFlg = 1 Then
        frmHosyoSet.Caption = frmHosyoSet.Caption & "（フリート契約）"
    Else
        frmHosyoSet.Caption = frmHosyoSet.Caption & "（ノンフリート明細付契約）"
    End If
    
    'コード値シート
    Dim wsCode As Worksheet
    If FleetTypeFlg = 1 Then
        Call subSetSheet(4, wsCode)         'シートオブジェクト(別紙　コード値)
    Else
        Call subSetSheet(16, wsCode)        'シートオブジェクト(別紙　コード値（ノンフリート）)（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
    End If
        
    '車両保険の種類
    With CmbHknSyurui
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("BJ2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("BJ2").Column + 1).End(xlUp)).Value
        .ColumnWidths = "-1;0"
    End With
    
    '車両免責金額
    With CmbSyaryoMskGaku
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("BN2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("BN2").Column + 1).End(xlUp)).Value
        .ColumnWidths = "-1;0"
    End With
        
    '対人賠償
    With CmbTaijinBaisyo
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("CD2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("CD2").Column).End(xlUp)).Value
    End With
        
    '対物賠償
    With CmbTaibutsuBaisyo
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("CH2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("CH2").Column).End(xlUp)).Value
    End With
    
    '対物免責金額
    With CmbTaibutsuMskGaku
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("BR2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("BR2").Column + 1).End(xlUp)).Value
        .ColumnWidths = "-1;0"
    End With
    
    '人身傷害(1名)
    With CmbZinshinSyougai_1Mei
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("CL2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("CL2").Column).End(xlUp)).Value
    End With
    
'    '搭乗者傷害(1名)
'    With CmbTouzyouSyougai_1Mei
'        .AddItem ""
'        .List = wsCode.Range(wsCode.Range("CP2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("CP2").Column).End(xlUp)).Value
'    End With
    
    '代車等セット特約
    With CmbDaisyaToku
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("CA2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("CA2").Column).End(xlUp)).Value
'        .ColumnWidths = "0;-1"
    End With
    
    'ファミリーバイク特約（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
    With CmbFamiryBikeToku
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("BV2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("BV2").Column + 1).End(xlUp)).Value
        .ColumnWidths = "-1;0"
    End With
    
    
    '画面下部のフレーム
    If FleetTypeFlg = 1 Then
        'フリート
        FrameNonFleet1.Visible = False
        FrameNonFleet2.Visible = False
    Else
        'ノンフリート（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        FrameFleet.Visible = False
    End If
    
    
    '保存情報がある場合反映する
    If fncFormRef(2, strSaveContent) Then

    Else
        Call subItemGet(strSaveContent)
    End If
         
    Set wsCode = Nothing
    
    On Error GoTo 0
    
    Exit Sub
    
Error:
    MsgBox "UserForm_Initialize" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"

End Sub

'保存情報を画面に反映
Private Sub subItemGet(ByVal strAllSetItem As String)
    Dim varSaveContent As Variant
    
    varSaveContent = Split(strAllSetItem, "/")

    '引数を項目にセット
    CmbHknSyurui.Value = varSaveContent(0)                      '保険の種類
    CmbSyaryoMskGaku.Value = varSaveContent(1)                  '車両免責金額
    ChkHknZnsnToku.Value = CBool(varSaveContent(2))             '保険全損臨費特約
    ChkSyaryoTyoukaToku.Value = CBool(varSaveContent(3))        '車両超過修理費用特約
    ChkSyaryoTonanToku.Value = CBool(varSaveContent(4))         '車両盗難対象外特約
    CmbTaijinBaisyo.Value = varSaveContent(5)                   '対人賠償
    CmbTaibutsuBaisyo.Value = varSaveContent(6)                 '対物賠償
    CmbTaibutsuMskGaku.Value = varSaveContent(7)                '対物免責金額
    ChkTaibutsuTyoukaToku.Value = CBool(varSaveContent(8))      '対物超過修理費用特約
    CmbZinshinSyougai_1Mei.Value = varSaveContent(9)            '人身傷害(1名)
    TxtZinshinSyougai_1Jiko.Value = varSaveContent(10)          '人身傷害(1事故)
    TxtTouzyouSyougai_1Mei.Value = varSaveContent(11)           '搭乗者傷害(1名)
    ChkTouzyouSyougai_1Mei_Taisyougai.Value = _
                                  CBool(varSaveContent(12))     '搭乗者傷害(1名)対象外
    TxtTouzyouSyougai_1Jiko.Value = varSaveContent(13)          '搭乗者傷害(1事故)
    ChkNissuToku.Value = CBool(varSaveContent(14))              '日数払特約
    ChkJigyouNushiToku.Value = CBool(varSaveContent(15))        '事業主費用特約
    ChkBengoshiToku.Value = CBool(varSaveContent(16))           '弁護士費用特約
    ChkJisonJikoToku.Value = CBool(varSaveContent(17))          '自損事故傷害特約
    ChkMuhokenToku.Value = CBool(varSaveContent(18))            '無保険車事故傷害特約
    CmbDaisyaToku.Value = varSaveContent(19)                    '事故代車・身の回り品補償特約
    ChkSyohiyou_Futekiyou.Value = CBool(varSaveContent(20))     '車両搬送時不適用特約
    ChkJugyouinToku.Value = CBool(varSaveContent(21))           '従業員等限定特約（フリートのみ）
    CmbFamiryBikeToku.Value = varSaveContent(22)                'ファミリーバイク特約（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
    CheckKojinBaisekiToku.Value = CBool(varSaveContent(23))     '個人賠償責任補償特約（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
    ChkCarJikoToku.Value = CBool(varSaveContent(24))            '自動車事故特約　　　（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）

End Sub


'「一括セット」ボタン押下
Private Sub BtnHosyouSet_Click()
    Dim strErrContent As String
    Dim strAllSetItem As String

    strErrContent = ""
    strAllSetItem = ""
    
    On Error GoTo Error

    '入力チェック
    Call fncEntryErrCheckHosyoSet(strErrContent)

    'エラー内容表示
    TxtErrBox = strErrContent & "チェック完了" & " [ " & Format(Time, "HH:MM:SS") & " ]"
    'エラーテキストボックスのスクロールバー移動
    TxtErrBox.SetFocus
    TxtErrBox.SelStart = 0

    'エラーがある場合は処理を抜ける
    If Not strErrContent = "" Then
        Exit Sub
    End If
    
    BtnHosyouSet.SetFocus

    Dim intConfirmMsg As Integer
    intConfirmMsg = MsgBox("入力内容を明細入力画面に反映します。" & vbCrLf & "よろしいですか?", vbYesNo, "確認ダイアログ")
    If intConfirmMsg = 7 Then
    Else
        Call subItemSet(strAllSetItem)
        
        Call subMeisaiUnProtect     'シートの保護の解除

        If fncFormSave(2, strAllSetItem) Then
            'シート・ブックの表示
            Call subBookUnProtect 'ブックの保護を解除 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
            Call subSheetVisible(True)
            Call subBookProtect   'ブックの保護 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
            
            Call subAllSet(strAllSetItem)
            
            Call subMeisaiProtect       'シートの保護
            Unload Me
        End If
    End If
    
    
    On Error GoTo 0
    
    Exit Sub
    
Error:
    MsgBox "BtnHosyouSet_Click" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"

End Sub

Private Function fncEntryErrCheckHosyoSet(ByRef strErrContent As String)
'関数名：EntryErrCheck
'内容　：エラーチェックを行いエラーがある場合は,エラー内容を返す。
'引数　：strErrContent = :エラー内容

    Dim strErrCheck As String
    Dim blnErrFlg As Boolean
    Dim strFormname As UserForm
    Set strFormname = frmHosyoSet

    strErrCheck = ""
    strErrContent = ""
    blnErrFlg = False

    With strFormname

        ''○車両保険エラー
        strErrCheck = fncHknSyuruiCheck(.CmbHknSyurui.Value, .CmbSyaryoMskGaku.Value, .CmbDaisyaToku.Value, _
                                        CStr(.ChkHknZnsnToku.Value), CStr(ChkSyaryoTonanToku.Value), CStr(ChkSyaryoTyoukaToku.Value))
        If strErrCheck = "" Then
        Else
            strErrContent = strErrContent & "○車両保険エラー" & vbCrLf
            strErrContent = strErrContent & strErrCheck & vbCrLf & vbCrLf
        End If

        ''対人賠償
        '必須チェック
        strErrCheck = fncNeedCheck(.CmbTaijinBaisyo.Value)
        If strErrCheck = "" Then
        Else
            strErrContent = strErrContent & "・対人賠償" & vbCrLf & " " & strErrCheck & vbCrLf & vbCrLf
        End If

        ''対物賠償
        '必須チェック
        strErrCheck = fncNeedCheck(.CmbTaibutsuBaisyo.Value)
        If strErrCheck = "" Then
            ''○対物賠償エラー
            strErrCheck = fncTaibutsuBaisyo(.CmbTaibutsuBaisyo.Value, .CmbTaibutsuMskGaku.Value, CStr(.ChkTaibutsuTyoukaToku.Value))
            If strErrCheck = "" Then
            Else
                strErrContent = strErrContent & "○対物賠償エラー" & vbCrLf
                strErrContent = strErrContent & strErrCheck & vbCrLf & vbCrLf
            End If
        Else
            strErrContent = strErrContent & "・対物賠償" & vbCrLf & " " & strErrCheck & vbCrLf & vbCrLf
        End If

        ''人身傷害(1名)
        '必須チェック
        strErrCheck = fncNeedCheck(.CmbZinshinSyougai_1Mei.Value)
        If strErrCheck = "" Then
            ''○人身傷害エラー
            strErrCheck = fncZinshinSyougai(.CmbZinshinSyougai_1Mei.Value, .TxtZinshinSyougai_1Jiko.Value)
            If strErrCheck = "" Then
            Else
                strErrContent = strErrContent & "○人身傷害エラー" & vbCrLf
                strErrContent = strErrContent & strErrCheck & vbCrLf & vbCrLf
                blnErrFlg = True
            End If
        Else
            strErrContent = strErrContent & "・人身傷害(1名)" & vbCrLf & " " & strErrCheck & vbCrLf & vbCrLf
        End If
        
        If blnErrFlg = False Then
            ''人身傷害(1事故)
            If Trim(.TxtZinshinSyougai_1Jiko.Value) = "" Then
            Else
                '数字チェック
                strErrCheck = fncNumCheck(Trim(.TxtZinshinSyougai_1Jiko.Value))
                If Not strErrCheck = "" Then
                    strErrContent = strErrContent & "・人身傷害(1事故)" & vbCrLf & " " & strErrCheck & vbCrLf & vbCrLf
                Else
                    '桁数チェック
                    strErrCheck = fncKetaCheck(Trim(.TxtZinshinSyougai_1Jiko.Value), 6, ">")
                    If Not strErrCheck = "" Then
                        strErrContent = strErrContent & "・人身傷害(1事故)" & vbCrLf & " " & strErrCheck & vbCrLf & vbCrLf
                    End If
                End If
            End If
        End If

        blnErrFlg = False

        ''○対人賠償エラー
        strErrCheck = fncTaijinBaisyoCheck(.CmbTaijinBaisyo.Value, CStr(ChkJisonJikoToku.Value), _
                                            .CmbZinshinSyougai_1Mei.Value, CStr(.ChkMuhokenToku.Value))
        If Not strErrCheck = "" Then
            strErrContent = strErrContent & "○対人賠償エラー" & vbCrLf
            strErrContent = strErrContent & strErrCheck & vbCrLf & vbCrLf
        End If

        ''搭乗者傷害(1名)
        '必須チェック
'        strErrCheck = fncNeedCheck(.TxtTouzyouSyougai_1Mei.Value)
'        If strErrCheck = "" Then
            ''○搭乗者傷害エラー
        strErrCheck = fncTouzyouSyougai(.TxtTouzyouSyougai_1Mei.Value, Trim(.ChkTouzyouSyougai_1Mei_Taisyougai.Value), _
                                        Trim(.TxtTouzyouSyougai_1Jiko.Value), CStr(.ChkNissuToku.Value))
        If strErrCheck = "" Then
        Else
            strErrContent = strErrContent & "○搭乗者傷害エラー" & vbCrLf
            strErrContent = strErrContent & strErrCheck & vbCrLf & vbCrLf
            blnErrFlg = True
        End If
'        Else
'            strErrContent = strErrContent & "・搭乗者傷害(1名)" & vbCrLf & " " & strErrCheck & vbCrLf & vbCrLf
'        End If
        
        If blnErrFlg = False Then
            ''搭乗者傷害(1事故)
            If Trim(.TxtTouzyouSyougai_1Jiko.Value) = "" Then
            Else
                '数字チェック
                strErrCheck = fncNumCheck(Trim(.TxtTouzyouSyougai_1Jiko.Value))
                If strErrCheck = "" Then
                    '桁数チェック
                    strErrCheck = fncKetaCheck(Trim(.TxtTouzyouSyougai_1Jiko.Value), 6, ">")
                    If Not strErrCheck = "" Then
                        strErrContent = strErrContent & "・搭乗者傷害(1事故)" & vbCrLf & " " & strErrCheck & vbCrLf & vbCrLf
                    End If
                Else
                    strErrContent = strErrContent & "・搭乗者傷害(1事故)" & vbCrLf & " " & strErrCheck & vbCrLf & vbCrLf
                End If
            End If
        End If
        
        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        'ファミリーバイク特約
        If FleetTypeFlg = 2 Then
            ''○ファミリーバイク特約エラー
            strErrCheck = fncFamilyBikeCheck(CmbFamiryBikeToku.Value, CmbZinshinSyougai_1Mei.Value)
            
            If strErrCheck = "" Then
            Else
                strErrContent = strErrContent & "○ファミリーバイク特約エラー" & vbCrLf
                strErrContent = strErrContent & strErrCheck & vbCrLf & vbCrLf
            End If
        End If
        
        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        '自動車事故特約
        If FleetTypeFlg = 2 Then
            ''○自動車事故特約エラー
            strErrCheck = fncJidoushaJikoCheck(ChkCarJikoToku.Value, CmbZinshinSyougai_1Mei.Value)
            
            If strErrCheck = "" Then
            Else
                strErrContent = strErrContent & "○自動車事故特約エラー" & vbCrLf
                strErrContent = strErrContent & strErrCheck & vbCrLf & vbCrLf
            End If
        End If
        
        ''エラー判定
        If strErrContent <> "" Then
            strErrContent = Left(strErrContent, Len(strErrContent) - 2)
            strErrContent = strErrContent & String(62, "-") & vbCrLf
        End If
        
    End With
    
End Function


'画面情報を取得
Private Sub subItemSet(ByRef strAllSetItem As String)
    
    '項目を引数にセット
    strAllSetItem = strAllSetItem & CmbHknSyurui.Value & "/"                    '保険の種類
    strAllSetItem = strAllSetItem & CmbSyaryoMskGaku.Value & "/"                '車両免責金額
    strAllSetItem = strAllSetItem & CStr((ChkHknZnsnToku.Value)) & "/"          '保険全損臨費特約
    strAllSetItem = strAllSetItem & CStr((ChkSyaryoTyoukaToku.Value)) & "/"     '車両超過修理費用特約
    strAllSetItem = strAllSetItem & CStr((ChkSyaryoTonanToku.Value)) & "/"      '車両盗難対象外特約
    strAllSetItem = strAllSetItem & CmbTaijinBaisyo.Value & "/"                 '対人賠償
    strAllSetItem = strAllSetItem & CmbTaibutsuBaisyo.Value & "/"               '対物賠償
    strAllSetItem = strAllSetItem & CmbTaibutsuMskGaku.Value & "/"              '対物免責金額
    strAllSetItem = strAllSetItem & CStr(ChkTaibutsuTyoukaToku.Value) & "/"     '対物超過修理費用特約
    strAllSetItem = strAllSetItem & CmbZinshinSyougai_1Mei.Value & "/"          '人身傷害(1名)
    strAllSetItem = strAllSetItem & TxtZinshinSyougai_1Jiko.Value & "/"         '人身傷害(1事故)
    strAllSetItem = strAllSetItem & TxtTouzyouSyougai_1Mei.Value & "/"          '搭乗者傷害(1名)
    strAllSetItem = strAllSetItem & _
                                  CStr(ChkTouzyouSyougai_1Mei_Taisyougai) & "/" '搭乗者傷害(1名)対象外
    strAllSetItem = strAllSetItem & TxtTouzyouSyougai_1Jiko.Value & "/"         '搭乗者傷害(1事故)
    strAllSetItem = strAllSetItem & CStr(ChkNissuToku.Value) & "/"              '日数払特約
    strAllSetItem = strAllSetItem & CStr(ChkJigyouNushiToku.Value) & "/"        '事業主費用特約
    strAllSetItem = strAllSetItem & CStr(ChkBengoshiToku.Value) & "/"           '弁護士費用特約
    strAllSetItem = strAllSetItem & CStr(ChkJisonJikoToku.Value) & "/"          '自損事故傷害特約
    strAllSetItem = strAllSetItem & CStr(ChkMuhokenToku.Value) & "/"            '無保険車事故傷害特約
    strAllSetItem = strAllSetItem & CmbDaisyaToku.Value & "/"                   '代車等セット特約
    strAllSetItem = strAllSetItem & CStr(ChkJugyouinToku.Value) & "/"           '従業員等限定特約
    strAllSetItem = strAllSetItem & CStr(ChkSyohiyou_Futekiyou.Value) & "/"     '車両搬送時不適用特約
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    strAllSetItem = strAllSetItem & CmbFamiryBikeToku.Value & "/"               'ファミリーバイク特約
    strAllSetItem = strAllSetItem & CStr(CheckKojinBaisekiToku.Value) & "/"     '個人賠償責任補償特約
    strAllSetItem = strAllSetItem & CStr(ChkCarJikoToku.Value) & "/"            '自動車事故特約
     
End Sub


'明細入力画面に一括セット
Private Sub subAllSet(ByVal strAllSetItem As String)
    Dim varSaveContent As Variant
    Dim i As Integer
    Dim strStartRow As String       '開始行
    Dim strAllCnt As String         '総付保台数
    Dim objAll As OLEObject
    Dim strAllCell As String

    i = 0
    
    Dim wsMeisai As Worksheet
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(1, wsMeisai)       'シートオブジェクト(明細入力)
    Else
        Call subSetSheet(17, wsMeisai)      'シートオブジェクト(明細入力（ノンフリート）)
    End If
    
    '計算用シートの開始行
    strStartRow = 21
    
    Dim objSouhuho As Object
    Set objSouhuho = wsMeisai.OLEObjects("txtSouhuho").Object
    '送付保台数のセル番号
    strAllCell = Left(objSouhuho.Value, Len(objSouhuho.Value) - 2)
    
    'フォームの値を配列化
    varSaveContent = Split(strAllSetItem, "/")

    'シートのセルに格納
    If FleetTypeFlg = 1 Then  'フリート
        For i = 0 To Val(strAllCell) - 1
        
            wsMeisai.Cells(strStartRow + i, 25) = varSaveContent(0)                                         '保険の種類
            wsMeisai.Cells(strStartRow + i, 27) = varSaveContent(1)                                         '車両免責金額
            wsMeisai.Cells(strStartRow + i, 28) = IIf(CBool(varSaveContent(2)) = True, "適用する", "")      '保険全損臨費特約
            wsMeisai.Cells(strStartRow + i, 29) = IIf(CBool(varSaveContent(3)) = True, "適用する", "")      '車両超過修理費用特約
            wsMeisai.Cells(strStartRow + i, 30) = IIf(CBool(varSaveContent(4)) = True, "対象外", "")        '車両盗難対象外特約
            wsMeisai.Cells(strStartRow + i, 31) = varSaveContent(5)                                         '対人賠償
            wsMeisai.Cells(strStartRow + i, 32) = IIf(CBool(varSaveContent(17)) = True, "対象外", "")       '自損事故傷害特約
            wsMeisai.Cells(strStartRow + i, 33) = IIf(CBool(varSaveContent(18)) = True, "対象外", "")       '無保険車事故傷害特約
            wsMeisai.Cells(strStartRow + i, 34) = varSaveContent(6)                                         '対物賠償
            wsMeisai.Cells(strStartRow + i, 35) = varSaveContent(7)                                         '対物免責金額
            wsMeisai.Cells(strStartRow + i, 36) = IIf(CBool(varSaveContent(8)) = True, "適用する", "")      '対物超過修理費用特約
            wsMeisai.Cells(strStartRow + i, 37) = varSaveContent(9)                                         '人身傷害(1名)
            wsMeisai.Cells(strStartRow + i, 38) = varSaveContent(10)                                        '人身傷害(1事故)
            wsMeisai.Cells(strStartRow + i, 39) = IIf(CBool(varSaveContent(12)) = _
                                                                        True, "対象外", varSaveContent(11)) '搭乗者傷害(1名)
            wsMeisai.Cells(strStartRow + i, 40) = varSaveContent(13)                                        '搭乗者傷害(1事故)
            wsMeisai.Cells(strStartRow + i, 41) = IIf(CBool(varSaveContent(14)) = True, "適用する", "")     '日数払特約
            wsMeisai.Cells(strStartRow + i, 42) = IIf(CBool(varSaveContent(15)) = True, "適用する", "")     '事業主費用特約
            wsMeisai.Cells(strStartRow + i, 43) = IIf(CBool(varSaveContent(16)) = True, "適用する", "")     '弁護士費用特約
            wsMeisai.Cells(strStartRow + i, 44) = IIf(CBool(varSaveContent(20)) = True, "限定", "")         '従業員等限定特約
            wsMeisai.Cells(strStartRow + i, 45) = varSaveContent(19)                                        '代車等セット特約
            wsMeisai.Cells(strStartRow + i, 46) = IIf(CBool(varSaveContent(21)) = True, "不適用", "")       '車両搬送時不適用特約

        Next
        
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    Else  'ノンフリート
         For i = 0 To Val(strAllCell) - 1
        
            wsMeisai.Cells(strStartRow + i, 33) = varSaveContent(0)                                         '保険の種類
            wsMeisai.Cells(strStartRow + i, 35) = varSaveContent(1)                                         '車両免責金額
            wsMeisai.Cells(strStartRow + i, 36) = IIf(CBool(varSaveContent(2)) = True, "適用する", "")      '保険全損臨費特約
            wsMeisai.Cells(strStartRow + i, 37) = IIf(CBool(varSaveContent(3)) = True, "適用する", "")      '車両超過修理費用特約
            wsMeisai.Cells(strStartRow + i, 38) = IIf(CBool(varSaveContent(4)) = True, "対象外", "")        '車両盗難対象外特約
            wsMeisai.Cells(strStartRow + i, 39) = varSaveContent(5)                                         '対人賠償
            wsMeisai.Cells(strStartRow + i, 40) = IIf(CBool(varSaveContent(17)) = True, "対象外", "")       '自損事故傷害特約
            wsMeisai.Cells(strStartRow + i, 41) = IIf(CBool(varSaveContent(18)) = True, "対象外", "")       '無保険車事故傷害特約
            wsMeisai.Cells(strStartRow + i, 42) = varSaveContent(6)                                         '対物賠償
            wsMeisai.Cells(strStartRow + i, 43) = varSaveContent(7)                                         '対物免責金額
            wsMeisai.Cells(strStartRow + i, 44) = IIf(CBool(varSaveContent(8)) = True, "適用する", "")      '対物超過修理費用特約
            wsMeisai.Cells(strStartRow + i, 45) = varSaveContent(9)                                         '人身傷害(1名)
            wsMeisai.Cells(strStartRow + i, 46) = varSaveContent(10)                                        '人身傷害(1事故)
            wsMeisai.Cells(strStartRow + i, 47) = IIf(CBool(varSaveContent(12)) = _
                                                                        True, "対象外", varSaveContent(11)) '搭乗者傷害(1名)
            wsMeisai.Cells(strStartRow + i, 48) = varSaveContent(13)                                        '搭乗者傷害(1事故)
            wsMeisai.Cells(strStartRow + i, 49) = IIf(CBool(varSaveContent(14)) = True, "適用する", "")     '日数払特約
            wsMeisai.Cells(strStartRow + i, 50) = IIf(CBool(varSaveContent(15)) = True, "適用する", "")     '事業主費用特約
            wsMeisai.Cells(strStartRow + i, 51) = IIf(CBool(varSaveContent(16)) = True, "適用する", "")     '弁護士費用特約
            wsMeisai.Cells(strStartRow + i, 52) = varSaveContent(22)                                        'ファミリーバイク特約
            wsMeisai.Cells(strStartRow + i, 53) = IIf(CBool(varSaveContent(23)) = True, "適用する", "")     '個人賠償責任補償特約
            wsMeisai.Cells(strStartRow + i, 54) = IIf(CBool(varSaveContent(24)) = True, "適用する", "")     '自動車事故特約
            wsMeisai.Cells(strStartRow + i, 55) = varSaveContent(19)                                        '代車等セット特約
            wsMeisai.Cells(strStartRow + i, 56) = IIf(CBool(varSaveContent(21)) = True, "不適用", "")       '車両搬送時不適用特約

            
        Next
    
    End If
    
    '明細入力画面のエラー用リスト初期化
    wsMeisai.OLEObjects("txtErrMsg").Object.Value = ""
    wsMeisai.OLEObjects("txtErrMsg").Activate
    wsMeisai.Range("A1").Activate
    
    Set wsMeisai = Nothing
    Set objSouhuho = Nothing
    
End Sub


'「戻る」ボタン押下
Private Sub BtnBack_Click()

    On Error GoTo Error

    Dim intConfirmMsg As Integer
    intConfirmMsg = MsgBox("入力内容を反映せずに明細入力画面に遷移します。" & vbCrLf & "よろしいですか?", vbYesNo, "確認ダイアログ")
    If intConfirmMsg = 7 Then
    Else
        'シート・ブックの表示
        Call subBookUnProtect 'ブックの保護を解除 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        Call subSheetVisible(True)
        Call subBookProtect   'ブックの保護 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加

        Unload Me
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


