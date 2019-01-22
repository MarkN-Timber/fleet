VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEntryMoushikomi 
   Caption         =   "申込書印刷"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10905
   OleObjectBlob   =   "frmEntryMoushikomi.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmEntryMoushikomi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intConfirmMsg As Integer


'初期表示
Private Sub UserForm_Initialize()
    Dim wsSetting As Worksheet

    On Error GoTo Error

    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    'フォームのタイトル設定
    If FleetTypeFlg = 1 Then
        frmEntryMoushikomi.Caption = frmEntryMoushikomi.Caption & "（フリート契約）"
    Else
        frmEntryMoushikomi.Caption = frmEntryMoushikomi.Caption & "（ノンフリート明細付契約）"
    End If

    Call subSetSheet(5, wsSetting)      'シートオブジェクト(別紙　各種設定)（ご注意事項文言保管）

    '団体入力欄　表示・非表示（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
    If FleetTypeFlg = 1 Then
        FrameDantai.Visible = False
        FremeToriatukaiInfo.Top = 378
        FrameWarningMemo.Top = 450
        frmEntryMoushikomi.Height = 537
        frmEntryMoushikomi.ScrollBars = fmScrollBarsNone
        frmEntryMoushikomi.ScrollHeight = 0
        'ご注意事項
        lblWarningMemo = wsSetting.Range("B6")
    Else
        lblWarningMemo = wsSetting.Range("B7")
    End If

    Set wsSetting = Nothing

    '保存情報がある場合反映する
    Dim strSaveContent As String
    Dim varSaveContent As Variant
    If fncMoushikomiFormRef(8, strSaveContent) Then
        strSaveContent = ""
        If fncMoushikomiFormRef(6, strSaveContent) Then
        Else
            varSaveContent = Split(strSaveContent, "/")
            Dim item As String
            Dim i As Long
            For i = 20 To 37
                item = item & varSaveContent(i)
            Next i
            
            If item = "" Then
            Else

                '申込書画面の各項目に保存内容をセット
                txtShokenNo = varSaveContent(37)                                                       '証券番号
                If Not varSaveContent(20) = "" Then
'                    txtPostNo_zen = Left(varSaveContent(20), InStr(varSaveContent(20), "－") - 1)      '郵便番号
'                    txtPostNo_kou = Right(Replace(varSaveContent(20), "－", "    "), 4)                '郵便番号
                    txtPostNo_zen = RTrim(Left(varSaveContent(20), 3))                                '郵便番号
                    If Len(varSaveContent(20)) > 3 Then
                        txtPostNo_kou = Mid(varSaveContent(20), 4)
                    Else
                        txtPostNo_kou = ""
                    End If
                Else
                    txtPostNo_zen = ""
                    txtPostNo_kou = ""
                End If
                txtKeiyakujusyo_kana = StrConv(StrConv(varSaveContent(21), vbKatakana), vbNarrow)      '契約者住所(ｶﾅ)
                txtKeiyakujusyo_kanji1 = RTrim(Left(varSaveContent(22), 40))                           '契約者住所(漢字)
                If Len(varSaveContent(22)) > 40 Then
                    txtKeiyakujusyo_kanji2 = Mid(varSaveContent(22), 41)                               '契約者住所(漢字)
                Else
                    txtKeiyakujusyo_kanji2 = ""
                End If
                txtHojin_kana = StrConv(StrConv(varSaveContent(23), vbKatakana), vbNarrow)             '法人名(ｶﾅ)
                txtHojin_kanji = varSaveContent(24)                                                    '法人名(漢字)
                txtYakusyoku_Shimei_kana = StrConv(StrConv(varSaveContent(25), vbKatakana), vbNarrow)  '役職名・氏名(ｶﾅ)
                txtYakusyoku_Shimei_kanji = varSaveContent(26)                                         '役職名・氏名(漢字)
                If Not varSaveContent(27) = "" Then
                    txtTelNo_Home_zen = Left(varSaveContent(27), InStr(varSaveContent(27), "-") - 1)  '連絡先1　自宅・携帯
                    txtTelNo_Home_chuu = Mid(Replace(varSaveContent(27), "-", "     "), _
                                                    InStr(varSaveContent(27), "-") + 5, 5)            '連絡先1　自宅・携帯
                    txtTelNo_Home_kou = Right(Replace(varSaveContent(27), "-", "    "), 4)            '連絡先1　自宅・携帯
                Else
                    txtTelNo_Home_zen = ""
                    txtTelNo_Home_chuu = ""
                    txtTelNo_Home_kou = ""
                End If
                If Not varSaveContent(28) = "" Then
                    txtTelNo_Kinmu_zen = Left(varSaveContent(28), InStr(varSaveContent(28), "-") - 1) '連絡先2　勤務先
                    txtTelNo_Kinmu_chuu = Mid(Replace(varSaveContent(28), "-", "     "), _
                                                  InStr(varSaveContent(28), "-") + 5, 5)              '連絡先2　勤務先
                    txtTelNo_Kinmu_kou = Right(Replace(varSaveContent(28), "-", "    "), 4)           '連絡先2　勤務先
                Else
                    txtTelNo_Kinmu_zen = ""
                    txtTelNo_Kinmu_chuu = ""
                    txtTelNo_Kinmu_kou = ""
                End If
                If Not varSaveContent(29) = "" Then
                    txtTelNo_Fax_zen = Left(varSaveContent(29), InStr(varSaveContent(29), "-") - 1)   '連絡先3　FAX
                    txtTelNo_Fax_chuu = Mid(Replace(varSaveContent(29), "-", "     "), _
                                                            InStr(varSaveContent(29), "-") + 5, 5)    '連絡先3　FAX
                    txtTelNo_Fax_kou = Right(Replace(varSaveContent(29), "-", "    "), 4)             '連絡先3　FAX
                Else
                    txtTelNo_Fax_zen = ""
                    txtTelNo_Fax_chuu = ""
                    txtTelNo_Fax_kou = ""
                End If
                '団体コードから社員コードまではノンフリートのみ
                If FleetTypeFlg <> 1 Then
                     'ノンフリートの場合
                    txtDantaimei = varSaveContent(30)                 '団体名
                    txtDantaiCode = varSaveContent(31)                '団体コード
                    txtDantaiToku = varSaveContent(32)                '団体扱に関する特約
                    txtShozoku = varSaveContent(33)                   '所属コード
                    txtSyainCode = varSaveContent(34)                 '社員コード
                    txtToriatsutenShop_code = varSaveContent(35)      '取扱店コード
                    txtDairiShop_code = varSaveContent(36)            '代理店コード
                Else
                    'フリートの場合
                    txtToriatsutenShop_code = varSaveContent(35)      '取扱店コード
                    txtDairiShop_code = varSaveContent(36)            '代理店コード
                End If
            End If
        End If
    Else
'        Dim varSaveContent As Variant
        varSaveContent = Split(strSaveContent, "/")

        '申込書画面の各項目に保存内容をセット
        txtShokenNo = varSaveContent(0)                       '証券番号
        txtPostNo_zen = varSaveContent(1)                     '郵便番号
        txtPostNo_kou = varSaveContent(2)                     '郵便番号
        txtKeiyakujusyo_kana = varSaveContent(3)              '契約者住所(カナ)
        txtKeiyakujusyo_kanji1 = varSaveContent(4)            '契約者住所(漢字)
        txtKeiyakujusyo_kanji2 = varSaveContent(5)            '契約者住所(漢字)
        txtHojin_kana = varSaveContent(6)                     '法人名(カナ)
        txtHojin_kanji = varSaveContent(7)                    '法人名(漢字)
        txtYakusyoku_Shimei_kana = varSaveContent(8)          '役職名・氏名(カナ)
        txtYakusyoku_Shimei_kanji = varSaveContent(9)         '役職名・氏名(漢字)
        txtTelNo_Home_zen = varSaveContent(10)                '連絡先1　自宅・携帯
        txtTelNo_Home_chuu = varSaveContent(11)               '連絡先1　自宅・携帯
        txtTelNo_Home_kou = varSaveContent(12)                '連絡先1　自宅・携帯
        txtTelNo_Kinmu_zen = varSaveContent(13)               '連絡先2　勤務先
        txtTelNo_Kinmu_chuu = varSaveContent(14)              '連絡先2　勤務先
        txtTelNo_Kinmu_kou = varSaveContent(15)               '連絡先2　勤務先
        txtTelNo_Fax_zen = varSaveContent(16)                 '連絡先3　FAX
        txtTelNo_Fax_chuu = varSaveContent(17)                '連絡先3　FAX
        txtTelNo_Fax_kou = varSaveContent(18)                 '連絡先3　FAX
        
        '団体コードから社員コードまではノンフリートのみ
        If FleetTypeFlg <> 1 Then
             'ノンフリートの場合
            txtDantaimei = varSaveContent(19)                 '団体名
            txtDantaiCode = varSaveContent(20)                '団体コード
            txtDantaiToku = varSaveContent(21)                '団体扱に関する特約
            txtShozoku = varSaveContent(22)                   '所属コード
            txtSyainCode = varSaveContent(23)                 '社員コード
        End If
        txtToriatsutenShop = varSaveContent(24)           '取扱店
        txtToriatsutenShop_code = varSaveContent(25)      '取扱店コード
        txtDairiShop = varSaveContent(26)                 '代理店
        txtDairiShop_code = varSaveContent(27)            '代理店コード
        txtBosyuuninID = varSaveContent(28)               '募集人ID
        
    End If
    
    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "UserForm_Initialize" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"
    
End Sub


'「×」ボタン押下時
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


'「戻る」ボタン押下
Private Sub BtnBack_Click()
    Dim ctlFormCtrl As Control
    Dim wsSave As Worksheet     '画面の状態が保存されているシート
    Dim i As Integer
    Dim j As Integer
    Dim objWorkSheet As Worksheet

    On Error GoTo Error

    intConfirmMsg = MsgBox("入力内容を削除して帳票選択画面に遷移します。" & vbCrLf & "よろしいですか？", vbYesNo, "確認ダイアログ")
    If intConfirmMsg = 6 Then

        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        If FleetTypeFlg = 1 Then  'フリート
            Call subSetSheet(18, wsSave)       'シートオブジェクト(明細書印刷)
        Else
            Call subSetSheet(19, wsSave)       'シートオブジェクト(明細書印刷（ノンフリート）)
        End If

        '明細書印刷シートの保護解除（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        subMeisaiPrtUnProtect

        'イベント無効
        Application.EnableEvents = False

        If FleetTypeFlg = 1 Then
            'フリート
            wsSave.Range(8 & ":" & wsSave.Rows.Count).Delete
            wsSave.Range("C7:F7") = ""

        Else

            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加

            '被保険者情報の行（C列(3列）7行目からAA列（27列）15行目まで）内容消去
             For i = 7 To 15
               For j = 3 To 27
                wsSave.Cells(i, j).Value = ""
              Next j
             Next i

            '車両情報の行（C列（3列）19行目からAE列（31列）27行目まで）内容消去
             For i = 19 To 27
               For j = 3 To 31
                wsSave.Cells(i, j).Value = ""
              Next j
             Next i
                    
            '前契約情報の行（C列（3列）31行目からAA列（27列）39行目まで）内容消去
             For i = 31 To 39
               For j = 3 To 27
                wsSave.Cells(i, j).Value = ""
              Next j
             Next i
             
            '証券番号と明細番号の間のハイフンを入力
             For i = 31 To 39
                wsSave.Cells(i, 4).Value = "-"
             Next i
        
            '非表示にしていた行を全て表示
            wsSave.Rows.Hidden = False
        
            '明細入力画面のエラー用リスト初期化
            wsSave.OLEObjects("txtErrMsg").Object.Value = ""
            
        End If
    
        Application.EnableEvents = True     'イベント有効
    
        subMeisaiPrtProtect  'シートの保護

        
        '画面初期化
        For Each ctlFormCtrl In frmEntryMoushikomi.Controls
            If TypeName(ctlFormCtrl) = "TextBox" Then _
                ctlFormCtrl.Value = ""
        Next ctlFormCtrl
        
        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        '保存情報削除
        Call subSetSheet(8, objWorkSheet) 'シートオブジェクト（申込書印刷画面内容）
        objWorkSheet.Cells.ClearContents
        
        Unload Me
        frmPrintMenu.Show vbModeless
        
    End If
    
    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "BtnBack_Click" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"
    
End Sub

'「次へ」ボタン押下
Private Sub btnNext_Click()
            
    Dim wsMeisai As Worksheet


    On Error GoTo Error
    
    '申込書印刷画面の内容を保存 （2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
    If FleetTypeFlg = 1 Then
        Call subMoushikomiFormSet 'フリート
    Else
        Call subMoushikomiNonfleetFormSet 'ノンフリート
    End If

    
    'シートの保護の解除　（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
    Call subMeisaiPrtUnProtect
    
    'フリート
    If FleetTypeFlg = 1 Then
        Call subSetSheet(18, wsMeisai)      'シートオブジェクト(明細書印刷)　2018/3 ﾌﾘｰﾄ明細印刷機能追加
    'ノンフリート
    Else
        Call subSetSheet(19, wsMeisai)      'シートオブジェクト(明細書印刷（ノンフリート）)　2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    End If
    
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    blnMoushikomiflg = True
    
    
    Call subBookUnProtect 'ブックの保護を解除 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    Call subSheetVisible(True) 'シート・ブックの表示
    Call subBookProtect   'ブックの保護 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        
        
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then
        ThisWorkbook.Worksheets("明細書印刷").Activate 'フリート
    Else
        ThisWorkbook.Worksheets("明細書印刷（ノンフリート）").Activate 'ノンフリート
    End If
        
    '申込書印刷画面のエラー用リスト初期化
    wsMeisai.OLEObjects("txtErrMsg").Object.Value = ""
    wsMeisai.OLEObjects("txtErrMsg").Activate
    wsMeisai.Range("C7").Activate

    Call subMeisaiPrtProtect       'シートの保護


    Set wsMeisai = Nothing

    Unload Me

    On Error GoTo 0

    Exit Sub
Error:
    MsgBox "エラー番号:" & Err.Number & vbCrLf & _
           "エラーの種類:" & Err.Description, vbExclamation, "btnNext_Click"

End Sub


Private Sub subMoushikomiFormSet()
    Dim i                 As Integer
    Dim j                 As Integer
    Dim strSave           As String
    Dim wsMoushiSet       As Worksheet
    
    Call subSetSheet(8, wsMoushiSet)         'シートオブジェクト(申込書印刷画面内容)

    '申込書印刷画面保存内容クリア
    wsMoushiSet.Cells.ClearContents

    '申込書印刷画面内容セット
    wsMoushiSet.Cells(1, 1) = txtShokenNo.Value
    wsMoushiSet.Cells(1, 2) = txtPostNo_zen.Value
    wsMoushiSet.Cells(1, 3) = txtPostNo_kou.Value
    wsMoushiSet.Cells(1, 4) = txtKeiyakujusyo_kana.Value
    wsMoushiSet.Cells(1, 5) = txtKeiyakujusyo_kanji1.Value
    wsMoushiSet.Cells(1, 6) = txtKeiyakujusyo_kanji2.Value
    wsMoushiSet.Cells(1, 7) = txtHojin_kana.Value
    wsMoushiSet.Cells(1, 8) = txtHojin_kanji.Value
    wsMoushiSet.Cells(1, 9) = txtYakusyoku_Shimei_kana.Value
    wsMoushiSet.Cells(1, 10) = txtYakusyoku_Shimei_kanji.Value
    wsMoushiSet.Cells(1, 11) = txtTelNo_Home_zen.Value
    wsMoushiSet.Cells(1, 12) = txtTelNo_Home_chuu.Value
    wsMoushiSet.Cells(1, 13) = txtTelNo_Home_kou.Value
    wsMoushiSet.Cells(1, 14) = txtTelNo_Kinmu_zen.Value
    wsMoushiSet.Cells(1, 15) = txtTelNo_Kinmu_chuu.Value
    wsMoushiSet.Cells(1, 16) = txtTelNo_Kinmu_kou.Value
    wsMoushiSet.Cells(1, 17) = txtTelNo_Fax_zen.Value
    wsMoushiSet.Cells(1, 18) = txtTelNo_Fax_chuu.Value
    wsMoushiSet.Cells(1, 19) = txtTelNo_Fax_kou.Value
    wsMoushiSet.Cells(1, 25) = txtToriatsutenShop.Value
    wsMoushiSet.Cells(1, 26) = txtToriatsutenShop_code.Value
    wsMoushiSet.Cells(1, 27) = txtDairiShop.Value
    wsMoushiSet.Cells(1, 28) = txtDairiShop_code.Value
    wsMoushiSet.Cells(1, 29) = txtBosyuuninID.Value
    
    For i = 3 To 4
        For j = 1 To 29
            Select Case i
                Case 3
                    Select Case j
                        Case 2, 3
                            If txtPostNo_zen.Value = "" And txtPostNo_kou.Value = "" Then
                                wsMoushiSet.Cells(i, j) = ""
                            Else
                                wsMoushiSet.Cells(i, j) = txtPostNo_zen.Value & "－" & txtPostNo_kou.Value
                            End If
                        Case 4
                            strSave = txtKeiyakujusyo_kana.Value
                            wsMoushiSet.Cells(i, j) = StrConv(StrConv(strSave, vbKatakana), vbNarrow)
                        Case 7, 9
                            If txtHojin_kana.Value = "" And txtYakusyoku_Shimei_kana.Value = "" Then
                                wsMoushiSet.Cells(i, j) = ""
                            Else
                                wsMoushiSet.Cells(i, j) = txtHojin_kana.Value & " " & txtYakusyoku_Shimei_kana.Value
                            End If
                        Case 11, 12, 13
                            If txtTelNo_Home_zen.Value = "" And txtTelNo_Home_chuu.Value = "" And txtTelNo_Home_kou.Value = "" Then
                                wsMoushiSet.Cells(i, j) = ""
                            Else
                                wsMoushiSet.Cells(i, j) = txtTelNo_Home_zen.Value & "－" & txtTelNo_Home_chuu.Value & "－" & txtTelNo_Home_kou.Value
                            End If
                        Case 14, 15, 16
                            If txtTelNo_Kinmu_zen.Value = "" And txtTelNo_Kinmu_chuu.Value = "" And txtTelNo_Kinmu_kou.Value = "" Then
                                wsMoushiSet.Cells(i, j) = ""
                            Else
                                wsMoushiSet.Cells(i, j) = txtTelNo_Kinmu_zen.Value & "－" & txtTelNo_Kinmu_chuu.Value & "－" & txtTelNo_Kinmu_kou.Value
                            End If
                        Case 17, 18, 19
                            If txtTelNo_Fax_zen.Value = "" And txtTelNo_Fax_chuu.Value = "" And txtTelNo_Fax_kou.Value = "" Then
                                wsMoushiSet.Cells(i, j) = ""
                            Else
                                wsMoushiSet.Cells(i, j) = txtTelNo_Fax_zen.Value & "－" & txtTelNo_Fax_chuu.Value & "－" & txtTelNo_Fax_kou.Value
                            End If
                            
                    End Select
                Case 4
                    Select Case j
                        Case 2, 3
                            If txtPostNo_zen.Value = "" And txtPostNo_kou.Value = "" Then
                                wsMoushiSet.Cells(i, j) = ""
                            Else
                                wsMoushiSet.Cells(i, j) = txtPostNo_zen.Value & txtPostNo_kou.Value
                            End If
                        Case 7
                            strSave = txtHojin_kana.Value
                            wsMoushiSet.Cells(i, j) = StrConv(StrConv(strSave, vbKatakana), vbNarrow)
                        Case 9
                            strSave = txtYakusyoku_Shimei_kana.Value
                            wsMoushiSet.Cells(i, j) = StrConv(StrConv(strSave, vbKatakana), vbNarrow)
                    End Select
            End Select
        Next j
    Next i

    Set wsMoushiSet = Nothing

End Sub


'2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
Private Sub subMoushikomiNonfleetFormSet()

    Dim i                 As Integer
    Dim j                 As Integer
    Dim strSave           As String
    Dim wsMoushiSet       As Worksheet
    
    Call subSetSheet(8, wsMoushiSet)         'シートオブジェクト(申込書印刷画面内容)

    '申込書印刷画面内容セット
    wsMoushiSet.Cells(1, 1) = txtShokenNo.Value
    wsMoushiSet.Cells(1, 2) = txtPostNo_zen.Value
    wsMoushiSet.Cells(1, 3) = txtPostNo_kou.Value
    wsMoushiSet.Cells(1, 4) = txtKeiyakujusyo_kana.Value
    wsMoushiSet.Cells(1, 5) = txtKeiyakujusyo_kanji1.Value
    wsMoushiSet.Cells(1, 6) = txtKeiyakujusyo_kanji2.Value
    wsMoushiSet.Cells(1, 7) = txtHojin_kana.Value
    wsMoushiSet.Cells(1, 8) = txtHojin_kanji.Value
    wsMoushiSet.Cells(1, 9) = txtYakusyoku_Shimei_kana.Value
    wsMoushiSet.Cells(1, 10) = txtYakusyoku_Shimei_kanji.Value
    wsMoushiSet.Cells(1, 11) = txtTelNo_Home_zen.Value
    wsMoushiSet.Cells(1, 12) = txtTelNo_Home_chuu.Value
    wsMoushiSet.Cells(1, 13) = txtTelNo_Home_kou.Value
    wsMoushiSet.Cells(1, 14) = txtTelNo_Kinmu_zen.Value
    wsMoushiSet.Cells(1, 15) = txtTelNo_Kinmu_chuu.Value
    wsMoushiSet.Cells(1, 16) = txtTelNo_Kinmu_kou.Value
    wsMoushiSet.Cells(1, 17) = txtTelNo_Fax_zen.Value
    wsMoushiSet.Cells(1, 18) = txtTelNo_Fax_chuu.Value
    wsMoushiSet.Cells(1, 19) = txtTelNo_Fax_kou.Value
    wsMoushiSet.Cells(1, 20) = txtDantaimei.Value
    wsMoushiSet.Cells(1, 21) = txtDantaiCode.Value
    wsMoushiSet.Cells(1, 22) = txtDantaiToku.Value
    wsMoushiSet.Cells(1, 23) = txtShozoku.Value
    wsMoushiSet.Cells(1, 24) = txtSyainCode.Value
    wsMoushiSet.Cells(1, 25) = txtToriatsutenShop.Value
    wsMoushiSet.Cells(1, 26) = txtToriatsutenShop_code.Value
    wsMoushiSet.Cells(1, 27) = txtDairiShop.Value
    wsMoushiSet.Cells(1, 28) = txtDairiShop_code.Value
    wsMoushiSet.Cells(1, 29) = txtBosyuuninID.Value
    
    'シートの2行目以降に帳票で使用する書式を併せて保存
    For i = 2 To 5
        For j = 1 To 29
            Select Case i
                Case 3
                    Select Case j
                        Case 2, 3
                            If txtPostNo_zen.Value = "" And txtPostNo_kou.Value = "" Then
                                wsMoushiSet.Cells(i, j) = ""
                            Else
                                wsMoushiSet.Cells(i, j) = txtPostNo_zen.Value & "－" & txtPostNo_kou.Value
                            End If
                        Case 4
                            strSave = txtKeiyakujusyo_kana.Value
                            wsMoushiSet.Cells(i, j) = StrConv(StrConv(strSave, vbKatakana), vbNarrow)
                        Case 7, 9
                            If txtHojin_kana.Value = "" And txtYakusyoku_Shimei_kana.Value = "" Then
                                wsMoushiSet.Cells(i, j) = ""
                            Else
                                wsMoushiSet.Cells(i, j) = txtHojin_kana.Value & " " & txtYakusyoku_Shimei_kana.Value
                            End If
                        Case 11, 12, 13
                            If txtTelNo_Home_zen.Value = "" And txtTelNo_Home_chuu.Value = "" And txtTelNo_Home_kou.Value = "" Then
                                wsMoushiSet.Cells(i, j) = ""
                            Else
                                wsMoushiSet.Cells(i, j) = txtTelNo_Home_zen.Value & "－" & txtTelNo_Home_chuu.Value & "－" & txtTelNo_Home_kou.Value
                            End If
                        Case 14, 15, 16
                            If txtTelNo_Kinmu_zen.Value = "" And txtTelNo_Kinmu_chuu.Value = "" And txtTelNo_Kinmu_kou.Value = "" Then
                                wsMoushiSet.Cells(i, j) = ""
                            Else
                                wsMoushiSet.Cells(i, j) = txtTelNo_Kinmu_zen.Value & "－" & txtTelNo_Kinmu_chuu.Value & "－" & txtTelNo_Kinmu_kou.Value
                            End If
                        Case 17, 18, 19
                            If txtTelNo_Fax_zen.Value = "" And txtTelNo_Fax_chuu.Value = "" And txtTelNo_Fax_kou.Value = "" Then
                                wsMoushiSet.Cells(i, j) = ""
                            Else
                                wsMoushiSet.Cells(i, j) = txtTelNo_Fax_zen.Value & "－" & txtTelNo_Fax_chuu.Value & "－" & txtTelNo_Fax_kou.Value
                            End If
                            
                    End Select
                Case 4
                    Select Case j
                        Case 2, 3
                            If txtPostNo_zen.Value = "" And txtPostNo_kou.Value = "" Then
                                wsMoushiSet.Cells(i, j) = ""
                            Else
                                wsMoushiSet.Cells(i, j) = txtPostNo_zen.Value & txtPostNo_kou.Value
                            End If
                        Case 7
                            strSave = txtHojin_kana.Value
                            wsMoushiSet.Cells(i, j) = StrConv(StrConv(strSave, vbKatakana), vbNarrow)
                        Case 9
                            strSave = txtYakusyoku_Shimei_kana.Value
                            wsMoushiSet.Cells(i, j) = StrConv(StrConv(strSave, vbKatakana), vbNarrow)
                    End Select
            End Select
        Next j
    Next i

    Set wsMoushiSet = Nothing

End Sub

