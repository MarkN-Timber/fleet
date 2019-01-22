Attribute VB_Name = "modChohyoFunctions"
Option Explicit

Dim intConfirmMsg As Integer

'エラーダイアログ出力
Function fncTextEntryErrChk(ByVal intForm As Integer) As Boolean
    Dim i           As Integer
    Dim intTotalCar As Integer
    Dim strTemp     As String
    Dim wsTextK     As Worksheet
    Dim wsTextM     As Worksheet
    
    fncTextEntryErrChk = False
    
    strTemp = ""
    
    Call subSetSheet(6, wsTextK)         'シートオブジェクト(テキスト内容(共通))
    Call subSetSheet(7, wsTextM)         'シートオブジェクト(テキスト内容(明細))

    '総付保台数
    intTotalCar = Val(wsTextK.Cells(1, 19))

    '日付チェック
    If intForm = 2 Then

        strTemp = Format(Left(wsTextK.Cells(1, 6), 6), "####/##")
        
        If strTemp = "" Then
            strTemp = "1900/01"
        End If
        
        '保険始期日が「保険始期日の属する月の前々月1日～翌月末日」でない場合エラー
        If Date >= DateSerial(Year(strTemp), Month(strTemp) - 2, 1) And _
            Date < DateSerial(Year(strTemp), Month(strTemp) + 2, 1) Then
        Else
            intConfirmMsg = MsgBox("取り込んだファイルは印刷対象外です。" & vbCrLf & "保険始期日をご確認ください。", vbOKOnly & vbExclamation, "エラーダイアログ")
            fncTextEntryErrChk = True
        End If
    End If
    
    '合計保険料チェック
    For i = 1 To intTotalCar
        If wsTextM.Cells(i, 32) = "" Then
            intConfirmMsg = MsgBox("取り込んだファイルは印刷対象外です。" & vbCrLf & "試算が完了していません。", vbOKOnly & vbExclamation, "エラーダイアログ")
            fncTextEntryErrChk = True
            Exit For
        End If
    Next i
    
    '先日付チェック（テキストファイルに「先日付フラグ」がある場合エラー）
    If intForm = 2 Then
        If wsTextK.Cells(1, 20) = "1" Then
            intConfirmMsg = MsgBox("先日付の契約です。" & vbCrLf & "保険料を再試算してください。", vbOKOnly & vbExclamation, "エラーダイアログ")
            fncTextEntryErrChk = True
        End If
    End If
    
    '総付保台数チェック
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
'        If intTotalCar < 10 Then
'            intConfirmMsg = MsgBox("取り込んだファイルは印刷対象外です。" & vbCrLf & "総付保台数が10台未満は印刷できません。", vbOKOnly & vbExclamation, "エラーダイアログ")
'            fncTextEntryErrChk = True
'        End If
    Else  'ノンフリート
        If intTotalCar > 9 Then
            intConfirmMsg = MsgBox("取り込んだファイルは印刷対象外です。" & vbCrLf & "総付保台数が10台以上は印刷できません。", vbOKOnly & vbExclamation, "エラーダイアログ")
            fncTextEntryErrChk = True
        End If
    End If
        
    Set wsTextK = Nothing
    Set wsTextM = Nothing

End Function


'警告ダイアログ出力
Function fncTextEntryWarChk(ByVal intForm As Integer) As Boolean
    Dim i           As Integer
    Dim intTotalCar As Integer
    Dim blnWarFlg_Rng_1 As Boolean
    Dim blnWarFlg_Rng_2 As Boolean
    Dim blnWarFlg_Carno As Boolean
    Dim wsTextK     As Worksheet
    Dim wsTextM     As Worksheet
    
    Dim wsTextMP     As Worksheet
    
    fncTextEntryWarChk = False
    blnWarFlg_Rng_1 = False
    blnWarFlg_Rng_2 = False
    blnWarFlg_Carno = False
        
    Call subSetSheet(6, wsTextK)         'シートオブジェクト(テキスト内容(共通))
    Call subSetSheet(7, wsTextM)         'シートオブジェクト(テキスト内容(明細))
    
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then
        Call subSetSheet(18, wsTextMP)
    Else
        Call subSetSheet(19, wsTextMP)
    End If
    
    
    '総付保台数
    intTotalCar = Val(wsTextK.Cells(1, 19))
    
    For i = 1 To intTotalCar
        '稟議フラグチェック
        If wsTextM.Cells(i, 68) <> "" Then
            If wsTextM.Cells(i, 68) = 1 Then
                fncTextEntryWarChk = True
                blnWarFlg_Rng_1 = True
            ElseIf wsTextM.Cells(i, 68) = 2 Then
                fncTextEntryWarChk = True
                blnWarFlg_Rng_2 = True
            End If
        End If
        
        If intForm = 2 Then
            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
            '車台番号・登録番号チェック
            If FleetTypeFlg = 1 Then
                If wsTextMP.Cells(6 + i, 3) = "" Or wsTextMP.Cells(6 + i, 4) = "" Or wsTextMP.Cells(6 + i, 5) = "" Then
                    fncTextEntryWarChk = True
                    blnWarFlg_Carno = True
                End If
            Else
                If wsTextMP.Cells(18 + i, 3) = "" Or wsTextMP.Cells(18 + i, 6) = "" Or wsTextMP.Cells(18 + i, 8) = "" Then
                    fncTextEntryWarChk = True
                    blnWarFlg_Carno = True
                End If
            End If
        End If
    Next i
    
    If blnWarFlg_Rng_1 Then
        intConfirmMsg = MsgBox("社員照会対象契約です。" & vbCrLf & "注意して手続きを進めてください。", vbOKOnly & vbInformation, "警告ダイアログ")
    ElseIf blnWarFlg_Rng_2 Then
        intConfirmMsg = MsgBox("保険料試算において警告事項が" & vbCrLf & "ありました。" & vbCrLf & "契約内容を確認してください。", vbOKOnly & vbInformation, "警告ダイアログ")
    End If
    
    If blnWarFlg_Carno Then
        intConfirmMsg = MsgBox("車台番号または登録番号(カナ・漢字)に未入力のものがあります。" & vbCrLf & "申込書印刷後に補記してください。", vbOKOnly & vbInformation, "警告ダイアログ")
    End If
    
    Set wsTextK = Nothing
    Set wsTextM = Nothing

End Function


'テキスト(編集)保存
Function fncTextEdit(ByVal intKbn As Integer, ByVal intKoumoku As Integer, ByVal strContent As String, _
                            ByVal intMeisaiCnt As Integer) As Variant
'テキスト内容を編集する｡
'i = 項目インデックス
'j = 明細行
'k = 編集種類インデックス

    Dim i               As Integer
    Dim j               As Integer
    Dim h               As Integer
    Dim lngTotalHkn     As Long
    Dim lngFstTotalHkn  As Long
    Dim lngYearTotalHkn As Long
    Dim strSave         As String
    Dim wsTextK         As Worksheet
    Dim wsTextM         As Worksheet
    
    i = intKoumoku
    j = intMeisaiCnt
    h = 1
    
    lngTotalHkn = 0
    lngFstTotalHkn = 0
    lngYearTotalHkn = 0
    
    Call subSetSheet(6, wsTextK)
    Call subSetSheet(7, wsTextM)
    
    If intKbn = 1 Then
        '共通項目編集
        With wsTextK
            Select Case i
                Case 1  'レコード区分
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
                    
                Case 2  '受付区分
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
                    
                    'コード値変換
                    strSave = fncFindCode(CStr(.Cells(1, i)), "C")
                    .Cells(2, i) = strSave
                    
                Case 3  '被保険者_個人法人区分
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
                    
                    'コード値変換
                    strSave = fncFindCode(CStr(.Cells(1, i)), "G")
                    .Cells(2, i) = strSave
                    
                Case 4  '保険種類
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
                    
                    'コード値変換
                    strSave = fncFindCode(CStr(.Cells(1, i)), "K")
                    .Cells(2, i) = strSave
                    
                Case 5  'フリート・ノンフリート区分
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
                    
                    'コード値変換
                    strSave = fncFindCode(CStr(.Cells(1, i)), "O")
                    .Cells(2, i) = strSave
                    
                    '帳票用変換1
                    '帳票用変換2
                    '帳票用変換3
                    strSave = CStr(.Cells(2, i))
                    If strSave = "フリート" Then
                        .Cells(3, i) = "フリート"
'                        .Cells(5, i) = "フリート"
'                        .Cells(5, i) = "◆フリート"
                        .Cells(5, i) = "◆総付保台数：10台以上（フリート・全車両一括なし）"
                    ElseIf strSave = "全車両一括" Then
                        .Cells(3, i) = "フリート"
                        .Cells(4, i) = "有"
'                        .Cells(5, i) = "フリート（全車両一括）"
'                        .Cells(5, i) = "◆フリート（全車両一括）"
                        .Cells(5, i) = "◆総付保台数：10台以上（フリート・全車両一括あり）"
                    ElseIf strSave = "全車両連結合算" Then
                        .Cells(3, i) = "フリート"
                        .Cells(4, i) = "有"
'                        .Cells(5, i) = "フリート（全車両連結合算）"
'                        .Cells(5, i) = "◆フリート（全車両連結合算）"
                        .Cells(5, i) = "◆総付保台数：10台以上（フリート・全車両連結合算）"
                    ElseIf strSave = "ノンフリート" Then
                        .Cells(3, i) = "ノンフリート"
'                        .Cells(5, i) = "ノンフリート"
'                        .Cells(5, i) = "◆ノンフリート"
                        .Cells(5, i) = "◆総付保台数：9台以下（ノンフリート）"
                    Else
                        .Cells(3, i) = strSave
                        .Cells(4, i) = strSave
                        .Cells(5, i) = strSave
                    End If
                    
                Case 6  '保険始期日
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
                    
                    '帳票用変換1
                    strSave = fncToWareki(CStr(.Cells(1, i)), 11)
                    If strSave = CStr(.Cells(1, i)) Then
                        .Cells(3, i) = strSave
                    Else
                        strSave = Replace(strSave, Mid(Right(strSave, 6), 1, 3), Val(Mid(Right(strSave, 6), 1, 2)) & Mid(Right(strSave, 6), 3, 1))
                        .Cells(3, i) = Replace(strSave, Mid(Right(strSave, 3), 1, 3), Val(Mid(Right(strSave, 3), 1, 2)) & Mid(Right(strSave, 3), 3, 1))
                    End If
                    
                    '帳票用変換2
                    strSave = CStr(.Cells(3, i))
                    .Cells(4, i) = CStr(Format(strSave, "eemmdd"))
                    
                Case 7  '保険始期時刻区分
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
                    
                    'コード値変換
                    strSave = fncFindCode(CStr(.Cells(1, i)), "S")
                    .Cells(2, i) = strSave
                    
                Case 8  '保険始期時刻区分
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
    
                Case 9  '保険終期日
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
                    
                    '帳票用変換1
                    strSave = fncToWareki(CStr(.Cells(1, i)), 11)
                    If .Cells(1, i) = strSave Then
                        .Cells(3, i) = strSave
                    Else
                        strSave = Replace(strSave, Mid(Right(strSave, 6), 1, 3), Val(Mid(Right(strSave, 6), 1, 2)) & Mid(Right(strSave, 6), 3, 1))
                        .Cells(3, i) = Replace(strSave, Mid(Right(strSave, 3), 1, 3), Val(Mid(Right(strSave, 3), 1, 2)) & Mid(Right(strSave, 3), 3, 1))
                    End If
                    
                    '帳票用変換2
                    strSave = CStr(.Cells(3, i))
                    .Cells(4, i) = CStr(Format(strSave, "eemmdd"))
                    
                Case 10 '計算方法
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
                    
                    'コード値変換
                    strSave = fncFindCode(CStr(.Cells(1, i)), "W")
                    .Cells(2, i) = strSave
                    
                Case 11 '保険期間_年
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
                    
                Case 12 '保険期間_月
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
                    
                Case 13 '保険期間_日
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
                    
                Case 14 '払込方法
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
                    
                    'コード値変換
                    strSave = fncFindCode(.Cells(1, i), "AY")
                    .Cells(2, i) = strSave
                    
                    '帳票用変換1
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    'strSave = CStr(.Cells(1, i))
                    '.Cells(3, i) = Left(strSave, 1)
                    Select Case CStr(.Cells(1, i))
                        Case "A  "
                            strSave = "A"
                        Case "B12"
                            strSave = "W"
                        Case "D  "
                            strSave = "M"
                        Case "D06"
                            strSave = "F"
                        Case "D12"
                            strSave = "G"
                        Case "E  "
                            strSave = "M"
                        Case "E06"
                            strSave = "F"
                        Case "E12"
                            strSave = "G"
                        Case "F02"
                            strSave = "H"
                        Case "F04"
                            strSave = "H"
                        Case "F06"
                            strSave = "H"
                        Case "F08"
                            strSave = "H"
                        Case "F10"
                            strSave = "H"
                        Case "F12"
                            strSave = "H"
                        Case "G02"
                            strSave = "Y"
                        Case "G04"
                            strSave = "Y"
                        Case "G06"
                            strSave = "Y"
                        Case "G08"
                            strSave = "Y"
                        Case "G10"
                            strSave = "Y"
                        Case "G12"
                            strSave = "Y"
                        Case Else
                            strSave = ""
                    End Select
                    .Cells(3, i) = strSave
                    
                    '帳票用変換2
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
'                    strSave = CStr(Val(Right(strSave, 2)))
'                    If strSave = 0 Then
'                        .Cells(4, i) = ""
'                    Else
'                        .Cells(4, i) = strSave
'                    End If
                    strSave = CStr(.Cells(1, i))
                    If strSave Like "A*" Or strSave Like "B*" Or strSave Like "D*" Or strSave Like "E*" Then
                        strSave = ""
                    ElseIf strSave Like "F*" Or strSave Like "G*" Then
                        strSave = CStr(Right(strSave, 2))
                    Else
                        strSave = ""
                    End If
                    
                    .Cells(4, i) = strSave
                    
                Case 15 'フリート優良割引
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
                    
                Case 16 '第一種デメ割増
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
                    
                Case 17 'フリート多数割引
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
                    
                    strSave = CStr(.Cells(1, i))
                    If strSave = "2 " Then
                        'コード値変換
                        .Cells(2, i) = "有"
                        '帳票用変換1
                        .Cells(3, i) = "5％"
                        '帳票用変換2
                        .Cells(4, i) = "フリート多数割引（5%）"
                    Else
                        .Cells(2, i) = strSave
                    End If
                    
                Case 18 'フリートコード
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
                    
                Case 19 '総付保台数
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
                    
                Case 20 '先日付フラグ
                    'テキスト貼り付け
                    .Cells(1, i) = strContent
                    
                Case 21 '郵便番号
                    .Cells(1, i) = strContent
                    
                Case 22 '契約者住所（カナ）
                    .Cells(1, i) = strContent
                    
                Case 23 '契約者住所（漢字）
                    .Cells(1, i) = strContent
                    
'                Case 24 '契約者住所（漢字）
'                    .Cells(1, i) = strContent

                Case 24 '法人名（カナ）
                    .Cells(1, i) = strContent
                    
                Case 25 '法人名（漢字）
                    .Cells(1, i) = strContent
                    
                Case 26 '役職名・氏名（カナ）
                    .Cells(1, i) = strContent
                    
                Case 27 '役職名・氏名（漢字）
                    .Cells(1, i) = strContent
                    
                Case 28 '連絡先１　自宅・携帯
                    .Cells(1, i) = strContent
                    
                Case 29 '連絡先２　勤務先
                    .Cells(1, i) = strContent
                    
                Case 30 '連絡先３　ＦＡＸ
                    .Cells(1, i) = strContent
                    
                Case 31 '団体名
                    .Cells(1, i) = strContent
                    
                Case 32 '団体コード
                    .Cells(1, i) = strContent
                    
                Case 33 '団体扱に関する特約
                    .Cells(1, i) = strContent
                    
                Case 34 '所属コード
                    .Cells(1, i) = strContent
                    
                Case 35 '社員コード
                    .Cells(1, i) = strContent
                    
                Case 36 '部課コード
                    .Cells(1, i) = strContent
                    
                Case 37 '代理店コード
                    .Cells(1, i) = strContent
                    
                Case 38 '証券番号
                    .Cells(1, i) = strContent

            End Select
        End With
    ElseIf intKbn = 2 Then
        '明細項目編集
        With wsTextM
            Select Case i
                Case 1 'レコード区分
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                
                Case 2  '用途車種
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    'コード値変換
                    strSave = fncFindCode(CStr(.Cells(j, i)), "AA")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = CStr(strSave)
                
                Case 3 '車名
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                
                Case 4  '型式
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                Case 5  '仕様
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                Case 6  '初度登録年月
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '帳票用変換1
                    strSave = fncToWareki(CStr(.Cells(j, i)) & "25", 8)
                    If strSave = CStr(.Cells(j, i)) & "25" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = CStr(.Cells(j, i))
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Replace(strSave, Mid(Right(strSave, 3), 1, 3), Val(Mid(Right(strSave, 3), 1, 2)) & Mid(Right(strSave, 3), 3, 1))
                    End If
                    
                    '帳票用変換2
                    strSave = CStr(.Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i))
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    '.Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = CStr(Format(strSave, "emmdd"))
                    .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = CStr(Format(strSave, "eemm"))
                    
                Case 7  '改造・不明車
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '帳票用変換1(型式)
                    strSave = fncFindCode(CStr(.Cells(j, 7)), "AE")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, 4) = strSave
                    
                    '帳票用変換1(型式)
                    If strSave = "改造車" And CStr(.Cells(j, 4)) <> "" Then
                       .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, 4) = CStr(.Cells(j, 4)) & "ｶｲ"
                    ElseIf strSave = "不明車" Then
                       .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, 4) = "ﾌﾒｲ"
                    ElseIf strSave = "" Then
                       .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, 4) = CStr(.Cells(j, 4))
                    Else
                       .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, 4) = ""
                    End If
                    
                Case 8  '排気量
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    '帳票用変換1
                    strSave = Format(CStr(.Cells(j, i)), "0.00")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 9  '2.5リットル超ディーゼル自家用小型乗用車
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                Case 10 '被保険者_生年月日
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    '帳票用変換1
                    strSave = fncToWareki(.Cells(j, i), 11)
                    .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                    '帳票用変換2
                    .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = CStr(Format(strSave, "geemmdd"))
                    
                Case 11 'ノンフリート等級
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '帳票用変換1
                    strSave = fncFindCode(CStr(.Cells(j, i)), "AI")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    '帳票用変換2
                    strSave = CStr(.Cells(j, i))
                    If IsNumeric(strSave) Then
                        strSave = Val(strSave)
                    End If
                    .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = StrConv(strSave, vbWide) & "等級"
                    
                Case 12 '事故有適用期間
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '帳票用変換1
                    strSave = fncFindCode(CStr(.Cells(j, i)), "AM")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 13 'ノンフリート多数割引
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '帳票用変換1
                    strSave = fncFindCode(CStr(.Cells(j, i)), "AQ")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                    '帳票用変換2
                    strSave = fncFindCode(CStr(.Cells(j, i)), "AQ")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 14 '団体割増引
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    '帳票用変換1
                    strSave = Format(CStr(.Cells(j, i)), "0.00")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 15 'ゴールド免許割引
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '帳票用変換1
                    If strContent = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "有"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strContent
                    End If
                    
                    '帳票用変換2
                    If strContent = "1" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "ゴールド免許割引"
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = strContent
                    End If
                    
                Case 16 '使用目的
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '帳票用変換1
                    If strContent = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "有"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strContent
                    End If
                    
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    '帳票用変換2
                    If strContent = "1" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "日常・レジャー"
                    ElseIf strContent = "3" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "通勤・通学使用"
                    ElseIf strContent = "4" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "業務使用"
                    End If
                    
                Case 17 '沖縄
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "3 " Then
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "沖縄"
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "○"
                        '帳票用変換2
                        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "沖縄料率"
                        '帳票用変換3
                        .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = "◆沖縄料率"
                    Else
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 18 'レンタカー
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "レンタカー"
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "○"
                        '帳票用変換2
                        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                        '.Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "レンタカー料率"
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "◆レンタカー料率"
                    Else
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 19 '教習車
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "5 " Then
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "教習車"
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "○"
                        '帳票用変換2
                        '.Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "教習車料率"
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "◆教習車料率"
                    Else
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 20 'ブーム対象外
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1 " Then
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "ブーム対象外"
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "○"
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "ブーム対象外特約"
                    Else
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 21 'リースカーオープンポリシー
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    'コード値変換
                    strSave = CStr(.Cells(j, i))
                    If strSave = "80" Then
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "付帯"
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "○"
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "リースカーオープンポリシー"
                    Else
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 22 'オープンポリシー多数割引
                    'テキスト貼り付け
                    .Cells(j, i) = strContent

                    'コード値変換
                    strSave = CStr(.Cells(j, i))
                    If strSave = "93" Then
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "付帯"
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "○"
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "オープンポリシー多数割引"
                    Else
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 23 '公有・準公有
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    'コード値変換
                    strSave = fncFindCode(CStr(.Cells(j, i)), "AU")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                
                Case 24 '車両料率クラス
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                
                Case 25 '対人料率クラス
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                
                Case 26 '対物料率クラス
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                
                Case 27 '障害料率クラス
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                
                Case 28 '新車割引
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "有"
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "○"
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "新車割引"
                    Else
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 29 '特種車
                    'テキスト貼り付け
                    .Cells(j, i) = strContent

                    strSave = CStr(.Cells(j, i))
                    If strSave = "8" Then
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "特種車"
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "特種用途自動車"
                    Else
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 30 '車両下限価格
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                Case 31 '車両上限価格
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                Case 32 '合計保険料
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '帳票用変換1
                    strSave = CStr(.Cells(j, i))
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                    
                    '帳票用変換2
                    If j = Val(wsTextK.Cells(1, 19)) Then
                        For h = 1 To Val(wsTextK.Cells(1, 19))
                            lngTotalHkn = lngTotalHkn + Val(.Cells(h, i))
                        Next h
                        
                        For h = 1 To Val(wsTextK.Cells(1, 19))
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + h, i) = Format(lngTotalHkn, "#,#")
                        Next h
                        
                        '帳票用変換3
                        strSave = CStr(wsTextK.Cells(1, 14))
                        If strSave Like "F*" Or strSave Like "G*" Then
                            If lngTotalHkn = 0 Or wsTextK.Cells(1, 14) = "" Then
                            Else
                                .Cells(4 * Val(wsTextK.Cells(1, 19)) + 1, i) = _
                                    Application.WorksheetFunction.Round(lngTotalHkn / Val(Right(wsTextK.Cells(1, 14), 2)), -1)
                                .Cells(4 * Val(wsTextK.Cells(1, 19)) + 1, i) = _
                                    Format(.Cells(4 * Val(wsTextK.Cells(1, 19)) + 1, i), "#,#")
                            End If
                        Else
                        End If
                        
                    End If
                    
                Case 33 '初回保険料
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '帳票用変換1
                    strSave = CStr(.Cells(j, i))
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                                        
                    '帳票用変換2
                    If j = Val(wsTextK.Cells(1, 19)) Then
                        For h = 1 To Val(wsTextK.Cells(1, 19))
                            lngFstTotalHkn = lngFstTotalHkn + Val(.Cells(h, i))
                        Next h
                        
                        For h = 1 To Val(wsTextK.Cells(1, 19))
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + h, i) = Format(lngFstTotalHkn, "#,#")
                        Next h
                    End If
                    
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    '帳票用変換3
                    If wsTextK.Cells(1, 14).Value = "B12" Then
                        strSave = Format(.Cells(j, i).Value, "#,#")
                    Else
                        strSave = ""
                    End If
                    .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 34 '年間保険料
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '帳票用変換1
                    strSave = CStr(.Cells(j, i))
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                    
                    '帳票用変換2
                    If j = Val(wsTextK.Cells(1, 19)) Then
                        For h = 1 To Val(wsTextK.Cells(1, 19))
                            lngYearTotalHkn = lngYearTotalHkn + Val(.Cells(h, i))
                        Next h
                        
                        For h = 1 To Val(wsTextK.Cells(1, 19))
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + h, i) = Format(lngYearTotalHkn, "#,#")
                        Next h
                    End If
                    
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    '帳票用変換3
                    If wsTextK.Cells(1, 14).Value = "B12" Or wsTextK.Cells(1, 14).Value Like "D*" Or wsTextK.Cells(1, 14).Value Like "E*" Then
                        strSave = Format(.Cells(j, i).Value, "#,#")
                    Else
                        strSave = ""
                    End If
                    .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 35 '年齢条件
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    'コード値変換
                    strSave = fncFindCode(CStr(.Cells(j, i)), "BC")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                
                Case 36 '高齢運転者対象外
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                
                    '帳票用変換1
                    If strContent = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "有"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strContent
                    End If
                    
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    '35年齢条件の帳票用変換1（36高齢運転者対象外が35年齢条件より後に処理されるためここで処理する）
                    If .Cells(j, 35).Value = "" Then
                        strSave = "対象外"
                    ElseIf .Cells(j, 35).Value = "5" And .Cells(j, 36).Value = "1" Then
                        strSave = "３５歳以上限定補償（高齢者対象外）"
                    Else
                        strSave = fncFindCode(CStr(.Cells(j, 35)), "BC")
                    End If
                    .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, 35) = strSave
                    
                Case 37 '運転者限定
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    'コード値変換
                    strSave = fncFindCode(CStr(.Cells(j, i)), "BG")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    '帳票用変換1
                    If wsTextK.Cells(1, 4).Value = "7 " Then
                        strSave = "対象外"
                    ElseIf .Cells(j, i).Value = "" Then
                        strSave = "なし"
                    Else
                        strSave = fncFindCode(CStr(.Cells(j, i)), "BG")
                    End If
                    .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 38 '運転者従業員等限定特約
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "有"
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "○"
                    Else
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 39 '車両保険種類
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    'コード値変換
                    strSave = fncFindCode(CStr(.Cells(j, i)), "BK")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 40 '車両保険金額
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '帳票用変換1
                    strSave = CStr(.Cells(j, i))
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                    
                Case 41 '車両免責金額
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    'コード値変換
                    strSave = fncFindCode(CStr(.Cells(j, i)), "BO")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                    '帳票用変換1
                    If strSave = "" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Val(strSave)
                    End If
                    
                    '帳票用変換2
                    If strSave = "" Then
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    ElseIf Val(strSave) = 0 Then
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "なし"
                    Else
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = Val(strSave)
                    End If
                    
                Case 42 '代車等セット
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '帳票用変換1
                    strSave = CStr(.Cells(j, i))
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    'If strSave = "" Then
                    '    .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "無"
                    'Else
'                    If strSave = "14" Then
'                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "5,000円"
'                    ElseIf strSave = "15" Then
'                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "7,000円"
'                    ElseIf strSave = "16" Then
'                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "10,000円"
'                    Else
'                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
'                    End If
                    If strSave = "34" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "5,000円"
                    ElseIf strSave = "35" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "7,000円"
                    ElseIf strSave = "36" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "10,000円"
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If

                    
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    '帳票用変換2
'                    strSave = CStr(.Cells(j, i))
'                    If strSave = "01" Or strSave = "02" Or strSave = "03" Then
'                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "無"
'                    ElseIf strSave = "11" Or strSave = "14" Or strSave = "17" Then
'                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "5,000円"
'                    ElseIf strSave = "12" Or strSave = "15" Or strSave = "18" Then
'                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "7,000円"
'                    ElseIf strSave = "13" Or strSave = "16" Or strSave = "19" Then
'                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "10,000円"
'                    Else
'                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
'                    End If
                    strSave = CStr(.Cells(j, i))
                    If strSave = "21" Then
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "無"
                    ElseIf strSave = "31" Or strSave = "34" Then
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "5,000円"
                    ElseIf strSave = "32" Or strSave = "35" Then
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "7,000円"
                    ElseIf strSave = "33" Or strSave = "36" Then
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "10,000円"
                    Else
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If

                    
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    '帳票用変換3
'                    strSave = CStr(.Cells(j, i))
'                    If strSave = "01" Or strSave = "02" Or strSave = "11" Or strSave = "12" Or strSave = "13" Or strSave = "17" Or strSave = "18" Or strSave = "19" Then
'                        .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = "有"
'                    ElseIf strSave = "03" Or strSave = "14" Or strSave = "15" Or strSave = "16" Then
'                        .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = "無"
'                    Else
'                        .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
'                    End If
                    
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    '帳票用変換4
'                    strSave = CStr(.Cells(j, i))
'                    If strSave = "01" Or strSave = "03" Or strSave = "11" Or strSave = "12" Or strSave = "13" Then
'                        .Cells(5 * Val(wsTextK.Cells(1, 19)) + j, i) = "30万円"
'                    ElseIf strSave = "02" Or strSave = "14" Or strSave = "15" Or strSave = "16" Or strSave = "17" Or strSave = "18" Or strSave = "19" Then
'                        .Cells(5 * Val(wsTextK.Cells(1, 19)) + j, i) = "無"
'                    Else
'                        .Cells(5 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
'                    End If
                    strSave = CStr(.Cells(j, i))
                    If strSave = "21" Or strSave = "31" Or strSave = "32" Or strSave = "33" Then
                        .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = "30万円"
                    ElseIf strSave = "34" Or strSave = "35" Or strSave = "36" Then
                        .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = "無"
                    Else
                        .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 43 '車両全損臨費特約
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "2" Then
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "適用"
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "○"
                        '帳票用変換2
'                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "車両全損臨費特約"
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "車両全損臨時費用特約"
                    Else
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 44 '車両盗難対象外特約
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "適用"
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "○"
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "車両盗難対象外特約"
                    Else
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 45 '車両超過修理費用特約
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "有"
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "○"
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "車両超過修理費用特約"
                    Else
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 46 '対人無制限
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    'コード値変換
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "無制限"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                Case 47 '対人対象外
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    'コード値変換
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "対象外"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                Case 48 '対人賠償保険金額
                    'テキスト貼り付け
                    .Cells(j, i) = strContent

                    'コード値変換
                    strSave = fncFindCode(CStr(.Cells(j, i)), "CE")
                    If strSave = "対象外" Or strSave = "無制限" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "1"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                    '帳票用変換1
                    strSave = CStr(.Cells(j, i))
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                    
                Case 49 '自損事故
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "対象外"
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "○"
                        '帳票用変換2
                        If CStr(.Cells(j, 47)) = "1" Then
                            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        Else
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "対象外"
                        End If
                    Else
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '帳票用変換2
                        If CStr(.Cells(j, 47)) = "1" Then
                            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        Else
                            If strSave = "" Then
                                .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "1,500万円"
                            Else
                                .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                            End If
                        End If
                    End If
                    
                Case 50 '無保険車傷害
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "対象外"
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "○"
                        '帳票用変換2
                        If CStr(.Cells(j, 47)) = "1" Then
                            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        Else
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "対象外"
                        End If
                    ElseIf strSave = "" Then
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        
                        If CStr(.Cells(j, 47)) = "1" Then
                            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        Else
                            If .Cells(Val(wsTextK.Cells(1, 19)) + j, 46) <> "" Then
                                '帳票用変換2
                                .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "無制限"
                            ElseIf .Cells(Val(wsTextK.Cells(1, 19)) + j, 47) <> "" Then
                                '帳票用変換2
                                .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "対象外"
                            ElseIf .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, 48) <> "" Then
                                '帳票用変換2
                                .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, 48) & "万円"
                            End If
                        End If
                    Else
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 51 '対物無制限
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    'コード値変換
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "無制限"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                Case 52 '対物対象外
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    'コード値変換
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "対象外"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                Case 53 '対物賠償保険金額
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    'コード値変換
                    strSave = fncFindCode(CStr(.Cells(j, i)), "CI")
                    If strSave = "対象外" Or strSave = "無制限" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "1"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                    '帳票用変換1
                    strSave = CStr(.Cells(j, i))
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                    
                Case 54 '対物免責金額
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    'コード値変換
                    strSave = fncFindCode(CStr(.Cells(j, i)), "BS")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                    '帳票用変換1
                    If strSave = "" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Val(strSave)
                    End If
                    
                    '帳票用変換2
                    If strSave = "" Then
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    ElseIf Val(strSave) = 0 Then
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "なし"
                    Else
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = Val(strSave)
                    End If
                    
                Case 55 '対物超過修理費用特約
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "有"
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "○"
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "対物超過修理費用特約"
                    Else
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 56 '人身傷害 1名
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    'コード値変換
                    strSave = CStr(fncFindCode(.Cells(j, i), "CM"))
                    If strSave = "対象外" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "1"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                    '帳票用変換1
                    strSave = .Cells(j, i)
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                    
                Case 57 '人傷対象外
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    'コード値変換
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "対象外"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                Case 58 '人身傷害 1事故
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '帳票用変換1
                    strSave = CStr(.Cells(j, i))
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                    
                Case 59 '自動車事故特約
                    'テキスト貼り付け
                    .Cells(j, i) = strContent

                    '帳票用変換1
                    If strContent = "2" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "有"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strContent
                    End If
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    '帳票用変換2
                    If strContent = "2" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "自動車事故特約"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 60 '死亡・後遺障害保険金額　1名
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    'コード値変換
                    strSave = fncFindCode(CStr(.Cells(j, i)), "CQ")
                    If strSave = "対象外" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "1"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                    '帳票用変換1
                    strSave = CStr(.Cells(j, i))
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                    
                Case 61 '搭傷対象外
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    'コード値変換
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "対象外"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                Case 62 '自動車事故特約
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '帳票用変換1
                    strSave = CStr(.Cells(j, i))
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                
                Case 63 '日数払特約
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "2" Then
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "日数払特約"
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "○"
                        '帳票用変換2
                        If CStr(.Cells(j, 61)) = "1" Then
                            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        Else
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "日数払特約"
                        End If
                    Else
                        'コード値変換
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        
                        If CStr(.Cells(j, 61)) = "1" Then
                            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                            '帳票用変換2
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        Else
                            If strSave = "" Then
                                '帳票用変換2
                                .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "部位・症状別払"
                            Else
                                '帳票用変換2
                                .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                            End If
                        End If
                    End If
                    
                Case 64 '事業主費用特約
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    If CStr(.Cells(j, i)) = "1" Then
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "○"
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "事業主費用特約"
'                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "搭傷事業主費用特約"
                    Else
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 65 '弁護士費用特約
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '帳票用変換1
                    If CStr(.Cells(j, i)) = "1" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "○"
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "有"
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 66 'ファミリーバイク特約
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '帳票用変換1
                    'コード値変換
                    strSave = fncFindCode(CStr(.Cells(j, i)), "BW")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    '帳票用変換2
                    If CStr(.Cells(j, i)) = "1" Then
                        strSave = "賠償・自損"
                    ElseIf CStr(.Cells(j, i)) = "2" Then
                        strSave = "賠償・人身"
                    End If
                    .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 67 '個人賠償責任補償特約
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    '帳票用変換2
                    If CStr(.Cells(j, i)) = "1" Then
                        strSave = "3億円（免責金額なし）"
                    Else
                        strSave = ""
                    End If
                    .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 68 '稟議警告フラグ
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    'コード値変換
                    If CStr(.Cells(j, i)) = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "稟議エラー有"
                    ElseIf CStr(.Cells(j, i)) = "2" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "警告有"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                Case 69 '登録番号
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                Case 70 '車台番号
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                Case 71 '車検満了日
                    'テキスト貼り付け
                    .Cells(j, i) = strContent
                    
                    '帳票用変換1
                    strSave = fncToWareki(CStr(.Cells(j, i)), 11)
                    If .Cells(j, i) = strSave Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    Else
                        strSave = Replace(strSave, Mid(Right(strSave, 6), 1, 3), Val(Mid(Right(strSave, 6), 1, 2)) & Mid(Right(strSave, 6), 3, 1))
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Replace(strSave, Mid(Right(strSave, 3), 1, 3), Val(Mid(Right(strSave, 3), 1, 2)) & Mid(Right(strSave, 3), 3, 1))
                    End If
                    
                   '帳票用変換2
                     strSave = CStr(.Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i))
                    .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = CStr(Format(strSave, "eemmdd"))
                    
                Case 72 '登録番号(カナ) '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    'テキスト貼り付け
                    .Cells(j, i) = strContent

                Case 73 'ASV割引
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "○"
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "ＡＳＶ割引"
                        '帳票用変換3
                        .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = "有"
                    End If
                    
                Case 74 '車両搬送時不適用特約
                    .Cells(j, i) = strContent
                    
                    'コード値変換
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        '帳票用変換1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "○"
                        '帳票用変換2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "有"
                    End If
                Case 75 '被保険者住所（カナ）
                    .Cells(j, i) = strContent
                Case 76 '被保険者氏名（カナ）
                    .Cells(j, i) = strContent
                Case 77 '被保険者氏名（漢字）
                    .Cells(j, i) = strContent
                Case 78 '免許証の色
                    .Cells(j, i) = strContent
                Case 79 '免許証有効期限
                    .Cells(j, i) = strContent
                    
                    strSave = fncToWareki(CStr(.Cells(j, i)), 11)
                    If .Cells(j, i) = strSave Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    Else
                        strSave = Replace(strSave, Mid(Right(strSave, 6), 1, 3), Val(Mid(Right(strSave, 6), 1, 2)) & Mid(Right(strSave, 6), 3, 1))
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Replace(strSave, Mid(Right(strSave, 3), 1, 3), Val(Mid(Right(strSave, 3), 1, 2)) & Mid(Right(strSave, 3), 3, 1))
                    End If
                Case 80 '車両保険者氏名（カナ）
                    .Cells(j, i) = strContent
                Case 81 '車両保険者氏名（漢字）
                    .Cells(j, i) = strContent
                Case 82 '所有権留保またはリース等
                    .Cells(j, i) = strContent
            End Select
        End With
    End If
    
    Set wsTextK = Nothing
    Set wsTextM = Nothing

End Function

'帳票出力
Public Sub subFormSetting(ByVal intChohyo As Integer, ByVal intFormNo As Integer, ByVal intColNo As Integer, _
                            ByVal strCell As String, ByVal intType As Integer, ByVal strFront As String, _
                            ByVal strBehind As String, ByVal strDate As String, Optional ByVal intEdpKbn As Integer, _
                            Optional ByVal intEdpIndex As Integer, Optional ByVal strEdpName As String, _
                            Optional ByVal strEdpVal As String, Optional ByRef valEdpSet As Variant, _
                            Optional intSame As Integer, Optional ByRef intMeisaiCnt As Integer, _
                            Optional ByVal intPageCnt As Integer, Optional ByVal blnFstflg As Boolean)
                            
    Dim wsTextK      As Worksheet
    Dim wsTextM      As Worksheet
    Dim wsAssistSave As Worksheet
    Dim wsChohyo     As Worksheet
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    Dim strEdpValTmp As String
    Dim wsTextMP     As Worksheet
    Dim strSave      As String
    Dim intEdpIdx    As Integer
    Dim intTotalCar  As Integer
    
    Call subSetSheet(6, wsTextK)         'シートオブジェクト(テキスト内容(共通))
    Call subSetSheet(7, wsTextM)         'シートオブジェクト(テキスト内容(明細))
    Call subSetSheet(8, wsAssistSave)    'シートオブジェクト(申込書印刷画面内容)
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then
        Call subSetSheet(18, wsTextMP)
    Else
        Call subSetSheet(19, wsTextMP)
    End If
    
    Select Case intChohyo
        Case 1
            Call subSetSheet(101, wsChohyo)         'シートオブジェクト(見積書WK)
        Case 2
            Call subSetSheet(102, wsChohyo)         'シートオブジェクト(車両明細書WK)
        Case 3
            Call subSetSheet(103, wsChohyo)         'シートオブジェクト(契約申込書1枚目WK)
        Case 4
            Call subSetSheet(104, wsChohyo)         'シートオブジェクト(契約申込書2枚目WK)
        Case 5
            Call subSetSheet(105, wsChohyo)         'シートオブジェクト(明細書WK)
        Case 6
            Call subSetSheet(106, wsChohyo)         'シートオブジェクト(申込書ＥＤＰWK)
        Case 7
            Call subSetSheet(107, wsChohyo)         'シートオブジェクト(明細書ＥＤＰWK)
    End Select
    
    '総付保台数
    intTotalCar = Val(wsTextK.Cells(1, 19))
    
    intType = IIf(intType = 0, 1, intType)

    '初期化 Or EDP配列作成
    If intSame >= intTotalCar + 1 Or blnFstflg Then
        If strCell = "" Then
        Else
            If Evaluate("ISREF(" & strCell & ")") = False Then
            Else
                If intEdpIndex > 0 Then
                    If intEdpKbn <= 0 Then
                        valEdpSet(0, intEdpIndex - 1) = strEdpName
                        valEdpSet(1, intEdpIndex - 1) = strEdpVal
                        Select Case intFormNo  '取得元区分
                            Case 1  '1(IFファイル共通情報レコード
                                If CStr(wsTextK.Cells(intType, intColNo)) = "" Then
                                Else
                                    valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                End If
                            Case 2  '2(IFファイル明細情報レコード)
                                '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                                'If CStr(wsTextK.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)) = "" Then
                                If CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)) = "" Then
                                Else
                                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                                    Select Case intColNo
                                        Case 32 '合計保険料
                                            strSave = wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)
                                            If strSave = "" Then
                                            Else
                                                Select Case intType
                                                    Case 5
                                                        '払込方法により設定
                                                        strSave = wsTextK.Cells(1, 14)
                                                        If strSave Like "F*" Or strSave Like "G*" Then
                                                            valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, Replace(CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), ",", ""), strBehind)
                                                        Else
                                                            valEdpSet(2, intEdpIndex - 1) = ""
                                                        End If
                                                    Case Else
                                                        valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, Replace(CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), ",", ""), strBehind)
                                                End Select
                                            End If
                                            
                                        Case 33  '初回保険料
                                            strSave = wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)
                                            If strSave = "" Then
                                            Else
                                                Select Case intType
                                                    Case 4
                                                        '払込方法により設定
                                                        strSave = wsTextK.Cells(1, 14)
                                                        If strSave = "B12" Then
                                                            valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, Replace(CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), ",", ""), strBehind)
                                                        End If
                                                   Case Else
                                                        valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, Replace(CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), ",", ""), strBehind)
                                                End Select
                                            End If
                                        
                                        Case Else
                                            valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End Select
                                End If
                            Case 4  '4(申込書補助画面)
                                If CStr(wsAssistSave.Cells(intType, intColNo)) = "" Then
                                Else
                                    valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsAssistSave.Cells(intType, intColNo)), strBehind)
                                End If
                            Case 5  '5(その他項目)
                                Select Case intColNo
                                    Case 1  '年月日時分(現在)
                                        If intType = 1 Then
                                            valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, CStr(strDate), strBehind)
                                        ElseIf intType = 3 Then
                                            valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, fncToWareki(CStr(Left(strDate, Len(strDate) - 4)), 11), strBehind)
                                        ElseIf intType = 4 Then
                                            valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, CStr(strDate) & " - " & CStr(Format((intPageCnt + 1), "0000")), strBehind)
                                        End If
                                    Case 3  '(空白)
                                        valEdpSet(2, intEdpIndex - 1) = strFront & strBehind
                                End Select
                        End Select
                    Else
                        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                        strEdpValTmp = valEdpSet(intEdpKbn - 1, 1, intEdpIndex - 1)
                        
                        valEdpSet(intEdpKbn - 1, 0, intEdpIndex - 1) = strEdpName
                        valEdpSet(intEdpKbn - 1, 1, intEdpIndex - 1) = strEdpVal
                        Select Case intFormNo
                            Case 1  '1(IFファイル共通情報レコード
                                If CStr(wsTextK.Cells(intType, intColNo)) = "" Then
                                Else
                                    valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                End If
                            Case 2  '2(IFファイル明細情報レコード)
                                If CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)) = "" Then
                                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                                    If FleetTypeFlg = 2 Then
                                        Select Case intColNo
                                            Case 46, 47, 48
                                                '46:対人賠償（無制限）　47:対人賠償（対象外）　48:対人賠償（保険金額）
                                                valEdpSet(intEdpKbn - 1, 1, intEdpIndex - 1) = strEdpValTmp
                                            Case 51, 52, 53
                                                '51:対物賠償（無制限）　52:対物賠償（対象外）　53:対物賠償（保険金額）
                                                valEdpSet(intEdpKbn - 1, 1, intEdpIndex - 1) = strEdpValTmp
                                            Case 56, 57
                                                '56:人身傷害（1名）（保険金額）　57:人身傷害（1名）（対象外）
                                                valEdpSet(intEdpKbn - 1, 1, intEdpIndex - 1) = strEdpValTmp
                                            Case 60, 61
                                                '60:搭乗者傷害（1名）（保険金額）　61:搭乗者傷害（1名）（対象外）
                                                valEdpSet(intEdpKbn - 1, 1, intEdpIndex - 1) = strEdpValTmp
                                        End Select
                                    End If
                                Else
                                    Select Case intColNo
                                        Case 33  '初回保険料
                                            strSave = wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)
                                            If strSave = "" Then
                                            Else
                                                '払込方法により設定
                                                strSave = wsTextK.Cells(1, 14)
                                                If strSave = "B12" Then
                                                    valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, Replace(CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), ",", ""), strBehind)
                                                End If
                                            End If
                                        Case Else
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End Select
'                                    valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 4  '4(申込書補助画面)
                                If CStr(wsAssistSave.Cells(intType, intColNo)) = "" Then
                                Else
                                    valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsAssistSave.Cells(intType, intColNo)), strBehind)
                                End If
                            Case 5  '5(その他項目)
                                Select Case intColNo
                                    Case 1  '年月日時分(現在)
                                        If intType = 1 Then
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(strDate), strBehind)
                                        ElseIf intType = 3 Then
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, fncToWareki(CStr(Left(strDate, Len(strDate) - 4)), 11), strBehind)
                                        ElseIf intType = 4 Then
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(strDate) & " - " & CStr(Format((intPageCnt + 1), "0000")), strBehind)
                                        End If
                                    Case 3  '空白
                                        valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = strFront & strBehind
                                End Select
                                
                            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                            Case 6  '6(明細書印刷画面)
                            
                                If FleetTypeFlg = 1 Then
                                
                                    'フリート
                                
                                    Select Case intColNo
                                        Case 6 '登録番号（漢字）
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(6 + intSame, 3)), strBehind)
                                        
                                        Case 7 '登録番号（カナ）
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(6 + intSame, 4)), strBehind)
                                        
                                        Case 8 '車台番号
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(6 + intSame, 5)), strBehind)
                                        
                                        Case 9 '車検満了日
                                                valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, Format(CStr(wsTextMP.Cells(6 + intSame, 6)), "eemmdd"), strBehind)
                                    End Select
                                    
                                
                                ElseIf FleetTypeFlg = 2 Then
                                
                                    'ノンフリート
                            
                                    Select Case intColNo
                            
                                        Case 1 '被保険者住所（カナ）
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, StrConv(CStr(wsTextMP.Cells(6 + intSame, 3)), vbKatakana + vbNarrow), strBehind)
                                            
                                        Case 2 '被保険者氏名（カナ）
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, StrConv(CStr(wsTextMP.Cells(6 + intSame, 8)), vbKatakana + vbNarrow), strBehind)
                                            
                                        Case 3 '被保険者氏名（漢字）
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(6 + intSame, 11)), strBehind)
                                        
                                        Case 4 '免許証の色
                                            If CStr(wsTextMP.Cells(6 + intSame, 24)) = "グリーン等" Then
                                                valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, "1", strBehind)
                                            ElseIf CStr(wsTextMP.Cells(6 + intSame, 24)) = "ブルー" Then
                                                valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, "2", strBehind)
                                            ElseIf CStr(wsTextMP.Cells(6 + intSame, 24)) = "ゴールド" Then
                                                valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, "3", strBehind)
                                            Else
                                                valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(6 + intSame, 24)), strBehind)
                                            End If
                                        
                                        Case 5 '免許証有効期限
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, Format(CStr(wsTextMP.Cells(6 + intSame, 27)), "eemmdd"), strBehind)
                                        
                                        Case 6 '登録番号（漢字）
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(18 + intSame, 3)), strBehind)
                                        
                                        Case 7 '登録番号（カナ）
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(18 + intSame, 6)), strBehind)
                                        
                                        Case 8 '車台番号
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(18 + intSame, 8)), strBehind)
                                        
                                        Case 9 '車検満了日
                                                valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, Format(CStr(wsTextMP.Cells(18 + intSame, 9)), "eemmdd"), strBehind)
                                        
                                        Case 10 '車両所有者氏名（カナ）
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, StrConv(CStr(wsTextMP.Cells(18 + intSame, 16)), vbKatakana + vbNarrow), strBehind)
                                        
                                        Case 11 '車両所有者氏名（漢字）
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(18 + intSame, 24)), strBehind)
                                        
                                        Case 12 '所有権留保またはリース等
                                            If CStr(wsTextMP.Cells(18 + intSame, 31)) = "所有権留保" Then
                                                valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, "1", strBehind)
                                            Else
                                                valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(18 + intSame, 31)), strBehind)
                                            End If
                                        
                                        Case 13 '証券番号
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 3)), strBehind)
                                        
                                        Case 14 '明細番号
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 5)), strBehind)
                                        
                                        Case 15 '前契約等級
                                            'valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 6)), strBehind)
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = Replace(fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 6)), strBehind), "等級", "")
                                        
                                        Case 16 '前契約事故有適用期間
                                            'valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 7)), strBehind)
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = Replace(fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 7)), strBehind), "年", "")
                                        
                                        Case 17 '前契約保険会社
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 8)), strBehind)
                                        
                                        Case 18 'コード
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 9)), strBehind)
                                        
                                        Case 19 '保険始期日
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, Format(CStr(wsTextMP.Cells(30 + intSame, 10)), "eemmdd"), strBehind)
                                        
                                        Case 20 '保険終期日
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, Format(CStr(wsTextMP.Cells(30 + intSame, 17)), "eemmdd"), strBehind)
                                        
                                        Case 21 '3等級ダウン事故
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 24)), strBehind)
                                        
                                        Case 22 '1等級ダウン事故
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 27)), strBehind)
                                    End Select
                                    
                            End If
                            
                        End Select
                        
                    End If
                    
                End If
                wsChohyo.Range(strCell) = ""
            End If
        End If
    Else
    'セルに値セット
        If (intChohyo = 6 Or intChohyo = 7) And intEdpIndex > 0 Then  '申込書ＥＤＰWK　or　明細書ＥＤＰWK
        Else
            '取得元区分
            Select Case intFormNo
                Case 1  '1(IFファイル共通情報レコード)
                    If strCell = "" Then
                    Else
                        wsChohyo.Range(strCell).NumberFormatLocal = "@"
                        Select Case intColNo  '項目No
                            Case 5  '5：フリート・ノンフリート区分
                                '帳票
                                Select Case intChohyo
                                    Case 3  '3：契約申込書1枚目WK
                                        If wsTextK.Cells(intType, intColNo) = "" Then
                                        Else
                                            'strCell：セル番号、intType：編集No
                                            wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        End If
                                    Case Else
                                        wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                End Select
                            Case 7  '保険始期時刻区分
                                Select Case intChohyo
                                    Case 3  '3：契約申込書1枚目WK
                                        If wsTextK.Cells(intType, intColNo) = "" Then
                                        Else
                                            wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        End If
                                    Case Else
                                        wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                End Select
                            Case 15  'フリート優良割引
                                Select Case intChohyo
                                    Case 3  '3：契約申込書1枚目WK
                                        If wsTextK.Cells(intType, intColNo) = "" Then
                                        Else
                                            'strCell：セル番号、intType：編集No
                                            wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        End If
                                    Case 5  '5：明細書WK
                                        If wsTextK.Cells(intType, intColNo) = "" Then
                                        Else
                                            wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        End If
                                    Case Else
                                        If wsTextK.Cells(intType, 16) = "" Then
                                            wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        Else
                                        End If
                                End Select
                            Case 16  '第一種デメ割増
                                Select Case intChohyo
                                    Case 3  '3：契約申込書1枚目WK
                                        If wsTextK.Cells(intType, intColNo) = "" Then
                                        Else
                                            wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        End If
                                    Case 5  '5：明細書WK
                                        If wsTextK.Cells(intType, intColNo) = "" Then
                                        Else
                                            wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        End If
                                    Case Else
                                        If wsTextK.Cells(intType, 15) = "" Then
                                            wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        Else
                                        End If
                                End Select
                            Case 17  'フリート多数割引
                                Select Case intChohyo
                                    Case 3  '3：契約申込書1枚目WK
                                        If wsTextK.Cells(intType, intColNo) = "" Then
                                        Else
                                            wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        End If
                                    Case 5  '5：明細書WK
                                        If wsTextK.Cells(intType, intColNo) = "" Then
                                        Else
                                            wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        End If
                                    Case Else
                                        wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                End Select
                            Case 18  'フリートコード
                                Select Case intChohyo
                                    Case 3
                                        If wsTextK.Cells(intType, intColNo) = "" Then
                                        Else
                                            wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        End If
                                    Case Else
                                        wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                End Select
                        Case Else
                            wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                        End Select
                    End If
                    
                Case 2  '2(IFファイル明細情報レコード)
                    If strCell = "" Then
                    Else
                        wsChohyo.Range(strCell).NumberFormatLocal = "@"
                        Select Case intColNo  '項目No
                            Case 8  '排気量
                                If intChohyo = 5 Then  '5：明細書WK
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                        wsChohyo.Range(strCell) = ""
                                    Else
                                        wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                            Case 11  'ノンフリート等級
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
'                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                        wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                            Case 12  '事故有適用期間
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                            Case 13  'ノンフリート多数割引
                                If intChohyo = 3 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                ElseIf intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                                
                            Case 14 '団体割増引
                                If intChohyo = 3 Or intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells(intType, intColNo)), strBehind)
                                End If

                            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                            Case 15  'ゴールド免許割引
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 17  '沖縄
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 18  'レンタカー
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 19  '教習車
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 20  'ブーム対象外
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 21  'リースカーオープンポリシー
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 22  'オープンポリシー多数割引
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 23  '公有・準公有車
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 28  '新車割引
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 32  '合計保険料
                                strSave = wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)
                                If strSave = "" Then
                                Else
                                    Select Case intType
                                        Case 5
                                            '払込方法により設定
                                            strSave = wsTextK.Cells(1, 14)
                                            If strSave Like "F*" Or strSave Like "G*" Then
                                                If intChohyo = 6 Then
                                                    wsChohyo.Range(strCell) = fncUnionString(strFront, Replace(CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), ",", ""), strBehind)
                                                Else
                                                   wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                                End If
                                            Else
                                                wsChohyo.Range(strCell) = ""
                                            End If
                                        Case Else
                                            If intChohyo = 6 Then
                                                wsChohyo.Range(strCell) = fncUnionString(strFront, Replace(CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), ",", ""), strBehind)
                                            Else
                                                wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                            End If
                                    End Select
                                End If
                            Case 33  '初回保険料
                                strSave = wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)
                                If strSave = "" Then
                                Else
                                    Select Case intType
                                        Case 4
                                            '払込方法により設定
                                            strSave = wsTextK.Cells(1, 14)
                                            If strSave = "B12" Then
                                                If intChohyo = 6 Then
                                                    wsChohyo.Range(strCell) = fncUnionString(strFront, Replace(CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), ",", ""), strBehind)
                                                Else
                                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                                End If
                                            End If
                                       Case Else
                                        If intChohyo = 6 Then
                                            wsChohyo.Range(strCell) = fncUnionString(strFront, Replace(CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), ",", ""), strBehind)
                                        Else
                                            wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                        End If
                                    End Select
                                End If
                            Case 34  '年間保険料
                                strSave = wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)
                                If strSave = "" Then
                                Else
                                    Select Case intType
                                        Case 4
                                            '払込方法により設定
                                            strSave = wsTextK.Cells(1, 14)
                                            If strSave Like "A*" Or strSave Like "F*" Or strSave Like "G*" Then
                                                wsChohyo.Range(strCell) = ""
                                            Else
                                                wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                            End If
                                        Case Else
                                            wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End Select
                                End If
                            
                            Case 43  '車両全損時臨費特約
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                        wsChohyo.Range(strCell) = ""
                                    Else
                                        wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 46  '対人無制限
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 47)) = "" And _
                                  CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 48)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 47  '対人対象外
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 46)) = "" And _
                                  CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 48)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 48  '対人賠償保険金額
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 46)) = "" And _
                                  CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 47)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 51  '対物無制限
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 52)) = "" And _
                                  CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 53)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 52  '対物対象外
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 51)) = "" And _
                                  CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 53)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 53  '対物賠償保険金額
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 51)) = "" And _
                                  CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 52)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 54  '対物免責金額
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                        wsChohyo.Range(strCell) = ""
                                    Else
                                        wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 56  '人身傷害 1名
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 57)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 57  '人傷対象外
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 56)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 58  '人身傷害 1事故
                                strSave = CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo))
                                If strSave = "" Then
                                    wsChohyo.Range(strCell) = ""
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, strSave, strBehind)
                                End If
                            Case 60  '死亡・後遺障害保険金額 1名
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 61)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 61  '搭傷対象外
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 60)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 62  '死亡・後遺障害保険金額 1事故
                                strSave = CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo))
                                If strSave = "" Then
                                    wsChohyo.Range(strCell) = ""
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, strSave, strBehind)
                                End If
                            Case 73  'ASV割引
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case Else
                                wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                        End Select
                    End If
                Case 3  '3(見積書補助画面)
                    If strCell = "" Then
                    Else
                        wsChohyo.Range(strCell).NumberFormatLocal = "@"  '@：表示形式指定→文字列
                        If wsAssistSave.Cells(intType, intColNo) = "" Then
                            wsChohyo.Range(strCell) = ""
                        Else
                            wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsAssistSave.Cells(intType, intColNo)), strBehind)
                        End If
                    End If
                    
                Case 4  '4(申込書補助画面)
                    If strCell = "" Then
                    Else
                        wsChohyo.Range(strCell).NumberFormatLocal = "@"  '@：表示形式指定→文字列
                        If wsAssistSave.Cells(intType, intColNo) = "" Then
                            wsChohyo.Range(strCell) = ""
                        Else
                            wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsAssistSave.Cells(intType, intColNo)), strBehind)
                        End If
                    End If
                    
                Case 5  '5(その他項目)
                    Select Case intColNo
                        Case 1  '年月日時分(現在)
                            If strCell = "" Then
                            Else
                                wsChohyo.Range(strCell).NumberFormatLocal = "@"  '@：表示形式指定→文字列
                                If intType = 1 Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(strDate), strBehind)
                                ElseIf intType = 3 Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, fncToWareki(CStr(Left(strDate, Len(strDate) - 4)), 11), strBehind)
                                ElseIf intType = 4 Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(strDate) & " - " & CStr(Format((intPageCnt + 1), "0000")), strBehind)
                                End If
                            End If
                        Case 2  '明細番号
                            If strCell = "" Then
                            Else
                                wsChohyo.Range(strCell).NumberFormatLocal = "@"  '@：表示形式指定→文字列
                                If intSame > intTotalCar Then
                                    wsChohyo.Range(strCell) = ""
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(Format(intMeisaiCnt, "0000")), strBehind)
                                End If
                            End If
                            intMeisaiCnt = intMeisaiCnt + 1
                            
                        Case 3  '空白
                            If strCell = "" Then
                            Else
                                wsChohyo.Range(strCell).NumberFormatLocal = "@"  '@：表示形式指定→文字列
                                wsChohyo.Range(strCell) = strFront & "" & strBehind
                            End If
                            
                    End Select
                    
                '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                Case 6  '6(明細書印刷画面)
                
                    If strCell = "" Then
                    Else
                    
                        If FleetTypeFlg = 1 Then
                        
                            'フリート
                        
                            Select Case intColNo
                                Case 1 '登録番号（漢字）
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(6 + intSame, 3)), strBehind)
                                
                                Case 2 '登録番号（カナ）
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(6 + intSame, 4)), strBehind)
                                
                                Case 3 '車台番号
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(6 + intSame, 5)), strBehind)
                                
                                Case 4 '車検満了日
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(6 + intSame, 6)), strBehind)
                            End Select
                        
                        
                        ElseIf FleetTypeFlg = 2 Then
                        
                            'ノンフリート
                        
                            Select Case intColNo
                            
                                Case 1 '被保険者住所（カナ）
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(5 + intMeisaiCnt, 3)), strBehind)
                                    
                                Case 2 '被保険者氏名（カナ）
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(5 + intMeisaiCnt, 8)), strBehind)
                                    
                                Case 3 '被保険者氏名（漢字）
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(5 + intMeisaiCnt, 11)), strBehind)
                                
                                Case 4 '免許証の色
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(5 + intMeisaiCnt, 24)), strBehind)
                                
                                Case 5 '免許証有効期限
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(5 + intMeisaiCnt, 27)), strBehind)
                                
                                Case 6 '登録番号（漢字）
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(17 + intMeisaiCnt, 3)), strBehind)
                                
                                Case 7 '登録番号（カナ）
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(17 + intMeisaiCnt, 6)), strBehind)
                                
                                Case 8 '車台番号
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(17 + intMeisaiCnt, 8)), strBehind)
                                
                                Case 9 '車検満了日
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(17 + intMeisaiCnt, 9)), strBehind)
                                
                                Case 10 '車両所有者氏名（カナ）
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(17 + intMeisaiCnt, 16)), strBehind)
                                
                                Case 11 '車両所有者氏名（漢字）
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(17 + intMeisaiCnt, 24)), strBehind)
                                
                                Case 12 '所有権留保またはリース等
                                    If intChohyo = 5 Then
                                        If CStr(wsTextMP.Cells(17 + intMeisaiCnt, 31)) = "所有権留保" Then
                                            wsChohyo.Range(strCell) = fncUnionString(strFront, "所有権留保またはリース等", strBehind)
                                        Else
                                            wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(17 + intMeisaiCnt, 31)), strBehind)
                                        End If
                                    Else
                                        wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(17 + intMeisaiCnt, 31)), strBehind)
                                    End If
                                
                                Case 13 '証券番号
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 3)), strBehind)
                                
                                Case 14 '明細番号
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 5)), strBehind)
                                
                                Case 15 '前契約等級
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 6)), strBehind)
                                
                                Case 16 '前契約事故有適用期間
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 7)), strBehind)
                                
                                Case 17 '前契約保険会社
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 8)), strBehind)
                                
                                Case 18 'コード
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 9)), strBehind)
                                
                                Case 19 '保険始期日
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 10)), strBehind)
                                
                                Case 20 '保険終期日
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 17)), strBehind)
                                
                                Case 21 '3等級ダウン事故
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 24)), strBehind)
                                
                                Case 22 '1等級ダウン事故
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 27)), strBehind)
                            
                            End Select
                        End If
                        
                    End If
                    
            End Select
        End If
    End If
    
    Set wsTextK = Nothing
    Set wsTextM = Nothing
    Set wsChohyo = Nothing
    Set wsAssistSave = Nothing
    
End Sub

Public Sub subFormEDPEdit(ByVal intForm As Integer, ByVal varEditEdp As Variant, ByVal intPageCnt As Integer, ByVal strStartRow As String)
    
    Dim i            As Integer
    Dim j            As Integer
    Dim intRowCnt    As Integer
    Dim intStartRow  As Integer
    Dim strSave      As String
    Dim wsChohyo     As Worksheet
    
    Select Case intForm
        Case 1
            Call subSetSheet(106, wsChohyo)         'シートオブジェクト(申込書ＥＤＰWK)
        Case 2
            Call subSetSheet(107, wsChohyo)         'シートオブジェクト(明細書ＥＤＰWK)
    End Select

    intRowCnt = 0
    intStartRow = wsChohyo.Range(strStartRow).Row + (85 * intPageCnt)

    i = intStartRow

    If intForm = 1 Then
        Do Until wsChohyo.Cells(i, 2).MergeArea(1) = ""
            If wsChohyo.Cells(i, 17).MergeArea(1) = "" Then
                wsChohyo.Cells(i, 2).MergeArea(1) = ""
                wsChohyo.Cells(i, 14).MergeArea(1) = ""
            End If

            i = i + 1
        Loop

        For i = 0 To UBound(varEditEdp, 2)
            If IsEmpty(varEditEdp(2, i)) Then
            Else
                If varEditEdp(2, i) = "" Then
                Else
                    wsChohyo.Cells(intStartRow + intRowCnt, 2) = varEditEdp(0, i)
                    wsChohyo.Cells(intStartRow + intRowCnt, 14) = varEditEdp(1, i)
                    wsChohyo.Cells(intStartRow + intRowCnt, 17) = varEditEdp(2, i)

                    intRowCnt = intRowCnt + 1
                End If
            End If

        Next i
    Else

'        Do Until wsChohyo.Cells(i, 2).MergeArea(1) = ""
'            If wsChohyo.Cells(i, 17).MergeArea(1) = "" Then
'                wsChohyo.Cells(i, 2).MergeArea(1) = ""
'                wsChohyo.Cells(i, 14).MergeArea(1) = ""
'            End If
'            i = i + 1
'        Loop
'
'        For j = 0 To UBound(varEditEdp, 3)
'            If IsEmpty(varEditEdp(0, 2, j)) then
'            Else
'                If varEditEdp(0, 2, j) = "" Then
'                Else
'                    wsChohyo.Cells(intStartRow + intRowCnt, 2) = varEditEdp(0, 0, j)
'                    wsChohyo.Cells(intStartRow + intRowCnt, 14) = varEditEdp(0, 1, j)
'                    wsChohyo.Cells(intStartRow + intRowCnt, 17) = varEditEdp(0, 2, j)
'
'                    intRowCnt = intRowCnt + 1
'                End If
'            End If
'
'        Next j
        Do Until wsChohyo.Cells(i, 2).MergeArea(1) = ""
            If wsChohyo.Cells(i, 13).MergeArea(1) = "" Then
                wsChohyo.Cells(i, 2).MergeArea(1) = ""
                wsChohyo.Cells(i, 10).MergeArea(1) = ""
            End If
            i = i + 1
        Loop
        i = intStartRow

        Do Until wsChohyo.Cells(i, 34).MergeArea(1) = ""
            If wsChohyo.Cells(i, 45).MergeArea(1) = "" Then
                wsChohyo.Cells(i, 34).MergeArea(1) = ""
                wsChohyo.Cells(i, 42).MergeArea(1) = ""
            End If
            i = i + 1
        Loop

        For j = 0 To UBound(varEditEdp, 3)
            If IsEmpty(varEditEdp(0, 2, j)) Then
            Else
                If varEditEdp(0, 2, j) = "" Then
                Else
                    If intRowCnt < 66 Then
                        wsChohyo.Cells(intStartRow + intRowCnt, 2) = varEditEdp(0, 0, j)
                        wsChohyo.Cells(intStartRow + intRowCnt, 10) = varEditEdp(0, 1, j)
                        wsChohyo.Cells(intStartRow + intRowCnt, 13) = varEditEdp(0, 2, j)

                        intRowCnt = intRowCnt + 1
                    Else
                        wsChohyo.Cells(intStartRow + intRowCnt - 66, 34) = varEditEdp(0, 0, j)
                        wsChohyo.Cells(intStartRow + intRowCnt - 66, 42) = varEditEdp(0, 1, j)
                        wsChohyo.Cells(intStartRow + intRowCnt - 66, 45) = varEditEdp(0, 2, j)

                        intRowCnt = intRowCnt + 1
                    End If
                End If
            End If

        Next j

    End If

    If intForm = 2 Then
        intRowCnt = 0
        intStartRow = 78 + (85 * intPageCnt)
        i = intStartRow

        Do Until wsChohyo.Cells(i, 2).MergeArea(1) = ""
            If wsChohyo.Cells(i, 10).MergeArea(1) = "" Then
                wsChohyo.Cells(i, 2).MergeArea(1) = ""
                wsChohyo.Cells(i, 7).MergeArea(1) = ""
            End If
            i = i + 1
        Loop

        i = intStartRow

        Do Until wsChohyo.Cells(i, 20).MergeArea(1) = ""
            If wsChohyo.Cells(i, 28).MergeArea(1) = "" Then
                wsChohyo.Cells(i, 20).MergeArea(1) = ""
                wsChohyo.Cells(i, 25).MergeArea(1) = ""
            End If
            i = i + 1
        Loop

        i = intStartRow
        Do Until wsChohyo.Cells(i, 38).MergeArea(1) = ""
            If wsChohyo.Cells(i, 43).MergeArea(1) = "" Then
                wsChohyo.Cells(i, 38).MergeArea(1) = ""
                wsChohyo.Cells(i, 46).MergeArea(1) = ""
            End If
            i = i + 1
        Loop

        i = intStartRow

        For i = 1 To UBound(varEditEdp, 1)
            intRowCnt = 0

            For j = 0 To UBound(varEditEdp, 3)
                If IsEmpty(varEditEdp(i, 2, j)) Then
                Else
                    If varEditEdp(i, 2, j) = "" Then
                    Else
                        Select Case i
                            Case 1
                                wsChohyo.Cells(intStartRow + intRowCnt, 2) = varEditEdp(i, 0, j)
                                wsChohyo.Cells(intStartRow + intRowCnt, 7) = varEditEdp(i, 1, j)
                                wsChohyo.Cells(intStartRow + intRowCnt, 10) = varEditEdp(i, 2, j)
                            Case 2
                                wsChohyo.Cells(intStartRow + intRowCnt, 20) = varEditEdp(i, 0, j)
                                wsChohyo.Cells(intStartRow + intRowCnt, 25) = varEditEdp(i, 1, j)
                                wsChohyo.Cells(intStartRow + intRowCnt, 28) = varEditEdp(i, 2, j)
                            Case 3
                                wsChohyo.Cells(intStartRow + intRowCnt, 38) = varEditEdp(i, 0, j)
                                wsChohyo.Cells(intStartRow + intRowCnt, 43) = varEditEdp(i, 1, j)
                                wsChohyo.Cells(intStartRow + intRowCnt, 46) = varEditEdp(i, 2, j)
                        End Select
                        intRowCnt = intRowCnt + 1
                    End If
                End If

            Next j

        Next i

    End If

    Set wsChohyo = Nothing

End Sub


'2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
Private Function fncUnionString(ByVal strFront As String, ByVal strValue As String, ByVal strBehind As String)

    If strValue = "" Then
        fncUnionString = strValue
    Else
        fncUnionString = strFront & strValue & strBehind
    End If

End Function

