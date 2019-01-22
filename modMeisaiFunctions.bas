Attribute VB_Name = "modMeisaiFunctions"
Option Explicit

''明細入力画面の処理群

Public Sub subMeisaiAdd(ByRef strAddrow As String, ByRef strAddMode As String)
'関数名：subMeisaiAdd
'内容　：明細行追加
'引数　：
'        strAddrow          = 追加行数
'        strAddMode         = 遷移元画面の判定("1" = 明細入力シート/ "2" = 明細入力シート以外)

    Dim i          As Integer               'ループ用カウント
    Dim blnCopyflg As Boolean               'コピーフラグ
    Dim strAllCnt  As String                '総付保台数
    Dim strStartAddrow As String            '追加開始行
    Dim strAddRowRange As String            '追加セル番号
    Dim strAddCnt      As String            '追加カウント
    Dim strLastCon As String                '最終行のコントロールの番号
    Dim strCopyRow As String                'コピー行
    Dim strCopyCon As String                'コピーコントロール
    Dim strCopyConValue As String           'コピーコントロール値
    Dim rngChkAll As Range                  'ループ用レンジ
    
    Dim wsMeisai As Worksheet           '明細入力ワークシート
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(1, wsMeisai)       'シートオブジェクト(明細入力)
    Else
        Call subSetSheet(17, wsMeisai)      'シートオブジェクト(明細入力（ノンフリート）)
    End If
    
    Dim objSouhuho As Object
    Set objSouhuho = wsMeisai.OLEObjects("txtSouhuho").Object
    
    strAllCnt = Left(objSouhuho.Value, Len(objSouhuho.Value) - 2) '総付保台数の値を設定
    strStartAddrow = 20 + Val(strAllCnt)                          '追加する行番号を設定
    Application.ScreenUpdating = False                            '描画停止
    Application.EnableEvents = False                              'イベント無効
    
    '追加カウントをリセット
    strAddCnt = 0
    
    '明細チェックにチェックが入っているか確認
    For Each rngChkAll In wsMeisai.Range("A21:" & wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp).Address)
        If InStr(rngChkAll.Value, "True") > 0 Then
            blnCopyflg = True
            Exit For
        End If
    Next rngChkAll
    
    Set rngChkAll = Nothing
    
    '最終行のコントロールの番号を設定
    strLastCon = Left(wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp), InStr(wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp), "/") - 1)
    
    '追加処理
    If blnCopyflg = True And strAddMode = "1" Then

        'コピーフラグ有かつ、明細入力画面から遷移してきた場合、明細行のコピー追加
        Set rngChkAll = Nothing
        For Each rngChkAll In wsMeisai.Range("A21:" & wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp).Address)

            If InStr(rngChkAll.Value, "True") > 0 Then

                '追加するセル番号を設定
                strAddRowRange = strStartAddrow + 1

                '行をコピー
                strCopyRow = rngChkAll.Row
                '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                If FleetTypeFlg = 1 Then  'フリート（最終はＡＸ列）
                    wsMeisai.Range("A" & strCopyRow & ":AX" & strCopyRow).Copy
                    wsMeisai.Range("A" & strStartAddrow + 1 & ":AX" & strStartAddrow + 1).PasteSpecial
                    wsMeisai.Rows(strStartAddrow + 1).EntireRow.RowHeight = 26.25
                Else
                    'ノンフリート（最終はＢＨ列）
                    wsMeisai.Range("A" & strCopyRow & ":BH" & strCopyRow).Copy
                    wsMeisai.Range("A" & strStartAddrow + 1 & ":BH" & strStartAddrow + 1).PasteSpecial
                    wsMeisai.Rows(strStartAddrow + 1).EntireRow.RowHeight = 26.25
                End If

                '明細Noを更新
                wsMeisai.Range("B" & strStartAddrow + 1) = strStartAddrow - 19

                'コピーするコントロールの番号を取得
                strCopyCon = Left(wsMeisai.Cells(strStartAddrow, 1), InStr(wsMeisai.Cells(strStartAddrow, 1), "/") - 1)

                '明細チェックボックスをコピー
                Call subConAdd("A" & strAddRowRange, "chkMeisai", "", strLastCon, False)
                '明細チェックのセルにコントロールの番号/チェック状態を設定
                Range("A" & strAddRowRange) = Val(strLastCon) + 1 & "/False"

                '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                If FleetTypeFlg = 1 Then  'フリート
                    '選択ボタンをコピー（Ｗ列）
                    Call subConAdd("W" & strAddRowRange, "btnOtherRate", "", strLastCon, True)
                Else
                    '選択ボタンをコピー（ＡＥ列）
                    Call subConAdd("AE" & strAddRowRange, "btnOtherRate", "", strLastCon, True)
                End If

                '最終行のコントロールの番号を更新
                strLastCon = Val(strLastCon) + 1
                '追加カウント
                strAddCnt = strAddCnt + 1
                '追加行を一つ下にずらす
                strStartAddrow = strStartAddrow + 1
                
                '追加した行数が追加予定の行数に達した場合、ループを抜ける
                If strAddrow = strAddCnt Then
                    Exit For
                End If
                
            End If
            
        Next rngChkAll
        
        Set rngChkAll = Nothing
        
        '追加行数 > チェック明細行数 の場合、チェックされている一番最後の行を不足分コピー
        Do While Val(strAddrow) > Val(strAddCnt)
            
            '追加するセル番号を設定
            strAddRowRange = strStartAddrow + 1
            
            '行をコピー
            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
            If FleetTypeFlg = 1 Then  'フリート（最終はＡＸ列）
                wsMeisai.Range("A" & strCopyRow & ":AX" & strCopyRow).Copy
                wsMeisai.Range("A" & strStartAddrow + 1 & ":AX" & strStartAddrow + 1).PasteSpecial
                wsMeisai.Rows(strStartAddrow + 1).EntireRow.RowHeight = 26.25
            Else
                'ノンフリート（最終はＢＨ列）
                wsMeisai.Range("A" & strCopyRow & ":BH" & strCopyRow).Copy
                wsMeisai.Range("A" & strStartAddrow + 1 & ":BH" & strStartAddrow + 1).PasteSpecial
                wsMeisai.Rows(strStartAddrow + 1).EntireRow.RowHeight = 26.25
            End If

            '明細Noを更新
            wsMeisai.Range("B" & strStartAddrow + 1) = strStartAddrow - 19
            
            '明細チェックボックスをコピー
            Call subConAdd("A" & strAddRowRange, "chkMeisai", "", strLastCon, False)
            '明細チェックのセルにコントロールの番号/チェック状態を設定
            wsMeisai.Range("A" & strAddRowRange) = Val(strLastCon) + 1 & "/False"
            
            '選択ボタンをコピー
            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
            If FleetTypeFlg = 1 Then  'フリート（Ｗ列）
                Call subConAdd("W" & strAddRowRange, "btnOtherRate", "", strLastCon, True)
            Else
                'ノンフリート（ＡＥ列）
                Call subConAdd("AE" & strAddRowRange, "btnOtherRate", "", strLastCon, True)
            End If
            
            '最終行のコントロールの番号を更新
            strLastCon = Val(strLastCon) + 1
            '追加カウント
            strAddCnt = strAddCnt + 1
            '追加行を一つ下にずらす
            strStartAddrow = strStartAddrow + 1
            
        Loop
        
    Else
        
        'コピーフラグ無、あるいは、明細入力画面以外から遷移してきた場合、明細行の新規追加
        Do While Val(strAddrow) > Val(strAddCnt)
            
            '追加するセル番号を設定
            strAddRowRange = strStartAddrow + 1
            
            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
            If FleetTypeFlg = 1 Then  'フリート（最終はＡＸ列）
                '行を新規追加
                wsMeisai.Range("A21:AX21").Copy
                '既存行の書式をコピー
                wsMeisai.Range("A" & strStartAddrow + 1 & ":AX" & strStartAddrow + 1).PasteSpecial Paste:=xlPasteFormats
                '既存行の入力規則をコピー
                wsMeisai.Range("A" & strStartAddrow + 1 & ":AX" & strStartAddrow + 1).PasteSpecial Paste:=xlPasteValidation
                wsMeisai.Rows(strStartAddrow + 1).RowHeight = 26.25
            Else
                'ノンフリート（最終はＢＨ列）
                '行を新規追加
                wsMeisai.Range("A21:BH21").Copy
                '既存行の書式をコピー
                wsMeisai.Range("A" & strStartAddrow + 1 & ":BH" & strStartAddrow + 1).PasteSpecial Paste:=xlPasteFormats
                '既存行の入力規則をコピー
                wsMeisai.Range("A" & strStartAddrow + 1 & ":BH" & strStartAddrow + 1).PasteSpecial Paste:=xlPasteValidation
                wsMeisai.Rows(strStartAddrow + 1).RowHeight = 26.25

            End If

            '明細Noを更新
            wsMeisai.Range("B" & strStartAddrow + 1) = strStartAddrow - 19

            '明細チェックボックスを新規追加
            Call subConAdd("A" & strAddRowRange, "chkMeisai", "", strLastCon, False)
            '明細チェックのセルにコントロールの番号/チェック状態を設定
            wsMeisai.Range("A" & strAddRowRange) = Val(strLastCon) + 1 & "/False"

            '選択ボタンを新規追加
            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
            If FleetTypeFlg = 1 Then  'フリート（Ｗ列）
                Call subConAdd("W" & strAddRowRange, "btnOtherRate", "", strLastCon, False)
            Else
                'ノンフリート（ＡＥ列）
                Call subConAdd("AE" & strAddRowRange, "btnOtherRate", "", strLastCon, False)
            End If

            '最終行のコントロールの番号を更新
            strLastCon = Val(strLastCon) + 1
            '追加カウント
            strAddCnt = strAddCnt + 1
            '追加行を一つ下にずらす
            strStartAddrow = strStartAddrow + 1

        Loop

    End If

    Application.CutCopyMode = False                                  'COPY選択解除
    objSouhuho.Value = CStr(Val(strAllCnt) + Val(strAddrow)) + " 台" '総付保台数更新
    wsMeisai.Range("A1").Select                                      'フォーカスを先頭に設定
    Application.EnableEvents = True                                  'イベント有効
    Application.ScreenUpdating = True                                '描画開始
    Call subCellProtect(Val(strAllCnt) + Val(strAddrow))             'セル入力可能範囲変更

    Set wsMeisai = Nothing
    Set objSouhuho = Nothing

End Sub


Public Sub subConAdd(ByRef strAddRowRange As String, strConName As String, ByRef strCopyValue As String, _
                     ByRef strLastCon As String, ByRef blnAddMode As Boolean)
'関数名：subConAdd
'内容　：コントロールの追加
'引数　：
'        strAddRowRange     = 追加セル
'        strConName         = 追加コントロールの名前
'        strCopyValue       = コピー元のオブジェクトの入力内容(新規作成・入力内容の存在しない場合は、ブランク)
'        strLastCon         = 明細入力シートに存在する最終行のコントロールの番号
'        blnAddMode         = 新規追加・コピーの判別フラグ

    Dim strCellLeft As String        'コントロールを作成するセルの左位置
    Dim strCellTop  As String        'コントロールを作成するセルの上位置
    Dim wsMeisai    As Worksheet     '明細入力ワークシート
    
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(1, wsMeisai)                       'シートオブジェクト(明細入力)
    Else
        Call subSetSheet(17, wsMeisai)                       'シートオブジェクト(明細入力)
    End If
    strCellLeft = wsMeisai.Range(strAddRowRange).Left   'コントロールを作成するセルの左位置を設定
    strCellTop = wsMeisai.Range(strAddRowRange).Top     'コントロールを作成するセルの上位置を設定
    
    'コントロールを新規作成
    If strConName Like "chkMeisai*" Then
        '明細チェックボックス
        wsMeisai.CheckBoxes.Add(strCellLeft + 8.25, strCellTop + 5, 24, 16.5).Select
        Selection.OnAction = "subClickchkMeisai"
        Selection.Characters.Text = ""
        
        'コピーの場合､チェックを付ける
'        If blnAddMode then
'            Selection.Value = xlOn
'        End If
    ElseIf strConName Like "btnOtherRate*" Then
        '選択ボタン
        wsMeisai.Buttons.Add(strCellLeft + 7, strCellTop + 4, 35, 18).Select
        Selection.OnAction = "subClickOtherBtn"
        Selection.Characters.Text = "選択"
        Selection.Font.Size = 8
    End If
    
    '作成したコントロールの名前を変更
    Selection.Name = strConName & Val(strLastCon) + 1
    
    Set wsMeisai = Nothing
    
End Sub

Public Sub subMeisaiDel(ByRef strDelRowCnt As String, ByRef strDelmode As String)
'関数名：subMeisaiDel
'内容　：明細行の削除
'引数　：
'        strDelRowCnt       = 削除行数
'        strDelmode         = 遷移元画面の判定("1" = 明細入力シート/ "2" = 明細入力シート以外)

    Dim i As Integer                'ループ用カウント
    Dim strAllCnt As String         '総付保台数
    Dim strDelCnt As String         '削除する行数
    Dim strDelRow As String         '削除する行番号
    Dim strDelCon As String         '削除するコントロールの番号
    Dim strDelRowArr() As String    '削除する行番号を格納する配列
    Dim strDelConArr() As String    '削除するコントロールの番号を格納する配列
    Dim strLastRow As String        '最終行の行番号
    Dim strLastCon As String        '最終行のコントロールの番号
    Dim rngChkAll  As Range         'ループ用レンジ

    Dim wsMeisai As Worksheet           '明細入力ワークシート
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(1, wsMeisai)       'シートオブジェクト(明細入力)
    Else
        Call subSetSheet(17, wsMeisai)      'シートオブジェクト(明細入力（ノンフリート）)
    End If

    Dim objSouhuho As Object
    Set objSouhuho = wsMeisai.OLEObjects("txtSouhuho").Object

    strAllCnt = Val(Left(objSouhuho.Value, Len(objSouhuho.Value) - 2))      '総付保台数を設定
    Application.ScreenUpdating = False                                      '描画停止
    Application.EnableEvents = False                                        'イベント無効


    '削除カウントをリセット
    strDelCnt = 0

    '削除処理
    If strDelmode = "1" Then

        '明細入力画面から遷移してきた場合、チェックの入っている明細行を削除対象に設定
        For Each rngChkAll In wsMeisai.Range("A21:" & wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp).Address)

            If InStr(rngChkAll.Value, "True") > 0 Then
                '削除行の情報を格納している配列の要素数を変更
                ReDim Preserve strDelRowArr(strDelCnt) As String
                ReDim Preserve strDelConArr(strDelCnt) As String

                'チェックされている明細チェックのコントロールの番号を設定
                strDelConArr(strDelCnt) = Left(rngChkAll, InStr(rngChkAll, "/") - 1)
                'チェックされている明細チェックの行番号を設定
                strDelRowArr(strDelCnt) = wsMeisai.Shapes("chkMeisai" & strDelConArr(strDelCnt)).TopLeftCell.Row

                '削除する行をカウント
                strDelCnt = strDelCnt + 1
                '総付保台数から削除数分を引く
                strAllCnt = strAllCnt - 1

            End If

        Next rngChkAll

        Set rngChkAll = Nothing

    Else

        '明細入力画面以外から遷移してきた場合、最終行から削除行数分を削除対象に設定
        '削除行の情報を格納している配列の要素数を変更
        ReDim Preserve strDelRowArr(Val(strDelRowCnt) - 1) As String
        ReDim Preserve strDelConArr(Val(strDelRowCnt) - 1) As String

        '最終行の行番号を設定
        strLastCon = Left(wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp), InStr(wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp), "/") - 1)
        '最終行のコントロールの番号を設定
        strLastRow = wsMeisai.Shapes("chkMeisai" & strLastCon).TopLeftCell.Row

        For i = 0 To Val(strDelRowCnt - 1)

            '最終行から削除行数分の行番号とコントロールの番号を設定
            strDelRow = Val(strLastRow) - i
            strDelCon = Left(wsMeisai.Range("A" & strDelRow), InStr(wsMeisai.Range("A" & strDelRow), "/") - 1)

            strDelConArr(strDelRowCnt - 1 - i) = strDelCon
            strDelRowArr(strDelRowCnt - 1 - i) = strDelRow

            '削除する行をカウント
            strDelCnt = strDelCnt + 1
            '総付保台数から削除数分を引く
            strAllCnt = strAllCnt - 1

        Next i

    End If

    strDelCnt = strDelCnt - 1

    '削除行数が0件以上の場合、明細行を削除
    If Val(strDelCnt) > -1 Then

        For i = LBound(strDelConArr) To UBound(strDelConArr)
            If IsNull(strDelConArr(strDelCnt - i)) = False Then

                If strAllCnt = 0 And i = UBound(strDelConArr) Then

                    '対象の明細行がシートに存在する最後の1行の場合、選択状態・入力内容のクリアのみ行い行を残す
                    Dim strChkValue As String

                    '行をクリア
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    If FleetTypeFlg = 1 Then  'フリート（最終はAX列）
                        wsMeisai.Range("B" & strDelRowArr(strDelCnt - i) & ":" & "AX" & strDelRowArr(strDelCnt - i)).ClearContents
                    Else
                        'ノンフリート（最終はBH列）
                        wsMeisai.Range("B" & strDelRowArr(strDelCnt - i) & ":" & "BH" & strDelRowArr(strDelCnt - i)).ClearContents
                    End If


                    '明細チェックボックスをクリア
                    wsMeisai.Shapes("chkMeisai" & strDelConArr(strDelCnt - i)).ControlFormat.Value = xlOff
                    strChkValue = wsMeisai.Cells(strDelRowArr(strDelCnt - i), 1)
                    wsMeisai.Cells(strDelRowArr(strDelCnt - i), 1) = Left(strChkValue, InStr(strChkValue, "/") - 1) & "/False"

'                    'その他料率テキストボックスをクリア
'                    wsMeisai.OLEObjects("txtOtherRate" & strDelConArr(strDelCnt - i)).Object.Value = ""

                    strAllCnt = strAllCnt + 1
                Else

                    '明細行が2行以上のこっている場合、対象の明細行を削除
                    '明細チェックボックスを削除
                    wsMeisai.Shapes("chkMeisai" & strDelConArr(strDelCnt - i)).Delete

                    '選択ボタンを削除
                    wsMeisai.Shapes("btnOtherRate" & strDelConArr(strDelCnt - i)).Delete

'                    'その他料率テキストボックスを削除
'                    wsMeisai.OLEObjects("txtOtherRate" & strDelConArr(strDelCnt - i)).Delete

                    '行を削除
                    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    If FleetTypeFlg = 1 Then  'フリート（最終はAX列）
                        wsMeisai.Range("A" & strDelRowArr(strDelCnt - i) & ":AX" & strDelRowArr(strDelCnt - i)).Select
                        Selection.Delete Shift:=xlUp
                    Else
                        'ノンフリート（最終はBH列）
                        wsMeisai.Range("A" & strDelRowArr(strDelCnt - i) & ":BH" & strDelRowArr(strDelCnt - i)).Select
                        Selection.Delete Shift:=xlUp
                    End If

                End If

            End If
        Next i

        '明細Noを更新
        Dim strMeisaiNoRow As String
        strMeisaiNoRow = wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp).Row

        For i = 0 To wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp).Row - 20
            '明細Noを更新
            wsMeisai.Range("B" & strMeisaiNoRow - i) = (strMeisaiNoRow - i) - 20

            '対象の行の明細Noとひとつ前の行の明細Noが連番になった場合、ループを抜ける
            If wsMeisai.Range("B" & strMeisaiNoRow - i) - wsMeisai.Range("B" & (strMeisaiNoRow - i) - 1) = 1 Then
                Exit For
            End If
        Next i

    End If

    objSouhuho.Value = strAllCnt + " 台"        '総付保台数更新
    wsMeisai.Range("A1").Select                 'フォーカスを先頭に設定
    Application.EnableEvents = True             'イベント有効
    Application.ScreenUpdating = True           '描画停止
    Call subCellProtect(strAllCnt)              'セル入力可能範囲変更

    Set wsMeisai = Nothing
    Set objSouhuho = Nothing

End Sub

Public Function fncTempSave_Kyotsu(strAllCell As String) As String
'関数名：fncTempSave_Kyotsu
'内容　：共通情報レコード（1行目）生成
'引数　：
'        strAllCell         = 総付保台数

    Dim wsKyoutsuU As Worksheet             'コード値ワークシート
    Call subSetSheet(2, wsKyoutsuU)         'シートオブジェクト(別紙　共通項目)

    Dim wsMoushikomiP As Worksheet
    Call subSetSheet(8, wsMoushikomiP)      'シートオブジェクト(申込書印刷画面内容）

    '「保険終期日」の計算(保険始期日より保険期間（1年）分足す)
    Dim strSaveDate As String
    strSaveDate = Left(wsKyoutsuU.Range("E2").Value, 4) & "/" & Mid(wsKyoutsuU.Range("E2").Value, 5, 2) & "/" & Mid(wsKyoutsuU.Range("E2").Value, 7, 2)
    strSaveDate = CStr(DateAdd("yyyy", wsKyoutsuU.Range("F2").Value, CDate(strSaveDate)))
    strSaveDate = Format(strSaveDate, "YYYYMMDD")

    '出力文字列生成
    Dim strSaveCont As String
    strSaveCont = strSaveCont & "1" & ","                           'レコード区分
    strSaveCont = strSaveCont & wsKyoutsuU.Range("A2").Value & ","  '受付区分
    strSaveCont = strSaveCont & wsKyoutsuU.Range("B2").Value & ","  '被保険者_個人法人区分
    strSaveCont = strSaveCont & wsKyoutsuU.Range("C2").Value & ","  '保険種類
    strSaveCont = strSaveCont & wsKyoutsuU.Range("D2").Value & ","  'フリート・ノンフリート区分
    strSaveCont = strSaveCont & wsKyoutsuU.Range("E2").Value & ","  '保険始期日
    strSaveCont = strSaveCont & "P" & ","                           '保険始期時刻区分
    strSaveCont = strSaveCont & "4" & ","                           '保険始期時刻
    strSaveCont = strSaveCont & strSaveDate & ","                   '保険終期日
    strSaveCont = strSaveCont & wsKyoutsuU.Range("G2").Value & ","  '計算方法
    strSaveCont = strSaveCont & wsKyoutsuU.Range("F2").Value & ","  '保険期間_年
    strSaveCont = strSaveCont & "" & ","                            '保険期間_月
    strSaveCont = strSaveCont & "" & ","                            '保険期間_日
    strSaveCont = strSaveCont & wsKyoutsuU.Range("H2").Value & ","  '払込方法
    strSaveCont = strSaveCont & IIf(wsKyoutsuU.Range("I2").Value = "" _
                    , "", Val(wsKyoutsuU.Range("I2").Value)) & ","  'フリート優良割引
    strSaveCont = strSaveCont & IIf(wsKyoutsuU.Range("J2").Value = "", _
                      "", Val(wsKyoutsuU.Range("J2").Value)) & ","  '第一種デメ割増
    strSaveCont = strSaveCont & wsKyoutsuU.Range("K2").Value & ","  'フリート多数割引
    strSaveCont = strSaveCont & wsKyoutsuU.Range("L2").Value & ","  'フリートコード
    strSaveCont = strSaveCont & strAllCell & ","                    '総付保台数
    strSaveCont = strSaveCont & "" & ","                              '先日付フラグ
    If blnChouhyouflg Then
        strSaveCont = strSaveCont & wsMoushikomiP.Range("B1").Value & _
                             String(3 - Len(wsMoushikomiP.Range("B1").Value), " ") '郵便番号（前）
        strSaveCont = strSaveCont & wsMoushikomiP.Range("C1").Value & "," '郵便番号（後）
        strSaveCont = strSaveCont & StrConv(wsMoushikomiP.Range("D1").Value, vbWide) & "," '契約者住所（カナ）
        strSaveCont = strSaveCont & wsMoushikomiP.Range("E1").Value & _
                             String(40 - Len(wsMoushikomiP.Range("E1").Value), " ") '契約者住所（漢字）
        strSaveCont = strSaveCont & wsMoushikomiP.Range("F1").Value & "," '契約者住所（漢字）
        strSaveCont = strSaveCont & StrConv(wsMoushikomiP.Range("G1").Value, vbWide) & "," '法人名（カナ）
        strSaveCont = strSaveCont & wsMoushikomiP.Range("H1").Value & "," '法人名（漢字）
        strSaveCont = strSaveCont & StrConv(wsMoushikomiP.Range("I1").Value, vbWide) & "," '役職名・氏名（カナ）
        strSaveCont = strSaveCont & wsMoushikomiP.Range("J1").Value & "," '役職名・氏名（漢字）
        strSaveCont = strSaveCont & StrConv(wsMoushikomiP.Range("K3").Value, vbNarrow) & "," '連絡先１　自宅・携帯
        strSaveCont = strSaveCont & StrConv(wsMoushikomiP.Range("N3").Value, vbNarrow) & "," '連絡先２　勤務先
        strSaveCont = strSaveCont & StrConv(wsMoushikomiP.Range("Q3").Value, vbNarrow) & "," '連絡先３　ＦＡＸ
        strSaveCont = strSaveCont & wsMoushikomiP.Range("T1").Value & "," '団体名
        strSaveCont = strSaveCont & wsMoushikomiP.Range("U1").Value & "," '団体コード
        strSaveCont = strSaveCont & wsMoushikomiP.Range("V1").Value & "," '団体扱に関する特約
        strSaveCont = strSaveCont & wsMoushikomiP.Range("W1").Value & "," '所属コード
        strSaveCont = strSaveCont & wsMoushikomiP.Range("X1").Value & "," '社員コード
        strSaveCont = strSaveCont & wsMoushikomiP.Range("Z1").Value & "," '部課コード
        strSaveCont = strSaveCont & wsMoushikomiP.Range("AB1").Value & "," '代理店コード
        strSaveCont = strSaveCont & wsMoushikomiP.Range("A1").Value & vbCrLf '証券番号
    Else
        strSaveCont = strSaveCont & wsKyoutsuU.Range("O2").Value & "," '郵便番号
        strSaveCont = strSaveCont & wsKyoutsuU.Range("P2").Value & "," '契約者住所（カナ）
        strSaveCont = strSaveCont & wsKyoutsuU.Range("Q2").Value & "," '契約者住所（漢字）
        strSaveCont = strSaveCont & wsKyoutsuU.Range("R2").Value & "," '法人名（カナ）
        strSaveCont = strSaveCont & wsKyoutsuU.Range("S2").Value & "," '法人名（漢字）
        strSaveCont = strSaveCont & wsKyoutsuU.Range("T2").Value & "," '役職名・氏名（カナ）
        strSaveCont = strSaveCont & wsKyoutsuU.Range("U2").Value & "," '役職名・氏名（漢字）
        strSaveCont = strSaveCont & wsKyoutsuU.Range("V2").Value & "," '連絡先１　自宅・携帯
        strSaveCont = strSaveCont & wsKyoutsuU.Range("W2").Value & "," '連絡先２　勤務先
        strSaveCont = strSaveCont & wsKyoutsuU.Range("X2").Value & "," '連絡先３　ＦＡＸ
        strSaveCont = strSaveCont & wsKyoutsuU.Range("Y2").Value & "," '団体名
        strSaveCont = strSaveCont & wsKyoutsuU.Range("Z2").Value & "," '団体コード
        strSaveCont = strSaveCont & wsKyoutsuU.Range("AA2").Value & "," '団体扱に関する特約
        strSaveCont = strSaveCont & wsKyoutsuU.Range("AB2").Value & "," '所属コード
        strSaveCont = strSaveCont & wsKyoutsuU.Range("AC2").Value & "," '社員コード
        strSaveCont = strSaveCont & wsKyoutsuU.Range("AD2").Value & "," '部課コード
        strSaveCont = strSaveCont & wsKyoutsuU.Range("AE2").Value & "," '代理店コード
        strSaveCont = strSaveCont & wsKyoutsuU.Range("AF2").Value & vbCrLf '証券番号
    End If
    
    fncTempSave_Kyotsu = strSaveCont
    
    Set wsKyoutsuU = Nothing
    
End Function

Public Function fncTempSave_Meisai(ByRef intCarCnt As Integer, ByVal strMeisai As String, _
                                    ByRef blnOutFlg As Boolean, ByVal blnDownloadflg) As String

    '関数名：fncTempSave_Meisai
    '内容　：明細情報レコード（2行目以降）生成
    '引数　：
    '        intCarCnt          = 総付保台数
    '        strMeisai          = 追加コントロールの名前
    '        objAllOtherRate    = コピー元のオブジェクトの入力内容(新規作成・入力内容の存在しない場合は、ブランク)
    '        intFileMaxCar      = 明細入力シートに存在する最終行のコントロールの番号
    '        blnOutFlg          = 新規追加・コピーの判別フラグ
    
    Dim strDate As String
    
    strDate = ""

    If blnOutFlg Then
        strMeisai = ""
        blnOutFlg = False
    End If
    
    Dim wsMeisai As Worksheet
    Call subSetSheet(1, wsMeisai)       'シートオブジェクト(明細入力)
    
    Dim wsTextM As Worksheet
    Call subSetSheet(7, wsTextM)
    
    '処理行
    Dim strTmpRow As Integer
    strTmpRow = 20 + intCarCnt
    
    With wsMeisai
        strMeisai = strMeisai & "2" & ","                                                              'レコード区分
        strMeisai = strMeisai & fncTrimComma(CStr(.Cells(strTmpRow, 4))) & ","                         '用途車種（コード）
        If blnDownloadflg Then
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 5))) & ","                     '車名
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 8))) & ","                     '型式
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 9))) & ","                     '仕様
            '初度登録年月
            strDate = fncWarekiCheck(Trim(.Cells(strTmpRow, 10)) & "1日", 6)
            If strDate = "" Then
                strDate = fncDateCheck(Trim(.Cells(strTmpRow, 10)), True, True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(Trim(.Cells(strTmpRow, 10))), 6) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 10))) & ","
                End If
            Else
                strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 10))) & ","
            End If

        Else
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 5)) & ","                           '車名
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 8)) & ","                           '型式
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 9)) & ","                           '仕様
            '初度登録年月
            strDate = fncWarekiCheck(.Cells(strTmpRow, 10) & "1日", 6)
            If strDate = "" Then
                strDate = fncDateCheck(.Cells(strTmpRow, 10), True, True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(.Cells(strTmpRow, 10)), 6) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 10)) & ","
                End If
            Else
                strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 10)) & ","
            End If
        End If
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 12).Value), "AD") & ","     '改造車
        
        '前0削除
        Dim i As String
        If blnDownloadflg Then
            i = fncTrimComma(CStr(Trim(.Cells(strTmpRow, 13))))              '排気量
        Else
            i = fncTrimComma(CStr(.Cells(strTmpRow, 13)))                    '排気量
        End If
        Do Until Len(i) <= 2
            If Left(i, 1) = "0" And IsNumeric(Mid(i, 2, 1)) Then
                i = Mid(i, 2)
            Else
                Exit Do
            End If
        Loop
        strMeisai = strMeisai & i & ","                    '排気量
        
        strMeisai = strMeisai & fncTrimComma(CStr(.Cells(strTmpRow, 14))) & ","                        '2.5リットル超ディーゼル自家用小型乗用車
        strMeisai = strMeisai & "" & ","                                                               '被保険者_生年月日    :ノンフリート
        strMeisai = strMeisai & "" & ","                                                               'ノンフリート等級     :ノンフリート
        strMeisai = strMeisai & "" & ","                                                               '事故有適用期間       :ノンフリート
        strMeisai = strMeisai & "" & ","                                                               'ノンフリート多数割引 :ノンフリート
        strMeisai = strMeisai & "" & ","                                                               '団体割増引           :ノンフリート
        strMeisai = strMeisai & "" & ","                                                               'ゴールド免許割引     :ノンフリート
        strMeisai = strMeisai & "" & ","                                                               '使用目的             :ノンフリート
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 24) Like "*沖縄*", "3 ", "") & ","                        '沖縄
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 24) Like "*レンタカー*", "1", "") & ","                   'レンタカー
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 24) Like "*教習車*", "5 ", "") & ","                      '教習車
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 24) Like "*ブーム対象外*", "1 ", "") & ","                'ブーム対象外
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 24) Like "*リースカーオープンポリシー*", "80", "") & ","  'リースカーオープンポリシー
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 24) Like "*オープンポリシー多数割引*", "93", "") & ","    'オープンポリシー多数割引
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 24) Like "*準公有*", fncFindName("準公有", "AT"), _
                                IIf(.Cells(strTmpRow, 24) Like "*公有*", fncFindName("公有", "AT"), "")) & ","  '公有・準公有車
        ''2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        If blnMoushikomiflg Then
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 16).Value) & ","                              '車両料率クラス
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 17).Value) & ","                              '対人料率クラス
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 18).Value) & ","                              '対物料率クラス
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 19).Value) & ","                              '傷害料率クラス
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 20).Value) & ","                              '新車割引
            strMeisai = strMeisai & IIf(.Cells(strTmpRow, 24) Like "*特種区分*", "8", "") & ","                  '特種区分
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 22).Value) & ","                              '車両下限価格
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 21).Value) & ","                              '車両上限価格
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 47).Value) & ","                              '合計保険料
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 48).Value) & ","                              '初回保険料
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 49).Value) & ","                              '年間保険料
        Else
            strMeisai = strMeisai & "" & ","                                                               '車両料率クラス
            strMeisai = strMeisai & "" & ","                                                               '対人料率クラス
            strMeisai = strMeisai & "" & ","                                                               '対物料率クラス
            strMeisai = strMeisai & "" & ","                                                               '傷害料率クラス
            strMeisai = strMeisai & "" & ","                                                               '新車割引
            strMeisai = strMeisai & "" & ","                                                               '特種区分
            strMeisai = strMeisai & "" & ","                                                               '車両下限価格
            strMeisai = strMeisai & "" & ","                                                               '車両上限価格
            strMeisai = strMeisai & "" & ","                                                               '合計保険料
            strMeisai = strMeisai & "" & ","                                                               '初回保険料
            strMeisai = strMeisai & "" & ","                                                               '年間保険料
        End If
        strMeisai = strMeisai & "" & ","                                                               '年齢条件             :ノンフリート
        strMeisai = strMeisai & "" & ","                                                               '高齢運転者対象外特約 :ノンフリート
        strMeisai = strMeisai & "" & ","                                                               '運転者限定特約       :ノンフリート
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 44)) <> "", "1", "") & ","          '従業員等限定特約
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 25).Value), "BJ") & ","     '車両保険の種類
        If blnDownloadflg Then
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 26)))) = "", "", Val(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 26)))))) & ","  '車両保険金額
        Else
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(.Cells(strTmpRow, 26))) = "", "", Val(fncTrimComma(CStr(.Cells(strTmpRow, 26))))) & ","   '車両保険金額
        End If
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 27)), "BN") & ","           '車両免責金額
        strMeisai = strMeisai & fncTrimComma(CStr(.Cells(strTmpRow, 45))) & ","                        '代車等セット
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 28)) <> "", "2", "") & ","          '車両全損臨費特約
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 30)) <> "", "1", "") & ","          '車両盗難対象外特約
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 29)) <> "", "1", "") & ","          '車両超過修理費用特約
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 31) <> "無制限", "", _
                                fncFindName(fncTrimComma(.Cells(strTmpRow, 31)), "CD")) & ","          '対人無制限
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 31) <> "対象外", "", _
                                fncFindName(fncTrimComma(.Cells(strTmpRow, 31)), "CD")) & ","          '対人対象外
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 31) = "無制限", "", _
                                IIf(.Cells(strTmpRow, 31) = "対象外", "", _
                                    fncTrimComma(fncFindName(.Cells(strTmpRow, 31), "CD")))) & ","     '対人賠償保険金額
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 32)) <> "", "1", "") & ","          '自損事故傷害特約
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 33)) <> "", "1", "") & ","          '無保険車事故傷害特約
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 34) <> "無制限", "", _
                                fncFindName(fncTrimComma(.Cells(strTmpRow, 34)), "CH")) & ","          '対物無制限
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 34) <> "対象外", "", _
                                fncFindName(fncTrimComma(.Cells(strTmpRow, 34)), "CH")) & ","          '対人対象外
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 34) = "無制限", "", _
                                IIf(.Cells(strTmpRow, 34) = "対象外", "", _
                                    fncTrimComma(fncFindName(.Cells(strTmpRow, 34), "CH")))) & ","     '対物賠償保険金額
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 35)), "BR") & ","           '対物免責金額
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 36)) <> "", "1", "") & ","          '対物超過修理費用特約
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 37) = "対象外", "", _
                                    fncTrimComma(fncFindName(.Cells(strTmpRow, 37), "CL"))) & ","      '人身傷害 1名
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 37) <> "対象外", "", _
                                    fncFindName(fncTrimComma(.Cells(strTmpRow, 37)), "CL")) & ","      '人身対象外
        If blnDownloadflg Then
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 38)))) = "", "", Val(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 38)))))) & ","  '人身傷害 1事故
        Else
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(.Cells(strTmpRow, 38))) = "", "", Val(fncTrimComma(CStr(.Cells(strTmpRow, 38))))) & ","   '人身傷害 1事故
        End If
        strMeisai = strMeisai & "" & ","                                                               '自動車事故特約 :ノンフリート
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 39) = "対象外", "", _
                                    fncTrimComma(fncFindName(.Cells(strTmpRow, 39), "CP"))) & ","      '死亡・後遺傷害保険金額 1名
        If .Cells(strTmpRow, 39) = "" Then
            .Cells(strTmpRow, 39) = "対象外"
        End If
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 39) <> "対象外", "", _
                                    fncFindName(fncTrimComma(.Cells(strTmpRow, 39)), "CP")) & ","      '搭傷対象外
        If blnDownloadflg Then
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 40)))) = "", "", Val(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 40)))))) & ","            '死亡・後遺傷害保険金額 1事故
        Else
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(.Cells(strTmpRow, 40))) = "", "", Val(fncTrimComma(CStr(.Cells(strTmpRow, 40))))) & ","                    '死亡・後遺傷害保険金額 1事故
        End If
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 41)) <> "", "2", "") & ","          '日数払特約
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 42)) <> "", "1", "") & ","          '事業主費用特約
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 43)) <> "", "1", "") & ","          '弁護士費用特約
        strMeisai = strMeisai & "" & ","                                                               'ファミリーバイク特約   :ノンフリート
        strMeisai = strMeisai & "" & ","                                                               '個人賠償責任補償フラグ :ノンフリート
        ''2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        If blnMoushikomiflg Then
            strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 50)) = "稟議エラー有", "1", _
                                    IIf(fncTrimComma(.Cells(strTmpRow, 50)) = "警告有", "2", "")) & "," '稟議警告フラグ
        Else
            strMeisai = strMeisai & "" & ","                                                           '稟議警告フラグ
        End If
        If blnDownloadflg Then
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 6))) & ","                     '登録番号
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 7))) & ","                     '車台番号
            '車検満了日
            strDate = fncWarekiCheck(Trim(.Cells(strTmpRow, 11)))
            If strDate = "" Then
                strDate = fncDateCheck(Trim(.Cells(strTmpRow, 11)), True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(Trim(.Cells(strTmpRow, 11))), 8) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 11))) & ","
                End If
            Else
                strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 11))) & ","
            End If
        Else
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 6)) & ","                           '登録番号
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 7)) & ","                           '車台番号
            '車検満了日
            strDate = fncWarekiCheck(.Cells(strTmpRow, 11))
            If strDate = "" Then
                strDate = fncDateCheck(.Cells(strTmpRow, 11), True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(.Cells(strTmpRow, 11)), 8) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 11)) & ","
                End If
                    
            Else
                strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 11)) & ","
            End If
        End If
        strTmpRow = strTmpRow - 20
        strMeisai = strMeisai & fncTrimComma(wsTextM.Cells(strTmpRow, 72)) & ","       '登録番号(カナ)
        strTmpRow = strTmpRow + 20
'        strMeisai = strMeisai & "" & ","                                                               '登録番号(カナ)
        'ASV割引
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 15)) <> "", "1", "") & ","          'ブランクでない場合「1:適用する」
        '車両搬送時不適用特約
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 46)) <> "", "1", "") & ","          'ブランクでない場合「1:不適用」
'        strMeisai = strMeisai & vbCrLf
    End With
    '被保険者住所 (カナ)
    strMeisai = strMeisai & ","
    '被保険者氏名 (カナ)
    strMeisai = strMeisai & ","
    '被保険者氏名 (漢字)
    strMeisai = strMeisai & ","
    '免許証の色
    strMeisai = strMeisai & ","
    '免許証有効期限
    strMeisai = strMeisai & ","
    '車両所有者氏名 (カナ)
    strMeisai = strMeisai & ","
    '車両所有者氏名 (漢字)
    strMeisai = strMeisai & ","
    '所有権留保またはリース等
    strMeisai = strMeisai & vbCrLf

        
    fncTempSave_Meisai = strMeisai

    Set wsMeisai = Nothing

End Function

Public Function fncTempSave_NonFleetMeisai(ByRef intCarCnt As Integer, ByVal strMeisai As String, _
                                    ByRef blnOutFlg As Boolean, ByVal blnDownloadflg) As String

    '関数名：fncTempSave_NonFleetMeisai （2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
    '内容　：明細情報レコード（2行目以降）生成
    '引数　：
    '        intCarCnt          = 総付保台数
    '        strMeisai          = 追加コントロールの名前
    '        objAllOtherRate    = コピー元のオブジェクトの入力内容(新規作成・入力内容の存在しない場合は、ブランク)
    '        intFileMaxCar      = 明細入力シートに存在する最終行のコントロールの番号
    '        blnOutFlg          = 新規追加・コピーの判別フラグ
    
    Dim strDate As String
    
    strDate = ""

    If blnOutFlg Then
        strMeisai = ""
        blnOutFlg = False
    End If
    
    Dim wsMeisai As Worksheet
    Call subSetSheet(17, wsMeisai)      'シートオブジェクト(明細入力（ノンフリート）)
    
    Dim wsTextM As Worksheet
    Call subSetSheet(7, wsTextM)
    
    '共通画面のノンフリート多数割引、団体割引の取得用（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
    Dim wsKyoutsuU As Worksheet             'コード値ワークシート
    Call subSetSheet(2, wsKyoutsuU)         'シートオブジェクト(別紙　共通項目)
    
    Dim wsMeisaiP As Worksheet
    Call subSetSheet(19, wsMeisaiP)

    '処理行
    Dim strTmpRow As Integer
    strTmpRow = 20 + intCarCnt

    With wsMeisai
        strMeisai = strMeisai & "2" & ","                                                              'レコード区分
        strMeisai = strMeisai & fncTrimComma(CStr(.Cells(strTmpRow, 4))) & ","                         '用途車種（コード）
        If blnDownloadflg Then
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 5))) & ","                     '車名
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 8))) & ","                     '型式
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 9))) & ","                     '仕様
            '初度登録年月
            strDate = fncWarekiCheck(Trim(.Cells(strTmpRow, 10)) & "1日", 6)
            If strDate = "" Then
                strDate = fncDateCheck(Trim(.Cells(strTmpRow, 10)), True, True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(Trim(.Cells(strTmpRow, 10))), 6) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 10))) & ","
                End If
            Else
                strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 10))) & ","
            End If

        Else
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 5)) & ","                           '車名
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 8)) & ","                           '型式
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 9)) & ","                           '仕様
            '初度登録年月
            strDate = fncWarekiCheck(.Cells(strTmpRow, 10) & "1日", 6)
            If strDate = "" Then
                strDate = fncDateCheck(.Cells(strTmpRow, 10), True, True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(.Cells(strTmpRow, 10)), 6) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 10)) & ","
                End If
            Else
                strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 10)) & ","
            End If
        End If
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 12).Value), "AD") & ","     '改造車
        
        '前0削除
        Dim i As String
        If blnDownloadflg Then
            i = fncTrimComma(CStr(Trim(.Cells(strTmpRow, 13))))              '排気量
        Else
            i = fncTrimComma(CStr(.Cells(strTmpRow, 13)))                    '排気量
        End If
        Do Until Len(i) <= 2
            If Left(i, 1) = "0" And IsNumeric(Mid(i, 2, 1)) Then
                i = Mid(i, 2)
            Else
                Exit Do
            End If
        Loop
        strMeisai = strMeisai & i & ","                    '排気量
'        If blnDownloadflg Then
'            strMeisai = strMeisai & fncTrimComma(CStr(Trim(.Cells(strTmpRow, 13)))) & ","              '排気量
'        Else
'            strMeisai = strMeisai & fncTrimComma(CStr(.Cells(strTmpRow, 13))) & ","                    '排気量
'        End If
        strMeisai = strMeisai & fncTrimComma(CStr(.Cells(strTmpRow, 14))) & ","                        '2.5リットル超ディーゼル自家用小型乗用車
        
       If blnDownloadflg Then
            '被保険者生年月日 （2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
            strDate = fncWarekiCheck(Trim(.Cells(strTmpRow, 15)))
            If strDate = "" Then
                strDate = fncDateCheck(Trim(.Cells(strTmpRow, 15)), True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(Trim(.Cells(strTmpRow, 15))), 8) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 15))) & ","
                End If
            Else
                strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 15))) & ","
            End If

        Else
            '被保険者生年月日
            strDate = fncWarekiCheck(.Cells(strTmpRow, 15))
            If strDate = "" Then
                strDate = fncDateCheck(.Cells(strTmpRow, 15), True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(.Cells(strTmpRow, 15)), 8) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 15)) & ","
                End If

            Else
                strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 15)) & ","
            End If
       End If

       'ノンフリート等級（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 16).Value), "AH") & ","

       '事故有適用期間（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 17).Value), "AL") & ","

        'ノンフリート多数割引（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        strMeisai = strMeisai & wsKyoutsuU.Range("M2").Value & ","    '「別紙　共通項目」シートより取得

        '団体割増引（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        strMeisai = strMeisai & wsKyoutsuU.Range("N2").Value & ","    '「別紙　共通項目」シートより取得

        'ゴールド免許割引（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 18)) <> "", "1", "") & ","          'ブランクでない場合「1:適用する」

        '使用目的（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 19).Value), "DC") & ","

        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 32) Like "*沖縄*", "3 ", "") & ","                        '沖縄
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 32) Like "*レンタカー*", "1", "") & ","                   'レンタカー
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 32) Like "*教習車*", "5 ", "") & ","                      '教習車
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 32) Like "*ブーム対象外*", "1 ", "") & ","                'ブーム対象外
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 32) Like "*リースカーオープンポリシー*", "80", "") & ","  'リースカーオープンポリシー
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 32) Like "*オープンポリシー多数割引*", "93", "") & ","    'オープンポリシー多数割引
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 32) Like "*準公有*", fncFindName("準公有", "AT"), _
                                IIf(.Cells(strTmpRow, 32) Like "*公有*", fncFindName("公有", "AT"), "")) & ","  '公有・準公有車


        ''2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        If blnMoushikomiflg Then
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 24).Value) & ","                        '車両料率クラス
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 25).Value) & ","                        '対人料率クラス
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 26).Value) & ","                        '対物料率クラス
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 27).Value) & ","                        '傷害料率クラス
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 28).Value) & ","                        '新車割引
            strMeisai = strMeisai & IIf(.Cells(strTmpRow, 32) Like "*特種区分*", "8", "") & ","            '特種区分
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 30).Value) & ","                        '車両下限価格
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 29).Value) & ","                        '車両上限価格
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 57).Value) & ","                        '合計保険料
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 58).Value) & ","                        '初回保険料
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 59).Value) & ","                        '年間保険料
        Else
            strMeisai = strMeisai & "" & ","                                                               '車両料率クラス
            strMeisai = strMeisai & "" & ","                                                               '対人料率クラス
            strMeisai = strMeisai & "" & ","                                                               '対物料率クラス
            strMeisai = strMeisai & "" & ","                                                               '傷害料率クラス
            strMeisai = strMeisai & "" & ","                                                               '新車割引
            strMeisai = strMeisai & "" & ","                                                               '特種区分
            strMeisai = strMeisai & "" & ","                                                               '車両下限価格
            strMeisai = strMeisai & "" & ","                                                               '車両上限価格
            strMeisai = strMeisai & "" & ","                                                               '合計保険料
            strMeisai = strMeisai & "" & ","                                                               '初回保険料
            strMeisai = strMeisai & "" & ","                                                               '年間保険料
        End If
        
        '年齢条件（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 20).Value), "BB") & ","
        
        '高齢運転者対象外特約（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 21)) <> "", "1", "") & ","          'ブランクでない場合「1:対象外」
        
        '運転者限定特約（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 22).Value), "BF") & ","
              
        '従業員等限定特約（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        strMeisai = strMeisai & "" & ","
        
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 33).Value), "BJ") & ","     '車両保険の種類
        If blnDownloadflg Then
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 34)))) = "", "", Val(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 34)))))) & ","  '車両保険金額
        Else
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(.Cells(strTmpRow, 34))) = "", "", Val(fncTrimComma(CStr(.Cells(strTmpRow, 34))))) & ","              '車両保険金額
        End If
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 35)), "BN") & ","           '車両免責金額
        strMeisai = strMeisai & fncTrimComma(CStr(.Cells(strTmpRow, 55))) & ","                        '事故代車・身の回り品特約
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 36)) <> "", "2", "") & ","          '車両全損臨費特約
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 38)) <> "", "1", "") & ","          '車両盗難対象外特約
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 37)) <> "", "1", "") & ","          '車両超過修理費用特約
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 39) <> "無制限", "", _
                                fncFindName(fncTrimComma(.Cells(strTmpRow, 39)), "CD")) & ","          '対人無制限
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 39) <> "対象外", "", _
                                fncFindName(fncTrimComma(.Cells(strTmpRow, 39)), "CD")) & ","          '対人対象外
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 39) = "無制限", "", _
                                IIf(.Cells(strTmpRow, 39) = "対象外", "", _
                                    fncTrimComma(fncFindName(.Cells(strTmpRow, 39), "CD")))) & ","     '対人賠償保険金額
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 40)) <> "", "1", "") & ","          '自損事故傷害特約
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 41)) <> "", "1", "") & ","          '無保険車事故傷害特約
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 42) <> "無制限", "", _
                                fncFindName(fncTrimComma(.Cells(strTmpRow, 42)), "CH")) & ","          '対物無制限
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 42) <> "対象外", "", _
                                fncFindName(fncTrimComma(.Cells(strTmpRow, 42)), "CH")) & ","          '対人対象外
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 42) = "無制限", "", _
                                IIf(.Cells(strTmpRow, 42) = "対象外", "", _
                                    fncTrimComma(fncFindName(.Cells(strTmpRow, 42), "CH")))) & ","     '対物賠償保険金額
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 43)), "BR") & ","           '対物免責金額
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 44)) <> "", "1", "") & ","          '対物超過修理費用特約
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 45) = "対象外", "", _
                                    fncTrimComma(fncFindName(.Cells(strTmpRow, 45), "CL"))) & ","      '人身傷害 1名
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 45) <> "対象外", "", _
                                    fncFindName(fncTrimComma(.Cells(strTmpRow, 45)), "CL")) & ","      '人身対象外
        If blnDownloadflg Then
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 46)))) = "", "", Val(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 46)))))) & ","  '人身傷害 1事故
        Else
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(.Cells(strTmpRow, 46))) = "", "", Val(fncTrimComma(CStr(.Cells(strTmpRow, 46))))) & ","   '人身傷害 1事故
        End If
        '自動車事故特約（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 54)) <> "", "2", "") & ","          'ブランクでない場合「2:適用する」

        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 47) = "対象外", "", _
                                    fncTrimComma(fncFindName(.Cells(strTmpRow, 47), "CP"))) & ","      '死亡・後遺傷害保険金額 1名
        If .Cells(strTmpRow, 47) = "" Then
            .Cells(strTmpRow, 47) = "対象外"
        End If
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 47) <> "対象外", "", _
                                    fncFindName(fncTrimComma(.Cells(strTmpRow, 47)), "CP")) & ","      '搭傷対象外
        If blnDownloadflg Then
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 48)))) = "", "", Val(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 48)))))) & ","        '死亡・後遺傷害保険金額 1事故
        Else
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(.Cells(strTmpRow, 48))) = "", "", Val(fncTrimComma(CStr(.Cells(strTmpRow, 48))))) & ","                    '死亡・後遺傷害保険金額 1事故
        End If
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 49)) <> "", "2", "") & ","          '日数払特約
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 50)) <> "", "1", "") & ","          '事業主費用特約
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 51)) <> "", "1", "") & ","          '弁護士費用特約

        'ファミリーバイク特約（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 52).Value), "BV") & ","

        '個人賠償責任補償特約
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 53)) <> "", "1", "") & ","          'ブランクでない場合「1:適用する」
        ''2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        If blnMoushikomiflg Then
            strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 60)) = "稟議エラー有", "1", _
                                    IIf(fncTrimComma(.Cells(strTmpRow, 60)) = "警告有", "2", "")) & "," '稟議警告フラグ

        Else
            strMeisai = strMeisai & "" & ","                                                           '稟議警告フラグ
        End If
        If blnDownloadflg Then
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 6))) & ","                     '登録番号
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 7))) & ","                     '車台番号
            '車検満了日
            strDate = fncWarekiCheck(Trim(.Cells(strTmpRow, 11)))
            If strDate = "" Then
                strDate = fncDateCheck(Trim(.Cells(strTmpRow, 11)), True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(Trim(.Cells(strTmpRow, 11))), 8) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 11))) & ","
                End If
            Else
                strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 11))) & ","
            End If
        Else
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 6)) & ","                           '登録番号
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 7)) & ","                           '車台番号
            '車検満了日
            strDate = fncWarekiCheck(.Cells(strTmpRow, 11))
            If strDate = "" Then
                strDate = fncDateCheck(.Cells(strTmpRow, 11), True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(.Cells(strTmpRow, 11)), 8) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 11)) & ","
                End If

            Else
                strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 11)) & ","
            End If
        End If
        strTmpRow = strTmpRow - 20
        strMeisai = strMeisai & fncTrimComma(wsTextM.Cells(strTmpRow, 72)) & "," '登録番号(カナ)
        strTmpRow = strTmpRow + 20
'        strMeisai = strMeisai & "" & ","                                                               '登録番号(カナ)
        'ASV割引
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 23)) <> "", "1", "") & ","          'ブランクでない場合「1:適用する」
        '車両搬送時不適用特約
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 56)) <> "", "1", "") & ","          'ブランクでない場合「1:不適用」
'        strMeisai = strMeisai & vbCrLf
    End With

    If blnChouhyouflg Then
        With wsMeisaiP
            strTmpRow = strTmpRow - 14
    '       被保険者住所 (ｶﾅ)
            strMeisai = strMeisai & StrConv(fncTrimComma(.Cells(strTmpRow, 3)), vbWide) & ","
    '       被保険者氏名 (ｶﾅ)
            strMeisai = strMeisai & StrConv(fncTrimComma(.Cells(strTmpRow, 8)), vbWide) & ","
    '       被保険者氏名 (漢字)
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 11)) & ","
    '       免許証の色
            strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 24).Value), "DT") & ","
    '       免許証有効期限
            strDate = fncDateCheck(.Cells(strTmpRow, 27), False)
            If strDate = "" Then
                strMeisai = strMeisai & fncToSeireki(fncTrimComma(.Cells(strTmpRow, 27)), 8) & ","
            Else
                strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 27)) & ","
            End If
    '        strMeisai = strMeisai & fncToSeireki(fncTrimComma(.Cells(strTmpRow, 27)), 8) & ","
            strTmpRow = strTmpRow + 12
    '       車両所有者氏名（ｶﾅ）
            strMeisai = strMeisai & StrConv(fncTrimComma(.Cells(strTmpRow, 16)), vbWide) & ","
    '       車両所有者氏名（漢字）
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 24)) & ","
    '       所有権留保またはリース等
            strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 31)) <> "", "1", "")
            strMeisai = strMeisai & vbCrLf
        End With
    Else
        With wsTextM
            strTmpRow = strTmpRow - 20
    '       被保険者住所 (ｶﾅ)
            strMeisai = strMeisai & .Cells(strTmpRow, 75) & ","
    '       被保険者氏名 (ｶﾅ)
            strMeisai = strMeisai & .Cells(strTmpRow, 76) & ","
    '       被保険者氏名 (漢字)
            strMeisai = strMeisai & .Cells(strTmpRow, 77) & ","
    '       免許証の色
            strMeisai = strMeisai & .Cells(strTmpRow, 78) & ","
    '       免許証有効期限
            strMeisai = strMeisai & .Cells(strTmpRow, 79) & ","
    '       車両所有者氏名（ｶﾅ）
            strMeisai = strMeisai & .Cells(strTmpRow, 80) & ","
    '       車両所有者氏名（漢字）
            strMeisai = strMeisai & .Cells(strTmpRow, 81) & ","
    '       所有権留保またはリース等
            strMeisai = strMeisai & .Cells(strTmpRow, 82)
            strMeisai = strMeisai & vbCrLf
        End With
    End If
        
    fncTempSave_NonFleetMeisai = strMeisai

    Set wsMeisai = Nothing

End Function
Public Sub subClearAll()
    Dim objAll As Object
    Dim strTagetRange As String
    
    Dim wsMeisai As Worksheet           '明細入力ワークシート
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(1, wsMeisai)       'シートオブジェクト(明細入力)
    Else
        Call subSetSheet(17, wsMeisai)      'シートオブジェクト(明細入力（ノンフリート）)
    End If
    
    '描画停止
    Application.ScreenUpdating = False
    
    For Each objAll In wsMeisai.DrawingObjects
        If objAll.Name Like "chkMeisai*" Then
            objAll.Value = xlOff
            strTagetRange = wsMeisai.Shapes(objAll.Name).TopLeftCell.Address
            wsMeisai.Range(strTagetRange) = Left(wsMeisai.Range(strTagetRange), InStr(wsMeisai.Range(strTagetRange), "/") - 1) & "/False"
        End If
    Next objAll
    
    Set wsMeisai = Nothing
    Set objAll = Nothing
    
    '描画停止
    Application.ScreenUpdating = True
        
End Sub

''明細のリスト設定
Public Sub subCmbInitialize()
    
    'フリート
    Call subCmbSet("C", "Z2", 1)       '用途車種
    Call subCmbSet("L", "AD2", 1)      '改造・不明車
    Call subCmbSet("O", "CT2", 1)      'ASV割引
    Call subCmbSet("Y", "BJ2", 1)      '車両保険の種類
    Call subCmbSet("AA", "BN2", 1)      '車両免責金額
    Call subCmbSet("AS", "CA2", 1)     '代車等セット特約
    Call subCmbSet("AE", "CD2", 1)     '対人賠償
    Call subCmbSet("AH", "CH2", 1)     '対物賠償
    Call subCmbSet("AI", "BR2", 1)     '対物免責金額
    Call subCmbSet("AK", "CL2", 1)     '人身傷害(1名)
'    Call subCmbSet("AM", "CP2", 1)     '搭乗者傷害(1名)
    Call subCmbSet("AB", "CT2", 1)     '車両全損臨費特約
    Call subCmbSet("AC", "CT2", 1)     '車両超過修理費用特約
    Call subCmbSet("AD", "CW2", 1)     '車両盗難対象外特約
    Call subCmbSet("AF", "CW2", 1)     '自損事故傷害特約
    Call subCmbSet("AG", "CW2", 1)     '無保険車事故傷害特約
    Call subCmbSet("AJ", "CT2", 1)     '対物超過修理費用特約
    Call subCmbSet("AO", "CT2", 1)     '日数払特約
    Call subCmbSet("AP", "CT2", 1)     '事業主費用特約
    Call subCmbSet("AQ", "CT2", 1)     '弁護士費用特約
    Call subCmbSet("AR", "CZ2", 1)     '従業員等限定特約
    Call subCmbSet("AT", "DC2", 1)     '車両搬送時不適用特約
    
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    Call subCmbSet("C", "Z2", 2)      '用途車種
    Call subCmbSet("L", "AD2", 2)      '改造・不明車
    Call subCmbSet("P", "DM2", 2)      'ノンフリート等級
    Call subCmbSet("Q", "AL2", 2)      '事故有係数適用期間
    Call subCmbSet("R", "CT2", 2)      'ゴールド免許割引
    Call subCmbSet("S", "DC2", 2)      '使用目的
    Call subCmbSet("T", "BB2", 2)      '年齢条件
    Call subCmbSet("U", "CW2", 2)      '高齢運転者対象外
    Call subCmbSet("V", "BF2", 2)      '運転者限定
    Call subCmbSet("W", "CT2", 2)      'ASV割引
    Call subCmbSet("AG", "BJ2", 2)     '車両保険の種類
    Call subCmbSet("AI", "BN2", 2)     '車両免責金額
    Call subCmbSet("BC", "CA2", 2)     '代車等セット特約
    Call subCmbSet("AM", "CD2", 2)     '対人賠償
    Call subCmbSet("AP", "CH2", 2)     '対物賠償
    Call subCmbSet("AQ", "BR2", 2)     '対物免責金額
    Call subCmbSet("AS", "CL2", 2)     '人身傷害(1名)
'    Call subCmbSet("AU", "CP2", 2)     '搭乗者傷害(1名)
    Call subCmbSet("AJ", "CT2", 2)     '車両全損臨費特約
    Call subCmbSet("AK", "CT2", 2)     '車両超過修理費用特約
    Call subCmbSet("AL", "CW2", 2)     '車両盗難対象外特約
    Call subCmbSet("AN", "CW2", 2)     '自損事故傷害特約
    Call subCmbSet("AO", "CW2", 2)     '無保険車事故傷害特約
    Call subCmbSet("AR", "CT2", 2)     '対物超過修理費用特約
    Call subCmbSet("AW", "CT2", 2)     '日数払特約
    Call subCmbSet("AX", "CT2", 2)     '事業主費用特約
    Call subCmbSet("AY", "CT2", 2)     '弁護士費用特約
    Call subCmbSet("AZ", "BV2", 2)     'ファミリーバイク特約
    Call subCmbSet("BA", "CT2", 2)     '個人賠償責任補償特約
    Call subCmbSet("BB", "CT2", 2)     '自動車事故特約
    Call subCmbSet("BD", "DQ2", 2)     '車両搬送時不適用特約
    
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    Call subCmbSetMeisaiPrint("X", "DG2", 2, 6)   '免許証の色
    Call subCmbSetMeisaiPrint("AE", "DJ2", 2, 18) '所有権留保またはリース等
    Call subCmbSetMeisaiPrint("F", "AH2", 2, 30)  '前契約等級
    Call subCmbSetMeisaiPrint("G", "AL2", 2, 30)  '前契約事故有係数適用期間
    
End Sub

''計算用シートのコンボボックスを設定
Public Sub subCmbSet(ByRef strTargetCol As String, ByRef strTargetCode As String, ByRef SheetFlg As Integer)
    Dim i As Integer                'ループカウント
    Dim intMaxRow As Integer        '初期の明細行数
    
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    Dim wsMeisai As Worksheet           '明細入力ワークシート
    If SheetFlg = 1 Then  'フリート
        Call subSetSheet(1, wsMeisai)       'シートオブジェクト(明細入力)
    Else
        Call subSetSheet(17, wsMeisai)      'シートオブジェクト(明細入力（ノンフリート）)
    End If
    
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    Dim wsCode As Worksheet                 '別紙コード値のオブジェクト
    If SheetFlg = 1 Then  'フリート
        Call subSetSheet(4, wsCode)         'シートオブジェクト(別紙　コード値)
    Else
        Call subSetSheet(16, wsCode)         'シートオブジェクト(別紙　コード値（ノンフリート））
    End If
    
    '総付保台数取得
    intMaxRow = Val(wsMeisai.OLEObjects("txtSouhuho").Object.Value)
    
    '初期作成されている明細行分だけループする
    With wsMeisai
        For i = 1 To intMaxRow
            '書式設定を削除
            .Range(strTargetCol & (20 + i)).Validation.Delete
            'セルをリストに変更し、該当のコード値正式名称を設定
            .Range(strTargetCol & (20 + i)).Validation.Add _
                Type:=xlValidateList, _
                Formula1:="=" & wsCode.Range(wsCode.Range(strTargetCode), wsCode.Cells(wsCode.Rows.Count, _
                            wsCode.Range(strTargetCode).Column).End(xlUp)).Address(External:=True)
            'リスト以外の文字も入力できるように設定
            If strTargetCode = "Z2" Or strTargetCode = "AD2" Then
                .Range(strTargetCol & (20 + i)).Validation.IMEMode = xlIMEModeOn
                .Range(strTargetCol & (20 + i)).Validation.ShowError = False
            End If
        Next i
    End With
    
    Set wsCode = Nothing
    
End Sub


'2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
''明細書印刷画面のコンボボックスを設定
Public Sub subCmbSetMeisaiPrint(ByRef strTargetCol As String, ByRef strTargetCode As String, ByRef SheetFlg As Integer, ByRef intHeadNum As Integer)
    Dim i As Integer                'ループカウント
    Dim intMaxRow As Integer        '初期の明細行数
    
    Dim wsMeisai As Worksheet           '明細入力ワークシート
    If SheetFlg = 1 Then  'フリート
        Call subSetSheet(18, wsMeisai)       'シートオブジェクト(明細入力)
    Else
        Call subSetSheet(19, wsMeisai)      'シートオブジェクト(明細入力（ノンフリート）)
    End If
    
    Dim wsCode As Worksheet                 '別紙コード値のオブジェクト
    If SheetFlg = 1 Then  'フリート
        Call subSetSheet(4, wsCode)         'シートオブジェクト(別紙　コード値)
    Else
        Call subSetSheet(16, wsCode)         'シートオブジェクト(別紙　コード値（ノンフリート））
    End If
    
    '総付保台数取得
    intMaxRow = 9
    
    '初期作成されている明細行分だけループする
    With wsMeisai
        For i = 1 To intMaxRow
            '書式設定を削除
            .Range(strTargetCol & (intHeadNum + i)).Validation.Delete
            'セルをリストに変更し、該当のコード値正式名称を設定
            .Range(strTargetCol & (intHeadNum + i)).Validation.Add _
                Type:=xlValidateList, _
                Formula1:="=" & wsCode.Range(wsCode.Range(strTargetCode), wsCode.Cells(wsCode.Rows.Count, _
                            wsCode.Range(strTargetCode).Column).End(xlUp)).Address(External:=True)
        Next i
    End With
    
    Set wsCode = Nothing
    
End Sub


'セル保護入力可能範囲設定
Public Sub subCellProtect(ByVal intRange As Integer)
    
    Dim i        As Integer
    Dim intCol   As Integer
    Dim strRange As String
    Dim varRange As Variant
    Dim wsMeisai  As Worksheet          '明細入力ワークシート
    Dim wsSetting As Worksheet          '各種設定ワークシート
    
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(1, wsMeisai)       'シートオブジェクト(明細入力)
    Else
        Call subSetSheet(17, wsMeisai)      'シートオブジェクト(明細入力（ノンフリート）)
    End If
    Call subSetSheet(5, wsSetting)      'シートオブジェクト(別紙　各種設定)
    
    strRange = ""
    intCol = 0
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        varRange = Array("$C$21:$C$21", "$E$21:$O$21", "$Y$21:$AT$21")  '入力可能セル"$X範囲（料率クラスや合計保険料等は入力不可）
    Else    'ノンフリート
        varRange = Array("$C$21:$C$21", "$E$21:$W$21", "$AG$21:$BD$21")
    End If
    
    'セル範囲設定が残っている場合、削除
    If wsMeisai.Protection.AllowEditRanges.Count = 0 Then
    Else
        wsMeisai.Protection.AllowEditRanges.item(1).Delete
    End If
    
    '入力可能セル範囲を総付保台数分広げる
    For i = 0 To UBound(varRange)
        intCol = Right(varRange(i), 2)
        intCol = intCol + intRange - 1
        varRange(i) = Left(varRange(i), Len(varRange(i)) - 2) & intCol
        
        strRange = strRange & "," & varRange(i)
    Next i
    
    strRange = Right(strRange, Len(strRange) - 1)
    
    '入力可能セル範囲を設定
    wsMeisai.Protection.AllowEditRanges.Add _
                Title:="EntryOK", _
                Range:=wsMeisai.Range(strRange)
    
    Set wsMeisai = Nothing
    Set wsSetting = Nothing
    
End Sub

'オブジェクトのNoを取得
Public Function fncGetObjectNo(strObjectFullName As String, strObjectName As String) As String
    fncGetObjectNo = Right(strObjectFullName, Len(strObjectFullName) - Len(strObjectName))
End Function

'カンマ削除
Public Function fncTrimComma(strVal As String) As String
    fncTrimComma = Replace(strVal, ",", "")
End Function


