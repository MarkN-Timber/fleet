Attribute VB_Name = "modErrorCheckFunctions"
Option Explicit

''基本エラー処理群

Public Function fncNeedCheck(ByVal strValue As String) As String
'関数名：fncNeedCheck
'内容　：必須チェック
'引数　：
'        strValue       = 入力内容

    If IsNull(strValue) Then
        fncNeedCheck = " 必須入力項目です。入力してください。"
    Else
        If strValue = "" Then
            fncNeedCheck = " 必須入力項目です。入力してください。"
        Else
            fncNeedCheck = ""
        End If
    End If
End Function


Public Function fncCommaCheck(ByVal strValue As String) As String
'関数名：fncCommaCheck
'内容　：,(カンマ)チェック
'引数　：
'        strValue       = 入力内容

    If strValue Like "*,*" Then
        fncCommaCheck = " 「，（カンマ）」は入力できません。"
    Else
        fncCommaCheck = ""
    End If
End Function


Public Function fncDateCheck(ByVal strValue As String, Optional blnWareki As Boolean = False, Optional blnSyodoTouroku As Boolean = False) As String
'関数名：fncDateCheck
'内容　：日付チェック
'引数　：
'        strValue       = 入力内容
'        blnWareki      = 和暦フラグ
    If blnWareki Then
    Dim strSeireki As String
        If strValue Like "*元年*" Then
            strValue = Left(strValue, InStr(strValue, "元") - 1) & "1" & Mid(strValue, InStr(strValue, "元") + 1)
        End If
        '新元号対応↓
        If IsNumeric(Mid(strValue, 3, InStr(strValue, "年") - 3)) And _
        IsNumeric(Mid(strValue, InStr(strValue, "年") + 1, InStr(strValue, "月") - InStr(strValue, "年") - 1)) And _
        (Mid(strValue, InStr(strValue, "年") + 1, InStr(strValue, "月") - InStr(strValue, "年") - 1)) >= 1 And _
        (Mid(strValue, InStr(strValue, "年") + 1, InStr(strValue, "月") - InStr(strValue, "年") - 1)) <= 12 Then
            If blnSyodoTouroku Then
                If strValue Like "大正*年*月" Then
                    If strValue Like "*7月*" Then
                        strValue = strValue & "30日"
                    ElseIf strValue Like "*12月*" Then
                        strValue = strValue & "24日"
                    Else
                        strValue = strValue & "1日"
                    End If
                ElseIf strValue Like "昭和*年*月" Then
                    If strValue Like "*12月*" Then
                        strValue = strValue & "25日"
                    ElseIf strValue Like "*1月*" Then
                        strValue = strValue & "7日"
                    Else
                        strValue = strValue & "1日"
                    End If
                ElseIf strValue Like "平成*" Then
                    If strValue Like "*1月*" Then
                        strValue = strValue & "8日"
                    ElseIf strValue Like "*4月*" Then
                        strValue = strValue & "30日"
                    Else
                        strValue = strValue & "1日"
                    End If
                ElseIf strValue Like "嗚呼*" Then
                    strValue = strValue & "1日"
                Else
                    fncDateCheck = " 年月日を確認のうえ、正しく入力してください。"
                End If
            End If

            If IsNumeric(Mid(strValue, InStr(strValue, "月") + 1, InStr(strValue, "日") - InStr(strValue, "月") - 1)) Then
                If fncDateCheck = "" Then
                    If strValue Like "大正*年*月*日" Then
                        If CDate(strValue) >= CDate("1912/07/30") And _
                           CDate(strValue) <= CDate("1926/12/24") Then
                            fncDateCheck = ""
                        Else
                            fncDateCheck = " 年月日を確認のうえ、正しく入力してください。"
                        End If
                    ElseIf strValue Like "昭和*年*月*日" Then
                        If CDate(strValue) >= CDate("1926/12/25") And _
                           CDate(strValue) <= CDate("1989/01/07") Then
                            fncDateCheck = ""
                        Else
                            fncDateCheck = " 年月日を確認のうえ、正しく入力してください。"
                        End If
                    ElseIf strValue Like "平成*年*月*日" Then
                        If CDate(strValue) >= CDate("1989/01/08") And _
                           CDate(strValue) <= CDate("2019/04/30") Then
                            fncDateCheck = ""
                        Else
                            fncDateCheck = " 年月日を確認のうえ、正しく入力してください。"
                        End If
                    ElseIf strValue Like "嗚呼*年*月*日" Then
                        strSeireki = Mid(strValue, 3, InStr(strValue, "年") - 3)
                        strSeireki = strSeireki + 2018 & Mid(strValue, InStr(strValue, "年"))
                        If CDate(strSeireki) >= CDate("2019/05/01") Then
'                        If fncToSeireki(strValue, 8) >= CDate("2019/05/01") Then
                            fncDateCheck = ""
                        Else
                            fncDateCheck = " 年月日を確認のうえ、正しく入力してください。"
                        End If
                    Else
                        fncDateCheck = " 年月日を確認のうえ、正しく入力してください。"
                    End If
                End If
            Else
                fncDateCheck = " 年月日を確認のうえ、正しく入力してください。"
            End If
        Else
            fncDateCheck = " 年月日を確認のうえ、正しく入力してください。"
        End If
        '新元号対応↑
    Else
        If IsDate(strValue) = False Then
            fncDateCheck = " 年月日を確認のうえ、正しく入力してください。"
        Else
'            If CDate(strValue) >= "1989/01/08" Then
            If CDate(strValue) >= "1912/07/30" Then
                fncDateCheck = ""
            Else
                fncDateCheck = " 年月日を確認のうえ、正しく入力してください。"
            End If
        End If
    End If
    
End Function


Public Function fncShikiCheck(ByVal strValue As String) As String
'関数名：fncShikiCheck
'内容　：保険始期チェック
'引数　：
'        strValue       = 入力内容

    If CDate(strValue) < "2019/01/01" Then
        fncShikiCheck = " 保険始期が2018年12月31日以前の契約には使用できません。"
    End If

End Function


'Public Function fncNumCheck(ByVal strValue As String) As String
''関数名：fncNumCheck
''内容　：数字チェック
''引数　：
''        strValue       = 入力内容
'
'    If strValue Like "*[!0-9]*" Then
'        fncNumCheck = " 数字のみを入力してください。"
'    Else
'        fncNumCheck = ""
'    End If
'End Function
'
Public Function fncDecimalCheck(ByVal strValue As String) As String
'関数名：fncNumCheck
'内容　：数字チェック(マイナス、少数点入り)
'引数　：
'        strValue       = 入力内容
    If IsNumeric(strValue) Then
        Dim strNum As String
        If strValue = Int(strValue) Then
            strNum = strValue
        Else
            strNum = Replace(strValue, ".", "")
        End If
        If strNum Like "[0-9]*" Or strNum Like "-[0-9]*" Then
            fncDecimalCheck = ""
        Else
            fncDecimalCheck = " 数字(半角)のみを入力してください。"
        End If
    Else
        fncDecimalCheck = " 数字(半角)のみを入力してください。"
    End If
End Function

Public Function fncNumCheck(ByVal strValue As String, Optional intNum As Integer) As String
'関数名：fncNumCheck
'内容　：数字チェック
'引数　：
'        strValue       = 入力内容
    Dim strNum As String
    strNum = strValue
    If intNum = 1 Then
        If IsNumeric(strValue) Then
            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
            'If strValue = Int(strValue) Then
            If strValue = CStr(Int(strValue)) Then
            Else
                strNum = Replace(strValue, ".", "")
            End If
        End If
    Else
    End If
    If strNum Like "*[!0-9]*" Then
        fncNumCheck = " 数字(半角)のみを入力してください。"
    Else
        fncNumCheck = ""
    End If
    
End Function


Public Function fncNumRangeCheck(ByVal intValue As Double, ByVal intRangeMin As Double, ByVal intRangeMax As Double) As String
'関数名：fncNumRangeCheck
'内容　：数値チェック
'引数　：
'        intValue           = 入力内容
'        intRabgeMin        = 最小値
'        intRabgeMax        = 最大値
    
    If intValue < intRangeMin Or intValue > intRangeMax Then
        fncNumRangeCheck = " 指定された範囲の数値を入力してください。"
    Else
        fncNumRangeCheck = ""
    End If
End Function


'2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
Public Function fncNonfleetSoufuhodaisuCheck(ByVal intValue As Double, ByVal intRangeMax As Double) As String
'関数名：fncNumRangeCheck
'内容　：総付保台数エラーチェック
'引数　：
'        intValue           = 入力内容
'        intRabgeMin        = 最小値
'        intRabgeMax        = 最大値
    
    If intValue > intRangeMax Then
        fncNonfleetSoufuhodaisuCheck = " 総付保台数に10台以上は入力できません。"
    Else
        fncNonfleetSoufuhodaisuCheck = ""
    End If
End Function


Public Function fncKetaCheck(ByVal strValue As String, ByVal intDigit As Integer, ByVal strType As String) As String
'関数名：fncKetaCheck
'内容　：桁数チェック
'引数　：
'        intValue           = 入力内容
'        intDigit           = 桁制限値
'        strType            = 比較方法

    If strType = "<" Then
        If Len(strValue) < intDigit Then
            fncKetaCheck = " 入力桁数が足りません。正しく入力してください。"
        Else
            fncKetaCheck = ""
        End If
    ElseIf strType = ">" Then
        If Len(strValue) > intDigit Then
            fncKetaCheck = " 入力できる桁数を超えています。正しく入力してください。"
        Else
            fncKetaCheck = ""
        End If
    ElseIf strType = "=" Then
        If Not Len(strValue) = intDigit Then
            fncKetaCheck = " 指定された桁数以外は入力できません。"
        Else
            fncKetaCheck = ""
        End If
    End If
End Function


Public Function fncZenkakuCheck(ByVal strValue As String) As String
'関数名：fncZenkakuCheck
'内容　：全角チェック
'引数　：
'        strValue           = 入力内容

    If Len(strValue) * 2 <> LenB(StrConv(strValue, vbFromUnicode)) Then
        fncZenkakuCheck = " 全角文字のみを入力してください。"
    Else
        fncZenkakuCheck = ""
    End If
End Function


Public Function fncHankakuCheck(ByVal strValue As String) As String
'関数名：fncHankakuCheck
'内容　：半角英数字チェック
'引数　：
'        strValue           = 入力内容

    If Len(strValue) <> LenB(StrConv(strValue, vbFromUnicode)) Then
        fncHankakuCheck = " 半角英数字のみを入力してください。"
    Else
        fncHankakuCheck = ""
    End If
End Function


Public Function fncListCheck(ByVal strValue As String, ByVal strListCell As String) As String
'関数名：fncListCheck
'内容　：リストチェック
'引数　：
'        strValue           = 入力内容
'        strListCell        =リストのセル番号

    Dim wsCode As Worksheet
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(4, wsCode)         'シートオブジェクト(別紙　コード値)
    Else
        Call subSetSheet(16, wsCode)         'シートオブジェクト(別紙　コード値（ノンフリート））
    End If
    
    If WorksheetFunction.CountIf(wsCode.Range(wsCode.Range(strListCell), wsCode.Cells(wsCode.Rows.Count, _
                                 wsCode.Range(strListCell).Column).End(xlUp)), strValue) = 0 Then
        fncListCheck = " 指定された値を入力してください。"
    Else
        fncListCheck = ""
    End If
    
    Set wsCode = Nothing
    
End Function


''個別エラー処理群

Public Function fncCodeCheck(ByVal strValue As String, ByVal strListCell As String) As String
'関数名：fncCodeCheck
'内容　：コード値チェック
'引数　：
'        strValue           = 入力内容
'        strListCell        =リストのセル番号

    Dim wsCode As Worksheet
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(4, wsCode)         'シートオブジェクト(別紙　コード値)
    Else
        Call subSetSheet(16, wsCode)         'シートオブジェクト(別紙　コード値（ノンフリート））
    End If
    
    If WorksheetFunction.CountIf(wsCode.Range(wsCode.Range(strListCell), wsCode.Cells(wsCode.Rows.Count, _
                                 wsCode.Range(strListCell).Column).End(xlUp)), strValue) = 0 Then
        fncCodeCheck = " 指定された値を入力してください。"
    Else
        fncCodeCheck = ""
    End If
    
    Set wsCode = Nothing
    
End Function


Public Function fncOneOrBlankCheck(ByVal strValue As String) As String
'関数名：fncOneOrBlankCheck
'内容　：1またはブランクチェック
'引数　：
'        strValue           = 入力内容

    fncOneOrBlankCheck = ""
    If IsNull(strValue) = False Then
        If strValue <> "" Then
            If strValue <> "1" Then
                fncOneOrBlankCheck = " 「1」以外は入力できません。"
            End If
        End If
    End If
End Function


Public Function fncWarekiCheck(ByVal strValue As String, Optional intKeta As Integer) As String
'関数名：fncWarekiCheck
'内容　：和暦チェック
'引数　：
'        strValue           = 入力内容

    If strValue Like "大正*年*月*日" Then
        fncWarekiCheck = ""
    ElseIf strValue Like "昭和*年*月*日" Then
        fncWarekiCheck = ""
    ElseIf strValue Like "平成*年*月*日" Then
        fncWarekiCheck = ""
    ElseIf strValue Like "嗚呼*年*月*日" Then
        fncWarekiCheck = ""
    Else
        If intKeta = 6 Then
            fncWarekiCheck = " 和暦で入力してください。（元号X年X月）"
        Else
            fncWarekiCheck = " 和暦で入力してください。（元号X年X月X日）"
        End If
    End If
End Function


''関連エラー処理群

Public Function fncHknSyuruiCheck(ByVal strHknSyurui As String, ByVal strSyaryoMskGaku As String, ByVal strDaisyaToku As String, _
                                  ByVal strHknZnsnToku As String, ByVal strSyaryoTonanToku As String, ByVal strSyaryoTyoukaToku As String, _
                                  Optional ByVal strHknKingaku As String = "NULL") As String
'関数名：fncHknSyuruiCheck
'内容　：車両保険エラー
'引数　：
'        strHknSyurui           =保険の種類
'        strHknKingaku          =車両保険金額
'        strSyaryoMskGaku       =車両免責金額
'        strDaisyaToku          =代車等セット特約
'        strHknZnsnToku         =車両全損臨費特約
'        strSyaryoTonanToku     =車両盗難対象外特約
'        strSyaryoTyoukaToku    =車両超過修理費用特約
    
    Dim blnErrFlg As Boolean
    blnErrFlg = False
    
    fncHknSyuruiCheck = ""
    If strHknSyurui = "" Then
        If strHknKingaku <> "NULL" Then
            If (strHknKingaku <> "" Or strSyaryoMskGaku <> "" Or strDaisyaToku <> "") Or _
               (CBool(fncTekiyouChange(strHknZnsnToku)) = True Or CBool(fncTaisyoChange(strSyaryoTonanToku)) = True Or _
                CBool(fncTekiyouChange(strSyaryoTyoukaToku)) = True) Then
                
                blnErrFlg = True
            End If
        Else
            If (strSyaryoMskGaku <> "" Or strDaisyaToku <> "") Or _
               (CBool(fncTekiyouChange(strHknZnsnToku)) = True Or CBool(fncTaisyoChange(strSyaryoTonanToku)) = True Or _
                CBool(fncTekiyouChange(strSyaryoTyoukaToku)) = True) Then
                
                blnErrFlg = True
            End If
        End If
        
        If blnErrFlg Then
            fncHknSyuruiCheck = " 車両保険の種類を選択してください。"
        End If
    Else
        If strHknKingaku <> "NULL" Then
            If strHknKingaku = "" Or strSyaryoMskGaku = "" Then
                fncHknSyuruiCheck = " 車両保険金額および車両免責金額を入力してください。"
            End If
        Else
            If strSyaryoMskGaku = "" Then
                fncHknSyuruiCheck = " 車両免責金額を入力してください。"
            End If
        End If
        
    End If
End Function

Public Function fncTaijinBaisyoCheck(ByVal strTaijinBaisyo As String, ByVal strJisonJikoToku As String, _
                                     ByVal strZinshinSyougai_1Mei As String, ByVal strMuhokenToku As String) As String
'関数名：fncTaijinBaisyoCheck
'内容　：対人賠償エラー
'引数　：
'        strTaijinBaisyo            =対人賠償
'        strJisonJikoToku           =自損事故傷害特約
'        strZinshinSyougai_1Mei     =人身傷害(1名)
'        strDaisyaToku              =代車等セット特約
'        strMuhokenToku             =無保険車事故傷害特約

    fncTaijinBaisyoCheck = ""
    If strTaijinBaisyo = "" Then
'        fncTaijinBaisyoCheck = " 対人賠償を選択してください。"
    ElseIf CBool(fncTaisyoChange(strTaijinBaisyo)) Then
        If CBool(fncTaisyoChange(strJisonJikoToku)) Then
            fncTaijinBaisyoCheck = " 対人対象外のとき自損事故対象外特約は付帯できません。"
        End If
        If CBool(fncTaisyoChange(strMuhokenToku)) Then
            If Not fncTaijinBaisyoCheck = "" Then
                fncTaijinBaisyoCheck = fncTaijinBaisyoCheck & vbCrLf & " 対人対象外のとき無保険車対象外特約は付帯できません。"
            Else
                fncTaijinBaisyoCheck = " 対人対象外のとき無保険車対象外特約は付帯できません。"
            End If
        End If
    End If
    
    If CBool(fncTaisyoChange(strZinshinSyougai_1Mei)) = False And CBool(fncTaisyoChange(strJisonJikoToku)) Then
        If strZinshinSyougai_1Mei <> "" Then
            If fncTaijinBaisyoCheck = "" Then
                fncTaijinBaisyoCheck = " 自損事故対象外特約を付帯する場合、人身傷害は付帯できません。"
            Else
                fncTaijinBaisyoCheck = fncTaijinBaisyoCheck & vbCrLf & " 自損事故対象外特約を付帯する場合、人身傷害は付帯できません。"
            End If
        End If
    End If
    
End Function


Public Function fncTaibutsuBaisyo(ByVal strTaibutsuBaisyo As String, ByVal strTaibutsuMskGaku As String, _
                                  ByVal strTaibutsuTyoukaToku As String) As String
'関数名：fncTaibutsuBaisyo
'内容　：対物賠償エラー
'引数　：
'        strTaibutsuBaisyo              =対物賠償
'        strTaibutsuMskGaku             =対物免責金額
'        strTaibutsuTyoukaToku          =対物超過修理費用特約

    fncTaibutsuBaisyo = ""
    If strTaibutsuBaisyo = "" Then
'        fncTaibutsuBaisyo = " 対物賠償を選択してください。"
    Else
        If strTaibutsuBaisyo = "対象外" And (strTaibutsuMskGaku <> "" Or CBool(fncTekiyouChange(strTaibutsuTyoukaToku)) = True) Then
            fncTaibutsuBaisyo = " 対物対象外のとき、対物免責金額および対物超過修理費用特約は入力不要です。"
        ElseIf strTaibutsuBaisyo <> "対象外" And strTaibutsuMskGaku = "" Then
            fncTaibutsuBaisyo = " 対物免責金額を選択してください。"
        End If
    End If
End Function


Public Function fncZinshinSyougai(ByVal strZinshinSyougai_1Mei As String, ByVal strZinshinSyougai_1Jiko As String) As String
'関数名：fncZinshinSyougai
'内容　：人身傷害エラー
'引数　：
'        strZinshinSyougai_1Mei             =人身傷害(1名)
'        strZinshinSyougai_1Jiko            =人身傷害(1事故)

    fncZinshinSyougai = ""
    If strZinshinSyougai_1Mei = "" Then
'        fncZinshinSyougai = " 人身傷害を選択してください。"
    ElseIf strZinshinSyougai_1Mei = "対象外" And strZinshinSyougai_1Jiko <> "" Then
        fncZinshinSyougai = " 人傷対象外のとき、人身傷害（１事故）は入力不要です。"
    End If
End Function


Public Function fncTouzyouSyougai(ByVal strTouzyouSyougai_1Mei As String, ByVal strTouzyouSyougai_1Mei_Taisyougai As String, _
                                  ByVal strTouzyouSyougai_1Jiko As String, ByVal strNissuToku As String)
'関数名：fncTouzyouSyougai
'内容　：搭乗者傷害エラー
'引数　：
'        strTouzyouSyougai_1Mei             =搭乗者傷害(1名)
'        strTouzyouSyougai_1Mei_Taisyougai            =搭乗者傷害(1事故)
'        strNissuToku                       =日数払特約

    fncTouzyouSyougai = ""
    If strTouzyouSyougai_1Mei = "" Then
        If strTouzyouSyougai_1Mei_Taisyougai = False Then
            fncTouzyouSyougai = " 搭乗者傷害（１名）保険金額か対象外のいずれかを入力してください。"
        Else
            If strTouzyouSyougai_1Jiko <> "" Or CBool(fncTekiyouChange(strNissuToku)) Then
                fncTouzyouSyougai = " 搭傷対象外のとき、搭乗者傷害（１事故）および日数払特約は入力不要です。"
            End If
        End If
'        fncTouzyouSyougai = " 搭乗者傷害を選択してください。"
    ElseIf IsNumeric(strTouzyouSyougai_1Mei) Then
        If strTouzyouSyougai_1Mei Mod 100 <> 0 Or strTouzyouSyougai_1Mei = 0 Then
            fncTouzyouSyougai = " 搭乗者傷害（１名）は100万円単位で入力してください。"
        ElseIf strTouzyouSyougai_1Mei > 5000 Then
            fncTouzyouSyougai = " 搭乗者傷害（１名）は5,000万円以下で入力してください。"
        ElseIf strTouzyouSyougai_1Mei_Taisyougai Then
            fncTouzyouSyougai = " 搭乗者傷害（１名）保険金額か対象外のいずれかを入力してください。"
        End If
    Else
        fncTouzyouSyougai = " 搭乗者傷害（１名）保険金額には数字のみ入力してください。"
    End If

End Function


Public Function fncTekiyouChange(ByVal strValue As String) As String
'関数名：fncTekiyouChange
'内容　：入力項目が「適用する」のものをTrue/Falseに変換
'引数　：
'        strValue           = 入力内容

    If strValue = "適用する" Or strValue = "True" Then
        fncTekiyouChange = "True"
    Else
        fncTekiyouChange = "False"
    End If
End Function


Public Function fncTaisyoChange(ByVal strValue As String) As String
'関数名：fncTaisyoChange
'内容　：入力項目が「対象外」のものをTrue/Falseに変換
'引数　：
'        strValue           = 入力内容

    If strValue = "対象外" Or strValue = "True" Then
        fncTaisyoChange = "True"
    Else
        fncTaisyoChange = "False"
    End If
End Function


Public Function fncGenteiChange(ByVal strValue As String) As String
'関数名：fncGenteiChange
'内容　：入力項目が「限定」のものをTrue/Falseに変換
'引数　：
'        strValue           = 入力内容

    If strValue = "限定" Or strValue = "True" Then
        fncGenteiChange = "True"
    Else
        fncGenteiChange = "False"
    End If
End Function


'2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
Public Function fncHokenSyuruiCheck(ByVal strHokenSyurui As String, ByVal strHoujin As String) As String
'関数名：fncNonfleetTawariCheck
'内容　：保険種類エラー
'引数　：
'        strHokenSyurui = 保険種類
'        strKojin = 被保険者

    If strHokenSyurui = "個人用総合自動車保険" And strHoujin = "True" Then
        fncHokenSyuruiCheck = " 個人用総合自動車保険は個人契約のみになります。"
    Else
        fncHokenSyuruiCheck = ""
    End If
End Function


'2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
Public Function fncNonfleetTawariCheck(ByVal strNonfleetTawari As String, ByVal strSouFuhoDaisu As String) As String
'関数名：fncNonfleetTawariCheck
'内容　：ﾉﾝﾌﾘｰﾄ多数割引エラー
'引数　：
'        strNonfleetTawari = ﾉﾝﾌﾘｰﾄ多数割引
'        strSouFuhoDaisu = 総付保台数

    fncNonfleetTawariCheck = ""
    If IsNumeric(strSouFuhoDaisu) Then
        If strSouFuhoDaisu >= 3 Then
            If strNonfleetTawari = "" Then
                fncNonfleetTawariCheck = " 総付保台数が３台以上の場合、ノンフリート多数割引が適用できます。"
            End If
        End If
    End If
End Function


'2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
Public Function fncFamilyBikeCheck(ByVal strFamilyBike As String, ByVal strJinshinSyougaiIchimei As String) As String
'関数名：fncNonfleetTawariCheck
'内容　：ファミリーバイク特約エラー
'引数　：
'        strFamilyBike = ファミリーバイク特約
'        strJinshinSyougaiIchimei = 人身傷害（1名）

    If strFamilyBike = "人身" And strJinshinSyougaiIchimei = "対象外" Then
        fncFamilyBikeCheck = " 人傷対象外のとき、ファミリーバイク特約「人身」は選択できません。"
    Else
        fncFamilyBikeCheck = ""
    End If
End Function


'2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
Public Function fncJidoushaJikoCheck(ByVal strJidoushaJiko As String, ByVal strJinshinSyougaiIchimei As String) As String
'関数名：fncNonfleetTawariCheck
'内容　：自動車事故特約エラー
'引数　：
'        strJidoushaJiko = 自動車事故特約
'        strJinshinSyougaiIchimei = 人身傷害（1名）

    If (strJidoushaJiko = "True" Or strJidoushaJiko = "適用する") And strJinshinSyougaiIchimei = "対象外" Then
        fncJidoushaJikoCheck = " 人傷対象外のとき、自動車事故特約は付帯できません。"
    Else
        fncJidoushaJikoCheck = ""
    End If
End Function


'2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
Public Function fncToukyuCheck1(ByVal strUketsukeKbn As String, ByVal strNonFleetToukyu As String) As String
'関数名：fncToukyuCheck1
'内容　：等級エラー１
'引数　：
'        strUketsukeKbn = 受付区分
'        strNonFleetToukyu = ノンフリート等級

    If strUketsukeKbn = "1" And strNonFleetToukyu <> "６Ｓ等級" Then
        fncToukyuCheck1 = " 受付区分が新規のとき、6S等級以外は選択できません。"
    Else
        fncToukyuCheck1 = ""
    End If
End Function


'2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
Public Function fncToukyuCheck2(ByVal strUketsukeKbn As String, ByVal strNonFleetToukyu As String) As String
'関数名：fncToukyuCheck2
'内容　：等級エラー２
'引数　：
'        strUketsukeKbn = 受付区分
'        strNonFleetToukyu = ノンフリート等級

    If strUketsukeKbn = "3" And (strNonFleetToukyu = "６Ｓ等級" Or strNonFleetToukyu = "７Ｓ等級") Then
        fncToukyuCheck2 = " 受付区分が継続のとき、6S等級、7S等級は選択できません。"
    Else
        fncToukyuCheck2 = ""
    End If
End Function



