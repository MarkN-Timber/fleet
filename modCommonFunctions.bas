Attribute VB_Name = "modCommonFunctions"
Option Explicit

'共通関数

'グローバル変数
Public FleetTypeFlg      As Integer     'フリート、ノンフリート判定
Public strTxtOther       As String      'その他料率文字列
Public blnSaveFlg        As Boolean     '管理者機能の保存フラグ
Public blnCloseFlg       As Boolean     'ブックの閉じる無効フラグ
Public blnSyaryouOpenFlg As Boolean     '車両情報取込ファイルの確認フラグ
Public blnFleetBtnFlg    As Boolean     'Top画面「フリート契約」の押下フラグ
Public blnNonFleetBtnFlg As Boolean     'Top画面「ノンフリート明細付契約」の押下フラグ
Public blnChouhyouflg    As Boolean     'Top画面「帳票出力」の押下フラグ
Public MeisaiBackFlg     As Integer     '明細入力画面「戻る」の押下時　0:帳票選択画面 1:共通項目画面
Public blnMoushikomiflg  As Boolean     '明細書印刷画面の押下フラグ
Public strTextName       As String      '試算ファイル名

Public Function fncFormSave(ByVal intForm As Integer, ByVal strSaveContent As String) As Boolean
'関数名：fncFormSave
'内容　：画面の状態をワークシートに保存
'引数　：
'        intForm        = 1 :共通項目画面
'                         2 :補償内容セット（一括）画面
'        strSaveContent = "":保存内容
    Dim intSheetType As Integer
    Dim varSaveConetent As Variant
    Dim i As Integer

    fncFormSave = False
    i = 0

    If intForm = 1 Then
        intSheetType = 2
    Else
        intSheetType = 3
    End If
    
    Dim wstSave As Worksheet
    Call subSetSheet(intSheetType, wstSave)         'シートオブジェクト(別紙　共通項目,別紙　補償内容セット（一括）)
    
    varSaveConetent = Split(strSaveContent, "/")
    
    Do While i < UBound(varSaveConetent)
        wstSave.Cells(2, 1 + i).Value = varSaveConetent(i)
        i = i + 1
    Loop
    
    Set wstSave = Nothing
    
    fncFormSave = True
    
End Function

Public Function fncFormRef(ByVal intForm As Integer, ByRef strSaveContent As String) As Boolean
'関数名：fncFormRef
'内容　：ワークシートに保存されている値を取得
'引数　：
'        intForm        = 1 :共通項目画面
'                         2 :補償内容セット（一括）画面
'        strSaveContent = "":保存内容

    Dim intSheetType As Integer
    Dim i As Integer
    Dim intLastCol As Integer
    
    fncFormRef = True
    
    i = 1
    intLastCol = 0
    
    If intForm = 1 Then
        intSheetType = 2
    Else
        intSheetType = 3
    End If
    
    Dim wstSave As Worksheet
    Call subSetSheet(intSheetType, wstSave)     'シートオブジェクト(別紙　共通項目,別紙　補償内容セット（一括）)
    
    Do Until wstSave.Cells(1, i) = ""
        i = i + 1
    Loop

    intLastCol = i
        
    For i = 1 To intLastCol - 1
        If Not wstSave.Cells(2, i) = "" Then
            fncFormRef = False
        End If
    
        strSaveContent = strSaveContent + wstSave.Cells(2, i) + "/"
    Next i
    
    Set wstSave = Nothing
    
End Function
Public Function fncMoushikomiFormRef(ByVal intForm As Integer, ByRef strSaveContent As String) As Boolean
'関数名：fncMoushikomiFormRef
'内容　：申込書印刷画面用ワークシートに保存されている値を取得
'引数　：
'        intForm　　　　= 8 :申込書印刷画面　'2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
'                       = 6 :テキスト内容(共通)
'        strSaveContent = "":保存内容

    Dim intSheetType As Integer
    Dim i As Integer
    Dim intLastCol As Integer
    
    fncMoushikomiFormRef = True
    
    i = 1
    intLastCol = 0
    
    If intForm = 8 Then
        intSheetType = 8 '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        If FleetTypeFlg = 1 Then
            intLastCol = 29 'フリート（申込書印刷画面の項目数は24）
        Else
            intLastCol = 29 'ノンフリート（申込書印刷画面の項目数は29）
        End If
        
    ElseIf intForm = 6 Then
        intSheetType = 6 'テキスト内容(共通)
        intLastCol = 38 '共通項目数は38
    End If
    
    Dim wstSave As Worksheet
    Call subSetSheet(intSheetType, wstSave)     'シートオブジェクト(申込書印刷画面)
        
    For i = 1 To intLastCol
        strSaveContent = strSaveContent + wstSave.Cells(1, i) + "/"
    Next i

    If strSaveContent = String(intLastCol, "/") Then
    Else
        fncMoushikomiFormRef = False
    End If
    Set wstSave = Nothing

End Function

'コード値検索
Public Function fncFindCode(ByVal strContent As String, ByVal strRow As String) As String
    Dim strFindRow As String
    Dim strFindColumn As String
    Dim rgFindNot As Range
    
    Dim wsCode As Worksheet
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(4, wsCode)         'シートオブジェクト(別紙　コード値)
    Else
        Call subSetSheet(16, wsCode)         'シートオブジェクト(別紙　コード値（ノンフリート））
    End If
    
    With wsCode
        
        Set rgFindNot = .Columns(strRow).Find(strContent, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True, MatchByte:=True)
        
        If rgFindNot Is Nothing Then
            fncFindCode = strContent
        Else
            strFindRow = rgFindNot.Row
            strFindColumn = rgFindNot.Column - 1
            
            fncFindCode = .Cells(Val(strFindRow), Val(strFindColumn)).Value
        End If

        If fncFindCode = "" Then
            fncFindCode = strContent
        End If

    End With
    
    Set wsCode = Nothing
    Set rgFindNot = Nothing

End Function

'コード値正式名称検索
Public Function fncFindName(ByVal strContent As String, ByVal strRow As String) As String
    Dim strFindRow      As String
    Dim strFindColumn   As String
    Dim rgFindNot      As Range

    Dim wsCode As Worksheet
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(4, wsCode)         'シートオブジェクト(別紙　コード値)
    Else
        Call subSetSheet(16, wsCode)         'シートオブジェクト(別紙　コード値（ノンフリート））
    End If
    
    With wsCode
        Set rgFindNot = .Columns(strRow).Find(strContent, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True, MatchByte:=True)
        
        If rgFindNot Is Nothing Then
            fncFindName = strContent
        Else
            strFindRow = .Columns(strRow).Find(strContent, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True, MatchByte:=True).Row
            strFindColumn = .Columns(strRow).Find(strContent, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True, MatchByte:=True).Column + 1
            
            fncFindName = .Cells(Val(strFindRow), Val(strFindColumn)).Value
            
        End If
        
        If fncFindName = "" Then
            fncFindName = strContent
        End If
        
    End With
    
    Set wsCode = Nothing
    
End Function


Public Function fncToWareki(ByVal strDate As String, ByVal intKeta As Integer) As String
    Dim strSaveDate   As String
    Dim intSaveDate   As Integer
    Dim strWareki     As String
    Dim strKetaErr    As String
    Dim strDateErr    As String
        
    strSaveDate = ""
    intSaveDate = 0
    strWareki = ""
    strKetaErr = ""
    strDateErr = ""
    
    fncToWareki = strDate
    
    If IsNumeric(strDate) = False Then
    Else
        strKetaErr = fncKetaCheck(strDate, 8, "=")
        
        If strKetaErr = "" Then
            strSaveDate = Format(strDate, "####/##/##")
            strDateErr = fncDateCheck(strSaveDate)
            
            If strDateErr = "" Then
                
                intSaveDate = Mid(strDate, 1, 4)
            
                If CDate(strSaveDate) >= "1912/07/30" And CDate(strSaveDate) <= "1926/12/24" Then
                    strWareki = "大正"
                    intSaveDate = intSaveDate - 1911
                
                ElseIf CDate(strSaveDate) >= "1926/12/25" And CDate(strSaveDate) <= "1989/01/07" Then
                    strWareki = "昭和"
                    intSaveDate = intSaveDate - 1925
                
                ElseIf CDate(strSaveDate) >= "1989/01/08" And CDate(strSaveDate) <= "2019/04/30" Then
                    strWareki = "平成"
                    intSaveDate = intSaveDate - 1988
                ElseIf CDate(strSaveDate) >= "2019/05/01" Then
                    strWareki = "嗚呼"
                    intSaveDate = intSaveDate - 2018
                End If
                
                fncToWareki = strWareki & CStr(intSaveDate) & "年" & Format(CDate(strSaveDate), "mm") & "月" & Format(CDate(strSaveDate), "dd") & "日"
                
                If intKeta = 8 Then
                    fncToWareki = strWareki & CStr(intSaveDate) & "年" & Format(CDate(strSaveDate), "mm") & "月"
                End If
                
            Else
            End If
        Else
        End If
    End If
    
End Function



'和暦→西暦　変換
Public Function fncToSeireki(ByVal strDate As String, ByVal intKeta As Date, Optional ByVal blnNengappiflg As Boolean, Optional ByVal blnMaezeroflg As Boolean) As String
    Dim intSaveDate    As Integer
    If strDate Like "*年*月*日" Then
        If strDate Like "*元年*" Then
            strDate = Left(strDate, InStr(strDate, "元") - 1) & "1" & Mid(strDate, InStr(strDate, "元") + 1)
        End If
        
        intSaveDate = Mid(strDate, 3, InStr(strDate, "年") - 3)
        
        If Left(strDate, 2) = "大正" Then
            intSaveDate = intSaveDate + 1911
        ElseIf Left(strDate, 2) = "昭和" Then
            intSaveDate = intSaveDate + 1925
        ElseIf Left(strDate, 2) = "平成" Then
            intSaveDate = intSaveDate + 1988
        ElseIf Left(strDate, 2) = "嗚呼" Then '新元号対応
            intSaveDate = intSaveDate + 2018
        End If
        
'        If intKeta = 6 Then
'            strDate = CStr(intSaveDate) & Mid(strDate, InStr(strDate, "年")) '新元号対応
'            fncToSeireki = CStr(intSaveDate) & Format(strDate, "mm")
'        Else
'            strDate = CStr(intSaveDate) & Mid(strDate, InStr(strDate, "年")) '新元号対応
'            fncToSeireki = CStr(intSaveDate) & Format(strDate, "mm") & Format(strDate, "dd")
'        End If
        If blnNengappiflg Then
            If intKeta = 8 Then
                strDate = CStr(intSaveDate) & Mid(strDate, InStr(strDate, "年")) '新元号対応
                fncToSeireki = CStr(intSaveDate) & "年" & Format(strDate, "m") & "月"
            ElseIf intKeta = 11 Then
                If blnMaezeroflg Then
                    strDate = CStr(intSaveDate) & Mid(strDate, InStr(strDate, "年")) '新元号対応
                    fncToSeireki = CStr(intSaveDate) & "年" & Format(strDate, "mm") & "月" & Format(strDate, "dd") & "日"
                Else
                    strDate = CStr(intSaveDate) & Mid(strDate, InStr(strDate, "年")) '新元号対応
                    fncToSeireki = CStr(intSaveDate) & "年" & Format(strDate, "m") & "月" & Format(strDate, "d") & "日"
                End If
            End If
        Else
            If intKeta = 6 Then
                strDate = CStr(intSaveDate) & Mid(strDate, InStr(strDate, "年")) '新元号対応
                fncToSeireki = CStr(intSaveDate) & Format(strDate, "mm")
            Else
                strDate = CStr(intSaveDate) & Mid(strDate, InStr(strDate, "年")) '新元号対応
                fncToSeireki = CStr(intSaveDate) & Format(strDate, "mm") & Format(strDate, "dd")
            End If
        End If
    End If
    
End Function

Public Sub subSaveDel()
'関数名：subSaveDel
'内容　：ワークシートに保存されている各画面の状態を削除
'引数　：
    
    Dim i As Integer            'ループカウント
    Dim objAll As Object     'ループ用オブジェクト
    
    Dim wsSave As Worksheet     '画面の状態が保存されているシート
    '共通項目
    Call subSetSheet(2, wsSave)             'シートオブジェクト(別紙　共通項目)
    wsSave.Rows("2:2").Delete
    
    '補償内容セット(一括)
    Call subSetSheet(3, wsSave)         'シートオブジェクト(別紙　補償内容セット（一括）)
    wsSave.Rows("2:2").Delete
    
    '明細入力
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(1, wsSave)       'シートオブジェクト(明細入力)
    Else
        Call subSetSheet(17, wsSave)      'シートオブジェクト(明細入力（ノンフリート）)
    End If
    
'    '初期状態(10行)より明細行が存在する場合は削除
'    If Val(wsSave.Range("F9")) > 10 Then
'        Call subMeisaiDel(Val(wsSave.Range("F9")) - 10, "2")
'    End If
    
    '選択状態・内容のクリアのみを行い行を残す
    
    Call subMeisaiUnProtect             'シートの保護の解除
    
    Application.EnableEvents = False    'イベント無効
    
    '共通項目
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        '保険期間
        wsSave.Range("B3") = "　保険期間　　　：　"
        '受付区分
        wsSave.Range("E3") = "　受付区分　　　：　"
        '被保険者
        wsSave.Range("G3") = "　被保険者　　　　　　：　"
        '保険種類
        wsSave.Range("B4") = "　保険種類　　　：　"
        'フリート区分
        wsSave.Range("E4") = "　フリート区分　：　"
        '全車両一括付保特約
        wsSave.Range("G4") = "　全車両一括付保特約　：　"
        '払込方法
        wsSave.Range("B5") = "　払込方法　　　：　"
        '優良割引
        wsSave.Range("E5") = "　優良割引　　　：　"
        '第一種デメ割増
        wsSave.Range("G5") = "　第一種デメ割増  　　：　"
        'ﾌﾘｰﾄ多数割引
        wsSave.Range("B6") = "　ﾌﾘｰﾄ多数割引　：　"
        'ﾌﾘｰﾄｺｰﾄﾞ
        wsSave.Range("E6") = "　ﾌﾘｰﾄｺｰﾄﾞ　　　：　"
        
        For i = 1 To Val(wsSave.OLEObjects("txtSouhuho").Object.Value)
            wsSave.Range("C" & 20 + i & ":" & "AX" & 20 + i).ClearContents  '明細行（最終はAX列）
            Windows(ThisWorkbook.Name).ScrollColumn = 1
            Windows(ThisWorkbook.Name).ScrollRow = 1
        Next i
    Else
        'ノンフリート（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        '保険期間
        wsSave.Range("B3") = "　保険期間　　　：　"
        '受付区分
        wsSave.Range("E3") = "　受付区分　　　：　"
        '被保険者
        wsSave.Range("G3") = "　被保険者　　　　　　：　"
        '保険種類
        wsSave.Range("B4") = "　保険種類　　　：　"
        'フリート区分
        wsSave.Range("E4") = "　フリート区分　：　"
        'ノンフリート多数割引
        wsSave.Range("G4") = "　ノンフリート多数割引：　"
        '払込方法
        wsSave.Range("B5") = "　払込方法　　　：　"
        '団体割増引
        wsSave.Range("E5") = "　団体割増引　　：　"
         '第一種デメ割増
        wsSave.Range("G5") = "　"
        'ﾌﾘｰﾄ多数割引
        wsSave.Range("B6") = "　"
        'ﾌﾘｰﾄｺｰﾄﾞ
        wsSave.Range("E6") = "　"
        
        For i = 1 To Val(wsSave.OLEObjects("txtSouhuho").Object.Value)
            wsSave.Range("C" & 20 + i & ":" & "BH" & 20 + i).ClearContents  '明細行（最終はBH列）
            Windows(ThisWorkbook.Name).ScrollColumn = 1
            Windows(ThisWorkbook.Name).ScrollRow = 1
        Next i
    End If
    
    Set objAll = Nothing
    
    '明細入力画面のエラー用リスト初期化
    wsSave.OLEObjects("txtErrMsg").Object.Value = ""
    '明細入力画面の明細追加テキストボックス初期化
    wsSave.OLEObjects("TxtMsaiAddCnt").Object.Value = ""
    
    Set wsSave = Nothing
    
    'チェックボックスのクリア
    Call subClearAll
    
    Application.EnableEvents = True     'イベント有効

    Call subMeisaiProtect               'シートの保護
    
'    '初期状態(10行)より明細行がすくない場合は追加
'    If wsSave.Range("F9") < 10 Then
'        Call subMeisaiAdd(10 - wsSave.Range("F9"), "2")
'    End If
    
End Sub


'「選択」ボタン押下（明細入力画面のその他料率）
Sub subClickOtherBtn()
    Dim intRow As Integer
    Dim intCol As Integer
    
    Dim wsMeisai As Worksheet           '明細入力ワークシート
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(1, wsMeisai)       'シートオブジェクト(明細入力)
    Else
        Call subSetSheet(17, wsMeisai)      'シートオブジェクト(明細入力（ノンフリート）)
    End If
    
    intRow = wsMeisai.Shapes(Application.Caller).TopLeftCell.Row
    intCol = wsMeisai.Shapes(Application.Caller).TopLeftCell.Column
    
    strTxtOther = intRow & ":" & intCol + 1
    
    Set wsMeisai = Nothing
    
    'シート・ブックの非表示
    Call subSheetVisible(False)
    
    frmOtherrate.Show vbModeless

End Sub

'
Sub subClickchkMeisai()
    Dim strChkValue As String
    
    Call subMeisaiUnProtect     'シートの保護の解除
    
    Dim wsMeisai As Worksheet           '明細入力ワークシート
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(1, wsMeisai)       'シートオブジェクト(明細入力)
    Else
        Call subSetSheet(17, wsMeisai)      'シートオブジェクト(明細入力（ノンフリート）)
    End If
    
    strChkValue = wsMeisai.Shapes(Application.Caller).TopLeftCell.Address
    If wsMeisai.CheckBoxes(Application.Caller).Value = xlOn Then
        wsMeisai.Range(strChkValue) = Left(wsMeisai.Range(strChkValue), InStr(wsMeisai.Range(strChkValue), "/") - 1) & "/True"
    Else
        wsMeisai.Range(strChkValue) = Left(wsMeisai.Range(strChkValue), InStr(wsMeisai.Range(strChkValue), "/") - 1) & "/False"
    End If
    
    Set wsMeisai = Nothing
    
    Call subMeisaiProtect       'シートの保護
    
End Sub


'管理者機能
Function fncAdmin() As Boolean

    Dim intMsg As Integer
    intMsg = MsgBox("編集を完了しファイルを保存しますか。", vbYesNo, "確認ダイアログ")
    
    If intMsg = vbYes Then
        
        Call subBookUnProtect           'ブックの保護を解除
        
        Application.ScreenUpdating = False                            '描画停止
        
        Application.OnKey "%{q}", ""
        ThisWorkbook.Worksheets("明細入力").Visible = True
        ThisWorkbook.Worksheets("別紙　コード値").Visible = False
        ThisWorkbook.Worksheets("別紙　各種設定").Visible = False
        ThisWorkbook.Worksheets("見積書").Visible = False
        ThisWorkbook.Worksheets("車両明細書").Visible = False
        ThisWorkbook.Worksheets("契約申込書1枚目").Visible = False
        ThisWorkbook.Worksheets("契約申込書2枚目").Visible = False
        ThisWorkbook.Worksheets("明細書").Visible = False
        ThisWorkbook.Worksheets("申込書ＥＤＰ").Visible = False
        ThisWorkbook.Worksheets("明細書ＥＤＰ").Visible = False
        ThisWorkbook.Worksheets("別紙　見積書設定").Visible = False
        ThisWorkbook.Worksheets("別紙　車両明細書設定").Visible = False
        ThisWorkbook.Worksheets("別紙　申込書(1枚目)設定").Visible = False
        ThisWorkbook.Worksheets("別紙　申込書(2枚目)設定").Visible = False
        ThisWorkbook.Worksheets("別紙　明細書設定").Visible = False
        ThisWorkbook.Worksheets("別紙　申込書ＥＤＰ設定").Visible = False
        ThisWorkbook.Worksheets("別紙　明細書ＥＤＰ設定").Visible = False
        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        ThisWorkbook.Worksheets("明細入力（ノンフリート）").Visible = False
        ThisWorkbook.Worksheets("明細書印刷").Visible = False
        ThisWorkbook.Worksheets("明細書印刷（ノンフリート）").Visible = False
        ThisWorkbook.Worksheets("別紙　コード値（ノンフリート）").Visible = False
        ThisWorkbook.Worksheets("見積書（ノンフリート）").Visible = False
        ThisWorkbook.Worksheets("車両明細書（ノンフリート）").Visible = False
        ThisWorkbook.Worksheets("明細書（ノンフリート）").Visible = False
        ThisWorkbook.Worksheets("申込書ＥＤＰ（ノンフリート）").Visible = False
        ThisWorkbook.Worksheets("明細書ＥＤＰ（ノンフリート）").Visible = False
        ThisWorkbook.Worksheets("別紙　見積書設定（ノンフリート）").Visible = False
        ThisWorkbook.Worksheets("別紙　車両明細書設定（ノンフリート）").Visible = False
        ThisWorkbook.Worksheets("別紙　申込書(1枚目)設定（ノンフリート）").Visible = False
        ThisWorkbook.Worksheets("別紙　申込書(2枚目)設定（ノンフリート）").Visible = False
        ThisWorkbook.Worksheets("別紙　明細書設定（ノンフリート）").Visible = False
        ThisWorkbook.Worksheets("別紙　申込書ＥＤＰ設定（ノンフリート）").Visible = False
        ThisWorkbook.Worksheets("別紙　明細書ＥＤＰ設定（ノンフリート）").Visible = False
        ThisWorkbook.Worksheets("テキスト内容(共通)").Visible = False
        ThisWorkbook.Worksheets("テキスト内容(明細)").Visible = False
        
        Application.ScreenUpdating = True                            '描画再開
        
        'ブックの保存
        blnSaveFlg = True
        ThisWorkbook.Save
        
        Call subMeisaiUnProtect     'シートの保護の解除
        Call subMeisaiPrtUnProtect  'シートの保護の解除 '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        Call subCmbInitialize       '計算用シートのコンボボックスを設定
        Call subMeisaiProtect       'シートの保護
        Call subMeisaiPrtProtect    'シートの保護 '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        
        Call subSheetVisible(False)     'シート・ブックの非表示
        Call subBookProtect             'ブックの保護
        
        frmTop.Show vbModeless
    End If

End Function


Public Sub subSetSheet(ByVal intSheetType As Integer, ByRef wsSheet As Worksheet)
    '関数名：subSetSheet
'内容　：シートオブジェクトに値を設定
'引数　：
'        intSheetType   =  1 :明細入力
'                          2 :別紙　共通項目
'                          3 :別紙　補償内容セット（一括）
'                          4 :別紙　コード値
'                          5 :別紙　各種設定
'                          6 :テキスト内容(共通)
'                          7 :テキスト内容(明細)
'                          8 :申込書印刷画面内容
'                          9 :別紙　見積書設定
'                         10 :別紙　車両明細書設定
'                         11 :別紙　申込書(1枚目)設定
'                         12 :別紙　申込書(2枚目)設定
'                         13 :別紙　明細書設定
'                         14 :別紙　申込書ＥＤＰ設定
'                         15 :別紙　明細書ＥＤＰ設定
'                         16 :別紙　コード値（ノンフリート）　　        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
'                         17 :明細入力（ノンフリート）　　　　　        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
'                         18 :明細書印刷　　　　　　　　　　　　        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
'                         19 :明細書印刷（ノンフリート）　　　　        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
'                         20 :別紙　見積書設定（ノンフリート）          '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
'                         21 :別紙　車両明細書設定（ノンフリート）      '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
'                         22 :別紙　申込書(1枚目)設定（ノンフリート）   '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
'                         23 :別紙　申込書(2枚目)設定（ノンフリート）   '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
'                         24 :別紙　明細書設定（ノンフリート）          '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
'                         25 :別紙　申込書ＥＤＰ設定（ノンフリート）    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
'                         26 :別紙　明細書ＥＤＰ設定（ノンフリート）    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
'
'                        101 :見積書WK
'                        102 :車両明細書WK
'                        103 :契約申込書1枚目WK
'                        104 :契約申込書2枚目WK
'                        105 :明細書WK
'                        106 :申込書ＥＤＰWK
'                        107 :明細書ＥＤＰWK
'        wsSheet        = "":シートオブジェクト（ノンフリート）
    Select Case intSheetType
        Case 1
            Set wsSheet = ThisWorkbook.Worksheets("明細入力")
        Case 2
            Set wsSheet = ThisWorkbook.Worksheets("別紙　共通項目")
        Case 3
            Set wsSheet = ThisWorkbook.Worksheets("別紙　補償内容セット（一括）")
        Case 4
            Set wsSheet = ThisWorkbook.Worksheets("別紙　コード値")
        Case 5
            Set wsSheet = ThisWorkbook.Worksheets("別紙　各種設定")
        Case 6
            Set wsSheet = ThisWorkbook.Worksheets("テキスト内容(共通)")
        Case 7
            Set wsSheet = ThisWorkbook.Worksheets("テキスト内容(明細)")
        Case 8
            Set wsSheet = ThisWorkbook.Worksheets("申込書印刷画面内容")
        Case 9
            Set wsSheet = ThisWorkbook.Worksheets("別紙　見積書設定")
        Case 10
            Set wsSheet = ThisWorkbook.Worksheets("別紙　車両明細書設定")
        Case 11
            Set wsSheet = ThisWorkbook.Worksheets("別紙　申込書(1枚目)設定")
        Case 12
            Set wsSheet = ThisWorkbook.Worksheets("別紙　申込書(2枚目)設定")
        Case 13
            Set wsSheet = ThisWorkbook.Worksheets("別紙　明細書設定")
        Case 14
            Set wsSheet = ThisWorkbook.Worksheets("別紙　申込書ＥＤＰ設定")
        Case 15
            Set wsSheet = ThisWorkbook.Worksheets("別紙　明細書ＥＤＰ設定")
        Case 16
            Set wsSheet = ThisWorkbook.Worksheets("別紙　コード値（ノンフリート）")         '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        Case 17
            Set wsSheet = ThisWorkbook.Worksheets("明細入力（ノンフリート）")               '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        Case 18
            Set wsSheet = ThisWorkbook.Worksheets("明細書印刷")                             '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        Case 19
            Set wsSheet = ThisWorkbook.Worksheets("明細書印刷（ノンフリート）")             '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        Case 20
            Set wsSheet = ThisWorkbook.Worksheets("別紙　見積書設定（ノンフリート）")       '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        Case 21
            Set wsSheet = ThisWorkbook.Worksheets("別紙　車両明細書設定（ノンフリート）")   '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        Case 22
            Set wsSheet = ThisWorkbook.Worksheets("別紙　申込書(1枚目)設定（ノンフリート）") '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        Case 23
            Set wsSheet = ThisWorkbook.Worksheets("別紙　申込書(2枚目)設定（ノンフリート）") '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        Case 24
            Set wsSheet = ThisWorkbook.Worksheets("別紙　明細書設定（ノンフリート）")       '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        Case 25
            Set wsSheet = ThisWorkbook.Worksheets("別紙　申込書ＥＤＰ設定（ノンフリート）") '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        Case 26
            Set wsSheet = ThisWorkbook.Worksheets("別紙　明細書ＥＤＰ設定（ノンフリート）") '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        Case 101
            Set wsSheet = ThisWorkbook.Worksheets("見積書WK")
        Case 102
            Set wsSheet = ThisWorkbook.Worksheets("車両明細書WK")
        Case 103
            Set wsSheet = ThisWorkbook.Worksheets("契約申込書1枚目WK")
        Case 104
            Set wsSheet = ThisWorkbook.Worksheets("契約申込書2枚目WK")
        Case 105
            Set wsSheet = ThisWorkbook.Worksheets("明細書WK")
        Case 106
            Set wsSheet = ThisWorkbook.Worksheets("申込書ＥＤＰWK")
        Case 107
            Set wsSheet = ThisWorkbook.Worksheets("明細書ＥＤＰWK")
    End Select
End Sub

Public Sub subSheetVisible(ByVal blnVisibleMode As Boolean)
    Dim blnOtherBookFlg As Boolean
    Dim sOpenBookSub    As Variant
    Dim wsMeisai        As Worksheet    '明細入力ワークシート
    Dim sOpenbookAll    As Workbook
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    Dim wsMeisaiNonfleet      As Worksheet    '明細入力（ノンフリート）ワークシート
    Dim wsMeisaiPrint         As Worksheet    '明細書印刷画面ワークシート
    Dim wsMeisaiPrintNonfleet As Worksheet    '明細書印刷画面（ノンフリート）ワークシート
    
    
    Call subSetSheet(1, wsMeisai)               'シートオブジェクト(明細入力)
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    Call subSetSheet(17, wsMeisaiNonfleet)      'シートオブジェクト(明細入力（ノンフリート）)
    Call subSetSheet(18, wsMeisaiPrint)         'シートオブジェクト(明細書印刷)
    Call subSetSheet(19, wsMeisaiPrintNonfleet) 'シートオブジェクト(明細書印刷（ノンフリート）)
    
    
    If blnVisibleMode Then
        '表示
        
        '他にブックが開いていないか確認
        For Each sOpenbookAll In Workbooks
            If sOpenbookAll.Name <> ThisWorkbook.Name Then
                blnOtherBookFlg = True
                Exit For
            End If
        Next sOpenbookAll
        
        If blnOtherBookFlg Then
            '他にブックが開かれている場合、シートを表示
            Windows(ThisWorkbook.Name).Visible = True
        Else
            '他にブックが開かれていない場合、エクセルごと表示
            Application.Visible = True
            Windows(ThisWorkbook.Name).Visible = True
        End If
        
        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        If blnMoushikomiflg Then
            '明細書印刷画面
            If FleetTypeFlg = 1 Then
                'フリート
                wsMeisaiPrint.Visible = True
                
                wsMeisai.Visible = False
                wsMeisaiNonfleet.Visible = False
                wsMeisaiPrintNonfleet.Visible = False
                
                wsMeisaiPrint.Activate
            Else
                'ノンフリート
                wsMeisaiPrintNonfleet.Visible = True
                
                wsMeisai.Visible = False
                wsMeisaiNonfleet.Visible = False
                wsMeisaiPrint.Visible = False
                
                wsMeisaiPrintNonfleet.Activate
            End If
        Else
            '明細入力画面
            If FleetTypeFlg = 1 Then
                'フリート
                wsMeisai.Visible = True
                
                wsMeisaiNonfleet.Visible = False
                wsMeisaiPrint.Visible = False
                wsMeisaiPrintNonfleet.Visible = False
                
                wsMeisai.Activate
            Else
                'ノンフリート
                wsMeisaiNonfleet.Visible = True
                
                wsMeisai.Visible = False
                wsMeisaiPrint.Visible = False
                wsMeisaiPrintNonfleet.Visible = False
                
                wsMeisaiNonfleet.Activate
            End If
        End If
        
    Else
        '非表示
        
        '他にブックが開いていないか確認
        For Each sOpenbookAll In Workbooks
            If sOpenbookAll.Name <> ThisWorkbook.Name Then
                blnOtherBookFlg = True
                Exit For
            End If
        Next sOpenbookAll
        
        If blnOtherBookFlg Then
            '他にブックが開かれている場合、シートを非表示
            Windows(ThisWorkbook.Name).Visible = False
        Else
            '他にブックが開かれていない場合、エクセルごと非表示
            Windows(ThisWorkbook.Name).Visible = False
            Application.Visible = False
        End If
    End If
    
    Set sOpenbookAll = Nothing
    Set sOpenBookSub = Nothing
    
End Sub

'ブックの保護
Public Sub subBookProtect()
    Dim wsSetting As Worksheet          '各種設定ワークシート
    Call subSetSheet(5, wsSetting)      'シートオブジェクト(別紙　各種設定)
    
    ThisWorkbook.Protect Password:=wsSetting.Range("B4").Value    'ブックの保護
    
    Set wsSetting = Nothing
    
End Sub

'ブックの保護を解除
Public Sub subBookUnProtect()
    Dim wsSetting As Worksheet          '各種設定ワークシート
    Call subSetSheet(5, wsSetting)      'シートオブジェクト(別紙　各種設定)
    
    ThisWorkbook.Unprotect Password:=wsSetting.Range("B4").Value    'ブックの保護を解除
    
    Set wsSetting = Nothing
    
End Sub

'シートの保護
Public Sub subMeisaiProtect()
    Dim wsMeisaiFleet  As Worksheet           '明細入力ワークシート
    Dim wsMeisaiNonfleet  As Worksheet        '明細入力ワークシート
    Dim wsSetting As Worksheet                '各種設定ワークシート
    
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 0 Then
        Call subSetSheet(1, wsMeisaiFleet)        'シートオブジェクト(明細入力)
        Call subSetSheet(17, wsMeisaiNonfleet)    'シートオブジェクト(明細入力（ノンフリート）)
    ElseIf FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(1, wsMeisaiFleet)        'シートオブジェクト(明細入力)
    Else
        Call subSetSheet(17, wsMeisaiNonfleet)    'シートオブジェクト(明細入力（ノンフリート）)
    End If
    Call subSetSheet(5, wsSetting)                'シートオブジェクト(別紙　各種設定)
    
    'シートの保護を解除
    If Not wsMeisaiFleet Is Nothing Then
        wsMeisaiFleet.Protect Password:=wsSetting.Range("B4").Value
    End If
    If Not wsMeisaiNonfleet Is Nothing Then
        wsMeisaiNonfleet.Protect Password:=wsSetting.Range("B4").Value
    End If
    
    Set wsMeisaiFleet = Nothing
    Set wsMeisaiNonfleet = Nothing
    Set wsSetting = Nothing
    
End Sub

'シートの保護を解除
Public Sub subMeisaiUnProtect()
    Dim wsMeisaiFleet  As Worksheet           '明細入力ワークシート
    Dim wsMeisaiNonfleet  As Worksheet           '明細入力ワークシート
    Dim wsSetting As Worksheet          '各種設定ワークシート
    
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 0 Then
        Call subSetSheet(1, wsMeisaiFleet)       'シートオブジェクト(明細入力)
        Call subSetSheet(17, wsMeisaiNonfleet)      'シートオブジェクト(明細入力（ノンフリート）)
    ElseIf FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(1, wsMeisaiFleet)       'シートオブジェクト(明細入力)
    Else
        Call subSetSheet(17, wsMeisaiNonfleet)      'シートオブジェクト(明細入力（ノンフリート）)
    End If
    Call subSetSheet(5, wsSetting)      'シートオブジェクト(別紙　各種設定)
    
    'シートの保護を解除
    If Not wsMeisaiFleet Is Nothing Then
        wsMeisaiFleet.Unprotect Password:=wsSetting.Range("B4").Value
    End If
    If Not wsMeisaiNonfleet Is Nothing Then
        wsMeisaiNonfleet.Unprotect Password:=wsSetting.Range("B4").Value
    End If
    
    Set wsMeisaiFleet = Nothing
    Set wsMeisaiNonfleet = Nothing
    Set wsSetting = Nothing
    
End Sub


'2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
'明細書印刷シートの保護
Public Sub subMeisaiPrtProtect()
    Dim wsMeisaiFleet As Worksheet      '明細書印刷ワークシート（フリート）
    Dim wsMeisaiNonfleet  As Worksheet  '明細書印刷ワークシート（ノンフリート）
    Dim wsSetting As Worksheet          '各種設定ワークシート
    
    If FleetTypeFlg = 0 Then
        'トップ画面等
        Call subSetSheet(18, wsMeisaiFleet)    'シートオブジェクト(明細書印刷)
        Call subSetSheet(19, wsMeisaiNonfleet) 'シートオブジェクト(明細書印刷（ノンフリート）)
    ElseIf FleetTypeFlg = 1 Then
        'フリート
        Call subSetSheet(18, wsMeisaiFleet)    'シートオブジェクト(明細書印刷)
    Else
        'ノンフリート
        Call subSetSheet(19, wsMeisaiNonfleet) 'シートオブジェクト(明細書印刷（ノンフリート）)
    End If
    Call subSetSheet(5, wsSetting)             'シートオブジェクト(別紙　各種設定)
    
    'シートの保護を解除
    If Not wsMeisaiFleet Is Nothing Then
        wsMeisaiFleet.Protect Password:=wsSetting.Range("B4").Value
    End If
    If Not wsMeisaiNonfleet Is Nothing Then
        wsMeisaiNonfleet.Protect Password:=wsSetting.Range("B4").Value
    End If
    
    Set wsMeisaiFleet = Nothing
    Set wsMeisaiNonfleet = Nothing
    Set wsSetting = Nothing
    
End Sub


'2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
'明細書印刷シートの保護を解除
Public Sub subMeisaiPrtUnProtect()
    Dim wsMeisaiFleet  As Worksheet      '明細書印刷ワークシート（フリート）
    Dim wsMeisaiNonfleet  As Worksheet   '明細書印刷ワークシート（ノンフリート）
    Dim wsSetting As Worksheet           '各種設定ワークシート
    
    If FleetTypeFlg = 0 Then
        'トップ画面等
        Call subSetSheet(18, wsMeisaiFleet)    'シートオブジェクト(明細書印刷)
        Call subSetSheet(19, wsMeisaiNonfleet) 'シートオブジェクト(明細書印刷（ノンフリート）)
    ElseIf FleetTypeFlg = 1 Then
        'フリート
        Call subSetSheet(18, wsMeisaiFleet)    'シートオブジェクト(明細書印刷)
    Else
        'ノンフリート
        Call subSetSheet(19, wsMeisaiNonfleet) 'シートオブジェクト(明細書印刷（ノンフリート）)
    End If
    Call subSetSheet(5, wsSetting)             'シートオブジェクト(別紙　各種設定)
    
    
    'シートの保護を解除
    If Not wsMeisaiFleet Is Nothing Then
        wsMeisaiFleet.Unprotect Password:=wsSetting.Range("B4").Value
    End If
    If Not wsMeisaiNonfleet Is Nothing Then
        wsMeisaiNonfleet.Unprotect Password:=wsSetting.Range("B4").Value
    End If
    
    Set wsMeisaiFleet = Nothing
    Set wsMeisaiNonfleet = Nothing
    Set wsSetting = Nothing
    
End Sub


Sub subShortCutKey(ByVal intFlg As Integer)
'intFlg = 1 'ショートカットキー無効
'intFlg = 2 'ショートカットキー有効

    If intFlg = 1 Then
    '無効
        Application.OnKey "^6", ""              'オブジェクトの表示・非表示
        Application.OnKey "^+@", ""             'ワークシートのセルの値の表示と、数式の表示を切り替え
        Application.OnKey "%{F8}", ""           '[マクロ]ダイアログボックスの表示
        Application.OnKey "%{F11}", ""          'VBAエディターの起動
    Else
    '有効
        Application.OnKey "^6"                  'オブジェクトの表示・非表示
        Application.OnKey "^+@"                 'ワークシートのセルの値の表示と、数式の表示を切り替え
        Application.OnKey "%{F8}"               '[マクロ]ダイアログボックスの表示
        Application.OnKey "%{F11}"              'VBAエディターの起動
    End If
    
End Sub

Function fncFormatDigit(ByVal str As String, _
                     ByVal strChar As String, _
                     ByVal lngdigit As Long) As String
'機能：指定文字埋め関数
'引数：str　：変換前の文字列
'　　　chr  ：埋める文字(１文字目のみ使用)
'　　　digit：桁数
'戻値：指定文字埋め後の文字列
    
    Dim strtmp As String
    strtmp = str
    If Len(str) < lngdigit And Len(strChar) > 0 Then
      strtmp = Right(String(lngdigit, strChar) & str, lngdigit)
    End If
    fncFormatDigit = strtmp

End Function

Private Sub subOtherClose()
    Dim USF As UserForm
    
    For Each USF In UserForms
        If TypeOf USF Is frmTop Then
            If frmTop.Visible Then
                Call subBookUnProtect           'ブックの保護を解除 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                Call subSheetVisible(True)      'シート・ブックの表示
                Call subBookProtect             'ブックの保護 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                Call subSheetVisible(False)     'シート・ブックの非表示
            End If
        End If
        If TypeOf USF Is frmKyoutsuu Then
            If frmKyoutsuu.Visible Then
                Call subBookUnProtect           'ブックの保護を解除 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                Call subSheetVisible(True)      'シート・ブックの表示
                Call subBookProtect             'ブックの保護 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                Call subSheetVisible(False)     'シート・ブックの非表示
            End If
        End If
        If TypeOf USF Is frmOtherrate Then
            If frmOtherrate.Visible Then
                Call subBookUnProtect           'ブックの保護を解除 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                Call subSheetVisible(True)      'シート・ブックの表示
                Call subBookProtect             'ブックの保護 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                Call subSheetVisible(False)     'シート・ブックの非表示
            End If
        End If
        If TypeOf USF Is frmSyaryou Then
            If frmSyaryou.Visible Then
                Call subBookUnProtect           'ブックの保護を解除 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                Call subSheetVisible(True)      'シート・ブックの表示
                Call subBookProtect             'ブックの保護 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                Call subSheetVisible(False)     'シート・ブックの非表示
            End If
        End If
        If TypeOf USF Is frmHosyoSet Then
            If frmHosyoSet.Visible Then
                Call subBookUnProtect           'ブックの保護を解除 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                Call subSheetVisible(True)      'シート・ブックの表示
                Call subBookProtect             'ブックの保護 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                Call subSheetVisible(False)     'シート・ブックの非表示
            End If
        End If
        If TypeOf USF Is frmPrintMenu Then
            If frmPrintMenu.Visible Then
                Call subBookUnProtect           'ブックの保護を解除 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                Call subSheetVisible(True)      'シート・ブックの表示
                Call subBookProtect             'ブックの保護 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                Call subSheetVisible(False)     'シート・ブックの非表示
            End If
        End If
        If TypeOf USF Is frmEntryMitsumori Then
            If frmEntryMitsumori.Visible Then
                Call subBookUnProtect           'ブックの保護を解除 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                Call subSheetVisible(True)      'シート・ブックの表示
                Call subBookProtect             'ブックの保護 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                Call subSheetVisible(False)     'シート・ブックの非表示
            End If
        End If
        If TypeOf USF Is frmEntryMoushikomi Then
            If frmEntryMoushikomi.Visible Then
                Call subBookUnProtect           'ブックの保護を解除 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                Call subSheetVisible(True)      'シート・ブックの表示
                Call subBookProtect             'ブックの保護 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                Call subSheetVisible(False)     'シート・ブックの非表示
            End If
        End If
    Next
    
End Sub

Public Sub subAppClose()
    Dim wsSetting As Worksheet          '各種設定ワークシート
    Call subSetSheet(5, wsSetting)      'シートオブジェクト(別紙　各種設定)
    wsSetting.Range("P1") = "True"

    Set wsSetting = Nothing
    
    'ショートカットキーの有効
    Call subShortCutKey(2)

    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    Application.DisplayStatusBar = True
    Application.DisplayFormulaBar = True
    Application.CommandBars("Ply").Enabled = True

    If Workbooks.Count = 1 Then
        blnCloseFlg = True
    End If

    'Auto_Closeにてブックを閉じる
    ThisWorkbook.RunAutoMacros Which:=xlAutoClose
        
End Sub

'オートクローズイベント
Private Sub Auto_Close()
    Dim wsSetting As Worksheet
    Call subSetSheet(5, wsSetting)      'シートオブジェクト(別紙　各種設定)
    
    If CBool(wsSetting.Range("P1")) Then
        Application.DisplayAlerts = False
        Application.Cursor = xlWait
        Application.OnTime Now, "my_Procedure"
    End If
End Sub
Private Sub my_Procedure()
    Application.Cursor = xlDefault
    If Workbooks.Count = 1 Then
        Application.Quit
    Else
        ThisWorkbook.Close SaveChanges:=False
    End If
End Sub

'明細入力シートに反映
Public Sub fncMeisaiEntry(ByVal varContent As Variant, ByVal intCar As Integer)
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim intActRow       As Integer
    Dim intSouhuho      As Integer
    Dim varMeisaiRow    As Variant
    Dim varMeisaiCell   As Variant
    Dim varMeisaiRgCell As Variant
    
    intActRow = 0
    
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then
        'フリート （明細項目数47）
        ReDim varMeisaiRgCell(intCar - 1, 47)
    Else
        'ノンフリート（明細項目数59）
        ReDim varMeisaiRgCell(intCar - 1, 59)
    End If
    
    Dim wstMeisai As Worksheet
    Dim wstTextM As Worksheet
    
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(1, wstMeisai)       'シートオブジェクト(明細入力)
        Call subSetSheet(7, wstTextM)          'シートオブジェクト(テキスト内容)
    Else
        Call subSetSheet(17, wstMeisai)       'シートオブジェクト(明細入力（ノンフリート）)
        Call subSetSheet(7, wstTextM)          'シートオブジェクト(テキスト内容)
    End If
    
    Dim objSouhuho As Object
    Set objSouhuho = wstMeisai.OLEObjects("txtSouhuho").Object
    intSouhuho = Val(Left(objSouhuho.Value, Len(objSouhuho.Value) - 2))

    Call subBookUnProtect 'ブックの保護を解除 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    Call subSheetVisible(True)      'シート・ブックの表示
    Call subBookProtect   'ブックの保護 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加

    ''明細行の追加・削除
    '明細行数 = 明細行数(デフォルト:10行)
    If intCar = intSouhuho Then
    '明細行数 < 明細行数-余剰分の明細行を削除
    ElseIf intCar < intSouhuho Then
        Call subMeisaiDel(intSouhuho - intCar, "2")
    '明細行数 > 明細行数-不足分の明細行を追加
    ElseIf intCar > intSouhuho Then
        Call subMeisaiAdd(intCar - intSouhuho, "2")
    End If
    Call subSheetVisible(False)      'シート・ブックの非表示
    
    ''明細行のシート部分に反映
    For i = 0 To UBound(varContent, 1)
        For j = 1 To UBound(varContent, 2)
            If UBound(Split(varContent(i, j))) = -1 Then Exit For
            
            varMeisaiRow = Split(varContent(i, j), ",")
            
            '明細シート用配列作成
            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
            If FleetTypeFlg = 1 Then  'フリート
                varMeisaiCell = fncMeisaiSetCell(varMeisaiRow)
            Else
                varMeisaiCell = fncNonFleetMeisaiSetCell(varMeisaiRow)
            End If
            
            '2次元配列の1次元目に1次元配列をコピー
            Call subMeisaiArray(varMeisaiRgCell, varMeisaiCell, intActRow)
            
'            If blnChouhyouflg Then
                For k = 1 To UBound(varMeisaiRow) + 1
                    wstTextM.Cells(intActRow + 1, k) = varMeisaiRow(k - 1)
                    Call fncTextEdit(2, k, varMeisaiRow(k - 1), intActRow + 1)
                Next k
'            End If
            
            intActRow = intActRow + 1
            
        Next j
    Next i
    
    '明細入力(Cell)に貼り付け
    Call subMeisaiCell(varMeisaiRgCell)
    
    Set wstMeisai = Nothing
    Set wstTextM = Nothing
    Set objSouhuho = Nothing
    Set varMeisaiRow = Nothing
    Set varMeisaiCell = Nothing
    Set varMeisaiRgCell = Nothing
        
End Sub

'明細シート用配列作成（セル）
Function fncMeisaiSetCell(ByVal varContent As Variant) As Variant
    Dim i           As Integer
    Dim strText     As String
    Dim strSaveDate As String
    Dim varCell(47) As Variant           'Cell用配列
    
    strText = ""
    strSaveDate = ""
    
    'シートのセルに格納
    '--------C列～L列--------
    varCell(0) = fncFindCode(varContent(1), "AA")             '用途車種名
    varCell(1) = varContent(1)                                '用途車種コード
    varCell(2) = varContent(2)                                '車名
    varCell(3) = varContent(68)                               '登録番号
    varCell(4) = varContent(69)                               '車台番号
    varCell(5) = varContent(3)                                '型式
    varCell(6) = varContent(4)                                '仕様
    '初度登録年月
    strSaveDate = fncToWareki(varContent(5) & "25", 8)
    If strSaveDate = varContent(5) & "25" Then
        varCell(7) = varContent(5)
    Else
        varCell(7) = strSaveDate
    End If
    '車検満了日
    strSaveDate = fncToWareki(CStr(varContent(70)), 11)
    varCell(8) = strSaveDate
    '改造・不明車
    varCell(9) = fncFindCode(varContent(6), "AE")
    
    '--------M列～U列--------
    varCell(10) = varContent(7)                                '排気量
    varCell(11) = varContent(8)                                '2.5リットル越ディーゼル自小乗
    varCell(12) = IIf(varContent(72) = "1", "適用する", varContent(72)) 'ASV割引
    If blnChouhyouflg = False Then
        varCell(13) = ""                                       '料率-車両
        varCell(14) = ""                                       '料率-対人
        varCell(15) = ""                                       '料率-対物
        varCell(16) = ""                                       '料率-傷害
        varCell(17) = ""                                       '新車割引
        varCell(18) = ""                                       '車両上限
        varCell(19) = ""                                       '車両下限
    Else
        varCell(13) = varContent(23)                           '料率-車両
        varCell(14) = varContent(24)                           '料率-対人
        varCell(15) = varContent(25)                           '料率-対物
        varCell(16) = varContent(26)                           '料率-傷害
        varCell(17) = varContent(27)                           '新車割引
        varCell(18) = varContent(30)                           '車両上限
        varCell(19) = varContent(29)                           '車両下限
    End If
    varCell(20) = ""                                           'その他料率-ボタン
    
    '--------W列～AF列--------
    strText = IIf(varContent(16) = "3 ", "沖縄／", "")
    strText = strText + IIf(varContent(17) = "1", "レンタカー／", "")
    strText = strText + IIf(varContent(18) = "5 ", "教習車／", "")
    strText = strText + IIf(varContent(19) = "1 ", "ブーム対象外／", "")
    strText = strText + IIf(varContent(20) = "80", "リースカーオープンポリシー／", "")
    strText = strText + IIf(varContent(21) = "93", "オープンポリシー多数割引／", "")
    strText = strText + IIf(varContent(22) = "1 ", "公有／", _
                        IIf(varContent(22) = "2 ", "準公有／", ""))
    If blnChouhyouflg = False Then
        strText = strText + ""
    Else
        strText = strText + IIf(varContent(28) = "8", "特種区分／", "")
    End If
    If strText <> "" Then
        varCell(21) = Left(strText, Len(strText) - 1)          'その他料率-テキスト
    End If
    varCell(22) = fncFindCode(varContent(38), "BK")            '車両保険の種類
    varCell(23) = varContent(39)                               '車両保険金額
    varCell(24) = fncFindCode(varContent(40), "BO")            '車両免責金額
    varCell(25) = IIf(varContent(42) = "2", _
                        "適用する", varContent(42))            '車両全損臨費特約
    varCell(26) = IIf(varContent(44) = "1", _
                        "適用する", varContent(44))            '車両超過修理費用特約
    varCell(27) = IIf(varContent(43) = "1", _
                        "対象外", varContent(43))              '車両盗難対象外特約
    If varContent(45) = "1" Then
        varCell(28) = "無制限"                                 '対人無制限
    ElseIf varContent(46) = "1" Then
        varCell(28) = "対象外"                                 '対人対象外
    Else
        varCell(28) = fncFindCode(varContent(47), "CE")        '対人賠償保険金額
        If varCell(28) = "無制限" Or varCell(28) = "対象外" Then
            varCell(28) = varContent(47)
        End If
    End If
    varCell(29) = IIf(varContent(48) = "1", _
                        "対象外", varContent(48))              '自損事故傷害特約
    
    '--------AG列～AP列--------
    varCell(30) = IIf(varContent(49) = "1", _
                        "対象外", varContent(49))              '無保険車事故傷害特約
    If varContent(50) = "1" Then
        varCell(31) = "無制限"                                 '対物無制限
    ElseIf varContent(51) = "1" Then
        varCell(31) = "対象外"                                 '対物対象外
    Else
        varCell(31) = fncFindCode(varContent(52), "CI")        '対物賠償保険金額
        If varCell(31) = "無制限" Or varCell(31) = "対象外" Then
            varCell(31) = varContent(52)
        End If
    End If
    varCell(32) = fncFindCode(varContent(53), "BS")            '対物免責金額
    varCell(33) = IIf(varContent(54) = "1", _
                        "適用する", varContent(54))            '対物超過修理費用特約
    If varContent(56) = 1 Then
        varCell(34) = "対象外"                                 '人身傷害対象外
    Else
        varCell(34) = fncFindCode(varContent(55), "CM")        '人身傷害(1名)
        If varCell(34) = "対象外" Then
            varCell(34) = varContent(55)
        End If
    End If
    varCell(35) = varContent(57)                               '人身傷害(1事故)
    If varContent(60) = "1" Then
        varCell(36) = "対象外"                                 '搭乗者傷害対象外
    Else
        varCell(36) = varContent(59)                           '搭乗者傷害(1名)
'        varCell(36) = fncFindCode(varContent(59), "CQ")        '搭乗者傷害(1名)
'        If varCell(36) = "対象外" Then
'            varCell(36) = varContent(59)
'        End If
    End If
    varCell(37) = varContent(61)                               '搭乗者傷害(1事故)
    varCell(38) = IIf(varContent(62) = "2", _
                        "適用する", varContent(62))            '日数払特約
    varCell(39) = IIf(varContent(63) = "1", _
                        "適用する", varContent(63))            '事業主費用特約
    
    '--------AQ列～AV列--------
    varCell(40) = IIf(varContent(64) = "1", _
                        "適用する", varContent(64))            '弁護士費用特約
    varCell(41) = IIf(varContent(37) = "1", _
                        "限定", varContent(37))                '従業員等限定特約
    varCell(42) = varContent(41)                               '事故代車・身の回り品特約
    varCell(43) = IIf(varContent(73) = "1", "不適用", varContent(73))   '車両搬送時不適用特約
    If blnChouhyouflg = False Then
        varCell(44) = ""                                           '合計保険料
        varCell(45) = ""                                           '初回保険料
        varCell(46) = ""                                           '年間保険料
        varCell(47) = ""                                           '稟議警告フラグ
    Else
        varCell(44) = varContent(31)                               '合計保険料
        varCell(45) = varContent(32)                               '初回保険料
        varCell(46) = varContent(33)                               '年間保険料
        varCell(47) = IIf(varContent(67) = "1", "稟議エラー有", IIf(varContent(67) = "2", "警告有", ""))  '稟議警告フラグ
    End If
    
    fncMeisaiSetCell = varCell
    
End Function
'ノンフリート明細シート用配列作成（セル）　 （2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
Function fncNonFleetMeisaiSetCell(ByVal varContent As Variant) As Variant
    Dim i           As Integer
    Dim strText     As String
    Dim strSaveDate As String
    Dim varCell(57) As Variant          'Cell用配列（ノンフリート明細付の列）
    
    strText = ""
    strSaveDate = ""
    
    'シートのセルに格納
    '--------C列～L列--------
    varCell(0) = fncFindCode(varContent(1), "AA")             '用途車種名
    varCell(1) = varContent(1)                                '用途車種コード
    varCell(2) = varContent(2)                                '車名
    varCell(3) = varContent(68)                               '登録番号
    varCell(4) = varContent(69)                               '車台番号
    varCell(5) = varContent(3)                                '型式
    varCell(6) = varContent(4)                                '仕様
    '初度登録年月
    strSaveDate = fncToWareki(varContent(5) & "25", 8)
    If strSaveDate = varContent(5) & "25" Then
        varCell(7) = varContent(5)
    Else
        varCell(7) = strSaveDate
    End If
    '車検満了日
    strSaveDate = fncToWareki(CStr(varContent(70)), 11)
    varCell(8) = strSaveDate
    '改造・不明車
    varCell(9) = fncFindCode(varContent(6), "AE")
    
    '--------M列～U列--------
    '排気量
    varCell(10) = varContent(7)
    
    '2.5リットル越ディーゼル自小乗
    varCell(11) = varContent(8)
    
    '被保険者生年月日
    varCell(12) = fncToWareki(CStr(varContent(9)), 11)
    
    'ノンフリート等級
    varCell(13) = fncFindCode(varContent(10), "AI")
    
    '事故有係数適用期間
    varCell(14) = fncFindCode(varContent(11), "AM")
    
    'ゴールド免許割引
    varCell(15) = IIf(varContent(14) = "1", "適用する", varContent(14))
    
    '使用目的
    varCell(16) = fncFindCode(varContent(15), "DD")
    
    '年齢条件
    varCell(17) = fncFindCode(varContent(34), "BC")
    
    '高齢運転者対象外
    varCell(18) = IIf(varContent(35) = "1", "対象外", varContent(35))
   
    '運転者限定
    varCell(19) = fncFindCode(varContent(36), "BG")
    
    varCell(20) = IIf(varContent(72) = "1", "適用する", varContent(72))
    
    '料率クラス
    If blnChouhyouflg = False Then
        varCell(21) = ""                                       '料率-車両
        varCell(22) = ""                                       '料率-対人
        varCell(23) = ""                                       '料率-対物
        varCell(24) = ""                                       '料率-傷害
        varCell(25) = ""                                       '新車割引
        varCell(26) = ""                                       '車両上限
        varCell(27) = ""                                       '車両下限
    Else
        varCell(21) = varContent(23)                           '料率-車両
        varCell(22) = varContent(24)                           '料率-対人
        varCell(23) = varContent(25)                           '料率-対物
        varCell(24) = varContent(26)                           '料率-傷害
        varCell(25) = varContent(27)                           '新車割引
        varCell(26) = varContent(30)                           '車両上限
        varCell(27) = varContent(29)                           '車両下限
    End If
    varCell(28) = ""                                           'その他料率-ボタン
    
    '--------W列～AF列--------
    strText = IIf(varContent(16) = "3 ", "沖縄／", "")
    strText = strText + IIf(varContent(17) = "1", "レンタカー／", "")
    strText = strText + IIf(varContent(18) = "5 ", "教習車／", "")
    strText = strText + IIf(varContent(19) = "1 ", "ブーム対象外／", "")
    strText = strText + IIf(varContent(20) = "80", "リースカーオープンポリシー／", "")
    strText = strText + IIf(varContent(21) = "93", "オープンポリシー多数割引／", "")
    strText = strText + IIf(varContent(22) = "1 ", "公有／", _
                        IIf(varContent(22) = "2 ", "準公有／", ""))
    If blnChouhyouflg = False Then
        strText = strText + ""
    Else
        strText = strText + IIf(varContent(28) = "8", "特種区分／", "")
    End If
    If strText <> "" Then
        varCell(29) = Left(strText, Len(strText) - 1)          'その他料率-テキスト
    End If
    
    varCell(30) = fncFindCode(varContent(38), "BK")            '車両保険の種類
    varCell(31) = varContent(39)                               '車両保険金額
    varCell(32) = fncFindCode(varContent(40), "BO")            '車両免責金額
    varCell(33) = IIf(varContent(42) = "2", _
                        "適用する", varContent(42))            '車両全損臨費特約
    varCell(34) = IIf(varContent(44) = "1", _
                        "適用する", varContent(44))            '車両超過修理費用特約
    varCell(35) = IIf(varContent(43) = "1", _
                        "対象外", varContent(43))              '車両盗難対象外特約
    If varContent(45) = "1" Then
        varCell(36) = "無制限"                                 '対人無制限
    ElseIf varContent(46) = "1" Then
        varCell(36) = "対象外"                                 '対人対象外
    Else
        varCell(36) = fncFindCode(varContent(47), "CE")        '対人賠償保険金額
        If varCell(36) = "無制限" Or varCell(36) = "対象外" Then
            varCell(36) = varContent(47)
        End If
    End If
    varCell(37) = IIf(varContent(48) = "1", _
                        "対象外", varContent(48))              '自損事故傷害特約
    
    '--------AG列～AP列--------
    varCell(38) = IIf(varContent(49) = "1", _
                        "対象外", varContent(49))              '無保険車事故傷害特約
    If varContent(50) = "1" Then
        varCell(39) = "無制限"                                 '対物無制限
    ElseIf varContent(51) = "1" Then
        varCell(39) = "対象外"                                 '対物対象外
    Else
        varCell(39) = fncFindCode(varContent(52), "CI")        '対物賠償保険金額
        If varCell(39) = "無制限" Or varCell(39) = "対象外" Then
            varCell(39) = varContent(52)
        End If
    End If
    varCell(40) = fncFindCode(varContent(53), "BS")            '対物免責金額
    varCell(41) = IIf(varContent(54) = "1", _
                        "適用する", varContent(54))            '対物超過修理費用特約
    If varContent(56) = 1 Then
        varCell(42) = "対象外"                                 '人身傷害対象外
    Else
        varCell(42) = fncFindCode(varContent(55), "CM")        '人身傷害(1名)
        If varCell(42) = "対象外" Then
            varCell(42) = varContent(55)
        End If
    End If
    varCell(43) = varContent(57)                               '人身傷害(1事故)
    If varContent(60) = "1" Then
        varCell(44) = "対象外"                                 '搭乗者傷害対象外
    Else
'        varCell(44) = fncFindCode(varContent(59), "CQ")        '搭乗者傷害(1名)
'        If varCell(44) = "対象外" Then
'            varCell(44) = varContent(59)
'        End If
    varCell(44) = varContent(59)
    End If
    varCell(45) = varContent(61)                               '搭乗者傷害(1事故)
    varCell(46) = IIf(varContent(62) = "2", _
                        "適用する", varContent(62))            '日数払特約
    varCell(47) = IIf(varContent(63) = "1", _
                        "適用する", varContent(63))            '事業主費用特約
    
    '--------AQ列～AV列--------
    varCell(48) = IIf(varContent(64) = "1", _
                        "適用する", varContent(64))            '弁護士費用特約
    
    '従業員等限定特約（ノンフリート不要）

    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    varCell(49) = fncFindCode(varContent(65), "BW")            'ファミリーバイク特約

    varCell(50) = IIf(varContent(66) = "1", "適用する", varContent(66))        '個人賠償責任補償特約

    varCell(51) = IIf(varContent(58) = "2", "適用する", varContent(58))        '自動車事故特約

    varCell(52) = varContent(41)                               '代車等セット特約
    
    varCell(53) = IIf(varContent(73) = "1", "不適用", varContent(73))  '車両搬送時諸費用特約の不適用に関する特約
   
    If blnChouhyouflg = False Then
        varCell(54) = ""                                           '合計保険料
        varCell(55) = ""                                           '初回保険料
        varCell(56) = ""                                           '年間保険料
        varCell(57) = ""                                           '稟議警告フラグ
    Else
        varCell(54) = varContent(31)                               '合計保険料
        varCell(55) = varContent(32)                               '初回保険料
        varCell(56) = varContent(33)                               '年間保険料
        varCell(57) = IIf(varContent(67) = "1", "稟議エラー有", IIf(varContent(67) = "2", "警告有", ""))   '稟議警告フラグ
    End If
    
    fncNonFleetMeisaiSetCell = varCell
    
End Function

'2次元配列の1次元目に1次元配列をコピー
Public Sub subMeisaiArray(ByRef varAllMeisai As Variant, ByVal varMeisai As Variant, ByVal intActRow As Integer)

    Dim i As Integer
    
    For i = 0 To UBound(varMeisai)
        varAllMeisai(intActRow, i) = varMeisai(i)
    Next i

End Sub


'明細入力(Cell)に貼り付け
Public Sub subMeisaiCell(ByVal varAllMeisai As Variant)
    Dim i As Integer
    Dim j As Integer
    Dim intStartRow As Integer
    Dim intStartCol As Integer
    
    Dim wstMeisai As Worksheet
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(1, wstMeisai)      'シートオブジェクト(明細入力)
    Else
        Call subSetSheet(17, wstMeisai)      'シートオブジェクト(明細入力)
    End If
    
    '明細開始行
    intStartRow = 21
    intStartCol = 3
    
    Application.EnableEvents = False
    For i = 0 To UBound(varAllMeisai, 1)
        For j = 0 To UBound(varAllMeisai, 2)
            wstMeisai.Cells(intStartRow + i, intStartCol + j) = varAllMeisai(i, j)
        Next
        intStartCol = 3
    Next i
    
    Set wstMeisai = Nothing
    
    Application.EnableEvents = True
End Sub


'2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
'PDFが開かれているか確認
Public Function fncIsFileOpen(ByVal strArgFile As String) As Boolean

    On Error GoTo FILE_ERR
    
    Open strArgFile For Binary Access Read Lock Read As #1
    Close #1
    fncIsFileOpen = False
    
    Exit Function
    
FILE_ERR:

    fncIsFileOpen = True
    
End Function

