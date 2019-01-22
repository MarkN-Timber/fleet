VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrintMenu 
   Caption         =   "帳票選択"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8145
   OleObjectBlob   =   "frmPrintMenu.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmPrintMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim intConfirmMsg As Integer

'初期表示
'2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
Public Sub UserForm_Initialize()
    'フォームのタイトル設定
    If FleetTypeFlg = 1 Then
        frmPrintMenu.Caption = frmPrintMenu.Caption & "（フリート契約）"
    Else
        frmPrintMenu.Caption = frmPrintMenu.Caption & "（ノンフリート明細付契約）"
    End If
End Sub


'「戻る」ボタン押下
Private Sub BtnBack_Click()
    Dim wstTextK As Worksheet
    Dim wstTextM As Worksheet
    Dim wstMitsuSave As Worksheet

    On Error GoTo Error

    intConfirmMsg = MsgBox("入力内容を削除してTOP画面に遷移します。" & vbCrLf & "よろしいですか？", vbYesNo, "確認ダイアログ")
    If intConfirmMsg = 6 Then
        Call subSetSheet(6, wstTextK)          'シートオブジェクト(テキスト内容(共通))
        Call subSetSheet(7, wstTextM)          'シートオブジェクト(テキスト内容(明細))
        Call subSetSheet(8, wstMitsuSave)          'シートオブジェクト(申込書印刷画面内容)
        
        '明細行の初期化
        Call subSaveDel
        
        'テキスト保存内容初期化
        wstTextK.Cells.ClearContents
        wstTextM.Cells.ClearContents
        wstMitsuSave.Cells.ClearContents
        
        Set wstTextK = Nothing
        Set wstTextM = Nothing
        Set wstMitsuSave = Nothing
        
        Unload Me
        frmTop.Show vbModeless
        
    End If
    
    On Error GoTo 0
    
    Exit Sub

Error:
    MsgBox "BtnBack_Click" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"
    
End Sub

'「計算用シート」ボタン押下
Private Sub BtnShisan_Click()
    Dim wsMeisai As Worksheet
    Dim wsTextK As Worksheet
    Dim wsTextM As Worksheet
    Dim wsSetting As Worksheet
    Dim wsKyoutsu As Worksheet
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim intMeisaiCnt As Integer
    Dim strTempK As String
    Dim strTempM As String

    On Error GoTo Error

    strTempK = ""
    strTempM = ""
    intMeisaiCnt = 0
    k = 0
    blnChouhyouflg = True
    
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(1, wsMeisai)        'シートオブジェクト(明細入力)
    Else
        Call subSetSheet(17, wsMeisai)       'シートオブジェクト(明細入力（ノンフリート）)
    End If
    Call subSetSheet(2, wsKyoutsu)      'シートオブジェクト(別紙　共通項目)
    Call subSetSheet(5, wsSetting)      'シートオブジェクト(別紙　各種設定)
    Call subSetSheet(6, wsTextK)        'シートオブジェクト(テキスト内容(共通))
    Call subSetSheet(7, wsTextM)        'シートオブジェクト(テキスト内容(明細))
    
    '明細入力画面ヘッダ設定
    Call subMeisaiUnProtect     'シートの保護の解除

    '保険期間
    wsMeisai.Range("B3") = "　保険期間　　：" & IIf(wsKyoutsu.Range("E2") = "", "", ("平成" & Format(CStr(Val(Left(wsKyoutsu.Range("E2"), 4)) - 1988), "00") & "年" & Mid(wsKyoutsu.Range("E2"), 5, 2) & "月" & Right(wsKyoutsu.Range("E2"), 2) & "日から1年間"))
    '受付区分
    wsMeisai.Range("E3") = "　受付区分　　：" & fncFindCode(wsKyoutsu.Range("A2"), "C")
    '被保険者
    wsMeisai.Range("G3") = "　被保険者　　　　　　：" & fncFindCode(wsKyoutsu.Range("B2"), "G")
    '保険種類
    wsMeisai.Range("B4") = "　保険種類　　：" & fncFindCode(wsKyoutsu.Range("C2"), "K")
    'フリート区分
    wsMeisai.Range("E4") = "　フリート区分：" & fncFindCode(wsKyoutsu.Range("D2"), "O")
    If FleetTypeFlg = 1 Then  'フリート
        '全車両一括付保特約
        wsMeisai.Range("G4") = "　全車両一括付保特約　：" & IIf(fncFindCode(wsKyoutsu.Range("D2"), "O") = "全車両一括" Or fncFindCode(wsKyoutsu.Range("D2"), "O") = "全車両連結合算", "有り", "無し")
    Else
        'ノンフリート多数割引（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        wsMeisai.Range("G4") = "　ノンフリート多数割引：" & fncFindCode(wsKyoutsu.Range("K2"), "AQ")
    End If
    '払込方法
    wsMeisai.Range("B5") = "　払込方法　　：" & fncFindCode(wsKyoutsu.Range("H2"), "AY")
    If FleetTypeFlg = 1 Then  'フリート
        '優良割引
        wsMeisai.Range("E5") = "　優良割引　　：" & IIf(Trim(wsKyoutsu.Range("I2")) = "", "", Trim(wsKyoutsu.Range("I2")) & "%")
    Else                      'ノンフリート
        '団体割増引（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
        wsMeisai.Range("E5") = "　団体割増引　：" & IIf(Trim(wsKyoutsu.Range("N2")) = "", "", Trim(wsKyoutsu.Range("N2")) & "%")
    End If
    If FleetTypeFlg = 1 Then  'フリート
        '第一種デメ割増
        wsMeisai.Range("G5") = "　第一種デメ割増  　　：" & IIf(Trim(wsKyoutsu.Range("J2")) = "", "", Trim(wsKyoutsu.Range("J2")) & "%")
        'ﾌﾘｰﾄ多数割引
        wsMeisai.Range("B6") = "　ﾌﾘｰﾄ多数割引：" & IIf(wsKyoutsu.Range("K2") = "2 ", "有り", "無し")
        'ﾌﾘｰﾄｺｰﾄﾞ
        wsMeisai.Range("E6") = "　ﾌﾘｰﾄｺｰﾄﾞ　　：" & wsKyoutsu.Range("L2")
    Else
        wsMeisai.Range("G5") = "　"
        wsMeisai.Range("B6") = "　"
        wsMeisai.Range("E6") = "　"
    End If
    
    '配列作成
    Dim varMeisai As Variant
    
    ReDim varMeisai(Val(wsTextK.Cells(1, 19)) \ Val(wsSetting.Range("B3").Value), Val(wsTextK.Cells(1, 19)))
    
'    For i = 1 To 20
    For i = 1 To 38
        strTempK = strTempK + CStr(wsTextK.Cells(1, i)) + ","
    Next i
    
    For i = 0 To Val(wsTextK.Cells(1, 19)) - 1
        
        If i = 0 Or i - 1 = 0 Then
        Else
            If (i - 1) Mod Val(wsSetting.Range("B3").Value) = 0 Then
                k = 0
                intMeisaiCnt = intMeisaiCnt + 1
            End If
        End If

        strTempM = ""

'        For j = 0 To 71
        For j = 0 To 81
            strTempM = strTempM + wsTextM.Cells(i + 1, j + 1) + ","
        Next j

        If k = 0 Then
            varMeisai(intMeisaiCnt, k) = strTempK
        End If
        varMeisai(intMeisaiCnt, k + 1) = strTempM

        k = k + 1

    Next i

    '明細入力シートに反映
    Call subBookUnProtect 'ブックの保護を解除 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    Call subSheetVisible(True)      'シート・ブックの表示
    Call subBookProtect   'ブックの保護 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    
    Call fncMeisaiEntry(varMeisai, Val(wsTextK.Cells(1, 19)))
    Call subSheetVisible(False)      'シート・ブックの非表示
    
'    Call subSetSheet(6, wsTextK)        'シートオブジェクト(テキスト内容(共通))
'    Call subSetSheet(7, wsTextM)        'シートオブジェクト(テキスト内容(明細))
    
    Call subBookUnProtect           'ブックの保護を解除
    Call subSheetVisible(True)      'シート・ブックの表示
    Call subBookProtect             'ブックの保護

    Call subMeisaiProtect       'シートの保護
    
    blnChouhyouflg = False
    
    Set wsKyoutsu = Nothing
    Set wsTextK = Nothing
    Set wsTextM = Nothing
    Set wsSetting = Nothing
    Set wsMeisai = Nothing

    '20190110対応
    MeisaiBackFlg = 0

    Unload Me

    On Error GoTo 0
    
    Exit Sub
    
Error:
    MsgBox "BtnShisan_Click" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"

End Sub


'「見積書・明細書」ボタン押下
Private Sub BtnMitsumori_Click()
    
    On Error GoTo Error
    
    'テキストファイル内容チェック
    If fncTextEntryErrChk(1) Then
        Exit Sub
    End If
    
    Me.Hide
    frmEntryMitsumori.Show vbModeless
    
    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "BtnMitsumori_Click" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"
    
End Sub


'「申込・明細書」ボタン押下
Private Sub BtnMoushikomi_Click()
    
    On Error GoTo Error
    
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    Dim wsSetting As Worksheet
    Dim intK As Integer
    Dim intM As Integer
    Dim intCnt As Integer
    Dim intMeisaiCnt As Integer
    Dim strTempK As String
    Dim strTempM As String
    
    Dim wsMeisai As Worksheet
    Dim wsTextK As Worksheet
    Dim wsTextM As Worksheet
    Dim wsMoushikomi As Worksheet
    Dim i As Integer
    Dim intStarRow As Integer
    Dim intSoufuho As Integer
    Dim intCol   As Integer
    Dim strRange As String
    Dim varRange As Variant
    Dim intRange As Variant
    
    
    'テキストファイル内容チェック
    If fncTextEntryErrChk(2) Then
        Exit Sub
    End If
    
    
    
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(18, wsMeisai)  'シートオブジェクト(明細書印刷)
    Else
        Call subSetSheet(19, wsMeisai)  'シートオブジェクト(明細書印刷（ノンフリート）)
    End If
    
    Call subSetSheet(5, wsSetting)      'シートオブジェクト(別紙　各種設定)
    Call subSetSheet(6, wsTextK)        'シートオブジェクト(テキスト内容(共通))
    Call subSetSheet(7, wsTextM)        'シートオブジェクト(テキスト内容(明細))
    Call subSetSheet(8, wsMoushikomi)   'シートオブジェクト(申込書印刷)
    
    
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    '配列作成
    ReDim varMeisai(Val(wsTextK.Cells(1, 19)) \ Val(wsSetting.Range("B3").Value), Val(wsTextK.Cells(1, 19)))
    
'    For intK = 1 To 20
    For intK = 1 To 38
        strTempK = strTempK + CStr(wsTextK.Cells(1, intK)) + ","
    Next intK
    
    For intK = 0 To Val(wsTextK.Cells(1, 19)) - 1
        
        If intK = 0 Or intK - 1 = 0 Then
        Else
            If (intK - 1) Mod Val(wsSetting.Range("B3").Value) = 0 Then
                intCnt = 0
                intMeisaiCnt = intMeisaiCnt + 1
            End If
        End If
        
        strTempM = ""
        
'        For intM = 0 To 71
        For intM = 0 To 81
            strTempM = strTempM + wsTextM.Cells(intK + 1, intM + 1) + ","
        Next intM
        
        If intCnt = 0 Then
            varMeisai(intMeisaiCnt, intCnt) = strTempK
            varMeisai(intMeisaiCnt, intCnt + 1) = strTempM
        Else
            varMeisai(intMeisaiCnt, intCnt + 1) = strTempM
        End If
        
        intCnt = intCnt + 1
        
    Next intK
    
    blnChouhyouflg = True
    Call subMeisaiUnProtect     'シートの保護の解除
    Call fncMeisaiEntry(varMeisai, Val(wsTextK.Cells(1, 19)))
    Call subMeisaiProtect       'シートの保護
    blnChouhyouflg = False
    


    'シートの保護の解除　（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
    Call subMeisaiPrtUnProtect
    
    ' 明細書印刷画面に取込テキストファイル情報を反映
    Call subSetSheet(6, wsTextK)
    Call subSetSheet(7, wsTextM)
    
    '総付保台数はテキストファイルの19カラム目
    intSoufuho = wsTextK.Cells(1, 19)
    
    '総付保台数を超える明細を非表示にする　（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
    Dim LastRow As Long
    Dim StartRow As Long
    Dim j As Integer
    Dim intCntMeisaiNo As Integer

    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    Application.EnableEvents = False
    If FleetTypeFlg = 1 Then  'フリート
    
        StartRow = 7
        intCntMeisaiNo = 1
        
        For j = StartRow To StartRow + intSoufuho - 1
            If j <> StartRow Then
                '明細行追加
                Call wsMeisai.Rows(StartRow).Copy(wsMeisai.Rows(j))
                
                '連番付与
                intCntMeisaiNo = intCntMeisaiNo + 1
                wsMeisai.Cells(j, 2) = Format(intCntMeisaiNo, "0000")
            End If
        Next j
        
         'テキストファイルより4項目をセット
        intStarRow = 6    '車両情報は7行目から
        For i = 1 To intSoufuho
            wsMeisai.Cells(intStarRow + i, 3) = wsTextM.Cells(i, 69)                    '' 登録番号（漢字）（C列＝3番目）
            wsMeisai.Cells(intStarRow + i, 4) = wsTextM.Cells(i, 72)                    '' 登録番号（カナ）（D列＝4番目）
            wsMeisai.Cells(intStarRow + i, 5) = wsTextM.Cells(i, 70)                    '' 車台番号　　　　（E列＝5番目）
            wsMeisai.Cells(intStarRow + i, 6) = wsTextM.Cells((intSoufuho * 2) + i, 71) '' 車検満了日　　　（F列＝6番目）
        Next i
    Else
        'ノンフリート
        '被保険者情報の行
        StartRow = 7
        LastRow = 15
        
        For j = (StartRow + intSoufuho) To LastRow
            wsMeisai.Rows(j).Hidden = True          '明細非表示
        Next j

        '車両情報の行
        StartRow = 19
        LastRow = 27
        For j = (StartRow + intSoufuho) To LastRow
            wsMeisai.Rows(j).Hidden = True          '明細非表示
        Next j

        '前契約情報の行
        StartRow = 31
        LastRow = 39
        For j = (StartRow + intSoufuho) To LastRow
            wsMeisai.Rows(j).Hidden = True          '明細非表示
        Next j
        
        
        'テキストファイルより12項目をセット
        intStarRow = 18    '車両情報は19行目から27行目
        For i = 1 To intSoufuho
            wsMeisai.Cells(6 + i, 3) = StrConv(StrConv(wsTextM.Cells(i, 75), vbKatakana), vbNarrow) ''被保険者住所（ｶﾅ）
            wsMeisai.Cells(6 + i, 8) = StrConv(StrConv(wsTextM.Cells(i, 76), vbKatakana), vbNarrow) ''被保険者氏名（ｶﾅ）
            wsMeisai.Cells(6 + i, 11) = wsTextM.Cells(i, 77)                            ''被保険者氏名（漢字）
            wsMeisai.Cells(6 + i, 24) = fncFindCode(wsTextM.Cells(i, 78), "DU")         ''免許証の色
            wsMeisai.Cells(6 + i, 27) = wsTextM.Cells((intSoufuho * 2) + i, 79)                            ''免許証有効期限
            wsMeisai.Cells(intStarRow + i, 3) = wsTextM.Cells(i, 69)                    '' 登録番号（漢字）（C列＝3番目）
            wsMeisai.Cells(intStarRow + i, 6) = wsTextM.Cells(i, 72)                    '' 登録番号（カナ）（F列＝6番目）
            wsMeisai.Cells(intStarRow + i, 8) = wsTextM.Cells(i, 70)                    '' 車台番号　　　　（H列＝8番目）
            wsMeisai.Cells(intStarRow + i, 9) = wsTextM.Cells((intSoufuho * 2) + i, 71) '' 車検満了日　　　（I列＝9番目）
            wsMeisai.Cells(intStarRow + i, 16) = StrConv(StrConv(wsTextM.Cells(i, 80), vbKatakana), vbNarrow) ''車両所有者氏名（カナ）
            wsMeisai.Cells(intStarRow + i, 24) = wsTextM.Cells(i, 81)                   ''車両所有者氏名（漢字）
            wsMeisai.Cells(intStarRow + i, 31) = fncFindCode(wsTextM.Cells(i, 82), "DY") ''所有権留保またはリース等

        Next i
    End If

    Application.EnableEvents = True

    '入力可能セル範囲（料率クラスや合計保険料等は入力不可）
    If FleetTypeFlg = 1 Then
        'フリート
        varRange = Array("$C$7:$J$7")
    Else
        'ノンフリート
        varRange = Array("$C$7:$AG$7", "$C$19:$AH$19", "$C$31:$C$31", "$E$31:$AC$31")
    End If

    'セル範囲設定が残っている場合、削除
    wsMeisai.Activate
    If wsMeisai.Protection.AllowEditRanges.Count = 0 Then
    Else
        wsMeisai.Protection.AllowEditRanges.item(1).Delete
    End If

    '入力可能セル範囲を総付保台数分広げる
    For i = 0 To UBound(varRange)
        If i = 0 Then
            intCol = Right(varRange(i), 1)
            intCol = intCol + intSoufuho - 1
            varRange(i) = Left(varRange(i), Len(varRange(i)) - 1) & intCol
        Else
            intCol = Right(varRange(i), 2)
            intCol = intCol + intSoufuho - 1
            varRange(i) = Left(varRange(i), Len(varRange(i)) - 2) & intCol
        End If
        
        strRange = strRange & "," & varRange(i)
    Next i
    
    strRange = Right(strRange, Len(strRange) - 1)
    
    '入力可能セル範囲を設定
    wsMeisai.Protection.AllowEditRanges.Add _
                Title:="EntryOK", _
                Range:=wsMeisai.Range(strRange)
    
    
    
    Call subMeisaiPrtProtect       'シートの保護
    
    
    Me.Hide
    frmEntryMoushikomi.Show vbModeless
    
    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "BtnMoushikomi_Click" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"

End Sub

''画面が閉じる前
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


