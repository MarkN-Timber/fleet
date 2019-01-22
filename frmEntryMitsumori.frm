VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEntryMitsumori 
   Caption         =   "見積書印刷"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10755
   OleObjectBlob   =   "frmEntryMitsumori.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmEntryMitsumori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intConfirmMsg As Integer

'初期表示
'2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
Public Sub UserForm_Initialize()
    'フォームのタイトル設定
    If FleetTypeFlg = 1 Then
        frmEntryMitsumori.Caption = frmEntryMitsumori.Caption & "（フリート契約）"
    Else
        frmEntryMitsumori.Caption = frmEntryMitsumori.Caption & "（ノンフリート明細付契約）"
    End If
End Sub

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

'戻るボタン
Private Sub BtnBack_Click()

    Dim ctlFormCtrl As Control
    Dim objWorkSheet As Worksheet
    
    On Error GoTo Error
    
    intConfirmMsg = MsgBox("入力内容を削除して帳票選択画面に遷移します。" & vbCrLf & "よろしいですか？", vbYesNo, "確認ダイアログ")
    If intConfirmMsg = 6 Then
        '画面初期化
        For Each ctlFormCtrl In frmEntryMitsumori.Controls
            If TypeName(ctlFormCtrl) = "TextBox" Then _
                ctlFormCtrl.Value = ""
        Next ctlFormCtrl
        
        '保存情報削除
        Call subSetSheet(8, objWorkSheet) 'シートオブジェクト（申込書印刷画面内容）
        objWorkSheet.Cells.ClearContents
        
        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        'Me.Hide
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

'' 見積書印刷ボタン
Private Sub BtnPrintMitsumori_Click()

    Dim intSame         As Integer
    Dim i               As Integer
    Dim intTotalCar     As Integer
    Dim intMeisaiCnt    As Integer
    Dim intPageCnt      As Integer
    Dim wsSaveCnt       As Integer
    Dim strIndex        As String
    Dim strSave         As String
    Dim strNowTime      As String
    Dim strFileName     As String
    Dim strOutputPath   As String
    Dim strFilePath     As String
    Dim wsAll           As Worksheet
    Dim wsTextK         As Worksheet
    Dim wsChohyo        As Worksheet
    Dim wsCarMeisaiSet  As Worksheet
    Dim wstMitsuSave    As Worksheet
    Dim wsMitsumoriSet  As Worksheet
    Dim wsSetting       As Worksheet
    Dim blnFstflg       As Boolean
    Dim intDaishaSet    As Integer '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加

    On Error GoTo Error

    intConfirmMsg = MsgBox("見積書を印刷します。" & vbCrLf & "よろしいですか？", vbYesNo, "確認ダイアログ")
    If intConfirmMsg = 6 Then

        i = 16                                                  '設定画面の開始行
        intMeisaiCnt = 1                                        '明細行
        intPageCnt = 0                                          'ページ数
        intSame = 1                                             '同項目数(が続いた場合カウントアップ)
        strIndex = ""                                           '同項目比較変数
        wsSaveCnt = 0
        blnFstflg = True

        Call subSetSheet(5, wsSetting)      'シートオブジェクト(別紙　各種設定)

        'テキストファイル出力パス取得
        strOuptutPath = wsSetting.Range("B5").Value

        '出力ファイルパス
        If strOuptutPath = "" Then
            strFilePath = CreateObject("WScript.Shell").SpecialFolders.item("Desktop") & "\"
        Else
            strFilePath = strOuptutPath & "\"
        End If

        strFileName = strFilePath & strTextName & "_" & "見積書・明細書.pdf"
        
        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        '同名ファイル確認
        If Dir(strFileName) <> "" Then
            intConfirmMsg = MsgBox("同じ名前のファイルが既に存在します。上書きしますか？", vbYesNo, "確認ダイアログ")
            If intConfirmMsg = 7 Then Exit Sub
        End If

        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        'PDFファイルが開かれていることを確認
        If fncIsFileOpen(strFileName) Then
            intConfirmMsg = MsgBox("PDFファイルが開かれています。" & vbCrLf & "閉じてからご使用ください。", vbOKOnly, "通知ダイアログ")
            Exit Sub
        End If

        '警告チェック
        If fncTextEntryWarChk(1) Then
        End If

        '現在日時取得(年月日時分)
        strNowTime = Format(Now, "yyyymmddHHMM")

        Call subMeisaiUnProtect         'シートの保護の解除
        Call subBookUnProtect           'ブックの保護を解除
        Call subSheetVisible(True)      'シート・ブックの表示

        Application.ScreenUpdating = False

        'シート検索・削除
        For Each wsAll In ThisWorkbook.Worksheets
            If wsAll.Name = "見積書WK" Then
                Application.DisplayAlerts = False
                Worksheets("見積書WK").Delete
                Application.DisplayAlerts = True
            ElseIf wsAll.Name = "車両明細書WK" Then
                Application.DisplayAlerts = False
                Worksheets("車両明細書WK").Delete
                Application.DisplayAlerts = True
            End If
        Next wsAll

        'シートコピー
        With ThisWorkbook
            If FleetTypeFlg = 1 Then
                '見積書
                .Worksheets("見積書").Visible = True
                .Worksheets("見積書").Copy After:=.Worksheets(.Worksheets.Count)
                ActiveSheet.Name = "見積書WK"
                .Worksheets("見積書").Visible = False
                '車両明細書
                .Worksheets("車両明細書").Visible = True
                .Worksheets("車両明細書").Copy After:=.Worksheets(Worksheets.Count)
                ActiveSheet.Name = "車両明細書WK"
                .Worksheets("車両明細書").Visible = False
            Else
                '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                '見積書
                .Worksheets("見積書（ノンフリート）").Visible = True
                .Worksheets("見積書（ノンフリート）").Copy After:=.Worksheets(.Worksheets.Count)
                ActiveSheet.Name = "見積書WK"
                .Worksheets("見積書（ノンフリート）").Visible = False
                '車両明細書
                .Worksheets("車両明細書（ノンフリート）").Visible = True
                .Worksheets("車両明細書（ノンフリート）").Copy After:=.Worksheets(Worksheets.Count)
                ActiveSheet.Name = "車両明細書WK"
                .Worksheets("車両明細書（ノンフリート）").Visible = False
            End If
        End With
        
        Call subSetSheet(6, wsTextK)         'シートオブジェクト(テキスト内容(共通))
        Call subSetSheet(8, wstMitsuSave)    'シートオブジェクト（申込書印刷画面内容）

        '総付保台数取得
        intTotalCar = Val(wsTextK.Cells(1, 19))
        
        '申込書印刷画面保存内容クリア(申込書印刷画面内容)
        If FleetTypeFlg = 1 Then
            Call subSetSheet(9, wsMitsumoriSet)  'シートオブジェクト(別紙　見積書設定)
            Call subSetSheet(10, wsCarMeisaiSet) 'シートオブジェクト(別紙　車両明細書設定)
        Else
            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
            Call subSetSheet(20, wsMitsumoriSet)  'シートオブジェクト(別紙　見積書設定（ノンフリート）)
            Call subSetSheet(21, wsCarMeisaiSet) 'シートオブジェクト(別紙　車両明細書設定（ノンフリート）)
        End If
        Call subSetSheet(102, wsChohyo)       'シートオブジェクト(車両明細書WK)
        '
        wstMitsuSave.Cells.ClearContents
        
        '申込書印刷画面保存内容更新
        wstMitsuSave.Cells(1, 1) = txtKeiyakusyaHoujin.Value
        wstMitsuSave.Cells(1, 2) = txtKeiyakusyaDaihyou.Value
        wstMitsuSave.Cells(1, 3) = txtDairiten.Value
        wstMitsuSave.Cells(1, 4) = txtTantousya.Value
        wstMitsuSave.Cells(1, 5) = txtComment.Value
        
        '帳票設定読み込み・帳票作成
        With wsMitsumoriSet
            Do Until .Cells(i, 1).MergeArea(1) = ""
                
                '更新用
                Call subFormSetting(1, Val(.Cells(i, 1).MergeArea(1)), Val(.Cells(i, 2).MergeArea(1)), _
                                    CStr(.Cells(i, 4).MergeArea(1)), Val(.Cells(i, 5).MergeArea(1)), _
                                    CStr(.Cells(i, 6).MergeArea(1)), .Cells(i, 7).MergeArea(1), strNowTime, , _
                                    , , , , intSame, , , blnFstflg)
    
                i = i + 1
                
                If .Cells(i, 1).MergeArea(1) = "" And blnFstflg Then
                    i = 16
                    blnFstflg = False
                End If
                
            Loop
        End With
        
        '帳票設定読み込み・帳票作成
        Do Until intMeisaiCnt >= intTotalCar + 1
            i = 16
            blnFstflg = True
            intDaishaSet = 0 '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                    
            With wsCarMeisaiSet
                Do Until .Cells(i, 1).MergeArea(1) = ""
                    
                    '同項目が続いた場合、intSameをカウントアップ
                    '前項目が現在の項目と違う場合、intSameをページ開始明細行数とする
                    If strIndex = .Cells(i, 1).MergeArea(1) & "," & Val(.Cells(i, 2).MergeArea(1)) Then
                        If FleetTypeFlg = 1 Then
                            intSame = intSame + 1
                        Else
                            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                            If blnFstflg = False And strIndex = "2,42" Then '代車等セット特約　特殊処理（1項目で複数セットするため）
                                intDaishaSet = intDaishaSet + 1
                                If intDaishaSet >= 2 Then
                                    intSame = 1 + (intPageCnt * 2) + 1
                                End If
                            Else
                                intSame = intSame + 1
                            End If
                        End If
                    Else
                        strIndex = .Cells(i, 1).MergeArea(1) & "," & Val(.Cells(i, 2).MergeArea(1))
                        If FleetTypeFlg = 1 Then
                            intSame = 1 + (intPageCnt * 10)
                        Else
                            '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                            intSame = 1 + (intPageCnt * 2)
                        End If
                    End If

                    If CStr(.Cells(i, 4).MergeArea(1)) = "" Then
                        strSave = CStr(.Cells(i, 4).MergeArea(1))
                    Else
                        If Evaluate("ISREF(" & CStr(.Cells(i, 4).MergeArea(1)) & ")") = False Then
                            strSave = ""
                        Else
                            If FleetTypeFlg = 1 Then
                                strSave = CStr(.Cells(Val(.Range(CStr(.Cells(i, 4).MergeArea(1))).Row + (49 * intPageCnt)), _
                                                        .Range(CStr(.Cells(i, 4).MergeArea(1))).Column).Address(False, False))
                            Else
                                '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                                strSave = CStr(.Cells(Val(.Range(CStr(.Cells(i, 4).MergeArea(1))).Row + (44 * intPageCnt)), _
                                                        .Range(CStr(.Cells(i, 4).MergeArea(1))).Column).Address(False, False))
                            End If
                        End If
                    End If

                    '更新用
                    Call subFormSetting(2, Val(.Cells(i, 1).MergeArea(1)), Val(.Cells(i, 2).MergeArea(1)), _
                                        strSave, Val(.Cells(i, 5).MergeArea(1)), CStr(.Cells(i, 6).MergeArea(1)), _
                                        .Cells(i, 7).MergeArea(1), strNowTime, , , , , , _
                                        intSame, intMeisaiCnt, intPageCnt, blnFstflg)

                    i = i + 1

                    If .Cells(i, 1).MergeArea(1) = "" And blnFstflg Then
                        i = 16
                        If wsSaveCnt = 0 Then
                            wsSaveCnt = intMeisaiCnt
                            intMeisaiCnt = 1
                        Else
                        End If
                        blnFstflg = False

                    End If
                Loop

                intPageCnt = intPageCnt + 1

                If intMeisaiCnt < intTotalCar + 1 Then
                    If FleetTypeFlg = 1 Then
                        ' 1 ～ 49 行目を 50 行目 へ貼り付け
                        Application.DisplayAlerts = False
                        wsChohyo.Range("1:49").Copy
                        wsChohyo.Range(CStr(49 * Val(intPageCnt) + 1 & ":" & 49 * Val(intPageCnt) + 1)).Select
                        wsChohyo.Paste
                        Application.DisplayAlerts = True
                    Else
                        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
                        ' 1 ～ 44 行目を 45 行目 へ貼り付け
                        Application.DisplayAlerts = False
                        wsChohyo.Range("1:44").Copy
                        wsChohyo.Range(CStr(44 * Val(intPageCnt) + 1 & ":" & 44 * Val(intPageCnt) + 1)).Select
                        wsChohyo.Paste
                        Application.DisplayAlerts = True
                    End If
                End If
            
            End With
            
        Loop
        
        Dim strPrintSheet(1) As String
        strPrintSheet(0) = "見積書WK"
        strPrintSheet(1) = "車両明細書WK"
        
        ThisWorkbook.Worksheets(strPrintSheet).Select
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=strFileName
        
        'シート削除
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("見積書WK").Delete
        ThisWorkbook.Worksheets("車両明細書WK").Delete
        Application.DisplayAlerts = True
        
        Call subSheetVisible(False)      'シート・ブックの非表示
        Call subBookProtect              'ブックの保護
        Call subMeisaiProtect            'シートの保護
        
        Application.ScreenUpdating = True
        
        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        'MsgBox "印刷が完了しました。", vbOKOnly, "通知ダイアログ"
        MsgBox "PDFファイルを出力しました。", vbOKOnly, "通知ダイアログ"
        
        Set wsTextK = Nothing
        Set wsChohyo = Nothing
        Set wstMitsuSave = Nothing
        Set wsCarMeisaiSet = Nothing
    
    End If
    
    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "BtnPrintMitsumori_Click" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"
        
End Sub

