VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTop 
   Caption         =   "TOP"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7320
   OleObjectBlob   =   "frmTop.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intConfirmMsg As Integer
'テキストファイル出力時の判別フラグ
Dim blnTextErrflg As Boolean


'アクティブ
Private Sub UserForm_Activate()
    Dim wsSetting As Worksheet
    
    On Error GoTo Error
    
    Call subSetSheet(5, wsSetting)      'シートオブジェクト(別紙　各種設定)

    Me.Caption = "TOP  ( Ver." & wsSetting.Range("B2") & " )"
    
    'ショートカットキーの有効
    Call subShortCutKey(2)

    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    Application.DisplayStatusBar = True
    Application.DisplayFormulaBar = True
    Application.CommandBars("Ply").Enabled = True
    
    Set wsSetting = Nothing
    
    On Error GoTo 0
    
    Exit Sub

Error:
    MsgBox "UserForm_Activate" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"

End Sub

'管理者機能
Sub BtnAdministrator_Click()
    Dim strInput As String
    Dim strPassword As String
    Dim wsSetting As Worksheet
    
    On Error GoTo Error
    
    Call subSetSheet(5, wsSetting)      'シートオブジェクト(別紙　各種設定)
    
    BtnFleet.SetFocus
    
    strPassword = wsSetting.Range("B4").Value
    
    Set wsSetting = Nothing
    
    strInput = InputBox("パスワードを入力してください", "パスワード入力ダイアログ")
    
    If StrPtr(strInput) = 0 Then
        Exit Sub
    End If
    
    If strInput = strPassword Then
        
        Call subBookUnProtect           'ブックの保護を解除
        Call subMeisaiUnProtect         'シートの保護を解除
        Call subSheetVisible(True)      'シート・ブックの表示
        
        Application.ScreenUpdating = False                            '描画停止
        
        ThisWorkbook.Worksheets("別紙　コード値").Visible = True
        ThisWorkbook.Worksheets("別紙　各種設定").Visible = True
        ThisWorkbook.Worksheets("明細入力").Visible = False
        ThisWorkbook.Worksheets("見積書").Visible = True
        ThisWorkbook.Worksheets("車両明細書").Visible = True
        ThisWorkbook.Worksheets("契約申込書1枚目").Visible = True
        ThisWorkbook.Worksheets("契約申込書2枚目").Visible = True
        ThisWorkbook.Worksheets("明細書").Visible = True
        ThisWorkbook.Worksheets("申込書ＥＤＰ").Visible = True
        ThisWorkbook.Worksheets("明細書ＥＤＰ").Visible = True
        ThisWorkbook.Worksheets("別紙　見積書設定").Visible = True
        ThisWorkbook.Worksheets("別紙　車両明細書設定").Visible = True
        ThisWorkbook.Worksheets("別紙　申込書(1枚目)設定").Visible = True
        ThisWorkbook.Worksheets("別紙　申込書(2枚目)設定").Visible = True
        ThisWorkbook.Worksheets("別紙　明細書設定").Visible = True
        ThisWorkbook.Worksheets("別紙　申込書ＥＤＰ設定").Visible = True
        ThisWorkbook.Worksheets("別紙　明細書ＥＤＰ設定").Visible = True
        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        ThisWorkbook.Worksheets("明細入力（ノンフリート）").Visible = False
        ThisWorkbook.Worksheets("明細書印刷").Visible = False
        ThisWorkbook.Worksheets("明細書印刷（ノンフリート）").Visible = False
        ThisWorkbook.Worksheets("別紙　コード値（ノンフリート）").Visible = True
        ThisWorkbook.Worksheets("見積書（ノンフリート）").Visible = True
        ThisWorkbook.Worksheets("車両明細書（ノンフリート）").Visible = True
        ThisWorkbook.Worksheets("明細書（ノンフリート）").Visible = True
        ThisWorkbook.Worksheets("申込書ＥＤＰ（ノンフリート）").Visible = True
        ThisWorkbook.Worksheets("明細書ＥＤＰ（ノンフリート）").Visible = True
        ThisWorkbook.Worksheets("別紙　見積書設定（ノンフリート）").Visible = True
        ThisWorkbook.Worksheets("別紙　車両明細書設定（ノンフリート）").Visible = True
        ThisWorkbook.Worksheets("別紙　申込書(1枚目)設定（ノンフリート）").Visible = True
        ThisWorkbook.Worksheets("別紙　申込書(2枚目)設定（ノンフリート）").Visible = True
        ThisWorkbook.Worksheets("別紙　明細書設定（ノンフリート）").Visible = True
        ThisWorkbook.Worksheets("別紙　申込書ＥＤＰ設定（ノンフリート）").Visible = True
        ThisWorkbook.Worksheets("別紙　明細書ＥＤＰ設定（ノンフリート）").Visible = True
        ThisWorkbook.Worksheets("テキスト内容(共通)").Visible = False
        ThisWorkbook.Worksheets("テキスト内容(明細)").Visible = False
        
        
        ThisWorkbook.Worksheets("別紙　コード値").Activate
        ThisWorkbook.Worksheets("別紙　コード値").Range("A1").Select
        
        Application.ScreenUpdating = True                            '描画再開
        
        Call subBookProtect             'ブックの保護
        
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
        
        Windows(ThisWorkbook.Name).Visible = True
                
        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        FleetTypeFlg = 0
        
        Application.OnKey "%{q}", "fncAdmin"
        Me.Hide
        
    Else
        MsgBox "パスワードが正しくありません", vbOKOnly, "エラーダイアログ"
    End If
    
    On Error GoTo 0
    
    Exit Sub

Error:
    MsgBox "BtnAdministrator_Click" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"

End Sub

'フリート契約ボタン押下
Private Sub BtnFleet_Click()
    
     On Error GoTo Error
   
    'フリート契約・ノンフリート明細付契約を判断用フラグ(フリートを設定)
    FleetTypeFlg = 1
    blnFleetBtnFlg = True
    
    '共通入力画面を表示
    Call BtnPathDelete_Click
    Me.Hide
    
    frmKyoutsuu.Show vbModeless
    
    On Error GoTo 0
    
    Exit Sub
    
Error:
    MsgBox "BtnFleet_Click" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"


End Sub

'ノンフリート明細付契約ボタン押下
Private Sub BtnNonFleet_Click()
    
    On Error GoTo Error
   
    'フリート契約・ノンフリート明細付契約を判断用フラグ(フリートを設定)
    FleetTypeFlg = 2
    blnNonFleetBtnFlg = True
    
    '共通入力画面を表示
    Call BtnPathDelete_Click
    Me.Hide
    
    frmKyoutsuu.Show vbModeless
    
    On Error GoTo 0
    
    Exit Sub
    
Error:
    MsgBox "BtnFleet_Click" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"
    
    
End Sub

'クリアボタン押下
Private Sub BtnPathDelete_Click()
    '選択されているパスをテキストから削除
    With txtFilePath
        .Value = ""
        
        .SetFocus
        .SelStart = 0
    End With
    
End Sub

'ファイル選択ボタン押下時
Private Sub BtnFileSelect_Click()
    Dim varFilePath  As Variant
    Dim varTxtPath   As Variant

    On Error GoTo Error

    'TXTファイルを選択
    varTxtPath = Application.GetOpenFilename(FileFilter:="TXTファイル,*.txt", MultiSelect:=True)

    'テキストボックス出力用に、TXTファイルパスを変換
    If IsArray(varTxtPath) Then

        If txtFilePath.Value <> "" Then
            txtFilePath.Value = txtFilePath.Value & vbCrLf
        End If

        'ファイルパスをテキストに格納
        For Each varFilePath In varTxtPath
            txtFilePath.Value = txtFilePath.Value & varFilePath & vbCrLf
        Next varFilePath

        txtFilePath.Value = Left(txtFilePath.Value, Len(txtFilePath.Value) - 2)

        txtFilePath.SetFocus
        txtFilePath.SelStart = 0

    End If

    Set varTxtPath = Nothing

    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "BtnFileSelect_Click" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"

End Sub

'再開ボタン押下
Private Sub BtnReuse_Click()
    Dim i             As Integer
    Dim intFileNumber As Integer      'ファイル番号
    Dim intActRow     As Integer      '読み込み行
    Dim intMaxRow     As Integer      'ファイル最大行数
    Dim intCar        As Integer      '総付保台数
    Dim intMsgBox     As Integer
    Dim blnFiles      As Boolean
    Dim strFileName   As String
    Dim varFilePath   As Variant      'ファイルパス
    Dim varFiles      As Variant      'ファイルパス退避用
    
    On Error GoTo Error

    If (txtFilePath.Value) = "" Then Exit Sub

    varFiles = Split(txtFilePath.Value, vbCrLf)

    ReDim varFilePath(UBound(varFiles), 0)

    blnFiles = False
    intMaxRow = 0
    intCar = 0
    strTextName = ""

    'ファイル名取得
    For i = 0 To UBound(varFiles)
        If i = 1 Then
            blnFiles = True
        End If

        varFilePath(i, 0) = varFiles(i)

    Next i

    'ファイル昇順入れ替え
    If blnFiles Then
        Call subArrayVar(varFilePath)
    End If

    'ファイル名エラーチェック
    If fncFileNameCheck(varFilePath, blnFiles) Then
        intConfirmMsg = MsgBox("ファイル名が不正です。", vbOKOnly + vbExclamation, "エラーダイアログ")

        If intConfirmMsg = 1 Then
            blnTextErrflg = True
            Exit Sub
        End If

    End If

    'ファイル名取得
    strFileName = Mid(varFilePath(0, 0), InStrRev(varFilePath(0, 0), "\") + 1)
    If blnFiles Then
        '複数
        strTextName = Mid(Right(strFileName, 19), 1, 12)
    Else
        '単数
        strTextName = Mid(Right(strFileName, 16), 1, 12)
    End If

    'ファイル内容を取得
    For i = 0 To UBound(varFilePath)
        '空いているファイル番号を取得
        intFileNumber = FreeFile

        '入力ファイルをInputモードで開く
        Open varFilePath(i, 0) For Input Lock Write As #intFileNumber

        '読み込み開始行
        intActRow = 1

        Do Until EOF(1)
            If intActRow >= intMaxRow Then
                '配列長を変更
                ReDim Preserve varFilePath(UBound(varFiles), intActRow - 1)
                intMaxRow = intActRow
            End If

            '1行ごとに読み込み
            Line Input #intFileNumber, varFilePath(i, intActRow - 1)

            intActRow = intActRow + 1
        Loop

        Close #intFileNumber
    Next i

    '共通項目エラーチェック および
    'フリート、ノンフリート判定（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
    If fncKyotsuErr(varFilePath, intCar) Then
        intConfirmMsg = MsgBox("ファイル内容が不正です。", vbOKOnly + vbExclamation, "エラーダイアログ")

        If intConfirmMsg = 1 Then
            blnTextErrflg = True
            Exit Sub
        End If
    End If

    '明細項目エラーチェック
    If fncMeisaiErr(varFilePath) Then
        intConfirmMsg = MsgBox("ファイル内容が不正です。", vbOKOnly + vbExclamation, "エラーダイアログ")

        If intConfirmMsg = 1 Then
            blnTextErrflg = True
            Exit Sub
        End If
    End If
    
    If blnChouhyouflg Then
        intMsgBox = MsgBox("帳票出力します。" & vbCrLf & "よろしいですか?", vbYesNo, "確認ダイアログ")
    Else
        intMsgBox = MsgBox("再開します。" & vbCrLf & "よろしいですか?", vbYesNo, "確認ダイアログ")
    End If

    If intMsgBox = 6 Then

        Call subMeisaiUnProtect     'シートの保護の解除

        '共通項目保存用シートに反映
        Call fncKyotsuEntry(varFilePath, intCar)

        '明細入力シートに反映
        Call fncMeisaiEntry(varFilePath, intCar)

        Call subMeisaiProtect       'シートの保護

        '共通入力画面を表示
        Call BtnPathDelete_Click
        Me.Hide

        If blnChouhyouflg Then
            frmPrintMenu.Show vbModeless
        Else
            frmKyoutsuu.Show vbModeless
        End If

    End If

    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "BtnReuse_Click" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"

End Sub

'帳票出力ボタン押下
Private Sub BtnFormOutput_Click()

    On Error GoTo Error

    blnChouhyouflg = True

    Call BtnReuse_Click

    If blnTextErrflg Then
        blnTextErrflg = False
        blnChouhyouflg = False
        Exit Sub
    End If

    blnChouhyouflg = False

    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "BtnFormOutput_Click" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"

End Sub

'ファイル名エラーチェック
Private Function fncFileNameCheck(ByVal varFilePath As Variant, blnFiles As Boolean) As Boolean
    Dim i             As Integer
    Dim strFileName   As String
    Dim strFiles      As String
    Dim intFileNo     As Integer

    fncFileNameCheck = False
    strFileName = ""
    strFiles = ""
    intFileNo = 0

    '単数・複数ファイル確認
    If blnFiles Then
    '複数の場合
        'リスト内のファイル数分ループ
        For i = 0 To UBound(varFilePath)
            strFileName = Mid(varFilePath(i, 0), InStrRev(varFilePath(i, 0), "\") + 1)

            'ファイル形式確認("_")
            If Mid(Right(strFileName, 7), 1, 1) = "_" Then
            
                'ファイル形式(任意文字+"YYYYMMDDhhmm_01.txt")の場合+それ以外
                If IsDate(Format(Mid(Right(strFileName, 19), 1, 12), "##/##/## ##:##")) Then
                    intFileNo = Val(Mid(Right(strFileName, 6), 1, 2))
                    If (intFileNo = i + 1) Then
                        'ファイル形式("YYYYMMDDhhmm")に相違がある場合
                        If strFiles = "" Then
                            strFiles = Mid(Right(strFileName, 19), 1, 12)
                        Else
                            If strFiles = Mid(Right(strFileName, 19), 1, 12) Then
                            Else
                                fncFileNameCheck = True
                            End If
                        End If
                    Else
                        fncFileNameCheck = True
                    End If
                Else
                'ファイル名が日付型に変換できない場合
                    fncFileNameCheck = True
                End If
            
            Else
                fncFileNameCheck = True
            End If
            If fncFileNameCheck Then
                Exit Function
            End If
        Next i
    Else
    '単数の場合
        strFileName = Mid(varFilePath(0, 0), InStrRev(varFilePath(0, 0), "\") + 1)
        'ファイル形式(任意文字+YYYYMMDDhhmm.txt)の場合+それ以外
        If IsDate(Format(Mid(Right(strFileName, 16), 1, 12), "##/##/## ##:##")) Then
        Else
        'ファイル名が日付型に変換できない場合
            fncFileNameCheck = True
        End If
    End If
End Function

'選択ファイル順変更
Private Sub subArrayVar(ByRef varContent As Variant)
    Dim varArray As Variant
    Dim varSave As Variant
    Dim i As Integer
    Dim j As Integer

    varArray = varContent

    '連番確認
    For i = 0 To UBound(varArray, 1)
        For j = UBound(varArray, 1) To i Step -1
            If Val(Mid(Right(varArray(i, 0), 6), 1, 2)) > Val(Mid(Right(varArray(j, 0), 6), 1, 2)) Then
                varSave = varArray(i, 0)
                varArray(i, 0) = varArray(j, 0)
                varArray(j, 0) = varSave
            End If
        Next j
    Next i

    'ファイル入れ替え
    For i = 0 To UBound(varContent, 1)
        varContent(i, 0) = varArray(i, 0)
    Next i
End Sub

'共通項目エラーチェック
Function fncKyotsuErr(ByVal varContent As Variant, ByRef intCarAll As Integer) As Boolean
    Dim i               As Integer
    Dim j               As Integer
    Dim intCar          As Integer
    Dim strBefore       As String
    Dim strKyotsu       As String
    Dim varKyotsuRow    As Variant

    fncKyotsuErr = False
    strBefore = ""
    intCar = 0

    'エラーチェック
    For i = 0 To UBound(varContent, 1)
        strKyotsu = ""

        '共通部分のカンマの数が37ではない場合、エラー
        If UBound(Split(varContent(i, 0), ",")) = 37 Then
            varKyotsuRow = Split(varContent(i, 0), ",")
        Else
            fncKyotsuErr = True
            Exit Function
        End If

        '共通部分の総付保台数が3桁以上または0桁の場合、エラー
        If Len(varKyotsuRow(18)) = 0 Then
            fncKyotsuErr = True
            Exit Function
        ElseIf Len(varKyotsuRow(18)) >= 3 Then
            fncKyotsuErr = True
            Exit Function
        End If

        '共通部分の総付保台数が0の場合、エラー
        If Val(varKyotsuRow(18)) = "0" Then
            fncKyotsuErr = True
            Exit Function
        End If

        '共通部分にファイル毎で相違がある場合、エラー
        For j = 0 To UBound(varKyotsuRow)
            If j = 18 Then
            Else
                strKyotsu = strKyotsu + varKyotsuRow(j)
            End If
        Next j

        If strBefore = "" Then
            strBefore = strKyotsu
        Else
            If strBefore = strKyotsu Then
            Else
                fncKyotsuErr = True
                Exit Function
            End If
        End If

        intCar = intCar + varKyotsuRow(18)
    Next i

    intCarAll = intCar

    If intCarAll > 999 Then
        fncKyotsuErr = True
        Exit Function
    End If


    '共通部分よりフリート区分を判定（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
    If Val(varKyotsuRow(4)) = "1" Then
        FleetTypeFlg = 2               'ノンフリート
    Else
        FleetTypeFlg = 1               'フリート
    End If

    'ノンフリートかつ複数ファイル選択時はエラーとする（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
    If FleetTypeFlg = 2 And UBound(varContent, 1) > 0 Then
        fncKyotsuErr = True
        Exit Function
    End If

    Set varKyotsuRow = Nothing

End Function


'明細項目エラーチェック
Function fncMeisaiErr(ByVal varContent As Variant) As Boolean
    Dim i               As Integer
    Dim j               As Integer
    Dim intCar          As Integer
    Dim varKyotsuRow    As Variant
    
    fncMeisaiErr = False
    
    'エラーチェック
    For i = 0 To UBound(varContent, 1)
        intCar = 0
        varKyotsuRow = Split(varContent(i, 0), ",")
        
        For j = 1 To UBound(varContent, 2)
            If UBound(Split(varContent(i, j))) = -1 Then Exit For
            
            '明細部分のカンマの数が81ではない場合、エラー
            If UBound(Split(varContent(i, j), ",")) = 81 Then
                intCar = intCar + 1
            Else
                fncMeisaiErr = True
                Exit Function
            End If
            
        Next j
        
        '総付保台数と明細のレコードが合致しない場合、エラー
        If intCar = varKyotsuRow(18) Then
        Else
            fncMeisaiErr = True
            Exit Function
        End If
        
    Next i
    
    Set varKyotsuRow = Nothing
    
End Function

'共通項目保存用シートに反映
Private Sub fncKyotsuEntry(ByVal varContent As Variant, ByVal intCar As Integer)
    Dim i                As Integer
    Dim strKyotsu        As String
    Dim varKyotsuRow     As Variant
    Dim varKyotsuCol(31) As Variant
    Dim wstKyotsu As Worksheet
    Dim varMeisaiRow     As Variant '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    
    Call subSetSheet(2, wstKyotsu)        'シートオブジェクト(別紙　共通項目)
    
    varKyotsuRow = Split(varContent(0, 0), ",")

    varKyotsuCol(0) = varKyotsuRow(1)   '受付区分
    varKyotsuCol(1) = varKyotsuRow(2)   '被保険者
    varKyotsuCol(2) = varKyotsuRow(3)   '保険種類
    varKyotsuCol(3) = varKyotsuRow(4)   'フリート区分
    varKyotsuCol(4) = varKyotsuRow(5)   '保険始期日
    varKyotsuCol(5) = varKyotsuRow(10)  '保険期間
    varKyotsuCol(6) = varKyotsuRow(9)   '計算方法
    varKyotsuCol(7) = varKyotsuRow(13)  '払込方法
    varKyotsuCol(8) = varKyotsuRow(14)  '優良割引
    varKyotsuCol(9) = varKyotsuRow(15)  '第一種デメ割増
    varKyotsuCol(10) = varKyotsuRow(16) 'フリート多数割引
    varKyotsuCol(11) = varKyotsuRow(17) 'フリートコード
    varKyotsuCol(14) = varKyotsuRow(20) '郵便番号
    varKyotsuCol(15) = varKyotsuRow(21) '契約者住所（カナ）
    varKyotsuCol(16) = varKyotsuRow(22) '契約者住所（漢字）
'    varKyotsuCol(17) = varKyotsuRow(23) '契約者住所（漢字）
    varKyotsuCol(17) = varKyotsuRow(23) '法人名（カナ）
    varKyotsuCol(18) = varKyotsuRow(24) '法人名（漢字）
    varKyotsuCol(19) = varKyotsuRow(25) '役職名・氏名（カナ）
    varKyotsuCol(20) = varKyotsuRow(26) '役職名・氏名（漢字）
    varKyotsuCol(21) = varKyotsuRow(27) '連絡先１　自宅・携帯
    varKyotsuCol(22) = varKyotsuRow(28) '連絡先２　勤務先
    varKyotsuCol(23) = varKyotsuRow(29) '連絡先３　ＦＡＸ
    varKyotsuCol(24) = varKyotsuRow(30) '団体名
    varKyotsuCol(25) = varKyotsuRow(31) '団体コード
    varKyotsuCol(26) = varKyotsuRow(32) '団体扱に関する特約
    varKyotsuCol(27) = varKyotsuRow(33) '所属コード
    varKyotsuCol(28) = varKyotsuRow(34) '社員コード
    varKyotsuCol(29) = varKyotsuRow(35) '部課コード
    varKyotsuCol(30) = varKyotsuRow(36) '代理店コード
    varKyotsuCol(31) = varKyotsuRow(37) '証券番号

    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 2 Then
        varMeisaiRow = Split(varContent(0, 1), ",")
        varKyotsuCol(12) = varMeisaiRow(12) 'ノンフリート多数割引
        varKyotsuCol(13) = varMeisaiRow(13) '団体割増引
    End If

    '1行目の項目名が空になるまで2行目に値を格納
    For i = 1 To wstKyotsu.Range("A1").End(xlToRight).Column
        wstKyotsu.Cells(2, i) = varKyotsuCol(i - 1)
    Next i

    'テキストファイルの内容をシートに反映
'    If blnChouhyouflg Then
        '帳票出力テキスト
        For i = 1 To UBound(varKyotsuRow) + 1
            If i = 19 Then
                varKyotsuRow(i - 1) = intCar
            End If
            Call fncTextEdit(1, i, varKyotsuRow(i - 1), 0)
        Next i

'    End If

    Set wstKyotsu = Nothing

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

