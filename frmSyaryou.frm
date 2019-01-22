VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSyaryou 
   Caption         =   "車両情報取込"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9240
   OleObjectBlob   =   "frmSyaryou.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSyaryou"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strBeforeFileName As String
Dim intConfirmMsg As Integer


'初期表示
'2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
Public Sub UserForm_Initialize()
    'フォームのタイトル設定
    If FleetTypeFlg = 1 Then
        frmSyaryou.Caption = frmSyaryou.Caption & "（フリート契約）"
    Else
        frmSyaryou.Caption = frmSyaryou.Caption & "（ノンフリート明細付契約）"
    End If
End Sub


'「選択」ボタン押下
Private Sub btnSelect_Click()
    Dim strBeforeFilePath As String
    
    On Error GoTo Error
    
    '選択されていたファイルのパス・ファイル名を取得
    strBeforeFilePath = frmSyaryou.txtFilePath
    strBeforeFileName = Mid(strBeforeFilePath, InStrRev(strBeforeFilePath, "\") + 1)
    
    'ファイル選択ダイアログ
    Dim strFilePath As String
    strFilePath = Application.GetOpenFilename(FileFilter:="Excelファイル,*.xlsx;*.xls;*.xlsm")
    
    If strFilePath <> "False" Then
    
        '「取込行数」取得
        Dim wbWorkBook As Workbook
        Dim wsWorkSheet As Worksheet
        Dim strFileName As String
        
        strFileName = Mid(strFilePath, InStrRev(strFilePath, "\") + 1)
         
        Dim blnFileFlg As Boolean
        Dim WB As Workbook
        For Each WB In Workbooks
            If WB.Name = strFileName Then
                blnFileFlg = True
            End If
        Next WB
        
        Set WB = Nothing
        
        If blnFileFlg Then
            Set wbWorkBook = Workbooks(strFileName)
        Else
            Set wbWorkBook = Workbooks.Open(strFilePath)
            Windows(wbWorkBook.Name).Visible = False
            
            Call subFileClose(strBeforeFileName)
        End If
        Set wsWorkSheet = wbWorkBook.Sheets(1)
        
        Dim i As Long
        Dim lngTmpRow As Long
        Dim lngMaxRow As Long
        
        '車両情報ファイル最後列まで取得（フリート:12列、ノンフリート:12列）
            For i = 1 To 12
                lngTmpRow = wsWorkSheet.Cells(wsWorkSheet.Rows.Count, i).End(xlUp).Row
                If lngTmpRow > lngMaxRow Then
                    lngMaxRow = lngTmpRow
                End If
            Next
        
        Set wbWorkBook = Nothing
        Set wsWorkSheet = Nothing
        
        'フォーム設定
        frmSyaryou.txtFilePath = strFilePath
        frmSyaryou.txtFilePath.SetFocus
        frmSyaryou.txtFilePath.SelStart = 0
        frmSyaryou.lblMaxRow = "取込台数：" & CStr(lngMaxRow - 1) & "台"
    
    End If

    Exit Sub
Error:
     MsgBox "btnSelect_Click" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"
    
End Sub


'「一括セット」ボタン押下
Private Sub btnSyaryouSet_Click()
    
    'エラーチェック
    Dim strMaxRow As String
    Dim strMaxMeisai As String
    Dim intConfirmMsg As Integer
    
    On Error GoTo Error
    
    strMaxMeisai = Replace(lblMaxMeisai, "総付保台数：", "")
    strMaxMeisai = Replace(strMaxMeisai, "台", "")
    strMaxRow = Replace(lblMaxRow, "取込台数：", "")
    strMaxRow = Replace(strMaxRow, "台", "")
    
    If Val(strMaxRow) > Val(strMaxMeisai) Then
        intConfirmMsg = MsgBox("取込台数が総付保台数よりも多いです。", vbOKOnly + vbExclamation, "エラーダイアログ")
        Exit Sub
    End If
    
    intConfirmMsg = MsgBox("車両情報を取り込みます。" & vbCrLf & "よろしいですか?", vbYesNo, "確認ダイアログ")
    If intConfirmMsg = 6 Then
        
        '取込データの書き込み
        Dim wbWorkBook As Workbook
        Dim wsWorkSheet As Worksheet
        
        Dim blnFileFlg As Boolean
        Dim WB As Workbook
        Dim strFilePath As String
        Dim strFileName As String
        
        strFilePath = frmSyaryou.txtFilePath
        
        strFileName = Mid(strFilePath, InStrRev(strFilePath, "\") + 1)
        
        For Each WB In Workbooks
            If WB.Name = strFileName Then
                blnFileFlg = True
            End If
        Next WB
        
        Set WB = Nothing
        
        If blnFileFlg Then
            Set wbWorkBook = Workbooks(strFileName)
        Else
            Set wbWorkBook = Workbooks.Open(txtFilePath)
            Windows(wbWorkBook.Name).Visible = False
            
            Call subFileClose(strBeforeFileName)
        End If
        Set wsWorkSheet = wbWorkBook.Sheets(1)
    
        Dim wsMeisai As Worksheet
        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        If FleetTypeFlg = 1 Then  'フリート
            Call subSetSheet(1, wsMeisai)       'シートオブジェクト(明細入力)
        Else
            Call subSetSheet(17, wsMeisai)      'シートオブジェクト(明細入力（ノンフリート）)
        End If
    
        Call subMeisaiUnProtect     'シートの保護の解除
    
        Dim i As Long
        For i = 0 To strMaxRow - 1
            wsMeisai.Cells(21 + i, 3) = wsWorkSheet.Cells(2 + i, 1)  '車検満了日
            wsMeisai.Cells(21 + i, 5) = wsWorkSheet.Cells(2 + i, 2)   '車名
            wsMeisai.Cells(21 + i, 6) = wsWorkSheet.Cells(2 + i, 3)   '登録番号
            wsMeisai.Cells(21 + i, 7) = wsWorkSheet.Cells(2 + i, 4)   '車台番号
            wsMeisai.Cells(21 + i, 8) = wsWorkSheet.Cells(2 + i, 5)   '型式
            wsMeisai.Cells(21 + i, 9) = wsWorkSheet.Cells(2 + i, 6)   '仕様
            wsMeisai.Cells(21 + i, 10) = wsWorkSheet.Cells(2 + i, 7)  '初度登録
            wsMeisai.Cells(21 + i, 11) = wsWorkSheet.Cells(2 + i, 8)  '車検満了日
            wsMeisai.Cells(21 + i, 12) = wsWorkSheet.Cells(2 + i, 9)  '改造・不明車
            wsMeisai.Cells(21 + i, 13) = wsWorkSheet.Cells(2 + i, 10) '排気量
            wsMeisai.Cells(21 + i, 14) = wsWorkSheet.Cells(2 + i, 11) '2.5リットル超ディーゼル自小乗
            If FleetTypeFlg = 1 Then
                wsMeisai.Cells(21 + i, 26) = wsWorkSheet.Cells(2 + i, 12) '車両保険金額
            ElseIf FleetTypeFlg = 2 Then
                wsMeisai.Cells(21 + i, 34) = wsWorkSheet.Cells(2 + i, 12) '車両保険金額
            End If
        Next
        
        If Windows(strFileName).Visible = False Then
            subFileClose (strFileName)
        End If
        
        'シート・ブックの表示
        Call subBookUnProtect 'ブックの保護を解除 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        Call subSheetVisible(True)
        Call subBookProtect   'ブックの保護 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        
        '明細入力画面のエラー用リスト初期化
        wsMeisai.OLEObjects("txtErrMsg").Object.Value = ""
        wsMeisai.OLEObjects("txtErrMsg").Activate
        wsMeisai.Range("A1").Activate
        
        Set wbWorkBook = Nothing
        Set wsWorkSheet = Nothing
        Set wsMeisai = Nothing
        
        Call subMeisaiProtect     'シートの保護
            
        Unload Me
    Else
    End If
    
    Exit Sub
    
Error:
     MsgBox "btnSyaryouSet_Click" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"

    
End Sub


'「戻る」ボタン押下
Private Sub BtnBack_Click()

    On Error GoTo Error
    
    Dim intMsgBox As Integer
    intMsgBox = MsgBox("車両情報ファイル内容を反映せずに明細入力画面に遷移します。" & vbCrLf & "よろしいですか？", vbYesNo, "確認ダイアログ")
    
    Dim strFileName As String

    strFileName = Mid(frmSyaryou.txtFilePath, InStrRev(frmSyaryou.txtFilePath, "\") + 1)
    If intMsgBox = 6 Then
        
        If Windows(strFileName).Visible = False Then
            subFileClose (strFileName)
        End If

        'シート・ブックの表示
        Call subBookUnProtect 'ブックの保護を解除 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        Call subSheetVisible(True)
        Call subBookProtect   'ブックの保護 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
            
        Unload Me
    End If
    
    Exit Sub
Error:
     MsgBox "BtnBack_Click" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"
            
End Sub


'「×」ボタン押下
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = vbFormControlMenu Then
        intConfirmMsg = MsgBox("ツールを終了します。" & vbCrLf & "よろしいですか?" & vbCrLf & "※入力内容は保存されません。", vbYesNo, "確認ダイアログ")
        If intConfirmMsg = 6 Then
            Dim strFileName As String
            
            strFileName = Mid(frmSyaryou.txtFilePath, InStrRev(frmSyaryou.txtFilePath, "\") + 1)
            If Windows(strFileName).Visible = False Then
                subFileClose (strFileName)
            End If
            
            Cancel = False
            Call subAppClose
        Else
            Cancel = True
        End If
    End If
    
End Sub


'選択して開いたファイルを閉じる
Public Sub subFileClose(ByVal strCloseFileName As String)
    Dim sOpenbookAll As Workbook
    Dim sOpenBookSub As Variant
    Dim blnOpenBookFlg As Boolean

    '選択したエクセルファイルがまだ開かれていることを確認
    For Each sOpenbookAll In Workbooks
        If sOpenbookAll.Name = strCloseFileName Then
            blnOpenBookFlg = True
            Exit For
        End If
    Next sOpenbookAll
    
    If blnOpenBookFlg Then
        'エクセルファイルを閉じる
        Application.DisplayAlerts = False
        Workbooks(strCloseFileName).Close
        Application.DisplayAlerts = True
    End If
    
    Set sOpenbookAll = Nothing
    Set sOpenBookSub = Nothing
    
End Sub


