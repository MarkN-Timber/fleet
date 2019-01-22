VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOtherrate 
   Caption         =   "その他料率"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12855
   OleObjectBlob   =   "frmOtherrate.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmOtherrate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intConfirmMsg As Integer

'初期表示
Private Sub UserForm_Initialize()
    
    On Error GoTo Error
    
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    'フォームのタイトル設定
    If FleetTypeFlg = 1 Then
        frmOtherrate.Caption = frmOtherrate.Caption & "（フリート契約）"
    Else
        frmOtherrate.Caption = frmOtherrate.Caption & "（ノンフリート明細付契約）"
    End If
    
    Dim wsCode As Worksheet
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(4, wsCode)         'シートオブジェクト(別紙　コード値)
    Else
        Call subSetSheet(16, wsCode)         'シートオブジェクト(別紙　コード値（ノンフリート））
    End If

    '公有・準公有
    With cmbKouyukbn
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("AT2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("AT2").Column + 1).End(xlUp)).Value
        .ColumnWidths = "-1;0"
    End With
    
    'レンタカー、教習車、オープンポリシー多数割引非活性（2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加）
    If FleetTypeFlg = 2 Then
        chkRentacar.Enabled = False
        chkLearncar.Enabled = False
        chkOpenpol.Enabled = False
    End If
    
    
    '明細入力画面に情報がある場合反映する
    Dim wsMeisai As Worksheet
    '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
    If FleetTypeFlg = 1 Then  'フリート
        Call subSetSheet(1, wsMeisai)       'シートオブジェクト(明細入力)
    Else
        Call subSetSheet(17, wsMeisai)      'シートオブジェクト(明細入力（ノンフリート）)
    End If
    
    Dim strParams() As String
    Dim varParam As Variant
    
    varParam = wsMeisai.Cells(Val(Left(strTxtOther, InStr(strTxtOther, ":") - 1)), Val(Mid(strTxtOther, InStr(strTxtOther, ":") + 1, Len(strTxtOther) - InStr(strTxtOther, ":"))))
    
    If varParam <> "" Then
        strParams() = Split(varParam, "／")
        
        For Each varParam In strParams()
            If varParam = "沖縄" Then chkOkinawa = True
            If varParam = "レンタカー" Then chkRentacar = True
            If varParam = "教習車" Then chkLearncar = True
            If varParam = "ブーム対象外" Then chkBoom = True
            If varParam = "リースカーオープンポリシー" Then chkLeasecar = True
            If varParam = "オープンポリシー多数割引" Then chkOpenpol = True
            If varParam = "" Then chkOkinawa = True
            If varParam = "公有" Then cmbKouyukbn = "公有"
            If varParam = "準公有" Then cmbKouyukbn = "準公有"
            If varParam = "特種区分" Then chkSpecialkbn = True
        Next
    End If
    
    Set wsCode = Nothing
    Set wsMeisai = Nothing
    
    On Error GoTo 0
    
    Exit Sub
    
Error:
    MsgBox "UserForm_Initialize" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"

End Sub


'「セット」ボタン押下
Private Sub btnSet_Click()

    Dim intMsgBox As Integer
    
    On Error GoTo Error

    intMsgBox = MsgBox("入力した内容を明細入力画面に反映します。" & vbCrLf & "よろしいですか？", vbYesNo, "確認ダイアログ")
    
    If intMsgBox = 6 Then
    
        '選択情報取得
        Dim varParam As String
        varParam = varParam & IIf(chkOkinawa, "沖縄／", "")
        varParam = varParam & IIf(chkRentacar, "レンタカー／", "")
        varParam = varParam & IIf(chkLearncar, "教習車／", "")
        varParam = varParam & IIf(chkBoom, "ブーム対象外／", "")
        varParam = varParam & IIf(chkLeasecar, "リースカーオープンポリシー／", "")
        varParam = varParam & IIf(chkOpenpol, "オープンポリシー多数割引／", "")
        varParam = varParam & IIf(cmbKouyukbn <> "", cmbKouyukbn & "／", "")
        varParam = varParam & IIf(chkSpecialkbn, "特種区分／", "")
        If varParam <> "" Then
            varParam = Left(varParam, Len(varParam) - 1)
        End If
        
        '明細入力画面に設定
        Dim wsMeisai As Worksheet
        '2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        If FleetTypeFlg = 1 Then  'フリート
            Call subSetSheet(1, wsMeisai)       'シートオブジェクト(明細入力)
        Else
            Call subSetSheet(17, wsMeisai)      'シートオブジェクト(明細入力（ノンフリート）)
        End If
        
        Call subMeisaiUnProtect     'シートの保護の解除
        
        wsMeisai.Cells(Val(Left(strTxtOther, InStr(strTxtOther, ":") - 1)), Val(Mid(strTxtOther, InStr(strTxtOther, ":") + 1, Len(strTxtOther) - InStr(strTxtOther, ":")))) = varParam
        
        Set wsMeisai = Nothing
        
        Call subMeisaiProtect       'シートの保護
        
        'シート・ブックの表示
        Call subBookUnProtect 'ブックの保護を解除 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        Call subSheetVisible(True)
        Call subBookProtect   'ブックの保護 2018/3 ﾉﾝﾌﾘｰﾄ明細付機能追加
        
        '閉じる
        Unload Me
    End If

    On Error GoTo 0
    
    Exit Sub
    
Error:
    MsgBox "btnSet_Click" & vbCrLf & _
            "エラー番号:" & Err.Number & vbCrLf & _
            "エラーの種類:" & Err.Description, vbExclamation, "予期せぬエラー"


End Sub


'「戻る」ボタン押下
Private Sub BtnBack_Click()
    
    Dim intMsgBox As Integer
    
    On Error GoTo Error

    intMsgBox = MsgBox("入力内容を反映せずに明細入力画面に遷移します。" & vbCrLf & "よろしいですか？", vbYesNo, "確認ダイアログ")
    
    If intMsgBox = 6 Then
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
