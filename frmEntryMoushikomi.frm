VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEntryMoushikomi 
   Caption         =   "�\�������"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10905
   OleObjectBlob   =   "frmEntryMoushikomi.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmEntryMoushikomi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intConfirmMsg As Integer


'�����\��
Private Sub UserForm_Initialize()
    Dim wsSetting As Worksheet

    On Error GoTo Error

    '2018/3 ���ذĖ��וt�@�\�ǉ�
    '�t�H�[���̃^�C�g���ݒ�
    If FleetTypeFlg = 1 Then
        frmEntryMoushikomi.Caption = frmEntryMoushikomi.Caption & "�i�t���[�g�_��j"
    Else
        frmEntryMoushikomi.Caption = frmEntryMoushikomi.Caption & "�i�m���t���[�g���וt�_��j"
    End If

    Call subSetSheet(5, wsSetting)      '�V�[�g�I�u�W�F�N�g(�ʎ��@�e��ݒ�)�i�����ӎ��������ۊǁj

    '�c�̓��͗��@�\���E��\���i2018/3 ���ذĖ��וt�@�\�ǉ��j
    If FleetTypeFlg = 1 Then
        FrameDantai.Visible = False
        FremeToriatukaiInfo.Top = 378
        FrameWarningMemo.Top = 450
        frmEntryMoushikomi.Height = 537
        frmEntryMoushikomi.ScrollBars = fmScrollBarsNone
        frmEntryMoushikomi.ScrollHeight = 0
        '�����ӎ���
        lblWarningMemo = wsSetting.Range("B6")
    Else
        lblWarningMemo = wsSetting.Range("B7")
    End If

    Set wsSetting = Nothing

    '�ۑ���񂪂���ꍇ���f����
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

                '�\������ʂ̊e���ڂɕۑ����e���Z�b�g
                txtShokenNo = varSaveContent(37)                                                       '�،��ԍ�
                If Not varSaveContent(20) = "" Then
'                    txtPostNo_zen = Left(varSaveContent(20), InStr(varSaveContent(20), "�|") - 1)      '�X�֔ԍ�
'                    txtPostNo_kou = Right(Replace(varSaveContent(20), "�|", "    "), 4)                '�X�֔ԍ�
                    txtPostNo_zen = RTrim(Left(varSaveContent(20), 3))                                '�X�֔ԍ�
                    If Len(varSaveContent(20)) > 3 Then
                        txtPostNo_kou = Mid(varSaveContent(20), 4)
                    Else
                        txtPostNo_kou = ""
                    End If
                Else
                    txtPostNo_zen = ""
                    txtPostNo_kou = ""
                End If
                txtKeiyakujusyo_kana = StrConv(StrConv(varSaveContent(21), vbKatakana), vbNarrow)      '�_��ҏZ��(��)
                txtKeiyakujusyo_kanji1 = RTrim(Left(varSaveContent(22), 40))                           '�_��ҏZ��(����)
                If Len(varSaveContent(22)) > 40 Then
                    txtKeiyakujusyo_kanji2 = Mid(varSaveContent(22), 41)                               '�_��ҏZ��(����)
                Else
                    txtKeiyakujusyo_kanji2 = ""
                End If
                txtHojin_kana = StrConv(StrConv(varSaveContent(23), vbKatakana), vbNarrow)             '�@�l��(��)
                txtHojin_kanji = varSaveContent(24)                                                    '�@�l��(����)
                txtYakusyoku_Shimei_kana = StrConv(StrConv(varSaveContent(25), vbKatakana), vbNarrow)  '��E���E����(��)
                txtYakusyoku_Shimei_kanji = varSaveContent(26)                                         '��E���E����(����)
                If Not varSaveContent(27) = "" Then
                    txtTelNo_Home_zen = Left(varSaveContent(27), InStr(varSaveContent(27), "-") - 1)  '�A����1�@����E�g��
                    txtTelNo_Home_chuu = Mid(Replace(varSaveContent(27), "-", "     "), _
                                                    InStr(varSaveContent(27), "-") + 5, 5)            '�A����1�@����E�g��
                    txtTelNo_Home_kou = Right(Replace(varSaveContent(27), "-", "    "), 4)            '�A����1�@����E�g��
                Else
                    txtTelNo_Home_zen = ""
                    txtTelNo_Home_chuu = ""
                    txtTelNo_Home_kou = ""
                End If
                If Not varSaveContent(28) = "" Then
                    txtTelNo_Kinmu_zen = Left(varSaveContent(28), InStr(varSaveContent(28), "-") - 1) '�A����2�@�Ζ���
                    txtTelNo_Kinmu_chuu = Mid(Replace(varSaveContent(28), "-", "     "), _
                                                  InStr(varSaveContent(28), "-") + 5, 5)              '�A����2�@�Ζ���
                    txtTelNo_Kinmu_kou = Right(Replace(varSaveContent(28), "-", "    "), 4)           '�A����2�@�Ζ���
                Else
                    txtTelNo_Kinmu_zen = ""
                    txtTelNo_Kinmu_chuu = ""
                    txtTelNo_Kinmu_kou = ""
                End If
                If Not varSaveContent(29) = "" Then
                    txtTelNo_Fax_zen = Left(varSaveContent(29), InStr(varSaveContent(29), "-") - 1)   '�A����3�@FAX
                    txtTelNo_Fax_chuu = Mid(Replace(varSaveContent(29), "-", "     "), _
                                                            InStr(varSaveContent(29), "-") + 5, 5)    '�A����3�@FAX
                    txtTelNo_Fax_kou = Right(Replace(varSaveContent(29), "-", "    "), 4)             '�A����3�@FAX
                Else
                    txtTelNo_Fax_zen = ""
                    txtTelNo_Fax_chuu = ""
                    txtTelNo_Fax_kou = ""
                End If
                '�c�̃R�[�h����Ј��R�[�h�܂ł̓m���t���[�g�̂�
                If FleetTypeFlg <> 1 Then
                     '�m���t���[�g�̏ꍇ
                    txtDantaimei = varSaveContent(30)                 '�c�̖�
                    txtDantaiCode = varSaveContent(31)                '�c�̃R�[�h
                    txtDantaiToku = varSaveContent(32)                '�c�̈��Ɋւ������
                    txtShozoku = varSaveContent(33)                   '�����R�[�h
                    txtSyainCode = varSaveContent(34)                 '�Ј��R�[�h
                    txtToriatsutenShop_code = varSaveContent(35)      '�戵�X�R�[�h
                    txtDairiShop_code = varSaveContent(36)            '�㗝�X�R�[�h
                Else
                    '�t���[�g�̏ꍇ
                    txtToriatsutenShop_code = varSaveContent(35)      '�戵�X�R�[�h
                    txtDairiShop_code = varSaveContent(36)            '�㗝�X�R�[�h
                End If
            End If
        End If
    Else
'        Dim varSaveContent As Variant
        varSaveContent = Split(strSaveContent, "/")

        '�\������ʂ̊e���ڂɕۑ����e���Z�b�g
        txtShokenNo = varSaveContent(0)                       '�،��ԍ�
        txtPostNo_zen = varSaveContent(1)                     '�X�֔ԍ�
        txtPostNo_kou = varSaveContent(2)                     '�X�֔ԍ�
        txtKeiyakujusyo_kana = varSaveContent(3)              '�_��ҏZ��(�J�i)
        txtKeiyakujusyo_kanji1 = varSaveContent(4)            '�_��ҏZ��(����)
        txtKeiyakujusyo_kanji2 = varSaveContent(5)            '�_��ҏZ��(����)
        txtHojin_kana = varSaveContent(6)                     '�@�l��(�J�i)
        txtHojin_kanji = varSaveContent(7)                    '�@�l��(����)
        txtYakusyoku_Shimei_kana = varSaveContent(8)          '��E���E����(�J�i)
        txtYakusyoku_Shimei_kanji = varSaveContent(9)         '��E���E����(����)
        txtTelNo_Home_zen = varSaveContent(10)                '�A����1�@����E�g��
        txtTelNo_Home_chuu = varSaveContent(11)               '�A����1�@����E�g��
        txtTelNo_Home_kou = varSaveContent(12)                '�A����1�@����E�g��
        txtTelNo_Kinmu_zen = varSaveContent(13)               '�A����2�@�Ζ���
        txtTelNo_Kinmu_chuu = varSaveContent(14)              '�A����2�@�Ζ���
        txtTelNo_Kinmu_kou = varSaveContent(15)               '�A����2�@�Ζ���
        txtTelNo_Fax_zen = varSaveContent(16)                 '�A����3�@FAX
        txtTelNo_Fax_chuu = varSaveContent(17)                '�A����3�@FAX
        txtTelNo_Fax_kou = varSaveContent(18)                 '�A����3�@FAX
        
        '�c�̃R�[�h����Ј��R�[�h�܂ł̓m���t���[�g�̂�
        If FleetTypeFlg <> 1 Then
             '�m���t���[�g�̏ꍇ
            txtDantaimei = varSaveContent(19)                 '�c�̖�
            txtDantaiCode = varSaveContent(20)                '�c�̃R�[�h
            txtDantaiToku = varSaveContent(21)                '�c�̈��Ɋւ������
            txtShozoku = varSaveContent(22)                   '�����R�[�h
            txtSyainCode = varSaveContent(23)                 '�Ј��R�[�h
        End If
        txtToriatsutenShop = varSaveContent(24)           '�戵�X
        txtToriatsutenShop_code = varSaveContent(25)      '�戵�X�R�[�h
        txtDairiShop = varSaveContent(26)                 '�㗝�X
        txtDairiShop_code = varSaveContent(27)            '�㗝�X�R�[�h
        txtBosyuuninID = varSaveContent(28)               '��W�lID
        
    End If
    
    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "UserForm_Initialize" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["
    
End Sub


'�u�~�v�{�^��������
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = vbFormControlMenu Then
        intConfirmMsg = MsgBox("�c�[�����I�����܂��B" & vbCrLf & "��낵���ł���?" & vbCrLf & "�����͓��e�͕ۑ�����܂���B", vbYesNo, "�m�F�_�C�A���O")
        If intConfirmMsg = 6 Then
            Cancel = False
            Call subAppClose
        Else
            Cancel = True
        End If
    End If
    
End Sub


'�u�߂�v�{�^������
Private Sub BtnBack_Click()
    Dim ctlFormCtrl As Control
    Dim wsSave As Worksheet     '��ʂ̏�Ԃ��ۑ�����Ă���V�[�g
    Dim i As Integer
    Dim j As Integer
    Dim objWorkSheet As Worksheet

    On Error GoTo Error

    intConfirmMsg = MsgBox("���͓��e���폜���Ē��[�I����ʂɑJ�ڂ��܂��B" & vbCrLf & "��낵���ł����H", vbYesNo, "�m�F�_�C�A���O")
    If intConfirmMsg = 6 Then

        '2018/3 ���ذĖ��וt�@�\�ǉ�
        If FleetTypeFlg = 1 Then  '�t���[�g
            Call subSetSheet(18, wsSave)       '�V�[�g�I�u�W�F�N�g(���׏����)
        Else
            Call subSetSheet(19, wsSave)       '�V�[�g�I�u�W�F�N�g(���׏�����i�m���t���[�g�j)
        End If

        '���׏�����V�[�g�̕ی�����i2018/3 ���ذĖ��וt�@�\�ǉ��j
        subMeisaiPrtUnProtect

        '�C�x���g����
        Application.EnableEvents = False

        If FleetTypeFlg = 1 Then
            '�t���[�g
            wsSave.Range(8 & ":" & wsSave.Rows.Count).Delete
            wsSave.Range("C7:F7") = ""

        Else

            '2018/3 ���ذĖ��וt�@�\�ǉ�

            '��ی��ҏ��̍s�iC��(3��j7�s�ڂ���AA��i27��j15�s�ڂ܂Łj���e����
             For i = 7 To 15
               For j = 3 To 27
                wsSave.Cells(i, j).Value = ""
              Next j
             Next i

            '�ԗ����̍s�iC��i3��j19�s�ڂ���AE��i31��j27�s�ڂ܂Łj���e����
             For i = 19 To 27
               For j = 3 To 31
                wsSave.Cells(i, j).Value = ""
              Next j
             Next i
                    
            '�O�_����̍s�iC��i3��j31�s�ڂ���AA��i27��j39�s�ڂ܂Łj���e����
             For i = 31 To 39
               For j = 3 To 27
                wsSave.Cells(i, j).Value = ""
              Next j
             Next i
             
            '�،��ԍ��Ɩ��הԍ��̊Ԃ̃n�C�t�������
             For i = 31 To 39
                wsSave.Cells(i, 4).Value = "-"
             Next i
        
            '��\���ɂ��Ă����s��S�ĕ\��
            wsSave.Rows.Hidden = False
        
            '���ד��͉�ʂ̃G���[�p���X�g������
            wsSave.OLEObjects("txtErrMsg").Object.Value = ""
            
        End If
    
        Application.EnableEvents = True     '�C�x���g�L��
    
        subMeisaiPrtProtect  '�V�[�g�̕ی�

        
        '��ʏ�����
        For Each ctlFormCtrl In frmEntryMoushikomi.Controls
            If TypeName(ctlFormCtrl) = "TextBox" Then _
                ctlFormCtrl.Value = ""
        Next ctlFormCtrl
        
        '2018/3 ���ذĖ��וt�@�\�ǉ�
        '�ۑ����폜
        Call subSetSheet(8, objWorkSheet) '�V�[�g�I�u�W�F�N�g�i�\���������ʓ��e�j
        objWorkSheet.Cells.ClearContents
        
        Unload Me
        frmPrintMenu.Show vbModeless
        
    End If
    
    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "BtnBack_Click" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["
    
End Sub

'�u���ցv�{�^������
Private Sub btnNext_Click()
            
    Dim wsMeisai As Worksheet


    On Error GoTo Error
    
    '�\���������ʂ̓��e��ۑ� �i2018/3 ���ذĖ��וt�@�\�ǉ��j
    If FleetTypeFlg = 1 Then
        Call subMoushikomiFormSet '�t���[�g
    Else
        Call subMoushikomiNonfleetFormSet '�m���t���[�g
    End If

    
    '�V�[�g�̕ی�̉����@�i2018/3 ���ذĖ��וt�@�\�ǉ��j
    Call subMeisaiPrtUnProtect
    
    '�t���[�g
    If FleetTypeFlg = 1 Then
        Call subSetSheet(18, wsMeisai)      '�V�[�g�I�u�W�F�N�g(���׏����)�@2018/3 �ذĖ��׈���@�\�ǉ�
    '�m���t���[�g
    Else
        Call subSetSheet(19, wsMeisai)      '�V�[�g�I�u�W�F�N�g(���׏�����i�m���t���[�g�j)�@2018/3 ���ذĖ��וt�@�\�ǉ�
    End If
    
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    blnMoushikomiflg = True
    
    
    Call subBookUnProtect '�u�b�N�̕ی������ 2018/3 ���ذĖ��וt�@�\�ǉ�
    Call subSheetVisible(True) '�V�[�g�E�u�b�N�̕\��
    Call subBookProtect   '�u�b�N�̕ی� 2018/3 ���ذĖ��וt�@�\�ǉ�
        
        
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then
        ThisWorkbook.Worksheets("���׏����").Activate '�t���[�g
    Else
        ThisWorkbook.Worksheets("���׏�����i�m���t���[�g�j").Activate '�m���t���[�g
    End If
        
    '�\���������ʂ̃G���[�p���X�g������
    wsMeisai.OLEObjects("txtErrMsg").Object.Value = ""
    wsMeisai.OLEObjects("txtErrMsg").Activate
    wsMeisai.Range("C7").Activate

    Call subMeisaiPrtProtect       '�V�[�g�̕ی�


    Set wsMeisai = Nothing

    Unload Me

    On Error GoTo 0

    Exit Sub
Error:
    MsgBox "�G���[�ԍ�:" & Err.Number & vbCrLf & _
           "�G���[�̎��:" & Err.Description, vbExclamation, "btnNext_Click"

End Sub


Private Sub subMoushikomiFormSet()
    Dim i                 As Integer
    Dim j                 As Integer
    Dim strSave           As String
    Dim wsMoushiSet       As Worksheet
    
    Call subSetSheet(8, wsMoushiSet)         '�V�[�g�I�u�W�F�N�g(�\���������ʓ��e)

    '�\���������ʕۑ����e�N���A
    wsMoushiSet.Cells.ClearContents

    '�\���������ʓ��e�Z�b�g
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
                                wsMoushiSet.Cells(i, j) = txtPostNo_zen.Value & "�|" & txtPostNo_kou.Value
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
                                wsMoushiSet.Cells(i, j) = txtTelNo_Home_zen.Value & "�|" & txtTelNo_Home_chuu.Value & "�|" & txtTelNo_Home_kou.Value
                            End If
                        Case 14, 15, 16
                            If txtTelNo_Kinmu_zen.Value = "" And txtTelNo_Kinmu_chuu.Value = "" And txtTelNo_Kinmu_kou.Value = "" Then
                                wsMoushiSet.Cells(i, j) = ""
                            Else
                                wsMoushiSet.Cells(i, j) = txtTelNo_Kinmu_zen.Value & "�|" & txtTelNo_Kinmu_chuu.Value & "�|" & txtTelNo_Kinmu_kou.Value
                            End If
                        Case 17, 18, 19
                            If txtTelNo_Fax_zen.Value = "" And txtTelNo_Fax_chuu.Value = "" And txtTelNo_Fax_kou.Value = "" Then
                                wsMoushiSet.Cells(i, j) = ""
                            Else
                                wsMoushiSet.Cells(i, j) = txtTelNo_Fax_zen.Value & "�|" & txtTelNo_Fax_chuu.Value & "�|" & txtTelNo_Fax_kou.Value
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


'2018/3 ���ذĖ��וt�@�\�ǉ�
Private Sub subMoushikomiNonfleetFormSet()

    Dim i                 As Integer
    Dim j                 As Integer
    Dim strSave           As String
    Dim wsMoushiSet       As Worksheet
    
    Call subSetSheet(8, wsMoushiSet)         '�V�[�g�I�u�W�F�N�g(�\���������ʓ��e)

    '�\���������ʓ��e�Z�b�g
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
    
    '�V�[�g��2�s�ڈȍ~�ɒ��[�Ŏg�p���鏑���𕹂��ĕۑ�
    For i = 2 To 5
        For j = 1 To 29
            Select Case i
                Case 3
                    Select Case j
                        Case 2, 3
                            If txtPostNo_zen.Value = "" And txtPostNo_kou.Value = "" Then
                                wsMoushiSet.Cells(i, j) = ""
                            Else
                                wsMoushiSet.Cells(i, j) = txtPostNo_zen.Value & "�|" & txtPostNo_kou.Value
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
                                wsMoushiSet.Cells(i, j) = txtTelNo_Home_zen.Value & "�|" & txtTelNo_Home_chuu.Value & "�|" & txtTelNo_Home_kou.Value
                            End If
                        Case 14, 15, 16
                            If txtTelNo_Kinmu_zen.Value = "" And txtTelNo_Kinmu_chuu.Value = "" And txtTelNo_Kinmu_kou.Value = "" Then
                                wsMoushiSet.Cells(i, j) = ""
                            Else
                                wsMoushiSet.Cells(i, j) = txtTelNo_Kinmu_zen.Value & "�|" & txtTelNo_Kinmu_chuu.Value & "�|" & txtTelNo_Kinmu_kou.Value
                            End If
                        Case 17, 18, 19
                            If txtTelNo_Fax_zen.Value = "" And txtTelNo_Fax_chuu.Value = "" And txtTelNo_Fax_kou.Value = "" Then
                                wsMoushiSet.Cells(i, j) = ""
                            Else
                                wsMoushiSet.Cells(i, j) = txtTelNo_Fax_zen.Value & "�|" & txtTelNo_Fax_chuu.Value & "�|" & txtTelNo_Fax_kou.Value
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

