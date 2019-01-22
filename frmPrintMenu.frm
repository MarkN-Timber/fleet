VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrintMenu 
   Caption         =   "���[�I��"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8145
   OleObjectBlob   =   "frmPrintMenu.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmPrintMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim intConfirmMsg As Integer

'�����\��
'2018/3 ���ذĖ��וt�@�\�ǉ�
Public Sub UserForm_Initialize()
    '�t�H�[���̃^�C�g���ݒ�
    If FleetTypeFlg = 1 Then
        frmPrintMenu.Caption = frmPrintMenu.Caption & "�i�t���[�g�_��j"
    Else
        frmPrintMenu.Caption = frmPrintMenu.Caption & "�i�m���t���[�g���וt�_��j"
    End If
End Sub


'�u�߂�v�{�^������
Private Sub BtnBack_Click()
    Dim wstTextK As Worksheet
    Dim wstTextM As Worksheet
    Dim wstMitsuSave As Worksheet

    On Error GoTo Error

    intConfirmMsg = MsgBox("���͓��e���폜����TOP��ʂɑJ�ڂ��܂��B" & vbCrLf & "��낵���ł����H", vbYesNo, "�m�F�_�C�A���O")
    If intConfirmMsg = 6 Then
        Call subSetSheet(6, wstTextK)          '�V�[�g�I�u�W�F�N�g(�e�L�X�g���e(����))
        Call subSetSheet(7, wstTextM)          '�V�[�g�I�u�W�F�N�g(�e�L�X�g���e(����))
        Call subSetSheet(8, wstMitsuSave)          '�V�[�g�I�u�W�F�N�g(�\���������ʓ��e)
        
        '���׍s�̏�����
        Call subSaveDel
        
        '�e�L�X�g�ۑ����e������
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
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["
    
End Sub

'�u�v�Z�p�V�[�g�v�{�^������
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
    
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(1, wsMeisai)        '�V�[�g�I�u�W�F�N�g(���ד���)
    Else
        Call subSetSheet(17, wsMeisai)       '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
    End If
    Call subSetSheet(2, wsKyoutsu)      '�V�[�g�I�u�W�F�N�g(�ʎ��@���ʍ���)
    Call subSetSheet(5, wsSetting)      '�V�[�g�I�u�W�F�N�g(�ʎ��@�e��ݒ�)
    Call subSetSheet(6, wsTextK)        '�V�[�g�I�u�W�F�N�g(�e�L�X�g���e(����))
    Call subSetSheet(7, wsTextM)        '�V�[�g�I�u�W�F�N�g(�e�L�X�g���e(����))
    
    '���ד��͉�ʃw�b�_�ݒ�
    Call subMeisaiUnProtect     '�V�[�g�̕ی�̉���

    '�ی�����
    wsMeisai.Range("B3") = "�@�ی����ԁ@�@�F" & IIf(wsKyoutsu.Range("E2") = "", "", ("����" & Format(CStr(Val(Left(wsKyoutsu.Range("E2"), 4)) - 1988), "00") & "�N" & Mid(wsKyoutsu.Range("E2"), 5, 2) & "��" & Right(wsKyoutsu.Range("E2"), 2) & "������1�N��"))
    '��t�敪
    wsMeisai.Range("E3") = "�@��t�敪�@�@�F" & fncFindCode(wsKyoutsu.Range("A2"), "C")
    '��ی���
    wsMeisai.Range("G3") = "�@��ی��ҁ@�@�@�@�@�@�F" & fncFindCode(wsKyoutsu.Range("B2"), "G")
    '�ی����
    wsMeisai.Range("B4") = "�@�ی���ށ@�@�F" & fncFindCode(wsKyoutsu.Range("C2"), "K")
    '�t���[�g�敪
    wsMeisai.Range("E4") = "�@�t���[�g�敪�F" & fncFindCode(wsKyoutsu.Range("D2"), "O")
    If FleetTypeFlg = 1 Then  '�t���[�g
        '�S�ԗ��ꊇ�t�ۓ���
        wsMeisai.Range("G4") = "�@�S�ԗ��ꊇ�t�ۓ���@�F" & IIf(fncFindCode(wsKyoutsu.Range("D2"), "O") = "�S�ԗ��ꊇ" Or fncFindCode(wsKyoutsu.Range("D2"), "O") = "�S�ԗ��A�����Z", "�L��", "����")
    Else
        '�m���t���[�g���������i2018/3 ���ذĖ��וt�@�\�ǉ��j
        wsMeisai.Range("G4") = "�@�m���t���[�g���������F" & fncFindCode(wsKyoutsu.Range("K2"), "AQ")
    End If
    '�������@
    wsMeisai.Range("B5") = "�@�������@�@�@�F" & fncFindCode(wsKyoutsu.Range("H2"), "AY")
    If FleetTypeFlg = 1 Then  '�t���[�g
        '�D�Ǌ���
        wsMeisai.Range("E5") = "�@�D�Ǌ����@�@�F" & IIf(Trim(wsKyoutsu.Range("I2")) = "", "", Trim(wsKyoutsu.Range("I2")) & "%")
    Else                      '�m���t���[�g
        '�c�̊������i2018/3 ���ذĖ��וt�@�\�ǉ��j
        wsMeisai.Range("E5") = "�@�c�̊������@�F" & IIf(Trim(wsKyoutsu.Range("N2")) = "", "", Trim(wsKyoutsu.Range("N2")) & "%")
    End If
    If FleetTypeFlg = 1 Then  '�t���[�g
        '����f������
        wsMeisai.Range("G5") = "�@����f������  �@�@�F" & IIf(Trim(wsKyoutsu.Range("J2")) = "", "", Trim(wsKyoutsu.Range("J2")) & "%")
        '�ذđ�������
        wsMeisai.Range("B6") = "�@�ذđ��������F" & IIf(wsKyoutsu.Range("K2") = "2 ", "�L��", "����")
        '�ذĺ���
        wsMeisai.Range("E6") = "�@�ذĺ��ށ@�@�F" & wsKyoutsu.Range("L2")
    Else
        wsMeisai.Range("G5") = "�@"
        wsMeisai.Range("B6") = "�@"
        wsMeisai.Range("E6") = "�@"
    End If
    
    '�z��쐬
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

    '���ד��̓V�[�g�ɔ��f
    Call subBookUnProtect '�u�b�N�̕ی������ 2018/3 ���ذĖ��וt�@�\�ǉ�
    Call subSheetVisible(True)      '�V�[�g�E�u�b�N�̕\��
    Call subBookProtect   '�u�b�N�̕ی� 2018/3 ���ذĖ��וt�@�\�ǉ�
    
    Call fncMeisaiEntry(varMeisai, Val(wsTextK.Cells(1, 19)))
    Call subSheetVisible(False)      '�V�[�g�E�u�b�N�̔�\��
    
'    Call subSetSheet(6, wsTextK)        '�V�[�g�I�u�W�F�N�g(�e�L�X�g���e(����))
'    Call subSetSheet(7, wsTextM)        '�V�[�g�I�u�W�F�N�g(�e�L�X�g���e(����))
    
    Call subBookUnProtect           '�u�b�N�̕ی������
    Call subSheetVisible(True)      '�V�[�g�E�u�b�N�̕\��
    Call subBookProtect             '�u�b�N�̕ی�

    Call subMeisaiProtect       '�V�[�g�̕ی�
    
    blnChouhyouflg = False
    
    Set wsKyoutsu = Nothing
    Set wsTextK = Nothing
    Set wsTextM = Nothing
    Set wsSetting = Nothing
    Set wsMeisai = Nothing

    '20190110�Ή�
    MeisaiBackFlg = 0

    Unload Me

    On Error GoTo 0
    
    Exit Sub
    
Error:
    MsgBox "BtnShisan_Click" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["

End Sub


'�u���Ϗ��E���׏��v�{�^������
Private Sub BtnMitsumori_Click()
    
    On Error GoTo Error
    
    '�e�L�X�g�t�@�C�����e�`�F�b�N
    If fncTextEntryErrChk(1) Then
        Exit Sub
    End If
    
    Me.Hide
    frmEntryMitsumori.Show vbModeless
    
    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "BtnMitsumori_Click" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["
    
End Sub


'�u�\���E���׏��v�{�^������
Private Sub BtnMoushikomi_Click()
    
    On Error GoTo Error
    
    '2018/3 ���ذĖ��וt�@�\�ǉ�
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
    
    
    '�e�L�X�g�t�@�C�����e�`�F�b�N
    If fncTextEntryErrChk(2) Then
        Exit Sub
    End If
    
    
    
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(18, wsMeisai)  '�V�[�g�I�u�W�F�N�g(���׏����)
    Else
        Call subSetSheet(19, wsMeisai)  '�V�[�g�I�u�W�F�N�g(���׏�����i�m���t���[�g�j)
    End If
    
    Call subSetSheet(5, wsSetting)      '�V�[�g�I�u�W�F�N�g(�ʎ��@�e��ݒ�)
    Call subSetSheet(6, wsTextK)        '�V�[�g�I�u�W�F�N�g(�e�L�X�g���e(����))
    Call subSetSheet(7, wsTextM)        '�V�[�g�I�u�W�F�N�g(�e�L�X�g���e(����))
    Call subSetSheet(8, wsMoushikomi)   '�V�[�g�I�u�W�F�N�g(�\�������)
    
    
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    '�z��쐬
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
    Call subMeisaiUnProtect     '�V�[�g�̕ی�̉���
    Call fncMeisaiEntry(varMeisai, Val(wsTextK.Cells(1, 19)))
    Call subMeisaiProtect       '�V�[�g�̕ی�
    blnChouhyouflg = False
    


    '�V�[�g�̕ی�̉����@�i2018/3 ���ذĖ��וt�@�\�ǉ��j
    Call subMeisaiPrtUnProtect
    
    ' ���׏������ʂɎ捞�e�L�X�g�t�@�C�����𔽉f
    Call subSetSheet(6, wsTextK)
    Call subSetSheet(7, wsTextM)
    
    '���t�ۑ䐔�̓e�L�X�g�t�@�C����19�J������
    intSoufuho = wsTextK.Cells(1, 19)
    
    '���t�ۑ䐔�𒴂��閾�ׂ��\���ɂ���@�i2018/3 ���ذĖ��וt�@�\�ǉ��j
    Dim LastRow As Long
    Dim StartRow As Long
    Dim j As Integer
    Dim intCntMeisaiNo As Integer

    '2018/3 ���ذĖ��וt�@�\�ǉ�
    Application.EnableEvents = False
    If FleetTypeFlg = 1 Then  '�t���[�g
    
        StartRow = 7
        intCntMeisaiNo = 1
        
        For j = StartRow To StartRow + intSoufuho - 1
            If j <> StartRow Then
                '���׍s�ǉ�
                Call wsMeisai.Rows(StartRow).Copy(wsMeisai.Rows(j))
                
                '�A�ԕt�^
                intCntMeisaiNo = intCntMeisaiNo + 1
                wsMeisai.Cells(j, 2) = Format(intCntMeisaiNo, "0000")
            End If
        Next j
        
         '�e�L�X�g�t�@�C�����4���ڂ��Z�b�g
        intStarRow = 6    '�ԗ�����7�s�ڂ���
        For i = 1 To intSoufuho
            wsMeisai.Cells(intStarRow + i, 3) = wsTextM.Cells(i, 69)                    '' �o�^�ԍ��i�����j�iC��3�Ԗځj
            wsMeisai.Cells(intStarRow + i, 4) = wsTextM.Cells(i, 72)                    '' �o�^�ԍ��i�J�i�j�iD��4�Ԗځj
            wsMeisai.Cells(intStarRow + i, 5) = wsTextM.Cells(i, 70)                    '' �ԑ�ԍ��@�@�@�@�iE��5�Ԗځj
            wsMeisai.Cells(intStarRow + i, 6) = wsTextM.Cells((intSoufuho * 2) + i, 71) '' �Ԍ��������@�@�@�iF��6�Ԗځj
        Next i
    Else
        '�m���t���[�g
        '��ی��ҏ��̍s
        StartRow = 7
        LastRow = 15
        
        For j = (StartRow + intSoufuho) To LastRow
            wsMeisai.Rows(j).Hidden = True          '���ה�\��
        Next j

        '�ԗ����̍s
        StartRow = 19
        LastRow = 27
        For j = (StartRow + intSoufuho) To LastRow
            wsMeisai.Rows(j).Hidden = True          '���ה�\��
        Next j

        '�O�_����̍s
        StartRow = 31
        LastRow = 39
        For j = (StartRow + intSoufuho) To LastRow
            wsMeisai.Rows(j).Hidden = True          '���ה�\��
        Next j
        
        
        '�e�L�X�g�t�@�C�����12���ڂ��Z�b�g
        intStarRow = 18    '�ԗ�����19�s�ڂ���27�s��
        For i = 1 To intSoufuho
            wsMeisai.Cells(6 + i, 3) = StrConv(StrConv(wsTextM.Cells(i, 75), vbKatakana), vbNarrow) ''��ی��ҏZ���i�Łj
            wsMeisai.Cells(6 + i, 8) = StrConv(StrConv(wsTextM.Cells(i, 76), vbKatakana), vbNarrow) ''��ی��Ҏ����i�Łj
            wsMeisai.Cells(6 + i, 11) = wsTextM.Cells(i, 77)                            ''��ی��Ҏ����i�����j
            wsMeisai.Cells(6 + i, 24) = fncFindCode(wsTextM.Cells(i, 78), "DU")         ''�Ƌ��؂̐F
            wsMeisai.Cells(6 + i, 27) = wsTextM.Cells((intSoufuho * 2) + i, 79)                            ''�Ƌ��ؗL������
            wsMeisai.Cells(intStarRow + i, 3) = wsTextM.Cells(i, 69)                    '' �o�^�ԍ��i�����j�iC��3�Ԗځj
            wsMeisai.Cells(intStarRow + i, 6) = wsTextM.Cells(i, 72)                    '' �o�^�ԍ��i�J�i�j�iF��6�Ԗځj
            wsMeisai.Cells(intStarRow + i, 8) = wsTextM.Cells(i, 70)                    '' �ԑ�ԍ��@�@�@�@�iH��8�Ԗځj
            wsMeisai.Cells(intStarRow + i, 9) = wsTextM.Cells((intSoufuho * 2) + i, 71) '' �Ԍ��������@�@�@�iI��9�Ԗځj
            wsMeisai.Cells(intStarRow + i, 16) = StrConv(StrConv(wsTextM.Cells(i, 80), vbKatakana), vbNarrow) ''�ԗ����L�Ҏ����i�J�i�j
            wsMeisai.Cells(intStarRow + i, 24) = wsTextM.Cells(i, 81)                   ''�ԗ����L�Ҏ����i�����j
            wsMeisai.Cells(intStarRow + i, 31) = fncFindCode(wsTextM.Cells(i, 82), "DY") ''���L�����ۂ܂��̓��[�X��

        Next i
    End If

    Application.EnableEvents = True

    '���͉\�Z���͈́i�����N���X�⍇�v�ی������͓��͕s�j
    If FleetTypeFlg = 1 Then
        '�t���[�g
        varRange = Array("$C$7:$J$7")
    Else
        '�m���t���[�g
        varRange = Array("$C$7:$AG$7", "$C$19:$AH$19", "$C$31:$C$31", "$E$31:$AC$31")
    End If

    '�Z���͈͐ݒ肪�c���Ă���ꍇ�A�폜
    wsMeisai.Activate
    If wsMeisai.Protection.AllowEditRanges.Count = 0 Then
    Else
        wsMeisai.Protection.AllowEditRanges.item(1).Delete
    End If

    '���͉\�Z���͈͂𑍕t�ۑ䐔���L����
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
    
    '���͉\�Z���͈͂�ݒ�
    wsMeisai.Protection.AllowEditRanges.Add _
                Title:="EntryOK", _
                Range:=wsMeisai.Range(strRange)
    
    
    
    Call subMeisaiPrtProtect       '�V�[�g�̕ی�
    
    
    Me.Hide
    frmEntryMoushikomi.Show vbModeless
    
    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "BtnMoushikomi_Click" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["

End Sub

''��ʂ�����O
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


