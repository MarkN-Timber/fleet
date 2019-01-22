VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKyoutsuu 
   Caption         =   "���ʍ���"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9630
   OleObjectBlob   =   "frmKyoutsuu.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmKyoutsuu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intConfirmMsg As Integer

'�����\��
Public Sub UserForm_Initialize()
    
    On Error GoTo Error
    
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    '�t�H�[���̃^�C�g���ݒ�
    If FleetTypeFlg = 1 Then
        frmKyoutsuu.Caption = frmKyoutsuu.Caption & "�i�t���[�g�_��j"
    Else
        frmKyoutsuu.Caption = frmKyoutsuu.Caption & "�i�m���t���[�g���וt�_��j"
    End If
    
    '���ד��̓V�[�g�擾
    Dim wsMeisai As Worksheet
    If FleetTypeFlg = 1 Then
        Call subSetSheet(1, wsMeisai)       '�V�[�g�I�u�W�F�N�g(���ד���)
    Else
        Call subSetSheet(17, wsMeisai)       '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)�i2018/3 ���ذĖ��וt�@�\�ǉ��j
    End If
    
    '�R�[�h�l�V�[�g�擾
    Dim wsCode As Worksheet
    If FleetTypeFlg = 1 Then
        Call subSetSheet(4, wsCode)         '�V�[�g�I�u�W�F�N�g(�ʎ��@�R�[�h�l)
    Else
        Call subSetSheet(16, wsCode)        '�V�[�g�I�u�W�F�N�g(�ʎ��@�R�[�h�l�i�m���t���[�g�j)�i2018/3 ���ذĖ��וt�@�\�ǉ��j
    End If
    
    '���ʍ��ڃV�[�g�擾
    Dim wsKyoutsuU As Worksheet
    Call subSetSheet(2, wsKyoutsuU)         '�V�[�g�I�u�W�F�N�g(�ʎ��@���ʍ���)

    '���t�ۑ䐔
    If blnFleetBtnFlg Then
        '�f�t�H���g�̍s��(10�s)��ݒ�
        txtSouFuhoDaisu.Value = "10"
        blnFleetBtnFlg = False
    ElseIf blnNonFleetBtnFlg Then
        '�f�t�H���g�̍s��(3�s)��ݒ�i2018/3 ���ذĖ��וt�@�\�ǉ��j
        txtSouFuhoDaisu.Value = "3"
        blnNonFleetBtnFlg = False
    Else
        '���ד��̓V�[�g�̑��t�ۑ䐔��ݒ�
        Dim objSouhuho As Object
        Set objSouhuho = wsMeisai.OLEObjects("txtSouhuho").Object
        txtSouFuhoDaisu.Value = Left(objSouhuho.Value, Len(objSouhuho.Value) - 2)
    End If
            
    '��t�敪
    With cmbUketsukekbn
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("B2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("B2").Column + 1).End(xlUp)).Value
        .ColumnWidths = "-1;0"
    End With
    
    '�ی����
    With cmbHokenSyurui
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("J2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("J2").Column + 1).End(xlUp)).Value
        .ColumnWidths = "-1;0"
    End With

    '�t���[�g�敪
    Dim intRows As Integer
    intRows = wsCode.Cells(wsCode.Rows.Count, wsCode.Range("N2").Column).End(xlUp).Row

    Dim strArray() As Variant

    If FleetTypeFlg = 1 Then    '2018/3 ���ذĖ��וt�@�\�ǉ�
        '�t���[�g
        Dim intListNo As Integer
        Dim intRow As Integer
        Dim strListName As String
        Dim strNonFleetName As String
        ReDim strArray(intRows - 3, 1) As Variant

        strNonFleetName = "�m���t���[�g"
        intRow = 0

        For intListNo = 0 To intRows - 2
            strListName = wsCode.Cells(intListNo + 2, wsCode.Range("N2").Column).Value
            If strListName <> strNonFleetName Then
                strArray(intRow, 0) = strListName
                strArray(intRow, 1) = wsCode.Cells(intListNo + 2, wsCode.Range("N2").Column + 1).Value
                intRow = intRow + 1
            End If
        Next intListNo
    Else
        '�m���t���[�g
        ReDim strArray(0, 1) As Variant
        strArray(0, 0) = wsCode.Cells(3, wsCode.Range("N2").Column).Value
        strArray(0, 1) = wsCode.Cells(3, wsCode.Range("O2").Column).Value
    End If

    With cmbFreetkbn
        .AddItem ""
        .List = strArray
        .ColumnWidths = "-1;0"
    End With

    '�t���[�g�敪��񊈐��ɕύX�@'2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg <> 1 Then
        cmbFreetkbn.ListIndex = 0
        cmbFreetkbn.Enabled = False
    End If

    '�������@
    With cmbHaraiHouhou
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("AX2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("AX2").Column + 1).End(xlUp)).Value
        .ColumnWidths = "-1;0"
    End With

    '��ی���
    optKojin = True

    '���ذđ��������i2018/3 ���ذĖ��וt�@�\�ǉ��j
    With cmbNonfleetTawari
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("AP2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("AP2").Column + 1).End(xlUp)).Value
        .ColumnWidths = "-1;0"
    End With

    '��ʉ����̃t���[��
    If FleetTypeFlg <> 1 Then
        '�m���t���[�g�i2018/3 ���ذĖ��וt�@�\�ǉ��j
        FrameFleet.Visible = False
        FrameNonFleet.Visible = True
    Else
        '�t���[�g�_��
        FrameFleet.Visible = True
        FrameNonFleet.Visible = False
    End If


    '�ۑ���񂪂���ꍇ���f����
    Dim strSaveContent As String
    If fncFormRef(1, strSaveContent) Then

    Else

        Dim varSaveContent As Variant
        varSaveContent = Split(strSaveContent, "/")

        Dim i As Integer

        '��t�敪
        For i = 0 To cmbUketsukekbn.ListCount - 1
            If cmbUketsukekbn.List(i, 1) = varSaveContent(0) Then
                cmbUketsukekbn.ListIndex = i
                Exit For
            End If
        Next

        '��ی���
        If varSaveContent(1) = "1" Then
            optKojin = True
            optHoujin = False
        Else
            optKojin = False
            optHoujin = True
        End If

        '�ی����
        For i = 0 To cmbHokenSyurui.ListCount - 1
            If cmbHokenSyurui.List(i, 1) = varSaveContent(2) Then
                cmbHokenSyurui.ListIndex = i
                Exit For
            End If
        Next

        '�t���[�g�敪
        For i = 0 To cmbFreetkbn.ListCount - 1
            If cmbFreetkbn.List(i, 1) = varSaveContent(3) Then
                cmbFreetkbn.ListIndex = i
                Exit For
            End If
        Next

        '�ی��n����
'        txtHokenStart_Nen = Format(Val(Left(varSaveContent(4), 4)) - 1988, "00")
        txtHokenStart_Nen = Format(Val(Left(varSaveContent(4), 4)) - 2000, "00") '�V�����Ή�
        txtHokenStart_Tsuki = Mid(varSaveContent(4), 5, 2)
        txtHokenStart_Hi = Right(varSaveContent(4), 2)

        '�������@
        For i = 0 To cmbHaraiHouhou.ListCount - 1
            If cmbHaraiHouhou.List(i, 1) = varSaveContent(7) Then
                cmbHaraiHouhou.ListIndex = i
                Exit For
            End If
        Next

        '�D�Ǌ���
        txtYuuryowari = varSaveContent(8)

        '����f������
        txtFirstDeme = varSaveContent(9)

        '�t���[�g��������
        chkFreetTasuu = IIf(varSaveContent(10) = "2 ", True, False)

        '�t���[�g�R�[�h
        txtFreetCode = varSaveContent(11)

        '���ذđ��������i2018/3 ���ذĖ��וt�@�\�ǉ��j
        For i = 0 To cmbNonfleetTawari.ListCount - 1
            If cmbNonfleetTawari.List(i, 1) = varSaveContent(12) Then
                cmbNonfleetTawari.ListIndex = i
                Exit For
            End If
        Next

        '�c�̊������i2018/3 ���ذĖ��וt�@�\�ǉ��j
        txtDantaiWarimashibiki = varSaveContent(13)

    End If

    Set wsMeisai = Nothing
    Set wsCode = Nothing
    Set wsKyoutsuU = Nothing
    Set objSouhuho = Nothing

    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "UserForm_Initialize" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["

End Sub

'�u���ցv�{�^������
Private Sub btnNext_Click()
    Dim strErrMsg As String

    strErrMsg = ""

    On Error GoTo Error

    '���̓`�F�b�N
    Call fncEntryCheckKyotsu(strErrMsg)

    '�G���[����
    If strErrMsg <> "" Then
        txtErrMsg = strErrMsg

        txtErrMsg.SetFocus
        txtErrMsg.SelStart = 0
        Exit Sub
    End If

    '���͏��擾
    Dim strSaveData As String
    Call subSaveData(strSaveData)

    '���͏��ۑ�
    Dim blnResult As Boolean
    blnResult = fncFormSave(1, strSaveData)

    '�V�[�g�̕ی�̉���
    Call subMeisaiUnProtect

    '���ד��͉�ʃw�b�_�ݒ�
    Dim wsMeisai As Worksheet
    '�t���[�g
    If FleetTypeFlg = 1 Then
        Call subSetSheet(1, wsMeisai)       '�V�[�g�I�u�W�F�N�g(���ד���)
    '�m���t���[�g
    Else
        Call subSetSheet(17, wsMeisai)      '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)�@2018/3 ���ذĖ��וt�@�\�ǉ�
    End If

    '�ی�����
'    wsMeisai.Range("B3") = "�@�ی����ԁ@�@�F����" & Format(Trim(txtHokenStart_Nen), "00") & "�N" & Format(Trim(txtHokenStart_Tsuki), "00") & "��" & Format(Trim(txtHokenStart_Hi), "00") & "������1�N��"
    wsMeisai.Range("B3") = "�@�ی����ԁ@�@�F20" & Format(Trim(txtHokenStart_Nen), "00") & "�N" & Format(Trim(txtHokenStart_Tsuki), "00") & "��" & Format(Trim(txtHokenStart_Hi), "00") & "������1�N��"
    '��t�敪
    wsMeisai.Range("E3") = "�@��t�敪�@�@�F" & cmbUketsukekbn
    '��ی���
    wsMeisai.Range("G3") = "�@��ی��ҁ@�@�@�@�@�@�F" & IIf(optKojin, "�l", "�@�l")
    '�ی����
    wsMeisai.Range("B4") = "�@�ی���ށ@�@�F" & cmbHokenSyurui
    '�t���[�g�敪
    wsMeisai.Range("E4") = "�@�t���[�g�敪�F" & cmbFreetkbn

    If FleetTypeFlg = 1 Then  '�t���[�g
        '�S�ԗ��ꊇ�t�ۓ���
        wsMeisai.Range("G4") = "�@�S�ԗ��ꊇ�t�ۓ���@�F" & IIf(cmbFreetkbn = "�S�ԗ��ꊇ" Or cmbFreetkbn = "�S�ԗ��A�����Z", "�L��", "����")
    Else
        '�m���t���[�g���������i2018/3 ���ذĖ��וt�@�\�ǉ��j
        wsMeisai.Range("G4") = "�@�m���t���[�g���������F" & cmbNonfleetTawari
    End If

    '�������@
    wsMeisai.Range("B5") = "�@�������@�@�@�F" & cmbHaraiHouhou

    If FleetTypeFlg = 1 Then  '�t���[�g
        '�D�Ǌ���
        wsMeisai.Range("E5") = "�@�D�Ǌ����@�@�F" & IIf(Trim(txtYuuryowari) = "", "", Trim(txtYuuryowari) & "%")
    Else                      '�m���t���[�g
        '�c�̊������i2018/3 ���ذĖ��וt�@�\�ǉ��j
        wsMeisai.Range("E5") = "�@�c�̊������@�F" & IIf(Trim(txtDantaiWarimashibiki) = "", "", Trim(txtDantaiWarimashibiki) & "%")
    End If

    If FleetTypeFlg = 1 Then  '�t���[�g
        '����f������
        wsMeisai.Range("G5") = "�@����f������  �@�@�F" & IIf(Trim(txtFirstDeme) = "", "", Trim(txtFirstDeme) & "%")
        '�ذđ�������
        wsMeisai.Range("B6") = "�@�ذđ��������F" & IIf(chkFreetTasuu, "�L��", "����")
        '�ذĺ���
        wsMeisai.Range("E6") = "�@�ذĺ��ށ@�@�F" & txtFreetCode
    Else
        wsMeisai.Range("G5") = "�@"
        wsMeisai.Range("B6") = "�@"
        wsMeisai.Range("E6") = "�@"
    End If

    Call subBookUnProtect           '�u�b�N�̕ی������
    Call subSheetVisible(True)      '�V�[�g�E�u�b�N�̕\��
    Call subBookProtect             '�u�b�N�̕ی�

    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        ThisWorkbook.Worksheets("���ד���").Activate
    Else
        ThisWorkbook.Worksheets("���ד��́i�m���t���[�g�j").Activate
    End If

    Dim intSouhuho As Integer
    Dim objSouhuho As Object
    Set objSouhuho = wsMeisai.OLEObjects("txtSouhuho").Object
    intSouhuho = Val(Left(objSouhuho.Value, Len(objSouhuho.Value) - 2))

    '���t�ۑ䐔���A���׍s��ǉ��E�폜
    If Val(Trim(txtSouFuhoDaisu)) > intSouhuho Then
        Call subMeisaiAdd(Val(Trim(txtSouFuhoDaisu)) - intSouhuho, "2")
    ElseIf Val(Trim(txtSouFuhoDaisu)) < Val(intSouhuho) Then
        Call subMeisaiDel(intSouhuho - Val(Trim(txtSouFuhoDaisu)), "2")
    End If

    '���ד��͉�ʂ̃G���[�p���X�g������
    wsMeisai.OLEObjects("txtErrMsg").Object.Value = ""
    wsMeisai.OLEObjects("txtErrMsg").Activate
    wsMeisai.Range("A1").Activate

    Set wsMeisai = Nothing
    Set objSouhuho = Nothing

    Call subMeisaiProtect       '�V�[�g�̕ی�

    '20190110�Ή�
    MeisaiBackFlg = 1

    Unload Me

    On Error GoTo 0

    Exit Sub
Error:
    MsgBox "�G���[�ԍ�:" & Err.Number & vbCrLf & _
           "�G���[�̎��:" & Err.Description, vbExclamation, "btnNext_Click"

End Sub


'���͏��擾
Private Sub subSaveData(ByRef strSaveData As String)

'    '�،��ԍ�
'    strSaveData = strSaveData & Trim(txtShokenNo) & "/"

    '��t�敪
    If cmbUketsukekbn.ListIndex > -1 Then
        strSaveData = strSaveData & cmbUketsukekbn.List(cmbUketsukekbn.ListIndex, 1) & "/"
    Else
        strSaveData = strSaveData & "" & "/"
    End If
    '��ی���_�l�@�l�敪
    strSaveData = strSaveData & IIf(optKojin, "1", "2") & "/"
    '�ی����
    If cmbHokenSyurui.ListIndex > -1 Then
        strSaveData = strSaveData & cmbHokenSyurui.List(cmbHokenSyurui.ListIndex, 1) & "/"
    Else
        strSaveData = strSaveData & "" & "/"
    End If
    '�t���[�g�E�m���t���[�g�敪
    If cmbFreetkbn.ListIndex > -1 Then
        strSaveData = strSaveData & cmbFreetkbn.List(cmbFreetkbn.ListIndex, 1) & "/"
    Else
        strSaveData = strSaveData & "" & "/"
    End If
    '�ی��n�����N����
'    strSaveData = strSaveData & Format(Val(Trim(txtHokenStart_Nen)) + 1988, "0000") & Format(Trim(txtHokenStart_Tsuki), "00") & Format(Trim(txtHokenStart_Hi), "00") & "/"
    strSaveData = strSaveData & Format(Val(Trim(txtHokenStart_Nen)) + 2000, "0000") & Format(Trim(txtHokenStart_Tsuki), "00") & Format(Trim(txtHokenStart_Hi), "00") & "/"
    strSaveData = strSaveData & "1" & "/"
    strSaveData = strSaveData & "0" & "/"
    If cmbHaraiHouhou.ListIndex > -1 Then
        strSaveData = strSaveData & cmbHaraiHouhou.List(cmbHaraiHouhou.ListIndex, 1) & "/"
    Else
        strSaveData = strSaveData & "" & "/"
    End If
    '�D�Ǌ���
    strSaveData = strSaveData & Trim(txtYuuryowari) & "/"
    '����f������
    strSaveData = strSaveData & Trim(txtFirstDeme) & "/"
    '�t���[�g��������
    strSaveData = strSaveData & IIf(chkFreetTasuu, "2 ", "") & "/"
    '�t���[�g�R�[�h
    strSaveData = strSaveData & Trim(txtFreetCode) & "/"

    '2018/3 ���ذĖ��וt�@�\�ǉ�
    '�m���t���[�g��������
    If cmbNonfleetTawari.ListIndex > -1 Then
        strSaveData = strSaveData & cmbNonfleetTawari.List(cmbNonfleetTawari.ListIndex, 1) & "/"
    Else
        strSaveData = strSaveData & "" & "/"
    End If
    '�c�̊�����
    strSaveData = strSaveData & Trim(txtDantaiWarimashibiki) & "/"

End Sub


'�u�߂�v�{�^������
Private Sub BtnBack_Click()
    Dim intMsgBox As Integer

    On Error GoTo Error

    intMsgBox = MsgBox("���͓��e���폜����TOP��ʂɑJ�ڂ��܂��B" & vbCrLf & "��낵���ł���?", vbYesNo, "�m�F�_�C�A���O")

    If intMsgBox = 6 Then

        '���׍s�̏�����
        Call subSaveDel

        Unload Me
        frmTop.Show vbModeless

    End If

    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "btnBack_Click" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["

End Sub


'�u�~�v�{�^������
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

Private Function fncEntryCheckKyotsu(ByRef strErrContent As String)
'�֐����FfncEntryErrCheckKyotsu
'���e�@�F�֘A�`�F�b�N�̃G���[�`�F�b�N���s���G���[������ꍇ��,�G���[���e��Ԃ��B
'�����@�FstrErrContent = :�G���[���e

    Dim strErrChkMsg As String
    Dim strErrKoumoku As String
    Dim strHokenStartErrMsg As String
    Dim strHokenStart As String

    strHokenStartErrMsg = ""
    strErrKoumoku = ""
    strErrChkMsg = ""

    '�E�ʃ`�F�b�N

    '���t�ۑ䐔
    strErrKoumoku = "�E���t�ۑ䐔" & vbCrLf
    strErrChkMsg = fncNeedCheck(Trim(txtSouFuhoDaisu))                      '�K�{�`�F�b�N

    If strErrChkMsg = "" Then
        strErrChkMsg = fncNumCheck(Trim(txtSouFuhoDaisu))                   '�����`�F�b�N
        If strErrChkMsg = "" Then

            '���l�`�F�b�N
            '2018/3 ���ذĖ��וt�@�\�ǉ�
'            strErrChkMsg = fncNumRangeCheck(Trim(txtSouFuhoDaisu), 1, 999)
            If FleetTypeFlg = 1 Then
                strErrChkMsg = fncNumRangeCheck(Trim(txtSouFuhoDaisu), 1, 999)
            Else
                strErrChkMsg = fncNonfleetSoufuhodaisuCheck(Trim(txtSouFuhoDaisu), 9)
                If strErrChkMsg = "" Then
                    strErrChkMsg = fncNumRangeCheck(Trim(txtSouFuhoDaisu), 1, 9)
                End If
            End If
        End If
    End If
    If strErrChkMsg = "" Then
    Else
        strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
    End If

    '��t�敪
    strErrChkMsg = ""

    strErrKoumoku = "�E��t�敪" & vbCrLf
    strErrChkMsg = fncNeedCheck(cmbUketsukekbn)                             '�K�{�`�F�b�N

    If strErrChkMsg = "" Then
    Else
        strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
    End If

    '��ی���
    strErrChkMsg = ""

    strErrKoumoku = "�E��ی���" & vbCrLf
    strErrChkMsg = fncNeedOptCheck(optKojin, optHoujin)                     '�K�{�`�F�b�N

    If strErrChkMsg = "" Then
    Else
        strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
    End If

    '�ی����
    strErrChkMsg = ""

    strErrKoumoku = "�E�ی����" & vbCrLf
    strErrChkMsg = fncNeedCheck(cmbHokenSyurui)                             '�K�{�`�F�b�N

    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If strErrChkMsg = "" Then
        strErrKoumoku = "���ی���ރG���[" & vbCrLf
        strErrChkMsg = fncHokenSyuruiCheck(cmbHokenSyurui.Value, optHoujin.Value) '�֘A�`�F�b�N
    End If

    If strErrChkMsg = "" Then
    Else
        strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
    End If

    '�t���[�g�敪
    strErrChkMsg = ""

    strErrKoumoku = "�E�t���[�g�敪" & vbCrLf
    strErrChkMsg = fncNeedCheck(cmbFreetkbn)                                '�K�{�`�F�b�N

    If strErrChkMsg = "" Then
    Else
        strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
    End If

    '�ی��n����_�N
    strErrChkMsg = ""

    strErrKoumoku = "�E�ی��n����_�N" & vbCrLf
    strErrChkMsg = fncNeedCheck(Trim(txtHokenStart_Nen))                      '�K�{�`�F�b�N

    If strErrChkMsg = "" Then
        strErrChkMsg = fncNumCheck(Trim(txtHokenStart_Nen))                   '�����`�F�b�N
        If strErrChkMsg = "" Then
            strErrChkMsg = fncNumRangeCheck(Trim(txtHokenStart_Nen), 1, 99)   '���l�`�F�b�N
        End If
    End If
    If strErrChkMsg = "" Then
    Else
        strHokenStartErrMsg = strHokenStartErrMsg & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
    End If

    '�ی��n����_��
    strErrChkMsg = ""

    strErrKoumoku = "�E�ی��n����_��" & vbCrLf
    strErrChkMsg = fncNeedCheck(Trim(txtHokenStart_Tsuki))                      '�K�{�`�F�b�N

    If strErrChkMsg = "" Then
        strErrChkMsg = fncNumCheck(Trim(txtHokenStart_Tsuki))                   '�����`�F�b�N
        If strErrChkMsg = "" Then
            strErrChkMsg = fncNumRangeCheck(Trim(txtHokenStart_Tsuki), 1, 12)   '���l�`�F�b�N
        End If
    End If
    If strErrChkMsg = "" Then
    Else
        strHokenStartErrMsg = strHokenStartErrMsg & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
    End If

    '�ی��n����_��
    strErrChkMsg = ""

    strErrKoumoku = "�E�ی��n����_��" & vbCrLf
    strErrChkMsg = fncNeedCheck(Trim(txtHokenStart_Hi))                    '�K�{�`�F�b�N

    If strErrChkMsg = "" Then
        strErrChkMsg = fncNumCheck(Trim(txtHokenStart_Hi))                 '�����`�F�b�N
        If strErrChkMsg = "" Then
            strErrChkMsg = fncNumRangeCheck(Trim(txtHokenStart_Hi), 1, 31) '���l�`�F�b�N
        End If
    End If
    If strErrChkMsg = "" Then
    Else
        strHokenStartErrMsg = strHokenStartErrMsg & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
    End If
    strErrContent = strErrContent & strHokenStartErrMsg

    '�ی��n����_�N����
    strErrChkMsg = ""
    strErrKoumoku = "�E�ی��n����" & vbCrLf
    '�V�����Ή���
'    If strHokenStartErrMsg = "" Then
'        strHokenStart = Format(Trim(txtHokenStart_Nen) + 1988, "0000") & "/" & _
'                        Format(Trim(txtHokenStart_Tsuki), "00") & "/" & _
'                        Format(Trim(txtHokenStart_Hi), "00")
'        strErrChkMsg = fncDateCheck(strHokenStart)                              '���t�`�F�b�N
'
'        If strErrChkMsg = "" Then
'            strErrChkMsg = fncShikiCheck(strHokenStart)                              '�ی��n���`�F�b�N
'        End If
'        If strErrChkMsg = "" Then
'        Else
'            strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
'        End If
'
'    End If
    If strHokenStartErrMsg = "" Then
        strHokenStart = Format(Trim(txtHokenStart_Nen) + 2000, "0000") & "/" & _
                        Format(Trim(txtHokenStart_Tsuki), "00") & "/" & _
                        Format(Trim(txtHokenStart_Hi), "00")
        strErrChkMsg = fncDateCheck(strHokenStart)                              '���t�`�F�b�N

        If strErrChkMsg = "" Then
            strErrChkMsg = fncShikiCheck(strHokenStart)                              '�ی��n���`�F�b�N
        End If
        If strErrChkMsg = "" Then
        Else
            strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
        End If

    End If
    '�V�����Ή���
    '�������@
    strErrChkMsg = ""

    strErrKoumoku = "�E�������@" & vbCrLf
    strErrChkMsg = fncNeedCheck(cmbHaraiHouhou)                             '�K�{�`�F�b�N

    If strErrChkMsg = "" Then
    Else
        strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
    End If

    '�D�Ǌ���
    If FleetTypeFlg = 1 Then '2018/3 ���ذĖ��וt�@�\�ǉ�
        If Trim(txtYuuryowari) = "" Then                                    '�u�����N�̏ꍇ�A�������Ȃ�
        Else
            strErrChkMsg = ""

            strErrKoumoku = "�E�D�Ǌ���" & vbCrLf
            strErrChkMsg = fncNumCheck(Trim(txtYuuryowari))                 '�����`�F�b�N
            If strErrChkMsg = "" Then
                strErrChkMsg = fncNumRangeCheck(Trim(txtYuuryowari), 1, 99) '���l�`�F�b�N
            End If
            If strErrChkMsg = "" Then
            Else
                strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
            End If
        End If
    End If

    '����f������
    If FleetTypeFlg = 1 Then '2018/3 ���ذĖ��וt�@�\�ǉ�
        If Trim(txtFirstDeme) = "" Then                                    '�u�����N�̏ꍇ�A�������Ȃ�
        Else
            strErrChkMsg = ""

            strErrKoumoku = "�E����f������" & vbCrLf
            strErrChkMsg = fncNumCheck(Trim(txtFirstDeme))                      '�����`�F�b�N
            If strErrChkMsg = "" Then
                strErrChkMsg = fncNumRangeCheck(Trim(txtFirstDeme), 1, 100)     '���l�`�F�b�N
            End If
            If strErrChkMsg = "" Then
            Else
                strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
            End If
        End If
    End If

    '�D�Ǌ����A����f������
    If FleetTypeFlg = 1 Then '2018/3 ���ذĖ��וt�@�\�ǉ�
        strErrChkMsg = ""

        strErrKoumoku = "�������G���[" & vbCrLf
        strErrChkMsg = fncWariCheck(Trim(txtYuuryowari), Trim(txtFirstDeme))                    '�֘A�`�F�b�N

        If strErrChkMsg = "" Then
        Else
            strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
        End If
    End If

    '�t���[�g��������
    If FleetTypeFlg = 1 Then '2018/3 ���ذĖ��וt�@�\�ǉ�
        strErrChkMsg = ""

        strErrKoumoku = "���t���[�g���������G���[" & vbCrLf
        strErrChkMsg = fncFleetTasuuCheck(cmbFreetkbn, chkFreetTasuu)                           '�֘A�`�F�b�N

        If strErrChkMsg = "" Then
        Else
            strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
        End If
    End If

    '�t���[�g�R�[�h
    If FleetTypeFlg = 1 Then '2018/3 ���ذĖ��וt�@�\�ǉ�
        strErrChkMsg = ""

        strErrKoumoku = "�E�t���[�g�R�[�h" & vbCrLf
        strErrChkMsg = fncNeedCheck(Trim(txtFreetCode))                     '�K�{�`�F�b�N

        If strErrChkMsg = "" Then
            strErrChkMsg = fncNumCheck(Trim(txtFreetCode))                  '�����`�F�b�N
            If strErrChkMsg = "" Then
                strErrChkMsg = fncKetaCheck(Trim(txtFreetCode), 5, "=") '����(����)�`�F�b�N
            End If
        End If
        If strErrChkMsg = "" Then
        Else
            strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
        End If
    End If

    '2018/3 ���ذĖ��וt�@�\�ǉ�
    '���ذđ�������
    If FleetTypeFlg = 2 Then
        strErrChkMsg = ""

        strErrKoumoku = "�����ذđ�������" & vbCrLf
        strErrChkMsg = fncNonfleetTawariCheck(cmbNonfleetTawari.Value, txtSouFuhoDaisu.Value) '���ذđ��������G���[�`�F�b�N

        If strErrChkMsg = "" Then
        Else
            strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
        End If
    End If

    '2018/3 ���ذĖ��וt�@�\�ǉ�
    '�c�̊�����
    If FleetTypeFlg = 2 Then
        strErrChkMsg = ""

        If txtDantaiWarimashibiki.Value = "" Then
        Else
            strErrKoumoku = "�E�c�̊�����" & vbCrLf

            strErrChkMsg = strErrChkMsg & fncDecimalCheck(txtDantaiWarimashibiki.Value) '�����`�F�b�N(�}�C�i�X�A�����_����)
            If strErrChkMsg = "" Then
                strErrChkMsg = strErrChkMsg & fncCommaCheck(txtDantaiWarimashibiki.Value) '�J���}�`�F�b�N
                If strErrChkMsg = "" Then
                    strErrChkMsg = strErrChkMsg & fncNumRangeCheck(Val(Trim(txtDantaiWarimashibiki.Value)), -50, 10) '���l�`�F�b�N
                End If
            End If
        End If

        If strErrChkMsg = "" Then
        Else
            strErrContent = strErrContent & strErrKoumoku & strErrChkMsg & vbCrLf & vbCrLf
        End If
    End If


    ''�G���[����
    If strErrContent <> "" Then
        strErrContent = Left(strErrContent, Len(strErrContent) - 2)
        strErrContent = strErrContent & String(62, "-") & vbCrLf & "�`�F�b�N����" & " [ " & Format(Time, "HH:MM:SS") & " ]"
    End If

End Function

Private Function fncNeedOptCheck(ByVal blnKojin As Boolean, ByVal blnHojin As Boolean) As String
'�֐����FfncNeedOptCheck
'���e�@�F�K�{�`�F�b�N
'�����@�F
'        blnValue       = ���͓��e
    
    fncNeedOptCheck = ""

    If blnKojin = False And blnHojin = False Then
         fncNeedOptCheck = " �K�{���͍��ڂł��B���͂��Ă��������B"
    End If
    
End Function

Private Function fncWariCheck(ByVal strYuryoWari As String, ByVal strDemeWari As String) As String
'�֐����FfncWariCheck
'���e�@�F�֘A�`�F�b�N
'�����@�F
'        blnValue       = ���͓��e
    
    fncWariCheck = ""

    If strYuryoWari <> "" And strDemeWari <> "" Then
         fncWariCheck = " �t���[�g�D�Ǌ����Ƒ���f�������͓����ɓ��͂ł��܂���B"
    End If

    If strYuryoWari = "" And strDemeWari = "" Then
         fncWariCheck = " �t���[�g�D�Ǌ����A����f�������̂����ꂩ����͂��Ă��������B"
    End If
    
End Function

Private Function fncFleetTasuuCheck(ByVal strFleetKbn As String, ByVal blnTasuuWari As Boolean) As String
'�֐����FfncFleetTasuuCheck
'���e�@�F�֘A�`�F�b�N
'�����@�F
'        blnValue       = ���͓��e

    fncFleetTasuuCheck = ""

    If strFleetKbn = "�S�ԗ��ꊇ" Or strFleetKbn = "�S�ԗ��A�����Z" Then
        If blnTasuuWari = False Then
            fncFleetTasuuCheck = " �S�ԗ��ꊇ�܂��͑S�ԗ��A���̏ꍇ�A�t���[�g�����������K�p�ł��܂��B"
        End If
    End If
    
End Function


