VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOtherrate 
   Caption         =   "���̑�����"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12855
   OleObjectBlob   =   "frmOtherrate.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmOtherrate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intConfirmMsg As Integer

'�����\��
Private Sub UserForm_Initialize()
    
    On Error GoTo Error
    
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    '�t�H�[���̃^�C�g���ݒ�
    If FleetTypeFlg = 1 Then
        frmOtherrate.Caption = frmOtherrate.Caption & "�i�t���[�g�_��j"
    Else
        frmOtherrate.Caption = frmOtherrate.Caption & "�i�m���t���[�g���וt�_��j"
    End If
    
    Dim wsCode As Worksheet
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(4, wsCode)         '�V�[�g�I�u�W�F�N�g(�ʎ��@�R�[�h�l)
    Else
        Call subSetSheet(16, wsCode)         '�V�[�g�I�u�W�F�N�g(�ʎ��@�R�[�h�l�i�m���t���[�g�j�j
    End If

    '���L�E�����L
    With cmbKouyukbn
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("AT2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("AT2").Column + 1).End(xlUp)).Value
        .ColumnWidths = "-1;0"
    End With
    
    '�����^�J�[�A���K�ԁA�I�[�v���|���V�[���������񊈐��i2018/3 ���ذĖ��וt�@�\�ǉ��j
    If FleetTypeFlg = 2 Then
        chkRentacar.Enabled = False
        chkLearncar.Enabled = False
        chkOpenpol.Enabled = False
    End If
    
    
    '���ד��͉�ʂɏ�񂪂���ꍇ���f����
    Dim wsMeisai As Worksheet
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(1, wsMeisai)       '�V�[�g�I�u�W�F�N�g(���ד���)
    Else
        Call subSetSheet(17, wsMeisai)      '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
    End If
    
    Dim strParams() As String
    Dim varParam As Variant
    
    varParam = wsMeisai.Cells(Val(Left(strTxtOther, InStr(strTxtOther, ":") - 1)), Val(Mid(strTxtOther, InStr(strTxtOther, ":") + 1, Len(strTxtOther) - InStr(strTxtOther, ":"))))
    
    If varParam <> "" Then
        strParams() = Split(varParam, "�^")
        
        For Each varParam In strParams()
            If varParam = "����" Then chkOkinawa = True
            If varParam = "�����^�J�[" Then chkRentacar = True
            If varParam = "���K��" Then chkLearncar = True
            If varParam = "�u�[���ΏۊO" Then chkBoom = True
            If varParam = "���[�X�J�[�I�[�v���|���V�[" Then chkLeasecar = True
            If varParam = "�I�[�v���|���V�[��������" Then chkOpenpol = True
            If varParam = "" Then chkOkinawa = True
            If varParam = "���L" Then cmbKouyukbn = "���L"
            If varParam = "�����L" Then cmbKouyukbn = "�����L"
            If varParam = "����敪" Then chkSpecialkbn = True
        Next
    End If
    
    Set wsCode = Nothing
    Set wsMeisai = Nothing
    
    On Error GoTo 0
    
    Exit Sub
    
Error:
    MsgBox "UserForm_Initialize" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["

End Sub


'�u�Z�b�g�v�{�^������
Private Sub btnSet_Click()

    Dim intMsgBox As Integer
    
    On Error GoTo Error

    intMsgBox = MsgBox("���͂������e�𖾍ד��͉�ʂɔ��f���܂��B" & vbCrLf & "��낵���ł����H", vbYesNo, "�m�F�_�C�A���O")
    
    If intMsgBox = 6 Then
    
        '�I�����擾
        Dim varParam As String
        varParam = varParam & IIf(chkOkinawa, "����^", "")
        varParam = varParam & IIf(chkRentacar, "�����^�J�[�^", "")
        varParam = varParam & IIf(chkLearncar, "���K�ԁ^", "")
        varParam = varParam & IIf(chkBoom, "�u�[���ΏۊO�^", "")
        varParam = varParam & IIf(chkLeasecar, "���[�X�J�[�I�[�v���|���V�[�^", "")
        varParam = varParam & IIf(chkOpenpol, "�I�[�v���|���V�[���������^", "")
        varParam = varParam & IIf(cmbKouyukbn <> "", cmbKouyukbn & "�^", "")
        varParam = varParam & IIf(chkSpecialkbn, "����敪�^", "")
        If varParam <> "" Then
            varParam = Left(varParam, Len(varParam) - 1)
        End If
        
        '���ד��͉�ʂɐݒ�
        Dim wsMeisai As Worksheet
        '2018/3 ���ذĖ��וt�@�\�ǉ�
        If FleetTypeFlg = 1 Then  '�t���[�g
            Call subSetSheet(1, wsMeisai)       '�V�[�g�I�u�W�F�N�g(���ד���)
        Else
            Call subSetSheet(17, wsMeisai)      '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
        End If
        
        Call subMeisaiUnProtect     '�V�[�g�̕ی�̉���
        
        wsMeisai.Cells(Val(Left(strTxtOther, InStr(strTxtOther, ":") - 1)), Val(Mid(strTxtOther, InStr(strTxtOther, ":") + 1, Len(strTxtOther) - InStr(strTxtOther, ":")))) = varParam
        
        Set wsMeisai = Nothing
        
        Call subMeisaiProtect       '�V�[�g�̕ی�
        
        '�V�[�g�E�u�b�N�̕\��
        Call subBookUnProtect '�u�b�N�̕ی������ 2018/3 ���ذĖ��וt�@�\�ǉ�
        Call subSheetVisible(True)
        Call subBookProtect   '�u�b�N�̕ی� 2018/3 ���ذĖ��וt�@�\�ǉ�
        
        '����
        Unload Me
    End If

    On Error GoTo 0
    
    Exit Sub
    
Error:
    MsgBox "btnSet_Click" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["


End Sub


'�u�߂�v�{�^������
Private Sub BtnBack_Click()
    
    Dim intMsgBox As Integer
    
    On Error GoTo Error

    intMsgBox = MsgBox("���͓��e�𔽉f�����ɖ��ד��͉�ʂɑJ�ڂ��܂��B" & vbCrLf & "��낵���ł����H", vbYesNo, "�m�F�_�C�A���O")
    
    If intMsgBox = 6 Then
        '�V�[�g�E�u�b�N�̕\��
        Call subBookUnProtect '�u�b�N�̕ی������ 2018/3 ���ذĖ��וt�@�\�ǉ�
        Call subSheetVisible(True)
        Call subBookProtect   '�u�b�N�̕ی� 2018/3 ���ذĖ��וt�@�\�ǉ�
        
        Unload Me
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
