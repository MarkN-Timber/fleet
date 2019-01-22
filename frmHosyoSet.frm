VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHosyoSet 
   Caption         =   "�⏞���e�Z�b�g�i�ꊇ�j"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14910
   OleObjectBlob   =   "frmHosyoSet.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmHosyoSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intConfirmMsg As Integer


Private Sub Frame89_Click()

End Sub

'�����\��
Private Sub UserForm_Initialize()
    Dim strSaveContent  As String

    strSaveContent = ""
    
    On Error GoTo Error
    
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    '�t�H�[���̃^�C�g���ݒ�
    If FleetTypeFlg = 1 Then
        frmHosyoSet.Caption = frmHosyoSet.Caption & "�i�t���[�g�_��j"
    Else
        frmHosyoSet.Caption = frmHosyoSet.Caption & "�i�m���t���[�g���וt�_��j"
    End If
    
    '�R�[�h�l�V�[�g
    Dim wsCode As Worksheet
    If FleetTypeFlg = 1 Then
        Call subSetSheet(4, wsCode)         '�V�[�g�I�u�W�F�N�g(�ʎ��@�R�[�h�l)
    Else
        Call subSetSheet(16, wsCode)        '�V�[�g�I�u�W�F�N�g(�ʎ��@�R�[�h�l�i�m���t���[�g�j)�i2018/3 ���ذĖ��וt�@�\�ǉ��j
    End If
        
    '�ԗ��ی��̎��
    With CmbHknSyurui
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("BJ2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("BJ2").Column + 1).End(xlUp)).Value
        .ColumnWidths = "-1;0"
    End With
    
    '�ԗ��ƐӋ��z
    With CmbSyaryoMskGaku
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("BN2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("BN2").Column + 1).End(xlUp)).Value
        .ColumnWidths = "-1;0"
    End With
        
    '�ΐl����
    With CmbTaijinBaisyo
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("CD2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("CD2").Column).End(xlUp)).Value
    End With
        
    '�Ε�����
    With CmbTaibutsuBaisyo
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("CH2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("CH2").Column).End(xlUp)).Value
    End With
    
    '�Ε��ƐӋ��z
    With CmbTaibutsuMskGaku
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("BR2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("BR2").Column + 1).End(xlUp)).Value
        .ColumnWidths = "-1;0"
    End With
    
    '�l�g���Q(1��)
    With CmbZinshinSyougai_1Mei
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("CL2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("CL2").Column).End(xlUp)).Value
    End With
    
'    '����ҏ��Q(1��)
'    With CmbTouzyouSyougai_1Mei
'        .AddItem ""
'        .List = wsCode.Range(wsCode.Range("CP2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("CP2").Column).End(xlUp)).Value
'    End With
    
    '��ԓ��Z�b�g����
    With CmbDaisyaToku
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("CA2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("CA2").Column).End(xlUp)).Value
'        .ColumnWidths = "0;-1"
    End With
    
    '�t�@�~���[�o�C�N����i2018/3 ���ذĖ��וt�@�\�ǉ��j
    With CmbFamiryBikeToku
        .AddItem ""
        .List = wsCode.Range(wsCode.Range("BV2"), wsCode.Cells(wsCode.Rows.Count, wsCode.Range("BV2").Column + 1).End(xlUp)).Value
        .ColumnWidths = "-1;0"
    End With
    
    
    '��ʉ����̃t���[��
    If FleetTypeFlg = 1 Then
        '�t���[�g
        FrameNonFleet1.Visible = False
        FrameNonFleet2.Visible = False
    Else
        '�m���t���[�g�i2018/3 ���ذĖ��וt�@�\�ǉ��j
        FrameFleet.Visible = False
    End If
    
    
    '�ۑ���񂪂���ꍇ���f����
    If fncFormRef(2, strSaveContent) Then

    Else
        Call subItemGet(strSaveContent)
    End If
         
    Set wsCode = Nothing
    
    On Error GoTo 0
    
    Exit Sub
    
Error:
    MsgBox "UserForm_Initialize" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["

End Sub

'�ۑ�������ʂɔ��f
Private Sub subItemGet(ByVal strAllSetItem As String)
    Dim varSaveContent As Variant
    
    varSaveContent = Split(strAllSetItem, "/")

    '���������ڂɃZ�b�g
    CmbHknSyurui.Value = varSaveContent(0)                      '�ی��̎��
    CmbSyaryoMskGaku.Value = varSaveContent(1)                  '�ԗ��ƐӋ��z
    ChkHknZnsnToku.Value = CBool(varSaveContent(2))             '�ی��S���Ք����
    ChkSyaryoTyoukaToku.Value = CBool(varSaveContent(3))        '�ԗ����ߏC����p����
    ChkSyaryoTonanToku.Value = CBool(varSaveContent(4))         '�ԗ�����ΏۊO����
    CmbTaijinBaisyo.Value = varSaveContent(5)                   '�ΐl����
    CmbTaibutsuBaisyo.Value = varSaveContent(6)                 '�Ε�����
    CmbTaibutsuMskGaku.Value = varSaveContent(7)                '�Ε��ƐӋ��z
    ChkTaibutsuTyoukaToku.Value = CBool(varSaveContent(8))      '�Ε����ߏC����p����
    CmbZinshinSyougai_1Mei.Value = varSaveContent(9)            '�l�g���Q(1��)
    TxtZinshinSyougai_1Jiko.Value = varSaveContent(10)          '�l�g���Q(1����)
    TxtTouzyouSyougai_1Mei.Value = varSaveContent(11)           '����ҏ��Q(1��)
    ChkTouzyouSyougai_1Mei_Taisyougai.Value = _
                                  CBool(varSaveContent(12))     '����ҏ��Q(1��)�ΏۊO
    TxtTouzyouSyougai_1Jiko.Value = varSaveContent(13)          '����ҏ��Q(1����)
    ChkNissuToku.Value = CBool(varSaveContent(14))              '����������
    ChkJigyouNushiToku.Value = CBool(varSaveContent(15))        '���Ǝ��p����
    ChkBengoshiToku.Value = CBool(varSaveContent(16))           '�ٌ�m��p����
    ChkJisonJikoToku.Value = CBool(varSaveContent(17))          '�������̏��Q����
    ChkMuhokenToku.Value = CBool(varSaveContent(18))            '���ی��Ԏ��̏��Q����
    CmbDaisyaToku.Value = varSaveContent(19)                    '���̑�ԁE�g�̉��i�⏞����
    ChkSyohiyou_Futekiyou.Value = CBool(varSaveContent(20))     '�ԗ��������s�K�p����
    ChkJugyouinToku.Value = CBool(varSaveContent(21))           '�]�ƈ����������i�t���[�g�̂݁j
    CmbFamiryBikeToku.Value = varSaveContent(22)                '�t�@�~���[�o�C�N����i2018/3 ���ذĖ��וt�@�\�ǉ��j
    CheckKojinBaisekiToku.Value = CBool(varSaveContent(23))     '�l�����ӔC�⏞����i2018/3 ���ذĖ��וt�@�\�ǉ��j
    ChkCarJikoToku.Value = CBool(varSaveContent(24))            '�����Ԏ��̓���@�@�@�i2018/3 ���ذĖ��וt�@�\�ǉ��j

End Sub


'�u�ꊇ�Z�b�g�v�{�^������
Private Sub BtnHosyouSet_Click()
    Dim strErrContent As String
    Dim strAllSetItem As String

    strErrContent = ""
    strAllSetItem = ""
    
    On Error GoTo Error

    '���̓`�F�b�N
    Call fncEntryErrCheckHosyoSet(strErrContent)

    '�G���[���e�\��
    TxtErrBox = strErrContent & "�`�F�b�N����" & " [ " & Format(Time, "HH:MM:SS") & " ]"
    '�G���[�e�L�X�g�{�b�N�X�̃X�N���[���o�[�ړ�
    TxtErrBox.SetFocus
    TxtErrBox.SelStart = 0

    '�G���[������ꍇ�͏����𔲂���
    If Not strErrContent = "" Then
        Exit Sub
    End If
    
    BtnHosyouSet.SetFocus

    Dim intConfirmMsg As Integer
    intConfirmMsg = MsgBox("���͓��e�𖾍ד��͉�ʂɔ��f���܂��B" & vbCrLf & "��낵���ł���?", vbYesNo, "�m�F�_�C�A���O")
    If intConfirmMsg = 7 Then
    Else
        Call subItemSet(strAllSetItem)
        
        Call subMeisaiUnProtect     '�V�[�g�̕ی�̉���

        If fncFormSave(2, strAllSetItem) Then
            '�V�[�g�E�u�b�N�̕\��
            Call subBookUnProtect '�u�b�N�̕ی������ 2018/3 ���ذĖ��וt�@�\�ǉ�
            Call subSheetVisible(True)
            Call subBookProtect   '�u�b�N�̕ی� 2018/3 ���ذĖ��וt�@�\�ǉ�
            
            Call subAllSet(strAllSetItem)
            
            Call subMeisaiProtect       '�V�[�g�̕ی�
            Unload Me
        End If
    End If
    
    
    On Error GoTo 0
    
    Exit Sub
    
Error:
    MsgBox "BtnHosyouSet_Click" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["

End Sub

Private Function fncEntryErrCheckHosyoSet(ByRef strErrContent As String)
'�֐����FEntryErrCheck
'���e�@�F�G���[�`�F�b�N���s���G���[������ꍇ��,�G���[���e��Ԃ��B
'�����@�FstrErrContent = :�G���[���e

    Dim strErrCheck As String
    Dim blnErrFlg As Boolean
    Dim strFormname As UserForm
    Set strFormname = frmHosyoSet

    strErrCheck = ""
    strErrContent = ""
    blnErrFlg = False

    With strFormname

        ''���ԗ��ی��G���[
        strErrCheck = fncHknSyuruiCheck(.CmbHknSyurui.Value, .CmbSyaryoMskGaku.Value, .CmbDaisyaToku.Value, _
                                        CStr(.ChkHknZnsnToku.Value), CStr(ChkSyaryoTonanToku.Value), CStr(ChkSyaryoTyoukaToku.Value))
        If strErrCheck = "" Then
        Else
            strErrContent = strErrContent & "���ԗ��ی��G���[" & vbCrLf
            strErrContent = strErrContent & strErrCheck & vbCrLf & vbCrLf
        End If

        ''�ΐl����
        '�K�{�`�F�b�N
        strErrCheck = fncNeedCheck(.CmbTaijinBaisyo.Value)
        If strErrCheck = "" Then
        Else
            strErrContent = strErrContent & "�E�ΐl����" & vbCrLf & " " & strErrCheck & vbCrLf & vbCrLf
        End If

        ''�Ε�����
        '�K�{�`�F�b�N
        strErrCheck = fncNeedCheck(.CmbTaibutsuBaisyo.Value)
        If strErrCheck = "" Then
            ''���Ε������G���[
            strErrCheck = fncTaibutsuBaisyo(.CmbTaibutsuBaisyo.Value, .CmbTaibutsuMskGaku.Value, CStr(.ChkTaibutsuTyoukaToku.Value))
            If strErrCheck = "" Then
            Else
                strErrContent = strErrContent & "���Ε������G���[" & vbCrLf
                strErrContent = strErrContent & strErrCheck & vbCrLf & vbCrLf
            End If
        Else
            strErrContent = strErrContent & "�E�Ε�����" & vbCrLf & " " & strErrCheck & vbCrLf & vbCrLf
        End If

        ''�l�g���Q(1��)
        '�K�{�`�F�b�N
        strErrCheck = fncNeedCheck(.CmbZinshinSyougai_1Mei.Value)
        If strErrCheck = "" Then
            ''���l�g���Q�G���[
            strErrCheck = fncZinshinSyougai(.CmbZinshinSyougai_1Mei.Value, .TxtZinshinSyougai_1Jiko.Value)
            If strErrCheck = "" Then
            Else
                strErrContent = strErrContent & "���l�g���Q�G���[" & vbCrLf
                strErrContent = strErrContent & strErrCheck & vbCrLf & vbCrLf
                blnErrFlg = True
            End If
        Else
            strErrContent = strErrContent & "�E�l�g���Q(1��)" & vbCrLf & " " & strErrCheck & vbCrLf & vbCrLf
        End If
        
        If blnErrFlg = False Then
            ''�l�g���Q(1����)
            If Trim(.TxtZinshinSyougai_1Jiko.Value) = "" Then
            Else
                '�����`�F�b�N
                strErrCheck = fncNumCheck(Trim(.TxtZinshinSyougai_1Jiko.Value))
                If Not strErrCheck = "" Then
                    strErrContent = strErrContent & "�E�l�g���Q(1����)" & vbCrLf & " " & strErrCheck & vbCrLf & vbCrLf
                Else
                    '�����`�F�b�N
                    strErrCheck = fncKetaCheck(Trim(.TxtZinshinSyougai_1Jiko.Value), 6, ">")
                    If Not strErrCheck = "" Then
                        strErrContent = strErrContent & "�E�l�g���Q(1����)" & vbCrLf & " " & strErrCheck & vbCrLf & vbCrLf
                    End If
                End If
            End If
        End If

        blnErrFlg = False

        ''���ΐl�����G���[
        strErrCheck = fncTaijinBaisyoCheck(.CmbTaijinBaisyo.Value, CStr(ChkJisonJikoToku.Value), _
                                            .CmbZinshinSyougai_1Mei.Value, CStr(.ChkMuhokenToku.Value))
        If Not strErrCheck = "" Then
            strErrContent = strErrContent & "���ΐl�����G���[" & vbCrLf
            strErrContent = strErrContent & strErrCheck & vbCrLf & vbCrLf
        End If

        ''����ҏ��Q(1��)
        '�K�{�`�F�b�N
'        strErrCheck = fncNeedCheck(.TxtTouzyouSyougai_1Mei.Value)
'        If strErrCheck = "" Then
            ''������ҏ��Q�G���[
        strErrCheck = fncTouzyouSyougai(.TxtTouzyouSyougai_1Mei.Value, Trim(.ChkTouzyouSyougai_1Mei_Taisyougai.Value), _
                                        Trim(.TxtTouzyouSyougai_1Jiko.Value), CStr(.ChkNissuToku.Value))
        If strErrCheck = "" Then
        Else
            strErrContent = strErrContent & "������ҏ��Q�G���[" & vbCrLf
            strErrContent = strErrContent & strErrCheck & vbCrLf & vbCrLf
            blnErrFlg = True
        End If
'        Else
'            strErrContent = strErrContent & "�E����ҏ��Q(1��)" & vbCrLf & " " & strErrCheck & vbCrLf & vbCrLf
'        End If
        
        If blnErrFlg = False Then
            ''����ҏ��Q(1����)
            If Trim(.TxtTouzyouSyougai_1Jiko.Value) = "" Then
            Else
                '�����`�F�b�N
                strErrCheck = fncNumCheck(Trim(.TxtTouzyouSyougai_1Jiko.Value))
                If strErrCheck = "" Then
                    '�����`�F�b�N
                    strErrCheck = fncKetaCheck(Trim(.TxtTouzyouSyougai_1Jiko.Value), 6, ">")
                    If Not strErrCheck = "" Then
                        strErrContent = strErrContent & "�E����ҏ��Q(1����)" & vbCrLf & " " & strErrCheck & vbCrLf & vbCrLf
                    End If
                Else
                    strErrContent = strErrContent & "�E����ҏ��Q(1����)" & vbCrLf & " " & strErrCheck & vbCrLf & vbCrLf
                End If
            End If
        End If
        
        '2018/3 ���ذĖ��וt�@�\�ǉ�
        '�t�@�~���[�o�C�N����
        If FleetTypeFlg = 2 Then
            ''���t�@�~���[�o�C�N����G���[
            strErrCheck = fncFamilyBikeCheck(CmbFamiryBikeToku.Value, CmbZinshinSyougai_1Mei.Value)
            
            If strErrCheck = "" Then
            Else
                strErrContent = strErrContent & "���t�@�~���[�o�C�N����G���[" & vbCrLf
                strErrContent = strErrContent & strErrCheck & vbCrLf & vbCrLf
            End If
        End If
        
        '2018/3 ���ذĖ��וt�@�\�ǉ�
        '�����Ԏ��̓���
        If FleetTypeFlg = 2 Then
            ''�������Ԏ��̓���G���[
            strErrCheck = fncJidoushaJikoCheck(ChkCarJikoToku.Value, CmbZinshinSyougai_1Mei.Value)
            
            If strErrCheck = "" Then
            Else
                strErrContent = strErrContent & "�������Ԏ��̓���G���[" & vbCrLf
                strErrContent = strErrContent & strErrCheck & vbCrLf & vbCrLf
            End If
        End If
        
        ''�G���[����
        If strErrContent <> "" Then
            strErrContent = Left(strErrContent, Len(strErrContent) - 2)
            strErrContent = strErrContent & String(62, "-") & vbCrLf
        End If
        
    End With
    
End Function


'��ʏ����擾
Private Sub subItemSet(ByRef strAllSetItem As String)
    
    '���ڂ������ɃZ�b�g
    strAllSetItem = strAllSetItem & CmbHknSyurui.Value & "/"                    '�ی��̎��
    strAllSetItem = strAllSetItem & CmbSyaryoMskGaku.Value & "/"                '�ԗ��ƐӋ��z
    strAllSetItem = strAllSetItem & CStr((ChkHknZnsnToku.Value)) & "/"          '�ی��S���Ք����
    strAllSetItem = strAllSetItem & CStr((ChkSyaryoTyoukaToku.Value)) & "/"     '�ԗ����ߏC����p����
    strAllSetItem = strAllSetItem & CStr((ChkSyaryoTonanToku.Value)) & "/"      '�ԗ�����ΏۊO����
    strAllSetItem = strAllSetItem & CmbTaijinBaisyo.Value & "/"                 '�ΐl����
    strAllSetItem = strAllSetItem & CmbTaibutsuBaisyo.Value & "/"               '�Ε�����
    strAllSetItem = strAllSetItem & CmbTaibutsuMskGaku.Value & "/"              '�Ε��ƐӋ��z
    strAllSetItem = strAllSetItem & CStr(ChkTaibutsuTyoukaToku.Value) & "/"     '�Ε����ߏC����p����
    strAllSetItem = strAllSetItem & CmbZinshinSyougai_1Mei.Value & "/"          '�l�g���Q(1��)
    strAllSetItem = strAllSetItem & TxtZinshinSyougai_1Jiko.Value & "/"         '�l�g���Q(1����)
    strAllSetItem = strAllSetItem & TxtTouzyouSyougai_1Mei.Value & "/"          '����ҏ��Q(1��)
    strAllSetItem = strAllSetItem & _
                                  CStr(ChkTouzyouSyougai_1Mei_Taisyougai) & "/" '����ҏ��Q(1��)�ΏۊO
    strAllSetItem = strAllSetItem & TxtTouzyouSyougai_1Jiko.Value & "/"         '����ҏ��Q(1����)
    strAllSetItem = strAllSetItem & CStr(ChkNissuToku.Value) & "/"              '����������
    strAllSetItem = strAllSetItem & CStr(ChkJigyouNushiToku.Value) & "/"        '���Ǝ��p����
    strAllSetItem = strAllSetItem & CStr(ChkBengoshiToku.Value) & "/"           '�ٌ�m��p����
    strAllSetItem = strAllSetItem & CStr(ChkJisonJikoToku.Value) & "/"          '�������̏��Q����
    strAllSetItem = strAllSetItem & CStr(ChkMuhokenToku.Value) & "/"            '���ی��Ԏ��̏��Q����
    strAllSetItem = strAllSetItem & CmbDaisyaToku.Value & "/"                   '��ԓ��Z�b�g����
    strAllSetItem = strAllSetItem & CStr(ChkJugyouinToku.Value) & "/"           '�]�ƈ����������
    strAllSetItem = strAllSetItem & CStr(ChkSyohiyou_Futekiyou.Value) & "/"     '�ԗ��������s�K�p����
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    strAllSetItem = strAllSetItem & CmbFamiryBikeToku.Value & "/"               '�t�@�~���[�o�C�N����
    strAllSetItem = strAllSetItem & CStr(CheckKojinBaisekiToku.Value) & "/"     '�l�����ӔC�⏞����
    strAllSetItem = strAllSetItem & CStr(ChkCarJikoToku.Value) & "/"            '�����Ԏ��̓���
     
End Sub


'���ד��͉�ʂɈꊇ�Z�b�g
Private Sub subAllSet(ByVal strAllSetItem As String)
    Dim varSaveContent As Variant
    Dim i As Integer
    Dim strStartRow As String       '�J�n�s
    Dim strAllCnt As String         '���t�ۑ䐔
    Dim objAll As OLEObject
    Dim strAllCell As String

    i = 0
    
    Dim wsMeisai As Worksheet
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(1, wsMeisai)       '�V�[�g�I�u�W�F�N�g(���ד���)
    Else
        Call subSetSheet(17, wsMeisai)      '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
    End If
    
    '�v�Z�p�V�[�g�̊J�n�s
    strStartRow = 21
    
    Dim objSouhuho As Object
    Set objSouhuho = wsMeisai.OLEObjects("txtSouhuho").Object
    '���t�ۑ䐔�̃Z���ԍ�
    strAllCell = Left(objSouhuho.Value, Len(objSouhuho.Value) - 2)
    
    '�t�H�[���̒l��z��
    varSaveContent = Split(strAllSetItem, "/")

    '�V�[�g�̃Z���Ɋi�[
    If FleetTypeFlg = 1 Then  '�t���[�g
        For i = 0 To Val(strAllCell) - 1
        
            wsMeisai.Cells(strStartRow + i, 25) = varSaveContent(0)                                         '�ی��̎��
            wsMeisai.Cells(strStartRow + i, 27) = varSaveContent(1)                                         '�ԗ��ƐӋ��z
            wsMeisai.Cells(strStartRow + i, 28) = IIf(CBool(varSaveContent(2)) = True, "�K�p����", "")      '�ی��S���Ք����
            wsMeisai.Cells(strStartRow + i, 29) = IIf(CBool(varSaveContent(3)) = True, "�K�p����", "")      '�ԗ����ߏC����p����
            wsMeisai.Cells(strStartRow + i, 30) = IIf(CBool(varSaveContent(4)) = True, "�ΏۊO", "")        '�ԗ�����ΏۊO����
            wsMeisai.Cells(strStartRow + i, 31) = varSaveContent(5)                                         '�ΐl����
            wsMeisai.Cells(strStartRow + i, 32) = IIf(CBool(varSaveContent(17)) = True, "�ΏۊO", "")       '�������̏��Q����
            wsMeisai.Cells(strStartRow + i, 33) = IIf(CBool(varSaveContent(18)) = True, "�ΏۊO", "")       '���ی��Ԏ��̏��Q����
            wsMeisai.Cells(strStartRow + i, 34) = varSaveContent(6)                                         '�Ε�����
            wsMeisai.Cells(strStartRow + i, 35) = varSaveContent(7)                                         '�Ε��ƐӋ��z
            wsMeisai.Cells(strStartRow + i, 36) = IIf(CBool(varSaveContent(8)) = True, "�K�p����", "")      '�Ε����ߏC����p����
            wsMeisai.Cells(strStartRow + i, 37) = varSaveContent(9)                                         '�l�g���Q(1��)
            wsMeisai.Cells(strStartRow + i, 38) = varSaveContent(10)                                        '�l�g���Q(1����)
            wsMeisai.Cells(strStartRow + i, 39) = IIf(CBool(varSaveContent(12)) = _
                                                                        True, "�ΏۊO", varSaveContent(11)) '����ҏ��Q(1��)
            wsMeisai.Cells(strStartRow + i, 40) = varSaveContent(13)                                        '����ҏ��Q(1����)
            wsMeisai.Cells(strStartRow + i, 41) = IIf(CBool(varSaveContent(14)) = True, "�K�p����", "")     '����������
            wsMeisai.Cells(strStartRow + i, 42) = IIf(CBool(varSaveContent(15)) = True, "�K�p����", "")     '���Ǝ��p����
            wsMeisai.Cells(strStartRow + i, 43) = IIf(CBool(varSaveContent(16)) = True, "�K�p����", "")     '�ٌ�m��p����
            wsMeisai.Cells(strStartRow + i, 44) = IIf(CBool(varSaveContent(20)) = True, "����", "")         '�]�ƈ����������
            wsMeisai.Cells(strStartRow + i, 45) = varSaveContent(19)                                        '��ԓ��Z�b�g����
            wsMeisai.Cells(strStartRow + i, 46) = IIf(CBool(varSaveContent(21)) = True, "�s�K�p", "")       '�ԗ��������s�K�p����

        Next
        
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    Else  '�m���t���[�g
         For i = 0 To Val(strAllCell) - 1
        
            wsMeisai.Cells(strStartRow + i, 33) = varSaveContent(0)                                         '�ی��̎��
            wsMeisai.Cells(strStartRow + i, 35) = varSaveContent(1)                                         '�ԗ��ƐӋ��z
            wsMeisai.Cells(strStartRow + i, 36) = IIf(CBool(varSaveContent(2)) = True, "�K�p����", "")      '�ی��S���Ք����
            wsMeisai.Cells(strStartRow + i, 37) = IIf(CBool(varSaveContent(3)) = True, "�K�p����", "")      '�ԗ����ߏC����p����
            wsMeisai.Cells(strStartRow + i, 38) = IIf(CBool(varSaveContent(4)) = True, "�ΏۊO", "")        '�ԗ�����ΏۊO����
            wsMeisai.Cells(strStartRow + i, 39) = varSaveContent(5)                                         '�ΐl����
            wsMeisai.Cells(strStartRow + i, 40) = IIf(CBool(varSaveContent(17)) = True, "�ΏۊO", "")       '�������̏��Q����
            wsMeisai.Cells(strStartRow + i, 41) = IIf(CBool(varSaveContent(18)) = True, "�ΏۊO", "")       '���ی��Ԏ��̏��Q����
            wsMeisai.Cells(strStartRow + i, 42) = varSaveContent(6)                                         '�Ε�����
            wsMeisai.Cells(strStartRow + i, 43) = varSaveContent(7)                                         '�Ε��ƐӋ��z
            wsMeisai.Cells(strStartRow + i, 44) = IIf(CBool(varSaveContent(8)) = True, "�K�p����", "")      '�Ε����ߏC����p����
            wsMeisai.Cells(strStartRow + i, 45) = varSaveContent(9)                                         '�l�g���Q(1��)
            wsMeisai.Cells(strStartRow + i, 46) = varSaveContent(10)                                        '�l�g���Q(1����)
            wsMeisai.Cells(strStartRow + i, 47) = IIf(CBool(varSaveContent(12)) = _
                                                                        True, "�ΏۊO", varSaveContent(11)) '����ҏ��Q(1��)
            wsMeisai.Cells(strStartRow + i, 48) = varSaveContent(13)                                        '����ҏ��Q(1����)
            wsMeisai.Cells(strStartRow + i, 49) = IIf(CBool(varSaveContent(14)) = True, "�K�p����", "")     '����������
            wsMeisai.Cells(strStartRow + i, 50) = IIf(CBool(varSaveContent(15)) = True, "�K�p����", "")     '���Ǝ��p����
            wsMeisai.Cells(strStartRow + i, 51) = IIf(CBool(varSaveContent(16)) = True, "�K�p����", "")     '�ٌ�m��p����
            wsMeisai.Cells(strStartRow + i, 52) = varSaveContent(22)                                        '�t�@�~���[�o�C�N����
            wsMeisai.Cells(strStartRow + i, 53) = IIf(CBool(varSaveContent(23)) = True, "�K�p����", "")     '�l�����ӔC�⏞����
            wsMeisai.Cells(strStartRow + i, 54) = IIf(CBool(varSaveContent(24)) = True, "�K�p����", "")     '�����Ԏ��̓���
            wsMeisai.Cells(strStartRow + i, 55) = varSaveContent(19)                                        '��ԓ��Z�b�g����
            wsMeisai.Cells(strStartRow + i, 56) = IIf(CBool(varSaveContent(21)) = True, "�s�K�p", "")       '�ԗ��������s�K�p����

            
        Next
    
    End If
    
    '���ד��͉�ʂ̃G���[�p���X�g������
    wsMeisai.OLEObjects("txtErrMsg").Object.Value = ""
    wsMeisai.OLEObjects("txtErrMsg").Activate
    wsMeisai.Range("A1").Activate
    
    Set wsMeisai = Nothing
    Set objSouhuho = Nothing
    
End Sub


'�u�߂�v�{�^������
Private Sub BtnBack_Click()

    On Error GoTo Error

    Dim intConfirmMsg As Integer
    intConfirmMsg = MsgBox("���͓��e�𔽉f�����ɖ��ד��͉�ʂɑJ�ڂ��܂��B" & vbCrLf & "��낵���ł���?", vbYesNo, "�m�F�_�C�A���O")
    If intConfirmMsg = 7 Then
    Else
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


