Attribute VB_Name = "modChohyoFunctions"
Option Explicit

Dim intConfirmMsg As Integer

'�G���[�_�C�A���O�o��
Function fncTextEntryErrChk(ByVal intForm As Integer) As Boolean
    Dim i           As Integer
    Dim intTotalCar As Integer
    Dim strTemp     As String
    Dim wsTextK     As Worksheet
    Dim wsTextM     As Worksheet
    
    fncTextEntryErrChk = False
    
    strTemp = ""
    
    Call subSetSheet(6, wsTextK)         '�V�[�g�I�u�W�F�N�g(�e�L�X�g���e(����))
    Call subSetSheet(7, wsTextM)         '�V�[�g�I�u�W�F�N�g(�e�L�X�g���e(����))

    '���t�ۑ䐔
    intTotalCar = Val(wsTextK.Cells(1, 19))

    '���t�`�F�b�N
    If intForm = 2 Then

        strTemp = Format(Left(wsTextK.Cells(1, 6), 6), "####/##")
        
        If strTemp = "" Then
            strTemp = "1900/01"
        End If
        
        '�ی��n�������u�ی��n�����̑����錎�̑O�X��1���`���������v�łȂ��ꍇ�G���[
        If Date >= DateSerial(Year(strTemp), Month(strTemp) - 2, 1) And _
            Date < DateSerial(Year(strTemp), Month(strTemp) + 2, 1) Then
        Else
            intConfirmMsg = MsgBox("��荞�񂾃t�@�C���͈���ΏۊO�ł��B" & vbCrLf & "�ی��n���������m�F���������B", vbOKOnly & vbExclamation, "�G���[�_�C�A���O")
            fncTextEntryErrChk = True
        End If
    End If
    
    '���v�ی����`�F�b�N
    For i = 1 To intTotalCar
        If wsTextM.Cells(i, 32) = "" Then
            intConfirmMsg = MsgBox("��荞�񂾃t�@�C���͈���ΏۊO�ł��B" & vbCrLf & "���Z���������Ă��܂���B", vbOKOnly & vbExclamation, "�G���[�_�C�A���O")
            fncTextEntryErrChk = True
            Exit For
        End If
    Next i
    
    '����t�`�F�b�N�i�e�L�X�g�t�@�C���Ɂu����t�t���O�v������ꍇ�G���[�j
    If intForm = 2 Then
        If wsTextK.Cells(1, 20) = "1" Then
            intConfirmMsg = MsgBox("����t�̌_��ł��B" & vbCrLf & "�ی������Ď��Z���Ă��������B", vbOKOnly & vbExclamation, "�G���[�_�C�A���O")
            fncTextEntryErrChk = True
        End If
    End If
    
    '���t�ۑ䐔�`�F�b�N
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
'        If intTotalCar < 10 Then
'            intConfirmMsg = MsgBox("��荞�񂾃t�@�C���͈���ΏۊO�ł��B" & vbCrLf & "���t�ۑ䐔��10�䖢���͈���ł��܂���B", vbOKOnly & vbExclamation, "�G���[�_�C�A���O")
'            fncTextEntryErrChk = True
'        End If
    Else  '�m���t���[�g
        If intTotalCar > 9 Then
            intConfirmMsg = MsgBox("��荞�񂾃t�@�C���͈���ΏۊO�ł��B" & vbCrLf & "���t�ۑ䐔��10��ȏ�͈���ł��܂���B", vbOKOnly & vbExclamation, "�G���[�_�C�A���O")
            fncTextEntryErrChk = True
        End If
    End If
        
    Set wsTextK = Nothing
    Set wsTextM = Nothing

End Function


'�x���_�C�A���O�o��
Function fncTextEntryWarChk(ByVal intForm As Integer) As Boolean
    Dim i           As Integer
    Dim intTotalCar As Integer
    Dim blnWarFlg_Rng_1 As Boolean
    Dim blnWarFlg_Rng_2 As Boolean
    Dim blnWarFlg_Carno As Boolean
    Dim wsTextK     As Worksheet
    Dim wsTextM     As Worksheet
    
    Dim wsTextMP     As Worksheet
    
    fncTextEntryWarChk = False
    blnWarFlg_Rng_1 = False
    blnWarFlg_Rng_2 = False
    blnWarFlg_Carno = False
        
    Call subSetSheet(6, wsTextK)         '�V�[�g�I�u�W�F�N�g(�e�L�X�g���e(����))
    Call subSetSheet(7, wsTextM)         '�V�[�g�I�u�W�F�N�g(�e�L�X�g���e(����))
    
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then
        Call subSetSheet(18, wsTextMP)
    Else
        Call subSetSheet(19, wsTextMP)
    End If
    
    
    '���t�ۑ䐔
    intTotalCar = Val(wsTextK.Cells(1, 19))
    
    For i = 1 To intTotalCar
        '�g�c�t���O�`�F�b�N
        If wsTextM.Cells(i, 68) <> "" Then
            If wsTextM.Cells(i, 68) = 1 Then
                fncTextEntryWarChk = True
                blnWarFlg_Rng_1 = True
            ElseIf wsTextM.Cells(i, 68) = 2 Then
                fncTextEntryWarChk = True
                blnWarFlg_Rng_2 = True
            End If
        End If
        
        If intForm = 2 Then
            '2018/3 ���ذĖ��וt�@�\�ǉ�
            '�ԑ�ԍ��E�o�^�ԍ��`�F�b�N
            If FleetTypeFlg = 1 Then
                If wsTextMP.Cells(6 + i, 3) = "" Or wsTextMP.Cells(6 + i, 4) = "" Or wsTextMP.Cells(6 + i, 5) = "" Then
                    fncTextEntryWarChk = True
                    blnWarFlg_Carno = True
                End If
            Else
                If wsTextMP.Cells(18 + i, 3) = "" Or wsTextMP.Cells(18 + i, 6) = "" Or wsTextMP.Cells(18 + i, 8) = "" Then
                    fncTextEntryWarChk = True
                    blnWarFlg_Carno = True
                End If
            End If
        End If
    Next i
    
    If blnWarFlg_Rng_1 Then
        intConfirmMsg = MsgBox("�Ј��Ɖ�Ώی_��ł��B" & vbCrLf & "���ӂ��Ď葱����i�߂Ă��������B", vbOKOnly & vbInformation, "�x���_�C�A���O")
    ElseIf blnWarFlg_Rng_2 Then
        intConfirmMsg = MsgBox("�ی������Z�ɂ����Čx��������" & vbCrLf & "����܂����B" & vbCrLf & "�_����e���m�F���Ă��������B", vbOKOnly & vbInformation, "�x���_�C�A���O")
    End If
    
    If blnWarFlg_Carno Then
        intConfirmMsg = MsgBox("�ԑ�ԍ��܂��͓o�^�ԍ�(�J�i�E����)�ɖ����͂̂��̂�����܂��B" & vbCrLf & "�\���������ɕ�L���Ă��������B", vbOKOnly & vbInformation, "�x���_�C�A���O")
    End If
    
    Set wsTextK = Nothing
    Set wsTextM = Nothing

End Function


'�e�L�X�g(�ҏW)�ۑ�
Function fncTextEdit(ByVal intKbn As Integer, ByVal intKoumoku As Integer, ByVal strContent As String, _
                            ByVal intMeisaiCnt As Integer) As Variant
'�e�L�X�g���e��ҏW����
'i = ���ڃC���f�b�N�X
'j = ���׍s
'k = �ҏW��ރC���f�b�N�X

    Dim i               As Integer
    Dim j               As Integer
    Dim h               As Integer
    Dim lngTotalHkn     As Long
    Dim lngFstTotalHkn  As Long
    Dim lngYearTotalHkn As Long
    Dim strSave         As String
    Dim wsTextK         As Worksheet
    Dim wsTextM         As Worksheet
    
    i = intKoumoku
    j = intMeisaiCnt
    h = 1
    
    lngTotalHkn = 0
    lngFstTotalHkn = 0
    lngYearTotalHkn = 0
    
    Call subSetSheet(6, wsTextK)
    Call subSetSheet(7, wsTextM)
    
    If intKbn = 1 Then
        '���ʍ��ڕҏW
        With wsTextK
            Select Case i
                Case 1  '���R�[�h�敪
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
                    
                Case 2  '��t�敪
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = fncFindCode(CStr(.Cells(1, i)), "C")
                    .Cells(2, i) = strSave
                    
                Case 3  '��ی���_�l�@�l�敪
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = fncFindCode(CStr(.Cells(1, i)), "G")
                    .Cells(2, i) = strSave
                    
                Case 4  '�ی����
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = fncFindCode(CStr(.Cells(1, i)), "K")
                    .Cells(2, i) = strSave
                    
                Case 5  '�t���[�g�E�m���t���[�g�敪
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = fncFindCode(CStr(.Cells(1, i)), "O")
                    .Cells(2, i) = strSave
                    
                    '���[�p�ϊ�1
                    '���[�p�ϊ�2
                    '���[�p�ϊ�3
                    strSave = CStr(.Cells(2, i))
                    If strSave = "�t���[�g" Then
                        .Cells(3, i) = "�t���[�g"
'                        .Cells(5, i) = "�t���[�g"
'                        .Cells(5, i) = "���t���[�g"
                        .Cells(5, i) = "�����t�ۑ䐔�F10��ȏ�i�t���[�g�E�S�ԗ��ꊇ�Ȃ��j"
                    ElseIf strSave = "�S�ԗ��ꊇ" Then
                        .Cells(3, i) = "�t���[�g"
                        .Cells(4, i) = "�L"
'                        .Cells(5, i) = "�t���[�g�i�S�ԗ��ꊇ�j"
'                        .Cells(5, i) = "���t���[�g�i�S�ԗ��ꊇ�j"
                        .Cells(5, i) = "�����t�ۑ䐔�F10��ȏ�i�t���[�g�E�S�ԗ��ꊇ����j"
                    ElseIf strSave = "�S�ԗ��A�����Z" Then
                        .Cells(3, i) = "�t���[�g"
                        .Cells(4, i) = "�L"
'                        .Cells(5, i) = "�t���[�g�i�S�ԗ��A�����Z�j"
'                        .Cells(5, i) = "���t���[�g�i�S�ԗ��A�����Z�j"
                        .Cells(5, i) = "�����t�ۑ䐔�F10��ȏ�i�t���[�g�E�S�ԗ��A�����Z�j"
                    ElseIf strSave = "�m���t���[�g" Then
                        .Cells(3, i) = "�m���t���[�g"
'                        .Cells(5, i) = "�m���t���[�g"
'                        .Cells(5, i) = "���m���t���[�g"
                        .Cells(5, i) = "�����t�ۑ䐔�F9��ȉ��i�m���t���[�g�j"
                    Else
                        .Cells(3, i) = strSave
                        .Cells(4, i) = strSave
                        .Cells(5, i) = strSave
                    End If
                    
                Case 6  '�ی��n����
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
                    
                    '���[�p�ϊ�1
                    strSave = fncToWareki(CStr(.Cells(1, i)), 11)
                    If strSave = CStr(.Cells(1, i)) Then
                        .Cells(3, i) = strSave
                    Else
                        strSave = Replace(strSave, Mid(Right(strSave, 6), 1, 3), Val(Mid(Right(strSave, 6), 1, 2)) & Mid(Right(strSave, 6), 3, 1))
                        .Cells(3, i) = Replace(strSave, Mid(Right(strSave, 3), 1, 3), Val(Mid(Right(strSave, 3), 1, 2)) & Mid(Right(strSave, 3), 3, 1))
                    End If
                    
                    '���[�p�ϊ�2
                    strSave = CStr(.Cells(3, i))
                    .Cells(4, i) = CStr(Format(strSave, "eemmdd"))
                    
                Case 7  '�ی��n�������敪
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = fncFindCode(CStr(.Cells(1, i)), "S")
                    .Cells(2, i) = strSave
                    
                Case 8  '�ی��n�������敪
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
    
                Case 9  '�ی��I����
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
                    
                    '���[�p�ϊ�1
                    strSave = fncToWareki(CStr(.Cells(1, i)), 11)
                    If .Cells(1, i) = strSave Then
                        .Cells(3, i) = strSave
                    Else
                        strSave = Replace(strSave, Mid(Right(strSave, 6), 1, 3), Val(Mid(Right(strSave, 6), 1, 2)) & Mid(Right(strSave, 6), 3, 1))
                        .Cells(3, i) = Replace(strSave, Mid(Right(strSave, 3), 1, 3), Val(Mid(Right(strSave, 3), 1, 2)) & Mid(Right(strSave, 3), 3, 1))
                    End If
                    
                    '���[�p�ϊ�2
                    strSave = CStr(.Cells(3, i))
                    .Cells(4, i) = CStr(Format(strSave, "eemmdd"))
                    
                Case 10 '�v�Z���@
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = fncFindCode(CStr(.Cells(1, i)), "W")
                    .Cells(2, i) = strSave
                    
                Case 11 '�ی�����_�N
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
                    
                Case 12 '�ی�����_��
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
                    
                Case 13 '�ی�����_��
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
                    
                Case 14 '�������@
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = fncFindCode(.Cells(1, i), "AY")
                    .Cells(2, i) = strSave
                    
                    '���[�p�ϊ�1
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    'strSave = CStr(.Cells(1, i))
                    '.Cells(3, i) = Left(strSave, 1)
                    Select Case CStr(.Cells(1, i))
                        Case "A  "
                            strSave = "A"
                        Case "B12"
                            strSave = "W"
                        Case "D  "
                            strSave = "M"
                        Case "D06"
                            strSave = "F"
                        Case "D12"
                            strSave = "G"
                        Case "E  "
                            strSave = "M"
                        Case "E06"
                            strSave = "F"
                        Case "E12"
                            strSave = "G"
                        Case "F02"
                            strSave = "H"
                        Case "F04"
                            strSave = "H"
                        Case "F06"
                            strSave = "H"
                        Case "F08"
                            strSave = "H"
                        Case "F10"
                            strSave = "H"
                        Case "F12"
                            strSave = "H"
                        Case "G02"
                            strSave = "Y"
                        Case "G04"
                            strSave = "Y"
                        Case "G06"
                            strSave = "Y"
                        Case "G08"
                            strSave = "Y"
                        Case "G10"
                            strSave = "Y"
                        Case "G12"
                            strSave = "Y"
                        Case Else
                            strSave = ""
                    End Select
                    .Cells(3, i) = strSave
                    
                    '���[�p�ϊ�2
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
'                    strSave = CStr(Val(Right(strSave, 2)))
'                    If strSave = 0 Then
'                        .Cells(4, i) = ""
'                    Else
'                        .Cells(4, i) = strSave
'                    End If
                    strSave = CStr(.Cells(1, i))
                    If strSave Like "A*" Or strSave Like "B*" Or strSave Like "D*" Or strSave Like "E*" Then
                        strSave = ""
                    ElseIf strSave Like "F*" Or strSave Like "G*" Then
                        strSave = CStr(Right(strSave, 2))
                    Else
                        strSave = ""
                    End If
                    
                    .Cells(4, i) = strSave
                    
                Case 15 '�t���[�g�D�Ǌ���
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
                    
                Case 16 '����f������
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
                    
                Case 17 '�t���[�g��������
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
                    
                    strSave = CStr(.Cells(1, i))
                    If strSave = "2 " Then
                        '�R�[�h�l�ϊ�
                        .Cells(2, i) = "�L"
                        '���[�p�ϊ�1
                        .Cells(3, i) = "5��"
                        '���[�p�ϊ�2
                        .Cells(4, i) = "�t���[�g���������i5%�j"
                    Else
                        .Cells(2, i) = strSave
                    End If
                    
                Case 18 '�t���[�g�R�[�h
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
                    
                Case 19 '���t�ۑ䐔
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
                    
                Case 20 '����t�t���O
                    '�e�L�X�g�\��t��
                    .Cells(1, i) = strContent
                    
                Case 21 '�X�֔ԍ�
                    .Cells(1, i) = strContent
                    
                Case 22 '�_��ҏZ���i�J�i�j
                    .Cells(1, i) = strContent
                    
                Case 23 '�_��ҏZ���i�����j
                    .Cells(1, i) = strContent
                    
'                Case 24 '�_��ҏZ���i�����j
'                    .Cells(1, i) = strContent

                Case 24 '�@�l���i�J�i�j
                    .Cells(1, i) = strContent
                    
                Case 25 '�@�l���i�����j
                    .Cells(1, i) = strContent
                    
                Case 26 '��E���E�����i�J�i�j
                    .Cells(1, i) = strContent
                    
                Case 27 '��E���E�����i�����j
                    .Cells(1, i) = strContent
                    
                Case 28 '�A����P�@����E�g��
                    .Cells(1, i) = strContent
                    
                Case 29 '�A����Q�@�Ζ���
                    .Cells(1, i) = strContent
                    
                Case 30 '�A����R�@�e�`�w
                    .Cells(1, i) = strContent
                    
                Case 31 '�c�̖�
                    .Cells(1, i) = strContent
                    
                Case 32 '�c�̃R�[�h
                    .Cells(1, i) = strContent
                    
                Case 33 '�c�̈��Ɋւ������
                    .Cells(1, i) = strContent
                    
                Case 34 '�����R�[�h
                    .Cells(1, i) = strContent
                    
                Case 35 '�Ј��R�[�h
                    .Cells(1, i) = strContent
                    
                Case 36 '���ۃR�[�h
                    .Cells(1, i) = strContent
                    
                Case 37 '�㗝�X�R�[�h
                    .Cells(1, i) = strContent
                    
                Case 38 '�،��ԍ�
                    .Cells(1, i) = strContent

            End Select
        End With
    ElseIf intKbn = 2 Then
        '���׍��ڕҏW
        With wsTextM
            Select Case i
                Case 1 '���R�[�h�敪
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                
                Case 2  '�p�r�Ԏ�
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = fncFindCode(CStr(.Cells(j, i)), "AA")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = CStr(strSave)
                
                Case 3 '�Ԗ�
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                
                Case 4  '�^��
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                Case 5  '�d�l
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                Case 6  '���x�o�^�N��
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '���[�p�ϊ�1
                    strSave = fncToWareki(CStr(.Cells(j, i)) & "25", 8)
                    If strSave = CStr(.Cells(j, i)) & "25" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = CStr(.Cells(j, i))
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Replace(strSave, Mid(Right(strSave, 3), 1, 3), Val(Mid(Right(strSave, 3), 1, 2)) & Mid(Right(strSave, 3), 3, 1))
                    End If
                    
                    '���[�p�ϊ�2
                    strSave = CStr(.Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i))
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    '.Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = CStr(Format(strSave, "emmdd"))
                    .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = CStr(Format(strSave, "eemm"))
                    
                Case 7  '�����E�s����
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '���[�p�ϊ�1(�^��)
                    strSave = fncFindCode(CStr(.Cells(j, 7)), "AE")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, 4) = strSave
                    
                    '���[�p�ϊ�1(�^��)
                    If strSave = "������" And CStr(.Cells(j, 4)) <> "" Then
                       .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, 4) = CStr(.Cells(j, 4)) & "��"
                    ElseIf strSave = "�s����" Then
                       .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, 4) = "�Ҳ"
                    ElseIf strSave = "" Then
                       .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, 4) = CStr(.Cells(j, 4))
                    Else
                       .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, 4) = ""
                    End If
                    
                Case 8  '�r�C��
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    '���[�p�ϊ�1
                    strSave = Format(CStr(.Cells(j, i)), "0.00")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 9  '2.5���b�g�����f�B�[�[�����Ɨp���^��p��
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                Case 10 '��ی���_���N����
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    '���[�p�ϊ�1
                    strSave = fncToWareki(.Cells(j, i), 11)
                    .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                    '���[�p�ϊ�2
                    .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = CStr(Format(strSave, "geemmdd"))
                    
                Case 11 '�m���t���[�g����
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '���[�p�ϊ�1
                    strSave = fncFindCode(CStr(.Cells(j, i)), "AI")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    '���[�p�ϊ�2
                    strSave = CStr(.Cells(j, i))
                    If IsNumeric(strSave) Then
                        strSave = Val(strSave)
                    End If
                    .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = StrConv(strSave, vbWide) & "����"
                    
                Case 12 '���̗L�K�p����
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '���[�p�ϊ�1
                    strSave = fncFindCode(CStr(.Cells(j, i)), "AM")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 13 '�m���t���[�g��������
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '���[�p�ϊ�1
                    strSave = fncFindCode(CStr(.Cells(j, i)), "AQ")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                    '���[�p�ϊ�2
                    strSave = fncFindCode(CStr(.Cells(j, i)), "AQ")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 14 '�c�̊�����
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    '���[�p�ϊ�1
                    strSave = Format(CStr(.Cells(j, i)), "0.00")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 15 '�S�[���h�Ƌ�����
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '���[�p�ϊ�1
                    If strContent = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�L"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strContent
                    End If
                    
                    '���[�p�ϊ�2
                    If strContent = "1" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "�S�[���h�Ƌ�����"
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = strContent
                    End If
                    
                Case 16 '�g�p�ړI
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '���[�p�ϊ�1
                    If strContent = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�L"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strContent
                    End If
                    
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    '���[�p�ϊ�2
                    If strContent = "1" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "����E���W���["
                    ElseIf strContent = "3" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "�ʋ΁E�ʊw�g�p"
                    ElseIf strContent = "4" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "�Ɩ��g�p"
                    End If
                    
                Case 17 '����
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "3 " Then
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "����"
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                        '���[�p�ϊ�2
                        '2018/3 ���ذĖ��וt�@�\�ǉ�
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "���ꗿ��"
                        '���[�p�ϊ�3
                        .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = "�����ꗿ��"
                    Else
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 18 '�����^�J�[
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�����^�J�["
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                        '���[�p�ϊ�2
                        '2018/3 ���ذĖ��וt�@�\�ǉ�
                        '.Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�����^�J�[����"
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�������^�J�[����"
                    Else
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 19 '���K��
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "5 " Then
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "���K��"
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                        '���[�p�ϊ�2
                        '.Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "���K�ԗ���"
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�����K�ԗ���"
                    Else
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 20 '�u�[���ΏۊO
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1 " Then
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�u�[���ΏۊO"
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�u�[���ΏۊO����"
                    Else
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 21 '���[�X�J�[�I�[�v���|���V�[
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = CStr(.Cells(j, i))
                    If strSave = "80" Then
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�t��"
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "���[�X�J�[�I�[�v���|���V�["
                    Else
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 22 '�I�[�v���|���V�[��������
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent

                    '�R�[�h�l�ϊ�
                    strSave = CStr(.Cells(j, i))
                    If strSave = "93" Then
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�t��"
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�I�[�v���|���V�[��������"
                    Else
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 23 '���L�E�����L
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = fncFindCode(CStr(.Cells(j, i)), "AU")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                
                Case 24 '�ԗ������N���X
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                
                Case 25 '�ΐl�����N���X
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                
                Case 26 '�Ε������N���X
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                
                Case 27 '��Q�����N���X
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                
                Case 28 '�V�Ԋ���
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�L"
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�V�Ԋ���"
                    Else
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 29 '�����
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent

                    strSave = CStr(.Cells(j, i))
                    If strSave = "8" Then
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�����"
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "����p�r������"
                    Else
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 30 '�ԗ��������i
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                Case 31 '�ԗ�������i
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                Case 32 '���v�ی���
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '���[�p�ϊ�1
                    strSave = CStr(.Cells(j, i))
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                    
                    '���[�p�ϊ�2
                    If j = Val(wsTextK.Cells(1, 19)) Then
                        For h = 1 To Val(wsTextK.Cells(1, 19))
                            lngTotalHkn = lngTotalHkn + Val(.Cells(h, i))
                        Next h
                        
                        For h = 1 To Val(wsTextK.Cells(1, 19))
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + h, i) = Format(lngTotalHkn, "#,#")
                        Next h
                        
                        '���[�p�ϊ�3
                        strSave = CStr(wsTextK.Cells(1, 14))
                        If strSave Like "F*" Or strSave Like "G*" Then
                            If lngTotalHkn = 0 Or wsTextK.Cells(1, 14) = "" Then
                            Else
                                .Cells(4 * Val(wsTextK.Cells(1, 19)) + 1, i) = _
                                    Application.WorksheetFunction.Round(lngTotalHkn / Val(Right(wsTextK.Cells(1, 14), 2)), -1)
                                .Cells(4 * Val(wsTextK.Cells(1, 19)) + 1, i) = _
                                    Format(.Cells(4 * Val(wsTextK.Cells(1, 19)) + 1, i), "#,#")
                            End If
                        Else
                        End If
                        
                    End If
                    
                Case 33 '����ی���
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '���[�p�ϊ�1
                    strSave = CStr(.Cells(j, i))
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                                        
                    '���[�p�ϊ�2
                    If j = Val(wsTextK.Cells(1, 19)) Then
                        For h = 1 To Val(wsTextK.Cells(1, 19))
                            lngFstTotalHkn = lngFstTotalHkn + Val(.Cells(h, i))
                        Next h
                        
                        For h = 1 To Val(wsTextK.Cells(1, 19))
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + h, i) = Format(lngFstTotalHkn, "#,#")
                        Next h
                    End If
                    
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    '���[�p�ϊ�3
                    If wsTextK.Cells(1, 14).Value = "B12" Then
                        strSave = Format(.Cells(j, i).Value, "#,#")
                    Else
                        strSave = ""
                    End If
                    .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 34 '�N�ԕی���
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '���[�p�ϊ�1
                    strSave = CStr(.Cells(j, i))
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                    
                    '���[�p�ϊ�2
                    If j = Val(wsTextK.Cells(1, 19)) Then
                        For h = 1 To Val(wsTextK.Cells(1, 19))
                            lngYearTotalHkn = lngYearTotalHkn + Val(.Cells(h, i))
                        Next h
                        
                        For h = 1 To Val(wsTextK.Cells(1, 19))
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + h, i) = Format(lngYearTotalHkn, "#,#")
                        Next h
                    End If
                    
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    '���[�p�ϊ�3
                    If wsTextK.Cells(1, 14).Value = "B12" Or wsTextK.Cells(1, 14).Value Like "D*" Or wsTextK.Cells(1, 14).Value Like "E*" Then
                        strSave = Format(.Cells(j, i).Value, "#,#")
                    Else
                        strSave = ""
                    End If
                    .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 35 '�N�����
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = fncFindCode(CStr(.Cells(j, i)), "BC")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                
                Case 36 '����^�]�ґΏۊO
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                
                    '���[�p�ϊ�1
                    If strContent = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�L"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strContent
                    End If
                    
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    '35�N������̒��[�p�ϊ�1�i36����^�]�ґΏۊO��35�N���������ɏ�������邽�߂����ŏ�������j
                    If .Cells(j, 35).Value = "" Then
                        strSave = "�ΏۊO"
                    ElseIf .Cells(j, 35).Value = "5" And .Cells(j, 36).Value = "1" Then
                        strSave = "�R�T�Έȏ����⏞�i����ґΏۊO�j"
                    Else
                        strSave = fncFindCode(CStr(.Cells(j, 35)), "BC")
                    End If
                    .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, 35) = strSave
                    
                Case 37 '�^�]�Ҍ���
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = fncFindCode(CStr(.Cells(j, i)), "BG")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    '���[�p�ϊ�1
                    If wsTextK.Cells(1, 4).Value = "7 " Then
                        strSave = "�ΏۊO"
                    ElseIf .Cells(j, i).Value = "" Then
                        strSave = "�Ȃ�"
                    Else
                        strSave = fncFindCode(CStr(.Cells(j, i)), "BG")
                    End If
                    .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 38 '�^�]�ҏ]�ƈ����������
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�L"
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                    Else
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 39 '�ԗ��ی����
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = fncFindCode(CStr(.Cells(j, i)), "BK")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 40 '�ԗ��ی����z
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '���[�p�ϊ�1
                    strSave = CStr(.Cells(j, i))
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                    
                Case 41 '�ԗ��ƐӋ��z
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = fncFindCode(CStr(.Cells(j, i)), "BO")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                    '���[�p�ϊ�1
                    If strSave = "" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Val(strSave)
                    End If
                    
                    '���[�p�ϊ�2
                    If strSave = "" Then
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    ElseIf Val(strSave) = 0 Then
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�Ȃ�"
                    Else
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = Val(strSave)
                    End If
                    
                Case 42 '��ԓ��Z�b�g
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '���[�p�ϊ�1
                    strSave = CStr(.Cells(j, i))
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    'If strSave = "" Then
                    '    .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                    'Else
'                    If strSave = "14" Then
'                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "5,000�~"
'                    ElseIf strSave = "15" Then
'                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "7,000�~"
'                    ElseIf strSave = "16" Then
'                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "10,000�~"
'                    Else
'                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
'                    End If
                    If strSave = "34" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "5,000�~"
                    ElseIf strSave = "35" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "7,000�~"
                    ElseIf strSave = "36" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "10,000�~"
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If

                    
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    '���[�p�ϊ�2
'                    strSave = CStr(.Cells(j, i))
'                    If strSave = "01" Or strSave = "02" Or strSave = "03" Then
'                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
'                    ElseIf strSave = "11" Or strSave = "14" Or strSave = "17" Then
'                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "5,000�~"
'                    ElseIf strSave = "12" Or strSave = "15" Or strSave = "18" Then
'                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "7,000�~"
'                    ElseIf strSave = "13" Or strSave = "16" Or strSave = "19" Then
'                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "10,000�~"
'                    Else
'                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
'                    End If
                    strSave = CStr(.Cells(j, i))
                    If strSave = "21" Then
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                    ElseIf strSave = "31" Or strSave = "34" Then
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "5,000�~"
                    ElseIf strSave = "32" Or strSave = "35" Then
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "7,000�~"
                    ElseIf strSave = "33" Or strSave = "36" Then
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "10,000�~"
                    Else
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If

                    
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    '���[�p�ϊ�3
'                    strSave = CStr(.Cells(j, i))
'                    If strSave = "01" Or strSave = "02" Or strSave = "11" Or strSave = "12" Or strSave = "13" Or strSave = "17" Or strSave = "18" Or strSave = "19" Then
'                        .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = "�L"
'                    ElseIf strSave = "03" Or strSave = "14" Or strSave = "15" Or strSave = "16" Then
'                        .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
'                    Else
'                        .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
'                    End If
                    
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    '���[�p�ϊ�4
'                    strSave = CStr(.Cells(j, i))
'                    If strSave = "01" Or strSave = "03" Or strSave = "11" Or strSave = "12" Or strSave = "13" Then
'                        .Cells(5 * Val(wsTextK.Cells(1, 19)) + j, i) = "30���~"
'                    ElseIf strSave = "02" Or strSave = "14" Or strSave = "15" Or strSave = "16" Or strSave = "17" Or strSave = "18" Or strSave = "19" Then
'                        .Cells(5 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
'                    Else
'                        .Cells(5 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
'                    End If
                    strSave = CStr(.Cells(j, i))
                    If strSave = "21" Or strSave = "31" Or strSave = "32" Or strSave = "33" Then
                        .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = "30���~"
                    ElseIf strSave = "34" Or strSave = "35" Or strSave = "36" Then
                        .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                    Else
                        .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 43 '�ԗ��S���Ք����
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "2" Then
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�K�p"
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                        '���[�p�ϊ�2
'                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�ԗ��S���Ք����"
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�ԗ��S���Վ���p����"
                    Else
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 44 '�ԗ�����ΏۊO����
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�K�p"
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�ԗ�����ΏۊO����"
                    Else
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 45 '�ԗ����ߏC����p����
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�L"
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�ԗ����ߏC����p����"
                    Else
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 46 '�ΐl������
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "������"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                Case 47 '�ΐl�ΏۊO
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�ΏۊO"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                Case 48 '�ΐl�����ی����z
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent

                    '�R�[�h�l�ϊ�
                    strSave = fncFindCode(CStr(.Cells(j, i)), "CE")
                    If strSave = "�ΏۊO" Or strSave = "������" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "1"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                    '���[�p�ϊ�1
                    strSave = CStr(.Cells(j, i))
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                    
                Case 49 '��������
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�ΏۊO"
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                        '���[�p�ϊ�2
                        If CStr(.Cells(j, 47)) = "1" Then
                            '2018/3 ���ذĖ��וt�@�\�ǉ�
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        Else
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�ΏۊO"
                        End If
                    Else
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '���[�p�ϊ�2
                        If CStr(.Cells(j, 47)) = "1" Then
                            '2018/3 ���ذĖ��וt�@�\�ǉ�
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        Else
                            If strSave = "" Then
                                .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "1,500���~"
                            Else
                                .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                            End If
                        End If
                    End If
                    
                Case 50 '���ی��ԏ��Q
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�ΏۊO"
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                        '���[�p�ϊ�2
                        If CStr(.Cells(j, 47)) = "1" Then
                            '2018/3 ���ذĖ��וt�@�\�ǉ�
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        Else
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�ΏۊO"
                        End If
                    ElseIf strSave = "" Then
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        
                        If CStr(.Cells(j, 47)) = "1" Then
                            '2018/3 ���ذĖ��וt�@�\�ǉ�
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        Else
                            If .Cells(Val(wsTextK.Cells(1, 19)) + j, 46) <> "" Then
                                '���[�p�ϊ�2
                                .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "������"
                            ElseIf .Cells(Val(wsTextK.Cells(1, 19)) + j, 47) <> "" Then
                                '���[�p�ϊ�2
                                .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�ΏۊO"
                            ElseIf .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, 48) <> "" Then
                                '���[�p�ϊ�2
                                .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, 48) & "���~"
                            End If
                        End If
                    Else
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 51 '�Ε�������
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "������"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                Case 52 '�Ε��ΏۊO
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�ΏۊO"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                Case 53 '�Ε������ی����z
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = fncFindCode(CStr(.Cells(j, i)), "CI")
                    If strSave = "�ΏۊO" Or strSave = "������" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "1"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                    '���[�p�ϊ�1
                    strSave = CStr(.Cells(j, i))
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                    
                Case 54 '�Ε��ƐӋ��z
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = fncFindCode(CStr(.Cells(j, i)), "BS")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                    '���[�p�ϊ�1
                    If strSave = "" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Val(strSave)
                    End If
                    
                    '���[�p�ϊ�2
                    If strSave = "" Then
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    ElseIf Val(strSave) = 0 Then
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�Ȃ�"
                    Else
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = Val(strSave)
                    End If
                    
                Case 55 '�Ε����ߏC����p����
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�L"
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�Ε����ߏC����p����"
                    Else
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 56 '�l�g���Q 1��
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = CStr(fncFindCode(.Cells(j, i), "CM"))
                    If strSave = "�ΏۊO" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "1"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                    '���[�p�ϊ�1
                    strSave = .Cells(j, i)
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                    
                Case 57 '�l���ΏۊO
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�ΏۊO"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                Case 58 '�l�g���Q 1����
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '���[�p�ϊ�1
                    strSave = CStr(.Cells(j, i))
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                    
                Case 59 '�����Ԏ��̓���
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent

                    '���[�p�ϊ�1
                    If strContent = "2" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�L"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strContent
                    End If
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    '���[�p�ϊ�2
                    If strContent = "2" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�����Ԏ��̓���"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 60 '���S�E����Q�ی����z�@1��
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = fncFindCode(CStr(.Cells(j, i)), "CQ")
                    If strSave = "�ΏۊO" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "1"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                    '���[�p�ϊ�1
                    strSave = CStr(.Cells(j, i))
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                    
                Case 61 '�����ΏۊO
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�ΏۊO"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                Case 62 '�����Ԏ��̓���
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '���[�p�ϊ�1
                    strSave = CStr(.Cells(j, i))
                    If strSave = "" Then
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Format(strSave, "#,#")
                    End If
                
                Case 63 '����������
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "2" Then
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "����������"
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                        '���[�p�ϊ�2
                        If CStr(.Cells(j, 61)) = "1" Then
                            '2018/3 ���ذĖ��וt�@�\�ǉ�
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        Else
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "����������"
                        End If
                    Else
                        '�R�[�h�l�ϊ�
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        
                        If CStr(.Cells(j, 61)) = "1" Then
                            '2018/3 ���ذĖ��וt�@�\�ǉ�
                            '���[�p�ϊ�2
                            .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        Else
                            If strSave = "" Then
                                '���[�p�ϊ�2
                                .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "���ʁE�Ǐ�ʕ�"
                            Else
                                '���[�p�ϊ�2
                                .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                            End If
                        End If
                    End If
                    
                Case 64 '���Ǝ��p����
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    If CStr(.Cells(j, i)) = "1" Then
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "���Ǝ��p����"
'                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�������Ǝ��p����"
                    Else
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 65 '�ٌ�m��p����
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '���[�p�ϊ�1
                    If CStr(.Cells(j, i)) = "1" Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�L"
                    Else
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = ""
                    End If
                    
                Case 66 '�t�@�~���[�o�C�N����
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '���[�p�ϊ�1
                    '�R�[�h�l�ϊ�
                    strSave = fncFindCode(CStr(.Cells(j, i)), "BW")
                    .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    '���[�p�ϊ�2
                    If CStr(.Cells(j, i)) = "1" Then
                        strSave = "�����E����"
                    ElseIf CStr(.Cells(j, i)) = "2" Then
                        strSave = "�����E�l�g"
                    End If
                    .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 67 '�l�����ӔC�⏞����
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    '���[�p�ϊ�2
                    If CStr(.Cells(j, i)) = "1" Then
                        strSave = "3���~�i�ƐӋ��z�Ȃ��j"
                    Else
                        strSave = ""
                    End If
                    .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    
                Case 68 '�g�c�x���t���O
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    If CStr(.Cells(j, i)) = "1" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�g�c�G���[�L"
                    ElseIf CStr(.Cells(j, i)) = "2" Then
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = "�x���L"
                    Else
                        .Cells(Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    End If
                    
                Case 69 '�o�^�ԍ�
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                Case 70 '�ԑ�ԍ�
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                Case 71 '�Ԍ�������
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent
                    
                    '���[�p�ϊ�1
                    strSave = fncToWareki(CStr(.Cells(j, i)), 11)
                    If .Cells(j, i) = strSave Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    Else
                        strSave = Replace(strSave, Mid(Right(strSave, 6), 1, 3), Val(Mid(Right(strSave, 6), 1, 2)) & Mid(Right(strSave, 6), 3, 1))
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Replace(strSave, Mid(Right(strSave, 3), 1, 3), Val(Mid(Right(strSave, 3), 1, 2)) & Mid(Right(strSave, 3), 3, 1))
                    End If
                    
                   '���[�p�ϊ�2
                     strSave = CStr(.Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i))
                    .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = CStr(Format(strSave, "eemmdd"))
                    
                Case 72 '�o�^�ԍ�(�J�i) '2018/3 ���ذĖ��וt�@�\�ǉ�
                    '�e�L�X�g�\��t��
                    .Cells(j, i) = strContent

                Case 73 'ASV����
                    .Cells(j, i) = strContent
                    
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�`�r�u����"
                        '���[�p�ϊ�3
                        .Cells(4 * Val(wsTextK.Cells(1, 19)) + j, i) = "�L"
                    End If
                    
                Case 74 '�ԗ��������s�K�p����
                    .Cells(j, i) = strContent
                    
                    '�R�[�h�l�ϊ�
                    strSave = CStr(.Cells(j, i))
                    If strSave = "1" Then
                        '���[�p�ϊ�1
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = "��"
                        '���[�p�ϊ�2
                        .Cells(3 * Val(wsTextK.Cells(1, 19)) + j, i) = "�L"
                    End If
                Case 75 '��ی��ҏZ���i�J�i�j
                    .Cells(j, i) = strContent
                Case 76 '��ی��Ҏ����i�J�i�j
                    .Cells(j, i) = strContent
                Case 77 '��ی��Ҏ����i�����j
                    .Cells(j, i) = strContent
                Case 78 '�Ƌ��؂̐F
                    .Cells(j, i) = strContent
                Case 79 '�Ƌ��ؗL������
                    .Cells(j, i) = strContent
                    
                    strSave = fncToWareki(CStr(.Cells(j, i)), 11)
                    If .Cells(j, i) = strSave Then
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = strSave
                    Else
                        strSave = Replace(strSave, Mid(Right(strSave, 6), 1, 3), Val(Mid(Right(strSave, 6), 1, 2)) & Mid(Right(strSave, 6), 3, 1))
                        .Cells(2 * Val(wsTextK.Cells(1, 19)) + j, i) = Replace(strSave, Mid(Right(strSave, 3), 1, 3), Val(Mid(Right(strSave, 3), 1, 2)) & Mid(Right(strSave, 3), 3, 1))
                    End If
                Case 80 '�ԗ��ی��Ҏ����i�J�i�j
                    .Cells(j, i) = strContent
                Case 81 '�ԗ��ی��Ҏ����i�����j
                    .Cells(j, i) = strContent
                Case 82 '���L�����ۂ܂��̓��[�X��
                    .Cells(j, i) = strContent
            End Select
        End With
    End If
    
    Set wsTextK = Nothing
    Set wsTextM = Nothing

End Function

'���[�o��
Public Sub subFormSetting(ByVal intChohyo As Integer, ByVal intFormNo As Integer, ByVal intColNo As Integer, _
                            ByVal strCell As String, ByVal intType As Integer, ByVal strFront As String, _
                            ByVal strBehind As String, ByVal strDate As String, Optional ByVal intEdpKbn As Integer, _
                            Optional ByVal intEdpIndex As Integer, Optional ByVal strEdpName As String, _
                            Optional ByVal strEdpVal As String, Optional ByRef valEdpSet As Variant, _
                            Optional intSame As Integer, Optional ByRef intMeisaiCnt As Integer, _
                            Optional ByVal intPageCnt As Integer, Optional ByVal blnFstflg As Boolean)
                            
    Dim wsTextK      As Worksheet
    Dim wsTextM      As Worksheet
    Dim wsAssistSave As Worksheet
    Dim wsChohyo     As Worksheet
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    Dim strEdpValTmp As String
    Dim wsTextMP     As Worksheet
    Dim strSave      As String
    Dim intEdpIdx    As Integer
    Dim intTotalCar  As Integer
    
    Call subSetSheet(6, wsTextK)         '�V�[�g�I�u�W�F�N�g(�e�L�X�g���e(����))
    Call subSetSheet(7, wsTextM)         '�V�[�g�I�u�W�F�N�g(�e�L�X�g���e(����))
    Call subSetSheet(8, wsAssistSave)    '�V�[�g�I�u�W�F�N�g(�\���������ʓ��e)
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then
        Call subSetSheet(18, wsTextMP)
    Else
        Call subSetSheet(19, wsTextMP)
    End If
    
    Select Case intChohyo
        Case 1
            Call subSetSheet(101, wsChohyo)         '�V�[�g�I�u�W�F�N�g(���Ϗ�WK)
        Case 2
            Call subSetSheet(102, wsChohyo)         '�V�[�g�I�u�W�F�N�g(�ԗ����׏�WK)
        Case 3
            Call subSetSheet(103, wsChohyo)         '�V�[�g�I�u�W�F�N�g(�_��\����1����WK)
        Case 4
            Call subSetSheet(104, wsChohyo)         '�V�[�g�I�u�W�F�N�g(�_��\����2����WK)
        Case 5
            Call subSetSheet(105, wsChohyo)         '�V�[�g�I�u�W�F�N�g(���׏�WK)
        Case 6
            Call subSetSheet(106, wsChohyo)         '�V�[�g�I�u�W�F�N�g(�\�����d�c�oWK)
        Case 7
            Call subSetSheet(107, wsChohyo)         '�V�[�g�I�u�W�F�N�g(���׏��d�c�oWK)
    End Select
    
    '���t�ۑ䐔
    intTotalCar = Val(wsTextK.Cells(1, 19))
    
    intType = IIf(intType = 0, 1, intType)

    '������ Or EDP�z��쐬
    If intSame >= intTotalCar + 1 Or blnFstflg Then
        If strCell = "" Then
        Else
            If Evaluate("ISREF(" & strCell & ")") = False Then
            Else
                If intEdpIndex > 0 Then
                    If intEdpKbn <= 0 Then
                        valEdpSet(0, intEdpIndex - 1) = strEdpName
                        valEdpSet(1, intEdpIndex - 1) = strEdpVal
                        Select Case intFormNo  '�擾���敪
                            Case 1  '1(IF�t�@�C�����ʏ�񃌃R�[�h
                                If CStr(wsTextK.Cells(intType, intColNo)) = "" Then
                                Else
                                    valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                End If
                            Case 2  '2(IF�t�@�C�����׏�񃌃R�[�h)
                                '2018/3 ���ذĖ��וt�@�\�ǉ�
                                'If CStr(wsTextK.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)) = "" Then
                                If CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)) = "" Then
                                Else
                                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                                    Select Case intColNo
                                        Case 32 '���v�ی���
                                            strSave = wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)
                                            If strSave = "" Then
                                            Else
                                                Select Case intType
                                                    Case 5
                                                        '�������@�ɂ��ݒ�
                                                        strSave = wsTextK.Cells(1, 14)
                                                        If strSave Like "F*" Or strSave Like "G*" Then
                                                            valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, Replace(CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), ",", ""), strBehind)
                                                        Else
                                                            valEdpSet(2, intEdpIndex - 1) = ""
                                                        End If
                                                    Case Else
                                                        valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, Replace(CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), ",", ""), strBehind)
                                                End Select
                                            End If
                                            
                                        Case 33  '����ی���
                                            strSave = wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)
                                            If strSave = "" Then
                                            Else
                                                Select Case intType
                                                    Case 4
                                                        '�������@�ɂ��ݒ�
                                                        strSave = wsTextK.Cells(1, 14)
                                                        If strSave = "B12" Then
                                                            valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, Replace(CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), ",", ""), strBehind)
                                                        End If
                                                   Case Else
                                                        valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, Replace(CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), ",", ""), strBehind)
                                                End Select
                                            End If
                                        
                                        Case Else
                                            valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End Select
                                End If
                            Case 4  '4(�\�����⏕���)
                                If CStr(wsAssistSave.Cells(intType, intColNo)) = "" Then
                                Else
                                    valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsAssistSave.Cells(intType, intColNo)), strBehind)
                                End If
                            Case 5  '5(���̑�����)
                                Select Case intColNo
                                    Case 1  '�N��������(����)
                                        If intType = 1 Then
                                            valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, CStr(strDate), strBehind)
                                        ElseIf intType = 3 Then
                                            valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, fncToWareki(CStr(Left(strDate, Len(strDate) - 4)), 11), strBehind)
                                        ElseIf intType = 4 Then
                                            valEdpSet(2, intEdpIndex - 1) = fncUnionString(strFront, CStr(strDate) & " - " & CStr(Format((intPageCnt + 1), "0000")), strBehind)
                                        End If
                                    Case 3  '(��)
                                        valEdpSet(2, intEdpIndex - 1) = strFront & strBehind
                                End Select
                        End Select
                    Else
                        '2018/3 ���ذĖ��וt�@�\�ǉ�
                        strEdpValTmp = valEdpSet(intEdpKbn - 1, 1, intEdpIndex - 1)
                        
                        valEdpSet(intEdpKbn - 1, 0, intEdpIndex - 1) = strEdpName
                        valEdpSet(intEdpKbn - 1, 1, intEdpIndex - 1) = strEdpVal
                        Select Case intFormNo
                            Case 1  '1(IF�t�@�C�����ʏ�񃌃R�[�h
                                If CStr(wsTextK.Cells(intType, intColNo)) = "" Then
                                Else
                                    valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                End If
                            Case 2  '2(IF�t�@�C�����׏�񃌃R�[�h)
                                If CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)) = "" Then
                                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                                    If FleetTypeFlg = 2 Then
                                        Select Case intColNo
                                            Case 46, 47, 48
                                                '46:�ΐl�����i�������j�@47:�ΐl�����i�ΏۊO�j�@48:�ΐl�����i�ی����z�j
                                                valEdpSet(intEdpKbn - 1, 1, intEdpIndex - 1) = strEdpValTmp
                                            Case 51, 52, 53
                                                '51:�Ε������i�������j�@52:�Ε������i�ΏۊO�j�@53:�Ε������i�ی����z�j
                                                valEdpSet(intEdpKbn - 1, 1, intEdpIndex - 1) = strEdpValTmp
                                            Case 56, 57
                                                '56:�l�g���Q�i1���j�i�ی����z�j�@57:�l�g���Q�i1���j�i�ΏۊO�j
                                                valEdpSet(intEdpKbn - 1, 1, intEdpIndex - 1) = strEdpValTmp
                                            Case 60, 61
                                                '60:����ҏ��Q�i1���j�i�ی����z�j�@61:����ҏ��Q�i1���j�i�ΏۊO�j
                                                valEdpSet(intEdpKbn - 1, 1, intEdpIndex - 1) = strEdpValTmp
                                        End Select
                                    End If
                                Else
                                    Select Case intColNo
                                        Case 33  '����ی���
                                            strSave = wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)
                                            If strSave = "" Then
                                            Else
                                                '�������@�ɂ��ݒ�
                                                strSave = wsTextK.Cells(1, 14)
                                                If strSave = "B12" Then
                                                    valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, Replace(CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), ",", ""), strBehind)
                                                End If
                                            End If
                                        Case Else
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End Select
'                                    valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 4  '4(�\�����⏕���)
                                If CStr(wsAssistSave.Cells(intType, intColNo)) = "" Then
                                Else
                                    valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsAssistSave.Cells(intType, intColNo)), strBehind)
                                End If
                            Case 5  '5(���̑�����)
                                Select Case intColNo
                                    Case 1  '�N��������(����)
                                        If intType = 1 Then
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(strDate), strBehind)
                                        ElseIf intType = 3 Then
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, fncToWareki(CStr(Left(strDate, Len(strDate) - 4)), 11), strBehind)
                                        ElseIf intType = 4 Then
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(strDate) & " - " & CStr(Format((intPageCnt + 1), "0000")), strBehind)
                                        End If
                                    Case 3  '��
                                        valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = strFront & strBehind
                                End Select
                                
                            '2018/3 ���ذĖ��וt�@�\�ǉ�
                            Case 6  '6(���׏�������)
                            
                                If FleetTypeFlg = 1 Then
                                
                                    '�t���[�g
                                
                                    Select Case intColNo
                                        Case 6 '�o�^�ԍ��i�����j
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(6 + intSame, 3)), strBehind)
                                        
                                        Case 7 '�o�^�ԍ��i�J�i�j
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(6 + intSame, 4)), strBehind)
                                        
                                        Case 8 '�ԑ�ԍ�
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(6 + intSame, 5)), strBehind)
                                        
                                        Case 9 '�Ԍ�������
                                                valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, Format(CStr(wsTextMP.Cells(6 + intSame, 6)), "eemmdd"), strBehind)
                                    End Select
                                    
                                
                                ElseIf FleetTypeFlg = 2 Then
                                
                                    '�m���t���[�g
                            
                                    Select Case intColNo
                            
                                        Case 1 '��ی��ҏZ���i�J�i�j
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, StrConv(CStr(wsTextMP.Cells(6 + intSame, 3)), vbKatakana + vbNarrow), strBehind)
                                            
                                        Case 2 '��ی��Ҏ����i�J�i�j
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, StrConv(CStr(wsTextMP.Cells(6 + intSame, 8)), vbKatakana + vbNarrow), strBehind)
                                            
                                        Case 3 '��ی��Ҏ����i�����j
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(6 + intSame, 11)), strBehind)
                                        
                                        Case 4 '�Ƌ��؂̐F
                                            If CStr(wsTextMP.Cells(6 + intSame, 24)) = "�O���[����" Then
                                                valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, "1", strBehind)
                                            ElseIf CStr(wsTextMP.Cells(6 + intSame, 24)) = "�u���[" Then
                                                valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, "2", strBehind)
                                            ElseIf CStr(wsTextMP.Cells(6 + intSame, 24)) = "�S�[���h" Then
                                                valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, "3", strBehind)
                                            Else
                                                valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(6 + intSame, 24)), strBehind)
                                            End If
                                        
                                        Case 5 '�Ƌ��ؗL������
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, Format(CStr(wsTextMP.Cells(6 + intSame, 27)), "eemmdd"), strBehind)
                                        
                                        Case 6 '�o�^�ԍ��i�����j
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(18 + intSame, 3)), strBehind)
                                        
                                        Case 7 '�o�^�ԍ��i�J�i�j
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(18 + intSame, 6)), strBehind)
                                        
                                        Case 8 '�ԑ�ԍ�
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(18 + intSame, 8)), strBehind)
                                        
                                        Case 9 '�Ԍ�������
                                                valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, Format(CStr(wsTextMP.Cells(18 + intSame, 9)), "eemmdd"), strBehind)
                                        
                                        Case 10 '�ԗ����L�Ҏ����i�J�i�j
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, StrConv(CStr(wsTextMP.Cells(18 + intSame, 16)), vbKatakana + vbNarrow), strBehind)
                                        
                                        Case 11 '�ԗ����L�Ҏ����i�����j
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(18 + intSame, 24)), strBehind)
                                        
                                        Case 12 '���L�����ۂ܂��̓��[�X��
                                            If CStr(wsTextMP.Cells(18 + intSame, 31)) = "���L������" Then
                                                valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, "1", strBehind)
                                            Else
                                                valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(18 + intSame, 31)), strBehind)
                                            End If
                                        
                                        Case 13 '�،��ԍ�
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 3)), strBehind)
                                        
                                        Case 14 '���הԍ�
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 5)), strBehind)
                                        
                                        Case 15 '�O�_�񓙋�
                                            'valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 6)), strBehind)
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = Replace(fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 6)), strBehind), "����", "")
                                        
                                        Case 16 '�O�_�񎖌̗L�K�p����
                                            'valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 7)), strBehind)
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = Replace(fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 7)), strBehind), "�N", "")
                                        
                                        Case 17 '�O�_��ی����
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 8)), strBehind)
                                        
                                        Case 18 '�R�[�h
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 9)), strBehind)
                                        
                                        Case 19 '�ی��n����
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, Format(CStr(wsTextMP.Cells(30 + intSame, 10)), "eemmdd"), strBehind)
                                        
                                        Case 20 '�ی��I����
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, Format(CStr(wsTextMP.Cells(30 + intSame, 17)), "eemmdd"), strBehind)
                                        
                                        Case 21 '3�����_�E������
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 24)), strBehind)
                                        
                                        Case 22 '1�����_�E������
                                            valEdpSet(intEdpKbn - 1, 2, intEdpIndex - 1) = fncUnionString(strFront, CStr(wsTextMP.Cells(30 + intSame, 27)), strBehind)
                                    End Select
                                    
                            End If
                            
                        End Select
                        
                    End If
                    
                End If
                wsChohyo.Range(strCell) = ""
            End If
        End If
    Else
    '�Z���ɒl�Z�b�g
        If (intChohyo = 6 Or intChohyo = 7) And intEdpIndex > 0 Then  '�\�����d�c�oWK�@or�@���׏��d�c�oWK
        Else
            '�擾���敪
            Select Case intFormNo
                Case 1  '1(IF�t�@�C�����ʏ�񃌃R�[�h)
                    If strCell = "" Then
                    Else
                        wsChohyo.Range(strCell).NumberFormatLocal = "@"
                        Select Case intColNo  '����No
                            Case 5  '5�F�t���[�g�E�m���t���[�g�敪
                                '���[
                                Select Case intChohyo
                                    Case 3  '3�F�_��\����1����WK
                                        If wsTextK.Cells(intType, intColNo) = "" Then
                                        Else
                                            'strCell�F�Z���ԍ��AintType�F�ҏWNo
                                            wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        End If
                                    Case Else
                                        wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                End Select
                            Case 7  '�ی��n�������敪
                                Select Case intChohyo
                                    Case 3  '3�F�_��\����1����WK
                                        If wsTextK.Cells(intType, intColNo) = "" Then
                                        Else
                                            wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        End If
                                    Case Else
                                        wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                End Select
                            Case 15  '�t���[�g�D�Ǌ���
                                Select Case intChohyo
                                    Case 3  '3�F�_��\����1����WK
                                        If wsTextK.Cells(intType, intColNo) = "" Then
                                        Else
                                            'strCell�F�Z���ԍ��AintType�F�ҏWNo
                                            wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        End If
                                    Case 5  '5�F���׏�WK
                                        If wsTextK.Cells(intType, intColNo) = "" Then
                                        Else
                                            wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        End If
                                    Case Else
                                        If wsTextK.Cells(intType, 16) = "" Then
                                            wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        Else
                                        End If
                                End Select
                            Case 16  '����f������
                                Select Case intChohyo
                                    Case 3  '3�F�_��\����1����WK
                                        If wsTextK.Cells(intType, intColNo) = "" Then
                                        Else
                                            wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        End If
                                    Case 5  '5�F���׏�WK
                                        If wsTextK.Cells(intType, intColNo) = "" Then
                                        Else
                                            wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        End If
                                    Case Else
                                        If wsTextK.Cells(intType, 15) = "" Then
                                            wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        Else
                                        End If
                                End Select
                            Case 17  '�t���[�g��������
                                Select Case intChohyo
                                    Case 3  '3�F�_��\����1����WK
                                        If wsTextK.Cells(intType, intColNo) = "" Then
                                        Else
                                            wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        End If
                                    Case 5  '5�F���׏�WK
                                        If wsTextK.Cells(intType, intColNo) = "" Then
                                        Else
                                            wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        End If
                                    Case Else
                                        wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                End Select
                            Case 18  '�t���[�g�R�[�h
                                Select Case intChohyo
                                    Case 3
                                        If wsTextK.Cells(intType, intColNo) = "" Then
                                        Else
                                            wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                        End If
                                    Case Else
                                        wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                                End Select
                        Case Else
                            wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextK.Cells(intType, intColNo)), strBehind)
                        End Select
                    End If
                    
                Case 2  '2(IF�t�@�C�����׏�񃌃R�[�h)
                    If strCell = "" Then
                    Else
                        wsChohyo.Range(strCell).NumberFormatLocal = "@"
                        Select Case intColNo  '����No
                            Case 8  '�r�C��
                                If intChohyo = 5 Then  '5�F���׏�WK
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                        wsChohyo.Range(strCell) = ""
                                    Else
                                        wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            '2018/3 ���ذĖ��וt�@�\�ǉ�
                            Case 11  '�m���t���[�g����
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
'                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                        wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            '2018/3 ���ذĖ��וt�@�\�ǉ�
                            Case 12  '���̗L�K�p����
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            '2018/3 ���ذĖ��וt�@�\�ǉ�
                            Case 13  '�m���t���[�g��������
                                If intChohyo = 3 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                ElseIf intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                                
                            Case 14 '�c�̊�����
                                If intChohyo = 3 Or intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells(intType, intColNo)), strBehind)
                                End If

                            '2018/3 ���ذĖ��וt�@�\�ǉ�
                            Case 15  '�S�[���h�Ƌ�����
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 17  '����
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 18  '�����^�J�[
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 19  '���K��
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 20  '�u�[���ΏۊO
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 21  '���[�X�J�[�I�[�v���|���V�[
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 22  '�I�[�v���|���V�[��������
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 23  '���L�E�����L��
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 28  '�V�Ԋ���
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 32  '���v�ی���
                                strSave = wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)
                                If strSave = "" Then
                                Else
                                    Select Case intType
                                        Case 5
                                            '�������@�ɂ��ݒ�
                                            strSave = wsTextK.Cells(1, 14)
                                            If strSave Like "F*" Or strSave Like "G*" Then
                                                If intChohyo = 6 Then
                                                    wsChohyo.Range(strCell) = fncUnionString(strFront, Replace(CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), ",", ""), strBehind)
                                                Else
                                                   wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                                End If
                                            Else
                                                wsChohyo.Range(strCell) = ""
                                            End If
                                        Case Else
                                            If intChohyo = 6 Then
                                                wsChohyo.Range(strCell) = fncUnionString(strFront, Replace(CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), ",", ""), strBehind)
                                            Else
                                                wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                            End If
                                    End Select
                                End If
                            Case 33  '����ی���
                                strSave = wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)
                                If strSave = "" Then
                                Else
                                    Select Case intType
                                        Case 4
                                            '�������@�ɂ��ݒ�
                                            strSave = wsTextK.Cells(1, 14)
                                            If strSave = "B12" Then
                                                If intChohyo = 6 Then
                                                    wsChohyo.Range(strCell) = fncUnionString(strFront, Replace(CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), ",", ""), strBehind)
                                                Else
                                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                                End If
                                            End If
                                       Case Else
                                        If intChohyo = 6 Then
                                            wsChohyo.Range(strCell) = fncUnionString(strFront, Replace(CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), ",", ""), strBehind)
                                        Else
                                            wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                        End If
                                    End Select
                                End If
                            Case 34  '�N�ԕی���
                                strSave = wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)
                                If strSave = "" Then
                                Else
                                    Select Case intType
                                        Case 4
                                            '�������@�ɂ��ݒ�
                                            strSave = wsTextK.Cells(1, 14)
                                            If strSave Like "A*" Or strSave Like "F*" Or strSave Like "G*" Then
                                                wsChohyo.Range(strCell) = ""
                                            Else
                                                wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                            End If
                                        Case Else
                                            wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End Select
                                End If
                            
                            Case 43  '�ԗ��S�����Ք����
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                        wsChohyo.Range(strCell) = ""
                                    Else
                                        wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 46  '�ΐl������
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 47)) = "" And _
                                  CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 48)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 47  '�ΐl�ΏۊO
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 46)) = "" And _
                                  CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 48)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 48  '�ΐl�����ی����z
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 46)) = "" And _
                                  CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 47)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 51  '�Ε�������
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 52)) = "" And _
                                  CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 53)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 52  '�Ε��ΏۊO
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 51)) = "" And _
                                  CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 53)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 53  '�Ε������ی����z
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 51)) = "" And _
                                  CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 52)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 54  '�Ε��ƐӋ��z
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                        wsChohyo.Range(strCell) = ""
                                    Else
                                        wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case 56  '�l�g���Q 1��
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 57)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 57  '�l���ΏۊO
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 56)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 58  '�l�g���Q 1����
                                strSave = CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo))
                                If strSave = "" Then
                                    wsChohyo.Range(strCell) = ""
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, strSave, strBehind)
                                End If
                            Case 60  '���S�E����Q�ی����z 1��
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 61)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 61  '�����ΏۊO
                                If CStr(wsTextM.Cells(Val(wsTextK.Cells(1, 19)) + intSame, 60)) = "" Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                Else
                                End If
                            Case 62  '���S�E����Q�ی����z 1����
                                strSave = CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo))
                                If strSave = "" Then
                                    wsChohyo.Range(strCell) = ""
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, strSave, strBehind)
                                End If
                            Case 73  'ASV����
                                If intChohyo = 5 Then
                                    If wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo) = "" Then
                                    Else
                                        wsChohyo.Range(strCell) = CStr(wsChohyo.Range(strCell)) & fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                    End If
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                                End If
                            Case Else
                                wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextM.Cells((intType - 1) * Val(wsTextK.Cells(1, 19)) + intSame, intColNo)), strBehind)
                        End Select
                    End If
                Case 3  '3(���Ϗ��⏕���)
                    If strCell = "" Then
                    Else
                        wsChohyo.Range(strCell).NumberFormatLocal = "@"  '@�F�\���`���w�聨������
                        If wsAssistSave.Cells(intType, intColNo) = "" Then
                            wsChohyo.Range(strCell) = ""
                        Else
                            wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsAssistSave.Cells(intType, intColNo)), strBehind)
                        End If
                    End If
                    
                Case 4  '4(�\�����⏕���)
                    If strCell = "" Then
                    Else
                        wsChohyo.Range(strCell).NumberFormatLocal = "@"  '@�F�\���`���w�聨������
                        If wsAssistSave.Cells(intType, intColNo) = "" Then
                            wsChohyo.Range(strCell) = ""
                        Else
                            wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsAssistSave.Cells(intType, intColNo)), strBehind)
                        End If
                    End If
                    
                Case 5  '5(���̑�����)
                    Select Case intColNo
                        Case 1  '�N��������(����)
                            If strCell = "" Then
                            Else
                                wsChohyo.Range(strCell).NumberFormatLocal = "@"  '@�F�\���`���w�聨������
                                If intType = 1 Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(strDate), strBehind)
                                ElseIf intType = 3 Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, fncToWareki(CStr(Left(strDate, Len(strDate) - 4)), 11), strBehind)
                                ElseIf intType = 4 Then
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(strDate) & " - " & CStr(Format((intPageCnt + 1), "0000")), strBehind)
                                End If
                            End If
                        Case 2  '���הԍ�
                            If strCell = "" Then
                            Else
                                wsChohyo.Range(strCell).NumberFormatLocal = "@"  '@�F�\���`���w�聨������
                                If intSame > intTotalCar Then
                                    wsChohyo.Range(strCell) = ""
                                Else
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(Format(intMeisaiCnt, "0000")), strBehind)
                                End If
                            End If
                            intMeisaiCnt = intMeisaiCnt + 1
                            
                        Case 3  '��
                            If strCell = "" Then
                            Else
                                wsChohyo.Range(strCell).NumberFormatLocal = "@"  '@�F�\���`���w�聨������
                                wsChohyo.Range(strCell) = strFront & "" & strBehind
                            End If
                            
                    End Select
                    
                '2018/3 ���ذĖ��וt�@�\�ǉ�
                Case 6  '6(���׏�������)
                
                    If strCell = "" Then
                    Else
                    
                        If FleetTypeFlg = 1 Then
                        
                            '�t���[�g
                        
                            Select Case intColNo
                                Case 1 '�o�^�ԍ��i�����j
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(6 + intSame, 3)), strBehind)
                                
                                Case 2 '�o�^�ԍ��i�J�i�j
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(6 + intSame, 4)), strBehind)
                                
                                Case 3 '�ԑ�ԍ�
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(6 + intSame, 5)), strBehind)
                                
                                Case 4 '�Ԍ�������
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(6 + intSame, 6)), strBehind)
                            End Select
                        
                        
                        ElseIf FleetTypeFlg = 2 Then
                        
                            '�m���t���[�g
                        
                            Select Case intColNo
                            
                                Case 1 '��ی��ҏZ���i�J�i�j
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(5 + intMeisaiCnt, 3)), strBehind)
                                    
                                Case 2 '��ی��Ҏ����i�J�i�j
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(5 + intMeisaiCnt, 8)), strBehind)
                                    
                                Case 3 '��ی��Ҏ����i�����j
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(5 + intMeisaiCnt, 11)), strBehind)
                                
                                Case 4 '�Ƌ��؂̐F
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(5 + intMeisaiCnt, 24)), strBehind)
                                
                                Case 5 '�Ƌ��ؗL������
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(5 + intMeisaiCnt, 27)), strBehind)
                                
                                Case 6 '�o�^�ԍ��i�����j
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(17 + intMeisaiCnt, 3)), strBehind)
                                
                                Case 7 '�o�^�ԍ��i�J�i�j
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(17 + intMeisaiCnt, 6)), strBehind)
                                
                                Case 8 '�ԑ�ԍ�
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(17 + intMeisaiCnt, 8)), strBehind)
                                
                                Case 9 '�Ԍ�������
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(17 + intMeisaiCnt, 9)), strBehind)
                                
                                Case 10 '�ԗ����L�Ҏ����i�J�i�j
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(17 + intMeisaiCnt, 16)), strBehind)
                                
                                Case 11 '�ԗ����L�Ҏ����i�����j
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(17 + intMeisaiCnt, 24)), strBehind)
                                
                                Case 12 '���L�����ۂ܂��̓��[�X��
                                    If intChohyo = 5 Then
                                        If CStr(wsTextMP.Cells(17 + intMeisaiCnt, 31)) = "���L������" Then
                                            wsChohyo.Range(strCell) = fncUnionString(strFront, "���L�����ۂ܂��̓��[�X��", strBehind)
                                        Else
                                            wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(17 + intMeisaiCnt, 31)), strBehind)
                                        End If
                                    Else
                                        wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(17 + intMeisaiCnt, 31)), strBehind)
                                    End If
                                
                                Case 13 '�،��ԍ�
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 3)), strBehind)
                                
                                Case 14 '���הԍ�
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 5)), strBehind)
                                
                                Case 15 '�O�_�񓙋�
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 6)), strBehind)
                                
                                Case 16 '�O�_�񎖌̗L�K�p����
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 7)), strBehind)
                                
                                Case 17 '�O�_��ی����
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 8)), strBehind)
                                
                                Case 18 '�R�[�h
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 9)), strBehind)
                                
                                Case 19 '�ی��n����
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 10)), strBehind)
                                
                                Case 20 '�ی��I����
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 17)), strBehind)
                                
                                Case 21 '3�����_�E������
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 24)), strBehind)
                                
                                Case 22 '1�����_�E������
                                    wsChohyo.Range(strCell) = fncUnionString(strFront, CStr(wsTextMP.Cells(29 + intMeisaiCnt, 27)), strBehind)
                            
                            End Select
                        End If
                        
                    End If
                    
            End Select
        End If
    End If
    
    Set wsTextK = Nothing
    Set wsTextM = Nothing
    Set wsChohyo = Nothing
    Set wsAssistSave = Nothing
    
End Sub

Public Sub subFormEDPEdit(ByVal intForm As Integer, ByVal varEditEdp As Variant, ByVal intPageCnt As Integer, ByVal strStartRow As String)
    
    Dim i            As Integer
    Dim j            As Integer
    Dim intRowCnt    As Integer
    Dim intStartRow  As Integer
    Dim strSave      As String
    Dim wsChohyo     As Worksheet
    
    Select Case intForm
        Case 1
            Call subSetSheet(106, wsChohyo)         '�V�[�g�I�u�W�F�N�g(�\�����d�c�oWK)
        Case 2
            Call subSetSheet(107, wsChohyo)         '�V�[�g�I�u�W�F�N�g(���׏��d�c�oWK)
    End Select

    intRowCnt = 0
    intStartRow = wsChohyo.Range(strStartRow).Row + (85 * intPageCnt)

    i = intStartRow

    If intForm = 1 Then
        Do Until wsChohyo.Cells(i, 2).MergeArea(1) = ""
            If wsChohyo.Cells(i, 17).MergeArea(1) = "" Then
                wsChohyo.Cells(i, 2).MergeArea(1) = ""
                wsChohyo.Cells(i, 14).MergeArea(1) = ""
            End If

            i = i + 1
        Loop

        For i = 0 To UBound(varEditEdp, 2)
            If IsEmpty(varEditEdp(2, i)) Then
            Else
                If varEditEdp(2, i) = "" Then
                Else
                    wsChohyo.Cells(intStartRow + intRowCnt, 2) = varEditEdp(0, i)
                    wsChohyo.Cells(intStartRow + intRowCnt, 14) = varEditEdp(1, i)
                    wsChohyo.Cells(intStartRow + intRowCnt, 17) = varEditEdp(2, i)

                    intRowCnt = intRowCnt + 1
                End If
            End If

        Next i
    Else

'        Do Until wsChohyo.Cells(i, 2).MergeArea(1) = ""
'            If wsChohyo.Cells(i, 17).MergeArea(1) = "" Then
'                wsChohyo.Cells(i, 2).MergeArea(1) = ""
'                wsChohyo.Cells(i, 14).MergeArea(1) = ""
'            End If
'            i = i + 1
'        Loop
'
'        For j = 0 To UBound(varEditEdp, 3)
'            If IsEmpty(varEditEdp(0, 2, j)) then
'            Else
'                If varEditEdp(0, 2, j) = "" Then
'                Else
'                    wsChohyo.Cells(intStartRow + intRowCnt, 2) = varEditEdp(0, 0, j)
'                    wsChohyo.Cells(intStartRow + intRowCnt, 14) = varEditEdp(0, 1, j)
'                    wsChohyo.Cells(intStartRow + intRowCnt, 17) = varEditEdp(0, 2, j)
'
'                    intRowCnt = intRowCnt + 1
'                End If
'            End If
'
'        Next j
        Do Until wsChohyo.Cells(i, 2).MergeArea(1) = ""
            If wsChohyo.Cells(i, 13).MergeArea(1) = "" Then
                wsChohyo.Cells(i, 2).MergeArea(1) = ""
                wsChohyo.Cells(i, 10).MergeArea(1) = ""
            End If
            i = i + 1
        Loop
        i = intStartRow

        Do Until wsChohyo.Cells(i, 34).MergeArea(1) = ""
            If wsChohyo.Cells(i, 45).MergeArea(1) = "" Then
                wsChohyo.Cells(i, 34).MergeArea(1) = ""
                wsChohyo.Cells(i, 42).MergeArea(1) = ""
            End If
            i = i + 1
        Loop

        For j = 0 To UBound(varEditEdp, 3)
            If IsEmpty(varEditEdp(0, 2, j)) Then
            Else
                If varEditEdp(0, 2, j) = "" Then
                Else
                    If intRowCnt < 66 Then
                        wsChohyo.Cells(intStartRow + intRowCnt, 2) = varEditEdp(0, 0, j)
                        wsChohyo.Cells(intStartRow + intRowCnt, 10) = varEditEdp(0, 1, j)
                        wsChohyo.Cells(intStartRow + intRowCnt, 13) = varEditEdp(0, 2, j)

                        intRowCnt = intRowCnt + 1
                    Else
                        wsChohyo.Cells(intStartRow + intRowCnt - 66, 34) = varEditEdp(0, 0, j)
                        wsChohyo.Cells(intStartRow + intRowCnt - 66, 42) = varEditEdp(0, 1, j)
                        wsChohyo.Cells(intStartRow + intRowCnt - 66, 45) = varEditEdp(0, 2, j)

                        intRowCnt = intRowCnt + 1
                    End If
                End If
            End If

        Next j

    End If

    If intForm = 2 Then
        intRowCnt = 0
        intStartRow = 78 + (85 * intPageCnt)
        i = intStartRow

        Do Until wsChohyo.Cells(i, 2).MergeArea(1) = ""
            If wsChohyo.Cells(i, 10).MergeArea(1) = "" Then
                wsChohyo.Cells(i, 2).MergeArea(1) = ""
                wsChohyo.Cells(i, 7).MergeArea(1) = ""
            End If
            i = i + 1
        Loop

        i = intStartRow

        Do Until wsChohyo.Cells(i, 20).MergeArea(1) = ""
            If wsChohyo.Cells(i, 28).MergeArea(1) = "" Then
                wsChohyo.Cells(i, 20).MergeArea(1) = ""
                wsChohyo.Cells(i, 25).MergeArea(1) = ""
            End If
            i = i + 1
        Loop

        i = intStartRow
        Do Until wsChohyo.Cells(i, 38).MergeArea(1) = ""
            If wsChohyo.Cells(i, 43).MergeArea(1) = "" Then
                wsChohyo.Cells(i, 38).MergeArea(1) = ""
                wsChohyo.Cells(i, 46).MergeArea(1) = ""
            End If
            i = i + 1
        Loop

        i = intStartRow

        For i = 1 To UBound(varEditEdp, 1)
            intRowCnt = 0

            For j = 0 To UBound(varEditEdp, 3)
                If IsEmpty(varEditEdp(i, 2, j)) Then
                Else
                    If varEditEdp(i, 2, j) = "" Then
                    Else
                        Select Case i
                            Case 1
                                wsChohyo.Cells(intStartRow + intRowCnt, 2) = varEditEdp(i, 0, j)
                                wsChohyo.Cells(intStartRow + intRowCnt, 7) = varEditEdp(i, 1, j)
                                wsChohyo.Cells(intStartRow + intRowCnt, 10) = varEditEdp(i, 2, j)
                            Case 2
                                wsChohyo.Cells(intStartRow + intRowCnt, 20) = varEditEdp(i, 0, j)
                                wsChohyo.Cells(intStartRow + intRowCnt, 25) = varEditEdp(i, 1, j)
                                wsChohyo.Cells(intStartRow + intRowCnt, 28) = varEditEdp(i, 2, j)
                            Case 3
                                wsChohyo.Cells(intStartRow + intRowCnt, 38) = varEditEdp(i, 0, j)
                                wsChohyo.Cells(intStartRow + intRowCnt, 43) = varEditEdp(i, 1, j)
                                wsChohyo.Cells(intStartRow + intRowCnt, 46) = varEditEdp(i, 2, j)
                        End Select
                        intRowCnt = intRowCnt + 1
                    End If
                End If

            Next j

        Next i

    End If

    Set wsChohyo = Nothing

End Sub


'2018/3 ���ذĖ��וt�@�\�ǉ�
Private Function fncUnionString(ByVal strFront As String, ByVal strValue As String, ByVal strBehind As String)

    If strValue = "" Then
        fncUnionString = strValue
    Else
        fncUnionString = strFront & strValue & strBehind
    End If

End Function

