Attribute VB_Name = "modMeisaiFunctions"
Option Explicit

''���ד��͉�ʂ̏����Q

Public Sub subMeisaiAdd(ByRef strAddrow As String, ByRef strAddMode As String)
'�֐����FsubMeisaiAdd
'���e�@�F���׍s�ǉ�
'�����@�F
'        strAddrow          = �ǉ��s��
'        strAddMode         = �J�ڌ���ʂ̔���("1" = ���ד��̓V�[�g/ "2" = ���ד��̓V�[�g�ȊO)

    Dim i          As Integer               '���[�v�p�J�E���g
    Dim blnCopyflg As Boolean               '�R�s�[�t���O
    Dim strAllCnt  As String                '���t�ۑ䐔
    Dim strStartAddrow As String            '�ǉ��J�n�s
    Dim strAddRowRange As String            '�ǉ��Z���ԍ�
    Dim strAddCnt      As String            '�ǉ��J�E���g
    Dim strLastCon As String                '�ŏI�s�̃R���g���[���̔ԍ�
    Dim strCopyRow As String                '�R�s�[�s
    Dim strCopyCon As String                '�R�s�[�R���g���[��
    Dim strCopyConValue As String           '�R�s�[�R���g���[���l
    Dim rngChkAll As Range                  '���[�v�p�����W
    
    Dim wsMeisai As Worksheet           '���ד��̓��[�N�V�[�g
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(1, wsMeisai)       '�V�[�g�I�u�W�F�N�g(���ד���)
    Else
        Call subSetSheet(17, wsMeisai)      '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
    End If
    
    Dim objSouhuho As Object
    Set objSouhuho = wsMeisai.OLEObjects("txtSouhuho").Object
    
    strAllCnt = Left(objSouhuho.Value, Len(objSouhuho.Value) - 2) '���t�ۑ䐔�̒l��ݒ�
    strStartAddrow = 20 + Val(strAllCnt)                          '�ǉ�����s�ԍ���ݒ�
    Application.ScreenUpdating = False                            '�`���~
    Application.EnableEvents = False                              '�C�x���g����
    
    '�ǉ��J�E���g�����Z�b�g
    strAddCnt = 0
    
    '���׃`�F�b�N�Ƀ`�F�b�N�������Ă��邩�m�F
    For Each rngChkAll In wsMeisai.Range("A21:" & wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp).Address)
        If InStr(rngChkAll.Value, "True") > 0 Then
            blnCopyflg = True
            Exit For
        End If
    Next rngChkAll
    
    Set rngChkAll = Nothing
    
    '�ŏI�s�̃R���g���[���̔ԍ���ݒ�
    strLastCon = Left(wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp), InStr(wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp), "/") - 1)
    
    '�ǉ�����
    If blnCopyflg = True And strAddMode = "1" Then

        '�R�s�[�t���O�L���A���ד��͉�ʂ���J�ڂ��Ă����ꍇ�A���׍s�̃R�s�[�ǉ�
        Set rngChkAll = Nothing
        For Each rngChkAll In wsMeisai.Range("A21:" & wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp).Address)

            If InStr(rngChkAll.Value, "True") > 0 Then

                '�ǉ�����Z���ԍ���ݒ�
                strAddRowRange = strStartAddrow + 1

                '�s���R�s�[
                strCopyRow = rngChkAll.Row
                '2018/3 ���ذĖ��וt�@�\�ǉ�
                If FleetTypeFlg = 1 Then  '�t���[�g�i�ŏI�͂`�w��j
                    wsMeisai.Range("A" & strCopyRow & ":AX" & strCopyRow).Copy
                    wsMeisai.Range("A" & strStartAddrow + 1 & ":AX" & strStartAddrow + 1).PasteSpecial
                    wsMeisai.Rows(strStartAddrow + 1).EntireRow.RowHeight = 26.25
                Else
                    '�m���t���[�g�i�ŏI�͂a�g��j
                    wsMeisai.Range("A" & strCopyRow & ":BH" & strCopyRow).Copy
                    wsMeisai.Range("A" & strStartAddrow + 1 & ":BH" & strStartAddrow + 1).PasteSpecial
                    wsMeisai.Rows(strStartAddrow + 1).EntireRow.RowHeight = 26.25
                End If

                '����No���X�V
                wsMeisai.Range("B" & strStartAddrow + 1) = strStartAddrow - 19

                '�R�s�[����R���g���[���̔ԍ����擾
                strCopyCon = Left(wsMeisai.Cells(strStartAddrow, 1), InStr(wsMeisai.Cells(strStartAddrow, 1), "/") - 1)

                '���׃`�F�b�N�{�b�N�X���R�s�[
                Call subConAdd("A" & strAddRowRange, "chkMeisai", "", strLastCon, False)
                '���׃`�F�b�N�̃Z���ɃR���g���[���̔ԍ�/�`�F�b�N��Ԃ�ݒ�
                Range("A" & strAddRowRange) = Val(strLastCon) + 1 & "/False"

                '2018/3 ���ذĖ��וt�@�\�ǉ�
                If FleetTypeFlg = 1 Then  '�t���[�g
                    '�I���{�^�����R�s�[�i�v��j
                    Call subConAdd("W" & strAddRowRange, "btnOtherRate", "", strLastCon, True)
                Else
                    '�I���{�^�����R�s�[�i�`�d��j
                    Call subConAdd("AE" & strAddRowRange, "btnOtherRate", "", strLastCon, True)
                End If

                '�ŏI�s�̃R���g���[���̔ԍ����X�V
                strLastCon = Val(strLastCon) + 1
                '�ǉ��J�E���g
                strAddCnt = strAddCnt + 1
                '�ǉ��s������ɂ��炷
                strStartAddrow = strStartAddrow + 1
                
                '�ǉ������s�����ǉ��\��̍s���ɒB�����ꍇ�A���[�v�𔲂���
                If strAddrow = strAddCnt Then
                    Exit For
                End If
                
            End If
            
        Next rngChkAll
        
        Set rngChkAll = Nothing
        
        '�ǉ��s�� > �`�F�b�N���׍s�� �̏ꍇ�A�`�F�b�N����Ă����ԍŌ�̍s��s�����R�s�[
        Do While Val(strAddrow) > Val(strAddCnt)
            
            '�ǉ�����Z���ԍ���ݒ�
            strAddRowRange = strStartAddrow + 1
            
            '�s���R�s�[
            '2018/3 ���ذĖ��וt�@�\�ǉ�
            If FleetTypeFlg = 1 Then  '�t���[�g�i�ŏI�͂`�w��j
                wsMeisai.Range("A" & strCopyRow & ":AX" & strCopyRow).Copy
                wsMeisai.Range("A" & strStartAddrow + 1 & ":AX" & strStartAddrow + 1).PasteSpecial
                wsMeisai.Rows(strStartAddrow + 1).EntireRow.RowHeight = 26.25
            Else
                '�m���t���[�g�i�ŏI�͂a�g��j
                wsMeisai.Range("A" & strCopyRow & ":BH" & strCopyRow).Copy
                wsMeisai.Range("A" & strStartAddrow + 1 & ":BH" & strStartAddrow + 1).PasteSpecial
                wsMeisai.Rows(strStartAddrow + 1).EntireRow.RowHeight = 26.25
            End If

            '����No���X�V
            wsMeisai.Range("B" & strStartAddrow + 1) = strStartAddrow - 19
            
            '���׃`�F�b�N�{�b�N�X���R�s�[
            Call subConAdd("A" & strAddRowRange, "chkMeisai", "", strLastCon, False)
            '���׃`�F�b�N�̃Z���ɃR���g���[���̔ԍ�/�`�F�b�N��Ԃ�ݒ�
            wsMeisai.Range("A" & strAddRowRange) = Val(strLastCon) + 1 & "/False"
            
            '�I���{�^�����R�s�[
            '2018/3 ���ذĖ��וt�@�\�ǉ�
            If FleetTypeFlg = 1 Then  '�t���[�g�i�v��j
                Call subConAdd("W" & strAddRowRange, "btnOtherRate", "", strLastCon, True)
            Else
                '�m���t���[�g�i�`�d��j
                Call subConAdd("AE" & strAddRowRange, "btnOtherRate", "", strLastCon, True)
            End If
            
            '�ŏI�s�̃R���g���[���̔ԍ����X�V
            strLastCon = Val(strLastCon) + 1
            '�ǉ��J�E���g
            strAddCnt = strAddCnt + 1
            '�ǉ��s������ɂ��炷
            strStartAddrow = strStartAddrow + 1
            
        Loop
        
    Else
        
        '�R�s�[�t���O���A���邢�́A���ד��͉�ʈȊO����J�ڂ��Ă����ꍇ�A���׍s�̐V�K�ǉ�
        Do While Val(strAddrow) > Val(strAddCnt)
            
            '�ǉ�����Z���ԍ���ݒ�
            strAddRowRange = strStartAddrow + 1
            
            '2018/3 ���ذĖ��וt�@�\�ǉ�
            If FleetTypeFlg = 1 Then  '�t���[�g�i�ŏI�͂`�w��j
                '�s��V�K�ǉ�
                wsMeisai.Range("A21:AX21").Copy
                '�����s�̏������R�s�[
                wsMeisai.Range("A" & strStartAddrow + 1 & ":AX" & strStartAddrow + 1).PasteSpecial Paste:=xlPasteFormats
                '�����s�̓��͋K�����R�s�[
                wsMeisai.Range("A" & strStartAddrow + 1 & ":AX" & strStartAddrow + 1).PasteSpecial Paste:=xlPasteValidation
                wsMeisai.Rows(strStartAddrow + 1).RowHeight = 26.25
            Else
                '�m���t���[�g�i�ŏI�͂a�g��j
                '�s��V�K�ǉ�
                wsMeisai.Range("A21:BH21").Copy
                '�����s�̏������R�s�[
                wsMeisai.Range("A" & strStartAddrow + 1 & ":BH" & strStartAddrow + 1).PasteSpecial Paste:=xlPasteFormats
                '�����s�̓��͋K�����R�s�[
                wsMeisai.Range("A" & strStartAddrow + 1 & ":BH" & strStartAddrow + 1).PasteSpecial Paste:=xlPasteValidation
                wsMeisai.Rows(strStartAddrow + 1).RowHeight = 26.25

            End If

            '����No���X�V
            wsMeisai.Range("B" & strStartAddrow + 1) = strStartAddrow - 19

            '���׃`�F�b�N�{�b�N�X��V�K�ǉ�
            Call subConAdd("A" & strAddRowRange, "chkMeisai", "", strLastCon, False)
            '���׃`�F�b�N�̃Z���ɃR���g���[���̔ԍ�/�`�F�b�N��Ԃ�ݒ�
            wsMeisai.Range("A" & strAddRowRange) = Val(strLastCon) + 1 & "/False"

            '�I���{�^����V�K�ǉ�
            '2018/3 ���ذĖ��וt�@�\�ǉ�
            If FleetTypeFlg = 1 Then  '�t���[�g�i�v��j
                Call subConAdd("W" & strAddRowRange, "btnOtherRate", "", strLastCon, False)
            Else
                '�m���t���[�g�i�`�d��j
                Call subConAdd("AE" & strAddRowRange, "btnOtherRate", "", strLastCon, False)
            End If

            '�ŏI�s�̃R���g���[���̔ԍ����X�V
            strLastCon = Val(strLastCon) + 1
            '�ǉ��J�E���g
            strAddCnt = strAddCnt + 1
            '�ǉ��s������ɂ��炷
            strStartAddrow = strStartAddrow + 1

        Loop

    End If

    Application.CutCopyMode = False                                  'COPY�I������
    objSouhuho.Value = CStr(Val(strAllCnt) + Val(strAddrow)) + " ��" '���t�ۑ䐔�X�V
    wsMeisai.Range("A1").Select                                      '�t�H�[�J�X��擪�ɐݒ�
    Application.EnableEvents = True                                  '�C�x���g�L��
    Application.ScreenUpdating = True                                '�`��J�n
    Call subCellProtect(Val(strAllCnt) + Val(strAddrow))             '�Z�����͉\�͈͕ύX

    Set wsMeisai = Nothing
    Set objSouhuho = Nothing

End Sub


Public Sub subConAdd(ByRef strAddRowRange As String, strConName As String, ByRef strCopyValue As String, _
                     ByRef strLastCon As String, ByRef blnAddMode As Boolean)
'�֐����FsubConAdd
'���e�@�F�R���g���[���̒ǉ�
'�����@�F
'        strAddRowRange     = �ǉ��Z��
'        strConName         = �ǉ��R���g���[���̖��O
'        strCopyValue       = �R�s�[���̃I�u�W�F�N�g�̓��͓��e(�V�K�쐬�E���͓��e�̑��݂��Ȃ��ꍇ�́A�u�����N)
'        strLastCon         = ���ד��̓V�[�g�ɑ��݂���ŏI�s�̃R���g���[���̔ԍ�
'        blnAddMode         = �V�K�ǉ��E�R�s�[�̔��ʃt���O

    Dim strCellLeft As String        '�R���g���[�����쐬����Z���̍��ʒu
    Dim strCellTop  As String        '�R���g���[�����쐬����Z���̏�ʒu
    Dim wsMeisai    As Worksheet     '���ד��̓��[�N�V�[�g
    
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(1, wsMeisai)                       '�V�[�g�I�u�W�F�N�g(���ד���)
    Else
        Call subSetSheet(17, wsMeisai)                       '�V�[�g�I�u�W�F�N�g(���ד���)
    End If
    strCellLeft = wsMeisai.Range(strAddRowRange).Left   '�R���g���[�����쐬����Z���̍��ʒu��ݒ�
    strCellTop = wsMeisai.Range(strAddRowRange).Top     '�R���g���[�����쐬����Z���̏�ʒu��ݒ�
    
    '�R���g���[����V�K�쐬
    If strConName Like "chkMeisai*" Then
        '���׃`�F�b�N�{�b�N�X
        wsMeisai.CheckBoxes.Add(strCellLeft + 8.25, strCellTop + 5, 24, 16.5).Select
        Selection.OnAction = "subClickchkMeisai"
        Selection.Characters.Text = ""
        
        '�R�s�[�̏ꍇ��`�F�b�N��t����
'        If blnAddMode then
'            Selection.Value = xlOn
'        End If
    ElseIf strConName Like "btnOtherRate*" Then
        '�I���{�^��
        wsMeisai.Buttons.Add(strCellLeft + 7, strCellTop + 4, 35, 18).Select
        Selection.OnAction = "subClickOtherBtn"
        Selection.Characters.Text = "�I��"
        Selection.Font.Size = 8
    End If
    
    '�쐬�����R���g���[���̖��O��ύX
    Selection.Name = strConName & Val(strLastCon) + 1
    
    Set wsMeisai = Nothing
    
End Sub

Public Sub subMeisaiDel(ByRef strDelRowCnt As String, ByRef strDelmode As String)
'�֐����FsubMeisaiDel
'���e�@�F���׍s�̍폜
'�����@�F
'        strDelRowCnt       = �폜�s��
'        strDelmode         = �J�ڌ���ʂ̔���("1" = ���ד��̓V�[�g/ "2" = ���ד��̓V�[�g�ȊO)

    Dim i As Integer                '���[�v�p�J�E���g
    Dim strAllCnt As String         '���t�ۑ䐔
    Dim strDelCnt As String         '�폜����s��
    Dim strDelRow As String         '�폜����s�ԍ�
    Dim strDelCon As String         '�폜����R���g���[���̔ԍ�
    Dim strDelRowArr() As String    '�폜����s�ԍ����i�[����z��
    Dim strDelConArr() As String    '�폜����R���g���[���̔ԍ����i�[����z��
    Dim strLastRow As String        '�ŏI�s�̍s�ԍ�
    Dim strLastCon As String        '�ŏI�s�̃R���g���[���̔ԍ�
    Dim rngChkAll  As Range         '���[�v�p�����W

    Dim wsMeisai As Worksheet           '���ד��̓��[�N�V�[�g
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(1, wsMeisai)       '�V�[�g�I�u�W�F�N�g(���ד���)
    Else
        Call subSetSheet(17, wsMeisai)      '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
    End If

    Dim objSouhuho As Object
    Set objSouhuho = wsMeisai.OLEObjects("txtSouhuho").Object

    strAllCnt = Val(Left(objSouhuho.Value, Len(objSouhuho.Value) - 2))      '���t�ۑ䐔��ݒ�
    Application.ScreenUpdating = False                                      '�`���~
    Application.EnableEvents = False                                        '�C�x���g����


    '�폜�J�E���g�����Z�b�g
    strDelCnt = 0

    '�폜����
    If strDelmode = "1" Then

        '���ד��͉�ʂ���J�ڂ��Ă����ꍇ�A�`�F�b�N�̓����Ă��閾�׍s���폜�Ώۂɐݒ�
        For Each rngChkAll In wsMeisai.Range("A21:" & wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp).Address)

            If InStr(rngChkAll.Value, "True") > 0 Then
                '�폜�s�̏����i�[���Ă���z��̗v�f����ύX
                ReDim Preserve strDelRowArr(strDelCnt) As String
                ReDim Preserve strDelConArr(strDelCnt) As String

                '�`�F�b�N����Ă��閾�׃`�F�b�N�̃R���g���[���̔ԍ���ݒ�
                strDelConArr(strDelCnt) = Left(rngChkAll, InStr(rngChkAll, "/") - 1)
                '�`�F�b�N����Ă��閾�׃`�F�b�N�̍s�ԍ���ݒ�
                strDelRowArr(strDelCnt) = wsMeisai.Shapes("chkMeisai" & strDelConArr(strDelCnt)).TopLeftCell.Row

                '�폜����s���J�E���g
                strDelCnt = strDelCnt + 1
                '���t�ۑ䐔����폜����������
                strAllCnt = strAllCnt - 1

            End If

        Next rngChkAll

        Set rngChkAll = Nothing

    Else

        '���ד��͉�ʈȊO����J�ڂ��Ă����ꍇ�A�ŏI�s����폜�s�������폜�Ώۂɐݒ�
        '�폜�s�̏����i�[���Ă���z��̗v�f����ύX
        ReDim Preserve strDelRowArr(Val(strDelRowCnt) - 1) As String
        ReDim Preserve strDelConArr(Val(strDelRowCnt) - 1) As String

        '�ŏI�s�̍s�ԍ���ݒ�
        strLastCon = Left(wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp), InStr(wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp), "/") - 1)
        '�ŏI�s�̃R���g���[���̔ԍ���ݒ�
        strLastRow = wsMeisai.Shapes("chkMeisai" & strLastCon).TopLeftCell.Row

        For i = 0 To Val(strDelRowCnt - 1)

            '�ŏI�s����폜�s�����̍s�ԍ��ƃR���g���[���̔ԍ���ݒ�
            strDelRow = Val(strLastRow) - i
            strDelCon = Left(wsMeisai.Range("A" & strDelRow), InStr(wsMeisai.Range("A" & strDelRow), "/") - 1)

            strDelConArr(strDelRowCnt - 1 - i) = strDelCon
            strDelRowArr(strDelRowCnt - 1 - i) = strDelRow

            '�폜����s���J�E���g
            strDelCnt = strDelCnt + 1
            '���t�ۑ䐔����폜����������
            strAllCnt = strAllCnt - 1

        Next i

    End If

    strDelCnt = strDelCnt - 1

    '�폜�s����0���ȏ�̏ꍇ�A���׍s���폜
    If Val(strDelCnt) > -1 Then

        For i = LBound(strDelConArr) To UBound(strDelConArr)
            If IsNull(strDelConArr(strDelCnt - i)) = False Then

                If strAllCnt = 0 And i = UBound(strDelConArr) Then

                    '�Ώۂ̖��׍s���V�[�g�ɑ��݂���Ō��1�s�̏ꍇ�A�I����ԁE���͓��e�̃N���A�̂ݍs���s���c��
                    Dim strChkValue As String

                    '�s���N���A
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    If FleetTypeFlg = 1 Then  '�t���[�g�i�ŏI��AX��j
                        wsMeisai.Range("B" & strDelRowArr(strDelCnt - i) & ":" & "AX" & strDelRowArr(strDelCnt - i)).ClearContents
                    Else
                        '�m���t���[�g�i�ŏI��BH��j
                        wsMeisai.Range("B" & strDelRowArr(strDelCnt - i) & ":" & "BH" & strDelRowArr(strDelCnt - i)).ClearContents
                    End If


                    '���׃`�F�b�N�{�b�N�X���N���A
                    wsMeisai.Shapes("chkMeisai" & strDelConArr(strDelCnt - i)).ControlFormat.Value = xlOff
                    strChkValue = wsMeisai.Cells(strDelRowArr(strDelCnt - i), 1)
                    wsMeisai.Cells(strDelRowArr(strDelCnt - i), 1) = Left(strChkValue, InStr(strChkValue, "/") - 1) & "/False"

'                    '���̑������e�L�X�g�{�b�N�X���N���A
'                    wsMeisai.OLEObjects("txtOtherRate" & strDelConArr(strDelCnt - i)).Object.Value = ""

                    strAllCnt = strAllCnt + 1
                Else

                    '���׍s��2�s�ȏ�̂����Ă���ꍇ�A�Ώۂ̖��׍s���폜
                    '���׃`�F�b�N�{�b�N�X���폜
                    wsMeisai.Shapes("chkMeisai" & strDelConArr(strDelCnt - i)).Delete

                    '�I���{�^�����폜
                    wsMeisai.Shapes("btnOtherRate" & strDelConArr(strDelCnt - i)).Delete

'                    '���̑������e�L�X�g�{�b�N�X���폜
'                    wsMeisai.OLEObjects("txtOtherRate" & strDelConArr(strDelCnt - i)).Delete

                    '�s���폜
                    '2018/3 ���ذĖ��וt�@�\�ǉ�
                    If FleetTypeFlg = 1 Then  '�t���[�g�i�ŏI��AX��j
                        wsMeisai.Range("A" & strDelRowArr(strDelCnt - i) & ":AX" & strDelRowArr(strDelCnt - i)).Select
                        Selection.Delete Shift:=xlUp
                    Else
                        '�m���t���[�g�i�ŏI��BH��j
                        wsMeisai.Range("A" & strDelRowArr(strDelCnt - i) & ":BH" & strDelRowArr(strDelCnt - i)).Select
                        Selection.Delete Shift:=xlUp
                    End If

                End If

            End If
        Next i

        '����No���X�V
        Dim strMeisaiNoRow As String
        strMeisaiNoRow = wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp).Row

        For i = 0 To wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp).Row - 20
            '����No���X�V
            wsMeisai.Range("B" & strMeisaiNoRow - i) = (strMeisaiNoRow - i) - 20

            '�Ώۂ̍s�̖���No�ƂЂƂO�̍s�̖���No���A�ԂɂȂ����ꍇ�A���[�v�𔲂���
            If wsMeisai.Range("B" & strMeisaiNoRow - i) - wsMeisai.Range("B" & (strMeisaiNoRow - i) - 1) = 1 Then
                Exit For
            End If
        Next i

    End If

    objSouhuho.Value = strAllCnt + " ��"        '���t�ۑ䐔�X�V
    wsMeisai.Range("A1").Select                 '�t�H�[�J�X��擪�ɐݒ�
    Application.EnableEvents = True             '�C�x���g�L��
    Application.ScreenUpdating = True           '�`���~
    Call subCellProtect(strAllCnt)              '�Z�����͉\�͈͕ύX

    Set wsMeisai = Nothing
    Set objSouhuho = Nothing

End Sub

Public Function fncTempSave_Kyotsu(strAllCell As String) As String
'�֐����FfncTempSave_Kyotsu
'���e�@�F���ʏ�񃌃R�[�h�i1�s�ځj����
'�����@�F
'        strAllCell         = ���t�ۑ䐔

    Dim wsKyoutsuU As Worksheet             '�R�[�h�l���[�N�V�[�g
    Call subSetSheet(2, wsKyoutsuU)         '�V�[�g�I�u�W�F�N�g(�ʎ��@���ʍ���)

    Dim wsMoushikomiP As Worksheet
    Call subSetSheet(8, wsMoushikomiP)      '�V�[�g�I�u�W�F�N�g(�\���������ʓ��e�j

    '�u�ی��I�����v�̌v�Z(�ی��n�������ی����ԁi1�N�j������)
    Dim strSaveDate As String
    strSaveDate = Left(wsKyoutsuU.Range("E2").Value, 4) & "/" & Mid(wsKyoutsuU.Range("E2").Value, 5, 2) & "/" & Mid(wsKyoutsuU.Range("E2").Value, 7, 2)
    strSaveDate = CStr(DateAdd("yyyy", wsKyoutsuU.Range("F2").Value, CDate(strSaveDate)))
    strSaveDate = Format(strSaveDate, "YYYYMMDD")

    '�o�͕����񐶐�
    Dim strSaveCont As String
    strSaveCont = strSaveCont & "1" & ","                           '���R�[�h�敪
    strSaveCont = strSaveCont & wsKyoutsuU.Range("A2").Value & ","  '��t�敪
    strSaveCont = strSaveCont & wsKyoutsuU.Range("B2").Value & ","  '��ی���_�l�@�l�敪
    strSaveCont = strSaveCont & wsKyoutsuU.Range("C2").Value & ","  '�ی����
    strSaveCont = strSaveCont & wsKyoutsuU.Range("D2").Value & ","  '�t���[�g�E�m���t���[�g�敪
    strSaveCont = strSaveCont & wsKyoutsuU.Range("E2").Value & ","  '�ی��n����
    strSaveCont = strSaveCont & "P" & ","                           '�ی��n�������敪
    strSaveCont = strSaveCont & "4" & ","                           '�ی��n������
    strSaveCont = strSaveCont & strSaveDate & ","                   '�ی��I����
    strSaveCont = strSaveCont & wsKyoutsuU.Range("G2").Value & ","  '�v�Z���@
    strSaveCont = strSaveCont & wsKyoutsuU.Range("F2").Value & ","  '�ی�����_�N
    strSaveCont = strSaveCont & "" & ","                            '�ی�����_��
    strSaveCont = strSaveCont & "" & ","                            '�ی�����_��
    strSaveCont = strSaveCont & wsKyoutsuU.Range("H2").Value & ","  '�������@
    strSaveCont = strSaveCont & IIf(wsKyoutsuU.Range("I2").Value = "" _
                    , "", Val(wsKyoutsuU.Range("I2").Value)) & ","  '�t���[�g�D�Ǌ���
    strSaveCont = strSaveCont & IIf(wsKyoutsuU.Range("J2").Value = "", _
                      "", Val(wsKyoutsuU.Range("J2").Value)) & ","  '����f������
    strSaveCont = strSaveCont & wsKyoutsuU.Range("K2").Value & ","  '�t���[�g��������
    strSaveCont = strSaveCont & wsKyoutsuU.Range("L2").Value & ","  '�t���[�g�R�[�h
    strSaveCont = strSaveCont & strAllCell & ","                    '���t�ۑ䐔
    strSaveCont = strSaveCont & "" & ","                              '����t�t���O
    If blnChouhyouflg Then
        strSaveCont = strSaveCont & wsMoushikomiP.Range("B1").Value & _
                             String(3 - Len(wsMoushikomiP.Range("B1").Value), " ") '�X�֔ԍ��i�O�j
        strSaveCont = strSaveCont & wsMoushikomiP.Range("C1").Value & "," '�X�֔ԍ��i��j
        strSaveCont = strSaveCont & StrConv(wsMoushikomiP.Range("D1").Value, vbWide) & "," '�_��ҏZ���i�J�i�j
        strSaveCont = strSaveCont & wsMoushikomiP.Range("E1").Value & _
                             String(40 - Len(wsMoushikomiP.Range("E1").Value), " ") '�_��ҏZ���i�����j
        strSaveCont = strSaveCont & wsMoushikomiP.Range("F1").Value & "," '�_��ҏZ���i�����j
        strSaveCont = strSaveCont & StrConv(wsMoushikomiP.Range("G1").Value, vbWide) & "," '�@�l���i�J�i�j
        strSaveCont = strSaveCont & wsMoushikomiP.Range("H1").Value & "," '�@�l���i�����j
        strSaveCont = strSaveCont & StrConv(wsMoushikomiP.Range("I1").Value, vbWide) & "," '��E���E�����i�J�i�j
        strSaveCont = strSaveCont & wsMoushikomiP.Range("J1").Value & "," '��E���E�����i�����j
        strSaveCont = strSaveCont & StrConv(wsMoushikomiP.Range("K3").Value, vbNarrow) & "," '�A����P�@����E�g��
        strSaveCont = strSaveCont & StrConv(wsMoushikomiP.Range("N3").Value, vbNarrow) & "," '�A����Q�@�Ζ���
        strSaveCont = strSaveCont & StrConv(wsMoushikomiP.Range("Q3").Value, vbNarrow) & "," '�A����R�@�e�`�w
        strSaveCont = strSaveCont & wsMoushikomiP.Range("T1").Value & "," '�c�̖�
        strSaveCont = strSaveCont & wsMoushikomiP.Range("U1").Value & "," '�c�̃R�[�h
        strSaveCont = strSaveCont & wsMoushikomiP.Range("V1").Value & "," '�c�̈��Ɋւ������
        strSaveCont = strSaveCont & wsMoushikomiP.Range("W1").Value & "," '�����R�[�h
        strSaveCont = strSaveCont & wsMoushikomiP.Range("X1").Value & "," '�Ј��R�[�h
        strSaveCont = strSaveCont & wsMoushikomiP.Range("Z1").Value & "," '���ۃR�[�h
        strSaveCont = strSaveCont & wsMoushikomiP.Range("AB1").Value & "," '�㗝�X�R�[�h
        strSaveCont = strSaveCont & wsMoushikomiP.Range("A1").Value & vbCrLf '�،��ԍ�
    Else
        strSaveCont = strSaveCont & wsKyoutsuU.Range("O2").Value & "," '�X�֔ԍ�
        strSaveCont = strSaveCont & wsKyoutsuU.Range("P2").Value & "," '�_��ҏZ���i�J�i�j
        strSaveCont = strSaveCont & wsKyoutsuU.Range("Q2").Value & "," '�_��ҏZ���i�����j
        strSaveCont = strSaveCont & wsKyoutsuU.Range("R2").Value & "," '�@�l���i�J�i�j
        strSaveCont = strSaveCont & wsKyoutsuU.Range("S2").Value & "," '�@�l���i�����j
        strSaveCont = strSaveCont & wsKyoutsuU.Range("T2").Value & "," '��E���E�����i�J�i�j
        strSaveCont = strSaveCont & wsKyoutsuU.Range("U2").Value & "," '��E���E�����i�����j
        strSaveCont = strSaveCont & wsKyoutsuU.Range("V2").Value & "," '�A����P�@����E�g��
        strSaveCont = strSaveCont & wsKyoutsuU.Range("W2").Value & "," '�A����Q�@�Ζ���
        strSaveCont = strSaveCont & wsKyoutsuU.Range("X2").Value & "," '�A����R�@�e�`�w
        strSaveCont = strSaveCont & wsKyoutsuU.Range("Y2").Value & "," '�c�̖�
        strSaveCont = strSaveCont & wsKyoutsuU.Range("Z2").Value & "," '�c�̃R�[�h
        strSaveCont = strSaveCont & wsKyoutsuU.Range("AA2").Value & "," '�c�̈��Ɋւ������
        strSaveCont = strSaveCont & wsKyoutsuU.Range("AB2").Value & "," '�����R�[�h
        strSaveCont = strSaveCont & wsKyoutsuU.Range("AC2").Value & "," '�Ј��R�[�h
        strSaveCont = strSaveCont & wsKyoutsuU.Range("AD2").Value & "," '���ۃR�[�h
        strSaveCont = strSaveCont & wsKyoutsuU.Range("AE2").Value & "," '�㗝�X�R�[�h
        strSaveCont = strSaveCont & wsKyoutsuU.Range("AF2").Value & vbCrLf '�،��ԍ�
    End If
    
    fncTempSave_Kyotsu = strSaveCont
    
    Set wsKyoutsuU = Nothing
    
End Function

Public Function fncTempSave_Meisai(ByRef intCarCnt As Integer, ByVal strMeisai As String, _
                                    ByRef blnOutFlg As Boolean, ByVal blnDownloadflg) As String

    '�֐����FfncTempSave_Meisai
    '���e�@�F���׏�񃌃R�[�h�i2�s�ڈȍ~�j����
    '�����@�F
    '        intCarCnt          = ���t�ۑ䐔
    '        strMeisai          = �ǉ��R���g���[���̖��O
    '        objAllOtherRate    = �R�s�[���̃I�u�W�F�N�g�̓��͓��e(�V�K�쐬�E���͓��e�̑��݂��Ȃ��ꍇ�́A�u�����N)
    '        intFileMaxCar      = ���ד��̓V�[�g�ɑ��݂���ŏI�s�̃R���g���[���̔ԍ�
    '        blnOutFlg          = �V�K�ǉ��E�R�s�[�̔��ʃt���O
    
    Dim strDate As String
    
    strDate = ""

    If blnOutFlg Then
        strMeisai = ""
        blnOutFlg = False
    End If
    
    Dim wsMeisai As Worksheet
    Call subSetSheet(1, wsMeisai)       '�V�[�g�I�u�W�F�N�g(���ד���)
    
    Dim wsTextM As Worksheet
    Call subSetSheet(7, wsTextM)
    
    '�����s
    Dim strTmpRow As Integer
    strTmpRow = 20 + intCarCnt
    
    With wsMeisai
        strMeisai = strMeisai & "2" & ","                                                              '���R�[�h�敪
        strMeisai = strMeisai & fncTrimComma(CStr(.Cells(strTmpRow, 4))) & ","                         '�p�r�Ԏ�i�R�[�h�j
        If blnDownloadflg Then
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 5))) & ","                     '�Ԗ�
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 8))) & ","                     '�^��
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 9))) & ","                     '�d�l
            '���x�o�^�N��
            strDate = fncWarekiCheck(Trim(.Cells(strTmpRow, 10)) & "1��", 6)
            If strDate = "" Then
                strDate = fncDateCheck(Trim(.Cells(strTmpRow, 10)), True, True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(Trim(.Cells(strTmpRow, 10))), 6) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 10))) & ","
                End If
            Else
                strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 10))) & ","
            End If

        Else
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 5)) & ","                           '�Ԗ�
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 8)) & ","                           '�^��
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 9)) & ","                           '�d�l
            '���x�o�^�N��
            strDate = fncWarekiCheck(.Cells(strTmpRow, 10) & "1��", 6)
            If strDate = "" Then
                strDate = fncDateCheck(.Cells(strTmpRow, 10), True, True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(.Cells(strTmpRow, 10)), 6) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 10)) & ","
                End If
            Else
                strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 10)) & ","
            End If
        End If
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 12).Value), "AD") & ","     '������
        
        '�O0�폜
        Dim i As String
        If blnDownloadflg Then
            i = fncTrimComma(CStr(Trim(.Cells(strTmpRow, 13))))              '�r�C��
        Else
            i = fncTrimComma(CStr(.Cells(strTmpRow, 13)))                    '�r�C��
        End If
        Do Until Len(i) <= 2
            If Left(i, 1) = "0" And IsNumeric(Mid(i, 2, 1)) Then
                i = Mid(i, 2)
            Else
                Exit Do
            End If
        Loop
        strMeisai = strMeisai & i & ","                    '�r�C��
        
        strMeisai = strMeisai & fncTrimComma(CStr(.Cells(strTmpRow, 14))) & ","                        '2.5���b�g�����f�B�[�[�����Ɨp���^��p��
        strMeisai = strMeisai & "" & ","                                                               '��ی���_���N����    :�m���t���[�g
        strMeisai = strMeisai & "" & ","                                                               '�m���t���[�g����     :�m���t���[�g
        strMeisai = strMeisai & "" & ","                                                               '���̗L�K�p����       :�m���t���[�g
        strMeisai = strMeisai & "" & ","                                                               '�m���t���[�g�������� :�m���t���[�g
        strMeisai = strMeisai & "" & ","                                                               '�c�̊�����           :�m���t���[�g
        strMeisai = strMeisai & "" & ","                                                               '�S�[���h�Ƌ�����     :�m���t���[�g
        strMeisai = strMeisai & "" & ","                                                               '�g�p�ړI             :�m���t���[�g
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 24) Like "*����*", "3 ", "") & ","                        '����
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 24) Like "*�����^�J�[*", "1", "") & ","                   '�����^�J�[
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 24) Like "*���K��*", "5 ", "") & ","                      '���K��
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 24) Like "*�u�[���ΏۊO*", "1 ", "") & ","                '�u�[���ΏۊO
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 24) Like "*���[�X�J�[�I�[�v���|���V�[*", "80", "") & ","  '���[�X�J�[�I�[�v���|���V�[
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 24) Like "*�I�[�v���|���V�[��������*", "93", "") & ","    '�I�[�v���|���V�[��������
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 24) Like "*�����L*", fncFindName("�����L", "AT"), _
                                IIf(.Cells(strTmpRow, 24) Like "*���L*", fncFindName("���L", "AT"), "")) & ","  '���L�E�����L��
        ''2018/3 ���ذĖ��וt�@�\�ǉ�
        If blnMoushikomiflg Then
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 16).Value) & ","                              '�ԗ������N���X
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 17).Value) & ","                              '�ΐl�����N���X
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 18).Value) & ","                              '�Ε������N���X
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 19).Value) & ","                              '���Q�����N���X
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 20).Value) & ","                              '�V�Ԋ���
            strMeisai = strMeisai & IIf(.Cells(strTmpRow, 24) Like "*����敪*", "8", "") & ","                  '����敪
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 22).Value) & ","                              '�ԗ��������i
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 21).Value) & ","                              '�ԗ�������i
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 47).Value) & ","                              '���v�ی���
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 48).Value) & ","                              '����ی���
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 49).Value) & ","                              '�N�ԕی���
        Else
            strMeisai = strMeisai & "" & ","                                                               '�ԗ������N���X
            strMeisai = strMeisai & "" & ","                                                               '�ΐl�����N���X
            strMeisai = strMeisai & "" & ","                                                               '�Ε������N���X
            strMeisai = strMeisai & "" & ","                                                               '���Q�����N���X
            strMeisai = strMeisai & "" & ","                                                               '�V�Ԋ���
            strMeisai = strMeisai & "" & ","                                                               '����敪
            strMeisai = strMeisai & "" & ","                                                               '�ԗ��������i
            strMeisai = strMeisai & "" & ","                                                               '�ԗ�������i
            strMeisai = strMeisai & "" & ","                                                               '���v�ی���
            strMeisai = strMeisai & "" & ","                                                               '����ی���
            strMeisai = strMeisai & "" & ","                                                               '�N�ԕی���
        End If
        strMeisai = strMeisai & "" & ","                                                               '�N�����             :�m���t���[�g
        strMeisai = strMeisai & "" & ","                                                               '����^�]�ґΏۊO���� :�m���t���[�g
        strMeisai = strMeisai & "" & ","                                                               '�^�]�Ҍ������       :�m���t���[�g
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 44)) <> "", "1", "") & ","          '�]�ƈ����������
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 25).Value), "BJ") & ","     '�ԗ��ی��̎��
        If blnDownloadflg Then
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 26)))) = "", "", Val(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 26)))))) & ","  '�ԗ��ی����z
        Else
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(.Cells(strTmpRow, 26))) = "", "", Val(fncTrimComma(CStr(.Cells(strTmpRow, 26))))) & ","   '�ԗ��ی����z
        End If
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 27)), "BN") & ","           '�ԗ��ƐӋ��z
        strMeisai = strMeisai & fncTrimComma(CStr(.Cells(strTmpRow, 45))) & ","                        '��ԓ��Z�b�g
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 28)) <> "", "2", "") & ","          '�ԗ��S���Ք����
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 30)) <> "", "1", "") & ","          '�ԗ�����ΏۊO����
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 29)) <> "", "1", "") & ","          '�ԗ����ߏC����p����
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 31) <> "������", "", _
                                fncFindName(fncTrimComma(.Cells(strTmpRow, 31)), "CD")) & ","          '�ΐl������
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 31) <> "�ΏۊO", "", _
                                fncFindName(fncTrimComma(.Cells(strTmpRow, 31)), "CD")) & ","          '�ΐl�ΏۊO
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 31) = "������", "", _
                                IIf(.Cells(strTmpRow, 31) = "�ΏۊO", "", _
                                    fncTrimComma(fncFindName(.Cells(strTmpRow, 31), "CD")))) & ","     '�ΐl�����ی����z
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 32)) <> "", "1", "") & ","          '�������̏��Q����
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 33)) <> "", "1", "") & ","          '���ی��Ԏ��̏��Q����
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 34) <> "������", "", _
                                fncFindName(fncTrimComma(.Cells(strTmpRow, 34)), "CH")) & ","          '�Ε�������
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 34) <> "�ΏۊO", "", _
                                fncFindName(fncTrimComma(.Cells(strTmpRow, 34)), "CH")) & ","          '�ΐl�ΏۊO
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 34) = "������", "", _
                                IIf(.Cells(strTmpRow, 34) = "�ΏۊO", "", _
                                    fncTrimComma(fncFindName(.Cells(strTmpRow, 34), "CH")))) & ","     '�Ε������ی����z
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 35)), "BR") & ","           '�Ε��ƐӋ��z
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 36)) <> "", "1", "") & ","          '�Ε����ߏC����p����
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 37) = "�ΏۊO", "", _
                                    fncTrimComma(fncFindName(.Cells(strTmpRow, 37), "CL"))) & ","      '�l�g���Q 1��
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 37) <> "�ΏۊO", "", _
                                    fncFindName(fncTrimComma(.Cells(strTmpRow, 37)), "CL")) & ","      '�l�g�ΏۊO
        If blnDownloadflg Then
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 38)))) = "", "", Val(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 38)))))) & ","  '�l�g���Q 1����
        Else
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(.Cells(strTmpRow, 38))) = "", "", Val(fncTrimComma(CStr(.Cells(strTmpRow, 38))))) & ","   '�l�g���Q 1����
        End If
        strMeisai = strMeisai & "" & ","                                                               '�����Ԏ��̓��� :�m���t���[�g
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 39) = "�ΏۊO", "", _
                                    fncTrimComma(fncFindName(.Cells(strTmpRow, 39), "CP"))) & ","      '���S�E��⏝�Q�ی����z 1��
        If .Cells(strTmpRow, 39) = "" Then
            .Cells(strTmpRow, 39) = "�ΏۊO"
        End If
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 39) <> "�ΏۊO", "", _
                                    fncFindName(fncTrimComma(.Cells(strTmpRow, 39)), "CP")) & ","      '�����ΏۊO
        If blnDownloadflg Then
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 40)))) = "", "", Val(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 40)))))) & ","            '���S�E��⏝�Q�ی����z 1����
        Else
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(.Cells(strTmpRow, 40))) = "", "", Val(fncTrimComma(CStr(.Cells(strTmpRow, 40))))) & ","                    '���S�E��⏝�Q�ی����z 1����
        End If
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 41)) <> "", "2", "") & ","          '����������
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 42)) <> "", "1", "") & ","          '���Ǝ��p����
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 43)) <> "", "1", "") & ","          '�ٌ�m��p����
        strMeisai = strMeisai & "" & ","                                                               '�t�@�~���[�o�C�N����   :�m���t���[�g
        strMeisai = strMeisai & "" & ","                                                               '�l�����ӔC�⏞�t���O :�m���t���[�g
        ''2018/3 ���ذĖ��וt�@�\�ǉ�
        If blnMoushikomiflg Then
            strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 50)) = "�g�c�G���[�L", "1", _
                                    IIf(fncTrimComma(.Cells(strTmpRow, 50)) = "�x���L", "2", "")) & "," '�g�c�x���t���O
        Else
            strMeisai = strMeisai & "" & ","                                                           '�g�c�x���t���O
        End If
        If blnDownloadflg Then
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 6))) & ","                     '�o�^�ԍ�
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 7))) & ","                     '�ԑ�ԍ�
            '�Ԍ�������
            strDate = fncWarekiCheck(Trim(.Cells(strTmpRow, 11)))
            If strDate = "" Then
                strDate = fncDateCheck(Trim(.Cells(strTmpRow, 11)), True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(Trim(.Cells(strTmpRow, 11))), 8) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 11))) & ","
                End If
            Else
                strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 11))) & ","
            End If
        Else
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 6)) & ","                           '�o�^�ԍ�
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 7)) & ","                           '�ԑ�ԍ�
            '�Ԍ�������
            strDate = fncWarekiCheck(.Cells(strTmpRow, 11))
            If strDate = "" Then
                strDate = fncDateCheck(.Cells(strTmpRow, 11), True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(.Cells(strTmpRow, 11)), 8) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 11)) & ","
                End If
                    
            Else
                strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 11)) & ","
            End If
        End If
        strTmpRow = strTmpRow - 20
        strMeisai = strMeisai & fncTrimComma(wsTextM.Cells(strTmpRow, 72)) & ","       '�o�^�ԍ�(�J�i)
        strTmpRow = strTmpRow + 20
'        strMeisai = strMeisai & "" & ","                                                               '�o�^�ԍ�(�J�i)
        'ASV����
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 15)) <> "", "1", "") & ","          '�u�����N�łȂ��ꍇ�u1:�K�p����v
        '�ԗ��������s�K�p����
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 46)) <> "", "1", "") & ","          '�u�����N�łȂ��ꍇ�u1:�s�K�p�v
'        strMeisai = strMeisai & vbCrLf
    End With
    '��ی��ҏZ�� (�J�i)
    strMeisai = strMeisai & ","
    '��ی��Ҏ��� (�J�i)
    strMeisai = strMeisai & ","
    '��ی��Ҏ��� (����)
    strMeisai = strMeisai & ","
    '�Ƌ��؂̐F
    strMeisai = strMeisai & ","
    '�Ƌ��ؗL������
    strMeisai = strMeisai & ","
    '�ԗ����L�Ҏ��� (�J�i)
    strMeisai = strMeisai & ","
    '�ԗ����L�Ҏ��� (����)
    strMeisai = strMeisai & ","
    '���L�����ۂ܂��̓��[�X��
    strMeisai = strMeisai & vbCrLf

        
    fncTempSave_Meisai = strMeisai

    Set wsMeisai = Nothing

End Function

Public Function fncTempSave_NonFleetMeisai(ByRef intCarCnt As Integer, ByVal strMeisai As String, _
                                    ByRef blnOutFlg As Boolean, ByVal blnDownloadflg) As String

    '�֐����FfncTempSave_NonFleetMeisai �i2018/3 ���ذĖ��וt�@�\�ǉ��j
    '���e�@�F���׏�񃌃R�[�h�i2�s�ڈȍ~�j����
    '�����@�F
    '        intCarCnt          = ���t�ۑ䐔
    '        strMeisai          = �ǉ��R���g���[���̖��O
    '        objAllOtherRate    = �R�s�[���̃I�u�W�F�N�g�̓��͓��e(�V�K�쐬�E���͓��e�̑��݂��Ȃ��ꍇ�́A�u�����N)
    '        intFileMaxCar      = ���ד��̓V�[�g�ɑ��݂���ŏI�s�̃R���g���[���̔ԍ�
    '        blnOutFlg          = �V�K�ǉ��E�R�s�[�̔��ʃt���O
    
    Dim strDate As String
    
    strDate = ""

    If blnOutFlg Then
        strMeisai = ""
        blnOutFlg = False
    End If
    
    Dim wsMeisai As Worksheet
    Call subSetSheet(17, wsMeisai)      '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
    
    Dim wsTextM As Worksheet
    Call subSetSheet(7, wsTextM)
    
    '���ʉ�ʂ̃m���t���[�g���������A�c�̊����̎擾�p�i2018/3 ���ذĖ��וt�@�\�ǉ��j
    Dim wsKyoutsuU As Worksheet             '�R�[�h�l���[�N�V�[�g
    Call subSetSheet(2, wsKyoutsuU)         '�V�[�g�I�u�W�F�N�g(�ʎ��@���ʍ���)
    
    Dim wsMeisaiP As Worksheet
    Call subSetSheet(19, wsMeisaiP)

    '�����s
    Dim strTmpRow As Integer
    strTmpRow = 20 + intCarCnt

    With wsMeisai
        strMeisai = strMeisai & "2" & ","                                                              '���R�[�h�敪
        strMeisai = strMeisai & fncTrimComma(CStr(.Cells(strTmpRow, 4))) & ","                         '�p�r�Ԏ�i�R�[�h�j
        If blnDownloadflg Then
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 5))) & ","                     '�Ԗ�
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 8))) & ","                     '�^��
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 9))) & ","                     '�d�l
            '���x�o�^�N��
            strDate = fncWarekiCheck(Trim(.Cells(strTmpRow, 10)) & "1��", 6)
            If strDate = "" Then
                strDate = fncDateCheck(Trim(.Cells(strTmpRow, 10)), True, True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(Trim(.Cells(strTmpRow, 10))), 6) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 10))) & ","
                End If
            Else
                strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 10))) & ","
            End If

        Else
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 5)) & ","                           '�Ԗ�
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 8)) & ","                           '�^��
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 9)) & ","                           '�d�l
            '���x�o�^�N��
            strDate = fncWarekiCheck(.Cells(strTmpRow, 10) & "1��", 6)
            If strDate = "" Then
                strDate = fncDateCheck(.Cells(strTmpRow, 10), True, True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(.Cells(strTmpRow, 10)), 6) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 10)) & ","
                End If
            Else
                strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 10)) & ","
            End If
        End If
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 12).Value), "AD") & ","     '������
        
        '�O0�폜
        Dim i As String
        If blnDownloadflg Then
            i = fncTrimComma(CStr(Trim(.Cells(strTmpRow, 13))))              '�r�C��
        Else
            i = fncTrimComma(CStr(.Cells(strTmpRow, 13)))                    '�r�C��
        End If
        Do Until Len(i) <= 2
            If Left(i, 1) = "0" And IsNumeric(Mid(i, 2, 1)) Then
                i = Mid(i, 2)
            Else
                Exit Do
            End If
        Loop
        strMeisai = strMeisai & i & ","                    '�r�C��
'        If blnDownloadflg Then
'            strMeisai = strMeisai & fncTrimComma(CStr(Trim(.Cells(strTmpRow, 13)))) & ","              '�r�C��
'        Else
'            strMeisai = strMeisai & fncTrimComma(CStr(.Cells(strTmpRow, 13))) & ","                    '�r�C��
'        End If
        strMeisai = strMeisai & fncTrimComma(CStr(.Cells(strTmpRow, 14))) & ","                        '2.5���b�g�����f�B�[�[�����Ɨp���^��p��
        
       If blnDownloadflg Then
            '��ی��Ґ��N���� �i2018/3 ���ذĖ��וt�@�\�ǉ��j
            strDate = fncWarekiCheck(Trim(.Cells(strTmpRow, 15)))
            If strDate = "" Then
                strDate = fncDateCheck(Trim(.Cells(strTmpRow, 15)), True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(Trim(.Cells(strTmpRow, 15))), 8) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 15))) & ","
                End If
            Else
                strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 15))) & ","
            End If

        Else
            '��ی��Ґ��N����
            strDate = fncWarekiCheck(.Cells(strTmpRow, 15))
            If strDate = "" Then
                strDate = fncDateCheck(.Cells(strTmpRow, 15), True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(.Cells(strTmpRow, 15)), 8) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 15)) & ","
                End If

            Else
                strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 15)) & ","
            End If
       End If

       '�m���t���[�g�����i2018/3 ���ذĖ��וt�@�\�ǉ��j
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 16).Value), "AH") & ","

       '���̗L�K�p���ԁi2018/3 ���ذĖ��וt�@�\�ǉ��j
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 17).Value), "AL") & ","

        '�m���t���[�g���������i2018/3 ���ذĖ��וt�@�\�ǉ��j
        strMeisai = strMeisai & wsKyoutsuU.Range("M2").Value & ","    '�u�ʎ��@���ʍ��ځv�V�[�g���擾

        '�c�̊������i2018/3 ���ذĖ��וt�@�\�ǉ��j
        strMeisai = strMeisai & wsKyoutsuU.Range("N2").Value & ","    '�u�ʎ��@���ʍ��ځv�V�[�g���擾

        '�S�[���h�Ƌ������i2018/3 ���ذĖ��וt�@�\�ǉ��j
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 18)) <> "", "1", "") & ","          '�u�����N�łȂ��ꍇ�u1:�K�p����v

        '�g�p�ړI�i2018/3 ���ذĖ��וt�@�\�ǉ��j
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 19).Value), "DC") & ","

        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 32) Like "*����*", "3 ", "") & ","                        '����
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 32) Like "*�����^�J�[*", "1", "") & ","                   '�����^�J�[
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 32) Like "*���K��*", "5 ", "") & ","                      '���K��
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 32) Like "*�u�[���ΏۊO*", "1 ", "") & ","                '�u�[���ΏۊO
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 32) Like "*���[�X�J�[�I�[�v���|���V�[*", "80", "") & ","  '���[�X�J�[�I�[�v���|���V�[
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 32) Like "*�I�[�v���|���V�[��������*", "93", "") & ","    '�I�[�v���|���V�[��������
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 32) Like "*�����L*", fncFindName("�����L", "AT"), _
                                IIf(.Cells(strTmpRow, 32) Like "*���L*", fncFindName("���L", "AT"), "")) & ","  '���L�E�����L��


        ''2018/3 ���ذĖ��וt�@�\�ǉ�
        If blnMoushikomiflg Then
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 24).Value) & ","                        '�ԗ������N���X
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 25).Value) & ","                        '�ΐl�����N���X
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 26).Value) & ","                        '�Ε������N���X
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 27).Value) & ","                        '���Q�����N���X
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 28).Value) & ","                        '�V�Ԋ���
            strMeisai = strMeisai & IIf(.Cells(strTmpRow, 32) Like "*����敪*", "8", "") & ","            '����敪
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 30).Value) & ","                        '�ԗ��������i
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 29).Value) & ","                        '�ԗ�������i
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 57).Value) & ","                        '���v�ی���
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 58).Value) & ","                        '����ی���
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 59).Value) & ","                        '�N�ԕی���
        Else
            strMeisai = strMeisai & "" & ","                                                               '�ԗ������N���X
            strMeisai = strMeisai & "" & ","                                                               '�ΐl�����N���X
            strMeisai = strMeisai & "" & ","                                                               '�Ε������N���X
            strMeisai = strMeisai & "" & ","                                                               '���Q�����N���X
            strMeisai = strMeisai & "" & ","                                                               '�V�Ԋ���
            strMeisai = strMeisai & "" & ","                                                               '����敪
            strMeisai = strMeisai & "" & ","                                                               '�ԗ��������i
            strMeisai = strMeisai & "" & ","                                                               '�ԗ�������i
            strMeisai = strMeisai & "" & ","                                                               '���v�ی���
            strMeisai = strMeisai & "" & ","                                                               '����ی���
            strMeisai = strMeisai & "" & ","                                                               '�N�ԕی���
        End If
        
        '�N������i2018/3 ���ذĖ��וt�@�\�ǉ��j
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 20).Value), "BB") & ","
        
        '����^�]�ґΏۊO����i2018/3 ���ذĖ��וt�@�\�ǉ��j
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 21)) <> "", "1", "") & ","          '�u�����N�łȂ��ꍇ�u1:�ΏۊO�v
        
        '�^�]�Ҍ������i2018/3 ���ذĖ��וt�@�\�ǉ��j
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 22).Value), "BF") & ","
              
        '�]�ƈ����������i2018/3 ���ذĖ��וt�@�\�ǉ��j
        strMeisai = strMeisai & "" & ","
        
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 33).Value), "BJ") & ","     '�ԗ��ی��̎��
        If blnDownloadflg Then
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 34)))) = "", "", Val(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 34)))))) & ","  '�ԗ��ی����z
        Else
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(.Cells(strTmpRow, 34))) = "", "", Val(fncTrimComma(CStr(.Cells(strTmpRow, 34))))) & ","              '�ԗ��ی����z
        End If
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 35)), "BN") & ","           '�ԗ��ƐӋ��z
        strMeisai = strMeisai & fncTrimComma(CStr(.Cells(strTmpRow, 55))) & ","                        '���̑�ԁE�g�̉��i����
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 36)) <> "", "2", "") & ","          '�ԗ��S���Ք����
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 38)) <> "", "1", "") & ","          '�ԗ�����ΏۊO����
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 37)) <> "", "1", "") & ","          '�ԗ����ߏC����p����
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 39) <> "������", "", _
                                fncFindName(fncTrimComma(.Cells(strTmpRow, 39)), "CD")) & ","          '�ΐl������
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 39) <> "�ΏۊO", "", _
                                fncFindName(fncTrimComma(.Cells(strTmpRow, 39)), "CD")) & ","          '�ΐl�ΏۊO
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 39) = "������", "", _
                                IIf(.Cells(strTmpRow, 39) = "�ΏۊO", "", _
                                    fncTrimComma(fncFindName(.Cells(strTmpRow, 39), "CD")))) & ","     '�ΐl�����ی����z
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 40)) <> "", "1", "") & ","          '�������̏��Q����
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 41)) <> "", "1", "") & ","          '���ی��Ԏ��̏��Q����
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 42) <> "������", "", _
                                fncFindName(fncTrimComma(.Cells(strTmpRow, 42)), "CH")) & ","          '�Ε�������
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 42) <> "�ΏۊO", "", _
                                fncFindName(fncTrimComma(.Cells(strTmpRow, 42)), "CH")) & ","          '�ΐl�ΏۊO
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 42) = "������", "", _
                                IIf(.Cells(strTmpRow, 42) = "�ΏۊO", "", _
                                    fncTrimComma(fncFindName(.Cells(strTmpRow, 42), "CH")))) & ","     '�Ε������ی����z
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 43)), "BR") & ","           '�Ε��ƐӋ��z
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 44)) <> "", "1", "") & ","          '�Ε����ߏC����p����
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 45) = "�ΏۊO", "", _
                                    fncTrimComma(fncFindName(.Cells(strTmpRow, 45), "CL"))) & ","      '�l�g���Q 1��
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 45) <> "�ΏۊO", "", _
                                    fncFindName(fncTrimComma(.Cells(strTmpRow, 45)), "CL")) & ","      '�l�g�ΏۊO
        If blnDownloadflg Then
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 46)))) = "", "", Val(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 46)))))) & ","  '�l�g���Q 1����
        Else
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(.Cells(strTmpRow, 46))) = "", "", Val(fncTrimComma(CStr(.Cells(strTmpRow, 46))))) & ","   '�l�g���Q 1����
        End If
        '�����Ԏ��̓���i2018/3 ���ذĖ��וt�@�\�ǉ��j
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 54)) <> "", "2", "") & ","          '�u�����N�łȂ��ꍇ�u2:�K�p����v

        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 47) = "�ΏۊO", "", _
                                    fncTrimComma(fncFindName(.Cells(strTmpRow, 47), "CP"))) & ","      '���S�E��⏝�Q�ی����z 1��
        If .Cells(strTmpRow, 47) = "" Then
            .Cells(strTmpRow, 47) = "�ΏۊO"
        End If
        strMeisai = strMeisai & IIf(.Cells(strTmpRow, 47) <> "�ΏۊO", "", _
                                    fncFindName(fncTrimComma(.Cells(strTmpRow, 47)), "CP")) & ","      '�����ΏۊO
        If blnDownloadflg Then
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 48)))) = "", "", Val(fncTrimComma(CStr(Trim(.Cells(strTmpRow, 48)))))) & ","        '���S�E��⏝�Q�ی����z 1����
        Else
            strMeisai = strMeisai & IIf(fncTrimComma(CStr(.Cells(strTmpRow, 48))) = "", "", Val(fncTrimComma(CStr(.Cells(strTmpRow, 48))))) & ","                    '���S�E��⏝�Q�ی����z 1����
        End If
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 49)) <> "", "2", "") & ","          '����������
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 50)) <> "", "1", "") & ","          '���Ǝ��p����
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 51)) <> "", "1", "") & ","          '�ٌ�m��p����

        '�t�@�~���[�o�C�N����i2018/3 ���ذĖ��וt�@�\�ǉ��j
        strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 52).Value), "BV") & ","

        '�l�����ӔC�⏞����
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 53)) <> "", "1", "") & ","          '�u�����N�łȂ��ꍇ�u1:�K�p����v
        ''2018/3 ���ذĖ��וt�@�\�ǉ�
        If blnMoushikomiflg Then
            strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 60)) = "�g�c�G���[�L", "1", _
                                    IIf(fncTrimComma(.Cells(strTmpRow, 60)) = "�x���L", "2", "")) & "," '�g�c�x���t���O

        Else
            strMeisai = strMeisai & "" & ","                                                           '�g�c�x���t���O
        End If
        If blnDownloadflg Then
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 6))) & ","                     '�o�^�ԍ�
            strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 7))) & ","                     '�ԑ�ԍ�
            '�Ԍ�������
            strDate = fncWarekiCheck(Trim(.Cells(strTmpRow, 11)))
            If strDate = "" Then
                strDate = fncDateCheck(Trim(.Cells(strTmpRow, 11)), True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(Trim(.Cells(strTmpRow, 11))), 8) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 11))) & ","
                End If
            Else
                strMeisai = strMeisai & fncTrimComma(Trim(.Cells(strTmpRow, 11))) & ","
            End If
        Else
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 6)) & ","                           '�o�^�ԍ�
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 7)) & ","                           '�ԑ�ԍ�
            '�Ԍ�������
            strDate = fncWarekiCheck(.Cells(strTmpRow, 11))
            If strDate = "" Then
                strDate = fncDateCheck(.Cells(strTmpRow, 11), True)
                If strDate = "" Then
                    strMeisai = strMeisai & fncToSeireki(fncTrimComma(.Cells(strTmpRow, 11)), 8) & ","
                Else
                    strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 11)) & ","
                End If

            Else
                strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 11)) & ","
            End If
        End If
        strTmpRow = strTmpRow - 20
        strMeisai = strMeisai & fncTrimComma(wsTextM.Cells(strTmpRow, 72)) & "," '�o�^�ԍ�(�J�i)
        strTmpRow = strTmpRow + 20
'        strMeisai = strMeisai & "" & ","                                                               '�o�^�ԍ�(�J�i)
        'ASV����
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 23)) <> "", "1", "") & ","          '�u�����N�łȂ��ꍇ�u1:�K�p����v
        '�ԗ��������s�K�p����
        strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 56)) <> "", "1", "") & ","          '�u�����N�łȂ��ꍇ�u1:�s�K�p�v
'        strMeisai = strMeisai & vbCrLf
    End With

    If blnChouhyouflg Then
        With wsMeisaiP
            strTmpRow = strTmpRow - 14
    '       ��ی��ҏZ�� (��)
            strMeisai = strMeisai & StrConv(fncTrimComma(.Cells(strTmpRow, 3)), vbWide) & ","
    '       ��ی��Ҏ��� (��)
            strMeisai = strMeisai & StrConv(fncTrimComma(.Cells(strTmpRow, 8)), vbWide) & ","
    '       ��ی��Ҏ��� (����)
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 11)) & ","
    '       �Ƌ��؂̐F
            strMeisai = strMeisai & fncFindName(fncTrimComma(.Cells(strTmpRow, 24).Value), "DT") & ","
    '       �Ƌ��ؗL������
            strDate = fncDateCheck(.Cells(strTmpRow, 27), False)
            If strDate = "" Then
                strMeisai = strMeisai & fncToSeireki(fncTrimComma(.Cells(strTmpRow, 27)), 8) & ","
            Else
                strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 27)) & ","
            End If
    '        strMeisai = strMeisai & fncToSeireki(fncTrimComma(.Cells(strTmpRow, 27)), 8) & ","
            strTmpRow = strTmpRow + 12
    '       �ԗ����L�Ҏ����i�Łj
            strMeisai = strMeisai & StrConv(fncTrimComma(.Cells(strTmpRow, 16)), vbWide) & ","
    '       �ԗ����L�Ҏ����i�����j
            strMeisai = strMeisai & fncTrimComma(.Cells(strTmpRow, 24)) & ","
    '       ���L�����ۂ܂��̓��[�X��
            strMeisai = strMeisai & IIf(fncTrimComma(.Cells(strTmpRow, 31)) <> "", "1", "")
            strMeisai = strMeisai & vbCrLf
        End With
    Else
        With wsTextM
            strTmpRow = strTmpRow - 20
    '       ��ی��ҏZ�� (��)
            strMeisai = strMeisai & .Cells(strTmpRow, 75) & ","
    '       ��ی��Ҏ��� (��)
            strMeisai = strMeisai & .Cells(strTmpRow, 76) & ","
    '       ��ی��Ҏ��� (����)
            strMeisai = strMeisai & .Cells(strTmpRow, 77) & ","
    '       �Ƌ��؂̐F
            strMeisai = strMeisai & .Cells(strTmpRow, 78) & ","
    '       �Ƌ��ؗL������
            strMeisai = strMeisai & .Cells(strTmpRow, 79) & ","
    '       �ԗ����L�Ҏ����i�Łj
            strMeisai = strMeisai & .Cells(strTmpRow, 80) & ","
    '       �ԗ����L�Ҏ����i�����j
            strMeisai = strMeisai & .Cells(strTmpRow, 81) & ","
    '       ���L�����ۂ܂��̓��[�X��
            strMeisai = strMeisai & .Cells(strTmpRow, 82)
            strMeisai = strMeisai & vbCrLf
        End With
    End If
        
    fncTempSave_NonFleetMeisai = strMeisai

    Set wsMeisai = Nothing

End Function
Public Sub subClearAll()
    Dim objAll As Object
    Dim strTagetRange As String
    
    Dim wsMeisai As Worksheet           '���ד��̓��[�N�V�[�g
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(1, wsMeisai)       '�V�[�g�I�u�W�F�N�g(���ד���)
    Else
        Call subSetSheet(17, wsMeisai)      '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
    End If
    
    '�`���~
    Application.ScreenUpdating = False
    
    For Each objAll In wsMeisai.DrawingObjects
        If objAll.Name Like "chkMeisai*" Then
            objAll.Value = xlOff
            strTagetRange = wsMeisai.Shapes(objAll.Name).TopLeftCell.Address
            wsMeisai.Range(strTagetRange) = Left(wsMeisai.Range(strTagetRange), InStr(wsMeisai.Range(strTagetRange), "/") - 1) & "/False"
        End If
    Next objAll
    
    Set wsMeisai = Nothing
    Set objAll = Nothing
    
    '�`���~
    Application.ScreenUpdating = True
        
End Sub

''���ׂ̃��X�g�ݒ�
Public Sub subCmbInitialize()
    
    '�t���[�g
    Call subCmbSet("C", "Z2", 1)       '�p�r�Ԏ�
    Call subCmbSet("L", "AD2", 1)      '�����E�s����
    Call subCmbSet("O", "CT2", 1)      'ASV����
    Call subCmbSet("Y", "BJ2", 1)      '�ԗ��ی��̎��
    Call subCmbSet("AA", "BN2", 1)      '�ԗ��ƐӋ��z
    Call subCmbSet("AS", "CA2", 1)     '��ԓ��Z�b�g����
    Call subCmbSet("AE", "CD2", 1)     '�ΐl����
    Call subCmbSet("AH", "CH2", 1)     '�Ε�����
    Call subCmbSet("AI", "BR2", 1)     '�Ε��ƐӋ��z
    Call subCmbSet("AK", "CL2", 1)     '�l�g���Q(1��)
'    Call subCmbSet("AM", "CP2", 1)     '����ҏ��Q(1��)
    Call subCmbSet("AB", "CT2", 1)     '�ԗ��S���Ք����
    Call subCmbSet("AC", "CT2", 1)     '�ԗ����ߏC����p����
    Call subCmbSet("AD", "CW2", 1)     '�ԗ�����ΏۊO����
    Call subCmbSet("AF", "CW2", 1)     '�������̏��Q����
    Call subCmbSet("AG", "CW2", 1)     '���ی��Ԏ��̏��Q����
    Call subCmbSet("AJ", "CT2", 1)     '�Ε����ߏC����p����
    Call subCmbSet("AO", "CT2", 1)     '����������
    Call subCmbSet("AP", "CT2", 1)     '���Ǝ��p����
    Call subCmbSet("AQ", "CT2", 1)     '�ٌ�m��p����
    Call subCmbSet("AR", "CZ2", 1)     '�]�ƈ����������
    Call subCmbSet("AT", "DC2", 1)     '�ԗ��������s�K�p����
    
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    Call subCmbSet("C", "Z2", 2)      '�p�r�Ԏ�
    Call subCmbSet("L", "AD2", 2)      '�����E�s����
    Call subCmbSet("P", "DM2", 2)      '�m���t���[�g����
    Call subCmbSet("Q", "AL2", 2)      '���̗L�W���K�p����
    Call subCmbSet("R", "CT2", 2)      '�S�[���h�Ƌ�����
    Call subCmbSet("S", "DC2", 2)      '�g�p�ړI
    Call subCmbSet("T", "BB2", 2)      '�N�����
    Call subCmbSet("U", "CW2", 2)      '����^�]�ґΏۊO
    Call subCmbSet("V", "BF2", 2)      '�^�]�Ҍ���
    Call subCmbSet("W", "CT2", 2)      'ASV����
    Call subCmbSet("AG", "BJ2", 2)     '�ԗ��ی��̎��
    Call subCmbSet("AI", "BN2", 2)     '�ԗ��ƐӋ��z
    Call subCmbSet("BC", "CA2", 2)     '��ԓ��Z�b�g����
    Call subCmbSet("AM", "CD2", 2)     '�ΐl����
    Call subCmbSet("AP", "CH2", 2)     '�Ε�����
    Call subCmbSet("AQ", "BR2", 2)     '�Ε��ƐӋ��z
    Call subCmbSet("AS", "CL2", 2)     '�l�g���Q(1��)
'    Call subCmbSet("AU", "CP2", 2)     '����ҏ��Q(1��)
    Call subCmbSet("AJ", "CT2", 2)     '�ԗ��S���Ք����
    Call subCmbSet("AK", "CT2", 2)     '�ԗ����ߏC����p����
    Call subCmbSet("AL", "CW2", 2)     '�ԗ�����ΏۊO����
    Call subCmbSet("AN", "CW2", 2)     '�������̏��Q����
    Call subCmbSet("AO", "CW2", 2)     '���ی��Ԏ��̏��Q����
    Call subCmbSet("AR", "CT2", 2)     '�Ε����ߏC����p����
    Call subCmbSet("AW", "CT2", 2)     '����������
    Call subCmbSet("AX", "CT2", 2)     '���Ǝ��p����
    Call subCmbSet("AY", "CT2", 2)     '�ٌ�m��p����
    Call subCmbSet("AZ", "BV2", 2)     '�t�@�~���[�o�C�N����
    Call subCmbSet("BA", "CT2", 2)     '�l�����ӔC�⏞����
    Call subCmbSet("BB", "CT2", 2)     '�����Ԏ��̓���
    Call subCmbSet("BD", "DQ2", 2)     '�ԗ��������s�K�p����
    
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    Call subCmbSetMeisaiPrint("X", "DG2", 2, 6)   '�Ƌ��؂̐F
    Call subCmbSetMeisaiPrint("AE", "DJ2", 2, 18) '���L�����ۂ܂��̓��[�X��
    Call subCmbSetMeisaiPrint("F", "AH2", 2, 30)  '�O�_�񓙋�
    Call subCmbSetMeisaiPrint("G", "AL2", 2, 30)  '�O�_�񎖌̗L�W���K�p����
    
End Sub

''�v�Z�p�V�[�g�̃R���{�{�b�N�X��ݒ�
Public Sub subCmbSet(ByRef strTargetCol As String, ByRef strTargetCode As String, ByRef SheetFlg As Integer)
    Dim i As Integer                '���[�v�J�E���g
    Dim intMaxRow As Integer        '�����̖��׍s��
    
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    Dim wsMeisai As Worksheet           '���ד��̓��[�N�V�[�g
    If SheetFlg = 1 Then  '�t���[�g
        Call subSetSheet(1, wsMeisai)       '�V�[�g�I�u�W�F�N�g(���ד���)
    Else
        Call subSetSheet(17, wsMeisai)      '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
    End If
    
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    Dim wsCode As Worksheet                 '�ʎ��R�[�h�l�̃I�u�W�F�N�g
    If SheetFlg = 1 Then  '�t���[�g
        Call subSetSheet(4, wsCode)         '�V�[�g�I�u�W�F�N�g(�ʎ��@�R�[�h�l)
    Else
        Call subSetSheet(16, wsCode)         '�V�[�g�I�u�W�F�N�g(�ʎ��@�R�[�h�l�i�m���t���[�g�j�j
    End If
    
    '���t�ۑ䐔�擾
    intMaxRow = Val(wsMeisai.OLEObjects("txtSouhuho").Object.Value)
    
    '�����쐬����Ă��閾�׍s���������[�v����
    With wsMeisai
        For i = 1 To intMaxRow
            '�����ݒ���폜
            .Range(strTargetCol & (20 + i)).Validation.Delete
            '�Z�������X�g�ɕύX���A�Y���̃R�[�h�l�������̂�ݒ�
            .Range(strTargetCol & (20 + i)).Validation.Add _
                Type:=xlValidateList, _
                Formula1:="=" & wsCode.Range(wsCode.Range(strTargetCode), wsCode.Cells(wsCode.Rows.Count, _
                            wsCode.Range(strTargetCode).Column).End(xlUp)).Address(External:=True)
            '���X�g�ȊO�̕��������͂ł���悤�ɐݒ�
            If strTargetCode = "Z2" Or strTargetCode = "AD2" Then
                .Range(strTargetCol & (20 + i)).Validation.IMEMode = xlIMEModeOn
                .Range(strTargetCol & (20 + i)).Validation.ShowError = False
            End If
        Next i
    End With
    
    Set wsCode = Nothing
    
End Sub


'2018/3 ���ذĖ��וt�@�\�ǉ�
''���׏������ʂ̃R���{�{�b�N�X��ݒ�
Public Sub subCmbSetMeisaiPrint(ByRef strTargetCol As String, ByRef strTargetCode As String, ByRef SheetFlg As Integer, ByRef intHeadNum As Integer)
    Dim i As Integer                '���[�v�J�E���g
    Dim intMaxRow As Integer        '�����̖��׍s��
    
    Dim wsMeisai As Worksheet           '���ד��̓��[�N�V�[�g
    If SheetFlg = 1 Then  '�t���[�g
        Call subSetSheet(18, wsMeisai)       '�V�[�g�I�u�W�F�N�g(���ד���)
    Else
        Call subSetSheet(19, wsMeisai)      '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
    End If
    
    Dim wsCode As Worksheet                 '�ʎ��R�[�h�l�̃I�u�W�F�N�g
    If SheetFlg = 1 Then  '�t���[�g
        Call subSetSheet(4, wsCode)         '�V�[�g�I�u�W�F�N�g(�ʎ��@�R�[�h�l)
    Else
        Call subSetSheet(16, wsCode)         '�V�[�g�I�u�W�F�N�g(�ʎ��@�R�[�h�l�i�m���t���[�g�j�j
    End If
    
    '���t�ۑ䐔�擾
    intMaxRow = 9
    
    '�����쐬����Ă��閾�׍s���������[�v����
    With wsMeisai
        For i = 1 To intMaxRow
            '�����ݒ���폜
            .Range(strTargetCol & (intHeadNum + i)).Validation.Delete
            '�Z�������X�g�ɕύX���A�Y���̃R�[�h�l�������̂�ݒ�
            .Range(strTargetCol & (intHeadNum + i)).Validation.Add _
                Type:=xlValidateList, _
                Formula1:="=" & wsCode.Range(wsCode.Range(strTargetCode), wsCode.Cells(wsCode.Rows.Count, _
                            wsCode.Range(strTargetCode).Column).End(xlUp)).Address(External:=True)
        Next i
    End With
    
    Set wsCode = Nothing
    
End Sub


'�Z���ی���͉\�͈͐ݒ�
Public Sub subCellProtect(ByVal intRange As Integer)
    
    Dim i        As Integer
    Dim intCol   As Integer
    Dim strRange As String
    Dim varRange As Variant
    Dim wsMeisai  As Worksheet          '���ד��̓��[�N�V�[�g
    Dim wsSetting As Worksheet          '�e��ݒ胏�[�N�V�[�g
    
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(1, wsMeisai)       '�V�[�g�I�u�W�F�N�g(���ד���)
    Else
        Call subSetSheet(17, wsMeisai)      '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
    End If
    Call subSetSheet(5, wsSetting)      '�V�[�g�I�u�W�F�N�g(�ʎ��@�e��ݒ�)
    
    strRange = ""
    intCol = 0
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        varRange = Array("$C$21:$C$21", "$E$21:$O$21", "$Y$21:$AT$21")  '���͉\�Z��"$X�͈́i�����N���X�⍇�v�ی������͓��͕s�j
    Else    '�m���t���[�g
        varRange = Array("$C$21:$C$21", "$E$21:$W$21", "$AG$21:$BD$21")
    End If
    
    '�Z���͈͐ݒ肪�c���Ă���ꍇ�A�폜
    If wsMeisai.Protection.AllowEditRanges.Count = 0 Then
    Else
        wsMeisai.Protection.AllowEditRanges.item(1).Delete
    End If
    
    '���͉\�Z���͈͂𑍕t�ۑ䐔���L����
    For i = 0 To UBound(varRange)
        intCol = Right(varRange(i), 2)
        intCol = intCol + intRange - 1
        varRange(i) = Left(varRange(i), Len(varRange(i)) - 2) & intCol
        
        strRange = strRange & "," & varRange(i)
    Next i
    
    strRange = Right(strRange, Len(strRange) - 1)
    
    '���͉\�Z���͈͂�ݒ�
    wsMeisai.Protection.AllowEditRanges.Add _
                Title:="EntryOK", _
                Range:=wsMeisai.Range(strRange)
    
    Set wsMeisai = Nothing
    Set wsSetting = Nothing
    
End Sub

'�I�u�W�F�N�g��No���擾
Public Function fncGetObjectNo(strObjectFullName As String, strObjectName As String) As String
    fncGetObjectNo = Right(strObjectFullName, Len(strObjectFullName) - Len(strObjectName))
End Function

'�J���}�폜
Public Function fncTrimComma(strVal As String) As String
    fncTrimComma = Replace(strVal, ",", "")
End Function


