Attribute VB_Name = "modCommonFunctions"
Option Explicit

'���ʊ֐�

'�O���[�o���ϐ�
Public FleetTypeFlg      As Integer     '�t���[�g�A�m���t���[�g����
Public strTxtOther       As String      '���̑�����������
Public blnSaveFlg        As Boolean     '�Ǘ��ҋ@�\�̕ۑ��t���O
Public blnCloseFlg       As Boolean     '�u�b�N�̕��閳���t���O
Public blnSyaryouOpenFlg As Boolean     '�ԗ����捞�t�@�C���̊m�F�t���O
Public blnFleetBtnFlg    As Boolean     'Top��ʁu�t���[�g�_��v�̉����t���O
Public blnNonFleetBtnFlg As Boolean     'Top��ʁu�m���t���[�g���וt�_��v�̉����t���O
Public blnChouhyouflg    As Boolean     'Top��ʁu���[�o�́v�̉����t���O
Public MeisaiBackFlg     As Integer     '���ד��͉�ʁu�߂�v�̉������@0:���[�I����� 1:���ʍ��ډ��
Public blnMoushikomiflg  As Boolean     '���׏������ʂ̉����t���O
Public strTextName       As String      '���Z�t�@�C����

Public Function fncFormSave(ByVal intForm As Integer, ByVal strSaveContent As String) As Boolean
'�֐����FfncFormSave
'���e�@�F��ʂ̏�Ԃ����[�N�V�[�g�ɕۑ�
'�����@�F
'        intForm        = 1 :���ʍ��ډ��
'                         2 :�⏞���e�Z�b�g�i�ꊇ�j���
'        strSaveContent = "":�ۑ����e
    Dim intSheetType As Integer
    Dim varSaveConetent As Variant
    Dim i As Integer

    fncFormSave = False
    i = 0

    If intForm = 1 Then
        intSheetType = 2
    Else
        intSheetType = 3
    End If
    
    Dim wstSave As Worksheet
    Call subSetSheet(intSheetType, wstSave)         '�V�[�g�I�u�W�F�N�g(�ʎ��@���ʍ���,�ʎ��@�⏞���e�Z�b�g�i�ꊇ�j)
    
    varSaveConetent = Split(strSaveContent, "/")
    
    Do While i < UBound(varSaveConetent)
        wstSave.Cells(2, 1 + i).Value = varSaveConetent(i)
        i = i + 1
    Loop
    
    Set wstSave = Nothing
    
    fncFormSave = True
    
End Function

Public Function fncFormRef(ByVal intForm As Integer, ByRef strSaveContent As String) As Boolean
'�֐����FfncFormRef
'���e�@�F���[�N�V�[�g�ɕۑ�����Ă���l���擾
'�����@�F
'        intForm        = 1 :���ʍ��ډ��
'                         2 :�⏞���e�Z�b�g�i�ꊇ�j���
'        strSaveContent = "":�ۑ����e

    Dim intSheetType As Integer
    Dim i As Integer
    Dim intLastCol As Integer
    
    fncFormRef = True
    
    i = 1
    intLastCol = 0
    
    If intForm = 1 Then
        intSheetType = 2
    Else
        intSheetType = 3
    End If
    
    Dim wstSave As Worksheet
    Call subSetSheet(intSheetType, wstSave)     '�V�[�g�I�u�W�F�N�g(�ʎ��@���ʍ���,�ʎ��@�⏞���e�Z�b�g�i�ꊇ�j)
    
    Do Until wstSave.Cells(1, i) = ""
        i = i + 1
    Loop

    intLastCol = i
        
    For i = 1 To intLastCol - 1
        If Not wstSave.Cells(2, i) = "" Then
            fncFormRef = False
        End If
    
        strSaveContent = strSaveContent + wstSave.Cells(2, i) + "/"
    Next i
    
    Set wstSave = Nothing
    
End Function
Public Function fncMoushikomiFormRef(ByVal intForm As Integer, ByRef strSaveContent As String) As Boolean
'�֐����FfncMoushikomiFormRef
'���e�@�F�\���������ʗp���[�N�V�[�g�ɕۑ�����Ă���l���擾
'�����@�F
'        intForm�@�@�@�@= 8 :�\���������ʁ@'2018/3 ���ذĖ��וt�@�\�ǉ�
'                       = 6 :�e�L�X�g���e(����)
'        strSaveContent = "":�ۑ����e

    Dim intSheetType As Integer
    Dim i As Integer
    Dim intLastCol As Integer
    
    fncMoushikomiFormRef = True
    
    i = 1
    intLastCol = 0
    
    If intForm = 8 Then
        intSheetType = 8 '2018/3 ���ذĖ��וt�@�\�ǉ�
        If FleetTypeFlg = 1 Then
            intLastCol = 29 '�t���[�g�i�\���������ʂ̍��ڐ���24�j
        Else
            intLastCol = 29 '�m���t���[�g�i�\���������ʂ̍��ڐ���29�j
        End If
        
    ElseIf intForm = 6 Then
        intSheetType = 6 '�e�L�X�g���e(����)
        intLastCol = 38 '���ʍ��ڐ���38
    End If
    
    Dim wstSave As Worksheet
    Call subSetSheet(intSheetType, wstSave)     '�V�[�g�I�u�W�F�N�g(�\����������)
        
    For i = 1 To intLastCol
        strSaveContent = strSaveContent + wstSave.Cells(1, i) + "/"
    Next i

    If strSaveContent = String(intLastCol, "/") Then
    Else
        fncMoushikomiFormRef = False
    End If
    Set wstSave = Nothing

End Function

'�R�[�h�l����
Public Function fncFindCode(ByVal strContent As String, ByVal strRow As String) As String
    Dim strFindRow As String
    Dim strFindColumn As String
    Dim rgFindNot As Range
    
    Dim wsCode As Worksheet
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(4, wsCode)         '�V�[�g�I�u�W�F�N�g(�ʎ��@�R�[�h�l)
    Else
        Call subSetSheet(16, wsCode)         '�V�[�g�I�u�W�F�N�g(�ʎ��@�R�[�h�l�i�m���t���[�g�j�j
    End If
    
    With wsCode
        
        Set rgFindNot = .Columns(strRow).Find(strContent, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True, MatchByte:=True)
        
        If rgFindNot Is Nothing Then
            fncFindCode = strContent
        Else
            strFindRow = rgFindNot.Row
            strFindColumn = rgFindNot.Column - 1
            
            fncFindCode = .Cells(Val(strFindRow), Val(strFindColumn)).Value
        End If

        If fncFindCode = "" Then
            fncFindCode = strContent
        End If

    End With
    
    Set wsCode = Nothing
    Set rgFindNot = Nothing

End Function

'�R�[�h�l�������̌���
Public Function fncFindName(ByVal strContent As String, ByVal strRow As String) As String
    Dim strFindRow      As String
    Dim strFindColumn   As String
    Dim rgFindNot      As Range

    Dim wsCode As Worksheet
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(4, wsCode)         '�V�[�g�I�u�W�F�N�g(�ʎ��@�R�[�h�l)
    Else
        Call subSetSheet(16, wsCode)         '�V�[�g�I�u�W�F�N�g(�ʎ��@�R�[�h�l�i�m���t���[�g�j�j
    End If
    
    With wsCode
        Set rgFindNot = .Columns(strRow).Find(strContent, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True, MatchByte:=True)
        
        If rgFindNot Is Nothing Then
            fncFindName = strContent
        Else
            strFindRow = .Columns(strRow).Find(strContent, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True, MatchByte:=True).Row
            strFindColumn = .Columns(strRow).Find(strContent, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True, MatchByte:=True).Column + 1
            
            fncFindName = .Cells(Val(strFindRow), Val(strFindColumn)).Value
            
        End If
        
        If fncFindName = "" Then
            fncFindName = strContent
        End If
        
    End With
    
    Set wsCode = Nothing
    
End Function


Public Function fncToWareki(ByVal strDate As String, ByVal intKeta As Integer) As String
    Dim strSaveDate   As String
    Dim intSaveDate   As Integer
    Dim strWareki     As String
    Dim strKetaErr    As String
    Dim strDateErr    As String
        
    strSaveDate = ""
    intSaveDate = 0
    strWareki = ""
    strKetaErr = ""
    strDateErr = ""
    
    fncToWareki = strDate
    
    If IsNumeric(strDate) = False Then
    Else
        strKetaErr = fncKetaCheck(strDate, 8, "=")
        
        If strKetaErr = "" Then
            strSaveDate = Format(strDate, "####/##/##")
            strDateErr = fncDateCheck(strSaveDate)
            
            If strDateErr = "" Then
                
                intSaveDate = Mid(strDate, 1, 4)
            
                If CDate(strSaveDate) >= "1912/07/30" And CDate(strSaveDate) <= "1926/12/24" Then
                    strWareki = "�吳"
                    intSaveDate = intSaveDate - 1911
                
                ElseIf CDate(strSaveDate) >= "1926/12/25" And CDate(strSaveDate) <= "1989/01/07" Then
                    strWareki = "���a"
                    intSaveDate = intSaveDate - 1925
                
                ElseIf CDate(strSaveDate) >= "1989/01/08" And CDate(strSaveDate) <= "2019/04/30" Then
                    strWareki = "����"
                    intSaveDate = intSaveDate - 1988
                ElseIf CDate(strSaveDate) >= "2019/05/01" Then
                    strWareki = "�j��"
                    intSaveDate = intSaveDate - 2018
                End If
                
                fncToWareki = strWareki & CStr(intSaveDate) & "�N" & Format(CDate(strSaveDate), "mm") & "��" & Format(CDate(strSaveDate), "dd") & "��"
                
                If intKeta = 8 Then
                    fncToWareki = strWareki & CStr(intSaveDate) & "�N" & Format(CDate(strSaveDate), "mm") & "��"
                End If
                
            Else
            End If
        Else
        End If
    End If
    
End Function



'�a�����@�ϊ�
Public Function fncToSeireki(ByVal strDate As String, ByVal intKeta As Date, Optional ByVal blnNengappiflg As Boolean, Optional ByVal blnMaezeroflg As Boolean) As String
    Dim intSaveDate    As Integer
    If strDate Like "*�N*��*��" Then
        If strDate Like "*���N*" Then
            strDate = Left(strDate, InStr(strDate, "��") - 1) & "1" & Mid(strDate, InStr(strDate, "��") + 1)
        End If
        
        intSaveDate = Mid(strDate, 3, InStr(strDate, "�N") - 3)
        
        If Left(strDate, 2) = "�吳" Then
            intSaveDate = intSaveDate + 1911
        ElseIf Left(strDate, 2) = "���a" Then
            intSaveDate = intSaveDate + 1925
        ElseIf Left(strDate, 2) = "����" Then
            intSaveDate = intSaveDate + 1988
        ElseIf Left(strDate, 2) = "�j��" Then '�V�����Ή�
            intSaveDate = intSaveDate + 2018
        End If
        
'        If intKeta = 6 Then
'            strDate = CStr(intSaveDate) & Mid(strDate, InStr(strDate, "�N")) '�V�����Ή�
'            fncToSeireki = CStr(intSaveDate) & Format(strDate, "mm")
'        Else
'            strDate = CStr(intSaveDate) & Mid(strDate, InStr(strDate, "�N")) '�V�����Ή�
'            fncToSeireki = CStr(intSaveDate) & Format(strDate, "mm") & Format(strDate, "dd")
'        End If
        If blnNengappiflg Then
            If intKeta = 8 Then
                strDate = CStr(intSaveDate) & Mid(strDate, InStr(strDate, "�N")) '�V�����Ή�
                fncToSeireki = CStr(intSaveDate) & "�N" & Format(strDate, "m") & "��"
            ElseIf intKeta = 11 Then
                If blnMaezeroflg Then
                    strDate = CStr(intSaveDate) & Mid(strDate, InStr(strDate, "�N")) '�V�����Ή�
                    fncToSeireki = CStr(intSaveDate) & "�N" & Format(strDate, "mm") & "��" & Format(strDate, "dd") & "��"
                Else
                    strDate = CStr(intSaveDate) & Mid(strDate, InStr(strDate, "�N")) '�V�����Ή�
                    fncToSeireki = CStr(intSaveDate) & "�N" & Format(strDate, "m") & "��" & Format(strDate, "d") & "��"
                End If
            End If
        Else
            If intKeta = 6 Then
                strDate = CStr(intSaveDate) & Mid(strDate, InStr(strDate, "�N")) '�V�����Ή�
                fncToSeireki = CStr(intSaveDate) & Format(strDate, "mm")
            Else
                strDate = CStr(intSaveDate) & Mid(strDate, InStr(strDate, "�N")) '�V�����Ή�
                fncToSeireki = CStr(intSaveDate) & Format(strDate, "mm") & Format(strDate, "dd")
            End If
        End If
    End If
    
End Function

Public Sub subSaveDel()
'�֐����FsubSaveDel
'���e�@�F���[�N�V�[�g�ɕۑ�����Ă���e��ʂ̏�Ԃ��폜
'�����@�F
    
    Dim i As Integer            '���[�v�J�E���g
    Dim objAll As Object     '���[�v�p�I�u�W�F�N�g
    
    Dim wsSave As Worksheet     '��ʂ̏�Ԃ��ۑ�����Ă���V�[�g
    '���ʍ���
    Call subSetSheet(2, wsSave)             '�V�[�g�I�u�W�F�N�g(�ʎ��@���ʍ���)
    wsSave.Rows("2:2").Delete
    
    '�⏞���e�Z�b�g(�ꊇ)
    Call subSetSheet(3, wsSave)         '�V�[�g�I�u�W�F�N�g(�ʎ��@�⏞���e�Z�b�g�i�ꊇ�j)
    wsSave.Rows("2:2").Delete
    
    '���ד���
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(1, wsSave)       '�V�[�g�I�u�W�F�N�g(���ד���)
    Else
        Call subSetSheet(17, wsSave)      '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
    End If
    
'    '�������(10�s)��薾�׍s�����݂���ꍇ�͍폜
'    If Val(wsSave.Range("F9")) > 10 Then
'        Call subMeisaiDel(Val(wsSave.Range("F9")) - 10, "2")
'    End If
    
    '�I����ԁE���e�̃N���A�݂̂��s���s���c��
    
    Call subMeisaiUnProtect             '�V�[�g�̕ی�̉���
    
    Application.EnableEvents = False    '�C�x���g����
    
    '���ʍ���
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        '�ی�����
        wsSave.Range("B3") = "�@�ی����ԁ@�@�@�F�@"
        '��t�敪
        wsSave.Range("E3") = "�@��t�敪�@�@�@�F�@"
        '��ی���
        wsSave.Range("G3") = "�@��ی��ҁ@�@�@�@�@�@�F�@"
        '�ی����
        wsSave.Range("B4") = "�@�ی���ށ@�@�@�F�@"
        '�t���[�g�敪
        wsSave.Range("E4") = "�@�t���[�g�敪�@�F�@"
        '�S�ԗ��ꊇ�t�ۓ���
        wsSave.Range("G4") = "�@�S�ԗ��ꊇ�t�ۓ���@�F�@"
        '�������@
        wsSave.Range("B5") = "�@�������@�@�@�@�F�@"
        '�D�Ǌ���
        wsSave.Range("E5") = "�@�D�Ǌ����@�@�@�F�@"
        '����f������
        wsSave.Range("G5") = "�@����f������  �@�@�F�@"
        '�ذđ�������
        wsSave.Range("B6") = "�@�ذđ��������@�F�@"
        '�ذĺ���
        wsSave.Range("E6") = "�@�ذĺ��ށ@�@�@�F�@"
        
        For i = 1 To Val(wsSave.OLEObjects("txtSouhuho").Object.Value)
            wsSave.Range("C" & 20 + i & ":" & "AX" & 20 + i).ClearContents  '���׍s�i�ŏI��AX��j
            Windows(ThisWorkbook.Name).ScrollColumn = 1
            Windows(ThisWorkbook.Name).ScrollRow = 1
        Next i
    Else
        '�m���t���[�g�i2018/3 ���ذĖ��וt�@�\�ǉ��j
        '�ی�����
        wsSave.Range("B3") = "�@�ی����ԁ@�@�@�F�@"
        '��t�敪
        wsSave.Range("E3") = "�@��t�敪�@�@�@�F�@"
        '��ی���
        wsSave.Range("G3") = "�@��ی��ҁ@�@�@�@�@�@�F�@"
        '�ی����
        wsSave.Range("B4") = "�@�ی���ށ@�@�@�F�@"
        '�t���[�g�敪
        wsSave.Range("E4") = "�@�t���[�g�敪�@�F�@"
        '�m���t���[�g��������
        wsSave.Range("G4") = "�@�m���t���[�g���������F�@"
        '�������@
        wsSave.Range("B5") = "�@�������@�@�@�@�F�@"
        '�c�̊�����
        wsSave.Range("E5") = "�@�c�̊������@�@�F�@"
         '����f������
        wsSave.Range("G5") = "�@"
        '�ذđ�������
        wsSave.Range("B6") = "�@"
        '�ذĺ���
        wsSave.Range("E6") = "�@"
        
        For i = 1 To Val(wsSave.OLEObjects("txtSouhuho").Object.Value)
            wsSave.Range("C" & 20 + i & ":" & "BH" & 20 + i).ClearContents  '���׍s�i�ŏI��BH��j
            Windows(ThisWorkbook.Name).ScrollColumn = 1
            Windows(ThisWorkbook.Name).ScrollRow = 1
        Next i
    End If
    
    Set objAll = Nothing
    
    '���ד��͉�ʂ̃G���[�p���X�g������
    wsSave.OLEObjects("txtErrMsg").Object.Value = ""
    '���ד��͉�ʂ̖��גǉ��e�L�X�g�{�b�N�X������
    wsSave.OLEObjects("TxtMsaiAddCnt").Object.Value = ""
    
    Set wsSave = Nothing
    
    '�`�F�b�N�{�b�N�X�̃N���A
    Call subClearAll
    
    Application.EnableEvents = True     '�C�x���g�L��

    Call subMeisaiProtect               '�V�[�g�̕ی�
    
'    '�������(10�s)��薾�׍s�������Ȃ��ꍇ�͒ǉ�
'    If wsSave.Range("F9") < 10 Then
'        Call subMeisaiAdd(10 - wsSave.Range("F9"), "2")
'    End If
    
End Sub


'�u�I���v�{�^�������i���ד��͉�ʂ̂��̑������j
Sub subClickOtherBtn()
    Dim intRow As Integer
    Dim intCol As Integer
    
    Dim wsMeisai As Worksheet           '���ד��̓��[�N�V�[�g
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(1, wsMeisai)       '�V�[�g�I�u�W�F�N�g(���ד���)
    Else
        Call subSetSheet(17, wsMeisai)      '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
    End If
    
    intRow = wsMeisai.Shapes(Application.Caller).TopLeftCell.Row
    intCol = wsMeisai.Shapes(Application.Caller).TopLeftCell.Column
    
    strTxtOther = intRow & ":" & intCol + 1
    
    Set wsMeisai = Nothing
    
    '�V�[�g�E�u�b�N�̔�\��
    Call subSheetVisible(False)
    
    frmOtherrate.Show vbModeless

End Sub

'
Sub subClickchkMeisai()
    Dim strChkValue As String
    
    Call subMeisaiUnProtect     '�V�[�g�̕ی�̉���
    
    Dim wsMeisai As Worksheet           '���ד��̓��[�N�V�[�g
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(1, wsMeisai)       '�V�[�g�I�u�W�F�N�g(���ד���)
    Else
        Call subSetSheet(17, wsMeisai)      '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
    End If
    
    strChkValue = wsMeisai.Shapes(Application.Caller).TopLeftCell.Address
    If wsMeisai.CheckBoxes(Application.Caller).Value = xlOn Then
        wsMeisai.Range(strChkValue) = Left(wsMeisai.Range(strChkValue), InStr(wsMeisai.Range(strChkValue), "/") - 1) & "/True"
    Else
        wsMeisai.Range(strChkValue) = Left(wsMeisai.Range(strChkValue), InStr(wsMeisai.Range(strChkValue), "/") - 1) & "/False"
    End If
    
    Set wsMeisai = Nothing
    
    Call subMeisaiProtect       '�V�[�g�̕ی�
    
End Sub


'�Ǘ��ҋ@�\
Function fncAdmin() As Boolean

    Dim intMsg As Integer
    intMsg = MsgBox("�ҏW���������t�@�C����ۑ����܂����B", vbYesNo, "�m�F�_�C�A���O")
    
    If intMsg = vbYes Then
        
        Call subBookUnProtect           '�u�b�N�̕ی������
        
        Application.ScreenUpdating = False                            '�`���~
        
        Application.OnKey "%{q}", ""
        ThisWorkbook.Worksheets("���ד���").Visible = True
        ThisWorkbook.Worksheets("�ʎ��@�R�[�h�l").Visible = False
        ThisWorkbook.Worksheets("�ʎ��@�e��ݒ�").Visible = False
        ThisWorkbook.Worksheets("���Ϗ�").Visible = False
        ThisWorkbook.Worksheets("�ԗ����׏�").Visible = False
        ThisWorkbook.Worksheets("�_��\����1����").Visible = False
        ThisWorkbook.Worksheets("�_��\����2����").Visible = False
        ThisWorkbook.Worksheets("���׏�").Visible = False
        ThisWorkbook.Worksheets("�\�����d�c�o").Visible = False
        ThisWorkbook.Worksheets("���׏��d�c�o").Visible = False
        ThisWorkbook.Worksheets("�ʎ��@���Ϗ��ݒ�").Visible = False
        ThisWorkbook.Worksheets("�ʎ��@�ԗ����׏��ݒ�").Visible = False
        ThisWorkbook.Worksheets("�ʎ��@�\����(1����)�ݒ�").Visible = False
        ThisWorkbook.Worksheets("�ʎ��@�\����(2����)�ݒ�").Visible = False
        ThisWorkbook.Worksheets("�ʎ��@���׏��ݒ�").Visible = False
        ThisWorkbook.Worksheets("�ʎ��@�\�����d�c�o�ݒ�").Visible = False
        ThisWorkbook.Worksheets("�ʎ��@���׏��d�c�o�ݒ�").Visible = False
        '2018/3 ���ذĖ��וt�@�\�ǉ�
        ThisWorkbook.Worksheets("���ד��́i�m���t���[�g�j").Visible = False
        ThisWorkbook.Worksheets("���׏����").Visible = False
        ThisWorkbook.Worksheets("���׏�����i�m���t���[�g�j").Visible = False
        ThisWorkbook.Worksheets("�ʎ��@�R�[�h�l�i�m���t���[�g�j").Visible = False
        ThisWorkbook.Worksheets("���Ϗ��i�m���t���[�g�j").Visible = False
        ThisWorkbook.Worksheets("�ԗ����׏��i�m���t���[�g�j").Visible = False
        ThisWorkbook.Worksheets("���׏��i�m���t���[�g�j").Visible = False
        ThisWorkbook.Worksheets("�\�����d�c�o�i�m���t���[�g�j").Visible = False
        ThisWorkbook.Worksheets("���׏��d�c�o�i�m���t���[�g�j").Visible = False
        ThisWorkbook.Worksheets("�ʎ��@���Ϗ��ݒ�i�m���t���[�g�j").Visible = False
        ThisWorkbook.Worksheets("�ʎ��@�ԗ����׏��ݒ�i�m���t���[�g�j").Visible = False
        ThisWorkbook.Worksheets("�ʎ��@�\����(1����)�ݒ�i�m���t���[�g�j").Visible = False
        ThisWorkbook.Worksheets("�ʎ��@�\����(2����)�ݒ�i�m���t���[�g�j").Visible = False
        ThisWorkbook.Worksheets("�ʎ��@���׏��ݒ�i�m���t���[�g�j").Visible = False
        ThisWorkbook.Worksheets("�ʎ��@�\�����d�c�o�ݒ�i�m���t���[�g�j").Visible = False
        ThisWorkbook.Worksheets("�ʎ��@���׏��d�c�o�ݒ�i�m���t���[�g�j").Visible = False
        ThisWorkbook.Worksheets("�e�L�X�g���e(����)").Visible = False
        ThisWorkbook.Worksheets("�e�L�X�g���e(����)").Visible = False
        
        Application.ScreenUpdating = True                            '�`��ĊJ
        
        '�u�b�N�̕ۑ�
        blnSaveFlg = True
        ThisWorkbook.Save
        
        Call subMeisaiUnProtect     '�V�[�g�̕ی�̉���
        Call subMeisaiPrtUnProtect  '�V�[�g�̕ی�̉��� '2018/3 ���ذĖ��וt�@�\�ǉ�
        Call subCmbInitialize       '�v�Z�p�V�[�g�̃R���{�{�b�N�X��ݒ�
        Call subMeisaiProtect       '�V�[�g�̕ی�
        Call subMeisaiPrtProtect    '�V�[�g�̕ی� '2018/3 ���ذĖ��וt�@�\�ǉ�
        
        Call subSheetVisible(False)     '�V�[�g�E�u�b�N�̔�\��
        Call subBookProtect             '�u�b�N�̕ی�
        
        frmTop.Show vbModeless
    End If

End Function


Public Sub subSetSheet(ByVal intSheetType As Integer, ByRef wsSheet As Worksheet)
    '�֐����FsubSetSheet
'���e�@�F�V�[�g�I�u�W�F�N�g�ɒl��ݒ�
'�����@�F
'        intSheetType   =  1 :���ד���
'                          2 :�ʎ��@���ʍ���
'                          3 :�ʎ��@�⏞���e�Z�b�g�i�ꊇ�j
'                          4 :�ʎ��@�R�[�h�l
'                          5 :�ʎ��@�e��ݒ�
'                          6 :�e�L�X�g���e(����)
'                          7 :�e�L�X�g���e(����)
'                          8 :�\���������ʓ��e
'                          9 :�ʎ��@���Ϗ��ݒ�
'                         10 :�ʎ��@�ԗ����׏��ݒ�
'                         11 :�ʎ��@�\����(1����)�ݒ�
'                         12 :�ʎ��@�\����(2����)�ݒ�
'                         13 :�ʎ��@���׏��ݒ�
'                         14 :�ʎ��@�\�����d�c�o�ݒ�
'                         15 :�ʎ��@���׏��d�c�o�ݒ�
'                         16 :�ʎ��@�R�[�h�l�i�m���t���[�g�j�@�@        '2018/3 ���ذĖ��וt�@�\�ǉ�
'                         17 :���ד��́i�m���t���[�g�j�@�@�@�@�@        '2018/3 ���ذĖ��וt�@�\�ǉ�
'                         18 :���׏�����@�@�@�@�@�@�@�@�@�@�@�@        '2018/3 ���ذĖ��וt�@�\�ǉ�
'                         19 :���׏�����i�m���t���[�g�j�@�@�@�@        '2018/3 ���ذĖ��וt�@�\�ǉ�
'                         20 :�ʎ��@���Ϗ��ݒ�i�m���t���[�g�j          '2018/3 ���ذĖ��וt�@�\�ǉ�
'                         21 :�ʎ��@�ԗ����׏��ݒ�i�m���t���[�g�j      '2018/3 ���ذĖ��וt�@�\�ǉ�
'                         22 :�ʎ��@�\����(1����)�ݒ�i�m���t���[�g�j   '2018/3 ���ذĖ��וt�@�\�ǉ�
'                         23 :�ʎ��@�\����(2����)�ݒ�i�m���t���[�g�j   '2018/3 ���ذĖ��וt�@�\�ǉ�
'                         24 :�ʎ��@���׏��ݒ�i�m���t���[�g�j          '2018/3 ���ذĖ��וt�@�\�ǉ�
'                         25 :�ʎ��@�\�����d�c�o�ݒ�i�m���t���[�g�j    '2018/3 ���ذĖ��וt�@�\�ǉ�
'                         26 :�ʎ��@���׏��d�c�o�ݒ�i�m���t���[�g�j    '2018/3 ���ذĖ��וt�@�\�ǉ�
'
'                        101 :���Ϗ�WK
'                        102 :�ԗ����׏�WK
'                        103 :�_��\����1����WK
'                        104 :�_��\����2����WK
'                        105 :���׏�WK
'                        106 :�\�����d�c�oWK
'                        107 :���׏��d�c�oWK
'        wsSheet        = "":�V�[�g�I�u�W�F�N�g�i�m���t���[�g�j
    Select Case intSheetType
        Case 1
            Set wsSheet = ThisWorkbook.Worksheets("���ד���")
        Case 2
            Set wsSheet = ThisWorkbook.Worksheets("�ʎ��@���ʍ���")
        Case 3
            Set wsSheet = ThisWorkbook.Worksheets("�ʎ��@�⏞���e�Z�b�g�i�ꊇ�j")
        Case 4
            Set wsSheet = ThisWorkbook.Worksheets("�ʎ��@�R�[�h�l")
        Case 5
            Set wsSheet = ThisWorkbook.Worksheets("�ʎ��@�e��ݒ�")
        Case 6
            Set wsSheet = ThisWorkbook.Worksheets("�e�L�X�g���e(����)")
        Case 7
            Set wsSheet = ThisWorkbook.Worksheets("�e�L�X�g���e(����)")
        Case 8
            Set wsSheet = ThisWorkbook.Worksheets("�\���������ʓ��e")
        Case 9
            Set wsSheet = ThisWorkbook.Worksheets("�ʎ��@���Ϗ��ݒ�")
        Case 10
            Set wsSheet = ThisWorkbook.Worksheets("�ʎ��@�ԗ����׏��ݒ�")
        Case 11
            Set wsSheet = ThisWorkbook.Worksheets("�ʎ��@�\����(1����)�ݒ�")
        Case 12
            Set wsSheet = ThisWorkbook.Worksheets("�ʎ��@�\����(2����)�ݒ�")
        Case 13
            Set wsSheet = ThisWorkbook.Worksheets("�ʎ��@���׏��ݒ�")
        Case 14
            Set wsSheet = ThisWorkbook.Worksheets("�ʎ��@�\�����d�c�o�ݒ�")
        Case 15
            Set wsSheet = ThisWorkbook.Worksheets("�ʎ��@���׏��d�c�o�ݒ�")
        Case 16
            Set wsSheet = ThisWorkbook.Worksheets("�ʎ��@�R�[�h�l�i�m���t���[�g�j")         '2018/3 ���ذĖ��וt�@�\�ǉ�
        Case 17
            Set wsSheet = ThisWorkbook.Worksheets("���ד��́i�m���t���[�g�j")               '2018/3 ���ذĖ��וt�@�\�ǉ�
        Case 18
            Set wsSheet = ThisWorkbook.Worksheets("���׏����")                             '2018/3 ���ذĖ��וt�@�\�ǉ�
        Case 19
            Set wsSheet = ThisWorkbook.Worksheets("���׏�����i�m���t���[�g�j")             '2018/3 ���ذĖ��וt�@�\�ǉ�
        Case 20
            Set wsSheet = ThisWorkbook.Worksheets("�ʎ��@���Ϗ��ݒ�i�m���t���[�g�j")       '2018/3 ���ذĖ��וt�@�\�ǉ�
        Case 21
            Set wsSheet = ThisWorkbook.Worksheets("�ʎ��@�ԗ����׏��ݒ�i�m���t���[�g�j")   '2018/3 ���ذĖ��וt�@�\�ǉ�
        Case 22
            Set wsSheet = ThisWorkbook.Worksheets("�ʎ��@�\����(1����)�ݒ�i�m���t���[�g�j") '2018/3 ���ذĖ��וt�@�\�ǉ�
        Case 23
            Set wsSheet = ThisWorkbook.Worksheets("�ʎ��@�\����(2����)�ݒ�i�m���t���[�g�j") '2018/3 ���ذĖ��וt�@�\�ǉ�
        Case 24
            Set wsSheet = ThisWorkbook.Worksheets("�ʎ��@���׏��ݒ�i�m���t���[�g�j")       '2018/3 ���ذĖ��וt�@�\�ǉ�
        Case 25
            Set wsSheet = ThisWorkbook.Worksheets("�ʎ��@�\�����d�c�o�ݒ�i�m���t���[�g�j") '2018/3 ���ذĖ��וt�@�\�ǉ�
        Case 26
            Set wsSheet = ThisWorkbook.Worksheets("�ʎ��@���׏��d�c�o�ݒ�i�m���t���[�g�j") '2018/3 ���ذĖ��וt�@�\�ǉ�
        Case 101
            Set wsSheet = ThisWorkbook.Worksheets("���Ϗ�WK")
        Case 102
            Set wsSheet = ThisWorkbook.Worksheets("�ԗ����׏�WK")
        Case 103
            Set wsSheet = ThisWorkbook.Worksheets("�_��\����1����WK")
        Case 104
            Set wsSheet = ThisWorkbook.Worksheets("�_��\����2����WK")
        Case 105
            Set wsSheet = ThisWorkbook.Worksheets("���׏�WK")
        Case 106
            Set wsSheet = ThisWorkbook.Worksheets("�\�����d�c�oWK")
        Case 107
            Set wsSheet = ThisWorkbook.Worksheets("���׏��d�c�oWK")
    End Select
End Sub

Public Sub subSheetVisible(ByVal blnVisibleMode As Boolean)
    Dim blnOtherBookFlg As Boolean
    Dim sOpenBookSub    As Variant
    Dim wsMeisai        As Worksheet    '���ד��̓��[�N�V�[�g
    Dim sOpenbookAll    As Workbook
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    Dim wsMeisaiNonfleet      As Worksheet    '���ד��́i�m���t���[�g�j���[�N�V�[�g
    Dim wsMeisaiPrint         As Worksheet    '���׏������ʃ��[�N�V�[�g
    Dim wsMeisaiPrintNonfleet As Worksheet    '���׏������ʁi�m���t���[�g�j���[�N�V�[�g
    
    
    Call subSetSheet(1, wsMeisai)               '�V�[�g�I�u�W�F�N�g(���ד���)
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    Call subSetSheet(17, wsMeisaiNonfleet)      '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
    Call subSetSheet(18, wsMeisaiPrint)         '�V�[�g�I�u�W�F�N�g(���׏����)
    Call subSetSheet(19, wsMeisaiPrintNonfleet) '�V�[�g�I�u�W�F�N�g(���׏�����i�m���t���[�g�j)
    
    
    If blnVisibleMode Then
        '�\��
        
        '���Ƀu�b�N���J���Ă��Ȃ����m�F
        For Each sOpenbookAll In Workbooks
            If sOpenbookAll.Name <> ThisWorkbook.Name Then
                blnOtherBookFlg = True
                Exit For
            End If
        Next sOpenbookAll
        
        If blnOtherBookFlg Then
            '���Ƀu�b�N���J����Ă���ꍇ�A�V�[�g��\��
            Windows(ThisWorkbook.Name).Visible = True
        Else
            '���Ƀu�b�N���J����Ă��Ȃ��ꍇ�A�G�N�Z�����ƕ\��
            Application.Visible = True
            Windows(ThisWorkbook.Name).Visible = True
        End If
        
        '2018/3 ���ذĖ��וt�@�\�ǉ�
        If blnMoushikomiflg Then
            '���׏�������
            If FleetTypeFlg = 1 Then
                '�t���[�g
                wsMeisaiPrint.Visible = True
                
                wsMeisai.Visible = False
                wsMeisaiNonfleet.Visible = False
                wsMeisaiPrintNonfleet.Visible = False
                
                wsMeisaiPrint.Activate
            Else
                '�m���t���[�g
                wsMeisaiPrintNonfleet.Visible = True
                
                wsMeisai.Visible = False
                wsMeisaiNonfleet.Visible = False
                wsMeisaiPrint.Visible = False
                
                wsMeisaiPrintNonfleet.Activate
            End If
        Else
            '���ד��͉��
            If FleetTypeFlg = 1 Then
                '�t���[�g
                wsMeisai.Visible = True
                
                wsMeisaiNonfleet.Visible = False
                wsMeisaiPrint.Visible = False
                wsMeisaiPrintNonfleet.Visible = False
                
                wsMeisai.Activate
            Else
                '�m���t���[�g
                wsMeisaiNonfleet.Visible = True
                
                wsMeisai.Visible = False
                wsMeisaiPrint.Visible = False
                wsMeisaiPrintNonfleet.Visible = False
                
                wsMeisaiNonfleet.Activate
            End If
        End If
        
    Else
        '��\��
        
        '���Ƀu�b�N���J���Ă��Ȃ����m�F
        For Each sOpenbookAll In Workbooks
            If sOpenbookAll.Name <> ThisWorkbook.Name Then
                blnOtherBookFlg = True
                Exit For
            End If
        Next sOpenbookAll
        
        If blnOtherBookFlg Then
            '���Ƀu�b�N���J����Ă���ꍇ�A�V�[�g���\��
            Windows(ThisWorkbook.Name).Visible = False
        Else
            '���Ƀu�b�N���J����Ă��Ȃ��ꍇ�A�G�N�Z�����Ɣ�\��
            Windows(ThisWorkbook.Name).Visible = False
            Application.Visible = False
        End If
    End If
    
    Set sOpenbookAll = Nothing
    Set sOpenBookSub = Nothing
    
End Sub

'�u�b�N�̕ی�
Public Sub subBookProtect()
    Dim wsSetting As Worksheet          '�e��ݒ胏�[�N�V�[�g
    Call subSetSheet(5, wsSetting)      '�V�[�g�I�u�W�F�N�g(�ʎ��@�e��ݒ�)
    
    ThisWorkbook.Protect Password:=wsSetting.Range("B4").Value    '�u�b�N�̕ی�
    
    Set wsSetting = Nothing
    
End Sub

'�u�b�N�̕ی������
Public Sub subBookUnProtect()
    Dim wsSetting As Worksheet          '�e��ݒ胏�[�N�V�[�g
    Call subSetSheet(5, wsSetting)      '�V�[�g�I�u�W�F�N�g(�ʎ��@�e��ݒ�)
    
    ThisWorkbook.Unprotect Password:=wsSetting.Range("B4").Value    '�u�b�N�̕ی������
    
    Set wsSetting = Nothing
    
End Sub

'�V�[�g�̕ی�
Public Sub subMeisaiProtect()
    Dim wsMeisaiFleet  As Worksheet           '���ד��̓��[�N�V�[�g
    Dim wsMeisaiNonfleet  As Worksheet        '���ד��̓��[�N�V�[�g
    Dim wsSetting As Worksheet                '�e��ݒ胏�[�N�V�[�g
    
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 0 Then
        Call subSetSheet(1, wsMeisaiFleet)        '�V�[�g�I�u�W�F�N�g(���ד���)
        Call subSetSheet(17, wsMeisaiNonfleet)    '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
    ElseIf FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(1, wsMeisaiFleet)        '�V�[�g�I�u�W�F�N�g(���ד���)
    Else
        Call subSetSheet(17, wsMeisaiNonfleet)    '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
    End If
    Call subSetSheet(5, wsSetting)                '�V�[�g�I�u�W�F�N�g(�ʎ��@�e��ݒ�)
    
    '�V�[�g�̕ی������
    If Not wsMeisaiFleet Is Nothing Then
        wsMeisaiFleet.Protect Password:=wsSetting.Range("B4").Value
    End If
    If Not wsMeisaiNonfleet Is Nothing Then
        wsMeisaiNonfleet.Protect Password:=wsSetting.Range("B4").Value
    End If
    
    Set wsMeisaiFleet = Nothing
    Set wsMeisaiNonfleet = Nothing
    Set wsSetting = Nothing
    
End Sub

'�V�[�g�̕ی������
Public Sub subMeisaiUnProtect()
    Dim wsMeisaiFleet  As Worksheet           '���ד��̓��[�N�V�[�g
    Dim wsMeisaiNonfleet  As Worksheet           '���ד��̓��[�N�V�[�g
    Dim wsSetting As Worksheet          '�e��ݒ胏�[�N�V�[�g
    
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 0 Then
        Call subSetSheet(1, wsMeisaiFleet)       '�V�[�g�I�u�W�F�N�g(���ד���)
        Call subSetSheet(17, wsMeisaiNonfleet)      '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
    ElseIf FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(1, wsMeisaiFleet)       '�V�[�g�I�u�W�F�N�g(���ד���)
    Else
        Call subSetSheet(17, wsMeisaiNonfleet)      '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
    End If
    Call subSetSheet(5, wsSetting)      '�V�[�g�I�u�W�F�N�g(�ʎ��@�e��ݒ�)
    
    '�V�[�g�̕ی������
    If Not wsMeisaiFleet Is Nothing Then
        wsMeisaiFleet.Unprotect Password:=wsSetting.Range("B4").Value
    End If
    If Not wsMeisaiNonfleet Is Nothing Then
        wsMeisaiNonfleet.Unprotect Password:=wsSetting.Range("B4").Value
    End If
    
    Set wsMeisaiFleet = Nothing
    Set wsMeisaiNonfleet = Nothing
    Set wsSetting = Nothing
    
End Sub


'2018/3 ���ذĖ��וt�@�\�ǉ�
'���׏�����V�[�g�̕ی�
Public Sub subMeisaiPrtProtect()
    Dim wsMeisaiFleet As Worksheet      '���׏�������[�N�V�[�g�i�t���[�g�j
    Dim wsMeisaiNonfleet  As Worksheet  '���׏�������[�N�V�[�g�i�m���t���[�g�j
    Dim wsSetting As Worksheet          '�e��ݒ胏�[�N�V�[�g
    
    If FleetTypeFlg = 0 Then
        '�g�b�v��ʓ�
        Call subSetSheet(18, wsMeisaiFleet)    '�V�[�g�I�u�W�F�N�g(���׏����)
        Call subSetSheet(19, wsMeisaiNonfleet) '�V�[�g�I�u�W�F�N�g(���׏�����i�m���t���[�g�j)
    ElseIf FleetTypeFlg = 1 Then
        '�t���[�g
        Call subSetSheet(18, wsMeisaiFleet)    '�V�[�g�I�u�W�F�N�g(���׏����)
    Else
        '�m���t���[�g
        Call subSetSheet(19, wsMeisaiNonfleet) '�V�[�g�I�u�W�F�N�g(���׏�����i�m���t���[�g�j)
    End If
    Call subSetSheet(5, wsSetting)             '�V�[�g�I�u�W�F�N�g(�ʎ��@�e��ݒ�)
    
    '�V�[�g�̕ی������
    If Not wsMeisaiFleet Is Nothing Then
        wsMeisaiFleet.Protect Password:=wsSetting.Range("B4").Value
    End If
    If Not wsMeisaiNonfleet Is Nothing Then
        wsMeisaiNonfleet.Protect Password:=wsSetting.Range("B4").Value
    End If
    
    Set wsMeisaiFleet = Nothing
    Set wsMeisaiNonfleet = Nothing
    Set wsSetting = Nothing
    
End Sub


'2018/3 ���ذĖ��וt�@�\�ǉ�
'���׏�����V�[�g�̕ی������
Public Sub subMeisaiPrtUnProtect()
    Dim wsMeisaiFleet  As Worksheet      '���׏�������[�N�V�[�g�i�t���[�g�j
    Dim wsMeisaiNonfleet  As Worksheet   '���׏�������[�N�V�[�g�i�m���t���[�g�j
    Dim wsSetting As Worksheet           '�e��ݒ胏�[�N�V�[�g
    
    If FleetTypeFlg = 0 Then
        '�g�b�v��ʓ�
        Call subSetSheet(18, wsMeisaiFleet)    '�V�[�g�I�u�W�F�N�g(���׏����)
        Call subSetSheet(19, wsMeisaiNonfleet) '�V�[�g�I�u�W�F�N�g(���׏�����i�m���t���[�g�j)
    ElseIf FleetTypeFlg = 1 Then
        '�t���[�g
        Call subSetSheet(18, wsMeisaiFleet)    '�V�[�g�I�u�W�F�N�g(���׏����)
    Else
        '�m���t���[�g
        Call subSetSheet(19, wsMeisaiNonfleet) '�V�[�g�I�u�W�F�N�g(���׏�����i�m���t���[�g�j)
    End If
    Call subSetSheet(5, wsSetting)             '�V�[�g�I�u�W�F�N�g(�ʎ��@�e��ݒ�)
    
    
    '�V�[�g�̕ی������
    If Not wsMeisaiFleet Is Nothing Then
        wsMeisaiFleet.Unprotect Password:=wsSetting.Range("B4").Value
    End If
    If Not wsMeisaiNonfleet Is Nothing Then
        wsMeisaiNonfleet.Unprotect Password:=wsSetting.Range("B4").Value
    End If
    
    Set wsMeisaiFleet = Nothing
    Set wsMeisaiNonfleet = Nothing
    Set wsSetting = Nothing
    
End Sub


Sub subShortCutKey(ByVal intFlg As Integer)
'intFlg = 1 '�V���[�g�J�b�g�L�[����
'intFlg = 2 '�V���[�g�J�b�g�L�[�L��

    If intFlg = 1 Then
    '����
        Application.OnKey "^6", ""              '�I�u�W�F�N�g�̕\���E��\��
        Application.OnKey "^+@", ""             '���[�N�V�[�g�̃Z���̒l�̕\���ƁA�����̕\����؂�ւ�
        Application.OnKey "%{F8}", ""           '[�}�N��]�_�C�A���O�{�b�N�X�̕\��
        Application.OnKey "%{F11}", ""          'VBA�G�f�B�^�[�̋N��
    Else
    '�L��
        Application.OnKey "^6"                  '�I�u�W�F�N�g�̕\���E��\��
        Application.OnKey "^+@"                 '���[�N�V�[�g�̃Z���̒l�̕\���ƁA�����̕\����؂�ւ�
        Application.OnKey "%{F8}"               '[�}�N��]�_�C�A���O�{�b�N�X�̕\��
        Application.OnKey "%{F11}"              'VBA�G�f�B�^�[�̋N��
    End If
    
End Sub

Function fncFormatDigit(ByVal str As String, _
                     ByVal strChar As String, _
                     ByVal lngdigit As Long) As String
'�@�\�F�w�蕶�����ߊ֐�
'�����Fstr�@�F�ϊ��O�̕�����
'�@�@�@chr  �F���߂镶��(�P�����ڂ̂ݎg�p)
'�@�@�@digit�F����
'�ߒl�F�w�蕶�����ߌ�̕�����
    
    Dim strtmp As String
    strtmp = str
    If Len(str) < lngdigit And Len(strChar) > 0 Then
      strtmp = Right(String(lngdigit, strChar) & str, lngdigit)
    End If
    fncFormatDigit = strtmp

End Function

Private Sub subOtherClose()
    Dim USF As UserForm
    
    For Each USF In UserForms
        If TypeOf USF Is frmTop Then
            If frmTop.Visible Then
                Call subBookUnProtect           '�u�b�N�̕ی������ 2018/3 ���ذĖ��וt�@�\�ǉ�
                Call subSheetVisible(True)      '�V�[�g�E�u�b�N�̕\��
                Call subBookProtect             '�u�b�N�̕ی� 2018/3 ���ذĖ��וt�@�\�ǉ�
                Call subSheetVisible(False)     '�V�[�g�E�u�b�N�̔�\��
            End If
        End If
        If TypeOf USF Is frmKyoutsuu Then
            If frmKyoutsuu.Visible Then
                Call subBookUnProtect           '�u�b�N�̕ی������ 2018/3 ���ذĖ��וt�@�\�ǉ�
                Call subSheetVisible(True)      '�V�[�g�E�u�b�N�̕\��
                Call subBookProtect             '�u�b�N�̕ی� 2018/3 ���ذĖ��וt�@�\�ǉ�
                Call subSheetVisible(False)     '�V�[�g�E�u�b�N�̔�\��
            End If
        End If
        If TypeOf USF Is frmOtherrate Then
            If frmOtherrate.Visible Then
                Call subBookUnProtect           '�u�b�N�̕ی������ 2018/3 ���ذĖ��וt�@�\�ǉ�
                Call subSheetVisible(True)      '�V�[�g�E�u�b�N�̕\��
                Call subBookProtect             '�u�b�N�̕ی� 2018/3 ���ذĖ��וt�@�\�ǉ�
                Call subSheetVisible(False)     '�V�[�g�E�u�b�N�̔�\��
            End If
        End If
        If TypeOf USF Is frmSyaryou Then
            If frmSyaryou.Visible Then
                Call subBookUnProtect           '�u�b�N�̕ی������ 2018/3 ���ذĖ��וt�@�\�ǉ�
                Call subSheetVisible(True)      '�V�[�g�E�u�b�N�̕\��
                Call subBookProtect             '�u�b�N�̕ی� 2018/3 ���ذĖ��וt�@�\�ǉ�
                Call subSheetVisible(False)     '�V�[�g�E�u�b�N�̔�\��
            End If
        End If
        If TypeOf USF Is frmHosyoSet Then
            If frmHosyoSet.Visible Then
                Call subBookUnProtect           '�u�b�N�̕ی������ 2018/3 ���ذĖ��וt�@�\�ǉ�
                Call subSheetVisible(True)      '�V�[�g�E�u�b�N�̕\��
                Call subBookProtect             '�u�b�N�̕ی� 2018/3 ���ذĖ��וt�@�\�ǉ�
                Call subSheetVisible(False)     '�V�[�g�E�u�b�N�̔�\��
            End If
        End If
        If TypeOf USF Is frmPrintMenu Then
            If frmPrintMenu.Visible Then
                Call subBookUnProtect           '�u�b�N�̕ی������ 2018/3 ���ذĖ��וt�@�\�ǉ�
                Call subSheetVisible(True)      '�V�[�g�E�u�b�N�̕\��
                Call subBookProtect             '�u�b�N�̕ی� 2018/3 ���ذĖ��וt�@�\�ǉ�
                Call subSheetVisible(False)     '�V�[�g�E�u�b�N�̔�\��
            End If
        End If
        If TypeOf USF Is frmEntryMitsumori Then
            If frmEntryMitsumori.Visible Then
                Call subBookUnProtect           '�u�b�N�̕ی������ 2018/3 ���ذĖ��וt�@�\�ǉ�
                Call subSheetVisible(True)      '�V�[�g�E�u�b�N�̕\��
                Call subBookProtect             '�u�b�N�̕ی� 2018/3 ���ذĖ��וt�@�\�ǉ�
                Call subSheetVisible(False)     '�V�[�g�E�u�b�N�̔�\��
            End If
        End If
        If TypeOf USF Is frmEntryMoushikomi Then
            If frmEntryMoushikomi.Visible Then
                Call subBookUnProtect           '�u�b�N�̕ی������ 2018/3 ���ذĖ��וt�@�\�ǉ�
                Call subSheetVisible(True)      '�V�[�g�E�u�b�N�̕\��
                Call subBookProtect             '�u�b�N�̕ی� 2018/3 ���ذĖ��וt�@�\�ǉ�
                Call subSheetVisible(False)     '�V�[�g�E�u�b�N�̔�\��
            End If
        End If
    Next
    
End Sub

Public Sub subAppClose()
    Dim wsSetting As Worksheet          '�e��ݒ胏�[�N�V�[�g
    Call subSetSheet(5, wsSetting)      '�V�[�g�I�u�W�F�N�g(�ʎ��@�e��ݒ�)
    wsSetting.Range("P1") = "True"

    Set wsSetting = Nothing
    
    '�V���[�g�J�b�g�L�[�̗L��
    Call subShortCutKey(2)

    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    Application.DisplayStatusBar = True
    Application.DisplayFormulaBar = True
    Application.CommandBars("Ply").Enabled = True

    If Workbooks.Count = 1 Then
        blnCloseFlg = True
    End If

    'Auto_Close�ɂău�b�N�����
    ThisWorkbook.RunAutoMacros Which:=xlAutoClose
        
End Sub

'�I�[�g�N���[�Y�C�x���g
Private Sub Auto_Close()
    Dim wsSetting As Worksheet
    Call subSetSheet(5, wsSetting)      '�V�[�g�I�u�W�F�N�g(�ʎ��@�e��ݒ�)
    
    If CBool(wsSetting.Range("P1")) Then
        Application.DisplayAlerts = False
        Application.Cursor = xlWait
        Application.OnTime Now, "my_Procedure"
    End If
End Sub
Private Sub my_Procedure()
    Application.Cursor = xlDefault
    If Workbooks.Count = 1 Then
        Application.Quit
    Else
        ThisWorkbook.Close SaveChanges:=False
    End If
End Sub

'���ד��̓V�[�g�ɔ��f
Public Sub fncMeisaiEntry(ByVal varContent As Variant, ByVal intCar As Integer)
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim intActRow       As Integer
    Dim intSouhuho      As Integer
    Dim varMeisaiRow    As Variant
    Dim varMeisaiCell   As Variant
    Dim varMeisaiRgCell As Variant
    
    intActRow = 0
    
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then
        '�t���[�g �i���׍��ڐ�47�j
        ReDim varMeisaiRgCell(intCar - 1, 47)
    Else
        '�m���t���[�g�i���׍��ڐ�59�j
        ReDim varMeisaiRgCell(intCar - 1, 59)
    End If
    
    Dim wstMeisai As Worksheet
    Dim wstTextM As Worksheet
    
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(1, wstMeisai)       '�V�[�g�I�u�W�F�N�g(���ד���)
        Call subSetSheet(7, wstTextM)          '�V�[�g�I�u�W�F�N�g(�e�L�X�g���e)
    Else
        Call subSetSheet(17, wstMeisai)       '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
        Call subSetSheet(7, wstTextM)          '�V�[�g�I�u�W�F�N�g(�e�L�X�g���e)
    End If
    
    Dim objSouhuho As Object
    Set objSouhuho = wstMeisai.OLEObjects("txtSouhuho").Object
    intSouhuho = Val(Left(objSouhuho.Value, Len(objSouhuho.Value) - 2))

    Call subBookUnProtect '�u�b�N�̕ی������ 2018/3 ���ذĖ��וt�@�\�ǉ�
    Call subSheetVisible(True)      '�V�[�g�E�u�b�N�̕\��
    Call subBookProtect   '�u�b�N�̕ی� 2018/3 ���ذĖ��וt�@�\�ǉ�

    ''���׍s�̒ǉ��E�폜
    '���׍s�� = ���׍s��(�f�t�H���g:10�s)
    If intCar = intSouhuho Then
    '���׍s�� < ���׍s��-�]�蕪�̖��׍s���폜
    ElseIf intCar < intSouhuho Then
        Call subMeisaiDel(intSouhuho - intCar, "2")
    '���׍s�� > ���׍s��-�s�����̖��׍s��ǉ�
    ElseIf intCar > intSouhuho Then
        Call subMeisaiAdd(intCar - intSouhuho, "2")
    End If
    Call subSheetVisible(False)      '�V�[�g�E�u�b�N�̔�\��
    
    ''���׍s�̃V�[�g�����ɔ��f
    For i = 0 To UBound(varContent, 1)
        For j = 1 To UBound(varContent, 2)
            If UBound(Split(varContent(i, j))) = -1 Then Exit For
            
            varMeisaiRow = Split(varContent(i, j), ",")
            
            '���׃V�[�g�p�z��쐬
            '2018/3 ���ذĖ��וt�@�\�ǉ�
            If FleetTypeFlg = 1 Then  '�t���[�g
                varMeisaiCell = fncMeisaiSetCell(varMeisaiRow)
            Else
                varMeisaiCell = fncNonFleetMeisaiSetCell(varMeisaiRow)
            End If
            
            '2�����z���1�����ڂ�1�����z����R�s�[
            Call subMeisaiArray(varMeisaiRgCell, varMeisaiCell, intActRow)
            
'            If blnChouhyouflg Then
                For k = 1 To UBound(varMeisaiRow) + 1
                    wstTextM.Cells(intActRow + 1, k) = varMeisaiRow(k - 1)
                    Call fncTextEdit(2, k, varMeisaiRow(k - 1), intActRow + 1)
                Next k
'            End If
            
            intActRow = intActRow + 1
            
        Next j
    Next i
    
    '���ד���(Cell)�ɓ\��t��
    Call subMeisaiCell(varMeisaiRgCell)
    
    Set wstMeisai = Nothing
    Set wstTextM = Nothing
    Set objSouhuho = Nothing
    Set varMeisaiRow = Nothing
    Set varMeisaiCell = Nothing
    Set varMeisaiRgCell = Nothing
        
End Sub

'���׃V�[�g�p�z��쐬�i�Z���j
Function fncMeisaiSetCell(ByVal varContent As Variant) As Variant
    Dim i           As Integer
    Dim strText     As String
    Dim strSaveDate As String
    Dim varCell(47) As Variant           'Cell�p�z��
    
    strText = ""
    strSaveDate = ""
    
    '�V�[�g�̃Z���Ɋi�[
    '--------C��`L��--------
    varCell(0) = fncFindCode(varContent(1), "AA")             '�p�r�Ԏ햼
    varCell(1) = varContent(1)                                '�p�r�Ԏ�R�[�h
    varCell(2) = varContent(2)                                '�Ԗ�
    varCell(3) = varContent(68)                               '�o�^�ԍ�
    varCell(4) = varContent(69)                               '�ԑ�ԍ�
    varCell(5) = varContent(3)                                '�^��
    varCell(6) = varContent(4)                                '�d�l
    '���x�o�^�N��
    strSaveDate = fncToWareki(varContent(5) & "25", 8)
    If strSaveDate = varContent(5) & "25" Then
        varCell(7) = varContent(5)
    Else
        varCell(7) = strSaveDate
    End If
    '�Ԍ�������
    strSaveDate = fncToWareki(CStr(varContent(70)), 11)
    varCell(8) = strSaveDate
    '�����E�s����
    varCell(9) = fncFindCode(varContent(6), "AE")
    
    '--------M��`U��--------
    varCell(10) = varContent(7)                                '�r�C��
    varCell(11) = varContent(8)                                '2.5���b�g���z�f�B�[�[��������
    varCell(12) = IIf(varContent(72) = "1", "�K�p����", varContent(72)) 'ASV����
    If blnChouhyouflg = False Then
        varCell(13) = ""                                       '����-�ԗ�
        varCell(14) = ""                                       '����-�ΐl
        varCell(15) = ""                                       '����-�Ε�
        varCell(16) = ""                                       '����-���Q
        varCell(17) = ""                                       '�V�Ԋ���
        varCell(18) = ""                                       '�ԗ����
        varCell(19) = ""                                       '�ԗ�����
    Else
        varCell(13) = varContent(23)                           '����-�ԗ�
        varCell(14) = varContent(24)                           '����-�ΐl
        varCell(15) = varContent(25)                           '����-�Ε�
        varCell(16) = varContent(26)                           '����-���Q
        varCell(17) = varContent(27)                           '�V�Ԋ���
        varCell(18) = varContent(30)                           '�ԗ����
        varCell(19) = varContent(29)                           '�ԗ�����
    End If
    varCell(20) = ""                                           '���̑�����-�{�^��
    
    '--------W��`AF��--------
    strText = IIf(varContent(16) = "3 ", "����^", "")
    strText = strText + IIf(varContent(17) = "1", "�����^�J�[�^", "")
    strText = strText + IIf(varContent(18) = "5 ", "���K�ԁ^", "")
    strText = strText + IIf(varContent(19) = "1 ", "�u�[���ΏۊO�^", "")
    strText = strText + IIf(varContent(20) = "80", "���[�X�J�[�I�[�v���|���V�[�^", "")
    strText = strText + IIf(varContent(21) = "93", "�I�[�v���|���V�[���������^", "")
    strText = strText + IIf(varContent(22) = "1 ", "���L�^", _
                        IIf(varContent(22) = "2 ", "�����L�^", ""))
    If blnChouhyouflg = False Then
        strText = strText + ""
    Else
        strText = strText + IIf(varContent(28) = "8", "����敪�^", "")
    End If
    If strText <> "" Then
        varCell(21) = Left(strText, Len(strText) - 1)          '���̑�����-�e�L�X�g
    End If
    varCell(22) = fncFindCode(varContent(38), "BK")            '�ԗ��ی��̎��
    varCell(23) = varContent(39)                               '�ԗ��ی����z
    varCell(24) = fncFindCode(varContent(40), "BO")            '�ԗ��ƐӋ��z
    varCell(25) = IIf(varContent(42) = "2", _
                        "�K�p����", varContent(42))            '�ԗ��S���Ք����
    varCell(26) = IIf(varContent(44) = "1", _
                        "�K�p����", varContent(44))            '�ԗ����ߏC����p����
    varCell(27) = IIf(varContent(43) = "1", _
                        "�ΏۊO", varContent(43))              '�ԗ�����ΏۊO����
    If varContent(45) = "1" Then
        varCell(28) = "������"                                 '�ΐl������
    ElseIf varContent(46) = "1" Then
        varCell(28) = "�ΏۊO"                                 '�ΐl�ΏۊO
    Else
        varCell(28) = fncFindCode(varContent(47), "CE")        '�ΐl�����ی����z
        If varCell(28) = "������" Or varCell(28) = "�ΏۊO" Then
            varCell(28) = varContent(47)
        End If
    End If
    varCell(29) = IIf(varContent(48) = "1", _
                        "�ΏۊO", varContent(48))              '�������̏��Q����
    
    '--------AG��`AP��--------
    varCell(30) = IIf(varContent(49) = "1", _
                        "�ΏۊO", varContent(49))              '���ی��Ԏ��̏��Q����
    If varContent(50) = "1" Then
        varCell(31) = "������"                                 '�Ε�������
    ElseIf varContent(51) = "1" Then
        varCell(31) = "�ΏۊO"                                 '�Ε��ΏۊO
    Else
        varCell(31) = fncFindCode(varContent(52), "CI")        '�Ε������ی����z
        If varCell(31) = "������" Or varCell(31) = "�ΏۊO" Then
            varCell(31) = varContent(52)
        End If
    End If
    varCell(32) = fncFindCode(varContent(53), "BS")            '�Ε��ƐӋ��z
    varCell(33) = IIf(varContent(54) = "1", _
                        "�K�p����", varContent(54))            '�Ε����ߏC����p����
    If varContent(56) = 1 Then
        varCell(34) = "�ΏۊO"                                 '�l�g���Q�ΏۊO
    Else
        varCell(34) = fncFindCode(varContent(55), "CM")        '�l�g���Q(1��)
        If varCell(34) = "�ΏۊO" Then
            varCell(34) = varContent(55)
        End If
    End If
    varCell(35) = varContent(57)                               '�l�g���Q(1����)
    If varContent(60) = "1" Then
        varCell(36) = "�ΏۊO"                                 '����ҏ��Q�ΏۊO
    Else
        varCell(36) = varContent(59)                           '����ҏ��Q(1��)
'        varCell(36) = fncFindCode(varContent(59), "CQ")        '����ҏ��Q(1��)
'        If varCell(36) = "�ΏۊO" Then
'            varCell(36) = varContent(59)
'        End If
    End If
    varCell(37) = varContent(61)                               '����ҏ��Q(1����)
    varCell(38) = IIf(varContent(62) = "2", _
                        "�K�p����", varContent(62))            '����������
    varCell(39) = IIf(varContent(63) = "1", _
                        "�K�p����", varContent(63))            '���Ǝ��p����
    
    '--------AQ��`AV��--------
    varCell(40) = IIf(varContent(64) = "1", _
                        "�K�p����", varContent(64))            '�ٌ�m��p����
    varCell(41) = IIf(varContent(37) = "1", _
                        "����", varContent(37))                '�]�ƈ����������
    varCell(42) = varContent(41)                               '���̑�ԁE�g�̉��i����
    varCell(43) = IIf(varContent(73) = "1", "�s�K�p", varContent(73))   '�ԗ��������s�K�p����
    If blnChouhyouflg = False Then
        varCell(44) = ""                                           '���v�ی���
        varCell(45) = ""                                           '����ی���
        varCell(46) = ""                                           '�N�ԕی���
        varCell(47) = ""                                           '�g�c�x���t���O
    Else
        varCell(44) = varContent(31)                               '���v�ی���
        varCell(45) = varContent(32)                               '����ی���
        varCell(46) = varContent(33)                               '�N�ԕی���
        varCell(47) = IIf(varContent(67) = "1", "�g�c�G���[�L", IIf(varContent(67) = "2", "�x���L", ""))  '�g�c�x���t���O
    End If
    
    fncMeisaiSetCell = varCell
    
End Function
'�m���t���[�g���׃V�[�g�p�z��쐬�i�Z���j�@ �i2018/3 ���ذĖ��וt�@�\�ǉ��j
Function fncNonFleetMeisaiSetCell(ByVal varContent As Variant) As Variant
    Dim i           As Integer
    Dim strText     As String
    Dim strSaveDate As String
    Dim varCell(57) As Variant          'Cell�p�z��i�m���t���[�g���וt�̗�j
    
    strText = ""
    strSaveDate = ""
    
    '�V�[�g�̃Z���Ɋi�[
    '--------C��`L��--------
    varCell(0) = fncFindCode(varContent(1), "AA")             '�p�r�Ԏ햼
    varCell(1) = varContent(1)                                '�p�r�Ԏ�R�[�h
    varCell(2) = varContent(2)                                '�Ԗ�
    varCell(3) = varContent(68)                               '�o�^�ԍ�
    varCell(4) = varContent(69)                               '�ԑ�ԍ�
    varCell(5) = varContent(3)                                '�^��
    varCell(6) = varContent(4)                                '�d�l
    '���x�o�^�N��
    strSaveDate = fncToWareki(varContent(5) & "25", 8)
    If strSaveDate = varContent(5) & "25" Then
        varCell(7) = varContent(5)
    Else
        varCell(7) = strSaveDate
    End If
    '�Ԍ�������
    strSaveDate = fncToWareki(CStr(varContent(70)), 11)
    varCell(8) = strSaveDate
    '�����E�s����
    varCell(9) = fncFindCode(varContent(6), "AE")
    
    '--------M��`U��--------
    '�r�C��
    varCell(10) = varContent(7)
    
    '2.5���b�g���z�f�B�[�[��������
    varCell(11) = varContent(8)
    
    '��ی��Ґ��N����
    varCell(12) = fncToWareki(CStr(varContent(9)), 11)
    
    '�m���t���[�g����
    varCell(13) = fncFindCode(varContent(10), "AI")
    
    '���̗L�W���K�p����
    varCell(14) = fncFindCode(varContent(11), "AM")
    
    '�S�[���h�Ƌ�����
    varCell(15) = IIf(varContent(14) = "1", "�K�p����", varContent(14))
    
    '�g�p�ړI
    varCell(16) = fncFindCode(varContent(15), "DD")
    
    '�N�����
    varCell(17) = fncFindCode(varContent(34), "BC")
    
    '����^�]�ґΏۊO
    varCell(18) = IIf(varContent(35) = "1", "�ΏۊO", varContent(35))
   
    '�^�]�Ҍ���
    varCell(19) = fncFindCode(varContent(36), "BG")
    
    varCell(20) = IIf(varContent(72) = "1", "�K�p����", varContent(72))
    
    '�����N���X
    If blnChouhyouflg = False Then
        varCell(21) = ""                                       '����-�ԗ�
        varCell(22) = ""                                       '����-�ΐl
        varCell(23) = ""                                       '����-�Ε�
        varCell(24) = ""                                       '����-���Q
        varCell(25) = ""                                       '�V�Ԋ���
        varCell(26) = ""                                       '�ԗ����
        varCell(27) = ""                                       '�ԗ�����
    Else
        varCell(21) = varContent(23)                           '����-�ԗ�
        varCell(22) = varContent(24)                           '����-�ΐl
        varCell(23) = varContent(25)                           '����-�Ε�
        varCell(24) = varContent(26)                           '����-���Q
        varCell(25) = varContent(27)                           '�V�Ԋ���
        varCell(26) = varContent(30)                           '�ԗ����
        varCell(27) = varContent(29)                           '�ԗ�����
    End If
    varCell(28) = ""                                           '���̑�����-�{�^��
    
    '--------W��`AF��--------
    strText = IIf(varContent(16) = "3 ", "����^", "")
    strText = strText + IIf(varContent(17) = "1", "�����^�J�[�^", "")
    strText = strText + IIf(varContent(18) = "5 ", "���K�ԁ^", "")
    strText = strText + IIf(varContent(19) = "1 ", "�u�[���ΏۊO�^", "")
    strText = strText + IIf(varContent(20) = "80", "���[�X�J�[�I�[�v���|���V�[�^", "")
    strText = strText + IIf(varContent(21) = "93", "�I�[�v���|���V�[���������^", "")
    strText = strText + IIf(varContent(22) = "1 ", "���L�^", _
                        IIf(varContent(22) = "2 ", "�����L�^", ""))
    If blnChouhyouflg = False Then
        strText = strText + ""
    Else
        strText = strText + IIf(varContent(28) = "8", "����敪�^", "")
    End If
    If strText <> "" Then
        varCell(29) = Left(strText, Len(strText) - 1)          '���̑�����-�e�L�X�g
    End If
    
    varCell(30) = fncFindCode(varContent(38), "BK")            '�ԗ��ی��̎��
    varCell(31) = varContent(39)                               '�ԗ��ی����z
    varCell(32) = fncFindCode(varContent(40), "BO")            '�ԗ��ƐӋ��z
    varCell(33) = IIf(varContent(42) = "2", _
                        "�K�p����", varContent(42))            '�ԗ��S���Ք����
    varCell(34) = IIf(varContent(44) = "1", _
                        "�K�p����", varContent(44))            '�ԗ����ߏC����p����
    varCell(35) = IIf(varContent(43) = "1", _
                        "�ΏۊO", varContent(43))              '�ԗ�����ΏۊO����
    If varContent(45) = "1" Then
        varCell(36) = "������"                                 '�ΐl������
    ElseIf varContent(46) = "1" Then
        varCell(36) = "�ΏۊO"                                 '�ΐl�ΏۊO
    Else
        varCell(36) = fncFindCode(varContent(47), "CE")        '�ΐl�����ی����z
        If varCell(36) = "������" Or varCell(36) = "�ΏۊO" Then
            varCell(36) = varContent(47)
        End If
    End If
    varCell(37) = IIf(varContent(48) = "1", _
                        "�ΏۊO", varContent(48))              '�������̏��Q����
    
    '--------AG��`AP��--------
    varCell(38) = IIf(varContent(49) = "1", _
                        "�ΏۊO", varContent(49))              '���ی��Ԏ��̏��Q����
    If varContent(50) = "1" Then
        varCell(39) = "������"                                 '�Ε�������
    ElseIf varContent(51) = "1" Then
        varCell(39) = "�ΏۊO"                                 '�Ε��ΏۊO
    Else
        varCell(39) = fncFindCode(varContent(52), "CI")        '�Ε������ی����z
        If varCell(39) = "������" Or varCell(39) = "�ΏۊO" Then
            varCell(39) = varContent(52)
        End If
    End If
    varCell(40) = fncFindCode(varContent(53), "BS")            '�Ε��ƐӋ��z
    varCell(41) = IIf(varContent(54) = "1", _
                        "�K�p����", varContent(54))            '�Ε����ߏC����p����
    If varContent(56) = 1 Then
        varCell(42) = "�ΏۊO"                                 '�l�g���Q�ΏۊO
    Else
        varCell(42) = fncFindCode(varContent(55), "CM")        '�l�g���Q(1��)
        If varCell(42) = "�ΏۊO" Then
            varCell(42) = varContent(55)
        End If
    End If
    varCell(43) = varContent(57)                               '�l�g���Q(1����)
    If varContent(60) = "1" Then
        varCell(44) = "�ΏۊO"                                 '����ҏ��Q�ΏۊO
    Else
'        varCell(44) = fncFindCode(varContent(59), "CQ")        '����ҏ��Q(1��)
'        If varCell(44) = "�ΏۊO" Then
'            varCell(44) = varContent(59)
'        End If
    varCell(44) = varContent(59)
    End If
    varCell(45) = varContent(61)                               '����ҏ��Q(1����)
    varCell(46) = IIf(varContent(62) = "2", _
                        "�K�p����", varContent(62))            '����������
    varCell(47) = IIf(varContent(63) = "1", _
                        "�K�p����", varContent(63))            '���Ǝ��p����
    
    '--------AQ��`AV��--------
    varCell(48) = IIf(varContent(64) = "1", _
                        "�K�p����", varContent(64))            '�ٌ�m��p����
    
    '�]�ƈ����������i�m���t���[�g�s�v�j

    '2018/3 ���ذĖ��וt�@�\�ǉ�
    varCell(49) = fncFindCode(varContent(65), "BW")            '�t�@�~���[�o�C�N����

    varCell(50) = IIf(varContent(66) = "1", "�K�p����", varContent(66))        '�l�����ӔC�⏞����

    varCell(51) = IIf(varContent(58) = "2", "�K�p����", varContent(58))        '�����Ԏ��̓���

    varCell(52) = varContent(41)                               '��ԓ��Z�b�g����
    
    varCell(53) = IIf(varContent(73) = "1", "�s�K�p", varContent(73))  '�ԗ�����������p����̕s�K�p�Ɋւ������
   
    If blnChouhyouflg = False Then
        varCell(54) = ""                                           '���v�ی���
        varCell(55) = ""                                           '����ی���
        varCell(56) = ""                                           '�N�ԕی���
        varCell(57) = ""                                           '�g�c�x���t���O
    Else
        varCell(54) = varContent(31)                               '���v�ی���
        varCell(55) = varContent(32)                               '����ی���
        varCell(56) = varContent(33)                               '�N�ԕی���
        varCell(57) = IIf(varContent(67) = "1", "�g�c�G���[�L", IIf(varContent(67) = "2", "�x���L", ""))   '�g�c�x���t���O
    End If
    
    fncNonFleetMeisaiSetCell = varCell
    
End Function

'2�����z���1�����ڂ�1�����z����R�s�[
Public Sub subMeisaiArray(ByRef varAllMeisai As Variant, ByVal varMeisai As Variant, ByVal intActRow As Integer)

    Dim i As Integer
    
    For i = 0 To UBound(varMeisai)
        varAllMeisai(intActRow, i) = varMeisai(i)
    Next i

End Sub


'���ד���(Cell)�ɓ\��t��
Public Sub subMeisaiCell(ByVal varAllMeisai As Variant)
    Dim i As Integer
    Dim j As Integer
    Dim intStartRow As Integer
    Dim intStartCol As Integer
    
    Dim wstMeisai As Worksheet
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(1, wstMeisai)      '�V�[�g�I�u�W�F�N�g(���ד���)
    Else
        Call subSetSheet(17, wstMeisai)      '�V�[�g�I�u�W�F�N�g(���ד���)
    End If
    
    '���׊J�n�s
    intStartRow = 21
    intStartCol = 3
    
    Application.EnableEvents = False
    For i = 0 To UBound(varAllMeisai, 1)
        For j = 0 To UBound(varAllMeisai, 2)
            wstMeisai.Cells(intStartRow + i, intStartCol + j) = varAllMeisai(i, j)
        Next
        intStartCol = 3
    Next i
    
    Set wstMeisai = Nothing
    
    Application.EnableEvents = True
End Sub


'2018/3 ���ذĖ��וt�@�\�ǉ�
'PDF���J����Ă��邩�m�F
Public Function fncIsFileOpen(ByVal strArgFile As String) As Boolean

    On Error GoTo FILE_ERR
    
    Open strArgFile For Binary Access Read Lock Read As #1
    Close #1
    fncIsFileOpen = False
    
    Exit Function
    
FILE_ERR:

    fncIsFileOpen = True
    
End Function

