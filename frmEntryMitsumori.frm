VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEntryMitsumori 
   Caption         =   "���Ϗ����"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10755
   OleObjectBlob   =   "frmEntryMitsumori.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmEntryMitsumori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intConfirmMsg As Integer

'�����\��
'2018/3 ���ذĖ��וt�@�\�ǉ�
Public Sub UserForm_Initialize()
    '�t�H�[���̃^�C�g���ݒ�
    If FleetTypeFlg = 1 Then
        frmEntryMitsumori.Caption = frmEntryMitsumori.Caption & "�i�t���[�g�_��j"
    Else
        frmEntryMitsumori.Caption = frmEntryMitsumori.Caption & "�i�m���t���[�g���וt�_��j"
    End If
End Sub

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

'�߂�{�^��
Private Sub BtnBack_Click()

    Dim ctlFormCtrl As Control
    Dim objWorkSheet As Worksheet
    
    On Error GoTo Error
    
    intConfirmMsg = MsgBox("���͓��e���폜���Ē��[�I����ʂɑJ�ڂ��܂��B" & vbCrLf & "��낵���ł����H", vbYesNo, "�m�F�_�C�A���O")
    If intConfirmMsg = 6 Then
        '��ʏ�����
        For Each ctlFormCtrl In frmEntryMitsumori.Controls
            If TypeName(ctlFormCtrl) = "TextBox" Then _
                ctlFormCtrl.Value = ""
        Next ctlFormCtrl
        
        '�ۑ����폜
        Call subSetSheet(8, objWorkSheet) '�V�[�g�I�u�W�F�N�g�i�\���������ʓ��e�j
        objWorkSheet.Cells.ClearContents
        
        '2018/3 ���ذĖ��וt�@�\�ǉ�
        'Me.Hide
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

'' ���Ϗ�����{�^��
Private Sub BtnPrintMitsumori_Click()

    Dim intSame         As Integer
    Dim i               As Integer
    Dim intTotalCar     As Integer
    Dim intMeisaiCnt    As Integer
    Dim intPageCnt      As Integer
    Dim wsSaveCnt       As Integer
    Dim strIndex        As String
    Dim strSave         As String
    Dim strNowTime      As String
    Dim strFileName     As String
    Dim strOutputPath   As String
    Dim strFilePath     As String
    Dim wsAll           As Worksheet
    Dim wsTextK         As Worksheet
    Dim wsChohyo        As Worksheet
    Dim wsCarMeisaiSet  As Worksheet
    Dim wstMitsuSave    As Worksheet
    Dim wsMitsumoriSet  As Worksheet
    Dim wsSetting       As Worksheet
    Dim blnFstflg       As Boolean
    Dim intDaishaSet    As Integer '2018/3 ���ذĖ��וt�@�\�ǉ�

    On Error GoTo Error

    intConfirmMsg = MsgBox("���Ϗ���������܂��B" & vbCrLf & "��낵���ł����H", vbYesNo, "�m�F�_�C�A���O")
    If intConfirmMsg = 6 Then

        i = 16                                                  '�ݒ��ʂ̊J�n�s
        intMeisaiCnt = 1                                        '���׍s
        intPageCnt = 0                                          '�y�[�W��
        intSame = 1                                             '�����ڐ�(���������ꍇ�J�E���g�A�b�v)
        strIndex = ""                                           '�����ڔ�r�ϐ�
        wsSaveCnt = 0
        blnFstflg = True

        Call subSetSheet(5, wsSetting)      '�V�[�g�I�u�W�F�N�g(�ʎ��@�e��ݒ�)

        '�e�L�X�g�t�@�C���o�̓p�X�擾
        strOuptutPath = wsSetting.Range("B5").Value

        '�o�̓t�@�C���p�X
        If strOuptutPath = "" Then
            strFilePath = CreateObject("WScript.Shell").SpecialFolders.item("Desktop") & "\"
        Else
            strFilePath = strOuptutPath & "\"
        End If

        strFileName = strFilePath & strTextName & "_" & "���Ϗ��E���׏�.pdf"
        
        '2018/3 ���ذĖ��וt�@�\�ǉ�
        '�����t�@�C���m�F
        If Dir(strFileName) <> "" Then
            intConfirmMsg = MsgBox("�������O�̃t�@�C�������ɑ��݂��܂��B�㏑�����܂����H", vbYesNo, "�m�F�_�C�A���O")
            If intConfirmMsg = 7 Then Exit Sub
        End If

        '2018/3 ���ذĖ��וt�@�\�ǉ�
        'PDF�t�@�C�����J����Ă��邱�Ƃ��m�F
        If fncIsFileOpen(strFileName) Then
            intConfirmMsg = MsgBox("PDF�t�@�C�����J����Ă��܂��B" & vbCrLf & "���Ă��炲�g�p���������B", vbOKOnly, "�ʒm�_�C�A���O")
            Exit Sub
        End If

        '�x���`�F�b�N
        If fncTextEntryWarChk(1) Then
        End If

        '���ݓ����擾(�N��������)
        strNowTime = Format(Now, "yyyymmddHHMM")

        Call subMeisaiUnProtect         '�V�[�g�̕ی�̉���
        Call subBookUnProtect           '�u�b�N�̕ی������
        Call subSheetVisible(True)      '�V�[�g�E�u�b�N�̕\��

        Application.ScreenUpdating = False

        '�V�[�g�����E�폜
        For Each wsAll In ThisWorkbook.Worksheets
            If wsAll.Name = "���Ϗ�WK" Then
                Application.DisplayAlerts = False
                Worksheets("���Ϗ�WK").Delete
                Application.DisplayAlerts = True
            ElseIf wsAll.Name = "�ԗ����׏�WK" Then
                Application.DisplayAlerts = False
                Worksheets("�ԗ����׏�WK").Delete
                Application.DisplayAlerts = True
            End If
        Next wsAll

        '�V�[�g�R�s�[
        With ThisWorkbook
            If FleetTypeFlg = 1 Then
                '���Ϗ�
                .Worksheets("���Ϗ�").Visible = True
                .Worksheets("���Ϗ�").Copy After:=.Worksheets(.Worksheets.Count)
                ActiveSheet.Name = "���Ϗ�WK"
                .Worksheets("���Ϗ�").Visible = False
                '�ԗ����׏�
                .Worksheets("�ԗ����׏�").Visible = True
                .Worksheets("�ԗ����׏�").Copy After:=.Worksheets(Worksheets.Count)
                ActiveSheet.Name = "�ԗ����׏�WK"
                .Worksheets("�ԗ����׏�").Visible = False
            Else
                '2018/3 ���ذĖ��וt�@�\�ǉ�
                '���Ϗ�
                .Worksheets("���Ϗ��i�m���t���[�g�j").Visible = True
                .Worksheets("���Ϗ��i�m���t���[�g�j").Copy After:=.Worksheets(.Worksheets.Count)
                ActiveSheet.Name = "���Ϗ�WK"
                .Worksheets("���Ϗ��i�m���t���[�g�j").Visible = False
                '�ԗ����׏�
                .Worksheets("�ԗ����׏��i�m���t���[�g�j").Visible = True
                .Worksheets("�ԗ����׏��i�m���t���[�g�j").Copy After:=.Worksheets(Worksheets.Count)
                ActiveSheet.Name = "�ԗ����׏�WK"
                .Worksheets("�ԗ����׏��i�m���t���[�g�j").Visible = False
            End If
        End With
        
        Call subSetSheet(6, wsTextK)         '�V�[�g�I�u�W�F�N�g(�e�L�X�g���e(����))
        Call subSetSheet(8, wstMitsuSave)    '�V�[�g�I�u�W�F�N�g�i�\���������ʓ��e�j

        '���t�ۑ䐔�擾
        intTotalCar = Val(wsTextK.Cells(1, 19))
        
        '�\���������ʕۑ����e�N���A(�\���������ʓ��e)
        If FleetTypeFlg = 1 Then
            Call subSetSheet(9, wsMitsumoriSet)  '�V�[�g�I�u�W�F�N�g(�ʎ��@���Ϗ��ݒ�)
            Call subSetSheet(10, wsCarMeisaiSet) '�V�[�g�I�u�W�F�N�g(�ʎ��@�ԗ����׏��ݒ�)
        Else
            '2018/3 ���ذĖ��וt�@�\�ǉ�
            Call subSetSheet(20, wsMitsumoriSet)  '�V�[�g�I�u�W�F�N�g(�ʎ��@���Ϗ��ݒ�i�m���t���[�g�j)
            Call subSetSheet(21, wsCarMeisaiSet) '�V�[�g�I�u�W�F�N�g(�ʎ��@�ԗ����׏��ݒ�i�m���t���[�g�j)
        End If
        Call subSetSheet(102, wsChohyo)       '�V�[�g�I�u�W�F�N�g(�ԗ����׏�WK)
        '
        wstMitsuSave.Cells.ClearContents
        
        '�\���������ʕۑ����e�X�V
        wstMitsuSave.Cells(1, 1) = txtKeiyakusyaHoujin.Value
        wstMitsuSave.Cells(1, 2) = txtKeiyakusyaDaihyou.Value
        wstMitsuSave.Cells(1, 3) = txtDairiten.Value
        wstMitsuSave.Cells(1, 4) = txtTantousya.Value
        wstMitsuSave.Cells(1, 5) = txtComment.Value
        
        '���[�ݒ�ǂݍ��݁E���[�쐬
        With wsMitsumoriSet
            Do Until .Cells(i, 1).MergeArea(1) = ""
                
                '�X�V�p
                Call subFormSetting(1, Val(.Cells(i, 1).MergeArea(1)), Val(.Cells(i, 2).MergeArea(1)), _
                                    CStr(.Cells(i, 4).MergeArea(1)), Val(.Cells(i, 5).MergeArea(1)), _
                                    CStr(.Cells(i, 6).MergeArea(1)), .Cells(i, 7).MergeArea(1), strNowTime, , _
                                    , , , , intSame, , , blnFstflg)
    
                i = i + 1
                
                If .Cells(i, 1).MergeArea(1) = "" And blnFstflg Then
                    i = 16
                    blnFstflg = False
                End If
                
            Loop
        End With
        
        '���[�ݒ�ǂݍ��݁E���[�쐬
        Do Until intMeisaiCnt >= intTotalCar + 1
            i = 16
            blnFstflg = True
            intDaishaSet = 0 '2018/3 ���ذĖ��וt�@�\�ǉ�
                    
            With wsCarMeisaiSet
                Do Until .Cells(i, 1).MergeArea(1) = ""
                    
                    '�����ڂ��������ꍇ�AintSame���J�E���g�A�b�v
                    '�O���ڂ����݂̍��ڂƈႤ�ꍇ�AintSame���y�[�W�J�n���׍s���Ƃ���
                    If strIndex = .Cells(i, 1).MergeArea(1) & "," & Val(.Cells(i, 2).MergeArea(1)) Then
                        If FleetTypeFlg = 1 Then
                            intSame = intSame + 1
                        Else
                            '2018/3 ���ذĖ��וt�@�\�ǉ�
                            If blnFstflg = False And strIndex = "2,42" Then '��ԓ��Z�b�g����@���ꏈ���i1���ڂŕ����Z�b�g���邽�߁j
                                intDaishaSet = intDaishaSet + 1
                                If intDaishaSet >= 2 Then
                                    intSame = 1 + (intPageCnt * 2) + 1
                                End If
                            Else
                                intSame = intSame + 1
                            End If
                        End If
                    Else
                        strIndex = .Cells(i, 1).MergeArea(1) & "," & Val(.Cells(i, 2).MergeArea(1))
                        If FleetTypeFlg = 1 Then
                            intSame = 1 + (intPageCnt * 10)
                        Else
                            '2018/3 ���ذĖ��וt�@�\�ǉ�
                            intSame = 1 + (intPageCnt * 2)
                        End If
                    End If

                    If CStr(.Cells(i, 4).MergeArea(1)) = "" Then
                        strSave = CStr(.Cells(i, 4).MergeArea(1))
                    Else
                        If Evaluate("ISREF(" & CStr(.Cells(i, 4).MergeArea(1)) & ")") = False Then
                            strSave = ""
                        Else
                            If FleetTypeFlg = 1 Then
                                strSave = CStr(.Cells(Val(.Range(CStr(.Cells(i, 4).MergeArea(1))).Row + (49 * intPageCnt)), _
                                                        .Range(CStr(.Cells(i, 4).MergeArea(1))).Column).Address(False, False))
                            Else
                                '2018/3 ���ذĖ��וt�@�\�ǉ�
                                strSave = CStr(.Cells(Val(.Range(CStr(.Cells(i, 4).MergeArea(1))).Row + (44 * intPageCnt)), _
                                                        .Range(CStr(.Cells(i, 4).MergeArea(1))).Column).Address(False, False))
                            End If
                        End If
                    End If

                    '�X�V�p
                    Call subFormSetting(2, Val(.Cells(i, 1).MergeArea(1)), Val(.Cells(i, 2).MergeArea(1)), _
                                        strSave, Val(.Cells(i, 5).MergeArea(1)), CStr(.Cells(i, 6).MergeArea(1)), _
                                        .Cells(i, 7).MergeArea(1), strNowTime, , , , , , _
                                        intSame, intMeisaiCnt, intPageCnt, blnFstflg)

                    i = i + 1

                    If .Cells(i, 1).MergeArea(1) = "" And blnFstflg Then
                        i = 16
                        If wsSaveCnt = 0 Then
                            wsSaveCnt = intMeisaiCnt
                            intMeisaiCnt = 1
                        Else
                        End If
                        blnFstflg = False

                    End If
                Loop

                intPageCnt = intPageCnt + 1

                If intMeisaiCnt < intTotalCar + 1 Then
                    If FleetTypeFlg = 1 Then
                        ' 1 �` 49 �s�ڂ� 50 �s�� �֓\��t��
                        Application.DisplayAlerts = False
                        wsChohyo.Range("1:49").Copy
                        wsChohyo.Range(CStr(49 * Val(intPageCnt) + 1 & ":" & 49 * Val(intPageCnt) + 1)).Select
                        wsChohyo.Paste
                        Application.DisplayAlerts = True
                    Else
                        '2018/3 ���ذĖ��וt�@�\�ǉ�
                        ' 1 �` 44 �s�ڂ� 45 �s�� �֓\��t��
                        Application.DisplayAlerts = False
                        wsChohyo.Range("1:44").Copy
                        wsChohyo.Range(CStr(44 * Val(intPageCnt) + 1 & ":" & 44 * Val(intPageCnt) + 1)).Select
                        wsChohyo.Paste
                        Application.DisplayAlerts = True
                    End If
                End If
            
            End With
            
        Loop
        
        Dim strPrintSheet(1) As String
        strPrintSheet(0) = "���Ϗ�WK"
        strPrintSheet(1) = "�ԗ����׏�WK"
        
        ThisWorkbook.Worksheets(strPrintSheet).Select
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=strFileName
        
        '�V�[�g�폜
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("���Ϗ�WK").Delete
        ThisWorkbook.Worksheets("�ԗ����׏�WK").Delete
        Application.DisplayAlerts = True
        
        Call subSheetVisible(False)      '�V�[�g�E�u�b�N�̔�\��
        Call subBookProtect              '�u�b�N�̕ی�
        Call subMeisaiProtect            '�V�[�g�̕ی�
        
        Application.ScreenUpdating = True
        
        '2018/3 ���ذĖ��וt�@�\�ǉ�
        'MsgBox "������������܂����B", vbOKOnly, "�ʒm�_�C�A���O"
        MsgBox "PDF�t�@�C�����o�͂��܂����B", vbOKOnly, "�ʒm�_�C�A���O"
        
        Set wsTextK = Nothing
        Set wsChohyo = Nothing
        Set wstMitsuSave = Nothing
        Set wsCarMeisaiSet = Nothing
    
    End If
    
    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "BtnPrintMitsumori_Click" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["
        
End Sub

