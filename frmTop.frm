VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTop 
   Caption         =   "TOP"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7320
   OleObjectBlob   =   "frmTop.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intConfirmMsg As Integer
'�e�L�X�g�t�@�C���o�͎��̔��ʃt���O
Dim blnTextErrflg As Boolean


'�A�N�e�B�u
Private Sub UserForm_Activate()
    Dim wsSetting As Worksheet
    
    On Error GoTo Error
    
    Call subSetSheet(5, wsSetting)      '�V�[�g�I�u�W�F�N�g(�ʎ��@�e��ݒ�)

    Me.Caption = "TOP  ( Ver." & wsSetting.Range("B2") & " )"
    
    '�V���[�g�J�b�g�L�[�̗L��
    Call subShortCutKey(2)

    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    Application.DisplayStatusBar = True
    Application.DisplayFormulaBar = True
    Application.CommandBars("Ply").Enabled = True
    
    Set wsSetting = Nothing
    
    On Error GoTo 0
    
    Exit Sub

Error:
    MsgBox "UserForm_Activate" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["

End Sub

'�Ǘ��ҋ@�\
Sub BtnAdministrator_Click()
    Dim strInput As String
    Dim strPassword As String
    Dim wsSetting As Worksheet
    
    On Error GoTo Error
    
    Call subSetSheet(5, wsSetting)      '�V�[�g�I�u�W�F�N�g(�ʎ��@�e��ݒ�)
    
    BtnFleet.SetFocus
    
    strPassword = wsSetting.Range("B4").Value
    
    Set wsSetting = Nothing
    
    strInput = InputBox("�p�X���[�h����͂��Ă�������", "�p�X���[�h���̓_�C�A���O")
    
    If StrPtr(strInput) = 0 Then
        Exit Sub
    End If
    
    If strInput = strPassword Then
        
        Call subBookUnProtect           '�u�b�N�̕ی������
        Call subMeisaiUnProtect         '�V�[�g�̕ی������
        Call subSheetVisible(True)      '�V�[�g�E�u�b�N�̕\��
        
        Application.ScreenUpdating = False                            '�`���~
        
        ThisWorkbook.Worksheets("�ʎ��@�R�[�h�l").Visible = True
        ThisWorkbook.Worksheets("�ʎ��@�e��ݒ�").Visible = True
        ThisWorkbook.Worksheets("���ד���").Visible = False
        ThisWorkbook.Worksheets("���Ϗ�").Visible = True
        ThisWorkbook.Worksheets("�ԗ����׏�").Visible = True
        ThisWorkbook.Worksheets("�_��\����1����").Visible = True
        ThisWorkbook.Worksheets("�_��\����2����").Visible = True
        ThisWorkbook.Worksheets("���׏�").Visible = True
        ThisWorkbook.Worksheets("�\�����d�c�o").Visible = True
        ThisWorkbook.Worksheets("���׏��d�c�o").Visible = True
        ThisWorkbook.Worksheets("�ʎ��@���Ϗ��ݒ�").Visible = True
        ThisWorkbook.Worksheets("�ʎ��@�ԗ����׏��ݒ�").Visible = True
        ThisWorkbook.Worksheets("�ʎ��@�\����(1����)�ݒ�").Visible = True
        ThisWorkbook.Worksheets("�ʎ��@�\����(2����)�ݒ�").Visible = True
        ThisWorkbook.Worksheets("�ʎ��@���׏��ݒ�").Visible = True
        ThisWorkbook.Worksheets("�ʎ��@�\�����d�c�o�ݒ�").Visible = True
        ThisWorkbook.Worksheets("�ʎ��@���׏��d�c�o�ݒ�").Visible = True
        '2018/3 ���ذĖ��וt�@�\�ǉ�
        ThisWorkbook.Worksheets("���ד��́i�m���t���[�g�j").Visible = False
        ThisWorkbook.Worksheets("���׏����").Visible = False
        ThisWorkbook.Worksheets("���׏�����i�m���t���[�g�j").Visible = False
        ThisWorkbook.Worksheets("�ʎ��@�R�[�h�l�i�m���t���[�g�j").Visible = True
        ThisWorkbook.Worksheets("���Ϗ��i�m���t���[�g�j").Visible = True
        ThisWorkbook.Worksheets("�ԗ����׏��i�m���t���[�g�j").Visible = True
        ThisWorkbook.Worksheets("���׏��i�m���t���[�g�j").Visible = True
        ThisWorkbook.Worksheets("�\�����d�c�o�i�m���t���[�g�j").Visible = True
        ThisWorkbook.Worksheets("���׏��d�c�o�i�m���t���[�g�j").Visible = True
        ThisWorkbook.Worksheets("�ʎ��@���Ϗ��ݒ�i�m���t���[�g�j").Visible = True
        ThisWorkbook.Worksheets("�ʎ��@�ԗ����׏��ݒ�i�m���t���[�g�j").Visible = True
        ThisWorkbook.Worksheets("�ʎ��@�\����(1����)�ݒ�i�m���t���[�g�j").Visible = True
        ThisWorkbook.Worksheets("�ʎ��@�\����(2����)�ݒ�i�m���t���[�g�j").Visible = True
        ThisWorkbook.Worksheets("�ʎ��@���׏��ݒ�i�m���t���[�g�j").Visible = True
        ThisWorkbook.Worksheets("�ʎ��@�\�����d�c�o�ݒ�i�m���t���[�g�j").Visible = True
        ThisWorkbook.Worksheets("�ʎ��@���׏��d�c�o�ݒ�i�m���t���[�g�j").Visible = True
        ThisWorkbook.Worksheets("�e�L�X�g���e(����)").Visible = False
        ThisWorkbook.Worksheets("�e�L�X�g���e(����)").Visible = False
        
        
        ThisWorkbook.Worksheets("�ʎ��@�R�[�h�l").Activate
        ThisWorkbook.Worksheets("�ʎ��@�R�[�h�l").Range("A1").Select
        
        Application.ScreenUpdating = True                            '�`��ĊJ
        
        Call subBookProtect             '�u�b�N�̕ی�
        
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
        
        Windows(ThisWorkbook.Name).Visible = True
                
        '2018/3 ���ذĖ��וt�@�\�ǉ�
        FleetTypeFlg = 0
        
        Application.OnKey "%{q}", "fncAdmin"
        Me.Hide
        
    Else
        MsgBox "�p�X���[�h������������܂���", vbOKOnly, "�G���[�_�C�A���O"
    End If
    
    On Error GoTo 0
    
    Exit Sub

Error:
    MsgBox "BtnAdministrator_Click" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["

End Sub

'�t���[�g�_��{�^������
Private Sub BtnFleet_Click()
    
     On Error GoTo Error
   
    '�t���[�g�_��E�m���t���[�g���וt�_��𔻒f�p�t���O(�t���[�g��ݒ�)
    FleetTypeFlg = 1
    blnFleetBtnFlg = True
    
    '���ʓ��͉�ʂ�\��
    Call BtnPathDelete_Click
    Me.Hide
    
    frmKyoutsuu.Show vbModeless
    
    On Error GoTo 0
    
    Exit Sub
    
Error:
    MsgBox "BtnFleet_Click" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["


End Sub

'�m���t���[�g���וt�_��{�^������
Private Sub BtnNonFleet_Click()
    
    On Error GoTo Error
   
    '�t���[�g�_��E�m���t���[�g���וt�_��𔻒f�p�t���O(�t���[�g��ݒ�)
    FleetTypeFlg = 2
    blnNonFleetBtnFlg = True
    
    '���ʓ��͉�ʂ�\��
    Call BtnPathDelete_Click
    Me.Hide
    
    frmKyoutsuu.Show vbModeless
    
    On Error GoTo 0
    
    Exit Sub
    
Error:
    MsgBox "BtnFleet_Click" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["
    
    
End Sub

'�N���A�{�^������
Private Sub BtnPathDelete_Click()
    '�I������Ă���p�X���e�L�X�g����폜
    With txtFilePath
        .Value = ""
        
        .SetFocus
        .SelStart = 0
    End With
    
End Sub

'�t�@�C���I���{�^��������
Private Sub BtnFileSelect_Click()
    Dim varFilePath  As Variant
    Dim varTxtPath   As Variant

    On Error GoTo Error

    'TXT�t�@�C����I��
    varTxtPath = Application.GetOpenFilename(FileFilter:="TXT�t�@�C��,*.txt", MultiSelect:=True)

    '�e�L�X�g�{�b�N�X�o�͗p�ɁATXT�t�@�C���p�X��ϊ�
    If IsArray(varTxtPath) Then

        If txtFilePath.Value <> "" Then
            txtFilePath.Value = txtFilePath.Value & vbCrLf
        End If

        '�t�@�C���p�X���e�L�X�g�Ɋi�[
        For Each varFilePath In varTxtPath
            txtFilePath.Value = txtFilePath.Value & varFilePath & vbCrLf
        Next varFilePath

        txtFilePath.Value = Left(txtFilePath.Value, Len(txtFilePath.Value) - 2)

        txtFilePath.SetFocus
        txtFilePath.SelStart = 0

    End If

    Set varTxtPath = Nothing

    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "BtnFileSelect_Click" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["

End Sub

'�ĊJ�{�^������
Private Sub BtnReuse_Click()
    Dim i             As Integer
    Dim intFileNumber As Integer      '�t�@�C���ԍ�
    Dim intActRow     As Integer      '�ǂݍ��ݍs
    Dim intMaxRow     As Integer      '�t�@�C���ő�s��
    Dim intCar        As Integer      '���t�ۑ䐔
    Dim intMsgBox     As Integer
    Dim blnFiles      As Boolean
    Dim strFileName   As String
    Dim varFilePath   As Variant      '�t�@�C���p�X
    Dim varFiles      As Variant      '�t�@�C���p�X�ޔ�p
    
    On Error GoTo Error

    If (txtFilePath.Value) = "" Then Exit Sub

    varFiles = Split(txtFilePath.Value, vbCrLf)

    ReDim varFilePath(UBound(varFiles), 0)

    blnFiles = False
    intMaxRow = 0
    intCar = 0
    strTextName = ""

    '�t�@�C�����擾
    For i = 0 To UBound(varFiles)
        If i = 1 Then
            blnFiles = True
        End If

        varFilePath(i, 0) = varFiles(i)

    Next i

    '�t�@�C����������ւ�
    If blnFiles Then
        Call subArrayVar(varFilePath)
    End If

    '�t�@�C�����G���[�`�F�b�N
    If fncFileNameCheck(varFilePath, blnFiles) Then
        intConfirmMsg = MsgBox("�t�@�C�������s���ł��B", vbOKOnly + vbExclamation, "�G���[�_�C�A���O")

        If intConfirmMsg = 1 Then
            blnTextErrflg = True
            Exit Sub
        End If

    End If

    '�t�@�C�����擾
    strFileName = Mid(varFilePath(0, 0), InStrRev(varFilePath(0, 0), "\") + 1)
    If blnFiles Then
        '����
        strTextName = Mid(Right(strFileName, 19), 1, 12)
    Else
        '�P��
        strTextName = Mid(Right(strFileName, 16), 1, 12)
    End If

    '�t�@�C�����e���擾
    For i = 0 To UBound(varFilePath)
        '�󂢂Ă���t�@�C���ԍ����擾
        intFileNumber = FreeFile

        '���̓t�@�C����Input���[�h�ŊJ��
        Open varFilePath(i, 0) For Input Lock Write As #intFileNumber

        '�ǂݍ��݊J�n�s
        intActRow = 1

        Do Until EOF(1)
            If intActRow >= intMaxRow Then
                '�z�񒷂�ύX
                ReDim Preserve varFilePath(UBound(varFiles), intActRow - 1)
                intMaxRow = intActRow
            End If

            '1�s���Ƃɓǂݍ���
            Line Input #intFileNumber, varFilePath(i, intActRow - 1)

            intActRow = intActRow + 1
        Loop

        Close #intFileNumber
    Next i

    '���ʍ��ڃG���[�`�F�b�N �����
    '�t���[�g�A�m���t���[�g����i2018/3 ���ذĖ��וt�@�\�ǉ��j
    If fncKyotsuErr(varFilePath, intCar) Then
        intConfirmMsg = MsgBox("�t�@�C�����e���s���ł��B", vbOKOnly + vbExclamation, "�G���[�_�C�A���O")

        If intConfirmMsg = 1 Then
            blnTextErrflg = True
            Exit Sub
        End If
    End If

    '���׍��ڃG���[�`�F�b�N
    If fncMeisaiErr(varFilePath) Then
        intConfirmMsg = MsgBox("�t�@�C�����e���s���ł��B", vbOKOnly + vbExclamation, "�G���[�_�C�A���O")

        If intConfirmMsg = 1 Then
            blnTextErrflg = True
            Exit Sub
        End If
    End If
    
    If blnChouhyouflg Then
        intMsgBox = MsgBox("���[�o�͂��܂��B" & vbCrLf & "��낵���ł���?", vbYesNo, "�m�F�_�C�A���O")
    Else
        intMsgBox = MsgBox("�ĊJ���܂��B" & vbCrLf & "��낵���ł���?", vbYesNo, "�m�F�_�C�A���O")
    End If

    If intMsgBox = 6 Then

        Call subMeisaiUnProtect     '�V�[�g�̕ی�̉���

        '���ʍ��ڕۑ��p�V�[�g�ɔ��f
        Call fncKyotsuEntry(varFilePath, intCar)

        '���ד��̓V�[�g�ɔ��f
        Call fncMeisaiEntry(varFilePath, intCar)

        Call subMeisaiProtect       '�V�[�g�̕ی�

        '���ʓ��͉�ʂ�\��
        Call BtnPathDelete_Click
        Me.Hide

        If blnChouhyouflg Then
            frmPrintMenu.Show vbModeless
        Else
            frmKyoutsuu.Show vbModeless
        End If

    End If

    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "BtnReuse_Click" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["

End Sub

'���[�o�̓{�^������
Private Sub BtnFormOutput_Click()

    On Error GoTo Error

    blnChouhyouflg = True

    Call BtnReuse_Click

    If blnTextErrflg Then
        blnTextErrflg = False
        blnChouhyouflg = False
        Exit Sub
    End If

    blnChouhyouflg = False

    On Error GoTo 0

    Exit Sub

Error:
    MsgBox "BtnFormOutput_Click" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["

End Sub

'�t�@�C�����G���[�`�F�b�N
Private Function fncFileNameCheck(ByVal varFilePath As Variant, blnFiles As Boolean) As Boolean
    Dim i             As Integer
    Dim strFileName   As String
    Dim strFiles      As String
    Dim intFileNo     As Integer

    fncFileNameCheck = False
    strFileName = ""
    strFiles = ""
    intFileNo = 0

    '�P���E�����t�@�C���m�F
    If blnFiles Then
    '�����̏ꍇ
        '���X�g���̃t�@�C���������[�v
        For i = 0 To UBound(varFilePath)
            strFileName = Mid(varFilePath(i, 0), InStrRev(varFilePath(i, 0), "\") + 1)

            '�t�@�C���`���m�F("_")
            If Mid(Right(strFileName, 7), 1, 1) = "_" Then
            
                '�t�@�C���`��(�C�ӕ���+"YYYYMMDDhhmm_01.txt")�̏ꍇ+����ȊO
                If IsDate(Format(Mid(Right(strFileName, 19), 1, 12), "##/##/## ##:##")) Then
                    intFileNo = Val(Mid(Right(strFileName, 6), 1, 2))
                    If (intFileNo = i + 1) Then
                        '�t�@�C���`��("YYYYMMDDhhmm")�ɑ��Ⴊ����ꍇ
                        If strFiles = "" Then
                            strFiles = Mid(Right(strFileName, 19), 1, 12)
                        Else
                            If strFiles = Mid(Right(strFileName, 19), 1, 12) Then
                            Else
                                fncFileNameCheck = True
                            End If
                        End If
                    Else
                        fncFileNameCheck = True
                    End If
                Else
                '�t�@�C���������t�^�ɕϊ��ł��Ȃ��ꍇ
                    fncFileNameCheck = True
                End If
            
            Else
                fncFileNameCheck = True
            End If
            If fncFileNameCheck Then
                Exit Function
            End If
        Next i
    Else
    '�P���̏ꍇ
        strFileName = Mid(varFilePath(0, 0), InStrRev(varFilePath(0, 0), "\") + 1)
        '�t�@�C���`��(�C�ӕ���+YYYYMMDDhhmm.txt)�̏ꍇ+����ȊO
        If IsDate(Format(Mid(Right(strFileName, 16), 1, 12), "##/##/## ##:##")) Then
        Else
        '�t�@�C���������t�^�ɕϊ��ł��Ȃ��ꍇ
            fncFileNameCheck = True
        End If
    End If
End Function

'�I���t�@�C�����ύX
Private Sub subArrayVar(ByRef varContent As Variant)
    Dim varArray As Variant
    Dim varSave As Variant
    Dim i As Integer
    Dim j As Integer

    varArray = varContent

    '�A�Ԋm�F
    For i = 0 To UBound(varArray, 1)
        For j = UBound(varArray, 1) To i Step -1
            If Val(Mid(Right(varArray(i, 0), 6), 1, 2)) > Val(Mid(Right(varArray(j, 0), 6), 1, 2)) Then
                varSave = varArray(i, 0)
                varArray(i, 0) = varArray(j, 0)
                varArray(j, 0) = varSave
            End If
        Next j
    Next i

    '�t�@�C������ւ�
    For i = 0 To UBound(varContent, 1)
        varContent(i, 0) = varArray(i, 0)
    Next i
End Sub

'���ʍ��ڃG���[�`�F�b�N
Function fncKyotsuErr(ByVal varContent As Variant, ByRef intCarAll As Integer) As Boolean
    Dim i               As Integer
    Dim j               As Integer
    Dim intCar          As Integer
    Dim strBefore       As String
    Dim strKyotsu       As String
    Dim varKyotsuRow    As Variant

    fncKyotsuErr = False
    strBefore = ""
    intCar = 0

    '�G���[�`�F�b�N
    For i = 0 To UBound(varContent, 1)
        strKyotsu = ""

        '���ʕ����̃J���}�̐���37�ł͂Ȃ��ꍇ�A�G���[
        If UBound(Split(varContent(i, 0), ",")) = 37 Then
            varKyotsuRow = Split(varContent(i, 0), ",")
        Else
            fncKyotsuErr = True
            Exit Function
        End If

        '���ʕ����̑��t�ۑ䐔��3���ȏ�܂���0���̏ꍇ�A�G���[
        If Len(varKyotsuRow(18)) = 0 Then
            fncKyotsuErr = True
            Exit Function
        ElseIf Len(varKyotsuRow(18)) >= 3 Then
            fncKyotsuErr = True
            Exit Function
        End If

        '���ʕ����̑��t�ۑ䐔��0�̏ꍇ�A�G���[
        If Val(varKyotsuRow(18)) = "0" Then
            fncKyotsuErr = True
            Exit Function
        End If

        '���ʕ����Ƀt�@�C�����ő��Ⴊ����ꍇ�A�G���[
        For j = 0 To UBound(varKyotsuRow)
            If j = 18 Then
            Else
                strKyotsu = strKyotsu + varKyotsuRow(j)
            End If
        Next j

        If strBefore = "" Then
            strBefore = strKyotsu
        Else
            If strBefore = strKyotsu Then
            Else
                fncKyotsuErr = True
                Exit Function
            End If
        End If

        intCar = intCar + varKyotsuRow(18)
    Next i

    intCarAll = intCar

    If intCarAll > 999 Then
        fncKyotsuErr = True
        Exit Function
    End If


    '���ʕ������t���[�g�敪�𔻒�i2018/3 ���ذĖ��וt�@�\�ǉ��j
    If Val(varKyotsuRow(4)) = "1" Then
        FleetTypeFlg = 2               '�m���t���[�g
    Else
        FleetTypeFlg = 1               '�t���[�g
    End If

    '�m���t���[�g�������t�@�C���I�����̓G���[�Ƃ���i2018/3 ���ذĖ��וt�@�\�ǉ��j
    If FleetTypeFlg = 2 And UBound(varContent, 1) > 0 Then
        fncKyotsuErr = True
        Exit Function
    End If

    Set varKyotsuRow = Nothing

End Function


'���׍��ڃG���[�`�F�b�N
Function fncMeisaiErr(ByVal varContent As Variant) As Boolean
    Dim i               As Integer
    Dim j               As Integer
    Dim intCar          As Integer
    Dim varKyotsuRow    As Variant
    
    fncMeisaiErr = False
    
    '�G���[�`�F�b�N
    For i = 0 To UBound(varContent, 1)
        intCar = 0
        varKyotsuRow = Split(varContent(i, 0), ",")
        
        For j = 1 To UBound(varContent, 2)
            If UBound(Split(varContent(i, j))) = -1 Then Exit For
            
            '���ו����̃J���}�̐���81�ł͂Ȃ��ꍇ�A�G���[
            If UBound(Split(varContent(i, j), ",")) = 81 Then
                intCar = intCar + 1
            Else
                fncMeisaiErr = True
                Exit Function
            End If
            
        Next j
        
        '���t�ۑ䐔�Ɩ��ׂ̃��R�[�h�����v���Ȃ��ꍇ�A�G���[
        If intCar = varKyotsuRow(18) Then
        Else
            fncMeisaiErr = True
            Exit Function
        End If
        
    Next i
    
    Set varKyotsuRow = Nothing
    
End Function

'���ʍ��ڕۑ��p�V�[�g�ɔ��f
Private Sub fncKyotsuEntry(ByVal varContent As Variant, ByVal intCar As Integer)
    Dim i                As Integer
    Dim strKyotsu        As String
    Dim varKyotsuRow     As Variant
    Dim varKyotsuCol(31) As Variant
    Dim wstKyotsu As Worksheet
    Dim varMeisaiRow     As Variant '2018/3 ���ذĖ��וt�@�\�ǉ�
    
    Call subSetSheet(2, wstKyotsu)        '�V�[�g�I�u�W�F�N�g(�ʎ��@���ʍ���)
    
    varKyotsuRow = Split(varContent(0, 0), ",")

    varKyotsuCol(0) = varKyotsuRow(1)   '��t�敪
    varKyotsuCol(1) = varKyotsuRow(2)   '��ی���
    varKyotsuCol(2) = varKyotsuRow(3)   '�ی����
    varKyotsuCol(3) = varKyotsuRow(4)   '�t���[�g�敪
    varKyotsuCol(4) = varKyotsuRow(5)   '�ی��n����
    varKyotsuCol(5) = varKyotsuRow(10)  '�ی�����
    varKyotsuCol(6) = varKyotsuRow(9)   '�v�Z���@
    varKyotsuCol(7) = varKyotsuRow(13)  '�������@
    varKyotsuCol(8) = varKyotsuRow(14)  '�D�Ǌ���
    varKyotsuCol(9) = varKyotsuRow(15)  '����f������
    varKyotsuCol(10) = varKyotsuRow(16) '�t���[�g��������
    varKyotsuCol(11) = varKyotsuRow(17) '�t���[�g�R�[�h
    varKyotsuCol(14) = varKyotsuRow(20) '�X�֔ԍ�
    varKyotsuCol(15) = varKyotsuRow(21) '�_��ҏZ���i�J�i�j
    varKyotsuCol(16) = varKyotsuRow(22) '�_��ҏZ���i�����j
'    varKyotsuCol(17) = varKyotsuRow(23) '�_��ҏZ���i�����j
    varKyotsuCol(17) = varKyotsuRow(23) '�@�l���i�J�i�j
    varKyotsuCol(18) = varKyotsuRow(24) '�@�l���i�����j
    varKyotsuCol(19) = varKyotsuRow(25) '��E���E�����i�J�i�j
    varKyotsuCol(20) = varKyotsuRow(26) '��E���E�����i�����j
    varKyotsuCol(21) = varKyotsuRow(27) '�A����P�@����E�g��
    varKyotsuCol(22) = varKyotsuRow(28) '�A����Q�@�Ζ���
    varKyotsuCol(23) = varKyotsuRow(29) '�A����R�@�e�`�w
    varKyotsuCol(24) = varKyotsuRow(30) '�c�̖�
    varKyotsuCol(25) = varKyotsuRow(31) '�c�̃R�[�h
    varKyotsuCol(26) = varKyotsuRow(32) '�c�̈��Ɋւ������
    varKyotsuCol(27) = varKyotsuRow(33) '�����R�[�h
    varKyotsuCol(28) = varKyotsuRow(34) '�Ј��R�[�h
    varKyotsuCol(29) = varKyotsuRow(35) '���ۃR�[�h
    varKyotsuCol(30) = varKyotsuRow(36) '�㗝�X�R�[�h
    varKyotsuCol(31) = varKyotsuRow(37) '�،��ԍ�

    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 2 Then
        varMeisaiRow = Split(varContent(0, 1), ",")
        varKyotsuCol(12) = varMeisaiRow(12) '�m���t���[�g��������
        varKyotsuCol(13) = varMeisaiRow(13) '�c�̊�����
    End If

    '1�s�ڂ̍��ږ�����ɂȂ�܂�2�s�ڂɒl���i�[
    For i = 1 To wstKyotsu.Range("A1").End(xlToRight).Column
        wstKyotsu.Cells(2, i) = varKyotsuCol(i - 1)
    Next i

    '�e�L�X�g�t�@�C���̓��e���V�[�g�ɔ��f
'    If blnChouhyouflg Then
        '���[�o�̓e�L�X�g
        For i = 1 To UBound(varKyotsuRow) + 1
            If i = 19 Then
                varKyotsuRow(i - 1) = intCar
            End If
            Call fncTextEdit(1, i, varKyotsuRow(i - 1), 0)
        Next i

'    End If

    Set wstKyotsu = Nothing

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

