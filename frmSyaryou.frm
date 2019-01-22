VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSyaryou 
   Caption         =   "�ԗ����捞"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9240
   OleObjectBlob   =   "frmSyaryou.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSyaryou"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strBeforeFileName As String
Dim intConfirmMsg As Integer


'�����\��
'2018/3 ���ذĖ��וt�@�\�ǉ�
Public Sub UserForm_Initialize()
    '�t�H�[���̃^�C�g���ݒ�
    If FleetTypeFlg = 1 Then
        frmSyaryou.Caption = frmSyaryou.Caption & "�i�t���[�g�_��j"
    Else
        frmSyaryou.Caption = frmSyaryou.Caption & "�i�m���t���[�g���וt�_��j"
    End If
End Sub


'�u�I���v�{�^������
Private Sub btnSelect_Click()
    Dim strBeforeFilePath As String
    
    On Error GoTo Error
    
    '�I������Ă����t�@�C���̃p�X�E�t�@�C�������擾
    strBeforeFilePath = frmSyaryou.txtFilePath
    strBeforeFileName = Mid(strBeforeFilePath, InStrRev(strBeforeFilePath, "\") + 1)
    
    '�t�@�C���I���_�C�A���O
    Dim strFilePath As String
    strFilePath = Application.GetOpenFilename(FileFilter:="Excel�t�@�C��,*.xlsx;*.xls;*.xlsm")
    
    If strFilePath <> "False" Then
    
        '�u�捞�s���v�擾
        Dim wbWorkBook As Workbook
        Dim wsWorkSheet As Worksheet
        Dim strFileName As String
        
        strFileName = Mid(strFilePath, InStrRev(strFilePath, "\") + 1)
         
        Dim blnFileFlg As Boolean
        Dim WB As Workbook
        For Each WB In Workbooks
            If WB.Name = strFileName Then
                blnFileFlg = True
            End If
        Next WB
        
        Set WB = Nothing
        
        If blnFileFlg Then
            Set wbWorkBook = Workbooks(strFileName)
        Else
            Set wbWorkBook = Workbooks.Open(strFilePath)
            Windows(wbWorkBook.Name).Visible = False
            
            Call subFileClose(strBeforeFileName)
        End If
        Set wsWorkSheet = wbWorkBook.Sheets(1)
        
        Dim i As Long
        Dim lngTmpRow As Long
        Dim lngMaxRow As Long
        
        '�ԗ����t�@�C���Ō��܂Ŏ擾�i�t���[�g:12��A�m���t���[�g:12��j
            For i = 1 To 12
                lngTmpRow = wsWorkSheet.Cells(wsWorkSheet.Rows.Count, i).End(xlUp).Row
                If lngTmpRow > lngMaxRow Then
                    lngMaxRow = lngTmpRow
                End If
            Next
        
        Set wbWorkBook = Nothing
        Set wsWorkSheet = Nothing
        
        '�t�H�[���ݒ�
        frmSyaryou.txtFilePath = strFilePath
        frmSyaryou.txtFilePath.SetFocus
        frmSyaryou.txtFilePath.SelStart = 0
        frmSyaryou.lblMaxRow = "�捞�䐔�F" & CStr(lngMaxRow - 1) & "��"
    
    End If

    Exit Sub
Error:
     MsgBox "btnSelect_Click" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["
    
End Sub


'�u�ꊇ�Z�b�g�v�{�^������
Private Sub btnSyaryouSet_Click()
    
    '�G���[�`�F�b�N
    Dim strMaxRow As String
    Dim strMaxMeisai As String
    Dim intConfirmMsg As Integer
    
    On Error GoTo Error
    
    strMaxMeisai = Replace(lblMaxMeisai, "���t�ۑ䐔�F", "")
    strMaxMeisai = Replace(strMaxMeisai, "��", "")
    strMaxRow = Replace(lblMaxRow, "�捞�䐔�F", "")
    strMaxRow = Replace(strMaxRow, "��", "")
    
    If Val(strMaxRow) > Val(strMaxMeisai) Then
        intConfirmMsg = MsgBox("�捞�䐔�����t�ۑ䐔���������ł��B", vbOKOnly + vbExclamation, "�G���[�_�C�A���O")
        Exit Sub
    End If
    
    intConfirmMsg = MsgBox("�ԗ�������荞�݂܂��B" & vbCrLf & "��낵���ł���?", vbYesNo, "�m�F�_�C�A���O")
    If intConfirmMsg = 6 Then
        
        '�捞�f�[�^�̏�������
        Dim wbWorkBook As Workbook
        Dim wsWorkSheet As Worksheet
        
        Dim blnFileFlg As Boolean
        Dim WB As Workbook
        Dim strFilePath As String
        Dim strFileName As String
        
        strFilePath = frmSyaryou.txtFilePath
        
        strFileName = Mid(strFilePath, InStrRev(strFilePath, "\") + 1)
        
        For Each WB In Workbooks
            If WB.Name = strFileName Then
                blnFileFlg = True
            End If
        Next WB
        
        Set WB = Nothing
        
        If blnFileFlg Then
            Set wbWorkBook = Workbooks(strFileName)
        Else
            Set wbWorkBook = Workbooks.Open(txtFilePath)
            Windows(wbWorkBook.Name).Visible = False
            
            Call subFileClose(strBeforeFileName)
        End If
        Set wsWorkSheet = wbWorkBook.Sheets(1)
    
        Dim wsMeisai As Worksheet
        '2018/3 ���ذĖ��וt�@�\�ǉ�
        If FleetTypeFlg = 1 Then  '�t���[�g
            Call subSetSheet(1, wsMeisai)       '�V�[�g�I�u�W�F�N�g(���ד���)
        Else
            Call subSetSheet(17, wsMeisai)      '�V�[�g�I�u�W�F�N�g(���ד��́i�m���t���[�g�j)
        End If
    
        Call subMeisaiUnProtect     '�V�[�g�̕ی�̉���
    
        Dim i As Long
        For i = 0 To strMaxRow - 1
            wsMeisai.Cells(21 + i, 3) = wsWorkSheet.Cells(2 + i, 1)  '�Ԍ�������
            wsMeisai.Cells(21 + i, 5) = wsWorkSheet.Cells(2 + i, 2)   '�Ԗ�
            wsMeisai.Cells(21 + i, 6) = wsWorkSheet.Cells(2 + i, 3)   '�o�^�ԍ�
            wsMeisai.Cells(21 + i, 7) = wsWorkSheet.Cells(2 + i, 4)   '�ԑ�ԍ�
            wsMeisai.Cells(21 + i, 8) = wsWorkSheet.Cells(2 + i, 5)   '�^��
            wsMeisai.Cells(21 + i, 9) = wsWorkSheet.Cells(2 + i, 6)   '�d�l
            wsMeisai.Cells(21 + i, 10) = wsWorkSheet.Cells(2 + i, 7)  '���x�o�^
            wsMeisai.Cells(21 + i, 11) = wsWorkSheet.Cells(2 + i, 8)  '�Ԍ�������
            wsMeisai.Cells(21 + i, 12) = wsWorkSheet.Cells(2 + i, 9)  '�����E�s����
            wsMeisai.Cells(21 + i, 13) = wsWorkSheet.Cells(2 + i, 10) '�r�C��
            wsMeisai.Cells(21 + i, 14) = wsWorkSheet.Cells(2 + i, 11) '2.5���b�g�����f�B�[�[��������
            If FleetTypeFlg = 1 Then
                wsMeisai.Cells(21 + i, 26) = wsWorkSheet.Cells(2 + i, 12) '�ԗ��ی����z
            ElseIf FleetTypeFlg = 2 Then
                wsMeisai.Cells(21 + i, 34) = wsWorkSheet.Cells(2 + i, 12) '�ԗ��ی����z
            End If
        Next
        
        If Windows(strFileName).Visible = False Then
            subFileClose (strFileName)
        End If
        
        '�V�[�g�E�u�b�N�̕\��
        Call subBookUnProtect '�u�b�N�̕ی������ 2018/3 ���ذĖ��וt�@�\�ǉ�
        Call subSheetVisible(True)
        Call subBookProtect   '�u�b�N�̕ی� 2018/3 ���ذĖ��וt�@�\�ǉ�
        
        '���ד��͉�ʂ̃G���[�p���X�g������
        wsMeisai.OLEObjects("txtErrMsg").Object.Value = ""
        wsMeisai.OLEObjects("txtErrMsg").Activate
        wsMeisai.Range("A1").Activate
        
        Set wbWorkBook = Nothing
        Set wsWorkSheet = Nothing
        Set wsMeisai = Nothing
        
        Call subMeisaiProtect     '�V�[�g�̕ی�
            
        Unload Me
    Else
    End If
    
    Exit Sub
    
Error:
     MsgBox "btnSyaryouSet_Click" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["

    
End Sub


'�u�߂�v�{�^������
Private Sub BtnBack_Click()

    On Error GoTo Error
    
    Dim intMsgBox As Integer
    intMsgBox = MsgBox("�ԗ����t�@�C�����e�𔽉f�����ɖ��ד��͉�ʂɑJ�ڂ��܂��B" & vbCrLf & "��낵���ł����H", vbYesNo, "�m�F�_�C�A���O")
    
    Dim strFileName As String

    strFileName = Mid(frmSyaryou.txtFilePath, InStrRev(frmSyaryou.txtFilePath, "\") + 1)
    If intMsgBox = 6 Then
        
        If Windows(strFileName).Visible = False Then
            subFileClose (strFileName)
        End If

        '�V�[�g�E�u�b�N�̕\��
        Call subBookUnProtect '�u�b�N�̕ی������ 2018/3 ���ذĖ��וt�@�\�ǉ�
        Call subSheetVisible(True)
        Call subBookProtect   '�u�b�N�̕ی� 2018/3 ���ذĖ��וt�@�\�ǉ�
            
        Unload Me
    End If
    
    Exit Sub
Error:
     MsgBox "BtnBack_Click" & vbCrLf & _
            "�G���[�ԍ�:" & Err.Number & vbCrLf & _
            "�G���[�̎��:" & Err.Description, vbExclamation, "�\�����ʃG���["
            
End Sub


'�u�~�v�{�^������
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = vbFormControlMenu Then
        intConfirmMsg = MsgBox("�c�[�����I�����܂��B" & vbCrLf & "��낵���ł���?" & vbCrLf & "�����͓��e�͕ۑ�����܂���B", vbYesNo, "�m�F�_�C�A���O")
        If intConfirmMsg = 6 Then
            Dim strFileName As String
            
            strFileName = Mid(frmSyaryou.txtFilePath, InStrRev(frmSyaryou.txtFilePath, "\") + 1)
            If Windows(strFileName).Visible = False Then
                subFileClose (strFileName)
            End If
            
            Cancel = False
            Call subAppClose
        Else
            Cancel = True
        End If
    End If
    
End Sub


'�I�����ĊJ�����t�@�C�������
Public Sub subFileClose(ByVal strCloseFileName As String)
    Dim sOpenbookAll As Workbook
    Dim sOpenBookSub As Variant
    Dim blnOpenBookFlg As Boolean

    '�I�������G�N�Z���t�@�C�����܂��J����Ă��邱�Ƃ��m�F
    For Each sOpenbookAll In Workbooks
        If sOpenbookAll.Name = strCloseFileName Then
            blnOpenBookFlg = True
            Exit For
        End If
    Next sOpenbookAll
    
    If blnOpenBookFlg Then
        '�G�N�Z���t�@�C�������
        Application.DisplayAlerts = False
        Workbooks(strCloseFileName).Close
        Application.DisplayAlerts = True
    End If
    
    Set sOpenbookAll = Nothing
    Set sOpenBookSub = Nothing
    
End Sub


