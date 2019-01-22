Attribute VB_Name = "modErrorCheckFunctions"
Option Explicit

''��{�G���[�����Q

Public Function fncNeedCheck(ByVal strValue As String) As String
'�֐����FfncNeedCheck
'���e�@�F�K�{�`�F�b�N
'�����@�F
'        strValue       = ���͓��e

    If IsNull(strValue) Then
        fncNeedCheck = " �K�{���͍��ڂł��B���͂��Ă��������B"
    Else
        If strValue = "" Then
            fncNeedCheck = " �K�{���͍��ڂł��B���͂��Ă��������B"
        Else
            fncNeedCheck = ""
        End If
    End If
End Function


Public Function fncCommaCheck(ByVal strValue As String) As String
'�֐����FfncCommaCheck
'���e�@�F,(�J���})�`�F�b�N
'�����@�F
'        strValue       = ���͓��e

    If strValue Like "*,*" Then
        fncCommaCheck = " �u�C�i�J���}�j�v�͓��͂ł��܂���B"
    Else
        fncCommaCheck = ""
    End If
End Function


Public Function fncDateCheck(ByVal strValue As String, Optional blnWareki As Boolean = False, Optional blnSyodoTouroku As Boolean = False) As String
'�֐����FfncDateCheck
'���e�@�F���t�`�F�b�N
'�����@�F
'        strValue       = ���͓��e
'        blnWareki      = �a��t���O
    If blnWareki Then
    Dim strSeireki As String
        If strValue Like "*���N*" Then
            strValue = Left(strValue, InStr(strValue, "��") - 1) & "1" & Mid(strValue, InStr(strValue, "��") + 1)
        End If
        '�V�����Ή���
        If IsNumeric(Mid(strValue, 3, InStr(strValue, "�N") - 3)) And _
        IsNumeric(Mid(strValue, InStr(strValue, "�N") + 1, InStr(strValue, "��") - InStr(strValue, "�N") - 1)) And _
        (Mid(strValue, InStr(strValue, "�N") + 1, InStr(strValue, "��") - InStr(strValue, "�N") - 1)) >= 1 And _
        (Mid(strValue, InStr(strValue, "�N") + 1, InStr(strValue, "��") - InStr(strValue, "�N") - 1)) <= 12 Then
            If blnSyodoTouroku Then
                If strValue Like "�吳*�N*��" Then
                    If strValue Like "*7��*" Then
                        strValue = strValue & "30��"
                    ElseIf strValue Like "*12��*" Then
                        strValue = strValue & "24��"
                    Else
                        strValue = strValue & "1��"
                    End If
                ElseIf strValue Like "���a*�N*��" Then
                    If strValue Like "*12��*" Then
                        strValue = strValue & "25��"
                    ElseIf strValue Like "*1��*" Then
                        strValue = strValue & "7��"
                    Else
                        strValue = strValue & "1��"
                    End If
                ElseIf strValue Like "����*" Then
                    If strValue Like "*1��*" Then
                        strValue = strValue & "8��"
                    ElseIf strValue Like "*4��*" Then
                        strValue = strValue & "30��"
                    Else
                        strValue = strValue & "1��"
                    End If
                ElseIf strValue Like "�j��*" Then
                    strValue = strValue & "1��"
                Else
                    fncDateCheck = " �N�������m�F�̂����A���������͂��Ă��������B"
                End If
            End If

            If IsNumeric(Mid(strValue, InStr(strValue, "��") + 1, InStr(strValue, "��") - InStr(strValue, "��") - 1)) Then
                If fncDateCheck = "" Then
                    If strValue Like "�吳*�N*��*��" Then
                        If CDate(strValue) >= CDate("1912/07/30") And _
                           CDate(strValue) <= CDate("1926/12/24") Then
                            fncDateCheck = ""
                        Else
                            fncDateCheck = " �N�������m�F�̂����A���������͂��Ă��������B"
                        End If
                    ElseIf strValue Like "���a*�N*��*��" Then
                        If CDate(strValue) >= CDate("1926/12/25") And _
                           CDate(strValue) <= CDate("1989/01/07") Then
                            fncDateCheck = ""
                        Else
                            fncDateCheck = " �N�������m�F�̂����A���������͂��Ă��������B"
                        End If
                    ElseIf strValue Like "����*�N*��*��" Then
                        If CDate(strValue) >= CDate("1989/01/08") And _
                           CDate(strValue) <= CDate("2019/04/30") Then
                            fncDateCheck = ""
                        Else
                            fncDateCheck = " �N�������m�F�̂����A���������͂��Ă��������B"
                        End If
                    ElseIf strValue Like "�j��*�N*��*��" Then
                        strSeireki = Mid(strValue, 3, InStr(strValue, "�N") - 3)
                        strSeireki = strSeireki + 2018 & Mid(strValue, InStr(strValue, "�N"))
                        If CDate(strSeireki) >= CDate("2019/05/01") Then
'                        If fncToSeireki(strValue, 8) >= CDate("2019/05/01") Then
                            fncDateCheck = ""
                        Else
                            fncDateCheck = " �N�������m�F�̂����A���������͂��Ă��������B"
                        End If
                    Else
                        fncDateCheck = " �N�������m�F�̂����A���������͂��Ă��������B"
                    End If
                End If
            Else
                fncDateCheck = " �N�������m�F�̂����A���������͂��Ă��������B"
            End If
        Else
            fncDateCheck = " �N�������m�F�̂����A���������͂��Ă��������B"
        End If
        '�V�����Ή���
    Else
        If IsDate(strValue) = False Then
            fncDateCheck = " �N�������m�F�̂����A���������͂��Ă��������B"
        Else
'            If CDate(strValue) >= "1989/01/08" Then
            If CDate(strValue) >= "1912/07/30" Then
                fncDateCheck = ""
            Else
                fncDateCheck = " �N�������m�F�̂����A���������͂��Ă��������B"
            End If
        End If
    End If
    
End Function


Public Function fncShikiCheck(ByVal strValue As String) As String
'�֐����FfncShikiCheck
'���e�@�F�ی��n���`�F�b�N
'�����@�F
'        strValue       = ���͓��e

    If CDate(strValue) < "2019/01/01" Then
        fncShikiCheck = " �ی��n����2018�N12��31���ȑO�̌_��ɂ͎g�p�ł��܂���B"
    End If

End Function


'Public Function fncNumCheck(ByVal strValue As String) As String
''�֐����FfncNumCheck
''���e�@�F�����`�F�b�N
''�����@�F
''        strValue       = ���͓��e
'
'    If strValue Like "*[!0-9]*" Then
'        fncNumCheck = " �����݂̂���͂��Ă��������B"
'    Else
'        fncNumCheck = ""
'    End If
'End Function
'
Public Function fncDecimalCheck(ByVal strValue As String) As String
'�֐����FfncNumCheck
'���e�@�F�����`�F�b�N(�}�C�i�X�A�����_����)
'�����@�F
'        strValue       = ���͓��e
    If IsNumeric(strValue) Then
        Dim strNum As String
        If strValue = Int(strValue) Then
            strNum = strValue
        Else
            strNum = Replace(strValue, ".", "")
        End If
        If strNum Like "[0-9]*" Or strNum Like "-[0-9]*" Then
            fncDecimalCheck = ""
        Else
            fncDecimalCheck = " ����(���p)�݂̂���͂��Ă��������B"
        End If
    Else
        fncDecimalCheck = " ����(���p)�݂̂���͂��Ă��������B"
    End If
End Function

Public Function fncNumCheck(ByVal strValue As String, Optional intNum As Integer) As String
'�֐����FfncNumCheck
'���e�@�F�����`�F�b�N
'�����@�F
'        strValue       = ���͓��e
    Dim strNum As String
    strNum = strValue
    If intNum = 1 Then
        If IsNumeric(strValue) Then
            '2018/3 ���ذĖ��וt�@�\�ǉ�
            'If strValue = Int(strValue) Then
            If strValue = CStr(Int(strValue)) Then
            Else
                strNum = Replace(strValue, ".", "")
            End If
        End If
    Else
    End If
    If strNum Like "*[!0-9]*" Then
        fncNumCheck = " ����(���p)�݂̂���͂��Ă��������B"
    Else
        fncNumCheck = ""
    End If
    
End Function


Public Function fncNumRangeCheck(ByVal intValue As Double, ByVal intRangeMin As Double, ByVal intRangeMax As Double) As String
'�֐����FfncNumRangeCheck
'���e�@�F���l�`�F�b�N
'�����@�F
'        intValue           = ���͓��e
'        intRabgeMin        = �ŏ��l
'        intRabgeMax        = �ő�l
    
    If intValue < intRangeMin Or intValue > intRangeMax Then
        fncNumRangeCheck = " �w�肳�ꂽ�͈͂̐��l����͂��Ă��������B"
    Else
        fncNumRangeCheck = ""
    End If
End Function


'2018/3 ���ذĖ��וt�@�\�ǉ�
Public Function fncNonfleetSoufuhodaisuCheck(ByVal intValue As Double, ByVal intRangeMax As Double) As String
'�֐����FfncNumRangeCheck
'���e�@�F���t�ۑ䐔�G���[�`�F�b�N
'�����@�F
'        intValue           = ���͓��e
'        intRabgeMin        = �ŏ��l
'        intRabgeMax        = �ő�l
    
    If intValue > intRangeMax Then
        fncNonfleetSoufuhodaisuCheck = " ���t�ۑ䐔��10��ȏ�͓��͂ł��܂���B"
    Else
        fncNonfleetSoufuhodaisuCheck = ""
    End If
End Function


Public Function fncKetaCheck(ByVal strValue As String, ByVal intDigit As Integer, ByVal strType As String) As String
'�֐����FfncKetaCheck
'���e�@�F�����`�F�b�N
'�����@�F
'        intValue           = ���͓��e
'        intDigit           = �������l
'        strType            = ��r���@

    If strType = "<" Then
        If Len(strValue) < intDigit Then
            fncKetaCheck = " ���͌���������܂���B���������͂��Ă��������B"
        Else
            fncKetaCheck = ""
        End If
    ElseIf strType = ">" Then
        If Len(strValue) > intDigit Then
            fncKetaCheck = " ���͂ł��錅���𒴂��Ă��܂��B���������͂��Ă��������B"
        Else
            fncKetaCheck = ""
        End If
    ElseIf strType = "=" Then
        If Not Len(strValue) = intDigit Then
            fncKetaCheck = " �w�肳�ꂽ�����ȊO�͓��͂ł��܂���B"
        Else
            fncKetaCheck = ""
        End If
    End If
End Function


Public Function fncZenkakuCheck(ByVal strValue As String) As String
'�֐����FfncZenkakuCheck
'���e�@�F�S�p�`�F�b�N
'�����@�F
'        strValue           = ���͓��e

    If Len(strValue) * 2 <> LenB(StrConv(strValue, vbFromUnicode)) Then
        fncZenkakuCheck = " �S�p�����݂̂���͂��Ă��������B"
    Else
        fncZenkakuCheck = ""
    End If
End Function


Public Function fncHankakuCheck(ByVal strValue As String) As String
'�֐����FfncHankakuCheck
'���e�@�F���p�p�����`�F�b�N
'�����@�F
'        strValue           = ���͓��e

    If Len(strValue) <> LenB(StrConv(strValue, vbFromUnicode)) Then
        fncHankakuCheck = " ���p�p�����݂̂���͂��Ă��������B"
    Else
        fncHankakuCheck = ""
    End If
End Function


Public Function fncListCheck(ByVal strValue As String, ByVal strListCell As String) As String
'�֐����FfncListCheck
'���e�@�F���X�g�`�F�b�N
'�����@�F
'        strValue           = ���͓��e
'        strListCell        =���X�g�̃Z���ԍ�

    Dim wsCode As Worksheet
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(4, wsCode)         '�V�[�g�I�u�W�F�N�g(�ʎ��@�R�[�h�l)
    Else
        Call subSetSheet(16, wsCode)         '�V�[�g�I�u�W�F�N�g(�ʎ��@�R�[�h�l�i�m���t���[�g�j�j
    End If
    
    If WorksheetFunction.CountIf(wsCode.Range(wsCode.Range(strListCell), wsCode.Cells(wsCode.Rows.Count, _
                                 wsCode.Range(strListCell).Column).End(xlUp)), strValue) = 0 Then
        fncListCheck = " �w�肳�ꂽ�l����͂��Ă��������B"
    Else
        fncListCheck = ""
    End If
    
    Set wsCode = Nothing
    
End Function


''�ʃG���[�����Q

Public Function fncCodeCheck(ByVal strValue As String, ByVal strListCell As String) As String
'�֐����FfncCodeCheck
'���e�@�F�R�[�h�l�`�F�b�N
'�����@�F
'        strValue           = ���͓��e
'        strListCell        =���X�g�̃Z���ԍ�

    Dim wsCode As Worksheet
    '2018/3 ���ذĖ��וt�@�\�ǉ�
    If FleetTypeFlg = 1 Then  '�t���[�g
        Call subSetSheet(4, wsCode)         '�V�[�g�I�u�W�F�N�g(�ʎ��@�R�[�h�l)
    Else
        Call subSetSheet(16, wsCode)         '�V�[�g�I�u�W�F�N�g(�ʎ��@�R�[�h�l�i�m���t���[�g�j�j
    End If
    
    If WorksheetFunction.CountIf(wsCode.Range(wsCode.Range(strListCell), wsCode.Cells(wsCode.Rows.Count, _
                                 wsCode.Range(strListCell).Column).End(xlUp)), strValue) = 0 Then
        fncCodeCheck = " �w�肳�ꂽ�l����͂��Ă��������B"
    Else
        fncCodeCheck = ""
    End If
    
    Set wsCode = Nothing
    
End Function


Public Function fncOneOrBlankCheck(ByVal strValue As String) As String
'�֐����FfncOneOrBlankCheck
'���e�@�F1�܂��̓u�����N�`�F�b�N
'�����@�F
'        strValue           = ���͓��e

    fncOneOrBlankCheck = ""
    If IsNull(strValue) = False Then
        If strValue <> "" Then
            If strValue <> "1" Then
                fncOneOrBlankCheck = " �u1�v�ȊO�͓��͂ł��܂���B"
            End If
        End If
    End If
End Function


Public Function fncWarekiCheck(ByVal strValue As String, Optional intKeta As Integer) As String
'�֐����FfncWarekiCheck
'���e�@�F�a��`�F�b�N
'�����@�F
'        strValue           = ���͓��e

    If strValue Like "�吳*�N*��*��" Then
        fncWarekiCheck = ""
    ElseIf strValue Like "���a*�N*��*��" Then
        fncWarekiCheck = ""
    ElseIf strValue Like "����*�N*��*��" Then
        fncWarekiCheck = ""
    ElseIf strValue Like "�j��*�N*��*��" Then
        fncWarekiCheck = ""
    Else
        If intKeta = 6 Then
            fncWarekiCheck = " �a��œ��͂��Ă��������B�i����X�NX���j"
        Else
            fncWarekiCheck = " �a��œ��͂��Ă��������B�i����X�NX��X���j"
        End If
    End If
End Function


''�֘A�G���[�����Q

Public Function fncHknSyuruiCheck(ByVal strHknSyurui As String, ByVal strSyaryoMskGaku As String, ByVal strDaisyaToku As String, _
                                  ByVal strHknZnsnToku As String, ByVal strSyaryoTonanToku As String, ByVal strSyaryoTyoukaToku As String, _
                                  Optional ByVal strHknKingaku As String = "NULL") As String
'�֐����FfncHknSyuruiCheck
'���e�@�F�ԗ��ی��G���[
'�����@�F
'        strHknSyurui           =�ی��̎��
'        strHknKingaku          =�ԗ��ی����z
'        strSyaryoMskGaku       =�ԗ��ƐӋ��z
'        strDaisyaToku          =��ԓ��Z�b�g����
'        strHknZnsnToku         =�ԗ��S���Ք����
'        strSyaryoTonanToku     =�ԗ�����ΏۊO����
'        strSyaryoTyoukaToku    =�ԗ����ߏC����p����
    
    Dim blnErrFlg As Boolean
    blnErrFlg = False
    
    fncHknSyuruiCheck = ""
    If strHknSyurui = "" Then
        If strHknKingaku <> "NULL" Then
            If (strHknKingaku <> "" Or strSyaryoMskGaku <> "" Or strDaisyaToku <> "") Or _
               (CBool(fncTekiyouChange(strHknZnsnToku)) = True Or CBool(fncTaisyoChange(strSyaryoTonanToku)) = True Or _
                CBool(fncTekiyouChange(strSyaryoTyoukaToku)) = True) Then
                
                blnErrFlg = True
            End If
        Else
            If (strSyaryoMskGaku <> "" Or strDaisyaToku <> "") Or _
               (CBool(fncTekiyouChange(strHknZnsnToku)) = True Or CBool(fncTaisyoChange(strSyaryoTonanToku)) = True Or _
                CBool(fncTekiyouChange(strSyaryoTyoukaToku)) = True) Then
                
                blnErrFlg = True
            End If
        End If
        
        If blnErrFlg Then
            fncHknSyuruiCheck = " �ԗ��ی��̎�ނ�I�����Ă��������B"
        End If
    Else
        If strHknKingaku <> "NULL" Then
            If strHknKingaku = "" Or strSyaryoMskGaku = "" Then
                fncHknSyuruiCheck = " �ԗ��ی����z����юԗ��ƐӋ��z����͂��Ă��������B"
            End If
        Else
            If strSyaryoMskGaku = "" Then
                fncHknSyuruiCheck = " �ԗ��ƐӋ��z����͂��Ă��������B"
            End If
        End If
        
    End If
End Function

Public Function fncTaijinBaisyoCheck(ByVal strTaijinBaisyo As String, ByVal strJisonJikoToku As String, _
                                     ByVal strZinshinSyougai_1Mei As String, ByVal strMuhokenToku As String) As String
'�֐����FfncTaijinBaisyoCheck
'���e�@�F�ΐl�����G���[
'�����@�F
'        strTaijinBaisyo            =�ΐl����
'        strJisonJikoToku           =�������̏��Q����
'        strZinshinSyougai_1Mei     =�l�g���Q(1��)
'        strDaisyaToku              =��ԓ��Z�b�g����
'        strMuhokenToku             =���ی��Ԏ��̏��Q����

    fncTaijinBaisyoCheck = ""
    If strTaijinBaisyo = "" Then
'        fncTaijinBaisyoCheck = " �ΐl������I�����Ă��������B"
    ElseIf CBool(fncTaisyoChange(strTaijinBaisyo)) Then
        If CBool(fncTaisyoChange(strJisonJikoToku)) Then
            fncTaijinBaisyoCheck = " �ΐl�ΏۊO�̂Ƃ��������̑ΏۊO����͕t�тł��܂���B"
        End If
        If CBool(fncTaisyoChange(strMuhokenToku)) Then
            If Not fncTaijinBaisyoCheck = "" Then
                fncTaijinBaisyoCheck = fncTaijinBaisyoCheck & vbCrLf & " �ΐl�ΏۊO�̂Ƃ����ی��ԑΏۊO����͕t�тł��܂���B"
            Else
                fncTaijinBaisyoCheck = " �ΐl�ΏۊO�̂Ƃ����ی��ԑΏۊO����͕t�тł��܂���B"
            End If
        End If
    End If
    
    If CBool(fncTaisyoChange(strZinshinSyougai_1Mei)) = False And CBool(fncTaisyoChange(strJisonJikoToku)) Then
        If strZinshinSyougai_1Mei <> "" Then
            If fncTaijinBaisyoCheck = "" Then
                fncTaijinBaisyoCheck = " �������̑ΏۊO�����t�т���ꍇ�A�l�g���Q�͕t�тł��܂���B"
            Else
                fncTaijinBaisyoCheck = fncTaijinBaisyoCheck & vbCrLf & " �������̑ΏۊO�����t�т���ꍇ�A�l�g���Q�͕t�тł��܂���B"
            End If
        End If
    End If
    
End Function


Public Function fncTaibutsuBaisyo(ByVal strTaibutsuBaisyo As String, ByVal strTaibutsuMskGaku As String, _
                                  ByVal strTaibutsuTyoukaToku As String) As String
'�֐����FfncTaibutsuBaisyo
'���e�@�F�Ε������G���[
'�����@�F
'        strTaibutsuBaisyo              =�Ε�����
'        strTaibutsuMskGaku             =�Ε��ƐӋ��z
'        strTaibutsuTyoukaToku          =�Ε����ߏC����p����

    fncTaibutsuBaisyo = ""
    If strTaibutsuBaisyo = "" Then
'        fncTaibutsuBaisyo = " �Ε�������I�����Ă��������B"
    Else
        If strTaibutsuBaisyo = "�ΏۊO" And (strTaibutsuMskGaku <> "" Or CBool(fncTekiyouChange(strTaibutsuTyoukaToku)) = True) Then
            fncTaibutsuBaisyo = " �Ε��ΏۊO�̂Ƃ��A�Ε��ƐӋ��z����ёΕ����ߏC����p����͓��͕s�v�ł��B"
        ElseIf strTaibutsuBaisyo <> "�ΏۊO" And strTaibutsuMskGaku = "" Then
            fncTaibutsuBaisyo = " �Ε��ƐӋ��z��I�����Ă��������B"
        End If
    End If
End Function


Public Function fncZinshinSyougai(ByVal strZinshinSyougai_1Mei As String, ByVal strZinshinSyougai_1Jiko As String) As String
'�֐����FfncZinshinSyougai
'���e�@�F�l�g���Q�G���[
'�����@�F
'        strZinshinSyougai_1Mei             =�l�g���Q(1��)
'        strZinshinSyougai_1Jiko            =�l�g���Q(1����)

    fncZinshinSyougai = ""
    If strZinshinSyougai_1Mei = "" Then
'        fncZinshinSyougai = " �l�g���Q��I�����Ă��������B"
    ElseIf strZinshinSyougai_1Mei = "�ΏۊO" And strZinshinSyougai_1Jiko <> "" Then
        fncZinshinSyougai = " �l���ΏۊO�̂Ƃ��A�l�g���Q�i�P���́j�͓��͕s�v�ł��B"
    End If
End Function


Public Function fncTouzyouSyougai(ByVal strTouzyouSyougai_1Mei As String, ByVal strTouzyouSyougai_1Mei_Taisyougai As String, _
                                  ByVal strTouzyouSyougai_1Jiko As String, ByVal strNissuToku As String)
'�֐����FfncTouzyouSyougai
'���e�@�F����ҏ��Q�G���[
'�����@�F
'        strTouzyouSyougai_1Mei             =����ҏ��Q(1��)
'        strTouzyouSyougai_1Mei_Taisyougai            =����ҏ��Q(1����)
'        strNissuToku                       =����������

    fncTouzyouSyougai = ""
    If strTouzyouSyougai_1Mei = "" Then
        If strTouzyouSyougai_1Mei_Taisyougai = False Then
            fncTouzyouSyougai = " ����ҏ��Q�i�P���j�ی����z���ΏۊO�̂����ꂩ����͂��Ă��������B"
        Else
            If strTouzyouSyougai_1Jiko <> "" Or CBool(fncTekiyouChange(strNissuToku)) Then
                fncTouzyouSyougai = " �����ΏۊO�̂Ƃ��A����ҏ��Q�i�P���́j����ѓ���������͓��͕s�v�ł��B"
            End If
        End If
'        fncTouzyouSyougai = " ����ҏ��Q��I�����Ă��������B"
    ElseIf IsNumeric(strTouzyouSyougai_1Mei) Then
        If strTouzyouSyougai_1Mei Mod 100 <> 0 Or strTouzyouSyougai_1Mei = 0 Then
            fncTouzyouSyougai = " ����ҏ��Q�i�P���j��100���~�P�ʂœ��͂��Ă��������B"
        ElseIf strTouzyouSyougai_1Mei > 5000 Then
            fncTouzyouSyougai = " ����ҏ��Q�i�P���j��5,000���~�ȉ��œ��͂��Ă��������B"
        ElseIf strTouzyouSyougai_1Mei_Taisyougai Then
            fncTouzyouSyougai = " ����ҏ��Q�i�P���j�ی����z���ΏۊO�̂����ꂩ����͂��Ă��������B"
        End If
    Else
        fncTouzyouSyougai = " ����ҏ��Q�i�P���j�ی����z�ɂ͐����̂ݓ��͂��Ă��������B"
    End If

End Function


Public Function fncTekiyouChange(ByVal strValue As String) As String
'�֐����FfncTekiyouChange
'���e�@�F���͍��ڂ��u�K�p����v�̂��̂�True/False�ɕϊ�
'�����@�F
'        strValue           = ���͓��e

    If strValue = "�K�p����" Or strValue = "True" Then
        fncTekiyouChange = "True"
    Else
        fncTekiyouChange = "False"
    End If
End Function


Public Function fncTaisyoChange(ByVal strValue As String) As String
'�֐����FfncTaisyoChange
'���e�@�F���͍��ڂ��u�ΏۊO�v�̂��̂�True/False�ɕϊ�
'�����@�F
'        strValue           = ���͓��e

    If strValue = "�ΏۊO" Or strValue = "True" Then
        fncTaisyoChange = "True"
    Else
        fncTaisyoChange = "False"
    End If
End Function


Public Function fncGenteiChange(ByVal strValue As String) As String
'�֐����FfncGenteiChange
'���e�@�F���͍��ڂ��u����v�̂��̂�True/False�ɕϊ�
'�����@�F
'        strValue           = ���͓��e

    If strValue = "����" Or strValue = "True" Then
        fncGenteiChange = "True"
    Else
        fncGenteiChange = "False"
    End If
End Function


'2018/3 ���ذĖ��וt�@�\�ǉ�
Public Function fncHokenSyuruiCheck(ByVal strHokenSyurui As String, ByVal strHoujin As String) As String
'�֐����FfncNonfleetTawariCheck
'���e�@�F�ی���ރG���[
'�����@�F
'        strHokenSyurui = �ی����
'        strKojin = ��ی���

    If strHokenSyurui = "�l�p���������ԕی�" And strHoujin = "True" Then
        fncHokenSyuruiCheck = " �l�p���������ԕی��͌l�_��݂̂ɂȂ�܂��B"
    Else
        fncHokenSyuruiCheck = ""
    End If
End Function


'2018/3 ���ذĖ��וt�@�\�ǉ�
Public Function fncNonfleetTawariCheck(ByVal strNonfleetTawari As String, ByVal strSouFuhoDaisu As String) As String
'�֐����FfncNonfleetTawariCheck
'���e�@�F���ذđ��������G���[
'�����@�F
'        strNonfleetTawari = ���ذđ�������
'        strSouFuhoDaisu = ���t�ۑ䐔

    fncNonfleetTawariCheck = ""
    If IsNumeric(strSouFuhoDaisu) Then
        If strSouFuhoDaisu >= 3 Then
            If strNonfleetTawari = "" Then
                fncNonfleetTawariCheck = " ���t�ۑ䐔���R��ȏ�̏ꍇ�A�m���t���[�g�����������K�p�ł��܂��B"
            End If
        End If
    End If
End Function


'2018/3 ���ذĖ��וt�@�\�ǉ�
Public Function fncFamilyBikeCheck(ByVal strFamilyBike As String, ByVal strJinshinSyougaiIchimei As String) As String
'�֐����FfncNonfleetTawariCheck
'���e�@�F�t�@�~���[�o�C�N����G���[
'�����@�F
'        strFamilyBike = �t�@�~���[�o�C�N����
'        strJinshinSyougaiIchimei = �l�g���Q�i1���j

    If strFamilyBike = "�l�g" And strJinshinSyougaiIchimei = "�ΏۊO" Then
        fncFamilyBikeCheck = " �l���ΏۊO�̂Ƃ��A�t�@�~���[�o�C�N����u�l�g�v�͑I���ł��܂���B"
    Else
        fncFamilyBikeCheck = ""
    End If
End Function


'2018/3 ���ذĖ��וt�@�\�ǉ�
Public Function fncJidoushaJikoCheck(ByVal strJidoushaJiko As String, ByVal strJinshinSyougaiIchimei As String) As String
'�֐����FfncNonfleetTawariCheck
'���e�@�F�����Ԏ��̓���G���[
'�����@�F
'        strJidoushaJiko = �����Ԏ��̓���
'        strJinshinSyougaiIchimei = �l�g���Q�i1���j

    If (strJidoushaJiko = "True" Or strJidoushaJiko = "�K�p����") And strJinshinSyougaiIchimei = "�ΏۊO" Then
        fncJidoushaJikoCheck = " �l���ΏۊO�̂Ƃ��A�����Ԏ��̓���͕t�тł��܂���B"
    Else
        fncJidoushaJikoCheck = ""
    End If
End Function


'2018/3 ���ذĖ��וt�@�\�ǉ�
Public Function fncToukyuCheck1(ByVal strUketsukeKbn As String, ByVal strNonFleetToukyu As String) As String
'�֐����FfncToukyuCheck1
'���e�@�F�����G���[�P
'�����@�F
'        strUketsukeKbn = ��t�敪
'        strNonFleetToukyu = �m���t���[�g����

    If strUketsukeKbn = "1" And strNonFleetToukyu <> "�U�r����" Then
        fncToukyuCheck1 = " ��t�敪���V�K�̂Ƃ��A6S�����ȊO�͑I���ł��܂���B"
    Else
        fncToukyuCheck1 = ""
    End If
End Function


'2018/3 ���ذĖ��וt�@�\�ǉ�
Public Function fncToukyuCheck2(ByVal strUketsukeKbn As String, ByVal strNonFleetToukyu As String) As String
'�֐����FfncToukyuCheck2
'���e�@�F�����G���[�Q
'�����@�F
'        strUketsukeKbn = ��t�敪
'        strNonFleetToukyu = �m���t���[�g����

    If strUketsukeKbn = "3" And (strNonFleetToukyu = "�U�r����" Or strNonFleetToukyu = "�V�r����") Then
        fncToukyuCheck2 = " ��t�敪���p���̂Ƃ��A6S�����A7S�����͑I���ł��܂���B"
    Else
        fncToukyuCheck2 = ""
    End If
End Function



