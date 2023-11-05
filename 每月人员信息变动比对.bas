Attribute VB_Name = "��Ա��Ϣ�ȶ�ģ��"
'ljffytax 0.0.2 2023-11-05
'GPL v2

'�����ļ���һ����ITS��ԭʼ����Ա�嵥������һ���ǵ���Ҫ�걨����Ա���ܱ�-�����ṩ
'�Ա������ļ������Ա��ITS�ﲻ���ڶ����ܱ�������򽫸ò�����Ա׷�ӵ�ITSβ����֮��
'��ITS�����Ա��עΪ�����������ҽ���������Ա����˹������ְ����


'��ȡ��ǰĿ¼�ڵĹ�˾����
Function getCompanyName() As String
    Dim st As String
    st = getITSFile()
    st = Left(st, Application.Find(".", st) - 1)
    getCompanyName = st
End Function

Function getITSFile() As String
    Dim st As String
    Dim path As String
    
    path = ThisWorkbook.path & "\"
    st = Dir(path)
    Do While st <> "" And 0 = InStr(1, st, "��˾", 1)
        st = Dir
        'Debug.Print (InStr(1, st, "��˾", 1))
    Loop
    getITSFile = st
End Function

Function personID2Birthday(PersonID As String) As String
    Dim m As String
    If Len(PersonID) = 18 Then
        PersonID = Mid(PersonID, 7, 8)
        m = "-" & Mid(PersonID, 5, 2) & "-"
        personID2Birthday = Left(PersonID, 4) & m & Right(PersonID, 2)
    Else
        personID2Birthday = ""
    End If
End Function

'���ҶԱ��������������֤���Ų���ITS��Ա��������ڹ��ʻ��ܱ������ʾ����ԱΪ����ְ��Ա
'������֤������ITS��Ա���ﲢ��״̬��Ϊ"����",���ʾ��Ա��������ְ������������ְ�ˣ�������Ա
'�ᱻ��ǳ�ǳ��ɫ����״̬�޸�Ϊ������
Function employedPerson(smWorkbook As Workbook, itsWorkbook As Workbook) As Integer
    'Set smWorkbook = Workbooks.Open(ThisWorkbook.path & "\����.xlsm")
    'Set itsWorkbook = Workbooks.Open(ThisWorkbook.path & "\" & getITSFile())
    Dim PersonID As String
    Dim iPerson As Integer
    Dim companyName As String
    Dim itsPersonID(29999) As String
    
    itsWorkbook.Worksheets(1).Columns("D:D").NumberFormatLocal = "@" '�������֤��Ϊ�ı���ʽ
    itsWorkbook.Worksheets(1).Columns("J:J").NumberFormatLocal = "@"
    itsWorkbook.Worksheets(1).Columns("K:L").NumberFormatLocal = "yyyy-mm-dd"
    itsWorkbook.Worksheets(1).Columns("G:G").NumberFormatLocal = "yyyy-mm-dd"
    companyName = getCompanyName()
    iPerson = 0
    
    '��ȡITS��Ա����ȫ����Ա���֤����
    With itsWorkbook.Worksheets(1)
        For i = 2 To 30000
            If (.Cells(i, 34) = "��Ա") Then
                itsPersonID(iPerson) = .Cells(i, 4)
                iPerson = iPerson + 1
            ElseIf (.Cells(i, 34) = "����") Then
                itsPersonID(iPerson) = "XXXX-XXXX"
                iPerson = iPerson + 1
            Else
                Exit For
            End If
        Next i
    End With
    
    iPerson = 0
    For iSM = 2 To 30000
        If (smWorkbook.Worksheets(1).Cells(iSM, 28) = companyName) Then
            PersonID = smWorkbook.Worksheets(1).Cells(iSM, 5) '���֤������
            If (PersonID = "") Then
                Exit For
            End If
            For iITS = 2 To 30000 '���ֻ�ȶԵ�3���ˣ���������Ч�ʵ���
                'If (0 = StrComp(itsWorkbook.Worksheets(1).Cells(iITS, 4), PersonID, vbTextCompare)) Then
                If (0 = StrComp(itsPersonID(iITS - 2), PersonID, vbTextCompare)) Then
                    If (itsWorkbook.Worksheets(1).Cells(iITS, 7) = "����") Then
                        Exit For
                    Else
                        itsWorkbook.Worksheets(1).Cells(iITS, 2).Interior.Color = RGB(110, 208, 255) '����ְ��Ա�������ǳ��ɫ
                        itsWorkbook.Worksheets(1).Cells(iITS, 7) = "����" '�޸�״̬Ϊ����
                        Exit For
                    End If
                Else
                     '�ұ���û�У���׷�ӵ�ĩβ��Ϊ��ֹ���м���ֿ���������������ȡ������������
                    'If (itsWorkbook.Worksheets(1).Cells(iITS, 4) = "" And itsWorkbook.Worksheets(1).Cells(iITS + 1, 4) = "") Then
                    If (itsPersonID(iITS - 2) = "" And itsPersonID(iITS - 1) = "") Then
                        itsPersonID(iITS - 2) = PersonID '������Ա׷�ӵ�����β����ʹ��������ITS��������һ��
                        With itsWorkbook.Worksheets(1)
                            .Cells(iITS, 2) = smWorkbook.Worksheets(1).Cells(iSM, 2) '����
                            .Cells(iITS, 3) = "�������֤"
                            .Cells(iITS, 4) = PersonID
                            .Cells(iITS, 5) = smWorkbook.Worksheets(1).Cells(iSM, 4) '�Ա�
                            .Cells(iITS, 6) = personID2Birthday(smWorkbook.Worksheets(1).Cells(iSM, 5)) '��������
                            .Cells(iITS, 7) = "����"
                            .Cells(iITS, 10) = smWorkbook.Worksheets(1).Cells(iSM, 29) '�绰
                            .Cells(iITS, 11) = smWorkbook.Worksheets(1).Cells(iSM, 8) '��ְ����
                            .Cells(iITS, 13) = "��" '�Ƿ��������
                            .Cells(iITS, 31) = "��"
                            .Cells(iITS, 32) = "��"
                            .Cells(iITS, 33) = "��"
                            .Cells(iITS, 34) = "��Ա"
                            .Cells(iITS, 36) = "�й�"
                        End With
                        iPerson = iPerson + 1
                        Exit For
                    End If
                End If
            Next iITS
        End If
    Next iSM
    employedPerson = iPerson
End Function

Function firedPerson(smWorkbook As Workbook, itsWorkbook As Workbook) As Integer
    Dim PersonID As String
    Dim iPerson As Integer
    Dim companyName As String
    Dim smPersonID(29999) As String
    Dim iSMPerson As Integer
    Dim itsPersonID As String
    
    iSMPerson = 0
    companyName = getCompanyName()
    
    '��ȡ���ܱ���ù�˾��ȫ����Ա���֤����
    With smWorkbook.Worksheets(1)
        For i = 2 To 30000
            If (.Cells(i, 28) = companyName) Then
                smPersonID(iSMPerson) = .Cells(i, 5)
                iSMPerson = iSMPerson + 1
            ElseIf (.Cells(i, 28) = "") Then
                Exit For
            End If
        Next i
    End With
    iPerson = 0
    'This is really bad,shit...
    For iITS = 2 To 30000
        itsPersonID = itsWorkbook.Worksheets(1).Cells(iITS, 4)
        If (itsWorkbook.Worksheets(1).Cells(iITS, 7) = "����" And itsWorkbook.Worksheets(1).Cells(iITS, 34) = "��Ա") Then
            For iSM = 0 To iSMPerson
                PersonID = smPersonID(iSM) '���֤������
                If (PersonID = "") Then '�����Ѿ��ҵ���β
                    itsWorkbook.Worksheets(1).Cells(iITS, 2).Interior.Color = RGB(255, 0, 0) '���������ɫ
                    itsWorkbook.Worksheets(1).Cells(iITS, 7) = "������"
                    iPerson = iPerson + 1
                    Exit For
                End If
                If (0 = StrComp(PersonID, itsPersonID, vbTextCompare)) Then
                    Exit For
                End If
            Next iSM
        ElseIf (itsWorkbook.Worksheets(1).Cells(iITS, 4) = "") Then '�ѵ���β
            Exit For
        End If
    Next iITS
    firedPerson = iPerson
End Function


Sub main()
    'r = MsgBox(getCompanyName(), vbOKOnly)
    ChDir (ThisWorkbook.path)
    Dim smWorkbook As Workbook
    Dim itsWorkbook As Workbook
    Set smWorkbook = Workbooks.Open(ThisWorkbook.path & "\����.xlsm")
    Set itsWorkbook = Workbooks.Open(ThisWorkbook.path & "\" & getITSFile())
    Application.ScreenUpdating = False
    iFi = firedPerson(smWorkbook, itsWorkbook)
    iEm = employedPerson(smWorkbook, itsWorkbook)
    Application.ScreenUpdating = ture
    smWorkbook.Close savechanges:=False
    r = MsgBox("��ְ" & CStr(iFi) & "�ˣ�����ְ" & CStr(iEm) & "��", vbOKOnly)
End Sub

