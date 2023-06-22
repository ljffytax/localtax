Attribute VB_Name = "ģ��1"
'Author ljffytax 2023-06-22 V0.01
'GNU GPLv2
'���ʱ��ͷ�ֶΣ�
'���|����|����|�Ա�|���֤����|����|��λ|��ְ����|��ְ����|˰ǰ����|���ϱ���|ҽ�Ʊ���|ʧҵ����|�б���|
'�����籣����|������|�ۼ���Ů����|�ۼƼ�������|�ۼ�ס��������Ϣ|�ۼ�ס�����|�ۼ���������|�ۼ�Ӥ�׶��ջ���|
'�ۼƸ������Ͻ�|ר��ӿ۳��ϼ�|��ְ������|����˰���������˰|˰�󲹷�|��˾����|�ֻ�����|�������|��ע|�Ƿ����޸�

Sub splitByCompanyName()
    Dim stName As String
    Dim companyWorkBook As Workbook
    Dim person As Integer
	'��˰�걨��ģ���ֶ�
    companyWorkBookTitle = Array("����", "����", "*֤������", "*֤�պ���", "*��������", _
    "������˰����", "�������ϱ��շ�", "����ҽ�Ʊ��շ�", "ʧҵ���շ�", "ס��������", "�ۼ���Ů����", _
    "�ۼƼ�������", "�ۼ�ס��������Ϣ", "�ۼ�ס�����", "�ۼ���������", "��ҵ(ְҵ)���", _
    "��ҵ��������", "˰�����ϱ���", "����", "׼��۳��ľ�����", "����˰��", "��ע")
    stName = "�������ʱ�"
    person = 0
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ChDir (ThisWorkbook.Path)
    Range("A2:AF1519").Select
    ThisWorkbook.Worksheets(stName).Sort.SortFields.Clear
    ThisWorkbook.Worksheets(stName).Sort.SortFields.Add2 Key:=Range("AB2") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ThisWorkbook.Worksheets(stName).Sort
        .SetRange Range("A2:AF1519")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    For low = 2 To 65535
        If ThisWorkbook.Worksheets(stName).Cells(low, 28) = "" Then
            Exit For
        End If
        For fast = low + 1 To 65535
            If ThisWorkbook.Worksheets(stName).Cells(low, 28) <> ThisWorkbook.Worksheets(stName).Cells(fast, 28) Then 'AB�У�����˾������
                Set companyWorkBook = createNewWorkBook(ThisWorkbook.Worksheets(stName).Cells(low, 28))
                For n = 0 To UBound(companyWorkBookTitle)
                    companyWorkBook.Worksheets("��������н��").Cells(1, n + 1) = companyWorkBookTitle(n)
                Next n
                ThisWorkbook.Worksheets(stName).Range("B" & CStr(low) & ":B" & CStr(fast - 1)).Copy _
                Destination:=companyWorkBook.Worksheets("��������н��").Range("B2")
                companyWorkBook.Worksheets("��������н��").Range("C2:C" & CStr(fast - low + 1)) = "�������֤"
                ThisWorkbook.Worksheets(stName).Range("E" & CStr(low) & ":E" & CStr(fast - 1)).Copy _
                Destination:=companyWorkBook.Worksheets("��������н��").Range("D2") '���֤��
                ThisWorkbook.Worksheets(stName).Range("J" & CStr(low) & ":J" & CStr(fast - 1)).Copy _
                Destination:=companyWorkBook.Worksheets("��������н��").Range("E2") '����
                ThisWorkbook.Worksheets(stName).Range("K" & CStr(low) & ":M" & CStr(fast - 1)).Copy _
                Destination:=companyWorkBook.Worksheets("��������н��").Range("G2") '����
                ThisWorkbook.Worksheets(stName).Range("P" & CStr(low) & ":P" & CStr(fast - 1)).Copy _
                Destination:=companyWorkBook.Worksheets("��������н��").Range("J2") '������
                If Application.Version > 11 Then 'execl 2003
                    companyWorkBook.CheckCompatibility = False '�����Լ��
                End If
                companyWorkBook.Save
                companyWorkBook.Close
                person = person + fast - low
                low = fast - 1 '��һ��ѭ�� For ���1������������ǰ��ȥ
                Exit For
            End If
        Next fast
    Next low
    Application.ScreenUpdating = Ture
    r = MsgBox("�����ɹ�" & CStr(person) & "��", vbOKOnly)
End Sub

Function createNewWorkBook(companyName As String) As Object
    ChDir (ThisWorkbook.Path)
    Set NewBook = Workbooks.Add
    With NewBook
        .Title = companyName
        .Subject = companyName
        .SaveAs Filename:=ThisWorkbook.Path & "\" & companyName & ".xls"
    End With
    NewBook.Sheets(1).Name = "��������н��"
    Set createNewWorkBook = NewBook
End Function

