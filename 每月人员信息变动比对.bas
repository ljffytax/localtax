Attribute VB_Name = "人员信息比对模块"
'ljffytax 0.0.2 2023-11-05
'GPL v2

'两个文件，一个是ITS里原始的人员清单，另外一个是当月要申报的人员汇总表-人力提供
'对比两个文件里的人员，ITS里不存在而汇总表里存在则将该部分人员追加到ITS尾；反之，
'则将ITS里的人员标注为非正常，并且将人名标红以便于人工添加离职日期


'获取当前目录内的公司名称
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
    Do While st <> "" And 0 = InStr(1, st, "公司", 1)
        st = Dir
        'Debug.Print (InStr(1, st, "公司", 1))
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

'查找对比两个表格，如果身份证件号不在ITS人员表里，但是在工资汇总表里则表示该人员为新入职人员
'如果身份证件号在ITS人员表里并且状态不为"正常",则表示该员工曾经离职但是又重新入职了，此人人员
'会被标记成浅蓝色并将状态修改为正常。
Function employedPerson(smWorkbook As Workbook, itsWorkbook As Workbook) As Integer
    'Set smWorkbook = Workbooks.Open(ThisWorkbook.path & "\汇总.xlsm")
    'Set itsWorkbook = Workbooks.Open(ThisWorkbook.path & "\" & getITSFile())
    Dim PersonID As String
    Dim iPerson As Integer
    Dim companyName As String
    Dim itsPersonID(29999) As String
    
    itsWorkbook.Worksheets(1).Columns("D:D").NumberFormatLocal = "@" '设置身份证列为文本格式
    itsWorkbook.Worksheets(1).Columns("J:J").NumberFormatLocal = "@"
    itsWorkbook.Worksheets(1).Columns("K:L").NumberFormatLocal = "yyyy-mm-dd"
    itsWorkbook.Worksheets(1).Columns("G:G").NumberFormatLocal = "yyyy-mm-dd"
    companyName = getCompanyName()
    iPerson = 0
    
    '获取ITS人员表里全部人员身份证件号
    With itsWorkbook.Worksheets(1)
        For i = 2 To 30000
            If (.Cells(i, 34) = "雇员") Then
                itsPersonID(iPerson) = .Cells(i, 4)
                iPerson = iPerson + 1
            ElseIf (.Cells(i, 34) = "其他") Then
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
            PersonID = smWorkbook.Worksheets(1).Cells(iSM, 5) '身份证件号码
            If (PersonID = "") Then
                Exit For
            End If
            For iITS = 2 To 30000 '最多只比对到3万人，暴力查找效率低下
                'If (0 = StrComp(itsWorkbook.Worksheets(1).Cells(iITS, 4), PersonID, vbTextCompare)) Then
                If (0 = StrComp(itsPersonID(iITS - 2), PersonID, vbTextCompare)) Then
                    If (itsWorkbook.Worksheets(1).Cells(iITS, 7) = "正常") Then
                        Exit For
                    Else
                        itsWorkbook.Worksheets(1).Cells(iITS, 2).Interior.Color = RGB(110, 208, 255) '再入职人员标记姓名浅蓝色
                        itsWorkbook.Worksheets(1).Cells(iITS, 7) = "正常" '修改状态为正常
                        Exit For
                    End If
                Else
                     '找遍表格都没有，则追加到末尾，为防止文中间出现空行引发问题这里取连续两个空行
                    'If (itsWorkbook.Worksheets(1).Cells(iITS, 4) = "" And itsWorkbook.Worksheets(1).Cells(iITS + 1, 4) = "") Then
                    If (itsPersonID(iITS - 2) = "" And itsPersonID(iITS - 1) = "") Then
                        itsPersonID(iITS - 2) = PersonID '新增人员追加到数组尾，以使得数组与ITS表内数据一致
                        With itsWorkbook.Worksheets(1)
                            .Cells(iITS, 2) = smWorkbook.Worksheets(1).Cells(iSM, 2) '姓名
                            .Cells(iITS, 3) = "居民身份证"
                            .Cells(iITS, 4) = PersonID
                            .Cells(iITS, 5) = smWorkbook.Worksheets(1).Cells(iSM, 4) '性别
                            .Cells(iITS, 6) = personID2Birthday(smWorkbook.Worksheets(1).Cells(iSM, 5)) '出生日期
                            .Cells(iITS, 7) = "正常"
                            .Cells(iITS, 10) = smWorkbook.Worksheets(1).Cells(iSM, 29) '电话
                            .Cells(iITS, 11) = smWorkbook.Worksheets(1).Cells(iSM, 8) '入职日期
                            .Cells(iITS, 13) = "是" '是否减除费用
                            .Cells(iITS, 31) = "否"
                            .Cells(iITS, 32) = "否"
                            .Cells(iITS, 33) = "否"
                            .Cells(iITS, 34) = "雇员"
                            .Cells(iITS, 36) = "中国"
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
    
    '获取汇总表里该公司的全部人员身份证件号
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
        If (itsWorkbook.Worksheets(1).Cells(iITS, 7) = "正常" And itsWorkbook.Worksheets(1).Cells(iITS, 34) = "雇员") Then
            For iSM = 0 To iSMPerson
                PersonID = smPersonID(iSM) '身份证件号码
                If (PersonID = "") Then '表明已经找到表尾
                    itsWorkbook.Worksheets(1).Cells(iITS, 2).Interior.Color = RGB(255, 0, 0) '标记姓名红色
                    itsWorkbook.Worksheets(1).Cells(iITS, 7) = "非正常"
                    iPerson = iPerson + 1
                    Exit For
                End If
                If (0 = StrComp(PersonID, itsPersonID, vbTextCompare)) Then
                    Exit For
                End If
            Next iSM
        ElseIf (itsWorkbook.Worksheets(1).Cells(iITS, 4) = "") Then '已到表尾
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
    Set smWorkbook = Workbooks.Open(ThisWorkbook.path & "\汇总.xlsm")
    Set itsWorkbook = Workbooks.Open(ThisWorkbook.path & "\" & getITSFile())
    Application.ScreenUpdating = False
    iFi = firedPerson(smWorkbook, itsWorkbook)
    iEm = employedPerson(smWorkbook, itsWorkbook)
    Application.ScreenUpdating = ture
    smWorkbook.Close savechanges:=False
    r = MsgBox("离职" & CStr(iFi) & "人；新入职" & CStr(iEm) & "人", vbOKOnly)
End Sub

