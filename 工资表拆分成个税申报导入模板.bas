Attribute VB_Name = "模块1"
'Author ljffytax 2023-06-30 V0.02
'GNU GPLv2
'工资表表头字段：
'序号|姓名|花名|性别|身份证号码|部门|岗位|入职日期|离职日期|税前收入|养老保险|医疗保险|失业保险|残保金|
'其他社保费用|公积金|累计子女教育|累计继续教育|累计住房贷款利息|累计住房租金|累计赡养老人|累计婴幼儿照护费|
'累计个人养老金|专项附加扣除合计|离职补偿金|本月税额个人所得税|税后补发|公司名称|手机号码|免除费用|备注|是否有修改

Sub splitByCompanyName()
    Dim stName As String
    Dim companyWorkBook As Workbook
    Dim person As Integer
    '个税申报表模板字段
    companyWorkBookTitle = Array("工号", "*姓名", "*证件类型", "*证件号码", "本期收入", _
    "本期免税收入", "基本养老保险费", "基本医疗保险费", "失业保险费", "住房公积金", "累计子女教育", _
    "累计继续教育", "累计住房贷款利息", "累计住房租金", "累计赡养老人", "累计3岁以下婴幼儿照护", _
    "累计个人养老金", "企业(职业)年金", "商业健康保险", "税延养老保险", "其他", "准予扣除的捐赠额", _
    "减免税额", "备注")
    stName = "人力工资表"
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
            If ThisWorkbook.Worksheets(stName).Cells(low, 28) <> ThisWorkbook.Worksheets(stName).Cells(fast, 28) Then 'AB列，即公司名字列
                Set companyWorkBook = createNewWorkBook(ThisWorkbook.Worksheets(stName).Cells(low, 28))
                With companyWorkBook.Worksheets("正常工资薪金收入")
                    For n = 0 To UBound(companyWorkBookTitle)
                        .Cells(1, n + 1) = companyWorkBookTitle(n)
                        .Cells(1, n + 1).Interior.Color = 16772300
                    Next n
                End With
                ThisWorkbook.Worksheets(stName).Range("B" & CStr(low) & ":B" & CStr(fast - 1)).Copy _
                Destination:=companyWorkBook.Worksheets("正常工资薪金收入").Range("B2")
                companyWorkBook.Worksheets("正常工资薪金收入").Range("C2:C" & CStr(fast - low + 1)) = "居民身份证"
                ThisWorkbook.Worksheets(stName).Range("E" & CStr(low) & ":E" & CStr(fast - 1)).Copy _
                Destination:=companyWorkBook.Worksheets("正常工资薪金收入").Range("D2") '身份证号
                ThisWorkbook.Worksheets(stName).Range("J" & CStr(low) & ":J" & CStr(fast - 1)).Copy _
                Destination:=companyWorkBook.Worksheets("正常工资薪金收入").Range("E2") '收入
                ThisWorkbook.Worksheets(stName).Range("K" & CStr(low) & ":M" & CStr(fast - 1)).Copy _
                Destination:=companyWorkBook.Worksheets("正常工资薪金收入").Range("G2") '三险
                ThisWorkbook.Worksheets(stName).Range("P" & CStr(low) & ":P" & CStr(fast - 1)).Copy _
                Destination:=companyWorkBook.Worksheets("正常工资薪金收入").Range("J2") '公积金
                If Application.Version > 11 Then 'execl 2003
                    companyWorkBook.CheckCompatibility = False '兼容性检查
                End If
                companyWorkBook.Save
                companyWorkBook.Close
                person = person + fast - low
                low = fast - 1 '下一轮循环 For 会加1，所以这里提前减去
                Exit For
            End If
        Next fast
    Next low
    Application.ScreenUpdating = Ture
    r = MsgBox("拆分完成共" & CStr(person) & "人", vbOKOnly)
End Sub

Function createNewWorkBook(companyName As String) As Object
    ChDir (ThisWorkbook.Path)
    Set NewBook = Workbooks.Add
    With NewBook
        .Title = companyName
        .Subject = companyName
        .SaveAs Filename:=ThisWorkbook.Path & "\" & companyName & ".xlsx"
    End With
    NewBook.Sheets(1).Name = "正常工资薪金收入"
    Set createNewWorkBook = NewBook
End Function

