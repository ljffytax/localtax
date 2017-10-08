Attribute VB_Name = "模块2"
'ljffytax
'2017-09-18
'2017-10-08

Public ValCodeArr
Public Wi

Sub Sfzhmxy()
    Dim res As Integer
    Dim dlg As String
    Dim sfzh As String
    Dim err As Integer
    Dim sfzhcd As Integer '身份证号长度
    MAX_ROWS = 1000
    Dim dlg1 As String
    'Dim ValCodeArr As Integer
    ValCodeArr = Array("1", "0", "X", "9", "8", "7", "6", "5", "4", "3", "2")
    Wi = Array("7", "9", "10", "5", "8", "4", "2", "1", "6", "3", "7", "9", "10", "5", "8", "4", "2")
    err = 0
    With ThisWorkbook.Worksheets("扣缴个人所得税报告表") '.Range("D11:E" & MAX_ROWS)
        For c = 11 To MAX_ROWS
            sfzh = .Cells(c, 5)
            If sfzh = "" Then
                Exit For
            End If
            If .Cells(c, 4) = "201|居民身份证" Then
            '校验身份证号
                If xysfzh(sfzh) Then
                    err = 0 + err
                Else
                    err = 1 + err
                    dlg1 = dlg1 & c & ";"
                End If
            End If
        Next
    End With
    If err = 0 Then
        res = MsgBox("校验完成，没有错误!", vbOKOnly)
    Else
        dlg = "发现" & err & "处错误!第" & dlg1 & "行"
        res = MsgBox(dlg, vbOKOnly)
    End If
        
End Sub
'检验表格是否符合要求
Sub Check()
    Dim res As Integer
    Dim c As Integer
    Dim dateq As String '日期起
    Dim y As String
    MAX_ROWS = 1000
    
    
    '检验表格结构是否破坏
    If (ThisWorkbook.Worksheets.Count <> 6) Then
        res = MsgBox("请不要更改整个表格的结构，不要添加、删除或更改工作表的名字！", vbCritical, "发现错误！")
        Exit Sub
    End If
    If (ThisWorkbook.Sheets(1).Name <> "扣缴个人所得税报告表") Then
        res = MsgBox("请不要更改整个表格的结构，不要添加、删除或更改工作表的名字！", vbCritical, "发现错误！")
        Exit Sub
    End If
    If (ThisWorkbook.Sheets(2).Name <> "减免事项报告表（减免事项部分）") Then
        res = MsgBox("请不要更改整个表格的结构，不要添加、删除或更改工作表的名字！", vbCritical, "发现错误！")
        Exit Sub
    End If
    If (ThisWorkbook.Sheets(3).Name <> "减免事项报告表（税收协定部分）") Then
        res = MsgBox("请不要更改整个表格的结构，不要添加、删除或更改工作表的名字！", vbCritical, "发现错误！")
        Exit Sub
    End If
    If (ThisWorkbook.Sheets(4).Name <> "商业保险明细表") Then
        res = MsgBox("请不要更改整个表格的结构，不要添加、删除或更改工作表的名字！", vbCritical, "发现错误！")
        Exit Sub
    End If
    If (ThisWorkbook.Sheets(5).Name <> "Sheet4") Then
        res = MsgBox("请不要更改整个表格的结构，不要添加、删除或更改工作表的名字！", vbCritical, "发现错误！")
        Exit Sub
    End If
    If (ThisWorkbook.Sheets(6).Name <> "Sheet5") Then
        res = MsgBox("请不要更改整个表格的结构，不要添加、删除或更改工作表的名字！", vbCritical, "发现错误！")
        Exit Sub
    End If
    If (ThisWorkbook.Sheets(1).Cells(6, 32) <> "备注") Then
        res = MsgBox("请不要对表格进行添加或删除列操作！", vbCritical, "发现错误！")
        Exit Sub
    End If
    
    '检验属期是否正确
    With ThisWorkbook.Worksheets("扣缴个人所得税报告表")
        dateq = .Cells(11, 7)
        For c = 12 To MAX_ROWS
            If .Cells(c, 7) = "" Then
                Exit For
            End If
            If .Cells(c, 7) <> dateq Then
                res = MsgBox("每一个模板文件只能填写一个月，不要将多个月份的数据填在一个模板内！", vbCritical, "发现错误！")
                Exit Sub
            End If
        Next
    End With
    
    '检验是否有未填项及已填项目是否都正确
    If (ThisWorkbook.Sheets(1).Cells(3, 4) = "") Then
        res = MsgBox("请填写纳税人识别号！", vbCritical, "发现错误！")
        Exit Sub
    End If
    If (ThisWorkbook.Sheets(1).Cells(3, 7) = "") Then
        res = MsgBox("请填写纳税人名称！", vbCritical, "发现错误！")
        Exit Sub
    End If
    If (ThisWorkbook.Sheets(1).Cells(3, 31) = "") Then
        res = MsgBox("请填写经办人姓名！", vbCritical, "发现错误！")
        Exit Sub
    End If
    If (ThisWorkbook.Sheets(1).Cells(11, 3) = "") Then
        res = MsgBox("请填写要申报员工的姓名！", vbCritical, "发现错误！")
        Exit Sub
    End If
    If (ThisWorkbook.Sheets(1).Cells(11, 5) = "") Then
        res = MsgBox("请填写要申报员工的证件号码，填写完后记得点击(身份证校验)！", vbCritical, "发现错误！")
        Exit Sub
    End If
    If (ThisWorkbook.Sheets(1).Cells(11, 9) = "") Then
        res = MsgBox("申报金额没有填写，如果是零，请填写0！", vbCritical, "发现错误！")
        Exit Sub
    End If
    If (ThisWorkbook.Sheets(1).Cells(11, 2) = "") Then
        res = MsgBox("请选择是否明细申报，一般应选(是)！", vbCritical, "发现错误！")
        Exit Sub
    End If
    With ThisWorkbook.Worksheets("扣缴个人所得税报告表")
        For c = 11 To MAX_ROWS
            If .Cells(c, 2) = "" Then
                Exit For
            End If
            If .Cells(c, 2) <> "是" Then
                res = MsgBox("是否明细申报，一般选(是)，除非你明白自己选(否)的目的！", vbCritical, "注意！")
                Exit For
            End If
        Next
    End With
    
    With ThisWorkbook.Worksheets("扣缴个人所得税报告表")
        If isGoodType(.Cells(11, 4), "sfzj") Then
            tmp = .Cells(11, 4)
        Else
            res = MsgBox("(身份证件类型)一列，应当从下拉列表中选择，禁止自己填写！", vbCritical, "错误！")
            Exit Sub
        End If
        For c = 12 To MAX_ROWS
            If .Cells(c, 4) = "" Then
                Exit For
            End If
            If .Cells(c, 4) <> tmp Then
                If isGoodType(.Cells(c, 4), "sfzj") Then
                    tmp = .Cells(c, 4)
                Else
                    res = MsgBox("(身份证件类型)一列，应当从下拉列表中选择，禁止自己填写！", vbCritical, "错误！")
                    Exit Sub
                End If
            End If
        Next
    End With
    
    With ThisWorkbook.Worksheets("扣缴个人所得税报告表")
        If isGoodType(.Cells(11, 6), "sdxm") Then
            tmp = .Cells(11, 6)
        Else
            If .Cells(11, 6) = "----工资薪金所得----" Then
                res = MsgBox("(所得项目)一列，如果是申报工资应该选(正常工资薪金)！", vbCritical, "错误！")
            ElseIf .Cells(11, 6) = "----财产转让所得----" Then
                res = MsgBox("(所得项目)一列，----财产转让所得----不是正确的品目！", vbCritical, "错误！")
            ElseIf .Cells(11, 6) = "----财产租赁所得----" Then
                res = MsgBox("(所得项目)一列，----财产租赁所得----不是正确的品目！", vbCritical, "错误！")
            Else
                res = MsgBox("(所得项目)一列，应当从下拉列表中选择，禁止自己填写！", vbCritical, "错误！")
            End If
            Exit Sub
        End If
        For c = 12 To MAX_ROWS
            If .Cells(c, 6) = "" Then
                Exit For
            End If
            If .Cells(c, 6) <> tmp Then
                If isGoodType(.Cells(c, 6), "sdxm") Then
                    tmp = .Cells(c, 6)
                Else
                    If .Cells(c, 6) = "----工资薪金所得----" Then
                        res = MsgBox("(所得项目)一列，工资应该选(正常工资薪金)！", vbCritical, "错误！")
                    ElseIf .Cells(c, 6) = "----财产转让所得----" Then
                        res = MsgBox("(所得项目)一列，----财产转让所得----不是正确的品目！", vbCritical, "错误！")
                    ElseIf .Cells(c, 6) = "----财产租赁所得----" Then
                        res = MsgBox("(所得项目)一列，----财产租赁所得----不是正确的品目！", vbCritical, "错误！")
                    Else
                        res = MsgBox("(所得项目)一列，应当从下拉列表中选择，禁止自己填写！", vbCritical, "错误！")
                    End If
                    Exit Sub
                End If
            End If
        Next
    End With
    
    res = MsgBox("校验完成，没发现错误,Good Luck！", vbOKOnly, "恭喜！")
End Sub


Function isGoodType(s As String, t As String) As Boolean
    'Dim sfzjlx As String '身份证件类型
    'Dim sdxmlx As String '所得项目类型

    sfzjlx = Array( _
    "201|居民身份证", _
    "210|港澳居民来往内地通行证", _
    "208|外国护照", _
    "213|台湾居民来往大陆通行证", _
    "219|香港永久性居民身份证", _
    "227|中国护照", _
    "202|军官证", _
    "203|武警警官证", _
    "204|士兵证", _
    "216|外交官证", _
    "220|台湾身份证", _
    "221|澳门特别行政区永久性居民身份证", _
    "233|外国人永久居留身份证（外国人永久居留证）" _
    )

    sdxmlx = Array( _
    "    正常工资薪金", _
    "    外籍人员正常工资薪金", _
    "    全年一次性奖金收入", _
    "劳务报酬所得", _
    "利息、股息、红利所得", _
    "    外籍人员数月奖金", _
    "    内退一次性补偿金", _
    "    解除劳动合同一次性补偿金", _
    "    个人股票期权行权收入", _
    "    提前退休一次性补贴", _
    "    企业年金", _
    "    其他财产转让所得", _
    "    股权转让所得", _
    "    个人房屋转让所得", _
    "    股票转让所得", _
    "    财产拍卖所得及回流文物拍卖所得", _
    "    个人房屋出租所得", _
    "    其他财产租赁所得", _
    "稿酬所得", _
    "特许权使用费所得", _
    "偶然所得", _
    "其他所得" _
    )
    
    If t = "sfzj" Then
        For i = 0 To 12
            If s = sfzjlx(i) Then
                isGoodType = True
                Exit Function
            End If
        Next
    ElseIf t = "sdxm" Then
        For i = 0 To 21
            If s = sdxmlx(i) Then
                isGoodType = True
                Exit Function
            End If
        Next
    End If
    isGoodType = False
End Function


Function getArray(s As String)
    Dim tmpArray(17) As Integer
    Dim ts As String
    
    ts = Left(s, 17)
    If IsNumeric(ts) Then
        For i = 1 To 17
            tmpArray(i) = Mid(s, i, 1)
        Next
    Else
        tmpArray(1) = 0
    End If
    getArray = tmpArray
End Function

Function xysfzh(sfz As String) As Boolean
    Dim cd As Integer '身份证号长度
    Dim xyw As String '校验位
    Dim sfzhArray() As Integer
    Dim TotalmulAiWi As Integer
    
    cd = Len(sfz)
    If cd = 15 And IsNumeric(sfz) Then
        xysfzh = True '身份证号对
    Else
        If cd = 18 Then
           sfzhArray = getArray(sfz)
           If sfzhArray(1) = 0 Then
               xysfzh = False '身份证号错
           Else
               TotalmulAiWi = 0
               For i = 1 To 17
                   TotalmulAiWi = sfzhArray(i) * Wi(i - 1) + TotalmulAiWi
               Next
               xyw = ValCodeArr(TotalmulAiWi Mod 11)
               If xyw = Right(sfz, 1) Then
                   xysfzh = True '身份证号对
               Else
                   xysfzh = False '身份证号错
               End If
           End If
        Else
           xysfzh = False '身份证号错
        End If
    End If
End Function
