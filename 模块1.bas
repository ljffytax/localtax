Attribute VB_Name = "模块1"
'陈梓源,lianjie,神州数码
'日期：2016-08-30,2017-07-13,2018-01-12,2018-01-22
'版权所有，请勿传播，勿修改

'法定费用扣除额赋值
Function FdfykceFz(sdxm As String, sdqjq As Date, sdqjz As Date, sre As Double, mssd As Double, hj As Double) As Double
    sdxm = VBA.LTrim(sdxm)
    If sdxm = "正常工资薪金" Then
        If sdqjq >= #1/9/1980# And sdqjz <= #2/29/2008# And sdqjq <= sdqjz Then
            FdfykceFz = 1600
        ElseIf sdqjq >= #1/3/2008# And sdqjz <= #8/31/2011# And sdqjq <= sdqjz Then
            FdfykceFz = 2000
        ElseIf sdqjq >= #1/9/2011# And sdqjz <= #12/31/2050# And sdqjq <= sdqjz Then
            FdfykceFz = 3500
        End If
        ElseIf sdxm = "外籍人员正常工资薪金" Then
            If sdqjq >= #1/9/1980# And sdqjz <= #12/31/2005# And sdqjq <= sdqjz Then
                FdfykceFz = 4000
            ElseIf sdqjq >= #1/1/2006# And sdqjz <= #12/31/2050# And sdqjq <= sdqjz Then
                FdfykceFz = 4800
            Else
                FdfykceFz = 0
            End If
        ElseIf sdxm = "劳务报酬所得" Or sdxm = "稿酬所得" Or sdxm = "特许权使用费所得" Or sdxm = "其他财产租赁所得" Or sdxm = "个人房屋出租所得" Then
            If (sre - mssd - hj) <= 4000 Then
                FdfykceFz = 800
            Else
                FdfykceFz = (sre - mssd - hj) * 0.2
            End If
    Else
        FdfykceFz = 0
    End If

End Function

'税率赋值
Function SlFz(sdxm As String, sdqjq As Date, sdqjz As Date, ynssde As Double) As String
    sdxm = VBA.LTrim(sdxm)
    Dim ysJe As Double
    If sdxm = "正常工资薪金" Or sdxm = "外籍人员正常工资薪金" Or sdxm = "全年一次性奖金收入" Or sdxm = "外籍人员数月奖金" Then
        If sdxm = "全年一次性奖金收入" Or sdxm = "外籍人员数月奖金" Then
            ysJe = ynssde / 12
        Else
            ysJe = ynssde
        End If
        
        If sdqjq >= #1/1/1980# And sdqjz <= #8/31/2011# Then
            If ysJe >= 0 And ysJe <= 500 Then
                SlFz = "0.05(0.00)"
            ElseIf ysJe > 500 And ysJe <= 2000 Then
                SlFz = "0.10(25.00)"
            ElseIf ysJe > 2000 And ysJe <= 5000 Then
                SlFz = "0.15(125.00)"
            ElseIf ysJe > 5000 And ysJe <= 20000 Then
                SlFz = "0.20(375.00)"
            ElseIf ysJe > 20000 And ysJe <= 40000 Then
                SlFz = "0.25(1,375.00)"
            ElseIf ysJe > 40000 And ysJe <= 60000 Then
                SlFz = "0.30(3,375.00)"
            ElseIf ysJe > 60000 And ysJe <= 80000 Then
                SlFz = "0.35(6,375.00)"
            ElseIf ysJe > 80000 And ysJe <= 100000 Then
                SlFz = "0.40(10,375.00)"
            ElseIf ysJe > 100000 Then
                SlFz = "0.45(15,375.00)"
            End If
        ElseIf sdqjq >= #9/1/2011# And sdqjz <= #12/31/2055# Then
            If ysJe >= 0 And ysJe <= 1500 Then
                SlFz = "0.03(0.00)"
            ElseIf ysJe > 1500 And ysJe <= 4500 Then
                SlFz = "0.10(105.00)"
            ElseIf ysJe > 4500 And ysJe <= 9000 Then
                SlFz = "0.20(555.00)"
            ElseIf ysJe > 9000 And ysJe <= 35000 Then
                SlFz = "0.25(1,005.00)"
            ElseIf ysJe > 35000 And ysJe <= 55000 Then
                SlFz = "0.30(2,755.00)"
            ElseIf ysJe > 55000 And ysJe <= 80000 Then
                SlFz = "0.35(5,505.00)"
            ElseIf ysJe > 80000 Then
                SlFz = "0.45(13,505.00)"
            End If
        End If
    ElseIf sdxm = "劳务报酬所得" Then
        If ynssde >= 0 And ynssde <= 20000 Then
            SlFz = "0.20(0.00)"
        ElseIf ynssde > 20000 And ynssde <= 50000 Then
            SlFz = "0.30(2,000.00)"
        ElseIf ynssde > 50000 Then
            SlFz = "0.40(7,000.00)"
        End If
    ElseIf sdxm = "稿酬所得" Then
        SlFz = "0.20(稿酬所得)"
    ElseIf sdxm = "特许权使用费所得" Then
        SlFz = "0.20(特许权使用费所得)"
    ElseIf sdxm = "利息、股息、红利所得" And ynssde >= 0 Then
        SlFz = "0.20(其他利息、股息、红利所得)"
    ElseIf sdxm = "其他财产转让所得" And ynssde >= 0 Then
        SlFz = "0.20(其他转让所得)"
    ElseIf sdxm = "个人房屋出租所得" And ynssde >= 0 Then
        SlFz = "0.10(个人房屋出租所得)"
    ElseIf sdxm = "偶然所得" And ynssde >= 0 Then
        SlFz = "0.20(偶然所得)"
    ElseIf sdxm = "其他所得" And ynssde >= 0 Then
        SlFz = "0.20(其他所得)"
    ElseIf sdxm = "个人房屋转让所得" And ynssde >= 0 Then
        SlFz = "0.20(房屋转让所得)"
    ElseIf sdxm = "其他财产租赁所得" Then
        SlFz = "0.20(其他财产租赁所得)"
    ElseIf sdxm = "股权转让所得" Or sdxm = "股票转让所得" Or sdxm = "财产拍卖所得及回流文物拍卖所得" Then
        SlFz = ""
    Else
        SlFz = ""
    End If
End Function


'所得期间至赋值
Function SdqjzFz(sdqjq As Date) As Date
    SdqjzFz = CDate(Year(sdqjq) & "-" & Month(sdqjq) & "-" & Day(DateSerial(Year(sdqjq), Month(sdqjq) + 1, 0)))
End Function



