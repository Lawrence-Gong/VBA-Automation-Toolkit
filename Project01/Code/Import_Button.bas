Attribute VB_Name = "模块5"
Sub 按钮7_Click()
If Sheets("合同信息").Range("F3").Value = "" Then
MsgBox ("请输入厂商名称")
ElseIf Sheets("合同信息").Range("G3").Value = "" Then
MsgBox ("请输入客编")
ElseIf Sheets("合同信息").Range("H3").Value = "" Then
MsgBox ("请输入BR")
ElseIf Sheets("合同信息").Range("I3").Value = "" Then
MsgBox ("请输入SO号")
Else
With Sheets("数据库")
    .Rows(3).Insert
    .Rows(5).Copy .Rows(3)
Sheets("数据库").Range("D3").Value = Sheets("合同信息").Range("B3").Value
Sheets("数据库").Range("E3").Value = Sheets("合同信息").Range("B5").Value
Sheets("数据库").Range("H3").Value = Sheets("合同信息").Range("B6").Value
Sheets("数据库").Range("I3").Value = Sheets("合同信息").Range("B7").Value
Sheets("数据库").Range("J3").Value = Sheets("合同信息").Range("B8").Value
Sheets("数据库").Range("K3").Value = Sheets("合同信息").Range("B12").Value
Sheets("数据库").Range("l3").Value = Sheets("合同信息").Range("B10").Value
Sheets("数据库").Range("M3").Value = Sheets("合同信息").Range("B9").Value
Sheets("数据库").Range("N3").Value = Sheets("合同信息").Range("B13").Value
Sheets("数据库").Range("O3").Value = Sheets("合同信息").Range("B14").Value
Sheets("数据库").Range("P3").Value = Sheets("合同信息").Range("B15").Value
Sheets("数据库").Range("R3").Value = Sheets("合同信息").Range("B17").Value
Sheets("数据库").Range("s3").Value = Sheets("合同信息").Range("B16").Value
Sheets("数据库").Range("t3").Value = Sheets("合同信息").Range("D3").Value
Sheets("数据库").Range("u3").Value = Sheets("合同信息").Range("d5").Value
Sheets("数据库").Range("x3").Value = Sheets("合同信息").Range("d6").Value
Sheets("数据库").Range("y3").Value = Sheets("合同信息").Range("d10").Value
Sheets("数据库").Range("z3").Value = Sheets("合同信息").Range("d7").Value
Sheets("数据库").Range("aa3").Value = Sheets("合同信息").Range("d13").Value
Sheets("数据库").Range("ab3").Value = Sheets("合同信息").Range("d14").Value
Sheets("数据库").Range("ac3").Value = Sheets("合同信息").Range("d15").Value
Sheets("数据库").Range("ae3").Value = Sheets("合同信息").Range("d17").Value
Sheets("数据库").Range("af3").Value = Sheets("合同信息").Range("d16").Value
Sheets("数据库").Range("AH3").Value = Sheets("合同信息").Range("I3").Value
Sheets("数据库").Range("A2").Value = Sheets("合同信息").Range("F3").Value
Sheets("数据库").Range("A3").Value = Sheets("数据库").Range("A2").Value
Sheets("数据库").Range("B3").Value = Sheets("合同信息").Range("G3").Value
Sheets("数据库").Range("C3").Value = Sheets("合同信息").Range("H3").Value
Sheets("数据库").Range("AI3") = Format(Now, "yyyy-mm-dd hh:mm:ss")
Rows(3).NumberFormat = Text
Sheets("数据库").Range("A3:AI3").Borders.LineStyle = True
End With
End If
End Sub

