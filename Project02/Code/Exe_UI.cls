VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Sheets("Exe").Range("B11").Value = UCase(Sheets("Exe").Range("B11").Value)
Sheets("Exe").Range("B13").Value = UCase(Sheets("Exe").Range("B13").Value)
Sheets("Exe").Range("B14").Value = UCase(Sheets("Exe").Range("B14").Value)
Sheets("Exe").Range("A1").Value = Sheets("CN").Range("A2").Value
Sheets("Exe").Range("B1").Value = Sheets("CN").Range("B2").Value
Sheets("Exe").Range("C1").Value = "商机清单-" & Sheets("Exe").Range("A1").Value & "-" & Sheets("Exe").Range("B9").Value & " " & Sheets("Exe").Range("B10").Value & ".xlsx"
Sheets("Exe").Range("B5").Value = "【Renew团队】请查看您的 " & Sheets("Exe").Range("B9").Value & " XXX产品 SERVICE RENEW商机信息摘要-" & Sheets("Exe").Range("A1").Value
Sheets("Exe").Range("B12").Value = Sheets("SN").UsedRange.Rows.Count
If Sheets("Exe").Range("D5").Value = "NonTest" Then
Sheets("Exe").Range("A5").Value = Sheets("Exe").Range("B1").Value
Else
Sheets("Exe").Range("A5").Value = Sheets("Exe").Range("B15").Value
End If
If Sheets("Exe").Range("A1").Value = "" Then
MsgBox ("全部完成")
End If
End Sub



