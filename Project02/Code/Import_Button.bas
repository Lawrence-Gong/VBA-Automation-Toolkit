Attribute VB_Name = "Ä£¿é1"
Sub °´Å¥7_Click()
Dim abc As Variant
'Dim okapp As Object
'Dim ok As Object
Dim efg As Variant
Dim st As Variant
Dim ed As Variant
Dim dls As String
Dim sum As String
Dim n As Integer
Dim m As Integer
While Sheets("CN").Range("A2").Value <> ""
abc = Sheets("Exe").Range("C1").Value
st = Sheets("Exe").Range("B13").Text
ed = Sheets("Exe").Range("B14").Text
dls = Sheets("Exe").Range("B11").Text
sum = Sheets("Exe").Range("B12").Text
n = Sheets("CN").Range("A2").CurrentRegion.Rows.Count
m = n Mod 2
Set rg = Sheets("SN").Range(dls & "1:" & dls & sum)
rg.AutoFilter Field:=1, Criteria1:=Sheets("Exe").Range("A1").Value
    Sheets("SN").Select
    Range(Cells(1, st), Cells(1, ed)).EntireColumn.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("C:C").ColumnWidth = 15
    Columns("B:B").ColumnWidth = 18
    Columns("D:D").ColumnWidth = 15
    Columns("G:G").ColumnWidth = 15
    Columns("E:E").ColumnWidth = 16.75
    Columns("J:J").ColumnWidth = 20.75
    Columns("H:H").ColumnWidth = 17.5
    Columns("I:I").ColumnWidth = 18.5
    Columns("F:F").ColumnWidth = 16
    Columns("A:A").ColumnWidth = 45
    Rows(1).RowHeight = 20
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & abc
    ActiveWindow.Close
    Sheets("Exe").Select
Set okapp = CreateObject("Outlook.Application")
Set ok = okapp.CreateItem(olMailItem)
efg = ThisWorkbook.Path & Sheets("Exe").Range("C1").Value
With ok
.from = Sheets("Exe").Range("B15").Value
.display
.to = Sheets("Exe").Range("A5").Value
.cc = Sheets("CN").Range("C2").Value
.Subject = Sheets("Exe").Range("B5").Value
.HTMLBody = Sheets("Exe").Range("C5").Value & .HTMLBody
.attachments.Add efg
.send
End With
Set objOL = Nothing
Set itmNewMail = Nothing
Kill efg
Sheets("CN").Rows(2).Delete
If m = 1 Then
Sheets("Exe").Range("A1").Select
Else
Sheets("Exe").Range("B1").Select
End If
Wend
End If
Sheets("Exe").Select
End Sub
