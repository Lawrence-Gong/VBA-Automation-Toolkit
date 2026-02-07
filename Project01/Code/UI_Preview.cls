VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Sheets("合同信息").Range("F3").Value = UCase(Sheets("合同信息").Range("F3").Value)
Sheets("合同信息").Range("G3").Value = UCase(Sheets("合同信息").Range("G3").Value)
End Sub
