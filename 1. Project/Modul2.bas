Attribute VB_Name = "Modul2"
Sub ChangeColorConditions()
Dim ws As Worksheet: Set ws = Worksheets("Jahresauswertung")

Dim MyRange As Range
Set MyRange = Range(ws.Cells(9, 4).Address, ws.Cells(100, 1000).Address)

Dim a1 As String
Dim a2 As String

Dim b1 As String
Dim b2 As String

Dim c1 As String
Dim C2 As String

a1 = ws.Cells(1, 5).Value
a2 = ws.Cells(1, 7).Value

b1 = ws.Cells(2, 5).Value
b2 = ws.Cells(2, 7).Value

c1 = ws.Cells(3, 5).Value
C2 = ws.Cells(3, 7).Value

MyRange.FormatConditions.Delete

For Each cell In MyRange.Cells
    If cell <> "" Then
        cell.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=" & a1, Formula2:="=" & a2
        cell.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
        
        cell.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=" & b1, Formula2:="=" & b2
        cell.FormatConditions(2).Interior.Color = RGB(255, 255, 0)
        
        cell.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=" & c1, Formula2:="=" & C2
        cell.FormatConditions(3).Interior.Color = RGB(0, 255, 0)
    End If
Next cell
End Sub
