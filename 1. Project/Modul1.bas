Attribute VB_Name = "Modul1"

Sub CalculateMaschineStatistics()
Dim monthWS As Worksheet: Set monthWS = Worksheets("Monatsübersicht")

Dim maschineCount As Integer
Dim arr() As Variant     '4 2 0 Maschine values in a 2D Array

Dim xStart As Integer
Dim yStart As Integer
Dim xA As Integer
Dim yA As Integer



If IsNumeric(monthWS.Cells(2, 3).Value) And IsNumeric(monthWS.Cells(3, 3).Value) Then


    xStart = monthWS.Cells(2, 3).Value
    yStart = monthWS.Cells(3, 3).Value
    
    xA = xStart - 1
    yA = yStart - 1
    
    If monthWS.Cells(yA, xA).Value = "Bereich" Then
    
        Call GetMaschineCount
        maschineCount = monthWS.Cells(1, 3).Value
    
        arr = Range(monthWS.Cells(yStart, xStart), monthWS.Cells(yA + maschineCount, xA + 31))
        
        Call CalculateMonthlyAvarage(monthWS, arr, yA, xA)
        Call CalculatePercent(monthWS, arr, yA, xA, maschineCount)
        Call CalculatePercentAvg(monthWS, arr, yA, xA, maschineCount)
        Call FillValuesIntoYearly(monthWS, yStart, xStart, maschineCount)
    Else
    MsgBox ("COORDINATES DO NOT LEAD TO TABLE / ARE INCORRECT")
    End If
Else
    MsgBox ("INVALID COORDINATES")
End If


End Sub
Sub CalculateMonthlyAvarage(monthWS As Worksheet, arr() As Variant, yA As Integer, xA As Integer)
Dim sum As Double
Dim cur As String
Dim usedDaysCount As Integer

For y = LBound(arr, 1) To UBound(arr, 1)
    For x = LBound(arr, 2) To UBound(arr, 2)
    
        cur = monthWS.Cells(y + yA, x + xA).Value
        If cur <> "" Then
            If cur = "2" Then sum = sum + 50
            If cur = "4" Then sum = sum + 100
            usedDaysCount = usedDaysCount + 1
        End If
        
    Next
    
    monthWS.Cells(y + yA, x + xA).Value = CalcAvg(sum, usedDaysCount) / 100
    sum = 0
    usedDaysCount = 0
    
Next
End Sub

Sub CalculatePercent(monthWS As Worksheet, arr() As Variant, yA As Integer, xA As Integer, maschineCount As Integer)
Dim cur As String

Dim wc As Double 'WorkingMaschineCount
Dim hwwc As Double 'HalfWayWorkingMaschineCount
Dim nwc As Double 'NotWorkingMaschineCount

y = 1   'initializing y so it isn't empty
For x = LBound(arr, 2) To UBound(arr, 2)

    cur = monthWS.Cells(y + yA, x + xA).Value

    If cur <> "" Then
        For y = LBound(arr, 1) To UBound(arr, 1)
            
            cur = monthWS.Cells(y + yA, x + xA).Value
            
            If cur = "4" Then wc = wc + 1
            If cur = "2" Then hwwc = hwwc + 1
            If cur = "0" Then nwc = nwc + 1
            
        Next
        
        monthWS.Cells(y + yA, x + xA).Value = CalcPercent(wc, maschineCount) / 100
        monthWS.Cells(y + yA + 1, x + xA).Value = CalcPercent(hwwc, maschineCount) / 100
        monthWS.Cells(y + yA + 2, x + xA).Value = CalcPercent(nwc, maschineCount) / 100
        y = 1 'resetting y before checking next current, else we'll start at 26
    End If
    
    nwc = 0
    wc = 0
    hwwc = 0
Next
End Sub
Sub CalculatePercentAvg(monthWS As Worksheet, arr() As Variant, yA As Integer, xA As Integer, maschineCount As Integer)

Dim sum As Double
Dim cur As String
Dim usedDaysCount As Integer

For y = LBound(arr, 1) To 3
    For x = LBound(arr, 2) To UBound(arr, 2)
        cur = monthWS.Cells(y + yA + maschineCount, x + xA).Value
        If cur <> "" Then
            sum = sum + cur
            usedDaysCount = usedDaysCount + 1
        End If
    Next
    
    If usedDaysCount <> 0 Then
    monthWS.Cells(y + yA + maschineCount, x + xA).Value = CalcPercent(sum, usedDaysCount) / 100
    Else
    monthWS.Cells(y + yA + maschineCount, x + xA).Value = 0
    End If
    
    sum = 0
    usedDaysCount = 0
    
Next
End Sub









Function CalcAvg(sum As Double, count As Integer)
CalcAvg = sum / count
End Function

Function CalcPercent(sum As Double, count As Integer)
CalcPercent = (sum * 100) / count
End Function
Sub GetMaschineCount()
Dim monthWS As Worksheet: Set monthWS = Worksheets("Monatsübersicht")

Dim cur As String
Dim count As Integer
Dim i As Integer

Dim xStart As Integer
Dim yStart As Integer
Dim xA                                                  'x adjusted with -3, since we need to go down the first column


    xStart = monthWS.Cells(2, 3).Value
    yStart = monthWS.Cells(3, 3).Value
    
    
    xA = xStart - 3
            
            
    cur = monthWS.Cells(yStart, xA).Value
    While cur <> "Verfügbar"
        count = count + 1
        i = i + 1
        cur = monthWS.Cells(yStart + i, xA).Value
    Wend
                
    monthWS.Cells(1, 3).Value = count                      'Set MaschineCount to it's static position
End Sub






Function GetMonth(monthWS As Worksheet, yStart As Integer, xStart As Integer) As Integer
Dim cur As String
Dim m As Integer 'month number

cur = monthWS.Cells(yStart - 2, xStart + 15)

Select Case cur

Case "Januar"
m = 1
Case "Februar"
m = 2
Case "März"
m = 3
Case "April"
m = 4
Case "Mai"
m = 5
Case "Juni"
m = 6
Case "Juli"
m = 7
Case "August"
m = 8
Case "September"
m = 9
Case "Oktober"
m = 10
Case "November"
m = 11
Case "Dezember"
m = 12
Case Else
m = -1
End Select

GetMonth = m
End Function

Function GetYearDifference(monthWS As Worksheet, yStart As Integer, xStart As Integer) As Integer
Dim cur As String

cur = monthWS.Cells(yStart - 2, xStart + 26).Value
If IsNumeric(cur) Then
    If (cur > 2021 And cur < 2100) Then
        GetYearDifference = cur - 2022
    Else
        GetYearDifference = -1
    End If
Else
    GetYearDifference = -1
End If

End Function

Function GetFinalXPosInYearlySheet(monthWS As Worksheet, yStart As Integer, xStart As Integer)
Dim monthNr As Integer
Dim yearlyDiff As Integer
Dim isValid As Boolean

monthNr = GetMonth(monthWS, yStart, xStart)
yearlyDiff = GetYearDifference(monthWS, yStart, xStart)
isValid = True

If monthNr = -1 Then
    MsgBox ("INVALID_MONTH")
    isValid = False
End If

If yearlyDiff = -1 Then
    MsgBox ("INVALID_YEAR")
    isValid = False
End If

If isValid = True Then
GetFinalXPosInYearlySheet = monthNr + (12 * yearlyDiff) + 5
Else: GetFinalXPosInYearlySheet = -1
End If
End Function





Sub FillValuesIntoYearly(monthWS As Worksheet, yStart As Integer, xStart As Integer, maschineCount As Integer)
Dim yearlyWS As Worksheet: Set yearlyWS = Worksheets("Jahresauswertung")
Dim xVal As Integer


xVal = GetFinalXPosInYearlySheet(monthWS, yStart, xStart)

If xVal <> -1 Then
    If yearlyWS.Cells(9, xVal).Value <> "" Then
        Dim answer As Integer
    
        answer = MsgBox("In der gegebenen Reihe ist schon information vorhanden. Soll diese überschrieben werden?", vbQuestion + vbYesNo + vbDefaultButton2, "ACHTUNG!")
    
        If answer = vbYes Then
            For i = 1 To (maschineCount + 3)
                yearlyWS.Cells(8 + i, xVal).Value = monthWS.Cells((yStart - 1) + i, xStart + 31)
            Next
        End If
    Else
        For i = 1 To (maschineCount + 3)
            yearlyWS.Cells(8 + i, xVal).Value = monthWS.Cells((yStart - 1) + i, xStart + 31)
        Next
    End If
End If


End Sub


