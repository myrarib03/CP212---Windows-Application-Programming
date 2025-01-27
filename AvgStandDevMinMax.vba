Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Myra Ribeiro
' Student ID: 169030590
' Date: 2025/01/21
' Program title: Task 1
' Description: reports the average, standard deviation
' minimum, and maximum of the scores in a message box
'===========================================================+

'We must declare functions instead of subs because
'we cannot call subs as if they are functions since they
'do not return values the same way.

Function CalculateAverage() As Double
CalculateAverage = Round(WorksheetFunction.Average(Range("A1:A100")), 2)
End Function

Function CalculateStandardDev()
CalculateStandardDev = Round(WorksheetFunction.StDev(Range("A1:A100")), 2)
End Function

Function CalculateMinimum() As Double
CalculateMinimum = Round(WorksheetFunction.Min(Range("A1:A100")), 2)
End Function

Function CalculateMaximum() As Double
CalculateMaximum = Round(WorksheetFunction.Max(Range("A1:A100")), 2)
End Function

'to call a function within this sub is easy.
'however, note that functions are not the same as
'macros and wont show up as such.

Sub TotalCalculations()
    Dim Average1 As Double
    Dim StandardDev As Double
    Dim Minimum As Double
    Dim Maximum As Double

    Average1 = CalculateAverage()
    StandardDev = CalculateStandardDev()
    Minimum = CalculateMinimum()
    Maximum = CalculateMaximum()

    MsgBox "Average: " & Average1 & vbNewLine & _
           "Stdev: " & StandardDev & vbNewLine & _
           "Min: " & Minimum & vbNewLine & _
           "Max: " & Maximum, vbInformation, "Calculation Results"
End Sub
