Option Explicit
Sub Macro1()
'
' Macro1 Macro
'
' Keyboard Shortcut: Ctrl+m
'
    Sheets.Add After:=ActiveSheet
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Regional Report"
    Range("A1").Select
    Selection.Style = "Title"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Name"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "District"
    ActiveWindow.LargeScroll Down:=1
    Range("B23").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "Sales Total"
    Range("A10").Select
    ActiveCell.FormulaR1C1 = "Done!"
    Range("A11").Select
    Sheets("Sheet5").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet5").Select
    ActiveSheet.Unprotect
    Range("A1").Select
End Sub
