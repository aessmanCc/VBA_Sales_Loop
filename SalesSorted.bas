Attribute VB_Name = "SalesSorted"
Option Explicit

Sub Fillin()
Attribute Fillin.VB_ProcData.VB_Invoke_Func = " \n14"

    Sheets("SALESMAN SORTED ").Select
    Range("B7").Select
    ActiveCell.FormulaR1C1 = "='Final Avg'!R[-5]C[13]"
    Range("B7").Select
    Selection.AutoFill Destination:=Range("B7:B126"), Type:=xlFillDefault
    Range("B7:B126").Select
    ActiveWindow.SmallScroll Down:=-190
    Range("C7").Select
    ActiveCell.FormulaR1C1 = "='Final Avg'!R[-5]C[-2]"
    Range("C7").Select
    Selection.AutoFill Destination:=Range("C7:C126"), Type:=xlFillDefault
    Range("C7:C126").Select
    ActiveWindow.SmallScroll Down:=-180
    Range("D7").Select
    ActiveCell.FormulaR1C1 = "='Final Avg'!R[-5]C[13]"
    Range("D7").Select
    Selection.AutoFill Destination:=Range("D7:D126"), Type:=xlFillDefault
    Range("D7:D126").Select
    ActiveWindow.SmallScroll Down:=-140
    Range("E7").Select
    ActiveCell.FormulaR1C1 = "='Final Avg'!R[-5]C[7]"
    Range("E7").Select
    Selection.AutoFill Destination:=Range("E7:E126"), Type:=xlFillDefault
    Range("E7:E126").Select
    ActiveWindow.SmallScroll Down:=-140
    Range("F7").Select
    ActiveCell.FormulaR1C1 = "='Final Avg'!R[-5]C[8]"
    Range("F7").Select
    Selection.AutoFill Destination:=Range("F7:F126"), Type:=xlFillDefault
    Range("F7:F126").Select
End Sub
