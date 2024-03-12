Attribute VB_Name = "Formatting"

Sub FormatData()

'Sheet Add & Data Clean up
Sheets.Add.Name = "Final Avg"
Worksheets("Final Avg").Select
Range("B2:I114").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
Worksheets("Table 1").Select
Range("J1:L114").Select
Selection.Copy
Worksheets("Final Avg").Select
Range("A2").Select
ActiveSheet.Paste
'New Pmts
Worksheets("Table 1").Select
Range("N1:N114").Select
Selection.Copy
Worksheets("Final Avg").Select
Range("J2").Select
ActiveSheet.Paste
'Instances & Store
Worksheets("Table 1").Select
Range("O1:P114").Select
Selection.Copy
Worksheets("Final Avg").Select
Range("O2").Select
ActiveSheet.Paste
'Cref
Worksheets("Table 1 (2)").Select
Range("K1:M114").Select
Selection.Copy
Worksheets("Final Avg").Select
Range("E2").Select
ActiveSheet.Paste
'Rec
Worksheets("Table 1 (3)").Select
Range("M1:M114").Select
Selection.Copy
Worksheets("Final Avg").Select
Range("D2").Select
ActiveSheet.Paste
'Del & Proc
Worksheets("Table 1 (4)").Select
Range("K1:L114").Select
Selection.Copy
Worksheets("Final Avg").Select
Range("H2").Select
ActiveSheet.Paste
'Cben
Worksheets("Table 1 (5)").Select
Range("K1:K114").Select
Selection.Copy
Worksheets("Final Avg").Select
Range("K2").Select
ActiveSheet.Paste
'SalesPerson Names
Worksheets("EmpMaster").Select
Range("B1:B114").Select
Selection.Copy
Worksheets("Final Avg").Select
Range("Q2").Select
ActiveSheet.Paste
Worksheets("EmpMaster").Select
Range("C1:C114").Select
Selection.Copy
Worksheets("Final Avg").Select
Range("R2").Select
ActiveSheet.Paste


'Formatting
Worksheets("Final Avg").Select
Range("A1").Value = "Sales#"
Range("B1").Value = "New"
Range("C1").Value = "Grp"
Range("D1").Value = "Rec"
Range("E1").Value = "Cref-New"
Range("F1").Value = "Cref-Grp"
Range("G1").Value = "Cref-Rec"
Range("H1").Value = "Delivery"
Range("I1").Value = "Processing"
Range("J1").Value = "Pmts"
Range("K1").Value = "Cben"
Range("L1").Value = "#New"
Range("M1").Value = "Total"
Range("N1").Value = "Avg"
Range("O1").Value = "Store"
Range("P1").Value = "Instance"
Range("Q1").Value = "Salesperson"
Range("R1").Value = "Month End"

'New Formula
Range("L2:L114").NumberFormat = "General"
Range("L2").Formula = "=J2 - K2"
Range("L2:L114").FillDown
Range("M2:M114").NumberFormat = "General"
Range("M2").Formula = "=B2 + C2 + D2 + E2 + F2 + G2 + H2"
Range("M2:M114").FillDown
Range("N2:N114").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
Range("N2").Formula = "=M2 / L2"
Range("N2:N114").FillDown

'Convert text to number
Range("S2").Formula = "=B2*1"
Range("S2:Z114").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
Range("AA2:AB114").NumberFormat = "General"
Range("S2:AB2").FillRight
Range("S2:AB114").FillDown
Range("S2:AB114").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
Range("J2:L114").NumberFormat = "General"
Range("B2:K114").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
Range("M2:M114").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
Range("S2:AB114").Select
Selection.Copy
Range("AC2").PasteSpecial xlPasteValues
Selection.Copy
Range("B2").PasteSpecial xlPasteValues
Range("S2:AL114").Select
Selection.Delete
Range("O2:P114").NumberFormat = "General"


'Delete Worksheets
Application.DisplayAlerts = False

Sheets("EmpMaster").Delete
Sheets("Table 1").Delete
Sheets("Table 1 (2)").Delete
Sheets("Table 1 (3)").Delete
Sheets("Table 1 (4)").Delete
Sheets("Table 1 (5)").Delete

'Populate Sales Sorted Spreadsheet
Call Fillin

'Save Report
ActiveWorkbook.SaveAs fileName:=("C:\Your Report Location" & VBA.Strings.Format(Date, "MM-DD-YYYY") & ".xlsx")

Application.DisplayAlerts = True
End Sub
