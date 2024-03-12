Attribute VB_Name = "xxxxTotalxxxxx"
Option Explicit

Sub totalAverage()
'
Application.ScreenUpdating = False

'Employee List
Call emplMaster

Worksheets("Table 1").Select

'New
Call Ost_New

'Cref
Call ImportCref

Worksheets("Table 1 (2)").Select

Call Ost_New

'Rec
Call ImportRec
Worksheets("Table 1 (3)").Select

Call Ost_New

'Del & Proc
Call ImportDelPro
Worksheets("Table 1 (4)").Select

Call Ost_DelPro

'Cben
Call ImportCben
Worksheets("Table 1 (5)").Select

Call Ost_Cben

'Get Sales Template
Call ImportTemplate

'Data Clean up and Formatting
Call FormatData

Application.ScreenUpdating = True
End Sub


