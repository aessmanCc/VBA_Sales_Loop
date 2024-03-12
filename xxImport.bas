Attribute VB_Name = "xxImport"
Sub ImportCref()

Dim directory As String, fileName As String, seet As Worksheet, total As Integer



directory = "C:\Yourdirectory\"
fileName = Dir(directory & "CF1*.xl??")


Workbooks.Open (directory & fileName)

For Each Sheet In Workbooks(fileName).Worksheets
    total = Workbooks("SL1.xlsx").Worksheets.Count
    Workbooks(fileName).Worksheets(Sheet.Name).Copy _
    after:=Workbooks("SL1.xlsx").Worksheets(total)
    

Next Sheet

Workbooks(fileName).Close

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Sub ImportRec()

Dim directory As String, fileName As String, seet As Worksheet, total As Integer


Application.ScreenUpdating = False
Application.DisplayAlerts = False

directory = "C:\Yourdirectory\"
fileName = Dir(directory & "RC1*.xl??")

Workbooks.Open (directory & fileName)

For Each Sheet In Workbooks(fileName).Worksheets
    total = Workbooks("SL1.xlsx").Worksheets.Count
    Workbooks(fileName).Worksheets(Sheet.Name).Copy _
    after:=Workbooks("SL1.xlsx").Worksheets(total)
    

Next Sheet

Workbooks(fileName).Close

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Sub ImportDelPro()

Dim directory As String, fileName As String, seet As Worksheet, total As Integer


Application.ScreenUpdating = False
Application.DisplayAlerts = False

directory = "C:\Yourdirectory\"
fileName = Dir(directory & "DL1*.xl??")

Workbooks.Open (directory & fileName)

For Each Sheet In Workbooks(fileName).Worksheets
    total = Workbooks("SL1.xlsx").Worksheets.Count
    Workbooks(fileName).Worksheets(Sheet.Name).Copy _
    after:=Workbooks("SL1.xlsx").Worksheets(total)
    

Next Sheet

Workbooks(fileName).Close

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Sub ImportCben()

Dim directory As String, fileName As String, seet As Worksheet, total As Integer


Application.ScreenUpdating = False
Application.DisplayAlerts = False

directory = "C:\Yourdirectory\"
fileName = Dir(directory & "CBEN1*.xl??")

Workbooks.Open (directory & fileName)

For Each Sheet In Workbooks(fileName).Worksheets
    total = Workbooks("SL1.xlsx").Worksheets.Count
    Workbooks(fileName).Worksheets(Sheet.Name).Copy _
    after:=Workbooks("SL1.xlsx").Worksheets(total)
    

Next Sheet

Workbooks(fileName).Close

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Sub ImportTemplate()

Dim directory As String, fileName As String, seet As Worksheet, total As Integer


Application.ScreenUpdating = False
Application.DisplayAlerts = False

directory = "C:\Yourdirectory\"
fileName = Dir(directory & "SalesMaster*.xl??")

Workbooks.Open (directory & fileName)

For Each Sheet In Workbooks(fileName).Worksheets
    total = Workbooks("SL1.xlsx").Worksheets.Count
    Workbooks(fileName).Worksheets(Sheet.Name).Copy _
    after:=Workbooks("SL1.xlsx").Worksheets(total)
    

Next Sheet

Workbooks(fileName).Close

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

