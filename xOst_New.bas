Attribute VB_Name = "xOst_New"

Sub Ost_New()
    
    Dim osTreet As Object
    
    
    Emp1 = Worksheets("EmpMaster").Range("A1").Value
    Emp2 = Worksheets("EmpMaster").Range("A2").Value
    Emp3 = Worksheets("EmpMaster").Range("A3").Value
    Emp4 = Worksheets("EmpMaster").Range("A4").Value
    Emp5 = Worksheets("EmpMaster").Range("A5").Value
    Emp6 = Worksheets("EmpMaster").Range("A6").Value
    
    
    Set osTreet = CreateObject("System.Collections.ArrayList")

    osTreet.Add Emp1
    osTreet.Add Emp2
    osTreet.Add Emp3
    osTreet.Add Emp4
    osTreet.Add Emp5
    osTreet.Add Emp6


    
    Dim i As Long, Counter As Integer, InstanceCount As Integer
    
    Counter = 0
    
   Application.ScreenUpdating = False
    
    
    For i = 0 To osTreet.Count - 1
    
    
    On Error GoTo pt1:
    
        Columns("F:F").Select
        Selection.Find(What:=osTreet(i), after:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(2, -5).Range("A1").Select
        Selection.Copy
        Range("K1").Offset(Counter, 0).Select
        ActiveSheet.Paste
    
        Columns("F:F").Select
        Selection.Find(What:=osTreet(i), after:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(2, -4).Range("A1").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("L1").Offset(Counter, 0).Select
        ActiveSheet.Paste
        Columns("F:F").Select
        Selection.Find(What:=osTreet(i), after:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=Fasle, SearchFormat:=False).Activate
        ActiveCell.Offset(4, -3).Range("A1").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("M1").Offset(Counter, 0).Select
        ActiveSheet.Paste
        Columns("F:F").Select
        Selection.FindNext(after:=ActiveCell).Activate
        ActiveCell.Offset(2, 2).Range("A1").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("N1").Offset(Counter, 0).Select
        ActiveSheet.Paste
    
pt1:
        Range("J1").Offset(Counter, 0).Select
        ActiveCell.FormulaR1C1 = osTreet(i)
        Range("O1").Offset(Counter, 0).Select
        ActiveCell.FormulaR1C1 = "2"
        InstanceCount = Application.WorksheetFunction.CountIf(Range("F:F"), osTreet(i))
        Range("P1").Offset(Counter, 0).Select
        ActiveCell.FormulaR1C1 = InstanceCount
        
    Resume pt2:
    
pt2:
        Counter = Counter + 1
    Next i
    
    Application.ScreenUpdating = True

End Sub


    
    
    
    
   

 
   
