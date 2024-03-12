Attribute VB_Name = "xxOst_DelPro"
Sub Ost_DelPro()

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

    
    Dim i As Long, Counter As Integer
    
    Counter = 0
    
   Application.ScreenUpdating = False
    
    
    For i = 0 To osTreet.Count - 1
    
    
    On Error GoTo pt1:
    
        Columns("E:E").Select
        Selection.Find(What:=osTreet(i), after:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(0, -2).Range("A1:A18").Select
        Selection.Find(What:="DELIVERY", after:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 1).Range("A1").Select
        Selection.Copy
        Range("L1").Offset(Counter, 0).Select
        ActiveSheet.Paste
        
        Columns("E:E").Select
        Selection.Find(What:=osTreet(i), after:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(0, -2).Range("A1:A18").Select
        Selection.Find(What:="DELIVERY", after:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 0).Range("A1").Select
        Selection.Copy
        Range("K1").Offset(Counter, 0).Select
        ActiveSheet.Paste
    
    
pt1:
        Range("J1").Offset(Counter, 0).Select
        ActiveCell.FormulaR1C1 = osTreet(i)
        Range("O1").Offset(Counter, 0).Select
        ActiveCell.FormulaR1C1 = "2"
        
        
    Resume pt2:
    
pt2:
        Counter = Counter + 1
    Next i
    
    Application.ScreenUpdating = True

End Sub

