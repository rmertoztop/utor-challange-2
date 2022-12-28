Attribute VB_Name = "Module1"
Sub Module2Challenge()

    Application.ScreenUpdating = False

    Dim Ticker As String
    Dim YearlyChange As String
    Dim PercentChange As String
    Dim TotalStockVolume As String
    Dim value As String

    Range("I1").value = "Ticker"
    Range("J1").value = "YearlyChange"
    Range("K1").value = "PercentChange"
    Range("L1").value = "TotalStockVolume"

    Range("I1") = "Ticker"
    Range("J1") = "YearlyChange"
    Range("K1") = "PercentChange"
    Range("L1") = "TotalStockVolume"


    
    Range("A2:B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    Selection.Copy
    Range("I2").Select
    ActiveSheet.Paste
    
    'last row for a = lra
    
    Dim lra As Double

    lra = Range("A2").End(xlDown).Row


    ActiveSheet.Range(Cells(2, 9), Cells(lra, 9)).RemoveDuplicates Columns:=1


    ActiveSheet.Range(Cells(2, 10), Cells(lra, 10)).RemoveDuplicates Columns:=1


    'last row for j = lrj
    
    Dim lrj As Double

    lrj = Range("J2").End(xlDown).Row
    
    
    'last row for i = lri
    
    Dim lri As Double

    lri = Range("I2").End(xlDown).Row
    

    Dim X As Double, Y As Double, Z As Double

    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).ClearContents

    For X = lrj To lra Step lrj
    For Y = 2 To lra Step lrj
    For Z = 2 To lri
    
        Cells(Z, "J") = Cells(X, "F") - Cells(Y, "C")
        
        X = X + lrj - 1
        Y = Y + lrj - 1
        
        Cells(Z, "J").NumberFormat = "0.00"
        
    Next
    Next
    Next


    For X = 2 To lrj
    For Y = 2 To lra Step lrj
    For Z = 2 To lri
    
        Cells(Z, "K") = (Cells(X, "J") / Cells(Y, "C"))
        
        X = X + 1
        Y = Y + lrj - 1
        
        Cells(Z, "K").NumberFormat = "0.00%"

    Next
    Next
    Next

    
Dim W As Long, WW As Long
    X = 2
    
    For Y = 2 To lra Step lrj - 1
    
        Cells(X, "L") = WorksheetFunction.Sum(Range("G" & Y).Resize(lrj - 1))
        X = X + 1
        
    Next

    
    
    
    
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0.00"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = vbGreen
    
    End With
    
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0.00"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = vbRed
    
    End With
    
    
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
    
     Range("I:I").Copy Range("S:S")
     Range("K:K").Copy Range("T:T")
     
     ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("T2:T3001") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("S1:T3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    
    Range("S2:T2").Copy Range("P2:Q2")
    
    
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("T2:T3001") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("S1:T3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("S2:T2").Copy Range("P3:Q3")
    
    Range("S:T").ClearContents
    
    
    
    Range("I:I").Copy Range("S:S")
    Range("L:L").Copy Range("T:T")
    
    
     ActiveSheet.Sort.SortFields.Clear
     ActiveSheet.Sort.SortFields.Add Key:=Range("T2:T3001") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("S1:T3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
     
     Range("S2:T2").Copy Range("P4:Q4")
     Range("S:T").ClearContents
     
    Columns("I:Q").EntireColumn.AutoFit
    Range("Q2:Q3").NumberFormat = "0.00%"
    Range("Q4").NumberFormat = "0.00E+00"
    Range("A1").Select
End Sub

