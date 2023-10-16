VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_analysis()
Dim i As Long
Dim j As Integer
Dim total As Double
Dim rowCount As Long
Dim start As Long
Dim percentChange As Double
Dim change As Double
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
start = 2
total = 0
change = 0
rowCount = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To rowCount
    If (Cells(i, 1).Value <> Cells(i + 1, 1).Value) Then
    
    total = total + Cells(i, 7).Value
         
        If total = 0 Then
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Range("J" & 2 + j).Value = 0
            Range("K" & 2 + j).Value = "%" & 0
            Range("L" & 2 + j).Value = 0
        Else
            If Cells(start, 3) = 0 Then
                For find_val = start To i
                    If Cells(find_val, 3) <> 0 Then
                        start = find_val
                        Exit For
                    End If
                Next find_val
            End If
            
            change = (Cells(i, 6) - Cells(start, 3))
            percentChange = change / Cells(start, 3)
            start = i + 1
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Range("J" & 2 + j).Value = change
            Range("J" & 2 + j).NumberFormat = "0.00"
            Range("K" & 2 + j).Value = percentChange
            Range("K" & 2 + j).NumberFormat = "0.00"
            Range("L" & 2 + j).Value = total
            
            If change > 0 Then
                Range("J" & 2 + j).Interior.ColorIndex = 4
            ElseIf change < 0 Then
                Range("J" & 2 + j).Interior.ColorIndex = 3
            Else
                Range("J" & 2 + j).Interior.ColorIndex = 0
            End If
             
        End If
        
        total = o
        change = 0
        j = j + 1
    
        Else
            total = total + Cells(i, 7).Value
    End If

Next i

Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowCount)) * 100
Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowCount)) * 100
Range("Q4") = WorksheetFunction.Max(Range("L2:L" & rowCount))

Range("P2") = Cells(WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0) + 1, 9)
Range("P3") = Cells(WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0) + 1, 9)
Range("P4") = Cells(WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowCount)), Range("L2:L" & rowCount), 0) + 1, 9)

End Sub
