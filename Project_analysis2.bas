Attribute VB_Name = "Module2"
Sub Project_analysis2()

    Dim ws As Worksheet
    Dim select_index As Double
    Dim first_row As Double
    Dim select_row As Double
    Dim LR As Double
    Dim year_op As Single
    Dim year_cl As Single
    Dim volume As Double

    
    For Each ws In Sheets
        Worksheets(ws.Name).Activate
        select_index = 2
        ticker_var = 2
        LR = ws.Cells(Rows.Count, 1).End(xlUp).Row
        volume = 0
        
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greastest Total Volume"
        
        
           For i = 2 To LR
           
           max_k = WorksheetFunction.Max(ActiveSheet.Columns("k"))
           min_k = WorksheetFunction.Min(ActiveSheet.Columns("k"))
           
           ws.Range("Q2").Value = max_k
           ws.Range("Q3").Value = min_k
           ws.Range("Q2:Q3").NumberFormat = "0.00%"
           
           If max_k = Cells(i, 11).Value Then
                ws.Range("P2").Value = Cells(i, 9).Value
            ElseIf min_k = Cells(i, 11).Value Then
                ws.Range("P3").Value = Cells(i, 9).Value
            End If
        Next i
           
            
    Next ws
                


End Sub
