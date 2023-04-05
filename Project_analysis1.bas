Attribute VB_Name = "Module1"
Sub Project_analysis()

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
            ticker_one = ws.Cells(i, 1).Value
            ticker_two = ws.Cells(i + 1, 1).Value
            If ticker_one <> ticker_two Then
                ws.Range("I" & ticker_var).Value = ticker_one
                ticker_var = ticker_var + 1
            End If
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                year_cl = ws.Cells(i, 6).Value
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                year_op = ws.Cells(i, 3).Value
            End If
            
            If year_op > 0 And year_cl > 0 Then
                increase = year_cl - year_op
                increase_per = increase / year_op
                ws.Cells(select_index, 10).Value = increase
                ws.Cells(select_index, 11).Value = FormatPercent(increase_per)
                year_cl = 0
                year_op = 0
                select_index = select_index + 1
                ws.Range("J" & ticker_var).NumberFormat = "0.00"
            End If
          Next i
          
          LR_two = ws.Cells(Rows.Count, 10).End(xlUp).Row
          
          For i = 2 To LR_two
          
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
            
         Next i
          
          Next ws
          
    End Sub
    
