Sub Assignment2()

    For Each ws In Worksheets
    
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
            
        Dim Ticker As String
        Dim Quarterly_Change As Double
        Dim Percent_Change As Double
        Dim Total_Stock_Volume As Double
        Dim Greatest_Increase_Ticker As String
        Dim Greatest_Increase As Double
        Dim Greatest_Decrease_Ticker As String
        Dim Greatest_Decrease As Double
        Dim Greatest_Volume_Ticker As String
        Dim Greatest_Volume As Double
                        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Quarterly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To LastRow
       
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Ticker = Cells(i, 1).Value
                Range("I" & Summary_Table_Row).Value = Ticker
                Quarterly_Change = Cells(i, 6) - Cells(i, 3)
                Range("J" & Summary_Table_Row).Value = Quarterly_Change
                Percent_Change = Quarterly_Change / Cells(i, 3)
                Range("K" & Summary_Table_Row).Value = Percent_Change
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                Summary_Table_Row = Summary_Table_Row + 1
                Total_Stock_Volume = 0
                Else
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            End If
            
            With Cells(i, 11)
                .NumberFormat = "0.00%"
                .Value = .Value
            End With
                        
            If Cells(i, 10).Value >= 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
                Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
                                 
        Greatest_Increase = WorksheetFunction.Max(Range("K" & Summary_Table_Row))
            Range("Q2").Value = Greatest_Increase
            Greatest_Increase_Ticker = Range("I" & Summary_Table_Row).Value
            Range("P2").Value = Greatest_Increase_Ticker
                            
        Greatest_Decrease = WorksheetFunction.Min(Range("K" & Summary_Table_Row))
            Range("Q3").Value = Greatest_Decrease
            Greatest_Decrease_Ticker = Range("I" & Summary_Table_Row).Value
            Range("P3").Value = Greatest_Decrease_Ticker
        
        Greatest_Volume = WorksheetFunction.Max(Range("L" & Summary_Table_Row))
            Range("Q4").Value = Greatest_Volume
            Greatest_Volume_Ticker = Range("I" & Summary_Table_Row).Value
            Range("P4").Value = Greatest_Volume_Ticker
                                                                    
        Next i
    
    Next ws

End Sub