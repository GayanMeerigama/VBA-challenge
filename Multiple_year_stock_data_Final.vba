Sub MultiyearStock()

 For Each ws In Worksheets  'VBA script to allow it to run on every worksheet (that is, every year) just by running the VBA script once


    Dim TickerSymbol As String
    
    Dim TotalVol As LongLong
    
    Dim YearlyChangeFirst  As Double
    
    Dim YearlyChangeLast  As Double
    
    Dim SummaryTableRow As Double
    
    
    Dim cell1 As Range, cell2 As Range
    
    Dim last_raw As Integer
    
    last_raw = Cells(Rows.Count, 1).End(xlUp).Row
    

    SummaryTableRow = 2
    
    YearlyChangeFirst = Cells(2, 3).Value
    
    For i = 2 To last_raw
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        TickerSymbol = Cells(i, 1).Value
        
        TotalVol = TotalVol + Cells(i, 7).Value
        Cells(SummaryTableRow, 9).Value = TickerSymbol   'The ticker symbol  
        Cells(SummaryTableRow, 12).Value = TotalVol         'The total stock volume of the stock
        SummaryTableRow = SummaryTableRow + 1
        
        TotalVol = 0
        
        YearlyChangeFirst = Cells(i + 1, 3).Value    'opening price at the beginning of a given year
             'Cells(SummaryTableRow, 12).Value = YearlyChangeFirst
        
    
        Else
        
           
        
            TotalVol = TotalVol + Cells(i, 7).Value
            
            YearlyChangeLast = Cells(i + 1, 6).Value   'closing price at the end of that year
             'Cells(SummaryTableRow, 11).Value = YearlyChangeLast
          
            
             
        
            
         End If
         
         Cells(SummaryTableRow, 10).Value = YearlyChangeLast - YearlyChangeFirst 'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year
         
            'highlight positive change in green and negative change in red.

            Set cell1 = Range("J" & i)
           
            If cell1.Value >= 0 Then cell1.Interior.Color = vbGreen
            If cell1.Value <= 0 Then cell1.Interior.Color = vbRed
            

            Cells(SummaryTableRow, 11).Value = ((YearlyChangeLast - YearlyChangeFirst) / YearlyChangeFirst) * 100 'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    
    
            'functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"

            Cells(2, 17).Value = Application.WorksheetFunction.Max(Range("K2:K3001"))
            Cells(3, 17).Value = Application.WorksheetFunction.Min(Range("K2:K3001"))
            Cells(4, 17).Value = Application.WorksheetFunction.Max(Range("L2:L3001"))
            
            Cells(2, 16).Value = WorksheetFunction.Index(Range("I2:I3001"), WorksheetFunction.Match(Cells(2, 17).Value, Range("K2:K3001"), 0))
            Cells(3, 16).Value = WorksheetFunction.Index(Range("I2:I3001"), WorksheetFunction.Match(Cells(3, 17).Value, Range("K2:K3001"), 0))
            Cells(4, 16).Value = WorksheetFunction.Index(Range("I2:I3001"), WorksheetFunction.Match(Cells(4, 17).Value, Range("L2:L3001"), 0))


            'last_raw = Cells(Rows.Count).End(xlUp).Row
    
            'Cells(2, 17).Value = Application.WorksheetFunction.Max(Range("K2" & last_raw))
            'Cells(3, 17).Value = Application.WorksheetFunction.Min(Range("K2" & last_raw))
            'Cells(4, 17).Value = Application.WorksheetFunction.Max(Range("L2" & last_raw))
            
            'Cells(2, 16).Value = WorksheetFunction.Index(Range("I2" & last_raw), WorksheetFunction.Match(Cells(2, 17).Value, Range("K2" & last_raw), 0))
            'Cells(3, 16).Value = WorksheetFunction.Index(Range("I2" & last_raw), WorksheetFunction.Match(Cells(3, 17).Value, Range("K2" & last_raw), 0))
            'Cells(4, 16).Value = WorksheetFunction.Index(Range("I2" & last_raw), WorksheetFunction.Match(Cells(4, 17).Value, Range("L2" & last_raw), 0))

               
    
    Next i


Next ws

End Sub

