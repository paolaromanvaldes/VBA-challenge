Attribute VB_Name = "Module1"
Sub alpha_data_analysis():

    For Each ws In Worksheets
        
    'set Variables
    
        Dim sheet_nm As String
        Dim ticker_nm As Long
        Dim data_last_row As Long
        Dim results_last_row As Long
        Dim percentage_change As Double
      
    'set column and row values
    
        Dim i As Long
        Dim j As Long

        
        'part one headers
        
        ws.Cells(1, 11).Value = "ticker"
        ws.Cells(1, 12).Value = "yearly change"
        ws.Cells(1, 13).Value = "percentage change"
        ws.Cells(1, 14).Value = "total ctock volume"
        ws.Cells(1, 18).Value = "ticker"
        ws.Cells(1, 19).Value = "value"
        ws.Cells(2, 17).Value = "max increase"
        ws.Cells(3, 17).Value = "max decrease"
        ws.Cells(4, 17).Value = "max total volume"
        
        data_last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
        ticker_nm = 2
        
        j = 2
        
        'part one check ticker and yearly change
            For i = 2 To data_last_row
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(ticker_nm, 11).Value = ws.Cells(i, 1).Value
                ws.Cells(ticker_nm, 12).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    'color conditonal based on change ang change to currency for yearly change
                    If ws.Cells(ticker_nm, 12).Value < 0 Then
                    ws.Cells(ticker_nm, 12).Interior.ColorIndex = 3
                    Else
                    ws.Cells(ticker_nm, 12).Interior.ColorIndex = 4
                    End If
                    Range("L:L").NumberFormat = "$#,##0.00"
                    
                    'percentage change
                    If ws.Cells(j, 3).Value <> 0 Then
                    percentage_change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    ws.Cells(ticker_nm, 13).Value = Format(percentage_change, "Percent")
                    Else
                    ws.Cells(ticker_nm, 11).Value = Format(0, "Percent")
                    End If
                    
                ws.Cells(ticker_nm, 14).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'ticker move on
                ticker_nm = ticker_nm + 1
                
                'change row of ticker
                j = i + 1
                
                End If
            
            Next i
            
    'part 2 loop
    
    'set variables part 2
        Dim greatest_increase As Double
        Dim greatest_decrease As Double
        Dim volume_change As Double
    
    'set headers part 2
        ws.Cells(1, 18).Value = "ticker"
        ws.Cells(1, 19).Value = "value"
        ws.Cells(2, 17).Value = "max increase percentage"
        ws.Cells(3, 17).Value = "max decrease percentage"
        ws.Cells(4, 17).Value = "greatest total volume"
        
        results_last_row = ws.Cells(Rows.Count, 11).End(xlUp).Row
   
        volume_change = ws.Cells(2, 14).Value
        greatest_increase = ws.Cells(2, 13).Value
        greatest_decrease = ws.Cells(2, 13).Value
    
            For i = 2 To results_last_row
            
                If ws.Cells(i, 14).Value > volume_change Then
                volume_change = ws.Cells(i, 14).Value
                ws.Cells(4, 18).Value = ws.Cells(i, 11).Value
                Else
                volume_change = volume_change
                End If
                
                If ws.Cells(i, 13).Value > greatest_increase Then
                ws.Cells(2, 18).Value = ws.Cells(i, 11).Value
                Else
                greatest_increase = greatest_increase
                End If
                
                If ws.Cells(i, 13).Value < greatest_decrease Then
                greatest_decrease = ws.Cells(i, 13).Value
                ws.Cells(3, 18).Value = ws.Cells(i, 11).Value
                Else
                greatest_decrease = greatest_decrease
                End If
                
            'part 2 results chart print
            
            ws.Cells(2, 19).Value = Format(greatest_increase, "Percent")
            ws.Cells(3, 19).Value = Format(greatest_decrease, "Percent")
            ws.Cells(4, 19).Value = Format(volume_change, "Scientific")
            
            Next i
            
    Next ws
        
End Sub

