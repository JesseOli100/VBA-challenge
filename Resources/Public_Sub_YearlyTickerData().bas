Public Sub YearlyTickerData()
    
    'Setting this as the variable for the worksheets
    Dim ws As Worksheet
    For Each ws In Worksheets
        Dim W_Name As String
        Dim g As Long 'This will set the current row
        Dim k As Long 'This starts the row for the ticker section
        Dim Ticker_Count As Long 'This will serve as the variable to fill out each Ticker per row
        Dim Last_R_A As Long 'This sets the last row for column A
        Dim Last_R_I As Long 'This sets the last row for column I
        Dim Percent_Change As Double 'This will help calculate the percent change
        Dim Greatest_Increase As Double 'This will help calculate the greatest increase change
        Dim Greatest_Decrease As Double 'This will help calculate the greatest decrease change
        Dim Greatest_Volume As Double 'This will help calculate the greatest volume change

        W_Name = ws.Name 'Sets the worksheet name
        
        'All of these just create the headers for each section
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Sets the row of where the ticker will start and where the rest of the inputs will start
        Ticker_Count = 2
        k = 2
        
        'Finds the last filled out cell in column A
        Last_R_A = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        'This loops through every row
        For g = 2 To Last_R_A
        
            'Checks for ticker value changes
            If ws.Cells(g + 1, 1).Value <> ws.Cells(g, 1).Value Then
            
                
                ws.Cells(Ticker_Count, 9).Value = ws.Cells(g, 1).Value 'This sets the ticker values on I
                ws.Cells(Ticker_Count, 10).Value = ws.Cells(g, 6).Value - ws.Cells(k, 3).Value 'This is calculation for the Yearly Change
                
                'Setting conditional formatting for color changes
                If ws.Cells(Ticker_Count, 10).Value < 0 Then
                    ws.Cells(Ticker_Count, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(Ticker_Count, 10).Interior.ColorIndex = 4
                End If
                
                'This sets the calculation for percent change and formatting
                If ws.Cells(k, 3).Value <> 0 Then
                    Percent_Change = ((ws.Cells(g, 6).Value - ws.Cells(k, 3).Value) / ws.Cells(k, 3).Value)
                    ws.Cells(Ticker_Count, 11).Value = Format(Percent_Change, "Percent")
                Else
                    ws.Cells(Ticker_Count, 11).Value = Format(0, "Percent")
                End If
                
                'This sets the calculation for the total volume
                ws.Cells(Ticker_Count, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(k, 7), ws.Cells(g, 7)))
                
                'By adding 1 to each iteration in the loop, this tells the script to run everything again once it finds a new ticker
                Ticker_Count = Ticker_Count + 1
                k = g + 1
            End If
        Next g

        'Finds the last filled out cell in I
        Last_R_I = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
        
        'Setting these up for the summary
        Greatest_Volume = ws.Cells(2, 12).Value
        Greatest_Increase = ws.Cells(2, 11).Value
        Greatest_Decrease = ws.Cells(2, 11).Value
        
        'Greatest Volume
        For g = 2 To Last_R_I
            If ws.Cells(g, 12).Value > Greatest_Volume Then
                Greatest_Volume = ws.Cells(g, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(g, 9).Value
            End If
            
            'Greatest Increase
            If ws.Cells(g, 11).Value > Greatest_Increase Then
                Greatest_Increase = ws.Cells(g, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(g, 9).Value
            End If
            
            'Greatest decrease
            If ws.Cells(g, 11).Value < Greatest_Decrease Then
                Greatest_Decrease = ws.Cells(g, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(g, 9).Value
            End If

            'Writes the results in the specified cells based on the necessary format
            ws.Cells(2, 17).Value = Format(Greatest_Increase, "Percent")
            ws.Cells(3, 17).Value = Format(Greatest_Decrease, "Percent")
            ws.Cells(4, 17).Value = Format(Greatest_Volume, "Scientific")
        
        Next g

        'Adjusts column width
        Worksheets(W_Name).Columns("A:Z").AutoFit
    Next ws
End Sub
