Sub VBAStocks()

 Dim ws As Worksheet
 Dim starting_ws As Worksheet
 Set starting_ws = ActiveSheet
 
 For Each ws In ThisWorkbook.Worksheets
       ws.Activate

            Dim LastRow As Long
            Dim Ticker As String
            Dim Next_Open_Price As Double
            Next_Open_Price = Cells(2, 3).Value 'reference first ticker open price
            Dim Close_Price As Double
            Close_Price = 0
            Dim YearlyChange As Double
            YearlyChange = 0
            Dim PercentChange As Double
            PercentChange = 0
            Dim T_Volume As Double
            Ticker_Volume = 0
            
         
            
            Dim Summary_Table_Row As Integer 'Keep track of the location for each ticker in the summary table
            Summary_Table_Row = 2 'Start on row 2 after header
            
           
            
            'Determine the last row
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row '***remove ws in front of Cells
           
            
            'Create headers
                Cells(1, 9).Value = "Ticker"
                Cells(1, 10).Value = "Yearly Change"
                Cells(1, 11).Value = "Percent Change"
                Cells(1, 12).Value = "Total Stock Volume"
                Cells(1, 16).Value = "Ticker"
                Cells(1, 17).Value = "Value"
            
            
            
            'looping through name of the tickers
            For I = 2 To LastRow
              'Ticker open value
              
              'Check to see if it's the same ticker name
                If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
                    'set ticker name
                    Ticker = Cells(I, 1).Value
                    T_Volume = T_Volume + Cells(I, 7).Value
                    Close_Price = Cells(I, 6).Value ' hold close price
                    YearlyChange = Close_Price - Next_Open_Price
                    'fix divide by zero error
                    If Next_Open_Price <> 0 Then
                        PercentChange = YearlyChange / Next_Open_Price
                    Else
                       PercentChange = 0
                    End If
                                
                    Range("P" & 2).Value = Greatest_Perc_Incr_Ticker
                    Range("I" & Summary_Table_Row).Value = Ticker 'Print ticker name in column I
                    Range("L" & Summary_Table_Row).Value = T_Volume  'Print ticker volume in column L
                    If YearlyChange > 0 Then
                       Range("J" & Summary_Table_Row).Value = YearlyChange
                       Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                     Else
                       Range("J" & Summary_Table_Row).Value = YearlyChange
                       Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
                    Range("K" & Summary_Table_Row).Value = PercentChange
                    Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                    
                    Summary_Table_Row = Summary_Table_Row + 1
                    T_Volume = 0 'Reset ticker volume total for next ticker
                    Close_Price = 0 'Reset ticker close price
                    Next_Open_Price = Cells(I + 1, 3).Value
                    'Open_Price = 0 'Reset ticker open price
                    YearlyChange = 0
                    PercenChange = 0
                Else
                    T_Volume = T_Volume + Cells(I, 7).Value
                    
                  
                End If
                 
                        
              
            Next I
            
            '--------------Greatest Increase, Decrease, and volume
            
            
            Dim Greatest_Perc_Decr_Ticker As String
            Dim Greatest_Tot_Vol_Ticker As String
            Dim Greatest_Perc_Incr_Value As Double
            Dim Greatest_Perc_Decr_Value As Double
            Dim Greatest_Tot_Vol_Value As Double
            
            
            Greatest_Perc_Incr_Ticker = Cells(2, 9).Value
            Greatest_Perc_Inc_Value = Cells(2, 11).Value
            
            Greatest_Perc_Decr_Ticker = Cells(2, 9).Value
            Greatest_Perc_Decr_Value = Cells(2, 11).Value
            
            Greatest_Tot_Vol_Ticker = Cells(2, 9).Value 'point to first ticker record, cells(2,9) row 2, col I
            Greatest_Tot_Vol_Value = Cells(2, 12).Value 'point to first volume record, row 2, col L
            
            For x = 2 To Range("I" & Rows.Count).End(xlUp).Row
                If Cells(x + 1, 11).Value > Greatest_Perc_Inc_Value Then
                    Greatest_Perc_Incr_Ticker = Cells(x + 1, 9).Value
                    Greatest_Perc_Inc_Value = Cells(x + 1, 11).Value 'comparing tickers yearly change percentage in column 11 (K)
                End If
                
                If Cells(x + 1, 11).Value < Greatest_Perc_Decr_Value Then
                    Greatest_Perc_Decr_Ticker = Cells(x + 1, 9).Value
                    Greatest_Perc_Decr_Value = Cells(x + 1, 11).Value 'comparing tickers yearly change percentage in column 11 (K)
                End If
                
                If Cells(x + 1, 12).Value > Greatest_Tot_Vol_Value Then
                    Greatest_Tot_Vol_Ticker = Cells(x + 1, 9).Value
                    Greatest_Tot_Vol_Value = Cells(x + 1, 12).Value
                End If
                
                
            Next x
            

            Cells(2, 15).Value = "Greatest % Increase"
            Cells(2, 16).Value = Greatest_Perc_Incr_Ticker
            Cells(2, 17).Value = Format(Greatest_Perc_Inc_Value, "0.00%")

            Cells(3, 15).Value = "Greatest % Decrease"
            Cells(3, 16).Value = Greatest_Perc_Decr_Ticker
            Cells(3, 17).Value = Format(Greatest_Perc_Decr_Value, "0.00%")

            Cells(4, 15).Value = "Greatest Total Volume"
            Cells(4, 16).Value = Greatest_Tot_Vol_Ticker
            Cells(4, 17).Value = Greatest_Tot_Vol_Value


        ws.Cells(1, 1) = 1
        
 Next ws
 
 starting_ws.Activate
 
 
End Sub

   


