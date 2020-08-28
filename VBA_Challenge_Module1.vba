Attribute VB_Name = "Module1"
Sub VBA_Challenge()

  ' Set an initial variable for holding the ticker
  Dim ticker As String

  ' Set an initial variable for holding the i
  Dim i As Variant
  
  ' Set an initial variable for holding the total for volume
  Dim ticker_volume As Double
  ticker_volume = 0
  
  ' Set an initial variable for holding the max_inc
  Dim low_perc, max_perc, max_vol, maxv_ticker As Variant
  
  ' Set an initial variable for holding the last_row
  Dim last_Row As Variant
  
  ' Set an initial variable for holding the max_ticker
  Dim low_ticker As Variant
  
  
    ' Set an initial variable for holding the yearly change
  Dim yearly_change As Variant
  yearly_change = 0
  
  ' Set an initial variable for holding the worksheet
  Dim ws As Worksheet

      
   
   
    ' Loop through all sheets
    For Each ws In Worksheets
        
        'Keep track of the location for each ticker brand in the summary table
         Dim Summary_Table_Row As Integer
         Summary_Table_Row = 2
      
           
      
        Dim open_price As Double
          open_price = 0
          
        Dim yearly_price As Double
          yearly_price = 0
          
        'Find last row
    last_Row = Cells(Rows.Count, 1).End(xlUp).Row
    
          
        ' Loop through all tickers and volumes WILL NEED LAST ROW
        For i = 2 To last_Row
           
            
           
            ' Check if we are still within the same ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
              
              ' Set the ticker, close and open names and values
                ticker = ws.Cells(i, 1).Value
                close_price = ws.Cells(i, 6).Value
                open_price_index = WorksheetFunction.Match(ticker, ws.Range("a2:a" & i), 0)
                open_price = ws.Cells(open_price_index + 1, 3).Value
                yearly_change = close_price - open_price
                     If open_price <> 0 Then
                     percent_change = yearly_change / open_price
                     End If
                     
                                        
                   
              ' Add to the volume Total
                ticker_volume = ticker_volume + ws.Cells(i, 7).Value
        
              ' Print headers to Summary_Table and greatest_table
                ws.Range("i1").Value = "Ticker"
                ws.Range("j1").Value = "Yearly Change"
                ws.Range("k1").Value = "Pecent Change"
                ws.Range("l1").Value = "Total Stock Volume"
                ws.Range("o1").Value = "Ticker"
                ws.Range("p1").Value = "Value"
                ws.Range("n2").Value = "Greatest % Increase"
                ws.Range("n3").Value = "Greatest % Decrease"
                ws.Range("n4").Value = "Greatest Total Volume"
              
                    
              
                ' Print the stock in the Summary Table
                ws.Range("i" & Summary_Table_Row).Value = ticker
            
            
                ' Print the Yearly Change in the Summary Table
                ws.Range("j" & Summary_Table_Row).Value = yearly_change
                ' Format yearly_change column to display negative in red and positive in green
                    If yearly_change > 0 Then
                    ws.Range("j" & Summary_Table_Row).Interior.ColorIndex = 4
                    Else
                    ws.Range("j" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
                     
                         
                  
                ' Print the Percent Change in the Summary Table
                ws.Range("k" & Summary_Table_Row).Value = percent_change
                ' Format percent_change column to display in percentage
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                    
                ' Print the volume to the Summary Table
                ws.Range("l" & Summary_Table_Row).Value = ticker_volume
                       
            
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                'Find the greatest decrease in percent
                low_perc = WorksheetFunction.Min(ws.Range("k2:k5000"))
                low_ticker = WorksheetFunction.Match(low_perc, ws.Range("k2:k5000"), 0)
                ws.Cells(3, 15).Value = Cells((low_ticker + 1), 9).Value
                ws.Cells(3, 16).Value = low_perc
                ws.Cells(3, 16).NumberFormat = "0.00%"
                
                'Find the greatest inc in percent
                max_perc = WorksheetFunction.Max(ws.Range("k2:k5000"))
                max_ticker = WorksheetFunction.Match(max_perc, ws.Range("k2:k5000"), 0)
                ws.Cells(2, 15).Value = Cells((max_ticker + 1), 9).Value
                ws.Cells(2, 16).Value = max_perc
                ws.Cells(2, 16).NumberFormat = "0.00%"
                'Find the greatest volume
                max_vol = WorksheetFunction.Max(ws.Range("l2:l5000"))
                maxv_ticker = WorksheetFunction.Match(max_vol, ws.Range("l2:l5000"), 0)
                ws.Cells(4, 15).Value = Cells((maxv_ticker + 1), 9).Value
                ws.Cells(4, 16).Value = max_vol
                ws.Cells(4, 16).NumberFormat = "#,###"
                               
                ' Reset the volume and open_price Total
                ticker_volume = 0
                  
                  
                    
                ' If the cell immediately following a row is the same brand...
                Else
        
                ' Add to the Brand Total
                ticker_volume = ticker_volume + ws.Cells(i, 7).Value
                       
              
            End If
                
                
          Next i
    
    
        ' Autofit to display data
       Columns("i:p").AutoFit
             
        
            
 
      
    Next ws
        
       
End Sub



