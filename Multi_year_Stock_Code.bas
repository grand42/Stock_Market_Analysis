Attribute VB_Name = "Module1"
Sub Stock_Market()

'Loop through each worksheet

For Each ws In Worksheets
 
'Find the last row with data
    
    Dim lastcolumn As Long
    Dim lastrow As Long
    
       
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'Create Summary Table
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "% Change"
    ws.Cells(1, 12).Value = "Total Stock Value"
   
'Define Variables
    
    Dim summary_table_row As Long
    summary_table_row = 2
    
    Dim TotalStock As LongLong
    Dim ticker As String
    Dim Stock_Open As Double
    Dim Stock_Close As Double
    Dim Yearly_Change As Double
    Dim PctChange As Double
    Dim i As Long
    
'Loop through the rows to sum the total stock value for each ticker
 
    TotalStock = 0
    Stock_Open = ws.Cells(2, 3).Value
    PctChange = 0
    Stock_Close = 0
    
    For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
        
        'Calculate Total Stock
            
            ticker = ws.Cells(i, 1).Value
            TotalStock = TotalStock + ws.Cells(i, 7).Value
            ws.Cells(summary_table_row, 9).Value = ticker
            ws.Cells(summary_table_row, 12).Value = TotalStock
            
        'Calculate Yearly Change
            
            Stock_Close = ws.Cells(i, 6).Value
            Yearly_Change = Stock_Close - Stock_Open
            ws.Cells(summary_table_row, 10).Value = Yearly_Change
            
        'Calculate Percent Change
            If Stock_Open <> 0 Then
                PctChange = (Stock_Close - Stock_Open) / (Stock_Open)
                ws.Cells(summary_table_row, 11).Value = PctChange
                ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
            
            End If
            
        'Color positive change as green and negative change as red
            
            If Yearly_Change < 0 Then
                ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
        
            Else
                ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
        
            End If
        
            summary_table_row = summary_table_row + 1
            TotalStock = 0
            Stock_Open = ws.Cells(i + 1, 3)
        
        
        Else
            TotalStock = TotalStock + CLng(ws.Cells(i, 7).Value)
       
        End If
    
    Next i
    
    
 'Find stock with greatest increase
 
 Dim ticker_num As Integer
 Dim max As LongLong
 Dim max_growth As Double
 Dim max_loss As Double
 Dim growth_range As Range
 Dim stock_range As Range

 
 
'Create Summary Table
 
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Stock Value"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
 
'Calculate Ranges

    ticker_num = ws.Cells(Rows.Count, 9).End(xlUp).Row
    

    Set growth_range = ws.Range(ws.Cells(2, 11), ws.Cells(ticker_num, 11))
    
    
    
    max_growth = Application.WorksheetFunction.max(growth_range)
    max_loss = Application.WorksheetFunction.Min(growth_range)
    
    ws.Cells(2, 16).Value = max_growth
    ws.Cells(2, 16).NumberFormat = "0.00%"
    ws.Cells(3, 16).Value = max_loss
    ws.Cells(3, 16).NumberFormat = "0.00%"

    Set stock_range = ws.Range(ws.Cells(2, 12), ws.Cells(ticker_num, 12))
    max = Application.WorksheetFunction.max(stock_range)
    
    ws.Cells(4, 16).Value = max
    
'Find ticker to correspond with values

Dim Max_growth_ticker As String
Dim Max_stock_ticker As String
Dim Max_loss_ticker As String
Dim j As Double
Dim ii As Double
Dim jj As Double

'Find ticker with  max growth

    For j = 2 To ticker_num

        If ws.Cells(j, 11).Value = max_growth Then
        
            Max_growth_ticker = ws.Cells(j, 9).Value
            ws.Cells(2, 15).Value = Max_growth_ticker
            Exit For
        End If
        
        Next j
        
'Find ticker with greatest decrease %

        For ii = 2 To ticker_num

        If ws.Cells(ii, 11).Value = max_loss Then
        
            Max_loss_ticker = ws.Cells(ii, 9).Value
            ws.Cells(3, 15).Value = Max_loss_ticker
            Exit For
        End If
        
        Next ii
        
'Find ticker with greatest stock value
        
        For jj = 2 To ticker_num

        If ws.Cells(jj, 12).Value = max Then
        
            Max_stock_ticker = ws.Cells(jj, 9).Value
            ws.Cells(4, 15).Value = Max_stock_ticker
            Exit For
        End If
        
        Next jj
            
       

 Next ws

End Sub
