# VBA-Challenge
VBA HW

Sub StockHW()

'Declare and set worksheet
Dim ws As Worksheet
    
    'Loop through all stocks for one year
    For Each ws In Worksheets
        
        'Set column headings
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Set variables for prices and stock changes
        Dim open_price As Double
        open_price = 0
        Dim close_price As Double
        close_price = 0
        Dim price_change As Double
        yearly_price_change = 0
        Dim price_change_percent As Double
        price_change_percent = 0
        
        'Define variable for Ticker
        Dim Ticker As String
        Ticker = " "
        
        'Define variable for Stock Volume
        Dim stock_volume As Double
        stock_volume = 0
        
        'Set initial and last row for worksheet
        Dim Lastrow As Long
        Dim i As Long
        Dim j As Integer
        
        'Define Lastrow of worksheet
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Do loop of current worksheet to Lastrow
        For i = 2 To Lastrow
            
            'Ticker symbol output
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            TickerRow = TickerRow + 1
            Ticker = ws.Cells(i, 1).Value
            ws.Cells(TickerRow, "I").Value = Ticker
            End If
            
            'Calculate Total Stock Volume 'Help from StackOverflow Notes
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
            stock_volume = stock_volume + ws.Cells(i, 7).Value
            
            ElseIf ws.Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            stock_volume = stock_volume + ws.Cells(i, 7).Value
            ws.Cells(i, 12).Value = ws.Cells(i, 1).Value
            ws.Cells(i, "L").Value = stock_volume
            j = j + 1
            stock_volume = 0
            End If
            
            'Calculate change in Stock Prices
            close_price = ws.Cells(i, 6).Value
            open_price = ws.Cells(i, 3).Value
            yearly_price_change = close_price - open_price
            ws.Cells(i, "J").Value = yearly_price_change
            
            'Fixing the open price equal zero problem 'Thanks Sharon Temple
            If open_price <> 0 Then
            price_change_percent = (close_price - open_price) / open_price
            ws.Cells(i, "K").Value = price_percent_change
            
            End If
            
            'Add colors to +/- percent change
            If yearly_change < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3 'Red color
            Else
            ws.Cells(i, 10).Interior.ColorIndex = 4 'Green color
            
            End If
    
        Next i
    
    Next ws

End Sub
