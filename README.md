VBA Challenge

Sub StockHW()

' Define & Initialize variables we'll use in script
' Thanks Sharon for setting me up with the variables

' Ticker symbol
Dim ticker As String
ticker = ""

' Number of tickers for each worksheet
Dim number_tickers As Integer
number_tickers = 0

' Last row in each worksheet
Dim lastRowState As Long

' Opening price for specific year
Dim opening_price As Double
opening_price = 0

' Closing price for specific year
Dim closing_price As Double

' Yearly change
Dim yearly_change As Double
yearly_change = 0

' Percent change
Dim percent_change As Double
percent_change = 0

' Total stock volume
Dim total_stock_volume As Double
total_stock_volume = 0

' BONUS variables
' Greatest percent increase value for specific year
Dim greatest_percent_increase As Double

' Ticker that has the greatest percent increase
Dim greatest_percent_increase_ticker As String

' Greatest percent decrease value for specific year
Dim greatest_percent_decrease As Double

' Ticker that has the greatest percent decrease
Dim greatest_percent_decrease_ticker As String

' Greatest stock volume value for specific year
Dim greatest_stock_volume As Double

' Ticker that has the greatest stock volume
Dim greatest_stock_volume_ticker As String

' Loop through each worksheet in the workbook
For Each ws In Worksheets

    ws.Activate

    ' Find the last row of each worksheet
    lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Header columns for each worksheet
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Loop through the list of tickers
    For i = 2 To lastRowState

        ' Get the value of the ticker symbol for each stock
        ticker = Cells(i, 1).Value
        
        ' Opening price for the ticker stock
        If opening_price = 0 Then
            opening_price = Cells(i, 3).Value
        End If
        
        ' Adding total stock volume values for a tickers
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
        ' Total for different tickers
        If Cells(i + 1, 1).Value <> ticker Then
            
            ' Add tickers for each row
            number_tickers = number_tickers + 1
            Cells(number_tickers + 1, 9) = ticker
            
            ' End of the year closing price for ticker
            closing_price = Cells(i, 6)
            
            ' Get yearly change value
            yearly_change = closing_price - opening_price
            
            ' Yearly change value in each worksheet
            Cells(number_tickers + 1, 10).Value = yearly_change
            
            ' Positive yearly change Green
            If yearly_change > 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
            
            ' Negative yearly change Red
            ElseIf yearly_change < 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
            
            ' No yearly change Yellow
            Else
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 6
            End If
            
            ' Percent change value for ticker
            If opening_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / opening_price)
            End If
            
            ' Format the percent_change value as a percent | Stack Overflow
            Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
            
            ' Set opening price back to 0 when we get to a different ticker in the list | Stack Overflow
            opening_price = 0
            
            ' Total stock volume value in each worksheet
            Cells(number_tickers + 1, 12).Value = total_stock_volume
            
            ' Total stock volume back to 0 when we get to a different ticker in the list
            total_stock_volume = 0
        End If
        
    Next i
    
    'BONUS Variables in ws
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    ' Last row
    lastRowState = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    ' Initialize variables and set values of variables initially to the first row in the list.
    greatest_percent_increase = Cells(2, 11).Value
    greatest_percent_increase_ticker = Cells(2, 9).Value
    greatest_percent_decrease = Cells(2, 11).Value
    greatest_percent_decrease_ticker = Cells(2, 9).Value
    greatest_stock_volume = Cells(2, 12).Value
    greatest_stock_volume_ticker = Cells(2, 9).Value
    
    
    ' Loop through the list of tickers excluding header
    For i = 2 To lastRowState
    
        ' Find the ticker with the greatest percent increase.
        If Cells(i, 11).Value > greatest_percent_increase Then
            greatest_percent_increase = Cells(i, 11).Value
            greatest_percent_increase_ticker = Cells(i, 9).Value
        End If
        
        ' Find the ticker with the greatest percent decrease.
        If Cells(i, 11).Value < greatest_percent_decrease Then
            greatest_percent_decrease = Cells(i, 11).Value
            greatest_percent_decrease_ticker = Cells(i, 9).Value
        End If
        
        ' Find the ticker with the greatest stock volume.
        If Cells(i, 12).Value > greatest_stock_volume Then
            greatest_stock_volume = Cells(i, 12).Value
            greatest_stock_volume_ticker = Cells(i, 9).Value
        End If
        
    Next i
    
    ' Add the values for greatest percent increase, decrease, and stock volume to each worksheet | Thank you to my peers for helping me with this
    Range("P2").Value = Format(greatest_percent_increase_ticker, "Percent")
    Range("Q2").Value = Format(greatest_percent_increase, "Percent")
    Range("P3").Value = Format(greatest_percent_decrease_ticker, "Percent")
    Range("Q3").Value = Format(greatest_percent_decrease, "Percent")
    Range("P4").Value = greatest_stock_volume_ticker
    Range("Q4").Value = greatest_stock_volume
    
Next ws
End Sub
