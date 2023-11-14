Option Explicit
 
Sub Stock_market()

'Declare and set worksheet
Dim ws As Worksheet

'Loop through all stocks for one year
For Each ws In Worksheets


'Create the column headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Define Ticker variable
Dim Ticker As String
Ticker = " "
Dim Ticker_volume As Double
Ticker_volume = 0

'Create variable to hold stock volume
Dim stock_volume As Double
stock_volume = 0

'Set initial and last row for worksheet
Dim Lastrow As Long
Dim i As Long
Dim j As Integer

'Define Lastrow of worksheet
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set new variables for prices and percent changes
Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0
Dim price_change As Double
price_change = 0
Dim price_change_percent As Double
price_change_percent = 0
Dim rowcount As Long
rowcount = 2

Dim TickerRow As Long: TickerRow = 1

'Do loop of current worksheet to Lastrow
For i = 2 To Lastrow

'Conditional to grab year open price
If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
open_price = ws.Cells(i, 3).Value

End If

'Total up the volume for each row to get the total stock volume for the year
stock_volume = stock_volume + ws.Cells(i, 7)

'Conditional to determine if the ticker symbol is changing
If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

'Move ticker symbol to summary table
ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value

'Move total stock volume to the summary table
ws.Cells(rowcount, 12).Value = stock_volume

'Grab year end price
close_price = ws.Cells(i, 6).Value

'Calculate the price change for the year and move it to the summary table.
price_change = close_price - open_price
ws.Cells(rowcount, 10).Value = price_change

'Conditional to format to highlight positive or negative change.
If price_change >= 0 Then
    ws.Cells(rowcount, 10).Interior.ColorIndex = 4
    
Else
    ws.Cells(rowcount, 10).Interior.ColorIndex = 3

End If

'Calculate the percent change for the year and move it to the summary table format as a percentage
'Conditional for calculating percent change
If open_price = 0 And close_price = 0 Then
    'Starting at zero and ending at zero will be a zero increase.  Cannot use a formula because
    'it would be dividing by zero.
    price_change_percent = 0
    ws.Cells(rowcount, 11).Value = price_change_percent
    ws.Cells(rowcount, 11).NumberFormat = "0.00%"

ElseIf open_price = 0 Then
    'If a stock starts at zero and increases, it grows by infinite percent.
    'Because of this, we only need to evaluate actual price increase by dollar amount and therefore put
    '"New Stock" as percent change.
    Dim price_change_percent_NA As String
    price_change_percent_NA = "New Stock"
    ws.Cells(rowcount, 11).Value = price_change_percent

Else
    price_change_percent = price_change / open_price
    ws.Cells(rowcount, 11).Value = price_change_percent
    ws.Cells(rowcount, 11).NumberFormat = "0.00%"

End If

'Add 1 to rowcount to move it to the next empty row in the summary table
 rowcount = rowcount + 1
 
'Reset total stock_volume, open_price, close_price, price_change, price_change_percent
    stock_volume = 0
    open_price = 0
    close_price = 0
    price_change = 0
    price_change_percent = 0
    
 End If

Next i
'Create a best/worst performance table

'Assign lastrow to count the number of rows in the summary table
    Lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
'Set variables to hold best performer, worst performer, and stock with the most volume
    Dim best_stock As String
    Dim best_value As Double
    
'Set best performer equal to the first stock
    best_value = ws.Cells(2, 11).Value

    Dim worst_stock As String
    Dim worst_value As Double

'Set worst performer equal to the first stock
    worst_value = ws.Cells(2, 11).Value
    
    Dim most_vol_stock As String
    Dim most_vol_value As Double
    
'Set most volume equal to the first stock
    most_vol_value = ws.Cells(2, 12).Value
    
'Loop to search through summary table
    For j = 2 To Lastrow
    
        'Conditional to determine best performer
        If ws.Cells(j, 11).Value > best_value Then
            best_value = ws.Cells(j, 11).Value
            best_stock = ws.Cells(j, 9).Value
        End If
        
        'Conditional to determine worst performer
        If ws.Cells(j, 11).Value < worst_value Then
            worst_value = ws.Cells(j, 11).Value
            worst_stock = ws.Cells(j, 9).Value
        End If
        
        'Conditional to determine stock with the greatest volume trade
        If ws.Cells(j, 12).Value > most_vol_value Then
            most_vol_value = ws.Cells(j, 12).Value
            most_vol_stock = ws.Cells(j, 9).Value
        End If
        
    Next j
'Move best performer, worst performer, and stock with the most volume items to the performance table
        ws.Cells(2, 16).Value = best_stock
        ws.Cells(2, 17).Value = best_value
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = worst_stock
        ws.Cells(3, 17).Value = worst_value
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = most_vol_stock
        ws.Cells(4, 17).Value = most_vol_value
        
    Next ws

End Sub
