Attribute VB_Name = "Module1"
Sub Stock_market()

'Declare and set worksheet
Dim ws As Worksheet

'Loop through all stocks for one year
For Each ws In ThisWorkbook.Worksheets

    'Create the column headings
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"

    'Define Ticker variable
    Dim Ticker As String
    Ticker = " "
    
    'Create variable to hold stock volume
    Dim stock_volume As Double
    stock_volume = 0

    'Set initial and last row for worksheet
    Dim Lastrow As Long
    Dim i As Long
    Dim k As Long
    Dim TickerRow As Long
    TickerRow = 2

    'Define Lastrow of worksheet
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Set new variables for prices and percent changes
    Dim open_price As Double
    open_price = ws.Cells(2, 3).Value
    Dim close_price As Double
    close_price = 0
    Dim price_change As Double
    price_change = 0
    Dim price_change_percent As Double
    price_change_percent = 0
    Dim total_volume As Double
    total_volume = 0

    'Do loop of current worksheet to Lastrow
    For i = 2 To Lastrow

        'Ticker symbol output
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            ws.Range("I" & TickerRow).Value = Ticker
        
            'Calculate change in Price
            close_price = ws.Cells(i, 6).Value
            price_change = close_price - open_price
            
            ws.Range("J" & TickerRow).Value = price_change

            If open_price <> 0 Then
                price_change_percent = (price_change / open_price) * 100
            Else
                price_change_percent = 0
            End If

            ws.Range("K" & TickerRow).Value = price_change_percent
            total_volume = total_volume + ws.Cells(i, 7).Value
            ws.Range("L" & TickerRow).Value = total_volume
            
    'Conditional Formatting
    If ws.Range("K" & TickerRow).Value > 0 Then
        ws.Range("K" & TickerRow).Interior.ColorIndex = 4
    ElseIf ws.Range("K" & TickerRow).Value < 0 Then
     ws.Range("K" & TickerRow).Interior.ColorIndex = 3
    Else
        ws.Range("K" & TickerRow).Interior.ColorIndex = 0
    End If
    
    If ws.Range("J" & TickerRow).Value > 0 Then
        ws.Range("J" & TickerRow).Interior.ColorIndex = 4
    ElseIf ws.Range("J" & TickerRow).Value < 0 Then
     ws.Range("J" & TickerRow).Interior.ColorIndex = 3
    Else
        ws.Range("J" & TickerRow).Interior.ColorIndex = 0
    End If
        
        'Reset values for next ticker
        TickerRow = TickerRow + 1
        open_price = ws.Cells(i + 1, 3).Value
        total_volume = ws.Cells(i + 1, 7).Value
        
        'Else is when ticker hasn't changed
        Else
            total_volume = total_volume + ws.Cells(i, 7).Value
        
        End If
    Next i
        
'Greatest Increase Decrease
    For k = 2 To Lastrow
    Dim Rng As Range
    Set Rng = ws.Range("K:K")
    ws.Range("P2").Value = WorksheetFunction.Max(Rng)
            If ws.Cells(k, 11).Value = ws.Cells(2, 16).Value Then
            ws.Cells(2, 15).Value = ws.Cells(k, 9).Value
        End If
        ws.Range("P3").Value = WorksheetFunction.Min(Rng)
        If ws.Cells(k, 11).Value = ws.Cells(3, 16).Value Then
            ws.Cells(3, 15).Value = ws.Cells(k, 9).Value
        End If

'Greatest Volume
    Dim GVol As Range
    Set GVol = ws.Range("L:L")
        ws.Range("P4").Value = WorksheetFunction.Max(GVol)
        If ws.Cells(k, 12).Value = ws.Cells(4, 16).Value Then
            ws.Cells(4, 15).Value = ws.Cells(k, 9).Value
        End If
    Next k

Next ws

End Sub
