Attribute VB_Name = "Module1"
Sub outputInfo()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
    Dim i As Long
    Dim currentTicker As String
    Dim openPrice As Single
    Dim closePrice As Single
    Dim totalStockVol As Single
    Dim tickerCount As Integer
    Dim maxIncreaseTicker As String
    Dim maxIncrease As Single
    Dim maxDecreaseTicker As String
    Dim maxDecrease As Single
    Dim maxTotalVolTicker As String
    Dim maxTotalVol As Single
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    currentTicker = Cells(2, 1).Value
    openPrice = Cells(2, 3).Value
    totalStockVol = 0
    tickerCount = 2
    maxIncrease = 0
    maxDecrease = 0
    maxTotalVol = 0
    
    For i = 2 To 753001
        If Cells(i, 1).Value = currentTicker Then
            totalStockVol = totalStockVol + Cells(i, 7).Value
            closePrice = Cells(i, 6).Value
        Else
            Cells(tickerCount, 9).Value = currentTicker
            Cells(tickerCount, 10).Value = closePrice - openPrice
            Cells(tickerCount, 11).Value = (closePrice / openPrice) - 1
            Cells(tickerCount, 12).Value = totalStockVol
            If Cells(tickerCount, 11).Value > maxIncrease Then
                maxIncrease = Cells(tickerCount, 11).Value
                maxIncreaseTicker = currentTicker
            End If
            If Cells(tickerCount, 11).Value < maxDecrease Then
                maxDecrease = Cells(tickerCount, 11).Value
                maxDecreaseTicker = currentTicker
            End If
            If totalStockVol > maxTotalVol Then
                maxTotalVol = totalStockVol
                maxTotalVolTicker = currentTicker
            End If
            currentTicker = Cells(i, 1).Value
            openPrice = Cells(i, 3).Value
            closePrice = Cells(i, 6).Value
            totalStockVol = Cells(i, 7).Value
            tickerCount = tickerCount + 1
        End If
    Next i
    
    Cells(2, 16).Value = maxIncreaseTicker
    Cells(2, 17).Value = maxIncrease
    Cells(3, 16).Value = maxDecreaseTicker
    Cells(3, 17).Value = maxDecrease
    Cells(4, 16).Value = maxTotalVolTicker
    Cells(4, 17).Value = maxTotalVol

Next ws

End Sub
