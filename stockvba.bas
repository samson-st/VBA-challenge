Attribute VB_Name = "Module1"

Sub stockx()

Dim tickers As String
Dim tickerVolume As Double
Dim tickerStartPrice As Double
Dim tickerEndPrice As Double
Dim tickerPriceChange As Double
Dim PercentChange As Double
Dim tickercounter As Double
Dim i As Double
Dim ws As Worksheet
Dim lastrow As Double

For Each ws In Worksheets
ws.Activate
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Range("I1").EntireColumn.Insert
ws.Cells(1, 9) = "tickers"
ws.Range("J1").EntireColumn.Insert
ws.Cells(1, 10) = "tickerPriceChange"
ws.Range("K1").EntireColumn.Insert
ws.Cells(1, 11) = "PercentChange"
ws.Range("L1").EntireColumn.Insert
ws.Cells(1, 12) = "tickerVolume"

    'scan across rows
    For i = 2 To lastrow
        tickers = Cells(i, 1).Value
            If tickerStartPrice = 0 Then
            tickerStartPrice = Cells(i, 3).Value
            End If
        
        If Cells(i + 1, 1).Value <> tickers Then
    
            'find start of ticker
            tickercounter = tickercounter + 1
            Cells(tickercounter + 1, 9) = tickers
                  
        
            'calculate ticker price
            tickerEndPrice = Cells(i, 6)
            tickerPriceChange = tickerEndPrice - tickerStartPrice
            Cells(tickercounter + 1, 10).Value = tickerPriceChange
        
        
            'calculate ticker volume
            tickerVolume = tickerVolume
            Cells(tickercounter + 1, 12).Value = tickerVolume
        
            'find percent change over year
            PercentChange = (tickerPriceChange / tickerStartPrice)
            Cells(tickercounter + 1, 11).Value = Format(PercentChange, "Percent")
        
        tickerVolume = 0
        tickerStartPrice = 0
        
        Else
        
        tickerVolume = tickerVolume + Cells(i, 7).Value
    
        End If
     
            If Cells(i, 10) > 0 Then
                Cells(i, 10).Interior.ColorIndex = 3
        
            ElseIf Cells(i, 10) < 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            
            Else
                Cells(i, 10).Interior.ColorIndex = 0
        End If
    Next i
Next ws
End Sub
