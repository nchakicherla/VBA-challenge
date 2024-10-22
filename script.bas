Attribute VB_Name = "Module1"
Sub summarizeTradingData()

    Dim sheet As Worksheet
    Dim finalRow As Long
    Dim symbol As String
    Dim firstPrice As Double
    Dim lastPrice As Double
    Dim percentDiff As Double
    Dim totalVolume As Double
    Dim changeValue As Double
    Dim rowIndex As Long
    Dim maxVolumeTicker As String
    Dim maxVolume As Double
    Dim maxPercentGain As Double
    Dim maxGainTicker As String
    Dim maxLossTicker As String
    Dim maxPercentLoss As Double
    
    For Each sheet In ThisWorkbook.Worksheets
    
        finalRow = sheet.Cells(Rows.Count, 1).End(xlUp).Row
        
        sheet.Cells(1, 9).Value = "Ticker"
        sheet.Cells(1, 10).Value = "Quarterly Change"
        sheet.Cells(1, 11).Value = "Percent Change"
        sheet.Cells(1, 12).Value = "Total Stock Volume"
        
        totalVolume = 0
        rowIndex = 2
        
        For i = 2 To finalRow
        
            symbol = sheet.Cells(i, 1).Value
            
            If symbol <> sheet.Cells(i - 1, 1).Value Then
            
                firstPrice = sheet.Cells(i, 3).Value
            
            End If
            
            totalVolume = totalVolume + sheet.Cells(i, 7).Value
            
            If symbol <> sheet.Cells(i + 1, 1).Value Then
                
                lastPrice = sheet.Cells(i, 6).Value
                changeValue = lastPrice - firstPrice
                percentDiff = ((lastPrice - firstPrice) / firstPrice)
                
                sheet.Cells(rowIndex, 9).Value = symbol
                sheet.Cells(rowIndex, 10).Value = changeValue
                sheet.Cells(rowIndex, 11).Value = percentDiff
                sheet.Cells(rowIndex, 12).Value = totalVolume
                
                totalVolume = 0
                
                If changeValue > 0 Then
                    sheet.Cells(rowIndex, 10).Interior.Color = vbGreen
                ElseIf changeValue < 0 Then
                    sheet.Cells(rowIndex, 10).Interior.Color = vbRed
                ElseIf changeValue = 0 Then
                    sheet.Cells(rowIndex, 10).Interior.Color = xlNone
                End If
                
                rowIndex = rowIndex + 1
            End If
        Next i
        
        finalRow = sheet.Cells(Rows.Count, 9).End(xlUp).Row
        
        sheet.Cells(1, 16).Value = "Ticker"
        sheet.Cells(1, 17).Value = "Value"
        sheet.Cells(2, 15).Value = "Greatest % Increase"
        sheet.Cells(3, 15).Value = "Greatest % Decrease"
        sheet.Cells(4, 15).Value = "Greatest Total Volume"
        
        maxPercentGain = 0
        maxPercentLoss = 0
        maxVolume = 0
        
        For i = 2 To finalRow
        
            If sheet.Cells(i, 11).Value > maxPercentGain Then
                maxPercentGain = sheet.Cells(i, 11).Value
                maxGainTicker = sheet.Cells(i, 9).Value
            
            ElseIf sheet.Cells(i, 11).Value < maxPercentLoss Then
                maxPercentLoss = sheet.Cells(i, 11).Value
                maxLossTicker = sheet.Cells(i, 9).Value
                
            End If
            
            If sheet.Cells(i, 12).Value > maxVolume Then
                maxVolume = sheet.Cells(i, 12).Value
                maxVolumeTicker = sheet.Cells(i, 9).Value
                
            End If
        Next i
        
        sheet.Cells(2, 16) = maxGainTicker
        sheet.Cells(2, 17) = maxPercentGain
        sheet.Cells(3, 16) = maxLossTicker
        sheet.Cells(3, 17) = maxPercentLoss
        sheet.Cells(4, 16) = maxVolumeTicker
        sheet.Cells(4, 17) = maxVolume
    Next sheet
End Sub


