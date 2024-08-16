Attribute VB_Name = "Module1"
Sub QuarterlyStockAnalysis_AllSheets()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim i As Long
    Dim outputRow As Long
    
    ' Loop through each sheet in the array
    For Each ws In ThisWorkbook.Sheets(sheetNames)
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        outputRow = 2 ' Reset output row for each sheet
        
        ' Variables to track greatest increase, decrease, and total volume for each sheet individually
        Dim greatestIncreaseTicker As String
        Dim greatestDecreaseTicker As String
        Dim greatestVolumeTicker As String
        Dim maxIncrease As Double
        Dim maxDecrease As Double
        Dim maxVolume As Double
        
        ' Initialize tracking variables
        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0
        
        ' Add titles in row 1 for new columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Volume"
        ws.Cells(1, 15).Value = "Metric"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' loop through each row of data in the current sheet
        For i = 2 To lastRow
            ' identify ticker transitions
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' it's a new ticker
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                totalVolume = 0
            End If
            
            ' sum of volumes for opening to close of current ticker
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' indentify the end of the quarter (if the next row is a new ticker or last row)
            If i = lastRow Or ws.Cells(i + 1, 1).Value <> ticker Then
                ' end of the quarter close price
                closePrice = ws.Cells(i, 6).Value
                
                ' quarterly change and percentage change
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentageChange = (quarterlyChange / openPrice) * 100
                Else ' prevents divide by 0 error
                    percentageChange = 0
                End If
                
                ' Output the results
                ws.Cells(outputRow, 9).Value = ticker ' Column I
                ws.Cells(outputRow, 10).Value = quarterlyChange ' Column J
                ws.Cells(outputRow, 11).Value = percentageChange ' Column K
                ws.Cells(outputRow, 12).Value = totalVolume ' Column L
                
                ' Conditional formatting - Column J
                If quarterlyChange > 0 Then
                    ws.Cells(outputRow, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf quarterlyChange < 0 Then
                    ws.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
                ' Update greatest increase, decrease, and volume trackers for the current sheet
                If percentageChange > maxIncrease Then
                    maxIncrease = percentageChange
                    greatestIncreaseTicker = ticker
                End If
                If percentageChange < maxDecrease Then
                    maxDecrease = percentageChange
                    greatestDecreaseTicker = ticker
                End If
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    greatestVolumeTicker = ticker
                End If
                
                ' Move to the next output row
                outputRow = outputRow + 1
            End If
        Next i
        
        ' Display greatest % increase, % decrease, and total volume for the current sheet
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = maxIncrease
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = maxDecrease
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(4, 17).Value = maxVolume
    Next ws
    

End Sub

