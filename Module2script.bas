Attribute VB_Name = "Module1"
Sub AnalyzeStockData()
    'Create variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim quarterStartDate As Date
    Dim outputRow As Long
    
    ' Create variables for tracking greatest values
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerIncrease As String
    Dim tickerDecrease As String
    Dim tickerVolume As String
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row with data for ticker
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Set up the headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        outputRow = 2
        ticker = ""
        
        ' Initialize greatest values
        greatestIncrease = -1
        greatestDecrease = 1
        greatestVolume = 0
        
        ' Loop through all rows
        For i = 2 To lastRow
            ' Check if this is the start of a new ticker or quarter
            If ws.Cells(i, 1).Value <> ticker Or _
               (ticker <> "" And DateDiff("q", quarterStartDate, ws.Cells(i, 2).Value) <> 0) Then
                
                ' Output the previous ticker's data (if not the first run)
                If ticker <> "" Then
                    ws.Cells(outputRow, 9).Value = ticker
                    ws.Cells(outputRow, 10).Value = closePrice - openPrice
                    ws.Cells(outputRow, 11).Value = (closePrice - openPrice) / openPrice
                    ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                    ws.Cells(outputRow, 12).Value = volume
                    
                    ' Check for greatest values
                    If ws.Cells(outputRow, 11).Value > greatestIncrease Then
                        greatestIncrease = ws.Cells(outputRow, 11).Value
                        tickerIncrease = ticker
                    End If
                    If ws.Cells(outputRow, 11).Value < greatestDecrease Then
                        greatestDecrease = ws.Cells(outputRow, 11).Value
                        tickerDecrease = ticker
                    End If
                    If volume > greatestVolume Then
                        greatestVolume = volume
                        tickerVolume = ticker
                    End If
                    
                    outputRow = outputRow + 1
                End If
                
                ' Reset variables for the new ticker/quarter
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                volume = 0
                quarterStartDate = DateSerial(Year(ws.Cells(i, 2).Value), (DatePart("q", ws.Cells(i, 2).Value) - 1) * 3 + 1, 1)
            End If
            
            ' Update the close price and add to the volume
            closePrice = ws.Cells(i, 6).Value
            volume = volume + ws.Cells(i, 7).Value
        Next i
        
        ' Output the last ticker's data
        If ticker <> "" Then
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = closePrice - openPrice
            ws.Cells(outputRow, 11).Value = (closePrice - openPrice) / openPrice
            ws.Cells(outputRow, 11).NumberFormat = "0.00%"
            ws.Cells(outputRow, 12).Value = volume
            
            ' Check for greatest values again
            If ws.Cells(outputRow, 11).Value > greatestIncrease Then
                greatestIncrease = ws.Cells(outputRow, 11).Value
                tickerIncrease = ticker
            End If
            If ws.Cells(outputRow, 11).Value < greatestDecrease Then
                greatestDecrease = ws.Cells(outputRow, 11).Value
                tickerDecrease = ticker
            End If
            If volume > greatestVolume Then
                greatestVolume = volume
                tickerVolume = ticker
            End If
        End If
        
        ' Apply conditional formatting to the Quarterly Change column
        ws.Range("J2:J" & outputRow).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
        ws.Range("J2:J" & outputRow).FormatConditions(1).Interior.Color = RGB(0, 255, 0)
        ws.Range("J2:J" & outputRow).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        ws.Range("J2:J" & outputRow).FormatConditions(2).Interior.Color = RGB(255, 0, 0)
        
        ' Output the greatest values
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 16).Value = tickerIncrease
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        ws.Cells(3, 16).Value = tickerDecrease
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 16).Value = tickerVolume
        ws.Cells(4, 17).Value = greatestVolume
        
        ' Autofit the columns
        ws.Columns("I:Q").AutoFit
    Next ws
End Sub

