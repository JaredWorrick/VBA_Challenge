Sub Multiple_year_stock_data():
    For Each ws In Worksheets
        Dim OWS As String
        Dim i As Long
        Dim j As Long
        Dim LatestBlankRow As Long
        Dim LastRow As Long
        Dim PercentChange As Double
        
        j = 2
        LatestBlankRow = 2
        TotalStockVolume = 0
        OWS = ws.Name
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
            For i = 2 To LastRow
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                    ' Ticker
                    ws.Range("i1").Value = "Ticker"
                    ws.Cells(LatestBlankRow, 9).Value = ws.Cells(i, 1).Value
            
                    ' Yearly Change
                    ws.Range("j1").Value = "Yearly Change"
                    ws.Cells(LatestBlankRow, 10) = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                        If ws.Cells(LatestBlankRow, 10).Value < 0 Then
                        ws.Cells(LatestBlankRow, 10).Interior.ColorIndex = 3
                        Else
                        ws.Cells(LatestBlankRow, 10).Interior.ColorIndex = 4
                        End If
        
                        ' Percent Change w/ Percantage Format
                        ws.Range("k1").Value = "Percent Change"
                        PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                        If ws.Cells(j, 3).Value <> 0 Then
                        ws.Cells(LatestBlankRow, 11).Value = Format(PercentChange, "Percent")
                        Else
                        ws.Cells(LatestBlankRow, 11).Value = Format(0, "Percent")
                        End If

                    ' Total Stock Volume
                    ws.Range("l1").Value = "Total Stock Volume"
                    ws.Cells(LatestBlankRow - 1, 12) = TotalStockVolume
                    TotalStockVolume = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                    LatestBlankRow = LatestBlankRow + 1
                    j = i + 1
            
                End If
            
            Next i
        
        Dim NewLastRow As Long
        Dim GI As Double
        Dim GD As Double
        Dim GTV As Double

        GI = ws.Cells(2, 11).Value
        GD = ws.Cells(2, 11).Value
        GTV = ws.Cells(2, 12).Value

        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"

        TickerLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

            For i = 2 To TickerLastRow

                'Greatest Percent Increase
                ws.Range("o2").Value = "Greatest % Increase"

                'Greatest Percent Decrease
                ws.Range("o3").Value = "Greatest % Decrease"

                'Greatest Total Volume
                ws.Range("o4").Value = "Greatest Total Volume"
            Next i
    Next ws
End Sub