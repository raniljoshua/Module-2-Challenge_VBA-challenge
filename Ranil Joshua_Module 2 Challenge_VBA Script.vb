'Ranil Joshua
'Module 2 Challenge

Sub Stocks()
    For Each ws In Worksheets
        Worksheets(ws.Name).Activate
        
        'Ticker = Loop for every different <Ticker> until end of column 1
        'Yearly Change = <Closing Price at End of Year> - <Opening Price at Start of Year>
        'Percent Change = -(1-(<Closing Price at End of Year> / <Opening Price at Start of Year>))
        'Total Stock Volume = SUM(<Volume at Start of Year> : <Volume at End of Year>)
        'Greatest Percent Increase = MAX(<Percent Change>), and corresponding <Ticker>
        'Greatest Percent Decrease = MIN(<Percent Change>), and corresponding <Ticker>
        'Greatest Total Volume = MAX(<Total Stock Volume>), and corresponding <Ticker>
        
        Range("I1") = "Ticker"
        Range("J1") = "Yearly Change"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Stock Volume"
        
        Range("O2") = "Greatest % Increase"
        Range("O3") = "Greatest % Decrease"
        Range("O4") = "Greatest Total Volume"
        Range("P1") = "Ticker"
        Range("Q1") = "Value"
        
        
        'Following solution for automatically finding last Row with data was found here:
        'https://www.wallstreetmojo.com/vba-last-row/
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox lastRow
        
        Cells(2, 9) = Cells(2, 1).Value
        CurrentTickerRow = 3
        For Row = 3 To lastRow
            If Cells(Row, 1).Value <> Cells(Row - 1, 1).Value Then
                Cells(CurrentTickerRow, 9) = Cells(Row, 1).Value
                CurrentTickerRow = CurrentTickerRow + 1
            End If
        Next Row
        
        'Col 10 = Yearly Change
        'Col 11 = Percent Change
        'Col 12 = Total Stock Volume
        
        CurrentTickerRow = 2
        Dim currentTickerOpenPrice As Double
        Dim currentTickerClosePrice As Double
        'Dim currentTickerTotalVol As Long
        
        currentTickerOpenPrice = Cells(2, 3).Value
        currentTickerTotalVol = Cells(2, 7).Value
        'MsgBox currentTickerTotalVol
        For Row = 3 To lastRow
            'If Current Ticker is last row
            If Row = lastRow Then
                'MsgBox ("Final row")
                currentTickerClosePrice = Cells(Row, 6).Value
                'Yearly Change = <Closing Price at End of Year> - <Opening Price at Start of Year>
                Cells(CurrentTickerRow, 10) = currentTickerClosePrice - currentTickerOpenPrice
                'Percent Change = -(1-(<Closing Price at End of Year> / <Opening Price at Start of Year>))
                Cells(CurrentTickerRow, 11) = -(1 - (currentTickerClosePrice / currentTickerOpenPrice))
                Cells(CurrentTickerRow, 11).Value = FormatPercent(Cells(CurrentTickerRow, 11))
                'Total Stock Volume = SUM(<Volume at Start of Year> : <Volume at End of Year>)
                currentTickerTotalVol = currentTickerTotalVol + Cells(Row, 7).Value
                Cells(CurrentTickerRow, 12) = currentTickerTotalVol
                'Set currentTickerOpenPrice for new Ticker
                currentTickerOpenPrice = Cells(Row, 3).Value
            'If Current Ticker is the same as the previous
            ElseIf Cells(Row, 1).Value = Cells(Row - 1, 1).Value Then
                'MsgBox (Cells(Row, 7).Value)
                currentTickerTotalVol = currentTickerTotalVol + Cells(Row, 7).Value
            'If Current Ticker is different than previous
            ElseIf Cells(Row, 1).Value <> Cells(Row - 1, 1).Value Then
                currentTickerClosePrice = Cells(Row - 1, 6).Value
                'Yearly Change = <Closing Price at End of Year> - <Opening Price at Start of Year>
                Cells(CurrentTickerRow, 10) = currentTickerClosePrice - currentTickerOpenPrice
                'Percent Change = -(1-(<Closing Price at End of Year> / <Opening Price at Start of Year>))
                Cells(CurrentTickerRow, 11) = -(1 - (currentTickerClosePrice / currentTickerOpenPrice))
                Cells(CurrentTickerRow, 11).Value = FormatPercent(Cells(CurrentTickerRow, 11))
                'Total Stock Volume = SUM(<Volume at Start of Year> : <Volume at End of Year>)
                Cells(CurrentTickerRow, 12) = currentTickerTotalVol
                'Set currentTickerOpenPrice for new Ticker
                currentTickerOpenPrice = Cells(Row, 3).Value
                'Reset currentTickerTotalVol to current Row
                currentTickerTotalVol = Cells(Row, 7).Value
                'Increment currentTickerRow
                CurrentTickerRow = CurrentTickerRow + 1
                'MsgBox currentTickerRow
            End If
        Next Row
        
        For Row = 2 To CurrentTickerRow
            If Cells(Row, 10).Value > 0 Then
                Cells(Row, 10).Interior.ColorIndex = 4
            Else
                Cells(Row, 10).Interior.ColorIndex = 3
            End If
        Next Row
        
        'Calculate Greatest % Increase, Greatest % Decrease, Greatest Total Volume
        greatestPercInc = WorksheetFunction.Max(Range("K:K"))
        'MsgBox greatestInc
        greatestPercDec = WorksheetFunction.Min(Range("K:K"))
        greatestTotalVol = WorksheetFunction.Max(Range("L:L"))
        
        
        CurrentTickerRow = WorksheetFunction.Match(greatestPercInc, Range("K:K"), 0)
        'MsgBox currentTickerRow
        Range("P2") = Cells(CurrentTickerRow, 9)
        Range("Q2") = greatestPercInc
        Range("Q2").Value = FormatPercent(Range("Q2"))
        
        CurrentTickerRow = WorksheetFunction.Match(greatestPercDec, Range("K:K"), 0)
        Range("P3") = Cells(CurrentTickerRow, 9)
        Range("Q3") = greatestPercDec
        Range("Q3").Value = FormatPercent(Range("Q3"))
        
        CurrentTickerRow = WorksheetFunction.Match(greatestTotalVol, Range("L:L"), 0)
        Range("P4") = Cells(CurrentTickerRow, 9)
        Range("Q4") = greatestTotalVol
        
        Columns("I:Q").AutoFit
    Next ws
End Sub

