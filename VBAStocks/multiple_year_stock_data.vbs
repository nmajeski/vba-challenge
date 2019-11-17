Sub MultipleYearStockData():
    For Each ws In Worksheets
        ' Set up headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Point to the beginning of the result table
        Dim CurrentResultRow As Long
        CurrentResultRow = 2
        ws.Cells(CurrentResultRow, 9).Value = ws.Cells(2, 1).Value

        Dim OpeningPrice As Double
        OpeningPrice = ws.Cells(2, 3).Value
        Dim ClosingPrice As Double
        
        ' Initialize variable for the total volume of each ticker
        Dim CurrentTotalVolume As LongLong
        CurrentTotalVolume = 0
        
        For i = 3 To LastRow
            ' Initialize previous and current row numbers
            Dim PreviousRow As Long
            Dim CurrentRow As Long
            PreviousRow = i - 1
            CurrentRow = i
            
            ' Get the previous row's and current row's ticker so that we can compare them later
            Dim PreviousRowTicker As String
            Dim CurrentRowTicker As String
            PreviousRowTicker = ws.Cells(PreviousRow, 1).Value
            CurrentRowTicker = ws.Cells(CurrentRow, 1).Value

            Dim YearlyChange As Double
            Dim PercentChange As Double
            
            ' Initialize Volume for use when calculating the total volume
            Dim Volume As Long
            
            ' We need to handle things differently when the previous and current tickers don't match
            If CurrentRowTicker <> PreviousRowTicker Then
                ' Get the closing price for the previous row's ticker and calculate the changes
                ClosingPrice = ws.Cells(PreviousRow, 6).Value
                YearlyChange = ClosingPrice - OpeningPrice
                ws.Cells(CurrentResultRow, 10).Value = YearlyChange
                If YearlyChange >= 0 Then
                    ws.Cells(CurrentResultRow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(CurrentResultRow, 10).Interior.ColorIndex = 3
                End If

                If OpeningPrice <> 0 Then
                    PercentChange = YearlyChange / OpeningPrice
                    ws.Cells(CurrentResultRow, 11).Value = PercentChange
                    ws.Cells(CurrentResultRow, 11).Style = "Percent"
                End If

                OpeningPrice = ws.Cells(CurrentRow, 3)

                ' Set the total volume cell for the previous ticker
                ws.Cells(CurrentResultRow, 12).Value = CurrentTotalVolume
                Volume = ws.Cells(CurrentRow, 7).Value
                CurrentTotalVolume = Volume
                
                ' Move onto the next ticker
                CurrentResultRow = CurrentResultRow + 1
                ws.Cells(CurrentResultRow, 9).Value = CurrentRowTicker
            Else
                ' Keep summing the volumes for the ticker
                Volume = ws.Cells(CurrentRow, 7).Value
                CurrentTotalVolume = CurrentTotalVolume + Volume
                
                ' If the CurrentRow is the last row, we won't have a next ticker to compare
                ' So, we need to update the price change, percent change, and total volume cells
                If CurrentRow = LastRow Then
                    ' Get the closing price for the previous row's ticker and calculate the changes
                    ClosingPrice = ws.Cells(PreviousRow, 6).Value
                    YearlyChange = ClosingPrice - OpeningPrice
                    ws.Cells(CurrentResultRow, 10).Value = YearlyChange
                    If YearlyChange >= 0 Then
                        ws.Cells(CurrentResultRow, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(CurrentResultRow, 10).Interior.ColorIndex = 3
                    End If

                    If OpeningPrice <> 0 Then
                        PercentChange = YearlyChange / OpeningPrice
                        ws.Cells(CurrentResultRow, 11).Value = PercentChange
                        ws.Cells(CurrentResultRow, 11).Style = "Percent"
                    End If

                    ws.Cells(CurrentResultRow, 12).Value = CurrentTotalVolume
                End If
            End If
            
        Next i

        ws.Columns("A:L").AutoFit
    Next
End Sub

