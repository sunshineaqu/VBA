Sub TotalStock()
   
    ' LOOP THROUGH ALL SHEETS
    For Each ws In Worksheets

        ' Created a Variable to Hold total volume, ticker, Last Row, and the position for output
        Dim TotalVolume As Double
        Dim Ticker As String
        Dim RowIndex As Integer
        Dim YearOpen as Double
        Dim YearClose as Double
        Dim YearChange as Double
        Dim PercChange as Double

        'set the value to 1st ticker (2nd row), in case the 1st ticker is the only one in ws
        TotalVolume = ws.Cells(2,7).Value
        RowIndex = 2
        YearOpen = ws.Cells(2,3).Value

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        
        ' Looping through the first colume to calculate each ticker
        For i = 2 To LastRow

            'for different ticker, output the value for the previous ticker, and start a new calculation
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then

                YearClose = ws.Cells(i,6).Value
                YearChange = YearClose - YearOpen
                PercChange = YearChange / YearOpen
                ws.Cells(RowIndex, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(RowIndex, 10).Value = YearChange
                ws.Cells(RowIndex, 11).Value = PercChange
                ws.Cells(RowIndex, 12).Value = TotalVolume
                'relocate the row for input, reset the total volume for the new ticker
                RowIndex = RowIndex + 1
                TotalVolume = 0
                YearOpen = 0


            'Define the 1st non 0 value for Year open, calculate the total stock volume for the same ticker
            Else
                TotalVolume = TotalVolume + ws.Cells(i+1, 7).Value
                If YearOpen = 0 Then
                    YearOpen = ws.Cells(i,3).Value
                End If    

            End If

        Next i

        ' format and print the head
        For i = 2 to RowIndex

            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i

        
        ws.Range("K2:K"&RowIndex).Style = "Percent"
        ws.Range("K2:K"&RowIndex).NumberFormat = "0.00%"
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
    Next ws

End Sub



