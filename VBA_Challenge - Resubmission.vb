Sub VBA_Challenge():

    'Create the loop between worksheets
    For Each ws in worksheets

        'Name the new columns and format sizing
        ws.Range("I1").Value="Ticker"
        ws.Range("J1").Value="Yearly Change"
        ws.Range("K1").Value="Percent Change"
        ws.Range("L1").Value="Total Stock Volume"
        ws.Range("J:L").ColumnWidth = 17

        'Create variables to hold ticker, opening price, closing price, and total volume
        Dim Ticker As String
        Dim TotalVolume as Long 
        Dim OpenPrice as Double
        Dim ClosePrice as Double
        Dim i as Long
        Dim LastRow as Long

        'Create a holder for the new outputs
        Dim Output as Double
        Output=2

        'Find the last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through ticker information
        For i=2 to LastRow

            'If the cell above is not equal to the current cell, then
            If ws.Cells (i-1,1).Value <> ws.Cells(i,1).Value Then

                'Establish the opening price
                OpenPrice=ws.Cells(i,3).Value

            End if

            'Or if the cell below is not equal to the current cell, then
            If ws.Cells (i+1,1).Value <> ws.Cells(i,1).Value Then
                
                'Establish the current ticker
                Ticker=ws.Cells(i,1).Value

                'Establish closing price
                ClosePrice=ws.Cells(i,6).Value

                'Add up the total stock volume
                TotalVolume=TotalVolume+ws.Cells(i,7).Value

                'Place current ticker in output table
                ws.Range("I"&Output).Value=Ticker

                'Place total stock volume in output table
                ws.Range("L"&Output).Value=TotalVolume

                'Place yearly change in output table, using opening and closing prices
                ws.Range("J"&Output).Value=ClosePrice-OpenPrice

                    If OpenPrice <>0 Then 
                    'Pleace percent change in output table, using opening and closing prices
                    ws.Range("K"&Output).Value=((ClosePrice-OpenPrice)/OpenPrice)
                    ws.Range("K:K").Numberformat="0.00%"
                    Else
                    End If

                'Add one to output table
                output=output+1

                'Reset total stock volume amd prices for next ticker
                TotalVolume=0

            Else

                'Or add to the stock volume
                TotalVolume=TotalVolume+ws.Cells(i,7).Value

            End If

            'If yearly change is less than 0, then the interior color is red
            If ws.Range("J"&Output).Value<0 Then
                ws.Range("J"&Output).Interior.ColorIndex=3

            'Or if the value is greater than 0, make it green
            ElseIf ws.Range("J"&Output).Value>=0 Then
                ws.Range("J"&Output).Interior.ColorIndex=4
                
            End If

        Next i

    Next ws

End Sub