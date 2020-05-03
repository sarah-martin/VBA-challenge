Sub VBA_Challenge ()

    'Create the loop between worksheets
    For each ws in worksheets

        'Name the new columns
        ws.Range("I1").Value="Ticker"
        ws.Range("J1").Value="Yearly Change"
        ws.Range("K1").Value="Percent Change"
        ws.Range("L1").Value="Total Stock Volume"

        'Create variables to hold ticker, opening price, closing price, and total volume
        Dim Ticker As String
        Dim TotalVolume as LongLong
        Dim OpenPrice as Double
        Dim ClosePrice as Double

            'Set all number variables to 0
            TotalVolume=0

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

            'Or if the cell below is not equal to the current cell, then
            Elseif ws.Cells (i+1,1).Value <> ws.Cells(i,1).Value Then
                
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
                    Else
                    End If

                'Add one to output table
                output=output+1

                'Reset total stock volume for next ticker
                TotalVolume=0

            Else

                'Or add to the stock volume
                TotalVolume=TotalVolume+ws.Cells(i,7).Value

            End If

            'Modify to percent style in column K
            ws.Cells(i, 11).Style = "Percent"

            'If yearly change is less than 10, then the interior color is red
            If ws.Cells(i,10).Value<0 Then
                ws.Cells(i,10).Interior.ColorIndex=3

            ElseIf ws.Cells(i,10).Value>=0 Then
                ws.Cells(i,10).Interior.ColorIndex=4
                
            Else
            End If

        Next i

    Next ws

End Sub