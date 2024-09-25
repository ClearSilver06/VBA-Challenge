Sub module2codes()

    ' Define variables
    Dim i As Long
    Dim ticker As String

    ' Define Worksheet Variables
    Dim ws As Worksheet

    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets

        Dim quarterly_change As Double
        quarterly_change = 0

        Dim open_price As Double
        open_price = 0

        Dim close_price As Double
        close_price = 0

        Dim total_stock_value As Double
        total_stock_value = 0

        ' Location of summary table row column (J)
        Dim summary_table_row As Long
        summary_table_row = 2

        ' Define dynamic last row
        Dim lastrow As Long
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Name columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Loop through tickers
        For i = 2 To lastrow

            ' First row or ticker change
            If i = 2 Or ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

                ' If not first ticker, calculate quarterly change
                If i > 2 Then
                    quarterly_change = close_price - open_price
                    
                    ' Print total stock volume and quarterly change
                    ws.Range("I" & summary_table_row).Value = ticker
                    ws.Range("L" & summary_table_row).Value = total_stock_value
                    ws.Range("J" & summary_table_row).Value = quarterly_change
                    

                    ' Calculate percent change 
                    If open_price <> 0 Then
                        ws.Range("K" & summary_table_row).Value = quarterly_change / open_price
                        ws.Range("K" & summary_table_row).NumberFormat = "0.00%" ' Convert to percentage
                    Else
                        ws.Range("K" & summary_table_row).Value = 0
                        ws.Range("K" & summary_table_row).NumberFormat = "0.00%" ' Convert to percentage
                    End If
                    

                    ' Add summary row
                    summary_table_row = summary_table_row + 1

                End If

                ' Set ticker name
                ticker = ws.Cells(i, 1).Value

                ' Reset total stock volume
                total_stock_value = ws.Cells(i, 7).Value
                
                ' Set open price for new ticker
                open_price = ws.Cells(i, 3).Value

            End If

            ' Update close price and total stock value
            close_price = ws.Cells(i, 6).Value
            total_stock_value = total_stock_value + ws.Cells(i, 7).Value
            
        Next i

        ' Final calculation for last ticker
        If open_price <> 0 Then
            quarterly_change = close_price - open_price
            ws.Range("I" & summary_table_row).Value = ticker
            ws.Range("L" & summary_table_row).Value = total_stock_value
            ws.Range("J" & summary_table_row).Value = quarterly_change
                    
            ' Calculate percent change
            If open_price <> 0 Then
                ws.Range("K" & summary_table_row).Value = quarterly_change / open_price
                ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
            Else
                ws.Range("K" & summary_table_row).Value = 0
                ws.Range("K" & summary_table_row).NumberFormat = "0.00%" 
            End If


        End If
        
        ' Color green for positive, Red for negative
        For j = 2 To summary_table_row 
            If ws.Range("J" & j).Value > 0 Then
                ws.Range("J" & j).Interior.ColorIndex = 4 ' Green 
            ElseIf ws.Range("J" & j).Value < 0 Then
                ws.Range("J" & j).Interior.ColorIndex = 3 ' Red
            End If


        Next j

    Next ws

End Sub
