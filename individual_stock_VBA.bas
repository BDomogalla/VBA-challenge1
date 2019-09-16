
Sub stockloop()

    Dim ws as Worksheet
    Dim summary_table_row as Long
    summary_table_row = 2
    Dim number_of_rows as Long
    Dim volume_count as Double
    Dim open_price as Double
    Dim close_price as Double 
    Dim yearly_change as Double
    Dim percent_change as Double
    Dim first_row as Long
    first_row = 0 
    Dim i as long

    For each ws in Worksheets
        ws.Cells(1,9).Value = "Ticker"
        ws.Cells(1,10).Value = "Yearly Change"
        ws.Cells(1,11).Value = "Percent Change"
        ws.Cells(1,12).Value = "Total Stock Volume"
        number_of_rows = ws.Cells(Rows.Count, 1).End(xlUp).Row


        For i = 2 to number_of_rows
            If ws.Cells(i,1).Value = ws.Cells(i+1,1).Value Then   
                first_row = first_row +1
                volume_count = volume_count + ws.Cells(i,7).Value
                If first_row = 1 Then
                open_price = ws.Cells(i,3)
                End If
            Else 
                volume_count = volume_count + ws.Cells(i,7).Value
                ws.Cells(summary_table_row,9).Value = ws.Cells(i,1).Value
                ws.Cells(summary_table_row,12) = volume_count
                close_price = ws.Cells(i,6).Value
                IF open_price <> 0 Then
                    yearly_change = close_price - open_price
                    percent_change = ((close_price - open_price)/open_price)
                Else
                    yearly_change = 0
                    percent_change = 0
                End If

                ws.Cells(summary_table_row,10).Value = yearly_change
                ws.Cells(summary_table_row,11).Value = percent_change
                ws.cells(summary_table_row,11).NumberFormat = "0.00%"

                    IF ws.Cells(summary_table_row,10).Value > 0 Then
                        ws.Cells(summary_table_row,10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(summary_table_row,10).Interior.ColorIndex = 3
                    End IF

                volume_count = 0
                summary_table_row = summary_table_row + 1
                first_row = 0
            End If
        Next i
    summary_table_row = 2
    Next ws
End Sub