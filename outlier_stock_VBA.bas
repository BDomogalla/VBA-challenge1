
Sub StockOutliers()
    Dim ws as Worksheet
    Dim i as Long
    Dim j as Long
    Dim high_num as Double
    Dim low_num as Double
    Dim row_number as Long
    Dim high_ticker As String
    Dim low_ticker as String
    Dim outlier_table_row as Integer
    

    For Each ws In Worksheets
        ws.Cells(1,15).Value = "Ticker"
        ws.Cells(1,16).Value = "Value"
        ws.Cells(2,14).Value = "Greatest % Increase"
        ws.Cells(3,14).Value = "Greatest % Decrease"
        ws.Cells(4,14).Value = "Greatest Total Volume"
        ws.Cells(5,14).Value = "Least Total Volume"

        high_num = 0
        low_num = 0
        outlier_table_row = 0

        row_number = ws.Cells(Rows.Count, 9).End(xlUp).Row


        For j = 11 to 12
            For i = 2 to row_number
                If ws.Cells(i,j).Value > high_num Then
                    high_num = ws.Cells(i,j).Value
                    high_ticker = ws.Cells(i,9).Value
                End If

                If ws.Cells(i,j).Value < low_num Then
                    low_num = ws.Cells(i,j).Value
                    low_ticker = ws.Cells(i,9).Value
                End If
            Next i
            ws.Cells(outlier_table_row + 2,15).Value = high_ticker
            ws.Cells(outlier_table_row + 2, 16).Value = high_num
            ws.Cells(outlier_table_row + 3, 15).Value = low_ticker
            Ws.Cells(outlier_table_row + 3,16).Value = low_num

            high_ticker = ""
            low_ticker = ""
            high_num = 0
            low_num = 0

            outlier_table_row = outlier_table_row + 2
        next j
    ws.Range("P2:P3").NumberFormat = "0.00%"
    ws.Range("N5:P5").Value = ""
    next ws
End Sub
        


