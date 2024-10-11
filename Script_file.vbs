Attribute VB_Name = "Module1"
Sub stock_loop()

' naming my variables
Dim I As Long
Dim ws As Worksheet
Dim tik_nm As String
Dim begin As Double

'Dim as Long because these values will be whole numbers
Dim lastRow As Long
Dim out_row As Long

'Dim as doubl for precise decimal
Dim start_qtr As Double
Dim end_qtr As Double
Dim qtr_total As Double

Dim Percent_change As Double
Dim stock_volume As Double

Dim greatest_percent As Double
Dim big_percent_tk As String
Dim smallest_percent As Double
Dim small_percent_tk As String
Dim greatest_stock As Double
Dim greatest_stock_tk As String



' Looping through each worksheet in the Workbook
For Each ws In ThisWorkbook.Worksheets

    ' initialize counters
    out_row = 2
    qtr_total = 0
    start_qtr = 0
    end_qtr = 0
    stock_volume = 0
    
    greatest_percent = -500
    smallest_percent = 500
    
    ' Find the last row of column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
        ' Loop through each row
        For I = 2 To lastRow
            
            'Capturing the open price of each ticker
            If ws.Cells(I - 1, 1).Value <> ws.Cells(I, 1).Value Then
            
                start_qtr = ws.Cells(I, 3).Value
            
            End If
    
            'For each ticker
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            
                ' capture the last value in column 6 for the each ticker
                end_qtr = ws.Cells(I, 6).Value
                
                ' calculating the quarter change
                qtr_total = end_qtr - start_qtr
                
                ' Formatting the colors of the quater change
                If qtr_total > 0 Then
                    'if qtr_total is positive fill cells green
                    ws.Cells(out_row, "J").Interior.ColorIndex = 4
                Else
                    'if qtr_total is negative fill cells red
                    ws.Cells(out_row, "J").Interior.ColorIndex = 3
                    
                End If
                
                If qtr_total = 0 Then
                    ws.Cells(out_row, "J").Interior.ColorIndex = 2
                End If
                
                
                ' calculate the stock volume
                stock_volume = stock_volume + ws.Cells(I, 7).Value
                
                
                'Calculating the percent change
                If start_qtr <> 0 Then
                    Percent_change = (qtr_total / start_qtr)
                Else
                    Percent_change = 0
                End If
                
                'Get the current ticker name from column A
                tik_nm = ws.Cells(I, 1).Value
                
                'capturing the biggest percent change and associated ticker
                If Percent_change > greatest_percent Then
                    greatest_percent = Percent_change
                    big_percent_tk = ws.Cells(I, 1).Value
                End If
                
                'capture biggest percent decrease and associated ticker
                If Percent_change < smallest_percent Then
                    smallest_percent = Percent_change
                    small_percent_tk = ws.Cells(I, 1).Value
                End If
                
                'caputure greatest stock volume
                If stock_volume > greatest_stock Then
                    greatest_stock = stock_volume
                    greatest_stock_tk = ws.Cells(I, 1).Value
                End If
                
                'print the ticker name to column I
                ws.Range("i" & out_row).Value = tik_nm
                
                'print quarter total
                ws.Range("j" & out_row).Value = qtr_total
                
                'print percent change
                ws.Range("k" & out_row).Value = Percent_change
                
                'print stock volume
                ws.Range("L" & out_row).Value = stock_volume
                
                'print the greatest percentage increase
                ws.Range("P" & 2).Value = greatest_percent
                
                'print the greatest percentage decrease
                ws.Range("P" & 3).Value = smallest_percent
                
                'print greatest total stock volume
                ws.Range("P" & 4).Value = greatest_stock
                
                'print ticker names
                ws.Range("O" & 2).Value = big_percent_tk
                ws.Range("O" & 3).Value = small_percent_tk
                ws.Range("O" & 4).Value = greatest_stock_tk
                
                
                'Print the summary table lables
                ws.Range("O" & 1).Value = "Ticker"
                ws.Range("P" & 1).Value = "Value"
                ws.Range("N" & 2).Value = "Greatest % Increase"
                ws.Range("N" & 3).Value = "Greatest % Decrease"
                ws.Range("N" & 4).Value = "Greatest Total Volume"
                ws.Range("I" & 1).Value = "Ticker"
                ws.Range("J" & 1).Value = "Quarterly Change"
                ws.Range("K" & 1).Value = "Percent Change"
                ws.Range("L" & 1).Value = "Total Stock Volume"

                
                'reset stock volume counter
                stock_volume = 0
                                
                out_row = out_row + 1
                
                ws.Columns("K:K").NumberFormat = "0.00%"
                ws.Columns("I:R").AutoFit

            Else
                'compile the stock_volume values to get the total
                stock_volume = stock_volume + ws.Cells(I, 7).Value
                
            End If
            
        Next I
    
Next ws

End Sub

