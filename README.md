'READ ME

Sub multiple_year_stock_data()
    'Dim ws As Worksheet
    Dim lastRow As Double
    Dim i As Long
    Dim total_volume As Double
    total_volume = 0
    
    'Extract stock data (ticker symbol, opening price, closing price, volume)
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim yeardate As Double
    Dim Stock_info As Double
   Stock_info = 2
    ' Set the worksheet containing the stock data
    For Each ws In Worksheets
        'Set ws = ThisWorkbook.Sheets("2018")
        ' Find the last row with data in column A
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' Loop through the rows to extract stock data
        For i = 2 To lastRow ' Assuming data starts from row 2
            'create loop to cycle through tickers and then create function to select the close value for the last ticker
            'ticker = ws.Cells(i, 1).Value ' Assuming ticker symbol is in column A
            '  openPrice = ws.Cells(i, 3).Value ' Assuming opening price is in column B
            ' closePrice = ws.Cells(i, 6).Value ' Assuming closing price is in column C
            'volume = ws.Cells(i, 7).Value ' Assuming volume is in column G
            'yeardate = ws.Cells(i, 2).Value ' DAte is in column 2
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ticker = Cells(i, 1).Value
                ws.Cells(Stock_info, "I") = ticker
                'Add to the volume total
                  total_volume = total_volume + ws.Cells(i, 7).Value
                'Print ticker symbol in ticker column
                ' Range("I" & Stock_info).Value = ticker
                'Print Volume in volume column
                'Range("j" & Stock_info).Value = total_volume
                'Reset volume total
                volume_total = 0
                'print total f
                ws.Cells(Stock_info, "L").Value = total_volume

                'Else
                '   volume_total = Cells(i, 7).Value
                Stock_info = Stock_info + 1
                    
                Else
                ws.Cells(Stock_info, "L") = total_volume
            End If
        Next i
    Next ws
End Sub

'Reference website: https://stackoverflow.com/questions/78156319/how-to-find-the-total-volume-of-a-stock-and-then-calculate-the-yearly-chang

