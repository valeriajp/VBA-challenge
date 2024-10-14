Sub Stock_Analysis()
  ' Create a variable to hold the counter
  Dim i As Integer
  Dim w As Worksheet
 
  ' Set an initial variable for holding the stock name, open price, close price, quarterly change and percentage change
  Dim Stock_Name As String
  Dim openPrice As Double
  Dim closePrice As Double
  Dim QuarterlyChange As Double
  Dim Percentagechange As Double

  ' Set an initial variable for holding the total per type of stock
  Dim Stock_Total As Double
 
  ' Keep track of the location for each credit card brand in the summary table
  Dim Summary_Table_Row As Integer
  
'variables for bonus
    Dim greatestIncrease As Double
    Dim greatesincreasestockname As String
    Dim greatestDecrease As Double
    Dim greatestdecreseestockname As String
    Dim greatestStockTotal As Double
    Dim greateststockname As String
    Dim lastrow As Long
    
     ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets

            'counts the number of rows
            Stock_Total = 0
            Summary_Table_Row = 2
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            openPrice = ws.Cells(2, 3).Value
            closePrice = 0
            greatestIncrease = 0
            greatestDecrease = 0
            greatestStockTotal = 0

       'Add Headers for Results
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Bonus headers
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
    
  ' Loop through all stock types
  For i = 2 To lastrow
  
    ' Check if we are still within the same stock name, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
      ' Set the Stock name
      Stock_Name = ws.Cells(i, 1).Value

      ' Add to the Brand Total
      Stock_Total = Stock_Total + ws.Cells(i, 7).Value
      
      'Set the Close Price
      closePrice = ws.Cells(i, 6).Value
      
      ' Print the Stock Type in the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = Stock_Name

      ' Print the Quarterly Change in the Summary Table
        ws.Range("J" & Summary_Table_Row).Value = closePrice - openPrice
            
        'Add conditional formating to summary table
        If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf ws.Range("J" & Summary_Table_Row).Value <= 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
          
      'Add value to percent change
     Percentagechange = (closePrice - openPrice) / openPrice
    
     'Print percent change to summary table
     ws.Range("K" & Summary_Table_Row).Value = Percentagechange
     
     'Format percent changed
     ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
      ' Print the Stock Amount to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Stock_Total


      'Max and Min calculations
            If Percentagechange > greatestIncrease Then
                greatestIncrease = Percentagechange
                greatesincreasestockname = Stock_Name
                
            ElseIf Percentagechange < greatestDecrease Then
                greatestDecrease = Percentagechange
                greatesincreasestockname = Stock_Name
            
            End If

        If Stock_Total > greatestStockTotal Then
            greatestStockTotal = Stock_Total
            greateststockname = Stock_Name
            End If
            
        'print to table
         ws.Cells(4, 16).Value = greateststockname
         ws.Cells(4, 17).Value = greatestStockTotal
         ws.Cells(2, 16).Value = greatesincreasestockname
         ws.Cells(2, 17).Value = greatestIncrease
         ws.Cells(3, 16).Value = greatesincreasestockname
         ws.Cells(3, 17).Value = greatestDecrease
        
        'format percentage
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
         
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1

      ' Reset the Stock Total, Open Price and Close Price
      Stock_Total = 0
      openPrice = ws.Cells(i + 1, 3).Value
      closePrice = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Stock Total
      Stock_Total = Stock_Total + Cells(i, 7).Value

    End If
    
    Next i
    
    Next ws
    
End Sub

