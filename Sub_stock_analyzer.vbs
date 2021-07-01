Sub stock_analyze()

'worksheet handling
Dim tot_rows As Double
Dim ws As Worksheet
'stock calculations
Dim stock_cntr As Integer
Dim curr_stock As String
Dim last_stock As String
Dim date_string As String
Dim curr_date As Date
Dim open_date As Date
Dim close_date As Date
Dim open_price As Double
Dim close_price As Double
Dim change As Double
Dim change_per As Double
Dim volume As Double
'bonus calculations
Dim max_increase As Double
Dim max_increase_stock As String
Dim max_decrease As Double
Dim max_decrease_stock As String
Dim max_volume As Double
Dim max_volume_stock As String

For Each ws In Worksheets

  ws.Select
  
  'clear meta stats
   max_increase = 0
   max_increase_stock = ""
   max_decrease = 0
   max_decrease_stock = ""
   max_volume = 0
   max_volume_stock = ""
  
  'get number of rows attribute
  tot_rows = ws.Cells(Rows.Count, "A").End(xlUp).Row
  
  'clear any contents & formatting from previous 'attempts'
   ws.Range("I1:Q" & tot_rows).ClearContents
   ws.Range("I1:Q" & tot_rows).ClearFormats
  
  'set up headers for new columns we're going to create and populate
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "Yearly Change"
  ws.Cells(1, 11).Value = "Percent Change"
  ws.Cells(1, 12).Value = "Total Stock Volume"
  ws.Range("I1:L1").Font.Bold = True

  'setup variables for first row of data
  stock_cntr = 0
  curr_stock = ws.Cells(2, 1).Value
  last_stock = curr_stock
  date_string = ws.Cells(2, 2).Value
  curr_date = DateSerial(Left(date_string, 4), Mid(date_string, 5, 2), Right(date_string, 2))
  open_date = curr_date
  close_date = curr_date

  'process ws data rows
  For i = 2 To tot_rows

    'current row attributes
    date_string = ws.Cells(i, 2).Value
    curr_date = DateSerial(Left(date_string, 4), Mid(date_string, 5, 2), Right(date_string, 2))
    curr_stock = ws.Cells(i, 1).Value
    
    If curr_stock = last_stock And ws.Cells(i + 1, 1).Value <> "" Then
      'first stock and all rows of first stock data and subsequent rows of stock data after stock ticker change
      'stock ticker change is handled below, on account of it's special...
      volume = volume + ws.Cells(i, 7).Value
      'open price on earliest date for stock is open price
      If curr_date <= open_date Then
        open_date = curr_date
        open_price = ws.Cells(i, 3).Value
      'close price on latest date for stock is close price
      ElseIf curr_date > close_date Then
        close_date = curr_date
        close_price = ws.Cells(i, 6).Value
      End If

    ElseIf curr_stock <> last_stock Then
      'stock ticker change - increment stock counter
      stock_cntr = stock_cntr + 1
      'calculate last stock results and populate results,
      ws.Cells(stock_cntr + 1, 9).Value = last_stock
      ws.Cells(stock_cntr + 1, 12).Value = volume
      'capture max volume before resetting volume
      If volume > max_volume Then
        max_volume = volume
        max_volume_stock = last_stock
      End If
      change = close_price - open_price
      ws.Cells(stock_cntr + 1, 10).Value = change
      'positive change is green, negative change is red
      If change < 0 Then
        ws.Cells(stock_cntr + 1, 10).Interior.ColorIndex = 3
      Else
        ws.Cells(stock_cntr + 1, 10).Interior.ColorIndex = 10
      End If
      'don't divide by zero.  just...don't.  ever.
      If open_price > 0 Then
        change_per = ((close_price - open_price) / open_price)
      Else
        change_per = close_price
      End If
      If change < 0 Then
        change_per = Abs(change_per) * -1
      End If
      ws.Cells(stock_cntr + 1, 11).Value = change_per
      'capture max and min change
      If change_per >= max_increase Then
        max_increase = change_per
        max_increase_stock = last_stock
      ElseIf change_per <= max_decrease Then
        max_decrease = change_per
        max_decrease_stock = last_stock
      End If
      'make sure format is correct
      ws.Cells(stock_cntr + 1, 11).NumberFormat = "0.00%"
      'setup variables for next stock
      open_date = curr_date
      close_date = curr_date
      open_price = ws.Cells(i, 3).Value
      close_price = ws.Cells(i, 6).Value
      volume = ws.Cells(i, 7).Value
  
    ElseIf curr_stock = last_stock And ws.Cells(i + 1, 1).Value = "" Then
      'last stock entry row in this sheet - increment counter, accumulate, and populate
      stock_cntr = stock_cntr + 1
      volume = volume + ws.Cells(i, 7).Value
      ws.Cells(stock_cntr + 1, 9).Value = curr_stock
      ws.Cells(stock_cntr + 1, 12).Value = volume
      'capture max volume of last stock
      If volume >= max_volume Then
        max_volume = volume
        max_volume_stock = curr_stock
      End If
      'open price on earliest date for stock is open price
      If curr_date <= open_date Then
        open_date = curr_date
        open_price = ws.Cells(i, 6).Value
      'close price on latest date for stock is close price
      ElseIf curr_date >= close_date Then
        close_date = curr_date
        close_price = ws.Cells(i, 6).Value
      End If
      change = close_price - open_price
      ws.Cells(stock_cntr + 1, 10).Value = change
      'positive change is green, negative change is red
      If change < 0 Then
        ws.Cells(stock_cntr + 1, 10).Interior.ColorIndex = 3
      Else
        ws.Cells(stock_cntr + 1, 10).Interior.ColorIndex = 10
      End If
      'I'll say it again: don't divide by zero...
      If open_price > 0 Then
        change_per = ((close_price - open_price) / open_price)
      Else
        change_per = close_price
      End If
      If change < 0 Then
        change_per = Abs(change_per) * -1
      End If
      ws.Cells(stock_cntr + 1, 11).Value = change_per
      'capture max and min change
      If change_per >= max_increase Then
        max_increase = change_per
        max_increase_stock = curr_stock
      ElseIf change_per <= max_decrease Then
        max_decrease = change_per
        max_decrease_stock = curr_stock
      End If
      'make sure format is correct
      ws.Cells(stock_cntr + 1, 11).NumberFormat = "0.00%"
    
    End If
  
    'this is kind of a big deal...
    last_stock = Cells(i, 1).Value
    
  Next i
  
  'populate the meta stats
  
  ws.Cells(1, 16).Value = "Ticker"
  ws.Cells(1, 17).Value = "Value"
  ws.Range("P1:Q1").Font.Bold = True
  
  ws.Cells(2, 15).Value = "Greatest % Increase"
  ws.Cells(2, 16).Value = max_increase_stock
  ws.Cells(2, 17).Value = max_increase
  ws.Cells(2, 17).NumberFormat = "0.00%"
  
  ws.Cells(3, 15).Value = "Greatest % Decrease"
  ws.Cells(3, 16).Value = max_decrease_stock
  ws.Cells(3, 17).Value = max_decrease
  ws.Cells(3, 17).NumberFormat = "0.00%"
  
  ws.Cells(4, 15).Value = "Greatest Total Volume"
  ws.Cells(4, 16).Value = max_volume_stock
  ws.Cells(4, 17).Value = max_volume

  
  'format the beautiful columns we just populated to show all the data
  ws.Range("I1").EntireColumn.AutoFit
  ws.Range("J1").EntireColumn.AutoFit
  ws.Range("K1").EntireColumn.AutoFit
  ws.Range("L1").EntireColumn.AutoFit
  ws.Range("O1").EntireColumn.AutoFit
  ws.Range("P1").EntireColumn.AutoFit
  ws.Range("Q1").EntireColumn.AutoFit

Next ws

For Each ws In Worksheets   'because I prefer to be left on the first worksheet
  ws.Select
  Exit For
Next ws

End Sub