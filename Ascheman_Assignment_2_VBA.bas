Attribute VB_Name = "Module1"
Sub stocks():

For Each ws In Worksheets
  
  ' ### Finding the last row - https://www.mrexcel.com/board/threads/loop-until-last-row-in-spreadsheet.302362/
  LastRow = Rows.CountLarge
  
  ' the grouping of tickers on the right
  Dim t_counter As Long
  t_counter = 1
  
  ' add column headings
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "Yearly Change"
  ws.Cells(1, 11).Value = "Percent Change"
  ws.Cells(1, 12).Value = "Total Stock Value"
  
  ' ### for every row in the year - https://www.mrexcel.com/board/threads/loop-until-last-row-in-spreadsheet.302362/
  For i = 2 To LastRow
    
    ' If last row, exit for loop
    If ws.Cells(i, 1).Value = "" Then Exit For
    
    ' If first ticker row
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
      
      ' new t_group
      t_counter = t_counter + 1
      ws.Cells(t_counter, 9).Value = ws.Cells(i, 1).Value
      ws.Cells(t_counter, 12).Value = 0
      
      ' store opening value
      open_val = ws.Cells(i, 3).Value
      
    End If
    
    ' If last ticker row
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
      
      ' calculate yearly and percentage change
      ws.Cells(t_counter, 10).Value = ws.Cells(i, 6).Value - open_val
      ws.Cells(t_counter, 11).Value = FormatPercent(ws.Cells(t_counter, 10).Value / open_val)
      
      ' color the yearly change cells
      If ws.Cells(t_counter, 10).Value > 0 Then ws.Cells(t_counter, 10).Interior.ColorIndex = 4
      Else:
      ws.Cells(t_counter, 10).Interior.ColorIndex = 3
      
    End If
    
    ' add to total stock value
    ws.Cells(t_counter, 12).Value = ws.Cells(t_counter, 12).Value + ws.Cells(i, 7).Value
    
  Next i
  
  ' labeling functions
  ws.Cells(1, 16).Value = "Ticker"
  ws.Cells(1, 17).Value = "Value"
  ws.Cells(2, 15).Value = "Greatest % Increase"
  ws.Cells(3, 15).Value = "Greatest % Decrease"
  ws.Cells(4, 15).Value = "Greatest Total Volume"
  
  ' ### Finding min and max values - https://www.statology.org/vba-find-max-value-in-range/
  ws.Cells(2, 17).Value = FormatPercent(Application.WorksheetFunction.Max(ws.Range("K2:K3001")))
  ws.Cells(3, 17).Value = FormatPercent(Application.WorksheetFunction.Min(ws.Range("K2:K3001")))
  ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range("L2:L3001"))
     
  ' loop through tickers to find matching value
    For j = 2 To LastRow:
    
      If ws.Cells(j, 11).Value = ws.Cells(2, 17).Value Then ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
      If ws.Cells(j, 11).Value = ws.Cells(3, 17).Value Then ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
      If ws.Cells(j, 12).Value = ws.Cells(4, 17).Value Then ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
    
    Next j
    
  Next ws

End Sub

