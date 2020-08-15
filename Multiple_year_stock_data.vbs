Attribute VB_Name = "Module1"
Sub Multiple_year_data()

    'For all the worksheets
    For Each ws In Worksheets
    
  ' Set an initial variable for holding the Ticker Symbol
  Dim Ticker As String

  ' Set an initial variable for holding the total Stock Volume per Ticker
  Dim Total_Volume As Double
  Total_Volume = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  'Set an initial variable for Yearly Change
  Dim Yearly_Change As Double
  Yearly_Change = 0
  
  'Set an initial variable for Percentage Change
  Dim Percentage_Change As Double
  Percentage_Change = 0
    
  'Naming the header for Ticker and Total Stock Volume
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 12).Value = "Total Stock Volume"
  
  'Naming the header for Yearly Change and Percentage Change
  ws.Cells(1, 10).Value = "Yearly Change"
  ws.Cells(1, 11).Value = "Percentage Change"
  
  'Set an initial variable for Opening amount
  Dim Open_amount As Double
  Open_amount = 0
  
  'Set an initial variable for Closing amount
  Dim Close_amount As Double
  Close_amount = 0
    
  'Set counter
  Dim Counter As Integer
  Counter = 0
  
   
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'To print the ticker , calculate the Yearly change and Percentage change
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  
' counts the number of rows in column 1
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through each row
  ' Use lastrow variable
   
  ' Loop through all tickers
   For i = 1 To lastrow
   

    ' Check if we are still within the same Ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker
      ws.Cells(Summary_Table_Row, 9).Value = ws.Cells(i + 1, 1).Value
      
      'Use open amount
      Open_amount = ws.Cells(i + 1, 3).Value
      
      'Looking in the ticker column and setting a counter to keep track of closing amount for the same tracker
      Counter = WorksheetFunction.CountIf(ws.Range("A:A"), ws.Cells(Summary_Table_Row, 9).Value)
      Close_amount = ws.Cells(i + Counter, 6).Value
      
      'Calculating Yearly change and printing its value in 0.00 format
      ws.Cells(Summary_Table_Row, 10).Value = Close_amount - Open_amount
      ws.Cells(Summary_Table_Row, 10).NumberFormat = "0.00"
      
      'Calculating the percentage change and printing its value in 0.00% format
      If Open_amount = 0 Then
            ws.Cells(Summary_Table_Row, 11).Value = 0
      Else
            ws.Cells(Summary_Table_Row, 11).Value = (Close_amount - Open_amount) / Open_amount
            ws.Range("K:K").NumberFormat = "0.00%"
      End If
      
           
      'Add one to the Summary_Table_Row counter
      Summary_Table_Row = Summary_Table_Row + 1
     

    End If

  Next i
  
  
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'To calculate the total stock volume and highlight positive change in green and negative change in red
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
   
   
  'counts the number of rows in the Ticker column
  lastrow = ws.Cells(Rows.Count, "I").End(xlUp).Row
  
  'Loop through the column I - tickers
  For j = 2 To lastrow
      
    'Finding the total stock volume per ticker
      ws.Cells(j, 12).Value = WorksheetFunction.SumIfs(ws.Range("G:G"), ws.Range("A:A"), ws.Cells(j, 9).Value)
  
  'Conditional formatting that will highlight positive change in green and negative change in red
    If ws.Cells(j, 10).Value >= 0 Then
       ws.Cells(j, 10).Interior.ColorIndex = 4
    Else
        ws.Cells(j, 10).Interior.ColorIndex = 3
        
    End If
    
  Next j
  
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'To calculate the Greatest % Increase, decrease and total Volume
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
     
     
  'Naming the cells for Greatest % Increase, decrease and total Volume
  ws.Cells(2, 16).Value = "Greatest % Increase"
  ws.Cells(3, 16).Value = "Greatest % Decrease"
  ws.Cells(4, 16).Value = "Greatest Total Volume"
  
  'Naming the Cells for Ticker and Value
   ws.Cells(1, 17).Value = "Ticker"
   ws.Cells(1, 18).Value = "Value"
  
  'Using the maximum and minimum functions for the range
  ws.Cells(2, 18).Value = WorksheetFunction.Max(ws.Range("K:K"))
  ws.Cells(3, 18).Value = WorksheetFunction.Min(ws.Range("K:K"))
  ws.Cells(4, 18).Value = WorksheetFunction.Max(ws.Range("L:L"))
  
  'formating the numbers in the cell
  ws.Cells(2, 18).NumberFormat = "0.00%"
  ws.Cells(3, 18).NumberFormat = "0.00%"
  'Cells(4, 18).NumberFormat = "0.00"
  
  'Creating another loop to find the Ticker corresponding to the Maximum and Minimum Value
  For k = 2 To lastrow
  
  'Compare the Percentage change column to the Value in the greatest % Increase cell, if the value is the same, then Print the corresponding Ticker
    If ws.Cells(k, 11).Value = ws.Cells(2, 18).Value Then
        ws.Cells(2, 17).Value = ws.Cells(k, 9).Value
   'Compare the Percentage change column to the Value in the greatest % Decrease cell, if the value is the same, then Print the corresponding Ticker
    ElseIf ws.Cells(k, 11).Value = ws.Cells(3, 18).Value Then
        ws.Cells(3, 17).Value = ws.Cells(k, 9).Value
   'Compare the Total Stock Volume column to the Value in the greatest Total Volume cell, if the value is the same, then Print the corresponding Ticker
    ElseIf ws.Cells(k, 12).Value = ws.Cells(4, 18).Value Then
        ws.Cells(4, 17).Value = ws.Cells(k, 9).Value
    End If
    
    Next k
    
  Next ws
  
End Sub



