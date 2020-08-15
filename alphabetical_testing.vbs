Attribute VB_Name = "Module1"
Sub Alphabetical_testing()

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
  Cells(1, 9).Value = "Ticker"
  Cells(1, 12).Value = "Total Stock Volume"
  
  'Naming the header for Yearly Change and Percentage Change
  Cells(1, 10).Value = "Yearly Change"
  Cells(1, 11).Value = "Percentage Change"
  
  'Set an initial variable for Opening amount
  Dim Open_amount As Double
  Open_amount = 0
  
  'Set an initial variable for Closing amount
  Dim Close_amount As Double
  Close_amount = 0
    
  'Set counter to capture close amount
  Dim Counter As Double
  Counter = 0
  
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'To print the ticker , calculate the Yearly change and Percentage change
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
  
' counts the number of rows in column 1
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through each row
  ' Use lastrow variable
   
  ' Loop through all tickers
   For i = 1 To lastrow
   

    ' Check if we are still within the same Ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker
      Cells(Summary_Table_Row, 9).Value = Cells(i + 1, 1).Value
      
      'Use open amount
      Open_amount = Cells(i + 1, 3).Value
      
      'Looking in the ticker column and setting a counter to keep track of opening and closing amount for the same tracker
      Counter = WorksheetFunction.CountIf(Range("A:A"), Cells(Summary_Table_Row, 9).Value)
      Close_amount = Cells(i + Counter, 6).Value
      
      'Calculating Yearly change and printing its value in 0.00 format
      Cells(Summary_Table_Row, 10).Value = Close_amount - Open_amount
      Cells(Summary_Table_Row, 10).NumberFormat = "0.00"
      
      'Calculating the percentage change and printing its value in 0.00% format
      If Open_amount = 0 Then
            Cells(Summary_Table_Row, 11).Value = 0
      Else
            Cells(Summary_Table_Row, 11).Value = (Close_amount - Open_amount) / Open_amount
            Columns("K").NumberFormat = "0.00%"
      End If
      
           
      'Add one to the Summary_Table_Row counter
      Summary_Table_Row = Summary_Table_Row + 1
     

    End If

  Next i
  
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'To calculate the total stock volume and highlight positive change in green and negative change in red
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    
  'counts the number of rows in the Ticker column
  lastrow = Cells(Rows.Count, "I").End(xlUp).Row
  
  'Loop through the column I - tickers
  For j = 2 To lastrow
      
    'Finding the total stock volume per ticker
      Cells(j, 12).Value = WorksheetFunction.SumIfs(Range("G:G"), Range("A:A"), Cells(j, 9).Value)
  
  'Conditional formatting that will highlight positive change in green and negative change in red
    If Cells(j, 10).Value >= 0 Then
       Cells(j, 10).Interior.ColorIndex = 4
    Else
        Cells(j, 10).Interior.ColorIndex = 3
        
    End If
    
  Next j
  
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'To calculate the Greatest % Increase, decrease and total Volume
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
   
  
  'Naming the cells for Greatest % Increase, decrease and total Volume
  Cells(2, 16).Value = "Greatest % Increase"
  Cells(3, 16).Value = "Greatest % Decrease"
  Cells(4, 16).Value = "Greatest Total Volume"
  
  'Naming the Cells for Ticker and Value
   Cells(1, 17).Value = "Ticker"
   Cells(1, 18).Value = "Value"
  
  'Using the maximum and minimum functions for the range
  Cells(2, 18).Value = WorksheetFunction.Max(Range("K:K"))
  Cells(3, 18).Value = WorksheetFunction.Min(Range("K:K"))
  Cells(4, 18).Value = WorksheetFunction.Max(Range("L:L"))
  
  'formating the numbers in the cell
  Cells(2, 18).NumberFormat = "0.00%"
  Cells(3, 18).NumberFormat = "0.00%"
  
  
  'Creating another loop to find the Ticker corresponding to the Maximum and Minimum Value
  For k = 2 To lastrow
  
  'Compare the Percentage change column to the Value in the greatest % Increase cell, if the value is the same, then Print the corresponding Ticker
    If Cells(k, 11).Value = Cells(2, 18).Value Then
        Cells(2, 17).Value = Cells(k, 9).Value
   'Compare the Percentage change column to the Value in the greatest % Decrease cell, if the value is the same, then Print the corresponding Ticker
    ElseIf Cells(k, 11).Value = Cells(3, 18).Value Then
        Cells(3, 17).Value = Cells(k, 9).Value
   'Compare the Total Stock Volume column to the Value in the greatest Total Volume cell, if the value is the same, then Print the corresponding Ticker
    ElseIf Cells(k, 12).Value = Cells(4, 18).Value Then
        Cells(4, 17).Value = Cells(k, 9).Value
    End If
    
    Next k
    
  
End Sub




