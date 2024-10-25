Attribute VB_Name = "Module1"
Sub Stock_Summary()


' Create a script that loops through all the stocks for each quarter and outputs the following information:

    ' Calculate the Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.

    ' Calculate the percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.

    ' Calculate the total stock volume of the stock.



'Set Variables UniqueTickers and loop through Cells (1,N)
    Dim UniqueTicker As String
    Dim i As Long
    Dim n As Long
    Dim ws As Worksheet

  
' Create variables to hold Opening price at beginning, Closing price at end of Quarter and volume traded.  Along with LastRoW of the Tickersybols.
    Dim Open_Price_Q As Double
    Dim Close_Price_Q As Double
    Dim Quarterly_change As Double
    Dim Percent_Change As Double
    Dim Vol_Total As Double
    Vol_Total = 0
    Close_Price_Q = 0
    Open_Price_Q = 0
    Percent_Change = 0
    Quarterly_change = 0
    Dim lastRow As Long
    Dim Greatest_Per_Increase As Double
    Dim GPI_Ticker As String
    Dim Greatest_Per_Decrease As Double
    Dim GPD_Ticker As String
    Dim Greatest_Total_Volume As Double
    Dim GTV_Ticker As String
    Greatest_Per_Increase = 0
    GPI_Ticker = ""
    Greatest_Per_Decrease = 0
    GPD_Ticker = ""
    Greatest_Total_Volume = 0
    GTV_Ticker = ""
    
    
    'Loop Through All Sheets
    For Each ws In Worksheets
    
'Calculate LastRow of Column A subtracting 1 for the header
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row + 1
        
 'Keep track of the location for each Ticker sybole in a summmary table and create headers
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
   ws.Cells(1, 9) = "Ticker"
   ws.Cells(1, 10) = "Quarterly Change"
   ws.Cells(1, 11) = "Percent Change"
   ws.Cells(1, 12) = "Total Stock Volume"
   ws.Cells(1, 9).ColumnWidth = 10
   ws.Cells(1, 10).ColumnWidth = 16
   ws.Cells(1, 11).ColumnWidth = 15
   ws.Cells(1, 12).ColumnWidth = 18
   
'Loop Through all Ticker items and pull UniqueTickers from

For i = 2 To lastRow
    
    ' Check if first Row of ticker to capture the opening price
      If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        Open_Price_Q = ws.Cells(i, 3).Value
        'Set the Ticker name
         UniqueTicker = ws.Cells(i, 1).Value
         'Print the Credit CArd Brand in the Summary Table
         ws.Range("I" & Summary_Table_Row).Value = UniqueTicker
         'Add to the Vol Total
        Vol_Total = Vol_Total + ws.Cells(i, 7).Value
         
         
    
    ' Check if we are still within the same Ticker, Capture closing price and last Volume amount.
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
              
       'Captuer Closing Price
        Close_Price_Q = ws.Cells(i, 6).Value
       'Add Stock Volume
         Vol_Total = Vol_Total + ws.Cells(i, 7).Value
       'Print Vol_total
        ws.Range("L" & Summary_Table_Row).Value = Vol_Total
       'Calculate Quarterly change
       Quarterly_change = (Close_Price_Q - Open_Price_Q)
       'Print % change
        ws.Range("J" & Summary_Table_Row).Value = Quarterly_change
       'Calculate %change
        Percent_Change = (Quarterly_change / Open_Price_Q)
        'Print % change
        ws.Range("K" & Summary_Table_Row).Value = Percent_Change
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
               
    'Reset the Vol_total
      Vol_Total = 0
      Close_Price_Q = 0
      Open_Price_Q = 0
      Percent_Change = 0
      Quarterly_change = 0
      Summary_Table_Row = Summary_Table_Row + 1
       
           
   ' If the cell immediately following a row is the same Ticker, capter opeing, and add Volume.
      Else
      ' Add to the Vol Total
        Vol_Total = Vol_Total + ws.Cells(i, 7).Value
 
    End If
    
    Next i
    
    'Set Color gradient for Quarterly Change
       Set Rng = ws.Range("J2:J" & lastRow)
       Rng.FormatConditions.Delete
    For Each cell In Rng
        ' Check the value of the cell and apply the corresponding color
        If cell.Value > 0 Then
            ' Change the cell's interior color to green
            cell.Interior.Color = RGB(0, 255, 0) ' Green
        ElseIf cell.Value = 0 Then
            ' Change the cell's interior color to white
            cell.Interior.Color = RGB(255, 255, 255) ' White
        Else
            ' Change the cell's interior color to red
            cell.Interior.Color = RGB(255, 0, 0) ' Red
        End If
   Next cell
   ws.Cells(1, 16) = "Ticker"
   ws.Cells(1, 17) = "Value"
   ws.Cells(2, 15) = "Greatest Percent Increase"
   ws.Cells(3, 15) = "Greatest Percent Decrease"
   ws.Cells(4, 15) = "Greatest Total Volume"
   ws.Cells(1, 15).ColumnWidth = 25
   ws.Cells(1, 16).ColumnWidth = 10
   ws.Cells(1, 17).ColumnWidth = 10
   
    Greatest_Per_Increase = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastRow))
        
     ws.Cells(2, 17).Value = Greatest_Per_Increase
     ws.Cells(2, 17).NumberFormat = "0.00%"
      For i = 2 To lastRow
        If ws.Cells(i, 11).Value = ws.Cells(2, 17).Value Then
        GPI_Ticker = ws.Cells(i, 9).Value
        ws.Cells(2, 16).Value = GPI_Ticker
        End If
      Next i
        
        
     Greatest_Per_Decrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastRow))
     ws.Cells(3, 17).Value = Greatest_Per_Decrease
     ws.Cells(3, 17).NumberFormat = "0.00%"
     For i = 2 To lastRow
        If ws.Cells(i, 11).Value = ws.Cells(3, 17).Value Then
        GPD_Ticker = ws.Cells(i, 9).Value
        ws.Cells(3, 16).Value = GPD_Ticker
        End If
      Next i
    
     Set Rng = ws.Range("L2:L" & lastRow)
     Greatest_Total_Volume = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
     ws.Cells(4, 17).Value = Greatest_Total_Volume
      For i = 2 To lastRow
        If ws.Cells(i, 12).Value = ws.Cells(4, 17).Value Then
        GTV_Ticker = ws.Cells(i, 9).Value
        ws.Cells(4, 16).Value = GTV_Ticker
        End If
      Next i
        
    Greatest_Per_Increase = 0
    GPI_Ticker = ""
    Greatest_Per_Decrease = 0
    GPD_Ticker = ""
    Greatest_Total_Volume = 0
    GTV_Ticker = ""
    
    lastRow = 0
    
    Next ws
End Sub
