Attribute VB_Name = "Module1"
Sub StockAnalysis()
'Set variables

'Total stock volume
    Dim Total As Double
'Variable that'll help control the loop
    Dim row As Long
'Variable that holds the number of rows in a sheet
    Dim rowCount As Double
'Variable that holds the change for each stock change quarterly
    Dim quarterlyChange As Double
'Variable that holds the percent change for each stock in the sheet
    Dim percentageChange As Double
'Variable that holds the rows of the summary table row
    Dim summaryTableRow As Long
'Variable that will hold the start of the stock's rows in the sheet
    Dim stockStartRow As Long
'start row for a stock (location of the first open)
    Dim startValue As Long
'finds the last ticker in the sheet
    Dim lastTicker As String
    
    'Loop through all the sheets in the excel workbook
    For Each ws In Worksheets
    

'Set Title Row for the new columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
'Set the title row of the totals section
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
'Values for the summary table row

'table row starts at 0 in the sheet so add 2 relative to the header
    summaryTableRow = 0
'total stock volume for a stock starts at 0
    Total = 0
'quarterly change starts at 0
    quartleryChange = 0
'first stock in the sheet is going to be on row 2
    stockStartRow = 2
'first of the first stock value is on row 2
    startValue = 2

'get the value of the last row in the current sheet
    rowCount = ws.Cells(Rows.Count, "A").End(xlUp).row
    
'find the last ticker so that we can break out of the loop
    lastTicker = ws.Cells(rowCount, 1).Value
    
'loop until we get to the end of the sheet
    For row = 2 To rowCount
    
'check for any changes in the tickers
    If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
    
'If there is a change in the Column A
'add to the total stock volume one last time
    Total = Total + ws.Cells(row, 7).Value

'check to see if the value of the total stock volume is 0
    If Total = 0 Then
'print the results in the summary tables I-L
    ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value  'Prints ticker value from colum A
    ws.Range("J" & 2 + summaryTableRow).Value = 0   'prints a 0 in column J (Quarterly chnage)
    ws.Range("K" & 2 + summaryTableRow).Value = 0   'prints a 0 in column K (Percent change)
    ws.Range("L" & 2 + summaryTableRow).Value = 0   'prints a 0 in column L (Total stock volume)

    
    Else
        'find the first non-zero first open value for the stock
    If ws.Cells(startValue, 3).Value = 0 Then
        'if the first open is 0, search for the first non-zero stock open value by moving to the next rows
        For findValue = startValue To row
        
        'check to see if the rows open value does not = 0
        If ws.Cells(findValue, 3).Value <> 0 Then
        'once we have a non-zero first open value, that value becomes the row where we can track our first open
            startValue = findValue
            'finally break from loop
            Exit For
         End If
        
        Next findValue
    End If
        
        
        'calculate the quarterly change (difference in the last close and first open)
        quarterlyChange = ws.Cells(row, 6).Value - ws.Cells(startValue, 3).Value
    
        'calculate the percent change (quarterly change / first open)
        percentChange = quarterlyChange / ws.Cells(startValue, 3).Value
        'Print the results
        ws.Range("I" & 2 + summaryTableRow).Value = Cells(row, 1).Value            'Prints ticker value from colum A
        ws.Range("J" & 2 + summaryTableRow).Value = quarterlyChange                'prints a 0 in column J (Quarterly chnage)
        ws.Range("K" & 2 + summaryTableRow).Value = percentChange                  'prints a 0 in column K (Percent change)
        ws.Range("L" & 2 + summaryTableRow).Value = Total                          'prints a 0 in column L (Total stock volume)
        
        
        'Color the quarterly change column in the summary based on our values
        If quarterlyChange > 0 Then
            'color green
            ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 4
        ElseIf quarterlyChange < 0 Then
            'color red
            ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 3
        Else
            'color the celle clear or keep it normal
            ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 0
        End If
        
        'reset / update the values for the next ticker
        Total = 0               'resets the total stock volume for the next ticker
        averageChange = 0       'resets the average change for the next ticker
        quarterlyChange = 0     'resets the     uarterly change for the next ticker
        startValue = row + 1    'move the start row
        'move to the next row in the summary table
        summaryTableRow = summaryTableRow + 1
        
      End If
        
    
    Else
    
' if we are in the same ticker keep adding to the total stock volume
    Total = Total + ws.Cells(row, 7).Value
'get the value from colum 7
    
      End If
       
    
    Next row
  
    'clean up incase we have extra data in the summary section
      'find the last row of the data in the summary
      
    'update the summary table row
    summaryTableRow = ws.Cells(Rows.Count, "I").End(xlUp).row
      
    'find the last data in the extra rows from columns J-L
    Dim lastExtraRow As Long
    lastExtra = ws.Cells(Rows.Count, "J").End(xlUp).row
      
    'loop that clears the extra data from columns I-L
        For e = summaryTableRow To lastExtraRow
          'For loop that goes through columns I-L (9-12)
          For Column = 9 To 12
            ws.Cells(e, Column).Value = ""
            ws.Cells(e, Column).Interior.ColorIndex = 0
            
        Next Column
    Next e
            
    'print the summary aggregates
    'after generating the information in the summary we're going to find the greatest percent increase and decrease
    ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & summaryTableRow + 2))
    ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & summaryTableRow + 2))
    ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & summaryTableRow + 2))
    
    'Use match() to find the row numbers of the ticker names associated with the greatest percent increase and decrease, then find the greatest total stock vol
    Dim greatestIncreaseRow As Double
    Dim greatestDecreaseRow As Double
    Dim greatestTotVolRow As Double
    greatestIncreaseRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & summaryTableRow + 2)), ws.Range("K2:K" & summaryTableRow + 2), 0)
    greatestDecreaseRow = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & summaryTableRow + 2)), ws.Range("K2:K" & summaryTableRow + 2), 0)
    greatestTotVolRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & summaryTableRow + 2)), ws.Range("L2:L" & summaryTableRow + 2), 0)
    
    
    'show ticker symbol for the greatest increase, greatest decrease, greatest total stock volume
    ws.Range("P2").Value = ws.Cells(greatestIncreaseRow + 1, 9).Value
    ws.Range("P3").Value = ws.Cells(greatestDecreaseRow + 1, 9).Value
    ws.Range("P4").Value = ws.Cells(greatestTotVolRow + 1, 9).Value
    
    'format the summary table columns
    For s = 0 To summaryTableRow
        ws.Range("J" & 2 + s).NumberFormat = "0.00"   'Formats quarterly change
        ws.Range("K" & 2 + s).NumberFormat = "0.00%"   'Formats percent change
        ws.Range("L" & 2 + s).NumberFormat = "#.###"   'Formats total stock volume
        
    Next s
    
    'Format the summary aggregate
    ws.Range("Q2").NumberFormat = "0.00%"     'format the greatest % increase
    ws.Range("Q3").NumberFormat = "0.00%"     'format the greatest number decrease
    ws.Range("Q4").NumberFormat = "#,###"      'format the greatest stock volume
    
    
'Fix/condition how the title in the columns look
    ws.Columns("A:Q").AutoFit
    
    
    Next ws
    
    
    
End Sub
