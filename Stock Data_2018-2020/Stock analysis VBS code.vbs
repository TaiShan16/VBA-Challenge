Sub stockData()
    
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Sheets
    
    ' Assign column name on the worksheet
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
    ' Assign cell name for greatest increase, greatest decrease, greatest volumn, ticker and volumn
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volumn"
        
    ' Make all column names bold
        ws.Range("A1:L1").Font.Bold = True
        
    ' Student's name & professior's name (just telling who prepared this workbook)
        ws.Range("S2").Value = "Prepared by: Ratima Chowadee"
        ws.Range("S3").Value = "Dr. Carl Arrington's Data Analytics Class"
        
    ' Autofit Columns
        ws.Range("I:Q").Columns.AutoFit
      
    
    ' track changes in column A
    ' add up totals in column G based on changes in column A
    ' each time the ticker changes in column A
    ' populate the name of the ticker in column I
    ' display the total stock volumn in column L
    ' reset the total and start tracking for the next ticker
    
    ' first find the last row in the sheet
    
    Dim lastRow As Long
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).row
    
    ' check on ticker names
    Dim tickerName As String
    
    ' delcare variable to hold the total volume
    Dim volumeTotal As Double
    volumeTotal = 0  ' start the initial total at 0
    
    ' assign variable to hold the rows in the total columns ( column I to L)
    Dim tickerRows As Integer
    tickerRows = 2  ' first row to poupluate in column I to L will be row 2
    
    ' declare variable to hold the row
    'Dim row As Integer
    Dim row As Long
    
    ' declare variable for open value and close value
    Dim closeValue As Double
    Dim openValue As Double
    openValue = 0   'initialize to zero before going through loop
    closeValue = 0  'initialize to zero before going through loop
    
    ' declare variable for year Change Row
    ' this keeps track of the row we will be putting values for Ticker, yearly change, % change, etc.
    Dim changeRow As Integer
    changeRow = 1  ' first row to populate
    
    ' declare variable for greatest increase, greatest decrease and greatest total volume
    Dim maxIncrease, maxDecrease, maxVolume As Double
    
    ' declare variable for index of tickers
    Dim maxIndex, minIndex, volIndex As String
    
    
    
    ' --------------------------------------------
    ' LOOP THE TICKER AND TOTAL VOLUME
    ' --------------------------------------------
    
    ' loop through the rows and check the changes in the tickers
    For row = 2 To lastRow
    
        ' check the changes in the tickers
        ' if current row ticker is NOT the same as the next row ticker
        ' it means that we are at the end of the current ticker
        If ws.Cells(row, 1).Value <> ws.Cells(row + 1, 1).Value Then
        
            ' if the ticker changes, then display the change
            'MsgBox (Cells(row, 1).Value + " -> " + Cells(row + 1, 1).Value)
            
            ' set the ticker name
            tickerName = ws.Cells(row, 1).Value    ' grabs the value from column A BEFORE the change
            
            ' add to the total volume
            volumeTotal = volumeTotal + ws.Cells(row, 7).Value  ' grab the value from column G BEFORE the change
            
            ' display the ticker name on the current row of the ticker column (Col I)
            ws.Cells(tickerRows, 9).Value = tickerName
            
            ' display the total volume on the current row of Col L
            ws.Cells(tickerRows, 12).Value = volumeTotal
            
            ' add 1 to the ticker total for the next ticker
            tickerRows = tickerRows + 1
            
            ' reset the ticker for the next ticker name
            volumeTotal = 0
            
        Else
        
            ' if there is no change in the ticker name, keep adding the total volume
            volumeTotal = volumeTotal + ws.Cells(row, 7).Value     ' Grabs the value from column G
        
        End If
        
    Next row
       
    ' --------------------------------------------
    ' LOOP YEAR CHANGE & PERCENT CHANGE
    ' --------------------------------------------
    
 
    
    For row = 2 To lastRow
    
    ' First, find value for "open"
    ' check to see if current row (ticker) equals the open in previous row
    ' if different, then we grab value for open
    
            ' current row  <>  previous row means we just start a new ticker
        If ws.Cells(row, 1).Value <> ws.Cells(row - 1, 1).Value Then
        
            ' put open value into openValue, this is the FIRST ticker found
            openValue = ws.Cells(row, 3).Value
            
            
            'need to check to see if openValue = 0 or not
            'if openvalue is 0, need to get value from next row,
            'one (or one after that) that is not zero or will be dividing by zero!!!
            If openValue = 0 Then
                            
                'once non-zero value found, save it as openValue
                openValue = ws.Cells(row + 1, 3).Value
                
                'MsgBox ("got NEXT ticker since current is zero, value = " + Str(openValue))
                      
            End If
             
             'because changeRow starts at 1 (see initialization), need to increment
             'so that values are put into correct row
             changeRow = changeRow + 1
   
        End If
        
    ' Second, find value for "close"
    ' check if current row (ticker) equals to nex row
    
        ' current row <> next row means that we are at the end of current ticker
        If ws.Cells(row, 1).Value <> ws.Cells(row + 1, 1).Value Then
            
            ' grab close value
            closeValue = ws.Cells(row, 6).Value
            
        End If
        
    ' Thrid, check if openValue has a value, and closeValue also has a value
    ' if both are not zero (has values), then we can perform calculations
        
        If (openValue <> 0 And closeValue <> 0) Then
        
            'MsgBox ("openvalue = " + Str(openValue) + ", closevalue = " + Str(closeValue))
            
            ' find the yearly change value and put the result in column J
            ' subtract open value from close value
            ws.Cells(changeRow, 10).Value = closeValue - openValue
                
            ' find percent change value
            ' declare variable for percentChange
            Dim percentChange As Double
 
            
            ' calculate percent change
            percentChange = (closeValue - openValue) / openValue
                
            ' display result in Column K
            ws.Cells(changeRow, 11).Value = percentChange
                
             
            ' reset openValue and closeValue
            ' this is so that we don't have wrong values for ticker calculation
            openValue = 0
            closeValue = 0
            
                
        End If
        
        ' --------------------------------------------
        ' CONDITIONAL FORMATTING
        ' --------------------------------------------
        
        ' Conditional formatting for Yearly Change: Negative Value = red, Positive Value = green
            If ws.Cells(changeRow, 10).Value <= 0 Then
                'change yearlyChange color to red
                ws.Cells(changeRow, 10).Interior.ColorIndex = 3
                
            Else
                'change yearlyChange color to green
                ws.Cells(changeRow, 10).Interior.ColorIndex = 4
                
            End If
            
         ' apply percentage style to all value in Column K (Percent Change)
                ws.Cells(changeRow, 11).NumberFormat = "0.00%"
        
    Next row
        
    
            
    
    '--------------------------------------------
    ' GREATEST INCREASE / DECREASE / TOTAL VOLUME
    ' --------------------------------------------
    
    ' declare variable for column K range
    Dim myRangeK As String
    myRangeK = "K2" & ":K" & changeRow
    
    ' delcare variable for column L range
    Dim myRangeL As String
    myRangeL = "L2" & ":L" & changeRow
    
        
    ' find maximum amount of percentage change
    maxIncrease = WorksheetFunction.Max(ws.Range(myRangeK))
    
    ' display result for the greatest increase
    ws.Range("Q2") = maxIncrease
    
    ' apply percentage style to the greatest increase
    ws.Range("Q2").NumberFormat = "0.00%"
    
    ' fine the lowest amount of percentage change
    maxDecrease = WorksheetFunction.Min(ws.Range(myRangeK))
    
    ' display result for the greatest decrease
    ws.Range("Q3") = maxDecrease
    
    ' apply percentage style to the greatest decrease
    ws.Range("Q3").NumberFormat = "0.00%"
    
    ' find maximum amount of the total volume
    maxVolume = WorksheetFunction.Max(ws.Range(myRangeL))
    
    ' display result for the greatest total volume
    ws.Range("Q4") = maxVolume
    
    ' find ticker for greatest increase
    maxIndex = WorksheetFunction.Match(maxIncrease, ws.Range(myRangeK), 0)
    
    
    ' display the ticker for greatest increase
    ws.Range("P2").Value = ws.Range("I" & maxIndex + 1).Value
    
    ' find ticker for greatest decrease
    minIndex = WorksheetFunction.Match(maxDecrease, ws.Range(myRangeK), 0)
    
    ' display the ticker for greatest increase
    ws.Range("P3").Value = ws.Range("I" & minIndex + 1).Value
    
    ' find ticker for greatest total volume
    volIndex = WorksheetFunction.Match(maxVolume, ws.Range(myRangeL), 0)
      
    ' display the ticker for greatest total volumn
    ws.Range("P4").Value = ws.Range("I" & volIndex + 1).Value
    

    
    
    Next ws
     
End Sub


' ------------------------------------------------------------------------------------------------------------------------------
' This VBA script is complied as a homework assignment for Data Analytics Boot Camp course @ Georgia Tech Professional Education
' This script is written by Ratima Chowadee

