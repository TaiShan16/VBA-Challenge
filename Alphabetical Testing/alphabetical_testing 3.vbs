Attribute VB_Name = "Module1"
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
        
    ' Autofit Columns
        ws.Range("I:Q").Columns.AutoFit
    
    
    
    ' track changes in column A
    ' add up totals in column G based on changes in column A
    ' each time the ticker changes in column A
    ' populate the name of the ticker in column I
    ' display the total stock volumn in column L
    ' reset the total and start tracking for the next ticker
    
    ' first find the last row in the sheet
    
    Dim lastRow
    lastRow = Cells(Rows.Count, 1).End(xlUp).row
    
    ' check on ticker names
    Dim tickerName As String
    
    ' delcare variable to hold the total volumn
    Dim volumnTotal As Double
    volumnTotal = 0  ' start the initial total at 0
    
    ' assign variable to hold the rows in the total columns ( column I to L)
    Dim tickerRows As Integer
    tickerRows = 2  ' first row to poupluate in column I to L will be row 2
    
    ' declare variable to hold the row
    Dim row As Integer
    
    ' declare variable for open value and close value
    Dim closeValue As Double
    Dim openValue As Double
    
    ' declare variable for yearChangeRow
    Dim changeRow As Integer
    
    changeRow = 1  ' first row to populate
    
    ' declare variable for greatest increase, greatest decrease and greatest total volumn
    Dim maxIncrease, maxDecrease, maxVolumn As Double
    
    ' declare variable for index of tickers
    Dim maxIndex, minIndex, volIndex As String
    
    
    
    ' --------------------------------------------
    ' LOOP THE TICKER AND TOTAL VOLUMN
    ' --------------------------------------------
    
    ' loop through the rows and check the changes in the tickers
    For row = 2 To lastRow
    
        ' check the changes in the tickers
        If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
        
            ' if the ticker changes, then display the change
            'MsgBox (Cells(row, 1).Value + " -> " + Cells(row + 1, 1).Value)
            
            ' set the ticker name
            tickerName = ws.Cells(row, 1).Value    ' grabs the value from column A BEFORE the change
            
            ' add to the total volumn
            volumnTotal = volumnTotal + ws.Cells(row, 7).Value  ' grab the value from column G BEFORE the change
            
            ' display the ticker name on the current row of the ticker column (Col I)
            ws.Cells(tickerRows, 9).Value = tickerName
            
            ' display the total volumn on the current row of Col L
            ws.Cells(tickerRows, 12).Value = volumnTotal
            
            ' add 1 to the ticker total for the next ticker
            tickerRows = tickerRows + 1
            
            ' reset the ticker for the next ticker name
            volumnTotal = 0
            
        Else
        
            ' if there is no change in the ticker name, keep adding the total volumn
            volumnTotal = volumnTotal + ws.Cells(row, 7).Value     ' Grabs the value from column G
        
        End If
        
    Next row
       
    ' --------------------------------------------
    ' LOOP YEAR CHANGE & PERCENT CHANGE
    ' --------------------------------------------
    
    For row = 2 To lastRow
    
    ' First, find value for "open"
    ' check to see if current row (ticker) equals the open in previous row
    ' if different, then we grab value for open
    
            ' current row      <>         previous row means we just start a new ticker
        If ws.Cells(row, 1).Value <> ws.Cells(row - 1, 1).Value Then
        
            ' put open value into openValue
            openValue = ws.Cells(row, 3).Value
            
            changeRow = changeRow + 1
           
            
        End If
        
    ' Second, find value for "close"
    ' check if current row (ticker) equals to nex row
    
                ' current row   <>        next row means that we are at the end of current ticker
        If ws.Cells(row, 1).Value <> ws.Cells(row + 1, 1).Value Then
            
            ' grab close value
            closeValue = ws.Cells(row, 6).Value
            
        End If
        
    ' Thrid, check if openValue has a value, and closeValue has a value
        
        If (openValue <> 0 And closeValue <> 0) Then
            
            ' subtract open value from close value and show the result in Yearly Change column
            ws.Cells(changeRow, 10).Value = closeValue - openValue
                
            ' declare variable for percentChange
            Dim percentChange As Double
            
            'YC = closeValue - openValue
            
            If openValue <> 0 Then
            
                ' calculate percent change
                percentChange = (closeValue - openValue) / openValue
                
                ' display result in Column K
                ws.Cells(changeRow, 11).Value = percentChange
                
            Else
                ' if openValue is 0, then value display is 0
                ws.Cells(changeRow, 11).Value = 0
                
            End If
            
             
            ' reset openValue and closeValue
            openValue = 0
            closeValue = 0
            
                
        End If
        
    Next row
        
        
    ' find maximum amount of percentage change
    maxIncrease = WorksheetFunction.Max(ws.Range("K2:K91"))
    
    ' display result for the greatest increase
    ws.Range("Q2") = maxIncrease
    
    ' apply percentage style to the greatest increase
    ws.Range("Q2").NumberFormat = "0.00%"
    
    ' fine the lowest amount of percentage change
    maxDecrease = WorksheetFunction.Min(ws.Range("K2:K91"))
    
    ' display result for the greatest decrease
    ws.Range("Q3") = maxDecrease
    
    ' apply percentage style to the greatest decrease
    ws.Range("Q3").NumberFormat = "0.00%"
    
    ' find maximum amount of the total volumn
    maxVolumn = WorksheetFunction.Max(ws.Range("L2:L91"))
    
    ' display result for the greatest total volumn
    ws.Range("Q4") = maxVolumn
    
    ' find ticker for greatest increase
    maxIndex = WorksheetFunction.Match(maxIncrease, ws.Range("K2:K91"), 0)
    
    ' display the ticker for greatest increase
    ws.Range("P2").Value = ws.Range("I" & maxIndex + 1).Value
    
    ' find ticker for greatest decrease
    minIndex = WorksheetFunction.Match(maxDecrease, ws.Range("K2:K91"), 0)
    
    ' display the ticker for greatest increase
    ws.Range("P3").Value = ws.Range("I" & minIndex + 1).Value
    
    ' find ticker for greatest total volumn
    volIndex = WorksheetFunction.Match(maxVolumn, ws.Range("L2:L91"), 0)
    
    ' display the ticker for greatest total volumn
    ws.Range("P4").Value = ws.Range("I" & volIndex + 1).Value
    
    Next ws
    
End Sub



