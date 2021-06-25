# Green Stock Analysis
## Project Overview
### Background
My friend, Steve, wants a green stock analyzed for his parents to see if investing is worth it. The Visual Basic Application in Excel was used to find the stock's daily volume and annual return. Eleven other green stocks were then analyzed to see how the original one compared to them. Finally, I was able to use this analysis to inform Steve the best investment option for his parent. 
### Purpose
The purpose of this project was to learn how to analyze multiple stocks using VBA efficiently. After running the initial analysis of the twelve different stocks, it became evident that the code could be run faster if refactored. This project analyzes the efiiciency of my refactored code. 
## Results
To make the code more efficient, I switched the nesting order of my for loops. In order to do this, I generated four separate arrays: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. The tickers array was used to create the ticker symbol of a stock. I created a tickerIndex variable to access the correct index across the four different arrays. This variable lets me assign the tickerVolumes, tickerStartingPrices, and tickerEndingPrices to each ticker symbol before iterating over all the rows in the dataset. This way, the analysis was completed much quicker as shown below.
### Original Code
```
Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime As Single
    
    'Ask client what year they would like to run with input box
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer
'1) Format the output sheet on the "All Stocks Analysis" worksheet.

Worksheets("All Stocks Analysis").Activate

Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
     Cells(3, 1).Value = "Ticker"
     Cells(3, 2).Value = "Total Daily Volume"
     Cells(3, 3).Value = "Return"
      
'2) Initialize an array of all tickers.

Dim tickers(12) As String

tickers(0) = "AY"
tickers(1) = "CSIQ"
tickers(2) = "DQ"
tickers(3) = "ENPH"
tickers(4) = "FSLR"
tickers(5) = "HASI"
tickers(6) = "JKS"
tickers(7) = "RUN"
tickers(8) = "SEDG"
tickers(9) = "SPWR"
tickers(10) = "TERP"
tickers(11) = "VSLR"

'3a) Initialize variables for the starting price and ending price.

    Dim startingPrice As String
    Dim endingPrice As String
    
'3b) Activate the data worksheet.

Worksheets(yearValue).Activate

'3c) Find the number of rows to loop over.

RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'4) Loop through the tickers.

For i = 0 To 11
    Ticker = tickers(i)
    totalVolume = 0
    
    '5) Loop through the rows in the data.
    Worksheets(yearValue).Activate
    For J = 2 To RowCount

        
    '5a) Find total volume for the current ticker.
     If Cells(J, 1).Value = Ticker Then
        
            totalVolume = totalVolume + Cells(J, 8).Value
    
    End If
    
    '5b) Find starting price for the current ticker.
       If Cells(J - 1, 1).Value <> Ticker And Cells(J, 1).Value = Ticker Then
            startingPrice = Cells(J, 6).Value
            
        End If
        
    '5c) Find ending price for the current ticker.
        If Cells(J + 1, 1).Value <> Ticker And Cells(J, 1).Value = Ticker Then
            endingPrice = Cells(J, 6).Value
            
        End If
        
        Next J
    
'6) Output the data for the current ticker.
Worksheets("All Stocks Analysis").Activate
Cells(4 + i, 1).Value = Ticker
Cells(4 + i, 2).Value = totalVolume
Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
    Next i
    
'Formatting
Worksheets("All Stocks Analysis").Activate
Range("A3:C3").Font.Bold = True
Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("A3:C3").Font.Size = 13
Range("B4:B15").NumberFormat = "$#,##0"
Range("C4:C15").NumberFormat = "0.00%"
Columns("B").AutoFit

dataRowStart = 4
dataRowEnd = 15
For i = dataRowStart To dataRowEnd

    If Cells(i, 3) > 0 Then
     
        'Change cell color to green
        Cells(i, 3).Interior.Color = vbGreen
    
    ElseIf Cells(i, 3) < 0 Then
    
        'Change cell color to red
        Cells(i, 3).Interior.Color = vbRed
    
    Else
    
        'Clear cell color
        Cells(i, 3).Interior.Color = xlNone
        
    End If
    

Next i

    endTime = Timer
    MsgBox "This code ran in" & (endTime - startTime) & "seconds for the year" & (yearValue)

End Sub
```

### Refactored Code
```
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    Dim tickerIndex As Single
    tickerIndex = 0

    '1b) Create three output arrays
    
    ReDim tickerVolumes(12) As Long
    ReDim tickerStartingPrices(12) As Single
    ReDim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
   
   For i = 0 To 11
    tickerVolumes(i) = 0
    
    Next i
            
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
           
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1

        End If
        
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = ((tickerEndingPrices(i) / tickerStartingPrices(i)) - 1)
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```
### Run-time for Both Sets of Code and yearValue
#### Original Code Run-time

#### Refactored Code Run-time

## Summary
### Advantages and Disadvantages of Refactoring Code
The main advantage of refactoring code is to make it quicker and more efficient. Disadvantages include potentially messing up the original code that works and even making it unusable. While refactoring can be very helpful, it is important to be careful and always save original code.
### Refactoring Code in VBA Script
Refactoring code in VBA script is advantageous since one can use as much as the original code as needed. It is also possible to place the new code next to the original code using different modules, which is useful. The main disadvantage with VBA script is the syntax sensitivity. If one does not understand the syntax, it will be difficult to rewrite the code more efficiently. 
