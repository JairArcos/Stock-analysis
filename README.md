**Overview of Project**

The purpose of this project was to edit, or refactor, previous All Stock Analysis and determine if refactoring previous code successfully made the VBA script run faster in order to make code more efficient by taking fewer steps, using less memory, or improving the logic of the code.

**Results**

Main change was to include tickerIndex and three outputs arrays. Basically what it was done was to reduce times looping code including new ticker variables for making calculus with Volumes, StartingPrices and EndingPrices. 

_All Stock Analysis comparison 2017_
![image](https://user-images.githubusercontent.com/96214761/148670813-20e710bb-b67e-48cd-8856-25d62fec7988.png)

_All Stock Analysis comparison 2018_
<img width="1046" alt="AllStocksAnalysis 2018" src="https://user-images.githubusercontent.com/96214761/148671737-ecc0779b-ccae-4667-8fce-1ec8b63393e9.png">

Original 2017 report code ran in 2.19 seconds, after code was refactored code ran in 0.31 seconds, same impact was observed in 2018 report code ran,  original ran in 2.16 seconds and refactoring in 0.32 seconds.

_Code_

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
    tickerIndex = 0

    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single

    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
         tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i

    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount

    '3a) Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    '3b) Check if the current row is the first row with the selected tickerIndex.
    'If  Then
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
    
    '3c) check if the current row is the last row with the selected ticker
    'If  Then
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If

        '3d Increase the tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If

    Next i

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
    
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
    
    Next i
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    'Drill
    Range("A3:C3").Font.Color = vbBlack
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
   
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

    Next i

    End Sub

**Summary**

There are several ways to prepare codes to resolve problems or making analysis, all of them can be OK, nonetheless if we are able to identify how to refactoring or editting our codes in order to improve their performace making more robust, they will be better in terms of memory usage, running time, logic code, etc.

_**Advantages**_

Original: 
  1. Smaller codes. 
 
Refactored: 
  1. Use less memory
  2. Run faster. 
  3. In case we need to pay for memory, less memory usage means less money.

_**Disadvantes**_

Original: 
  1. Takes a more time to run
  2. More memory usage. 
 
Refactored: 
  1. Codes are bigger so that increase the risk of having issue while running. 
