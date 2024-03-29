Attribute VB_Name = "Module7"
Sub AllStocksAnalysisRefactoredtrial()

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
    
    Dim TickerVolumes(12) As Long
    Dim TickerStartingPrices(12) As Single
    Dim TickerEndingPrices(12) As Single
    
''2a) Create a for loop to initialize the tickerVolumes to zero.

    For b = 0 To 11
    
        TickerVolumes(b) = 0
       
    Next b

''2b) Loop over all the rows in the spreadsheet.

     For j = 2 To RowCount
    
'3a) Increase volume for current ticker

         TickerVolumes(tickerIndex) = TickerVolumes(tickerIndex) + Cells(j, 8).Value
    
        
'3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then

        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
        
            TickerStartingPrices(tickerIndex) = Cells(j, 6).Value
           
        End If
        
'3c) check if the current row is the last row with the selected ticker
        'If the next row�s ticker doesn�t match, increase the tickerIndex.
        'If  Then
        
         If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
         
            TickerEndingPrices(tickerIndex) = Cells(j, 6).Value
         
         End If

        '3d Increase the tickerIndex.
        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerIndex = tickerIndex + 1
        
        End If
        
        Next j
           
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For tickerIndex = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + tickerIndex, 1).Value = tickers(tickerIndex)
        Cells(4 + tickerIndex, 2).Value = TickerVolumes(tickerIndex)
        Cells(4 + tickerIndex, 3).Value = TickerEndingPrices(tickerIndex) / TickerStartingPrices(tickerIndex) - 1
            
    Next tickerIndex
    
    'Formatting
    
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For L = dataRowStart To dataRowEnd
        
        If Cells(L, 3) > 0 Then
        
        Cells(L, 3).Interior.Color = vbGreen
            
        Else
        
        Cells(L, 3).Interior.Color = vbRed
            
        End If
        
    Next L
 
    endTime = Timer
    
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

