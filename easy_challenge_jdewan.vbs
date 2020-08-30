' -----------------------------------------------------------------
'
' Instructions:
'       Create a script that will loop through all the stocks for one year and output the following information:
'       1. The ticker symbol
'       2. Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'       3. The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'       4. The total stock volume of the stock
'           image: easy_solution.png
'
' -----------------------------------------------------------------

' Part-1: Solution for single sheet
' ----------------------------
' Function for Easy Solution: as per the given format in the Homework (easy_solution.png)
'----------------------------
Sub EasySingleSheet()

    ' Define Ticker, Total Stock Volume and Ticker Counter
    Dim tickerName As String
    Dim totalStockVolume As Double
    Dim myTickerCounter As Integer
    
    ' Initailize values
    '
    myTickerCounter = 2
    totalStockVolume = 0
    
    ' Write header:  "Ticker" in cell I1 and "Total Stock Volume" in cell J1
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"
    
    ' Determine the ast row to be used for the counter
    '
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' ------------------------------------------------------
    ' Loop through all the stock
    ' i) Get ticker name and calculate total volume (using column 7) for the ticker in each row (for an ticker symbol)
    ' ii) if you hit to a differnt ticker then sumarize  "totaStockVolume" for the particular sticker
    '  then push the ticker name and totaStockVolume to the specified cells
    '
    ' ------------------------------------------------------

    
    For i = 2 To lastRow
        '  get tickername and calculate total stock volume for the same ticker
        '
        tickerName = Cells(i, 1).Value
        totalStockVolume = totalStockVolume + Cells(i, 7).Value
                
        ' Compare current ticker (i,1) with next  ticker (i+1, 1) if they are different
        ' If they are different then get tickerName and totalStockVolume, save
        ' them to cell (i, 9) and (i, 10) respectively
        '
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        '
        ' Put current tickerName to cell (i, 9) and  totalStockVolume to cell(i, 10) if different ticker is found
        '
            Cells(myTickerCounter, 9).Value = tickerName
            Cells(myTickerCounter, 10).Value = totalStockVolume
            
            ' Reset total_stock_volume  to 0 as new sticker is found
            '
            totalStockVolume = 0
            
            '  Advance counter as new sticker is found
            myTickerCounter = myTickerCounter + 1
        
            
        End If
        
    ' Go to the next itration and loop through the processingfor the new sticker
    ' until the end of last row the current worksheet
    '
    Next i

' Auto fit column J with all the numners (as numbers could be large)
'
    Columns("J").AutoFit
    
MsgBox ("Completed Easy solution for current Worksheet")

End Sub

' ------------------------------------------------------------------
' Part-2: Easy Challenge
'
' This VBA script is to run on every worksheet, i.e., every year, just by running the VBA script once.
'
' ------------------------------------------------------------------
Sub EasyMultiSheetChallenge()

    ' Define Ticker name, Total Stock Volume and Ticker Counter
    Dim tickerName As String
    Dim totalStockVolume As Double
    Dim myTickerCounter As Integer
    
    
    For Each ws In Worksheets
        ' Initailize values
        '
        myTickerCounter = 2
        totalStockVolume = 0
        
        ' Write header:  "Ticker" in cell I1 and "Total Stock Volume" in cell J1 for each sheet
        '
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"
        
        ' Determine the last row for the counter
        '
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
         
        ' ------------------------------------------------------
        ' Loop through all the stock in the current sheet
        ' i) Get ticker name and calculate total volume (using column 7) for the ticker in each row (for an ticker symbol)
        ' ii) if you hit to a differnt ticker then sumarize  "total_stock_volume" for the particular sticker
        '  then push the ticker name and total_stock_volume to the specified cells
        ' ------------------------------------------------------
        For i = 2 To lastRow
            '
             ' get tickername and calculate total stock volume for the same ticker
            '
            tickerName = ws.Cells(i, 1).Value
            totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
            
            '
            ' Compare current ticker (i,1) with next  ticker (i+1, 1) if they are different
             ' If they are different then get tickerName and totalStockVolume, save
            ' them to cell (i, 9) and (i, 10) respectively
            '
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(myTickerCounter, 9).Value = tickerName
                ws.Cells(myTickerCounter, 10).Value = totalStockVolume
                
                ' ii) reset totalStockVolume  to 0 for the for the different sticker and advance the counter
                '
                totalStockVolume = 0
                myTickerCounter = myTickerCounter + 1
            End If
        
        ' Go to the Next ticker
        '
        Next i

        'Auto fit column J to fit all the numners (as numbers could be large)
        '
        ws.Columns("J").AutoFit
    
    Next ws

MsgBox ("Completed Easy Challenge Solutions for all the Worksheets")

End Sub

' Clear worksheet  for the Easy Solution
'
Sub CleanupEasy()
    Columns("I:J").ClearContents
    Columns("I:J").ClearFormats
    Columns("I:J").UseStandardWidth = True
End Sub

' Cleanup all the worksheets under test
'
Sub CleanupEasyChallenge()
    ' Cleanup result data for all Worksheet: traverse through all the sheets one by one
    For Each ws In Worksheets
        ws.Columns("I:J").ClearContents
        ws.Columns("I:J").ClearFormats
        ws.Columns("I:J").UseStandardWidth = True
    Next ws
End Sub
