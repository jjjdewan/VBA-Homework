' ---------------------------------------------------------------------------------
' Homework Description: Moderate
' Create a script that will loop through all the stocks for one year and output the following information.
' 1. The ticker symbol.
' 2. Yearly change from opening price at the beginning of a given year
'     to the closing price at the end of that year.
' 3. The percent change from opening price at the beginning of a given year
'      to the closing price at the end of that year.
' 4. The total stock volume of the stock.
' 5. Use conditional formatting that will highlight positive change in green and negative change in red
'
' Challenges:
' 6. Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet,
'       i.e., every year, just by running the VBA script once.
'  The result should look as follows.
'   moderate_solution.png
'
'------------------------------------------------------------------------------------
' Part-1: Solution for single sheet
' -------------------------------------------------------------------------------------
'
' This subroutine is used to solve the problems from 1-5 as above
'
Sub ModerateSingleSheet()
    
    ' Define Variables
    '
    Dim tickerName As String                                      ' ticker name
    Dim totalStockVolume As Double                           ' total ticker volume
    Dim myTickerCounter As Integer                           ' counter to keep track of row saving open-close values
    Dim myOpenCloseCounter As Double                    ' keep an eye on open close vlaue in a row
    Dim yearBeginPrice, yearEndPrice As Double         ' Price at the begging and end of the year
    
    ' Initailize values
    '
    ' Init row # to write ticker summary
    myTickerCounter = 2
    
    ' init counter (row #) to keep any eye on open-close vlaues
    myOpenCloseCounter = 2
    
    ' Initialize total ticker volume to zero
    totalStockVolume = 0
    
    ' Write header:  "Ticker" in cell I1, "Total Stock Volume" in cell J1,
    '                       "Percent Change" in cell K1 and "Percent Change" in cell L1
    '
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    ' Determine the ast row to be used for the counter
    '
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    '
    ' ------------------------------------------------------
    ' Loop through all the stock
    ' i) Get ticker name and calculate total volume (using column 7) for the ticker in each row (for an ticker symbol)
    ' ii) if you hit to a differnt ticker then sumarize  "totaStockVolume" for the particular sticker
    '  then push the ticker name and totaStockVolume to the specified cells
    '
    ' ------------------------------------------------------
      
    For i = 2 To lastRow

        tickerName = Cells(i, 1).Value
        yearlyOpenPrice = Cells(myOpenCloseCounter, 3)
        totalStockVolume = totalStockVolume + Cells(i, 7).Value
        
       ' Compare current ticker (i,1) with next  ticker (i+1, 1) if they are different
        '  then get  all I need for current ticker: ticketName, yearlyClosePrice,and yearly change in price
        '
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ' Get sticker name, closing price and yearly change in price
            '
            Cells(myTickerCounter, 9).Value = tickerName
            yearlyClosePrice = Cells(i, 6)
            
            ' Price change from the begin to the end date of the year
            Cells(myTickerCounter, 10).Value = yearlyClosePrice - yearlyOpenPrice
            
            ' If we have opening value = 0, then just set cell to null
            ' to avoid dividing by 0
            
            ' Check if the yearlyOpenPrice is 0 then set the cell to Null or can have issue in divide by 0 for
            ' % calculation
            '
            If yearlyOpenPrice = 0 Then
                ' Make the cell Null value if yearly opening pricing is "0"
                Cells(myTickerCounter, 11).Value = Null
            Else
                ' If not begin price is not 0 then calculate rate of change over begin and
                ' end price of the year to find % change
                '
                Cells(myTickerCounter, 11).Value = (yearlyClosePrice - yearlyOpenPrice) / yearlyOpenPrice
                
            End If
            
            ' Now put  totalStockVolume to cell
            '
            Cells(myTickerCounter, 12).Value = totalStockVolume
            
            ' Format yearly change in price with color (column 10 cells ) as per below:
            ' positive change in green  (if >0) and negative change in red (if < 0)
            '
            If Cells(myTickerCounter, 10).Value > 0 Then
                Cells(myTickerCounter, 10).Interior.ColorIndex = 4
            Else
                Cells(myTickerCounter, 10).Interior.ColorIndex = 3
            End If
            
            ' Format column 11 cell with %
            '
            Cells(myTickerCounter, 11).NumberFormat = "0.00%"
            
            ' Reset v
            totalStockVolume = 0                                    ' reset total volume to '0'
            myTickerCounter = myTickerCounter + 1       ' advance ticket counter for next row
            myOpenCloseCounter = i + 1                         '  go to next ticker of the set
        End If
        
    Next i

    Columns("J").AutoFit
    
    
    Columns("K").AutoFit
    Columns("L").AutoFit

End Sub

'------------------------------------------------------------------------------------
' Part-2: Solution for multiple sheet to prododuce data as per example: moderate_solution.png
' -------------------------------------------------------------------------------------

' This subroutine is used to solve the problems from 1-6 in the Homework Description
' with the following challenes:
'  The VBA script runs on every worksheet and produces results as per the given
'  sample solution format: moderate_solution.png
'
Sub ModerateMultiSheetChallenge()
    
    ' Define Variables
    '
    Dim tickerName As String                                      ' ticker name
    Dim totalStockVolume As Double                           ' total ticker volume
    Dim myTickerCounter As Integer                           ' counter to keep track of row saving open-close values
    Dim myOpenCloseCounter As Double                    ' keep an eye on open close vlaue in a row
    Dim yearBeginPrice, yearEndPrice As Double         ' Price at the begging and end of the year
    
    '  Go through each of the worksheets and update them
    '
    For Each ws In Worksheets
        
         ' Initailize values
        '
        ' Init row # to write ticker summary
        myTickerCounter = 2
        
        ' init counter (row #) to keep any eye on open-close vlaues
        myOpenCloseCounter = 2
        
        ' Initialize total ticker volume to zero
        totalStockVolume = 0
    
        
        ' Write header:  "Ticker" in cell I1, "Total Stock Volume" in cell J1,
        '                       "Percent Change" in cell K1 and "Percent Change" in cell L1
        '
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Determine the last row to be used for the counter
        '
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastRow
        
            '
             ' get tickername, calculate total stock volume and price at the beginning of
             ' of the year for the same ticker
            '
            tickerName = ws.Cells(i, 1).Value
            totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
            yearBeginPrice = ws.Cells(myOpenCloseCounter, 3)
            
             '
             ' Compare current ticker (i,1) with next  ticker (i+1, 1)
             ' If they are different then get tickerName and price difference
             ' from beginning and end day of the year for the current one and put them
             ' in proper cells of the sheet
            '
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(myTickerCounter, 9).Value = tickerName
                yearEndPrice = ws.Cells(i, 6)
                ws.Cells(myTickerCounter, 10).Value = yearEndPrice - yearBeginPrice
                
                ' Check if the yearlyOpenPrice is 0 then set the cell to Null or can have issue in divide by 0 for
                '   % calculation
                If yearBeginPrice = 0 Then
                    ws.Cells(myTickerCounter, 11).Value = Null
                Else
                    ' If begin price is not 0 then calculate rate of change over begin and
                    ' end price of the year to find % change
                    '
                    ws.Cells(myTickerCounter, 11).Value = (yearBeginPrice - yearEndPrice) / yearBeginPrice
                
                End If   ' End of check for 0
                
                
                '
                ' Get totalStockVolume and then copy to cell in column 12
                ws.Cells(myTickerCounter, 12).Value = totalStockVolume
                
                 ' Format yearly change in price with color (column 10 cells ) as per below:
                ' positive change in green  (if >0) and negative change in red (if < 0)
                '
                If ws.Cells(myTickerCounter, 10).Value > 0 Then
                    ws.Cells(myTickerCounter, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(myTickerCounter, 10).Interior.ColorIndex = 3
                End If
                
                ' Format the value in cell in column 11 of of the current row to %
                ' to reflect % change in value
                '
                ws.Cells(myTickerCounter, 11).NumberFormat = "0.00%"
                
                ' reset volume count to 0,
                ' move to next row to write ticker summary to in new table,
                ' update to first row of ticker group
                
                ' Reset value:
                '
                totalStockVolume = 0                                      ' reset total volume count to '0'
                myTickerCounter = myTickerCounter + 1       ' advance ticket counter for next row
                myOpenCloseCounter = i + 1                         '  go to next ticker of the set
                
            
            End If
            
            ' Go to next group of ticker and do the same processing that group
            
        Next i
        
        '
        ' Auto fit the columns J, K and L in order to fit the values in the column
        ws.Columns("J").AutoFit
        ws.Columns("K").AutoFit
        ws.Columns("L").AutoFit

    '
    ' Go to next worksheet and perform the same operation as it was done for current sheet
    '
    Next ws
End Sub

' Clean up the data generated by ModerateSingleSheet() sub
'
Sub CleanupModerateSingleSheet()
    Columns("I:L").ClearContents
    Columns("I:L").ClearFormats
    Columns("I:L").UseStandardWidth = True
End Sub

' Clean up the data generated by ModerateMultiSheetChallenge() sub
'
Sub CleanupModerateMutiSheetChallenge()

    ' Cleanup result data for all Worksheet: traverse through all the sheets one by one
    For Each ws In Worksheets
        ws.Columns("I:L").ClearContents
        ws.Columns("I:L").ClearFormats
        ws.Columns("I:L").UseStandardWidth = True
    Next ws
End Sub


