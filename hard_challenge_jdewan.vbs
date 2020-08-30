' ----------------------------------------------------------------------------------------
' Description:
' Part-1: This subroutine to perfrom as per the given instructions in the Homework and
'             produce results similar to the screen shot given in
'             ![hard_solution](Images/hard_solution.png)
'             Processing:
'             1) Call  sub ModerateSingleSheet() as in file "moderate_challenge_jdewan.vbs" to
'                 generate the data for "Ticker", "Yearly Change", "Percentage Change" and "Total Stock Volume"
'             2) Generate data in the single sheet as per below:
'                        cell "O2" = "Greatest % Increase"
'                        cell "O3" = "Greatest % Decrease"
'                        cell "O4" = "Greatest Total Volume"
'                        cell "P1" = "Ticker"
'                        cell "Q1" = "Value"
'             3. Provide Stock ticker name that has "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
'                       in a year
'                       The solutiion should be similar to the given figure "hard_solution.png" as shown on the top right corner
'
'  Notes to run:
'            1. I have loaded "sub ModerateSingleSheet() " in Module1 and then use Module2
'               run "HardSingleSheet()" to produce data for 1 year in single sheet
'
'

'  Subroutine: To display: Ticker, Yearly Change, Percent Change, Total Stock Volume,  Ticker with Greatest % Increase,
'  Greatest % Decrease and Greatest Total Volume
'
Sub HardSingleSheet()
    ' Run sub ModerateSingleSheet() at the beginning to produce data for
    ' "Ticker",  "Yearly Change", "Percent Change" and "Percent Change"
    '  They will be used in following part of the program to prodcuce final results
    '
    Call ModerateSingleSheet

    '
    '  Define range of cells for for which results will be produced
    '
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
    Range("P1") = "Ticker"                        ' Ticker names that will have the above results
    Range("Q1") = "Value"                        ' Value colum to store the above results
    
    '
    ' Define variables
    '
    Dim lowestValue, highestValue As Double         ' lowest value of a Ticker
    Dim lowestRowCount, highestRowCount, greatestTotalVolumeCount As Integer    ' define counter index
    Dim greatestTotalVolume As Double                  ' Greatest or Maximum  total volume of a Ticker
    
    
    ' Initailize values
    lowestValue = 0
    highestValue = 0
    greatestTotalVolume = 0
    

    ' Determine the ast row of column 9 to be used for the counter
    ' column 9 contains the "Ticker" to be used for procesing further
    '
    lastRow = Cells(Rows.Count, 9).End(xlUp).Row
    
    For i = 2 To lastRow
        
       '  Check for the Highest Value among all the Ticker
       '  If value is higher than than the current higher then replace cell value with the new higher value from
       '  Cell(i, 11), where "i" is the row # used as counter in the for loop
       '
        If Cells(i, 11) > highestValue Then
            highestValue = Cells(i, 11)
            highestRowCount = i                 ' save the counter value to be used later to write cell value
        End If
        
        '  Check for the lowest Value among all the Tickers
        '  If value is less than than the current lowest then replace cell value with the new lowest from
        '   Cell(i, 11), where "i" is the row # used as counter in the for loop
        '
        If Cells(i, 11) < lowestValue Then
            lowestValue = Cells(i, 11)
            lowestRowCount = i               ' save the counter value to be used later to write cell value
        End If
        
        '  Check for the greatestTotalVolume among all the Tickers
        '  If Cells(i, 12) value is greater than than the cuurent greatestTotalVolume then
        '  update the greatestTotalVolume with the content in cell Cells(i, 12)
        '  where "i" s the row # used as counter in the for loop
        '
        If Cells(i, 12) > greatestTotalVolume Then
            greatestTotalVolume = Cells(i, 12)
            greatestTotalVolumeCount = i                '  save the counter value to be used later to write cell value
        End If
        
        ' Go to the next row until the lastRow
    Next i
    
End Sub


' -----------------------------------------------------------------------------------------------
'   Part-2: HardMultiSheetChallenge
'
'   Descriiption:
'   The objective of this subroutine is to run on all the sheets at once and produce the results as below:
'   1) Generate data in the single sheet as per below:
'                        cell "O2" = "Greatest % Increase"
'                        cell "O3" = "Greatest % Decrease"
'                        cell "O4" = "Greatest Total Volume"
'                        cell "P1" = "Ticker"
'                        cell "Q1" = "Value"
'   2) Provide Stock ticker name that has "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
'                       in a year
'                       The solutiion should be similar to the given figure "hard_solution.png" as shown on the top right corner
'
'  Notes to run:
'            1. Load  "sub ModerateSingleSheet() " in Module1 and then use Module2
'               run "HardMultiSheetChallenge()" to produce data for all years in multi0le sheets
'
'-----------------------------------------------------------------------------------------------
'  Subroutine: To display: Ticker, Yearly Change, Percent Change, Total Stock Volume,  Ticker with Greatest % Increase,
'  Greatest % Decrease and Greatest Total Volume
'
Sub HardMultiSheetChallenge()

    ' 1. Run sub ModerateMultiSheetChallenge() at the beginning to produce data for
    ' "Ticker",  "Yearly Change", "Percent Change" and "Percent Change"
    '  They will be used in following part of the program to prodcuce final results
    '
    ' 2. call ModerateMultiSheetChallenge function to populate data
    '
    Call ModerateMultiSheetChallenge
    
    
    '
    '  Loop through each of the Worksheets and do the processing
    '
    For Each ws In Worksheets
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        
        '
        ' Define variables
        '
        Dim lowestValue, highestValue As Double         ' lowest value of a Ticker
        Dim lowestRowCount, highestRowCount, greatestTotalVolumeCount As Integer    ' define counter index
        Dim greatestTotalVolume As Double                  ' Greatest or Maximum  total volume of a Ticker
        
        ' Initailize values
        lowestValue = 0
        highestValue = 0
        greatestTotalVolume = 0
        
        
        
        ' Determine the ast row of column 9 to be used for the counter
        ' column 9 contains the "Ticker" to be used for procesing further
        '
        lastRow = Cells(Rows.Count, 9).End(xlUp).Row
        
         
        ' Loop through all the rows starting from row 2 and do further processing to find the results
        ' that will be written later  for "Greatest % Increase", Greatest % Decrease", "Greatest Total Volume"
        '  "Ticker" name and "Value"
        '
        For i = 2 To lastRow
            '  Check for the Highest Value among all the Ticker
            '  If value is higher than than the current higher then replace cell value with the new higher value from
            '  ws.Cells(i, 11), where "i" is the row # used as counter in the for loop
            '
            If ws.Cells(i, 11) > highestValue Then
                highestValue = ws.Cells(i, 11)
                 highestRowCount = i
            End If
            
            '  Check for the lowest Value among all the Tickers
            '  If value is less than than the current lowest then replace cell value with the new lowest from
            '   ws.Cells(r, 11), where "r" s the row # used as counter in the for loop
            '
            If ws.Cells(i, 11) < lowestValue Then
                lowestValue = ws.Cells(i, 11)
                lowestRowCount = i
            End If
            
            '  Check for the greatestTotalVolume among all the Tickers
            '  If ws.Cells(i, 12) value is greater than than the cuurent greatestTotalVolume then
            '  update the greatestTotalVolume with the content in cell ws.Cells(i, 12)
            '  where "i" s the row # used as counter in the for loop
            '
            If ws.Cells(i, 12) > greatestTotalVolume Then
                greatestTotalVolume = ws.Cells(i, 12)
                greatestTotalVolumeCount = i
            End If
        Next i
        
        ' Write all the results  in cell P2, P3 and P4  for the Ticker names that
        ' satified the critera: Greatest % Increase, Greatest % Decrease, Greatest Total Volume
        '
        ws.Range("P2") = ws.Cells(highestRowCount, 9).Value
        ws.Range("P3") = ws.Cells(lowestRowCount, 9).Value
        ws.Range("P4") = ws.Cells(greatestTotalVolumeCount, 9).Value
        
        ' Write all the results  in cell Q2, Q3 and Q4  for the Ticker names that
        ' satified the critera: Greatest % Increase, Greatest % Decrease, Greatest Total Volume
        '
        ws.Range("Q2") = highestValue
        ws.Range("Q3") = lowestValue
        ws.Range("Q4") = greatestTotalVolume
        
        '
        ' Format Q2 and Q3 data in % form
        '
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"

        ' Just fit all the columns O, P and Q so that data is properly visiable
        '
        ws.Columns("O").AutoFit
        ws.Columns("P").AutoFit
        ws.Columns("Q").AutoFit
    
    Next ws
End Sub

' Cleanup data for single sheet
'
Sub CleanupHardSingleSheet()

    Call CleanupModerateSingleSheet

    Columns("O:Q").ClearContents
    Columns("O:Q").ClearFormats
    Columns("O:Q").UseStandardWidth = True
End Sub

'
' Cleanup data for multiple sheets
'
Sub CleanupHardMultiSheetChallenge()
    Call CleanupModerateMutiSheetChallenge
    
    ' Cleanup result data for all Worksheet: traverse through all the sheets one by one
    For Each ws In Worksheets
        ws.Columns("O:Q").ClearContents
        ws.Columns("O:Q").ClearFormats
        ws.Columns("O:Q").UseStandardWidth = True
    Next ws
End Sub

