Attribute VB_Name = "StockAnalysis"
Sub SummaryData()

    '========================================================================================
    'Define variables
    
    Dim i               As Long
    Dim j               As Long
    Dim sht             As Worksheet
    Dim LastRow         As Long
    Dim LastColumn      As Long
    
    Dim CurentTicker    As String
    Dim CurentVolume    As Double
    
    Dim NextTicker      As String
    Dim TotalVolume     As Double
    
    Dim CurrOpenAmt     As Double
    Dim CurrCloseAmt    As Double
            
    Dim NextOpenAmt     As Double
    Dim NextCloseAmt    As Double
    
    
    Dim MaxPercTicker   As String
    Dim MinPercTicker   As String
    Dim MaxVolumeTicker As String
    Dim MaxPerc         As Double
    Dim MinPerc         As Double
    Dim MaxVolume       As Double
    
    
    '========================================================================================
    'Summary Table
    
    'Set the Headings of the summary table
    
    Set sht = ActiveSheet

    'Get Last Used Column
    LastColumn = sht.Cells(1, sht.Columns.Count).End(xlToLeft).Column
    
    'Set the text for the columns
    Cells(1, LastColumn + 2) = "Ticker"
    Cells(1, LastColumn + 3) = "Yearly Change"
    Cells(1, LastColumn + 4) = "Percentage Change"
    Cells(1, LastColumn + 5) = "Total Stock Volume"
    
    'Format the headings
    Rows("1:1").RowHeight = 30
    Range("I1:L1").WrapText = True
    
    '--------------------------------
    'Populate the summary data set
    TotalVolume = 0
    MaxVolume = 0
    j = 2
    
    'Get the last row of column A
    LastRow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row
    
    NextOpenAmt = Cells(j, 3)
    
    For i = 2 To LastRow
    
        CurentTicker = Cells(i, 1).Value
        NextTicker = Cells(i + 1, 1).Value
        
        CurentVolume = Cells(i, 7)
        
        TotalVolume = TotalVolume + CurentVolume
        
        If CurentTicker <> NextTicker Then
        
            CurrOpenAmt = NextOpenAmt
            CurrCloseAmt = Cells(i, 6).Value
            NextOpenAmt = Cells(i + 1, 3).Value
                    
            Cells(j, LastColumn + 2) = CurentTicker
            Cells(j, LastColumn + 3) = CurrCloseAmt - CurrOpenAmt
            'cater for divide by zero
            If CurrOpenAmt <> 0 Then
                Cells(j, LastColumn + 4) = (CurrCloseAmt - CurrOpenAmt) / CurrOpenAmt
            Else
                Cells(j, LastColumn + 4) = 0
            End If
            Cells(j, LastColumn + 5) = TotalVolume
            
            'Back ground colour for annual change amount, red negative, green positive
            If Cells(j, LastColumn + 3) < 0 Then
                Cells(j, LastColumn + 3).Interior.ColorIndex = 3
            Else
                Cells(j, LastColumn + 3).Interior.ColorIndex = 4
            End If
            
            TotalVolume = 0
            
            j = j + 1
        End If
     
    Next i
          
    'Format the numbers in the summary data set
    LastRow = sht.Cells(sht.Rows.Count, "J").End(xlUp).Row
  
    Range("J2:J" & LastRow).NumberFormat = "#,##0.00"
    Range("K2:K" & LastRow).NumberFormat = "0.00%"
    Range("L2:L" & LastRow).NumberFormat = "#,##0"
     

    '========================================================================================
    'Analysis Data Set
    
    'Set the text for the columns
    Cells(1, LastColumn + 9) = "Ticker"
    Cells(1, LastColumn + 10) = "Value"
    Cells(2, LastColumn + 8) = "Greatest % Increase"
    Cells(3, LastColumn + 8) = "Greatest % Decrease"
    Cells(4, LastColumn + 8) = "Greatest Total Volume"
    
    'Set width of column is the analysis data set
    Columns("L").ColumnWidth = 15
    Columns("O").ColumnWidth = 20
    Columns("Q").ColumnWidth = 15
    
    'Format the numbers in the analysis data set
    Cells(2, LastColumn + 10).NumberFormat = "0.00%"
    Cells(3, LastColumn + 10).NumberFormat = "0.00%"
    Cells(4, LastColumn + 10).NumberFormat = "#,###"
    
 
    'Last column of summary data set
    LastRow = sht.Cells(sht.Rows.Count, "I").End(xlUp).Row
    
    j = 11
    
    MaxPerc = Cells(2, j)
    MinPerc = Cells(2, j)
    MaxVolume = Cells(2, j + 1)
    
    'Find the values
    For i = 2 To LastRow
    
        If Cells(i, j) > MaxPerc Then
            MaxPerc = Cells(i, j)
            MaxPercTicker = Cells(i, j - 2)
            
        End If
        
        If Cells(i, j) < MinPerc Then
            MinPerc = Cells(i, j)
            MinPercTicker = Cells(i, j - 2)
            
        End If
        
        If Cells(i, j + 1) > MaxVolume Then
            MaxVolume = Cells(i, j + 1)
            MaxVolumeTicker = Cells(i, j - 2)
            
        End If
    
    Next i
    
    'Populate table
    Cells(2, LastColumn + 9) = MaxPercTicker
    Cells(2, LastColumn + 10) = MaxPerc
    
    Cells(3, LastColumn + 9) = MinPercTicker
    Cells(3, LastColumn + 10) = MinPerc

    Cells(4, LastColumn + 9) = MaxVolumeTicker
    Cells(4, LastColumn + 10) = MaxVolume
    

End Sub


Sub PopAllSheets()

    Dim sht As Worksheet
    
    For Each sht In ThisWorkbook.Worksheets
    
        sht.Activate
        
        SummaryData
    
    Next sht
    
    'Return to the first worksheet
    Worksheets(1).Activate
    
End Sub


Sub ClearAllSheets()

    Dim sht As Worksheet
    
    For Each sht In ThisWorkbook.Worksheets
    
        sht.Activate
        
        ClearData
    
    Next sht
    
    'Return to the first worksheet
    Worksheets(1).Activate

End Sub



Sub ClearData()

    Dim LastRow As Long
    Dim sht     As Worksheet
    Dim rng     As Range
    
    Set sht = ActiveSheet
    
    LastRow = sht.Cells(sht.Rows.Count, "I").End(xlUp).Row
    
    Set rng = sht.Range("I1:Q" & LastRow)
    
    rng.Clear
         
End Sub
