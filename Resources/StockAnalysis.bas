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
    Cells(1, LastColumn + 2).Value = "Ticker"
    Cells(1, LastColumn + 3).Value = "Yearly Change"
    Cells(1, LastColumn + 4).Value = "Percentage Change"
    Cells(1, LastColumn + 5).Value = "Total Stock Volume"
    
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
                    
            Cells(j, LastColumn + 2).Value = CurentTicker
            Cells(j, LastColumn + 3).Value = CurrCloseAmt - CurrOpenAmt
            'cater for divide by zero
            If CurrOpenAmt <> 0 Then
                Cells(j, LastColumn + 4).Value = (CurrCloseAmt - CurrOpenAmt) / CurrOpenAmt
            Else
                Cells(j, LastColumn + 4).Value = 0
            End If
            Cells(j, LastColumn + 5).Value = TotalVolume
            
            'Back ground colour for annual change amount, red negative, green positive
            If Cells(j, LastColumn + 3).Value < 0 Then
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
    Cells(1, LastColumn + 9).Value = "Ticker"
    Cells(1, LastColumn + 10).Value = "Value"
    Cells(2, LastColumn + 8).Value = "Greatest % Increase"
    Cells(3, LastColumn + 8).Value = "Greatest % Decrease"
    Cells(4, LastColumn + 8).Value = "Greatest Total Volume"
    
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
    
    MaxPerc = Cells(2, j).Value
    MinPerc = Cells(2, j).Value
    MaxVolume = Cells(2, j + 1).Value
    
    'Find the values
    For i = 2 To LastRow
    
        If Cells(i, j).Value > MaxPerc Then
            MaxPerc = Cells(i, j).Value
            MaxPercTicker = Cells(i, j - 2).Value
            
        End If
        
        If Cells(i, j).Value < MinPerc Then
            MinPerc = Cells(i, j).Value
            MinPercTicker = Cells(i, j - 2).Value
            
        End If
        
        If Cells(i, j + 1).Value > MaxVolume Then
            MaxVolume = Cells(i, j + 1).Value
            MaxVolumeTicker = Cells(i, j - 2).Value
            
        End If
    
    Next i
    
    'Populate table
    Cells(2, LastColumn + 9).Value = MaxPercTicker
    Cells(2, LastColumn + 10).Value = MaxPerc
    
    Cells(3, LastColumn + 9).Value = MinPercTicker
    Cells(3, LastColumn + 10).Value = MinPerc

    Cells(4, LastColumn + 9).Value = MaxVolumeTicker
    Cells(4, LastColumn + 10).Value = MaxVolume
    

End Sub


Sub PopAllSheets()
    
    'Main routine to populate each sheet
    
    Dim sht As Worksheet
    
    'Check if data is cleared or empty
    If Not IsEmpty(Cells(1, 9).Value) Then
        MsgBox ("Clear Summary Data Before Re-running")
        Exit Sub
    End If
    
    'Loop through each sheet
    For Each sht In ThisWorkbook.Worksheets
    
        sht.Activate
        
        SummaryData  'call sub routine
    
    Next sht
    
    'Return to the first worksheet
    Worksheets(1).Activate
    
End Sub


Sub ClearAllSheets()
    'Main routine to clear the populated data

    Dim sht As Worksheet
    
    'Loop through each sheet
    For Each sht In ThisWorkbook.Worksheets
    
        sht.Activate
        
        ClearData    'call sub routine
    
    Next sht
    
    'Return to the first worksheet
    Worksheets(1).Activate

End Sub



Sub ClearData()

    'Clear the Summary & Analysis data, current sheet
    
    Dim LastRow As Long
    Dim sht     As Worksheet
    Dim rng     As Range
    
    Set sht = ActiveSheet
    
    LastRow = sht.Cells(sht.Rows.Count, "I").End(xlUp).Row
    
    Set rng = sht.Range("I1:Q" & LastRow)
    
    rng.Clear
         
End Sub
