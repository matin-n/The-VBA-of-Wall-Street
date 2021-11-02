Attribute VB_Name = "Module1"
Sub DQAnalysis()

    ' Get user input on the year they would like to run analysis on
    yearValue = InputBox("What year would you like to run the analysis on?")

    Worksheets("DQ Analysis").Activate
    
    Range("A1").Value = "DAQO (Ticker: DQ)"
    
    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
         
    Worksheets(yearValue).Activate
    Dim totalVolume, rowStart, rowEnd As Integer
    Dim startingPrice, endingPrice As Double
    totalVolume = 0
    rowStart = 2
    ' rowEnd = 3013
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row ' https://stackoverflow.com/a/61706147/9393975
    
    
    For i = rowStart To rowEnd
    
        If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
            ' Set starting price
            startingPrice = Cells(i, 6).Value
        ElseIf Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then
            ' Set ending price
            endingPrice = Cells(i, 6).Value
        End If
    
        ' Increase totalVolume if ticker is "DQ"
        If Cells(i, 1).Value = "DQ" Then
            totalVolume = totalVolume + Cells(i, 8).Value
        End If
        
    Next i
    
    Worksheets("DQ Analysis").Activate
    Range("A4").Value = yearValue
    Range("B4").Value = totalVolume
    Range("C4").Value = (endingPrice / startingPrice) - 1
    
End Sub

Sub AllStocksAnalysis()

    ' Get user input on the year they would like to run analysis on
    yearValue = InputBox("What year would you like to run the analysis on?")
      
    ' 1) Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    ' Create a header row
    Range("A3").Value = "Ticker"
    Range("B3").Value = "Total Daily Volume"
    Range("C3").Value = "Return"
    
    ' 2) Initialize array of all tickers
    Dim tickers(11) As String
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
    
    ' 3a) Initialize variables for starting price and ending price
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    ' 3b) Activate data worksheet
    Worksheets(yearValue).Activate
    
    ' 3c) Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' 4) Loop through tickers
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        ' 5) loop through rows in the data
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
            ' 5a) Find total volume for current ticker
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            ' 5b) Find starting price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            ' 5c) Find ending price for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
        Next j
        
        '6) Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Next i
    
    Call formatAllStocksAnalysisTable
End Sub

'Sub yearValueAnalysis()

    ' yearValue = InputBox("What year would you like to run the analysis on?")
    ' Range("A1").Value = "All Stocks (" + yearValue + ")"

'End Sub

Sub formatAllStocksAnalysisTable()

    ' Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True ' Header
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "$#,##0.00" ' Total Daily Volume
    Range("C4:C15").NumberFormat = "0.00%" ' Return %
    Columns("B").AutoFit
    
    ' Conditional color coding on return %
    Dim dataRowStart, dataRowEnd As Integer
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
        If Cells(i, 3) > 0 Then
            ' Color the cell green if postive return
            Cells(i, 3).Interior.Color = vbGreen
        ElseIf Cells(i, 3) < 0 Then
            ' Color the cell red the cell red if negative return
            Cells(i, 3).Interior.Color = vbRed
        Else
            ' Clear the cell color in unhandled cases
            Cells(i, 3).Interior.Color = xlNone
        End If
    Next i
    
End Sub

Sub ClearWorksheet()
    Cells.Clear
End Sub