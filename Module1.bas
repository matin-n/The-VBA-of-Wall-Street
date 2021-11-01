Attribute VB_Name = "Module1"
Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate
    
    Range("A1").Value = "DAQO (Ticker: DQ)"

    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
         
    Worksheets("2018").Activate
    Dim totalVolume, rowStart, rowEnd As Integer
    Dim startingPrice, endingPrice As Double
    totalVolume = 0
    rowStart = 2
    ' rowEnd = 3013
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row ' https://stackoverflow.com/a/61706147/9393975
    
    
    For i = rowStart To rowEnd
    
        If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
            'set starting price
            startingPrice = Cells(i, 6).Value
        ElseIf Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then
            'set ending price
            endingPrice = Cells(i, 6).Value
        End If
    
        'increase totalVolume if ticker is "DQ"
        If Cells(i, 1).Value = "DQ" Then
            totalVolume = totalVolume + Cells(i, 8).Value
        End If
        
    Next i
    
    Worksheets("DQ Analysis").Activate
    Range("A4").Value = 2018
    Range("B4").Value = totalVolume
    Range("C4").Value = (endingPrice / startingPrice) - 1
    
End Sub

Sub AllStocksAnalysis()
    
    '1) Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (2018)"
    
    ' Create a header row
    Range("A3").Value = "Ticker"
    Range("B3").Value = "Total Daily Volume"
    Range("C3").Value = "Return"
    
    '2) Initialize array of all tickers
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
    
    '3a) Initialize variables for starting price and ending price
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    '3b) Activate data worksheet
    Worksheets("2018").Activate
    
    '3c) Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '4) Loop through tickers
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        '5) loop through rows in the data
        Worksheets("2018").Activate
        For j = 2 To RowCount
            '5a) Find total volume for current ticker
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            '5b) Find starting price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            '5c) Find ending price for current ticker
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
    
End Sub
