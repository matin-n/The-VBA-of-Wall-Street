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

    Worksheets("All Stocks Analysis").Activate

    Range("A1").Value = "All Stocks (2018)"

    ' Create a header row
    Range("A3").Value = "Ticker"
    Range("B3").Value = "Total Daily Volume"
    Range("C3").Value = "Return"
    
    ' Array to hold 12 tickers
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
    
    Dim startValue, endValue As Integer
    startValue = 1
    endValue = 10
    
    For i = startValue To endValue
        ticker = tickers(i)
        For j = startValue To endValue
        Cells(i, j).Value = i + j
        Cells(i, j).Clear
        Next j
    Next i

End Sub
