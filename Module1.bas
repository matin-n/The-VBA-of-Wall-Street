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
