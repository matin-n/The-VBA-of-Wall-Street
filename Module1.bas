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
    totalVolume = 0
    rowStart = 2
    rowEnd = 3013
    
    For i = rowStart To rowEnd
        'increase totalVolume if ticker is "DQ"
        If Cells(i, 1).Value = "DQ" Then
            totalVolume = totalVolume + Cells(i, 8).Value
        End If
    Next i
    
    Worksheets("DQ Analysis").Activate
    Range("A4").Value = 2018
    Range("B4").Value = totalVolume
    
    
' MsgBox (totalVolume)

End Sub
