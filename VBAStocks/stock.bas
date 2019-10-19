Sub stock()

Dim ws As Worksheet
Dim MinMax As Integer
Dim tickername As String
Dim tickertotal As Variant
Dim tablerow As Integer
Dim tickerrows As Integer
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim yearchange As Double
Dim percentchange As Variant
Dim lastrow As Long
Dim MaxPercent As Double
Dim MaxTickerName As String
Dim MinPercent As Double
Dim MinTickerName As String
Dim MaxVolume As Variant
Dim MaxVolumeTickerName As String
    
    For Each ws In ThisWorkbook.Worksheets
    
        ws.Activate
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        tablerow = 2
        tickerrows = 0
        MinMax = 2
        MaxPercent = 0
        MinPercent = 0
        MaxVolume = 0

        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        
        For i = 2 To lastrow
    
            If Cells(i, 1) <> Cells(i + 1, 1) Then
            
                ClosePrice = Cells(i, 6).Value
                tickername = Cells(i, 1).Value
                yearchange = ClosePrice - OpenPrice
                If OpenPrice = 0 Then
                    percentchange = "N/A"
                Else
                    percentchange = (ClosePrice / OpenPrice) - 1
                    If percentchange > MaxPercent Then
                        MaxPercent = percentchange
                        MaxTickerName = tickername
                    End If
                    If percentchange < MinPercent Then
                        MinPercent = percentchange
                        MinTickerName = tickername
                    End If
                End If
                tickertotal = tickertotal + Cells(i, 7).Value
                Range("I" & tablerow).Value = tickername
                Range("J" & tablerow).Value = yearchange
                If percentchange <> "N/A" Then
                    Range("K" & tablerow).Value = percentchange
                    Cells(tablerow, 11).NumberFormatLocal = "0.00%"
                Else
                    Range("K" & tablerow).Value = percentchange
                End If
                If percentchange > 0 And percentchange <> "N/A" Then
                    Range("K" & tablerow).Interior.ColorIndex = 4
                Else
                    If percentchange < 0 Then
                       Range("K" & tablerow).Interior.ColorIndex = 3
                    Else
                        Range("K" & tablerow).Interior.ColorIndex = 0
                    End If
                End If
                Range("L" & tablerow).Value = tickertotal
                If MaxVolume < tickertotal Then
                    MaxVolume = tickertotal
                    MaxVolumeTickerName = tickername
                End If
                tablerow = tablerow + 1
                tickertotal = 0
                tickerrows = 0
            
            Else
                If tickerrows = 0 Then
                  OpenPrice = Cells(i, 3).Value
                 tickerrows = tickerrows + 1
                    tickertotal = tickertotal + Cells(i, 7).Value
                Else
                    tickerrows = tickerrows + 1
                    tickertotal = tickertotal + Cells(i, 7).Value
                    
                End If

            End If
            
        Next
        
        Range("O" & MinMax).Value = "Greatest % Increase"
        Range("P" & MinMax).Value = MaxTickerName
        Range("Q" & MinMax).Value = MaxPercent
        Cells(MinMax, 17).NumberFormatLocal = "0.00%"
        MinMax = MinMax + 1
        Range("O" & MinMax).Value = "Greatest % Decrease"
        Range("P" & MinMax).Value = MinTickerName
        Range("Q" & MinMax).Value = MinPercent
        Cells(MinMax, 17).NumberFormatLocal = "0.00%"
        MinMax = MinMax + 1
        Range("O" & MinMax).Value = "Greatest Total Volume"
        Range("P" & MinMax).Value = MaxVolumeTickerName
        Range("Q" & MinMax).Value = MaxVolume
        
        ws.Columns("I:Q").AutoFit


    Next ws
        

End Sub
