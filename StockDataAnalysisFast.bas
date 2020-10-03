Attribute VB_Name = "StockDataAnalysisFast"
Sub ResetCells()
    Dim ws As Worksheet
    For Each ws In Worksheets
        With ws
            .Range("I:Q").ClearContents
            .Range("I:Q").ClearFormats
        End With
    Next ws
 
End Sub


Sub MainRun()
    Dim ws As Worksheet
    Dim Start, Finish, TotalTime
    'Set Start time
    Start = Timer
    For Each ws In ActiveWorkbook.Worksheets
        Call StockYearlyChangeFast(ws)
        Call GreatestYearlyChange(ws)
    Next ws
    'Set end time
    Finish = Timer
    TotalTime = Finish - Start
    Application.StatusBar = "Overall Elapsed Time: " & TotalTime & " seconds"
    
    
    
End Sub

Sub GreatestYearlyChange(ws As Worksheet)
    Dim LastRow, GreatIncrease, GreatDecrease, GreatVolume As Double
    Dim TickerIncrease, TickerDecrease, TickerVolume As String
    
    With ws
        
        .Cells(1, 16).Value = "Ticker"
        .Cells(1, 17).Value = "Value"
        .Cells(2, 15).Value = "Greatest % Increase"
        .Cells(3, 15).Value = "Greatest % Increase"
        .Cells(4, 15).Value = "Greatest % Increase"
        .Columns("O").AutoFit
        
        LastRow = .Cells(Rows.Count, 9).End(xlUp).Row
        'Check the greatest increase & decrease percentage
        For m = 2 To LastRow
            If .Cells(m, 11).Value > GreatIncrease Then
                GreatIncrease = .Cells(m, 11).Value
                TickerIncrease = .Cells(m, 9).Value
            ElseIf .Cells(m, 11).Value < GreatDecrease Then
                GreatDecrease = .Cells(m, 11).Value
                TickerDecrease = .Cells(m, 9).Value
            End If
            
            If .Cells(m, 12).Value > GreatVolume Then
                GreatVolume = .Cells(m, 12).Value
                TickerVolume = .Cells(m, 9).Value
            End If
        Next m
        
        .Cells(2, 16).Value = TickerIncrease
        .Cells(3, 16).Value = TickerDecrease
        .Cells(4, 16).Value = TickerVolume
        
        .Cells(2, 17).Value = Format(GreatIncrease, "0.00%")
        .Cells(3, 17).Value = Format(GreatDecrease, "0.00%")
        .Cells(4, 17).Value = GreatVolume
               
    End With
End Sub


Sub StockYearlyChangeFast(ws As Worksheet)

    Dim LastRow, OpenPrice, ClosePrice, PercentChange, TotalVolume, SummaryRow As Double
    Dim Headers() As Variant
    SummaryRow = 2
    TotalVolume = 0
    
        With ws
            'Get the last column of worksheet
            LastRow = .Cells(Rows.Count, 1).End(xlUp).Row

            
            'Input headers in the summary table
            Headers = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
            For m = 0 To 3
                .Cells(1, m + 9).Value = Headers(m)
            Next m
            
            For i = 2 To LastRow
            
                'Check if a new ticker section starts, and also check the 1st volume is not 0
                If TotalVolume = 0 Then
                    If .Cells(i - 1, 7).Value = 0 Then
                        OpenPrice = .Cells(i - 1, 3).Value
                    Else
                        OpenPrice = .Cells(i, 3).Value
                    End If
                End If
                
                'Calculate the total volume
                TotalVolume = TotalVolume + .Cells(i, 7).Value
                
                If .Cells(i + 1, 1).Value <> .Cells(i, 1).Value Then
                    'Record the first unknown ticker
                    .Cells(SummaryRow, 9).Value = .Cells(i, 1).Value
                    'Calculate the total volume for this ticker
                    .Cells(SummaryRow, 12).Value = TotalVolume
                    'Get the Close Price
                    ClosePrice = .Cells(i, 6).Value
                    'MsgBox ("Close Price " & ClosePrice)
                    'Record the Yearly Change
                    .Cells(SummaryRow, 10).Value = ClosePrice - OpenPrice
                    
                    'Filling colour based on the change
                    If .Cells(SummaryRow, 10).Value > 0 Then
                        .Cells(SummaryRow, 10).Interior.ColorIndex = 4
                    ElseIf .Cells(SummaryRow, 10).Value < 0 Then
                        .Cells(SummaryRow, 10).Interior.ColorIndex = 3
                    End If
                
                    If (OpenPrice > 0) And (ClosePrice > 0) Then
                        .Cells(SummaryRow, 11).Value = Format(((ClosePrice - OpenPrice) / OpenPrice), "0.00%")
                    ElseIf OpenPrice = 0 Then
                        .Cells(SummaryRow, 11).Value = Format(0, "0.00%")
                    
                    End If
                    
                    'Move to next summary row
                    SummaryRow = SummaryRow + 1
                    'Reset the total volume
                    TotalVolume = 0
                End If
                
                If i Mod 5000 = 0 Or i = LastRow Then

                    'Show bar for the progress
                    Application.StatusBar = "Progress: " & i & " of " & LastRow & " (" & Format(i / LastRow, "0%") & ")"
                End If
            Next i
        End With
            
End Sub

