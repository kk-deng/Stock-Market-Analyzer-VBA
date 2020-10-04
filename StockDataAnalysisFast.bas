Attribute VB_Name = "StockDataAnalysisFast"
Sub MainRun()
    Dim ws As Worksheet
    Dim Start, Finish, TotalTime
    Dim RunOption As String
    
    'Display a msgbox for user to choose between functions
    RunOption = MsgBox("Do you want to reset summary cells before processing?", 3)
    
    'Set a start timer
    Start = Timer
    
    'If resetting cells is chosen, call the sub to clear all previous summary cell values
    If RunOption = vbYes Then
        Call ResetCells
    ElseIf RunOption = vbCancel Then
        Exit Sub
    End If
    
    'Main functions to generate stock stats for all worksheets
    For Each ws In ActiveWorkbook.Worksheets
        Call StockYearlyChangeFast(ws)
        Call GreatestYearlyChange(ws)
    Next ws
    
    'Calculate the elaspsed time
    Finish = Timer
    TotalTime = Finish - Start
    Application.StatusBar = "Overall Elapsed Time: " & TotalTime & " seconds"
    
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
            
            'Main function to loop from row 2 to the last row
            For i = 2 To LastRow
            
                'Check if a new ticker section starts by checking the total volume for the current ticker
                If TotalVolume = 0 Then
                    'If previous row stock volume is 0, then get the open price from previous date
                    'Otherwise, get open price from the current date (Avoid 0 open price and volume)
                    If .Cells(i - 1, 7).Value = 0 Then
                        OpenPrice = .Cells(i - 1, 3).Value
                    Else
                        OpenPrice = .Cells(i, 3).Value
                    End If
                End If
                
                'Calculate the total volume
                TotalVolume = TotalVolume + .Cells(i, 7).Value
                
                'When the current ticker reaches the last date
                If .Cells(i + 1, 1).Value <> .Cells(i, 1).Value Then
                    
                    'Record the current ticker
                    .Cells(SummaryRow, 9).Value = .Cells(i, 1).Value
                    
                    'Record the total volume for this ticker
                    .Cells(SummaryRow, 12).Value = TotalVolume
                    
                    'Get the Close Price
                    ClosePrice = .Cells(i, 6).Value
                    
                    'Record the Yearly Change
                    .Cells(SummaryRow, 10).Value = ClosePrice - OpenPrice
                    
                    'Filling colour format based on the change, (Green when >0, Red when <0)
                    If .Cells(SummaryRow, 10).Value > 0 Then
                        .Cells(SummaryRow, 10).Interior.ColorIndex = 4
                    ElseIf .Cells(SummaryRow, 10).Value < 0 Then
                        .Cells(SummaryRow, 10).Interior.ColorIndex = 3
                    End If
                    
                    'Check the open and close prices are > 0 for percentage calculation
                    If (OpenPrice > 0) And (ClosePrice > 0) Then
                        .Cells(SummaryRow, 11).Value = Format(((ClosePrice - OpenPrice) / OpenPrice), "0.00%")
                    ElseIf OpenPrice = 0 Then
                        .Cells(SummaryRow, 11).Value = Format(0, "0.00%")
                    End If
                    
                    'Move to next summary row
                    SummaryRow = SummaryRow + 1
                    
                    'Reset the total volume for new ticker
                    TotalVolume = 0
                    
                End If
                
                'Update the progress of looping on the Application Status Bar
                'To reduct the update frequencey, it runs only every 5000 rows
                If i Mod 5000 = 0 Or i = LastRow Then
                    Application.StatusBar = "Progress: " & i & " of " & LastRow & " (" & Format(i / LastRow, "0%") & ")"
                End If
            
            Next i
        
        End With
            
End Sub


Sub GreatestYearlyChange(ws As Worksheet)
    Dim LastRow, GreatIncrease, GreatDecrease, GreatVolume As Double
    Dim TickerIncrease, TickerDecrease, TickerVolume As String
    
    'This is the function for the challenge question (calculate greatest percentages and volume)
    With ws
        
        'Enter information of headers
        .Cells(1, 16).Value = "Ticker"
        .Cells(1, 17).Value = "Value"
        .Cells(2, 15).Value = "Greatest % Increase"
        .Cells(3, 15).Value = "Greatest % Decrease"
        .Cells(4, 15).Value = "Greatest Total Volume"
        
        'Get the last row of current worksheet
        LastRow = .Cells(Rows.Count, 9).End(xlUp).Row
        
        'Check the greatest increase & decrease percentage for summary table
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
        
        'Record results to cells
        .Cells(2, 16).Value = TickerIncrease
        .Cells(3, 16).Value = TickerDecrease
        .Cells(4, 16).Value = TickerVolume
        
        .Cells(2, 17).Value = Format(GreatIncrease, "0.00%")
        .Cells(3, 17).Value = Format(GreatDecrease, "0.00%")
        .Cells(4, 17).Value = GreatVolume
        
        'Auto adjust column width for better view
        .Columns("O").AutoFit
        .Columns("Q").AutoFit
               
    End With
    
End Sub

Sub ResetCells()
    For Each ws In ActiveWorkbook.Worksheets
        
        'This is a function to clear all previous run results in summary area (From I to Q)
        With ws
            .Range("I:Q").ClearContents
            .Range("I:Q").ClearFormats
        End With
    Next ws

End Sub
