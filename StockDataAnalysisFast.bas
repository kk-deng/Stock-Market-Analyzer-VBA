Attribute VB_Name = "Module2"
Sub ResetCells()
    Dim ws As Worksheet
    For Each ws In Worksheets
        With ws
            .Range("I:L").ClearContents
            .Range("I:L").ClearFormats
        End With
    Next ws
 
End Sub


Sub forEachWs()
    Dim ws As Worksheet
    Dim Start, Finish, TotalTime
    Dim AllRows As Double
    AllRows = 0
    'Set Start time
    Start = Timer
    For Each ws In ActiveWorkbook.Worksheets
        Call StockYearlyChangeFast(ws)
    Next ws
    'Set end time
    Finish = Timer
    TotalTime = Finish - Start
    Application.StatusBar = "Stats: " & AllRows & " rows processed. Elapsed Time: " & TotalTime & " s"
    
End Sub


Sub StockYearlyChangeFast(ws As Worksheet)

    Dim LastRow, OpenPrice, ClosePrice, PercentChange, TotalVolume, SummaryRow As Double
    Dim Headers() As Variant
    SummaryRow = 2
    TotalVolume = 0
    
        With ws
            'Get the last column of worksheet
            LastRow = .Cells(Rows.Count, 1).End(xlUp).Row
            AllRows = AllRows + LastRow
            
            'Input headers in the summary table
            Headers = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
            For m = 0 To 3
                .Cells(1, m + 9).Value = Headers(m)
            Next m
            
            For i = 2 To LastRow
            
                'Check if a new ticker section starts, and also check the 1st volume is not 0
                If TotalVolume = 0 And .Cells(i - 1, 7).Value <> 0 Then
                    If .Cells(i, 3).Value <> 0 Then
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
                    If Cells(SummaryRow, 10).Value > 0 Then
                        .Cells(SummaryRow, 10).Interior.ColorIndex = 4
                    ElseIf Cells(SummaryRow, 10).Value < 0 Then
                        .Cells(SummaryRow, 10).Interior.ColorIndex = 3
                    End If
                
                    If (OpenPrice > 0) And (ClosePrice > 0) Then
                        .Cells(SummaryRow, 11).Value = Format(((ClosePrice - OpenPrice) / OpenPrice), "0.00%")
                    
                    End If
                    
                    'Move to next summary row
                    SummaryRow = SummaryRow + 1
                    'Reset the total volume
                    TotalVolume = 0
                End If
                
                If i Mod 5000 = 0 Then

                    'Show bar for the progress
                    Application.StatusBar = "Progress: " & i & " of " & LastRow & " (" & Format(i / LastRow, "0%") & ")"
                End If
            Next i
        End With

            
End Sub

