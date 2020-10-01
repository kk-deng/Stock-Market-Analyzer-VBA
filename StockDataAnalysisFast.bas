Attribute VB_Name = "StockDataAnalysisFast"
Sub StockDataAnalysisFast():
    
    'Get the last column of worksheet
    Dim LastRow As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim TotalVolume As Double
    Dim SummaryRow As Double
    Dim Headers As Variant
    Dim ws As Worksheet
    Dim Start, Finish, TotalTime
    
    SummaryRow = 2
    TotalVolume = 0
    
    'Set Start time
    Start = Timer
    
    For Each ws In Worksheets
        With ws
        
            LastRow = .Cells(Rows.Count, 1).End(xlUp).Row
            'MsgBox (LastRow)
        
            'Input headers in the summary table
            Headers = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
            For m = 0 To 3
                .Cells(1, m + 9).Value = Headers(m)
            Next m
        
        End With
        
    Next ws
    
End Sub

