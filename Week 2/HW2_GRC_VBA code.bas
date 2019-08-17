Attribute VB_Name = "Module1"

Sub TotalVolume()
   
    
    Dim TickerOn As String
    Dim TickerCurrentRow As String
    Dim RunningTotal As LongLong
    Dim ResultRow As LongLong
    
  
    
    For Each YearWkst In Worksheets
        YearWkst.Cells(1, 9) = "Ticker Symbol"
        YearWkst.Cells(1, 10) = "Total Volume"
    
        TickerOn = ""
        RunningTotal = 0
        ResultRow = 2
    
        lastRow = YearWkst.Cells(Rows.Count, 1).End(xlUp).Row
        
        For RowIndex = 2 To lastRow
            Dim CurrentRowTicker As String
            Dim CurrentRowVol As Long
            
            CurrentRowTicker = YearWkst.Cells(RowIndex, 1)
            CurrentRowVol = YearWkst.Cells(RowIndex, 7)
            
            If TickerOn <> CurrentRowTicker Then
                YearWkst.Cells(ResultRow, 9).Value = CurrentRowTicker
                    
                If TickerOn <> "" Then
                    YearWkst.Cells(ResultRow - 1, 10).Value = RunningTotal
                End If
                
                
                TickerOn = CurrentRowTicker
                RunningTotal = 0
                ResultRow = ResultRow + 1
            End If
            RunningTotal = RunningTotal + CurrentRowVol
        
        Next RowIndex
    
    Next
    
End Sub

    
