Attribute VB_Name = "Module11"

Sub TotalVolume()
   
    
    Dim TickerOn As String
    Dim TickerCurrentRow As String
    Dim RunningTotal As LongLong
    Dim ResultRow As LongLong
    
    Dim OpenPrice As Double
    Dim ClosePrice As Double
  
    
    For Each YearWkst In Worksheets
        YearWkst.Cells(1, 9) = "Ticker Symbol"
        YearWkst.Cells(1, 12) = "Total Volume"
        YearWkst.Cells(1, 10) = "Yearly Change"
        YearWkst.Cells(1, 11) = "Percent Change"
        
    
        TickerOn = ""
        RunningTotal = 0
        ResultRow = 2
    
        lastRow = YearWkst.Cells(Rows.Count, 1).End(xlUp).Row
        
        For RowIndex = 2 To lastRow
            Dim CurrentRowTicker As String
            Dim CurrentRowVol As Long
            Dim YearlyChange As Double
            Dim PercentChange As Double
            
            
            CurrentRowTicker = YearWkst.Cells(RowIndex, 1)
            CurrentRowVol = YearWkst.Cells(RowIndex, 7)
            
            
            If TickerOn <> CurrentRowTicker Then
                YearWkst.Cells(ResultRow, 9).Value = CurrentRowTicker
                    
                If TickerOn <> "" Then
                    YearWkst.Cells(ResultRow - 1, 12).Value = RunningTotal
                End If
                
                OpenPrice = YearWkst.Cells(RowIndex, 3)
                
                TickerOn = CurrentRowTicker
                RunningTotal = 0
                ResultRow = ResultRow + 1
               
                
            End If
            RunningTotal = RunningTotal + CurrentRowVol
           
            ClosePrice = YearWkst.Cells(RowIndex, 6)
            YearlyChange = ClosePrice - OpenPrice
            YearWkst.Cells(ResultRow - 1, 10).Value = YearlyChange
            YearWkst.Cells(ResultRow - 1, 10).NumberFormat = "0.0000"
                If YearlyChange >= 0 Then
                    YearWkst.Cells(ResultRow - 1, 10).Interior.ColorIndex = 4
                Else
                    YearWkst.Cells(ResultRow - 1, 10).Interior.ColorIndex = 3
                End If
            
            YearWkst.Cells(ResultRow - 1, 11).NumberFormat = "0.00"
                If OpenPrice > 0 Then
                    PercentChange = YearlyChange / OpenPrice * 100
                End If
            YearWkst.Cells(ResultRow - 1, 11).Value = PercentChange
            
            
              
        Next RowIndex
    
    Next
    
End Sub

    
