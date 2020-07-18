Sub Ticker()

    'Worksheet variables
    Dim WS_Count As Integer
    WS_Count = ActiveWorkbook.Worksheets.Count
    '
    'Find the last row###########
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Summary Table
    Dim SumTabRow As Integer
    
    'Stocks variables
    Dim Ticker As String
    Dim ClosePrice As Integer
    Dim OpenPrice As Integer
    Dim YearChange As Integer
    Dim PercentChange As Double
    
    'Total Stock Volume variable
    Dim TSV As Double
    TSV = 0
    
    'Loop through worksheets
    For i = 1 To WS_Count
    
        'Activate current worksheet
        Worksheets(i).Activate
        
        'Populate headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 14).Value = "Year End"
        Cells(1, 15).Value = "Year Start"
        
        'Set Summary Table to 2 each worksheet loop
        SumTabRow = 2
        
        'Stocks loop
        
        If i = 7 Then
        
            'Sheet P has 261 rows per stock
            For k = 2 To LastRow
            
                If Cells(k, 1).Value <> Cells(k + 1, 1).Value Then
            
                    'Populate Ticker
                    Ticker = Cells(k, 1).Value
                    Range("I" & SumTabRow).Value = Ticker
                    
                    'Populate Yearly Change
                    ClosePrice = Cells(k, 6).Value
                    OpenPrice = Cells(k - 260, 3).Value
                    YearChange = ClosePrice - OpenPrice
                    Range("N" & SumTabRow).Value = ClosePrice
                    Range("O" & SumTabRow).Value = OpenPrice
                    Range("J" & SumTabRow).Value = YearChange
                    
                    'Populate Percent Change
                    PercentChange = YearChange / OpenPrice
                    Range("K" & SumTabRow).Value = PercentChange
                    
                    'Check Percent Change for Condition
                    If PercentChange >= 0 Then
                        Range("K" & SumTabRow).Interior.Color = RGB(0, 255, 0)
                    Else
                        Range("K" & SumTabRow).Interior.Color = RGB(255, 0, 0)
                    End If
                    
                    'Populate Total Stock Volume
                    TSV = TSV + Cells(k, 7).Value
                    Range("L" & SumTabRow).Value = TSV
                    
                    SumTabRow = SumTabRow + 1
                    TSV = 0
                    
                Else
                    TSV = TSV + Cells(k, 7).Value
                
                End If
            Next k
        
        Else
        
            'Sheets A through F have 262 rows per stock
            For j = 2 To LastRow
            
                If Cells(j, 1).Value <> Cells(j + 1, 1).Value Then
            
                    'Populate Ticker
                    Ticker = Cells(j, 1).Value
                    Range("I" & SumTabRow).Value = Ticker
                    
                    'Populate Yearly Change
                    ClosePrice = Cells(j, 6).Value
                    OpenPrice = Cells(j - 261, 3).Value
                    YearChange = ClosePrice - OpenPrice
                    Range("N" & SumTabRow).Value = ClosePrice
                    Range("O" & SumTabRow).Value = OpenPrice
                    Range("J" & SumTabRow).Value = YearChange
                    
                    'Populate Percent Change
                    PercentChange = YearChange / OpenPrice
                    Range("K" & SumTabRow).Value = PercentChange
                    
                    'Check Percent Change for Condition
                    If PercentChange >= 0 Then
                        Range("K" & SumTabRow).Interior.Color = RGB(0, 255, 0)
                    Else
                        Range("K" & SumTabRow).Interior.Color = RGB(255, 0, 0)
                    End If
                    
                    'Populate Total Stock Volume
                    TSV = TSV + Cells(j, 7).Value
                    Range("L" & SumTabRow).Value = TSV
                    
                    SumTabRow = SumTabRow + 1
                    TSV = 0
                    
                Else
                    TSV = TSV + Cells(j, 7).Value
                    
                End If
            Next j
        End If
        
        'Worksheet identification
        MsgBox ActiveWorkbook.Worksheets(i).Name
    
    Next i

    'Returns to first worksheet
    Worksheets("A").Activate

End Sub
            

