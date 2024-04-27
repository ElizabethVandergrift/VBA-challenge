# VBA-challenge
Assignment Files

    Sub Stock_Market()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim TickerRow As Long
    Dim Stock_Code As String
    Dim Start_Price As Double
    Dim End_Price As Double
    Dim Change As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim Start As Long
    Dim j As Long
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolume As Double
    Dim MaxIncreaseTicker As String
    Dim MaxDecreaseTicker As String
    Dim MaxVolumeTicker As String
    Dim SummaryRow As Integer
    
    For Each ws In Worksheets
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Add titles for the summary table
        ws.Range("I1").Value = "Ticker Code"
        ws.Range("J1").Value = "Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        
        ' Initialize variables
        TickerRow = 2 ' Start from row 2 for the summary table
        Start = 2
        TotalVolume = 0
        j = 0
        MaxIncrease = 0
        MaxDecrease = 0
        MaxVolume = 0
        
        For i = 2 To LastRow
        
            Stock_Code = ws.Cells(i, 1)
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' Calculate Total Volume
                TotalVolume = TotalVolume + CDbl(ws.Cells(i, 7).Value)
                
                ' Check if volume is 0
                If TotalVolume <> 0 Then
                    If Start <> i Then
                        Change = ws.Cells(i, 6).Value - ws.Cells(Start, 3)
                        PercentChange = Change / ws.Cells(Start, 3)
                    Else
                        Change = 0
                        PercentChange = 0
                    End If
                    
                    ' Output summary information to the summary table
                    ws.Cells(TickerRow, 9).Value = Stock_Code ' Ticker code
                    ws.Cells(TickerRow, 10).Value = Change ' Change in price
                    ws.Cells(TickerRow, 11).Value = PercentChange & "%" ' Percent change
                    ws.Cells(TickerRow, 12).Value = TotalVolume ' Total volume
            
                    ' Color the cell based on the change
                    Select Case Change
                        Case Is > 0
                            ws.Range("J" & TickerRow).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & TickerRow).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("J" & TickerRow).Interior.ColorIndex = 0
                    End Select
                    
                    ' Check for greatest increase, decrease, and volume
                    If PercentChange > MaxIncrease Then
                        MaxIncrease = PercentChange
                        MaxIncreaseTicker = Stock_Code
                    End If
                    If PercentChange < MaxDecrease Then
                        MaxDecrease = PercentChange
                        MaxDecreaseTicker = Stock_Code
                    End If
                    If TotalVolume > MaxVolume Then
                        MaxVolume = TotalVolume
                        MaxVolumeTicker = Stock_Code
                    End If
                End If
                
                ' Update variables for the next stock
                TickerRow = TickerRow + 1
                Start = i + 1
                TotalVolume = 0
               
            Else
                ' Accumulate total volume for the current stock
                If IsNumeric(ws.Cells(i, 7).Value) Then
                    TotalVolume = TotalVolume + CDbl(ws.Cells(i, 7).Value)
                End If
            End If
        Next i
        
        ' Output the greatest increase, decrease, and volume to specified cells
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(2, 16).Value = MaxIncreaseTicker
        ws.Cells(3, 16).Value = MaxDecreaseTicker
        ws.Cells(4, 16).Value = MaxVolumeTicker
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 17).Value = MaxIncrease
        ws.Cells(3, 17).Value = MaxDecrease
        ws.Cells(4, 17).Value = MaxVolume
    Next ws
    
    ' Format the summary table columns
    For Each ws In Worksheets
        ws.Range("J:L").NumberFormat = "0.00"
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("L:L").NumberFormat = "#,##0"
    Next ws
End Sub
    
        

