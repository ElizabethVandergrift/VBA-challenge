# VBA-challenge
Assignment Files
Sub Stock_Market()

    'Create Variable for Worksheets
    
    Dim Current As Worksheet

    'Indentify Last Row

    For Each Current In Worksheets

        Dim LR As Long

        LR = Cells(Rows.Count, 1).End(xlUp).Row
    
        'Set an initial variable for holding the stock abbreviation
        Dim Stock_Code As String
    
        'Set an initial variable for holding the change in price, Date, Open and Close Prices
        
        Dim Change As Double
        Dim Total_Vol As Long
        Dim i As Long
        Dim j As Long
        Dim percentChange
        
        Current.Range("I1").Value = "Ticker"
        Current.Range("J1").Value = "Change"
        Current.Range("K1").Value = "Percent"
        Current.Range("L1").Value = "Total_Volume"
        
        'Give values to variables
        Change = 0
        Total_Vol = 0
        Start = 2
        
        'Keep track of the location for each stock code in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'Loop through all days of the year for stock outcomes
        For i = 2 To LR
            
            'Check if we are still within the same stock abbreviation, if not...
            If Current.Cells(i + 1, 1).Value <> Current.Cells(i, 1).Value Then
                
                'Set the Stock Abbreviation
                Stock_Code = Current.Cells(i, 1).Value

                'Print the Stock Abbreviation in the Summary Table
                Current.Range("I" & Summary_Table_Row).Value = Stock_Code
            
                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
            
            Else
                
            End If
            
            
        Next i
        
        
        
Next Current

End Sub
