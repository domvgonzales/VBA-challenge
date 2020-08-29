Sub WorksheetLoop()
        Dim Ticker As String
        Dim Yearly_Change As Double
            Yearly_Change = 0
        Dim Percent_Change As Double
            Percent_Change = 0
        Dim Total_Stock_Volume As Double
            Total_Stock_Volume = 0
        Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2
        Dim counter As Integer
            counter = 0
        Dim ann_change As Variant
        Dim rng As Range
        
            
        For i = 2 To 797711
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value
            
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
            fin_price = Cells(i, 6).Value
            in_price = Cells(i - counter, 6).Value
            ann_change = fin_price - in_price
            Yearly_Change = fin_price - in_price
            
            
        If in_price <> 0 Then
                    Percent_Change = CDec(ann_change) / in_price
        End If
        
            Range("I" & Summary_Table_Row).Value = Ticker
            Range("J" & Summary_Table_Row).Value = Yearly_Change
            Range("K" & Summary_Table_Row).Value = Percent_Change
            Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            Total_Stock_Volume = 0
        
            counter = 0
        Else
        
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
            counter = counter + 1
        End If
        
        If Cells(i, 10).Value < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        Else
            Cells(i, 10).Interior.ColorIndex = 4
        End If
        
        Next i

    
End Sub

