Attribute VB_Name = "Module1"
Sub stock_market()

    For Each ws In Worksheets

        Dim Stock_Name As String
    
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
        Dim opening_price As Double
        opening_price = ws.Cells(2, 3).Value
    
        Dim closing_price As Double
    
        Dim yearly_change As Double
    
        Dim percent_change As Double
    
        Dim Stock_Vol As Double
        Stock_Vol = 0
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastrow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                Stock_Name = ws.Cells(i, 1).Value
                
                closing_price = ws.Cells(i, 6).Value
                
                yearly_change = (closing_price - opening_price)
                
                If (opening_price = 0) Then
                
                    percent_change = 0
                    
                Else
                    
                    percent_change = (yearly_change / opening_price)
                    
                End If
                            
                Stock_Vol = Stock_Vol + Cells(i, 7).Value
          
                ws.Range("I" & Summary_Table_Row).Value = Stock_Name
                
                ws.Range("J" & Summary_Table_Row).Value = yearly_change
                
                ws.Range("K" & Summary_Table_Row).Value = percent_change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                ws.Range("L" & Summary_Table_Row).Value = Stock_Vol
    
                Summary_Table_Row = Summary_Table_Row + 1
                
                opening_price = ws.Cells(i + 1, 3)
                
                Stock_Vol = 0
    
            Else
            
                Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value
    
            End If
        
        Next i
        
        LastRow_Yearly_Change = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For j = 2 To LastRow_Yearly_Change
        
            If ws.Cells(j, 10).Value > 0 Then
            
                ws.Cells(j, 10).Interior.ColorIndex = 4
            
            Else
                
                ws.Cells(j, 10).Interior.ColorIndex = 3
            
            End If
        
        Next j
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        For k = 2 To LastRow_Yearly_Change
        
            If ws.Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRow_Yearly_Change)) Then
            
                ws.Cells(2, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(k, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"
            
            End If
            
            If ws.Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRow_Yearly_Change)) Then
            
                ws.Cells(3, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(k, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            End If
            
            If ws.Cells(k, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRow_Yearly_Change)) Then
            
                ws.Cells(4, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(k, 12).Value
            
            End If
        
        Next k
    
    Next ws

End Sub
