Sub VBAHomework():

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        
        Dim LastRow1 As Long
            LastRow1 = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
        Dim Ticker As String
        
        Dim Total_Stock As Double
            Total_Stock = 0
        
        Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2
            
        Dim Opening_Price As Double
            Opening_Price = 0
          
        Dim Closing_Price As Double
            Closing_Price = 0
              
        Dim Yearly_Change As Double
            Yearly_Change = 0
        
        Dim Percent_Change As Double
            Percent_Change = 0
        
        WorksheetName = ws.Name
            
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        Opening_Price = ws.Cells(2, 3).Value
        
        For i = 2 To LastRow1
         
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                Ticker = ws.Cells(i, 1).Value
                
                Closing_Price = ws.Cells(i, 6).Value
                
                Yearly_Change = Closing_Price - Opening_Price
                
                If Opening_Price <> 0 Then
                
                    Percent_Change = (Yearly_Change / Opening_Price) * 100
                    
                End If
                
                Total_Stock = Total_Stock + ws.Cells(i, 7).Value
                   
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                    If Yearly_Change > 0 Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        
                    ElseIf Yearly_Change <= 0 Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    
                    End If
                
                ws.Range("K" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")
                
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock
                
                Summary_Table_Row = Summary_Table_Row + 1
                
                Closing_Price = 0
                
                Yearly_Change = 0
                
                Percent_Change = 0
                    
                Opening_Price = ws.Cells(i + 1, 3).Value
          
            Else
            
                Total_Stock = Total_Stock + ws.Cells(i, 7).Value
                
            End If
            
        Next i
                 
    Next ws
   
End Sub

