Attribute VB_Name = "Module1"
Sub stocks()


Dim Ticker_Row As Long
Dim Stock_Volume As Double
Dim LastRow As Long
Dim ws As Worksheet


Ticker_Row = 2
Stock_Volume = 0


For Each ws In Worksheets
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Opening_Price = ws.Cells(2, 3).Value
        
        'Loop through all rows in the worksheet
    For i = 2 To LastRow
        
            ' start adding up the stock volume in column G
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
    
            'if at anypoint we sense a new ticker code, do the following:
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                'Print the ticker code and volume total in Columns I and L
                
                ws.Cells(Ticker_Row, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(Ticker_Row, 12).Value = Stock_Volume
                
                Closing_Price = ws.Cells(i, 6).Value
                
                Yearly_Change = Closing_Price - Opening_Price
                
                ws.Cells(Ticker_Row, 10) = Yearly_Change
                
                    If (Opening_Price > 0) Then
                        Percentage_Change = Yearly_Change / Opening_Price
        
                    Else
                        Percentage_Change = 0
        
                    End If
                
                ws.Cells(Ticker_Row, 11) = Percentage_Change
                
                    If Yearly_Change > 0 Then
                        ws.Cells(Ticker_Row, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(Ticker_Row, 10).Interior.ColorIndex = 3
                    End If
                
                Opening_Price = ws.Cells(i + 1, 3).Value
                Stock_Volume = 0
                Ticker_Row = Ticker_Row + 1
        
        End If
        
    Next i
        
Next ws
    

    
    
End Sub
