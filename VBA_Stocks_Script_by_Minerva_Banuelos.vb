'HOMEWORK INSTRUCTIONS:

    'Create a script that will loop through all the stocks for one year and output the following information.

    '- The ticker symbol
    '- Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    '- The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    '- The total stock volume of the stock

    'You should also have conditional formatting that will highlight positive change in green and negative change in red.

Sub Stock_Data()

        Dim Ticker As String
        Dim Yearly_Change As Double
        Yearly_Change = Cells(2, 3).Value
        Dim Percent_Change As Double
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        Dim Output As Integer
        Output = 2
        Dim LastRow
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
                
            For i = 2 To LastRow
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    Ticker = Cells(i, 1).Value
                    If Yearly_Change = 0 Then
                        Percent_Change = 0
                                            
                    Else
                        Percent_Change = (Cells(i, 6).Value - Yearly_Change) / Yearly_Change
                    End If
                    
                    Yearly_Change = Cells(i, 6).Value - Yearly_Change
                    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                
                    Range("I" & Output).Value = Ticker
                    Range("J" & Output).Value = Yearly_Change
                    Range("K" & Output).Value = Percent_Change
                    Range("L" & Output).Value = Total_Stock_Volume
                
                    Output = Output + 1
                    Total_Stock_Volume = 0
                    Yearly_Change = Cells(i + 1, 3)
            
                Else
                    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                
                End If
            
            Next i
        
            For i = 2 To LastRow
                For j = 10 To 10
                    If Cells(i, j).Value > 0 Then
                        Cells(i, 10).Interior.ColorIndex = 4
                
                    ElseIf Cells(i, j).Value < 0 Then
                        Cells(i, 10).Interior.ColorIndex = 3
                
                    End If
                Next j
            Next i
        Range("K:K").NumberFormat = "0.00%"
End Sub

