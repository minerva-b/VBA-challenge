'HOMEWORK INSTRUCTIONS:

    'Create a script that will loop through all the stocks for one year and output the following information.

    '- The ticker symbol
    '- Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    '- The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    '- The total stock volume of the stock

    'You should also have conditional formatting that will highlight positive change in green and negative change in red.

'-------------------------------------------------------------------------------------------------------------------------------

'CHALLENGES:

    '-1- Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
    '-2- Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.
    
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Other Considerations:

    'Use the sheet alphabetical_testing.xlsx while developing your code. This data set is smaller and will allow you to test faster. Your code should run on this file in less than 3-5 minutes.
    'Make sure that the script acts the same on each sheet. The joy of VBA is to take the tediousness out of repetitive task and run over and over again with a click of the button.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub Stock_Data()

    For Each ws In Worksheets
        ws.Activate
        
        Dim Ticker As String
        Dim Yearly_Change As Double
        Yearly_Change = ws.Cells(2, 3).Value
        Dim Percent_Change As Double
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        Dim Output As Integer
        Output = 2
        Dim LastRow
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                
            For i = 2 To LastRow
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    Ticker = ws.Cells(i, 1).Value
                    If Yearly_Change = 0 Then
                        Percent_Change = 0
                                            
                    Else
                        Percent_Change = (ws.Cells(i, 6).Value - Yearly_Change) / Yearly_Change
                    End If
                    
                    Yearly_Change = ws.Cells(i, 6).Value - Yearly_Change
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                    ws.Range("I" & Output).Value = Ticker
                    ws.Range("J" & Output).Value = Yearly_Change
                    ws.Range("K" & Output).Value = Percent_Change
                    ws.Range("L" & Output).Value = Total_Stock_Volume
                
                    Output = Output + 1
                    Total_Stock_Volume = 0
                    Yearly_Change = ws.Cells(i + 1, 3)
            
                Else
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                End If
            
            Next i
        
                For i = 2 To LastRow
                    For j = 10 To 10
                        If ws.Cells(i, j).Value > 0 Then
                        ws.Cells(i, 10).Interior.ColorIndex = 4
                
                        ElseIf ws.Cells(i, j).Value < 0 Then
                        ws.Cells(i, 10).Interior.ColorIndex = 3
                
                        End If
                    Next j
                Next i
            ws.Range("K:K").NumberFormat = "0.00%"
            
            'NEW PART ADDED FOR CHALLENGE 1
            Dim Greatest_Percent_Increase As Double
            Dim Greatest_Percent_Decrease As Double
            Dim Greatest_Total_Volume As Double
            
            ws.Range("N2").Value = "Greatest % Increase"
            ws.Range("N3").Value = "Greatest % Decrease"
            ws.Range("N4").Value = "Greatest Total Volume"
            ws.Range("O1").Value = "Ticker"
            ws.Range("P1").Value = "Value"
            
            Dim my_range As Range
            Dim my_range2 As Range
            Dim LastRow2
            
            Set my_range = ws.Range("K:K")
            Set my_range2 = ws.Range("L:L")
            LastRow2 = ws.Cells(Rows.Count, 1).End(xlUp).Row
            max_value = Application.WorksheetFunction.Max(my_range)
            min_value = Application.WorksheetFunction.Min(my_range)
            max_value2 = Application.WorksheetFunction.Max(my_range2)
                        
            For i = 2 To LastRow2
                If ws.Cells(i, 11).Value = max_value Then
                    ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
                    ws.Cells(2, 16).Value = ws.Cells(i, 11).Value
                End If
                
                If ws.Cells(i, 11).Value = min_value Then
                    ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
                    ws.Cells(3, 16).Value = ws.Cells(i, 11).Value
                End If
                If ws.Cells(i, 12).Value = max_value2 Then
                    ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
                    ws.Cells(4, 16).Value = ws.Cells(i, 12).Value
                End If

            Next i
        ws.Range("P2:P3").NumberFormat = "0.00%"
            
    Next ws
End Sub