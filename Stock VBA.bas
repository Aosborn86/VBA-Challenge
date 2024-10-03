Attribute VB_Name = "Module1"
Sub reset_file(): 'Resets all sheets to pre-analysis state
    Dim i As Integer
    'Loop to cycle through all workbook sheets and delete columns I through Q - This also resets formating
    For i = 1 To Sheets.Count
        With Sheets(i)
            .Columns("I:Q").Delete
        End With
    Next i
End Sub
Sub ProcessStockData()

    Dim i As Long
    Dim j As Long
    Dim Ticker As String
    Dim Total_Volume As Double
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Quarterly_Change As Double
    Dim Price1 As Double
    Dim Price2 As Double
    Dim Percentage_change As Double
    Dim Summary_table_Row As Long
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Summary_table_Row = 2
        Total_Volume = 0
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        ' Loop through each row of the worksheet
        For i = 2 To LastRow
            ' If we are at the last row or the next ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                Price2 = ws.Cells(i, 3).Value
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                
                ' Calculate Quarterly Change
                If i > 2 Then ' Ensure there's a previous price
                    Quarterly_Change = Price2 - Price1
                End If
                
                ' Output results
                ws.Cells(Summary_table_Row, 9).Value = Ticker ' Column I
                ws.Cells(Summary_table_Row, 10).Value = Quarterly_Change ' Column J
                ws.Cells(Summary_table_Row, 12).Value = Total_Volume ' Column K
                
                ' Calculate Percentage Change
                If Price1 <> 0 Then
                    Percentage_change = Round((Price2 / Price1 - 1) * 100, 2)
                    ws.Cells(Summary_table_Row, 11).Value = Percentage_change ' Column L
                    ws.Cells(Summary_table_Row, 11).NumberFormat = "0.00%" ' Format for Percentage Change
                End If
                
                ' Conditional Formatting
                Select Case Quarterly_Change
                    Case Is > 0
                        ws.Cells(Summary_table_Row, 10).Interior.ColorIndex = 4 ' Green
                    Case Is < 0
                        ws.Cells(Summary_table_Row, 10).Interior.ColorIndex = 3 ' Red
                    Case Else
                        ws.Cells(Summary_table_Row, 10).Interior.ColorIndex = 0 ' No color
                End Select
                
                ' Move to the next summary row
                Summary_table_Row = Summary_table_Row + 1
                Total_Volume = 0 ' Reset for next ticker
            Else
                ' Accumulate total volume
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            End If
            
            ' Set Price1 for the next iteration
            Price1 = ws.Cells(i, 3).Value
            
            

        Next i
            ws.Columns("I:Q").AutoFit
    
    Next ws

End Sub

