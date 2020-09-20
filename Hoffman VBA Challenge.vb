Sub All_Data()

'create variables for printing values
Dim TickerName As String
Dim TotalVolume As Double

'create variables for change calculations
Dim stock_open As Double
Dim stock_close As Double
Dim stock_change As Double
Dim percent_change As Double

'create variables for moving through sheet/workbook
Dim summaryrow As Integer
Dim ws As Worksheet

'loop through the worksheets
For Each ws In Worksheets

'insert column headings to right of data
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly_Change"
ws.Cells(1, 11).Value = "Percent_Change"
ws.Cells(1, 12).Value = "Total_Stock_Volume"

'determine last row
 LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'format Percent_Change as percent
    For I = 2 To LastRow
        ws.Cells(I, 11).Style = "Percent"
    Next I

'Determine LastRow
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'start count for summaryrow (printing in the summary table)
summaryrow = 2

        For j = 2 To LastRow
        
        If ws.Cells(j - 1, 1).Value <> ws.Cells(j, 1).Value Then
            
            'hold stock_open as first value of a ticker
            stock_open = ws.Cells(j, 3).Value
        
        End If
                
        If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
           
            'Hold Ticker Name
            TickerName = ws.Cells(j, 1).Value
            
            'Hold total volume
            TotalVolume = TotalVolume + ws.Cells(j, 7).Value
            
            'Hold stock_close
            stock_close = ws.Cells(j, 6).Value
                
            'Print Ticker in column
            ws.Range("I" & summaryrow).Value = TickerName
            'Print total volume in colum
            ws.Range("L" & summaryrow).Value = TotalVolume
        
            'print Stock Change data
            stock_change = stock_close - stock_open
            ws.Range("J" & summaryrow).Value = stock_change
                If stock_change > 0 Then
                    ws.Range("J" & summaryrow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & summaryrow).Interior.ColorIndex = 3
                End If
                
            'calculate and print percent change
                If stock_open <> 0 Then
                    percent_change = stock_change / stock_open
                    ws.Range("K" & summaryrow).Value = percent_change
                End If
                            
            'Reset Total Volume
            TotalVolume = 0
        
            'Add to summary row
            summaryrow = summaryrow + 1
        
        Else
            'Add to Total Volume
            TotalVolume = TotalVolume + ws.Cells(j, 7).Value
    
        End If

    Next j

Next ws

End Sub