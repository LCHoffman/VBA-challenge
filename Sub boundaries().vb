Sub boundaries()
Dim ws As Worksheet
Dim rngP As Range
Dim rngT As Range
Dim Total As Double
Dim perMax As Double
Dim perMin As Double
Dim TotMax As Double

For Each ws In Worksheets

'print headings
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"

'Set range variables
Set rngP = ws.Range("K:K")
Set rngT = ws.Range("L:L")

'Find the max and print increase
perMax = Application.WorksheetFunction.max(rngP)
ws.Cells(2, 15).Value = perMax
ws.Cells(2, 15).Style = "Percent"

'Find the min and print decrease
perMin = Application.WorksheetFunction.Min(rngP)
ws.Cells(3, 15).Value = perMin
ws.Cells(3, 15).Style = "Percent"

'Find and print max Volume
TotMax = Application.WorksheetFunction.max(rngT)
ws.Cells(4, 15).Value = TotMax

Next ws

End Sub
