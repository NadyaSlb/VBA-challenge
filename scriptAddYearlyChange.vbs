Sub addYearlyChange()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

' Set headers and format cells
ws.Range("J1").Value = "Yearly Change"
ws.Columns("J").AutoFit

' Define last row in the table
Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

' Set initial variables
Dim openPrice As Double
Dim closePrice As Double
Dim yearlyChange As Double

' Keep track of the location for each ticker in the summary table
Dim tableRow As Integer
tableRow = 2

' TickerRow count
Dim tickerRow As Integer
tickerRow = 0

Dim i As Long

' Loop through the first column
For i = 2 To lastRow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Define and print yearly change
closePrice = ws.Cells(i, 6).Value
openPrice = ws.Cells(i - tickerRow, 3).Value
yearlyChange = closePrice - openPrice
ws.Range("J" & tableRow).Value = yearlyChange
ws.Range("J" & tableRow).NumberFormat = "0.00"

If yearlyChange >= 0 Then
ws.Range("J" & tableRow).Interior.ColorIndex = 4
Else
ws.Range("J" & tableRow).Interior.ColorIndex = 3
End If
tableRow = tableRow + 1

tickerRow = 0

Else
tickerRow = tickerRow + 1

End If

Next i

Next ws

End Sub

