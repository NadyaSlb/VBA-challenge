Sub addPercentageChange()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

 ' Set headers and format cells
ws.Range("K1").Value = "Percent Change"
ws.Columns("K").AutoFit

' Define last row in the table
Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

 ' Set initial variables
Dim openPrice As Double
Dim closePrice As Double
Dim yearlyChange As Double
Dim percentChange As Double

 ' Keep track of the location for each ticker in the summary table
Dim tableRow As Integer
tableRow = 2

'tickerRow count
Dim tickerRow As Integer
tickerRow = 0

Dim i As Long

' Loop through the first column
For i = 2 To lastRow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Define and print yearly change and percentage change
closePrice = ws.Cells(i, 6).Value
openPrice = ws.Cells(i - tickerRow, 3).Value
yearlyChange = closePrice - openPrice
percentChange = yearlyChange / openPrice
ws.Range("K" & tableRow).Value = percentChange
ws.Range("K" & tableRow).NumberFormat = "0.00%"
If percentChange >= 0 Then
ws.Range("K" & tableRow).Interior.ColorIndex = 4
Else
ws.Range("K" & tableRow).Interior.ColorIndex = 3
End If
tableRow = tableRow + 1

tickerRow = 0

Else
tickerRow = tickerRow + 1

End If

Next i

Next ws

End Sub
