Sub marketData()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

 'Set headers and format cells
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"
ws.Columns("J").AutoFit
ws.Columns("K").AutoFit
ws.Columns("L").AutoFit
ws.Columns("N").AutoFit


' Define last row in the table
Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

 ' Set initial variables
Dim ticker As String
Dim openPrice As Double
Dim closePrice As Double
Dim yearlyChange As Double
Dim percentChange As Double

' Set the initial variable for holding total stock
Dim totalStock As Double
totalStock = 0

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

' Get ticker
ticker = ws.Cells(i, 1).Value

' Get total stock
totalStock = totalStock + ws.Cells(i, 7).Value

' Print the ticker to the summary table
ws.Range("I" & tableRow).Value = ticker

'Print the total stock to the summary table
ws.Range("L" & tableRow).Value = totalStock

'Define and print yearly change and percentage change
closePrice = ws.Cells(i, 6).Value
openPrice = ws.Cells(i - tickerRow, 3).Value
yearlyChange = closePrice - openPrice
ws.Range("J" & tableRow).Value = yearlyChange
ws.Range("J" & tableRow).NumberFormat = "0.00"
percentChange = yearlyChange / openPrice
ws.Range("K" & tableRow).Value = percentChange
ws.Range("K" & tableRow).NumberFormat = "0.00%"

If yearlyChange >= 0 Then
ws.Range("J" & tableRow).Interior.ColorIndex = 4
Else
ws.Range("J" & tableRow).Interior.ColorIndex = 3
End If

If percentChange >= 0 Then
ws.Range("K" & tableRow).Interior.ColorIndex = 4
Else
ws.Range("K" & tableRow).Interior.ColorIndex = 3
End If

tableRow = tableRow + 1

' Reset the total stock and tickerRow
totalStock = 0
tickerRow = 0

Else
totalStock = totalStock + ws.Cells(i, 7).Value
tickerRow = tickerRow + 1

End If

Next i

'Functionality to return greatest percentage increase
Dim maxIncrease As Double
Dim incRow As Long
Dim incTicker As String
maxIncrease = WorksheetFunction.Max(ws.Range("K2:K" & lastRow))
ws.Range("P2").Value = maxIncrease
ws.Range("P2").NumberFormat = "0.00%"
incRow = WorksheetFunction.Match(maxIncrease, ws.Range("K2:K" & lastRow), 0)
incTicker = ws.Range("I" & incRow + 1).Value
ws.Range("O2").Value = incTicker

'Functionality to return greatest percentage decrease
Dim maxDecrease As Double
Dim decRow As Long
Dim decTicker As String
maxDecrease = WorksheetFunction.Min(ws.Range("K2:K" & lastRow))
ws.Range("P3").Value = maxDecrease
ws.Range("P3").NumberFormat = "0.00%"
decRow = WorksheetFunction.Match(maxDecrease, ws.Range("K2:K" & lastRow), 0)
decTicker = ws.Range("I" & decRow + 1).Value
ws.Range("O3").Value = decTicker

'Functionality to return greatest percentage total volume
Dim maxVolume As Double
Dim volRow As Long
Dim volTicker As String
maxVolume = WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
ws.Range("P4").Value = maxVolume
'Range("P4").NumberFormat = "0.00%"
volRow = WorksheetFunction.Match(maxVolume, ws.Range("L2:L" & lastRow), 0)
volTicker = ws.Range("I" & volRow + 1).Value
ws.Range("O4").Value = volTicker
Next ws
End Sub