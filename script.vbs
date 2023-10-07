Sub marketData()

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

' Define last row in the table
Dim ws As Worksheet
Dim lastRow As Long
Set ws = ThisWorkbook.Worksheets("A")
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

'tickerRow count
Dim tickerRow As Integer
tickerRow = 0

Dim i As Integer

' Loop through the first column
For i = 2 To lastRow
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

' Get ticker
ticker = Cells(i, 1).Value

' Get total stock
totalStock = totalStock + Cells(i, 7).Value

' Print the ticker to the summary table
Range("I" & tableRow).Value = ticker

'Print the total stock to the summary table
Range("L" & tableRow).Value = totalStock

'Define and print yearly change and percentage change
closePrice = Cells(i, 6).Value
openPrice = Cells(i - tickerRow, 3).Value
yearlyChange = closePrice - openPrice
Range("J" & tableRow).Value = yearlyChange
percentChange = yearlyChange / openPrice
Range("K" & tableRow).Value = percentChange
Range("K" & tableRow).NumberFormat = "0.00%"

If yearlyChange >= 0 Then
Range("J" & tableRow).Interior.ColorIndex = 4
Else
Range("J" & tableRow).Interior.ColorIndex = 3
End If
tableRow = tableRow + 1

' Reset the total stock and tickerRow
totalStock = 0
tickerRow = 0

Else
totalStock = totalStock + Cells(i, 7).Value
tickerRow = tickerRow + 1

End If

Next i

End Sub
