Sub addTotalStock()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

 'Set headers and format cells
ws.Range("L1").Value = "Total Stock Volume"
ws.Columns("L").AutoFit

' Define last row in the table
Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

' Set the initial variable for holding total stock
Dim totalStock As Double
totalStock = 0

 ' Keep track of the location for each ticker in the summary table
Dim tableRow As Integer
tableRow = 2

Dim i As Long

' Loop through the first column
For i = 2 To lastRow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

' Get total stock
totalStock = totalStock + ws.Cells(i, 7).Value

'Print the total stock to the summary table
ws.Range("L" & tableRow).Value = totalStock

tableRow = tableRow + 1

' Reset the total stock and tickerRow
totalStock = 0

Else
totalStock = totalStock + ws.Cells(i, 7).Value

End If

Next i

Next ws
End Sub
