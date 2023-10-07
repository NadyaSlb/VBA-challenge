Sub getTicker()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

' Set headers and format cells
ws.Range("I1").Value = "Ticker"

' Define last row in the table
Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

 ' Set initial variables
Dim ticker As String

 ' Keep track of the location for each ticker in the summary table
Dim tableRow As Integer
tableRow = 2

Dim i As Long

' Loop through the first column
For i = 2 To lastRow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

' Get ticker
ticker = ws.Cells(i, 1).Value

' Print the ticker to the summary table
ws.Range("I" & tableRow).Value = ticker

tableRow = tableRow + 1

End If

Next i

Next ws

End Sub