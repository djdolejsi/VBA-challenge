Sub Stock_Exchange()

Dim ws As Worksheet
Dim Ticker_Name As String
Dim Ticker_Volume As Double
Dim Ticker_Summary As Integer
Ticker_Volume = 0

For Each ws In Worksheets

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percentage Change"
ws.Range("L1").Value = "Total Stock Volume"

Ticker_Summary = 2

lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1

For i = 2 To lastRow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
	Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
	Ticker_Name = ws.Cells(i, 1).Value
	Closing_Value = ws.Cells(i,6).Value
	Yearly_Change = Closing_Value - Opening_Value
	Percentage_Change = Yearly_Change / Opening_Value

ws.Range("I" & Ticker_Summary).Value = Ticker_Name
ws.Range("J" & Ticker_Summary).Value = Yearly_Change
ws.Range("K" & Ticker_Summary).Value = Percentage_Change
ws.Range("L" & Ticker_Summary).Value = Ticker_Volume

Ticker_Summary = Ticker_Summary + 1

ElseIf ws.Cells(i-1,1).Value <> ws.Cells(i,1).Value Then
	Opening_Value = ws.Cells(i, 3).Value

Ticker_Volume = 0

Else

Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value

End If

Next i

lastRow2 = ws.Cells(Rows.Count, "J").End(xlUp).Row + 1

For j = 2 To lastRow2

ws.Cells(j,11).NumberFormat = "0.00%"

If ws.Cells(j, 10).Value > 0 Then
    ws.Cells(j, 10).Interior.ColorIndex = 4

ElseIf ws.Cells(j, 10).Value < 0 Then
    ws.Cells(j, 10).Interior.ColorIndex = 3

End If

Next j

Next ws

End Sub

