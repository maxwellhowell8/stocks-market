Sub stocks()
Dim ticker As String
Dim open_value, close_value, stock_volume, percent_change, diff  As Double
Dim start, j As Integer




Dim ws As Worksheet


'Loop through all stocks for one year
For Each ws In Worksheets


'Create the column headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Defining
start = 2
volume = 0
j = 2


Last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To Last_row

'sum of volume
If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

volume = volume + ws.Cells(i, 7).Value
ticker = ws.Cells(i, 1).Value
ws.Cells(j, 9).Value = ticker
ws.Cells(j, 12).Value = volume

' Yearly Change

open_value = ws.Cells(start, 3).Value
close_value = ws.Cells(i, 6).Value

'percent change

diff= close_value - open_value
ws.Range("J" & j).Value = diff
ws.Range("K" & j).Value = diff





start = i + 1
volume = 0
j = j + 1


Else

volume = volume + ws.Cells(i, 7).Value


End If




Next i

Next ws


End Sub






