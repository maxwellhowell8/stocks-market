Sub stocks()
Dim total As Double
Dim i As Long
Dim change As Double
Dim j As Integer
Dim start As Long
Dim rowCount As Long
Dim percentChange As Double
Dim days As Integer
Dim dailyChange As Double
Dim averageChange As Double
j = 0
total = 0
change = 0
start = 2
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
Last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To Last_row
'sum of volume
   If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
       volume = volume + ws.Cells(i, 7).Value
       ticker = ws.Cells(i, 1).Value
       If total = 0 Then
               ' print the results
               Range("I" & 2 + j).Value = Cells(i, 1).Value
               Range("J" & 2 + j).Value = 0
               Range("K" & 2 + j).Value = "%" & 0
               Range("L" & 2 + j).Value = 0
           Else
               ' Find First non zero starting value
               If Cells(start, 3) = 0 Then
                   For find_value = start To i
                       If Cells(find_value, 3).Value <> 0 Then
                           start = find_value
                           Exit For
                       End If
                    Next find_value
               End If
               ' Calculate Change
               change = (Cells(i, 6) - Cells(start, 3))
               percentChange = Round((change / Cells(start, 3) * 100), 2)
               ' start of the next stock ticker
               start = i + 1
               ' print the results
               Range("I" & 2 + j).Value = Cells(i, 1).Value
               Range("J" & 2 + j).Value = Round(change, 2)
               Range("K" & 2 + j).Value = "%" & percentChange
               Range("L" & 2 + j).Value = total
               ' colors positives green and negatives red
               Select Case change
                   Case Is > 0
                       Range("J" & 2 + j).Interior.ColorIndex = 4
                   Case Is < 0
                       Range("J" & 2 + j).Interior.ColorIndex = 3
                   Case Else
                       Range("J" & 2 + j).Interior.ColorIndex = 0
               End Select
           End If
           ' reset variables for new stock ticker
           total = 0
           change = 0
           j = j + 1
           days = 0
       ' If ticker is still the same add results
       Else
           total = total + Cells(i, 7).Value
       End If
   Next i
Next ws

End Sub

       





