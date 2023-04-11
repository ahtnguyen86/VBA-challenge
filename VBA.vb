
Sub tickerclean()

For Each ws In Worksheets
    Dim WorksheetName As String
        
    'find last row number
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    WorksheetName = ws.Name

Dim tickers As String
Dim yearlychange As Double
Dim percentchange As Double
Dim totalstock As Long

Dim headers(7) As String
headers(0) = "Ticker"
headers(1) = "Yearly Change"
headers(2) = "Percent Change"
headers(3) = "Total Stock Volume"
headers(4) = "Value"
headers(5) = "Greatest % Increase"
headers(6) = "Greatest % Decrease"
headers(7) = "Greatest Total Volume"

ws.Range("I1").Value = headers(0)
ws.Range("J1").Value = headers(1)
ws.Range("K1").Value = headers(2)
ws.Range("L1").Value = headers(3)
ws.Range("O2").Value = headers(5)
ws.Range("O3").Value = headers(6)
ws.Range("O4").Value = headers(7)
ws.Range("P1").Value = headers(0)
ws.Range("Q1").Value = headers(4)


tickers = 2

j = 2
For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
    ws.Cells(tickers, 9).Value = ws.Cells(i, 1).Value
    ws.Cells(tickers, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
    If ws.Cells(tickers, 10).Value < 0 Then
    
    ws.Cells(tickers, 10).Interior.ColorIndex = 3
    Else
    ws.Cells(tickers, 10).Interior.ColorIndex = 4
    End If

    If ws.Cells(j, 3).Value <> 0 Then
    percentchange = (ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value

    ws.Cells(tickers, 11).Value = percentchange
    ws.Cells(tickers, 11).NumberFormat = "0.00%"

    Else
    ws.Cells(tickers, 11).Value = "0"
    ws.Cells(tickers, 11).NumberFormat = "0.00%"
    End If

    ws.Cells(tickers, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
    ws.Cells(tickers, 12).NumberFormat = "0"

    tickers = tickers + 1

    j = i + 1
    End If

Next i


lastrowt = ws.Cells(Rows.Count, 9).End(xlUp).Row

Greatvolume = ws.Cells(2, 12).Value
Greatincrease = ws.Cells(2, 11).Value
Greatdecrease = ws.Cells(2, 11).Value

For i = 2 To lastrowt

If ws.Cells(i, 12).Value > Greatvolume Then

Greatvolume = ws.Cells(i, 12).Value
ws.Cells(4, 16).Value = ws.Cells(i, 9).Value

Else

Greatvolume = Greatvolume

End If

If ws.Cells(i, 11).Value > Greatincrease Then
Greatincrease = ws.Cells(i, 11).Value
ws.Cells(2, 16).Value = ws.Cells(i, 9).Value

Else

Greatincrease = Greatincrease

End If

If ws.Cells(i, 11).Value < Greatdecrease Then
Greatdecrease = ws.Cells(i, 11).Value
ws.Cells(3, 16).Value = ws.Cells(i, 9).Value

Else

Greatdecrease = Greatdecrease

End If
ws.Cells(2, 17).Value = Greatincrease
ws.Cells(3, 17).Value = Greatdecrease
ws.Cells(4, 17).Value = Greatvolume



ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("Q4").NumberFormat = "0.00E+00"

Next i

ws.Columns("A:Q").AutoFit

Next ws

End Sub


