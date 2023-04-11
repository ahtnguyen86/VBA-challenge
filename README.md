I utilized https://excelchamps.com/vba/sum/

for ws.Cells(tickers, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
    ws.Cells(tickers, 12).NumberFormat = "0"