Sub StockScript()

For Each ws In Worksheets
    'declare variables
    Dim ticker As String
    Dim totalStockVolume As Variant
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim LastRow As LongLong
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Dim summaryTableRow As Integer
    summaryTableRow = 2
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim daysSinceOpen As Integer
    daysSinceOpen = 0

    'Label Summary Columns for summary table
         
         ws.Range("j1").Value = "Ticker Symbol"
         ws.Range("k1").Value = "Yearly change"
         ws.Range("l1").Value = "Percentage Change"
         ws.Range("m1").Value = "Total Stock Volume"

    For Row = 2 To LastRow
        'Record the ticker symbols in summary table
        
        If (ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value) And ws.Range("C" & Row).Value > 0 Then
        
            'set the ticker symbol
            ticker = ws.Range("A" & Row).Value
            'Record in summary table
            ws.Range("J" & summaryTableRow).Value = ticker
            'Add last volume
            totalStockVolume = totalStockVolume + ws.Range("G" & Row).Value
            'Record summary table
            ws.Range("M" & summaryTableRow).Value = totalStockVolume
            'Get first open price value
            openingPrice = ws.Range("C" & Row - daysSinceOpen).Value
            'Get last closing price value
            closingPrice = ws.Range("F" & Row).Value
            'Calculate yearly change
            yearlyChange = closingPrice - openingPrice
            'Put Yearly change in summary table
            ws.Range("K" & summaryTableRow).Value = yearlyChange
            
            'Calculate Percent Change
            percentChange = (closingPrice - openingPrice) / openingPrice
            'Record percent change in summary table
            ws.Range("L" & summaryTableRow).Value = percentChange
            'change formatting to percentage
            ws.Range("L" & summaryTableRow).NumberFormat = "0.00%"
            'Add conditional formatting
            'Add row to table
            summaryTableRow = summaryTableRow + 1
            'Reset the total stock volume
            totalStockVolume = 0
            daysSinceOpen = 0
        ElseIf (ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value) And ws.Range("C" & Row - daysSinceOpen).Value = 0 Then
            
            ws.Range("M" & summaryTableRow).Value = "XXX"
            summaryTableRow = summaryTableRow + 1
            daysSinceOpen = 0
            totalStockVolume = 0
        'Confirm if still in same ticker symbol
        ElseIf ws.Range("A" & Row).Value = ws.Range("A" & Row + 1).Value And ws.Range("C" & Row).Value = 0 Then
            daysSinceOpen = 0
            totalStockVolume = 0
        Else
            totalStockVolume = totalStockVolume + ws.Range("G" & Row).Value
            daysSinceOpen = daysSinceOpen + 1
            
        End If

    Next Row

    For Row1 = 2 To LastRow

        If ws.Range("K" & Row1).Value > 0 Then
                ws.Range("K" & Row1).Interior.colorIndex = 10
            ElseIf ws.Range("K" & Row1).Value < 0 Then
                ws.Range("K" & Row1).Interior.colorIndex = 3
            End If

    Next Row1

Next ws

End Sub