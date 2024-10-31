VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentTicker As String
    Dim startRow As Long
    Dim endRow As Long
    Dim quarterStartPrice As Double
    Dim quarterEndPrice As Double
    Dim totalVolume As Double
    Dim currentQuarter As Integer
    Dim currentDate As Date
    Dim outputRow As Long
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        startRow = 2
        currentTicker = ws.Cells(startRow, 1).Value
        currentQuarter = DatePart("q", ws.Cells(startRow, 2).Value)
        quarterStartPrice = ws.Cells(startRow, 3).Value
        totalVolume = 0
        
        ' Output headers for the current sheet
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Quarterly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Volume"
        outputRow = 2
        
        For i = 2 To lastRow
            currentDate = ws.Cells(i, 2).Value
            If ws.Cells(i, 1).Value <> currentTicker Or DatePart("q", currentDate) <> currentQuarter Then
                ' Calculate end of quarter
                endRow = i - 1
                quarterEndPrice = ws.Cells(endRow, 6).Value
                
                ' Calculate percent change
                Dim percentChange As Double
                percentChange = (quarterEndPrice - quarterStartPrice) / quarterStartPrice * 100
                
                ' Output results to the current sheet
                ws.Cells(outputRow, 10).Value = currentTicker
                ws.Cells(outputRow, 11).Value = quarterEndPrice - quarterStartPrice
                ws.Cells(outputRow, 12).Value = percentChange
                ws.Cells(outputRow, 13).Value = totalVolume
                outputRow = outputRow + 1
                
                ' Reset for next quarter or ticker
                currentTicker = ws.Cells(i, 1).Value
                currentQuarter = DatePart("q", currentDate)
                startRow = i
                quarterStartPrice = ws.Cells(startRow, 3).Value
                totalVolume = 0
            End If
            
            ' Accumulate volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        Next i
        
        ' Output last quarter
        quarterEndPrice = ws.Cells(lastRow, 6).Value
        Dim finalPercentChange As Double
        finalPercentChange = (quarterEndPrice - quarterStartPrice) / quarterStartPrice * 100
        ws.Cells(outputRow, 10).Value = currentTicker
        ws.Cells(outputRow, 11).Value = quarterEndPrice - quarterStartPrice
        ws.Cells(outputRow, 12).Value = finalPercentChange
        ws.Cells(outputRow, 13).Value = totalVolume
    Next ws
End Sub

