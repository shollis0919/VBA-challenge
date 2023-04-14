Attribute VB_Name = "Module1"
Sub stockExchange()

For Each ws In Worksheets

Dim ticker As String
Dim volumeSum As Double
    volumeSum = 0
Dim openPrice As Double
Dim closePrice As Double
Dim yearChange As Double
Dim percentChange As Double
Dim sumTableRow As Integer
    sumTableRow = 1

ws.Cells(1, 9).Value = "Ticker"
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest % Total Voilume"
For i = 2 To lastRow
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        sumTableRow = sumTableRow + 1
        ticker = ws.Cells(i, 1).Value
        ws.Range("I" & sumTableRow).Value = ticker
        volumeSum = 0
    End If
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value And ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
        openPrice = ws.Cells(i, 3).Value
        
    End If
    If ws.Cells(i - 1, 1).Value = ws.Cells(i, 1).Value And ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        closePrice = ws.Cells(i, 6).Value
    End If
    
    yearChange = closePrice - openPrice
    ws.Range("J" & sumTableRow).Value = yearChange
    ws.Range("J" & sumTableRow).NumberFormat = "0.##"
    
    percentChange = ((yearChange / openPrice))
    ws.Range("K" & sumTableRow).Value = percentChange
    ws.Range("K" & sumTableRow).NumberFormat = ".##%"
    
    volumeSum = volumeSum + ws.Cells(i, 7).Value
    ws.Range("L" & sumTableRow).Value = volumeSum
Next i

For i = 2 To lastRow
    If ws.Cells(i, 11).Value > ws.Cells(i + 1, 11).Value And ws.Cells(i, 11).Value > ws.Cells(2, 17).Value Then
        ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
        ws.Cells(2, 17).NumberFormat = ".##%"
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        
    ElseIf ws.Cells(i, 11).Value < ws.Cells(i + 1, 11).Value And ws.Cells(i, 11).Value < ws.Cells(3, 17).Value Then
        ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
        ws.Cells(3, 17).NumberFormat = ".##%"
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        
    End If
Next i
For i = 2 To lastRow

    If ws.Cells(i, 12).Value > ws.Cells(i + 1, 12).Value And ws.Cells(i, 12).Value > ws.Cells(4, 17).Value Then
       ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
       ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    End If
     
Next i

For i = 2 To lastRow
    If ws.Cells(i, 10) >= 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If

Next i

Next ws

End Sub



