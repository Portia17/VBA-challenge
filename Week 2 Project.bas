Attribute VB_Name = "Module1"
Sub Week2()

'Get it to apply to all worksheets
For Each ws In Worksheets

'Name Column
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'Get it to print in the right column
Dim tickercount As Long
tickercount = 2

Dim firstopen As Double
Dim lastclose As Double
Dim TotalRecords As Long
Dim i As Long
Dim column As Variant


'Get end row of data
TotalRecords = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Auto adjust the width of the columns in table one
            ws.Columns("J:Q").AutoFit

firstopen = 2

For i = 2 To TotalRecords


    'Get it to print ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ws.Cells(tickercount, 9).Value = ws.Cells(i, 1).Value
    
    'Get it to print yearly change
    ws.Cells(tickercount, 10).Value = ws.Cells(i, 6) - ws.Cells(firstopen, 3).Value
    
    'get it to print percent change
    ws.Cells(tickercount, 11).Value = (ws.Cells(i, 6).Value / ws.Cells(firstopen, 3).Value) - 1
    
    'sum ticker
    ws.Cells(tickercount, 12).Value = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(firstopen, 7), ws.Cells(i, 7)))
    
    'add color to yearly change
        If ws.Cells(tickercount, 10).Value > 0 Then
       ws.Cells(tickercount, 10).Interior.ColorIndex = 4
       
        ElseIf ws.Cells(tickercount, 10).Value <= 0 Then
       ws.Cells(tickercount, 10).Interior.ColorIndex = 3
       
       End If
    
    'Get greatest percentage
    If ws.Cells(tickercount, 11).Value > ws.Cells(2, 17).Value Then
    ws.Cells(2, 17).Value = ws.Cells(tickercount, 11).Value
    ws.Cells(2, 16).Value = ws.Cells(tickercount, 9).Value
    
    End If
    
    'Get smallest percentage
    If ws.Cells(tickercount, 11).Value < ws.Cells(3, 17).Value Then
    ws.Cells(3, 17).Value = ws.Cells(tickercount, 11).Value
    ws.Cells(3, 16).Value = ws.Cells(tickercount, 9).Value
    End If
    
    'Get volume
    If ws.Cells(tickercount, 12).Value > ws.Cells(4, 17) Then
    ws.Cells(4, 17).Value = ws.Cells(tickercount, 12).Value
    ws.Cells(4, 16).Value = ws.Cells(tickercount, 9).Value
    End If
    
    ws.Columns(11).NumberFormat = "0.00%"
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    
    tickercount = tickercount + 1
    firstopen = i + 1
    
    End If
    
    TotalRecords = TotalRecords + 1

Next i



Next ws

MsgBox ("Complete")

End Sub

