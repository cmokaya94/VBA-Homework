Attribute VB_Name = "Module1"
Sub stockinfo()
Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

 ws.Range("O2").Value = "Greatest Increase"
 ws.Range("O3").Value = "Greatest Decrease"
 ws.Range("O4").Value = "Greatest Total Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"


Dim ts As String
Dim last_row As Long
Dim startingprice As Double
Dim endingprice As Double
Dim yearlychange As Double
Dim percentchange As Double

Dim ts_total As Double
ts_total = 0


Dim maxv As Double
maxv = WorksheetFunction.Max(ws.Range("K:K"))
ws.Range("Q2").Value = maxv


Dim minv As Double
minv = WorksheetFunction.Min(ws.Range("K:K"))
ws.Range("Q3").Value = minv

Dim GTV As Double
GTV = WorksheetFunction.Max(ws.Range("L:L"))
ws.Range("Q4").Value = GTV

Dim row As Long

Dim counter As Long
counter = 2

last_row = Cells(Rows.Count, 1).End(xlUp).row
For row = 2 To last_row

    If ws.Cells(row, 1).Value <> ws.Cells(row - 1, 1).Value Then
        startingprice = ws.Cells(row, 3).Value
    End If
    
If ws.Cells(row, 1).Value <> ws.Cells(row + 1, 1).Value Then

    ts = ws.Cells(row, 1).Value
    endingprice = ws.Cells(row, 6)
    
    yearlychange = endingprice - startingprice
    ts_total = ts_total + ws.Cells(row, 7).Value
    
    ws.Range("I" & counter).Value = ts
    ws.Range("J" & counter).Value = yearlychange
    
    
    If ws.Range("J" & counter).Value >= 0 Then
        ws.Range("J" & counter).Interior.ColorIndex = 4
    ElseIf ws.Range("J" & counter).Value < 0 Then
        ws.Range("J" & counter).Interior.ColorIndex = 3
    End If
    
    If (startingprice = 0) Then
        percentchange = 0
        ws.Range("K" & counter).Value = percentchange
    Else
        percentchange = yearlychange / startingprice
        ws.Range("K" & counter).Value = percentchange
    End If
    
    ws.Range("K" & counter).NumberFormat = "0.00%"
    ws.Range("L" & counter).Value = ts_total
    
    counter = counter + 1
    ts_total = 0
    
    Else
    ts_total = ts_total + ws.Cells(row, 7).Value
    yearlychange = endingprice - startingprice
    
    End If
Next row

last_row = ws.Cells(Rows.Count, 9).End(xlUp).row

For row = 2 To last_row
    If ws.Range("K" & row).Value = maxv Then
    ws.Range("Q2").Value = maxv
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("P2").Value = ws.Range("I" & row).Value
    End If
    
    If ws.Range("K" & row).Value = minv Then
    ws.Range("Q3").Value = minv
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("P3").Value = ws.Range("I" & row).Value
    End If
    
    If ws.Range("L" & row).Value = GTV Then
    ws.Range("Q4").Value = ws.Range("L" & row).Value
    ws.Range("P4").Value = ws.Range("I" & row).Value
    End If
Next row
ws.Columns("I:Q").AutoFit

Next ws

End Sub
    
    
       


