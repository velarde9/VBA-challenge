# VBA-challenge
Sub stockloop()

Dim ws As Worksheet
Dim wb As Workbook

Set wb = ActiveWorkbook

'analyze each worksheet
For Each ws In Worksheets

'summary table
Dim summary_table As Long
summary_table = 2

'headings for summary table
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "YearlyChange"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'variables for ticker
Dim Ticker As String
Ticker = ""
Dim min_pc_ticker As String
min_pc_ticker = ""
Dim max_pc_ticker As String
max_pc_ticker = ""
Dim max_tsv_ticker As String
max_tsv_ticker = ""

'variables for yearly change
Dim yc As Double
yc = 0
Dim min_yc As Double
min_yc = 0
Dim max_yc As Double
max_yc = 0

'variables for percent change
Dim pc As Double
pc = 0
Dim min_pc As Double
min_pc = 0
Dim max_pc As Double
max_pc = 0

'variables for total stock volume
Dim tsv As Double
tsv = 0
Dim max_tsv As Double
max_tsv = 0

'variables for price
Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0
Dim price_change As Double
price_change = 0

'last row
Dim last_row As Long
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

open_price = ws.Cells(2, 3).Value

'calculations
For i = 2 To last_row
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        close_price = ws.Cells(i, 6).Value
        yc = close_price - open_price
        pc = (yc / open_price) * 100
        tsv = tsv + Cells(i, 7).Value

'outputs
    ws.Range("I" & summary_table).Value = Ticker
    ws.Range("J" & summary_table).Value = yc
    ws.Range("K" & summary_table).Value = CStr(pc) & "%"
    ws.Range("L" & summary_table).Value = tsv

'colour fill for yearly change
        If (pc > 0) Then
            ws.Range("J" & summary_table).Interior.ColorIndex = 4
        ElseIf (pc <= 0) Then
            ws.Range("J" & summary_table).Interior.ColorIndex = 3
        End If

'calculate next ticker
        summary_table = summary_table + 1
        open_price = ws.Cells(i + 1, 3).Value

'max and min percent change
        If (pc > max_pc) Then
            max_pc = pc
            max_pc_ticker = Ticker

            ElseIf (pc < min_pc) Then
                min_pc = pc
                min_pc_ticker = Ticker
        End If

'max total stock value
        If (tsv > max_tsv) Then
            max_tsv = tsv
            max_tsv_ticker = Ticker
        End If

'reset total stock value
    tsv = 0

        Else
        tsv = tsv + ws.Cells(i, 7).Value
    End If

Next i

'analysis table
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("P2").Value = max_pc_ticker
    ws.Range("P3").Value = min_pc_ticker
    ws.Range("P4").Value = max_tsv_ticker
    ws.Range("Q1").Value = "Value"
    ws.Range("Q2").Value = CStr(max_pc) & "%"
    ws.Range("Q3").Value = CStr(min_pc) & "%"
    ws.Range("Q4").Value = max_tsv

Next ws

End Sub
