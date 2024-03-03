Attribute VB_Name = "Module1"
Sub VBA_challenge_all():
    'Defining the variables
    Dim ws As Worksheet
    Dim ticker As String
    Dim volume_total As Double
    Dim table As Integer
    Dim open_value As Double
    Dim close_value As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    For Each ws In Worksheets
        ws.Activate
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    'declaring values for the variables
        volume_total = 0
        table = 2
        open_value = ws.Cells(2, 3)
        close_value = 0
        yearly_change = 0
        percent_change = 0
    'adding values to the ticker, volume_total, yearly_change, and percent_change columns
        For i = 2 To 800000
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                close_value = ws.Cells(i, 6) 'setting the close value
                yearly_change = close_value - open_value 'creating varible for yearly change
                volume_total = volume_total + ws.Cells(i, 7).Value
                percent_change = ((yearly_change) / open_value)
                ws.Range("I" & table).Value = ticker
                ws.Range("L" & table).Value = volume_total
                ws.Range("J" & table).Value = yearly_change
                ws.Range("K" & table).Value = percent_change
                table = table + 1
                volume_total = 0
                open_value = ws.Cells(i + 1, 3) 'resetting the open value
            Else
                volume_total = volume_total + ws.Cells(i, 7).Value
            End If
        Next i
    'formatting the percent_change column to %
        ws.Range("K2:K5000").NumberFormat = "0.00%"
    'formatting yearly_change to red for below zero, green for above zero
        Dim last_row As Integer
        last_row = ws.Cells(Rows.Count, 10).End(xlUp).Row
        For i = 2 To last_row
            If ws.Cells(i, 10) <= 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 4
            End If
        Next i
    'creating the table headers for greatest % increase/decrease and volume
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
    'finding the greatest % increase
        Dim percent_change_range As Range
        Set percent_change_range = ws.Range("K2:K5000")
        Dim greatest_percent As Double
        greatest_percent = Application.WorksheetFunction.Max(percent_change_range)
        ws.Range("Q2").Value = greatest_percent
    'putting the greatest % increase ticker in the table
        For i = 2 To last_row
            If ws.Cells(i, 11) = greatest_percent Then
                ws.Range("P2") = ws.Cells(i, 9)
            End If
        Next i
    'finding the greatest % decrease
        Dim greatest_decrease As Double
        greatest_decrease = Application.WorksheetFunction.Min(percent_change_range)
        ws.Range("Q3").Value = greatest_decrease
    'putting the greatest % decrease ticker in the table
        For i = 2 To last_row
            If ws.Cells(i, 11) = greatest_decrease Then
                ws.Range("P3") = ws.Cells(i, 9)
            End If
        Next i
    'finding the greatest volume
        Dim volume_range As Range
        Set volume_range = ws.Range("L2:L5000")
        Dim greatest_volume As Double
        greatest_volume = Application.WorksheetFunction.Max(volume_range)
        ws.Range("Q4").Value = greatest_volume
    'putting the greatest total volume ticker in the table
        For i = 2 To last_row
            If ws.Cells(i, 12) = greatest_volume Then
                ws.Range("p4").Value = ws.Cells(i, 9)
            End If
        Next i
    'formatting cells to %
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
    'adjusting column width
        ws.Columns("I").ColumnWidth = 9.22
        ws.Columns("J:L").ColumnWidth = 16.22
        ws.Columns("O").ColumnWidth = 19
        ws.Columns("Q").ColumnWidth = 16
    Next ws
End Sub
