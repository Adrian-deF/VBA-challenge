# VBA-challenge

Sub Stock_Info()

Dim ws As Worksheet
Dim ticker As String
Dim volume As Double
volume = 0
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double

For Each ws In ThisWorkbook.Worksheets

    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Volume"

    Dim last_row As Long
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    For i = 2 To last_row

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ticker = Cells(i, 1).Value
        Range("J" & Summary_Table_Row).Value = ticker
        
        volume = volume + Cells(i, 7).Value
        Range("M" & Summary_Table_Row).Value = volume
        
        year_open = Cells(i, 3).Value
        year_close = Cells(i, 6).Value
    
        yearly_change = year_close - year_open
        Range("K" & Summary_Table_Row).Value = yearly_change
        
        percent_change = yearly_change / year_open * 100
        Range("L" & Summary_Table_Row).Value = percent_change
    
        Summary_Table_Row = Summary_Table_Row + 1
        
        volume = 0
        
        Else
            volume = volume + Cells(i, 7).Value

        End If
        
    Next i

End Sub

