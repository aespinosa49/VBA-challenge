Sub TickerRun()
For Each ws In ThisWorkbook.Worksheets
ws.Activate
    Range("H1").Value = "Ticker Symbol"
    Range("I1").Value = "Yearly Change"
    Range("J1").Value = "Percent Change"
    Range("K1").Value = "Stock Volume"
    Dim EndRows As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Dim Ticker_Symbol As String
    Dim Ticker_Rows As Integer
    Ticker_Rows = 2
    Dim Ticker_Volume As Double
    Ticker_Volume = 0
    Dim Open_Price As Double
    Open_Price = Cells(2, 3).Value
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker_Symbol = Cells(i, 1).Value
        Range("H" & Ticker_Rows).Value = Ticker_Symbol
        Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
        Range("K" & Ticker_Rows).Value = Ticker_Volume
        Close_Price = Cells(i, 6).Value
        Yearly_Change = Close_Price - Open_Price
        Range("I" & Ticker_Rows).Value = Yearly_Change
            If (Open_Price = 0) Then
                Percent_Change = 0
            Else
                Percent_Change = Yearly_Change / Open_Price
            End If
        Range("J" & Ticker_Rows).Value = Percent_Change
        Range("J" & Ticker_Rows).NumberFormat = "0.00%"
        Ticker_Rows = Ticker_Rows + 1
        Ticker_Volume = 0
        Open_Price = Cells(i + 1, 3)
        Else
        Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
        End If
    Next i
    For i = 1 To LastRow
        If Cells(i + 1, 9).Value > 0 Then
        Cells(i + 1, 9).Interior.ColorIndex = 4
        Else
        Cells(i + 1, 9).Interior.ColorIndex = 3
        End If
    Next i
Next ws

End Sub
