Sub WallStStockData()

'Loop through all sheets
For Each WS In ActiveWorkbook.Worksheets
WS.Activate

'Find Last Row
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

'Add all Column Headers
WS.Cells(1, "I").Value = "Ticker"
WS.Cells(1, "J").Value = "Yearly Change"
WS.Cells(1, "K").Value = "Percent Change"
WS.Cells(1, "L").Value = "Total Stock Volume"

'Set Variables' data types
Dim Column As Long
Dim Row As Long

Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Ticker_Symbol As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Stock_Volume As Double
Dim i As Long

'Opening_Price assignment
Row = 2
Column = 1
Opening_Price = Cells(2, Column + 2).Value

'Loop through tickers
For i = 2 To LastRow
If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then

'Ticker_Symbol assignment
Ticker_Symbol = Cells(i, Column).Value
Cells(Row, Column + 8).Value = Ticker_Symbol

'Closing_Price assignment
Closing_Price = Cells(i, Column + 5).Value

'Yearly_Change assignment
Yearly_Change = Closing_Price - Opening_Price
Cells(Row, Column + 9).Value = Yearly_Change

'Percent_Change assignment
If (Opening_Price = 0 And Closing_Price = 0) Then
Percent_Change = 0
ElseIf (Opening_Price = 0 And Closing_Price <> 0) Then
Percent_Change = 1
Else
Percent_Change = Yearly_Change / Opening_Price

Cells(Row, Column + 10).Value = Percent_Change
Cells(Row, Column + 10).NumberFormat = "0.00%"
End If

'conditional matching for ticker
Stock_Volume = Stock_Volume + Cells(i, Column + 6).Value

Cells(Row, Column + 11).Value = Stock_Volume

Row = Row + 1
'reset Opening_Price value
Opening_Price = Cells(i + 1, Column + 2)
Stock_Volume = 0

Else
Stock_Volume = Stock_Volume + Cells(i, Column + 6).Value
End If

Next i


'Cell color formatting
For j = 2 To WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
Cells(j, Column + 9).Interior.ColorIndex = 4
ElseIf Cells(j, Column + 9).Value < 0 Then
Cells(j, Column + 9).Interior.ColorIndex = 3
End If

Next j
        
Next WS
        
End Sub




