Attribute VB_Name = "Module1"
Sub firstTable()

Dim Ticker As String

Dim Yearly_Change As Double
Yearly_Change = 0

Dim Percent_change As Double
Percent_change = 0

Dim TSV As Double
TSV = 0


Dim Summary_Table_Row As Single

Summary_Table_Row = 2


'Naming Table columns

Range("i1").Value = "Ticker"
Range("j1").Value = "Yearly Change"
Range("k1").Value = "Percent Change"
Range("l1").Value = "Total Stock Volume"

'Check if I am still within the same ticker, if not..

For i = 2 To 753001

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'Ticker column
Ticker = Cells(i, 1).Value
Range("i" & Summary_Table_Row).Value = Ticker

'Yearly Change column
Yearly_Change = Format(Yearly_Change + Cells(i, 6).Value - Cells(i, 3).Value, "0.00")
Range("j" & Summary_Table_Row).Value = Yearly_Change
Yearly_Change = 0

'Percent Change column
Range("k2:k3001").NumberFormat = "0.00%"
Percent_change = Percent_change + ((Cells(i, 6).Value - Cells(i, 3).Value) / Cells(i, 3).Value)
Range("k" & Summary_Table_Row).Value = Percent_change
Percent_change = 0

'Stock volume column
TSV = TSV + Cells(i, 7).Value
Range("l" & Summary_Table_Row).Value = TSV
TSV = 0

Summary_Table_Row = Summary_Table_Row + 1

Else

Yearly_Change = Yearly_Change + Cells(i, 6).Value - Cells(i, 3).Value
Percent_change = Percent_change + ((Cells(i, 6).Value - Cells(i, 3).Value) / Cells(i, 3).Value)
TSV = TSV + Cells(i, 7).Value

End If

For j = 2 To 3001

' Adding colour to the Yearly Change Column
If Cells(j, 10) < 0 Then
Cells(j, 10).Interior.ColorIndex = 3
Else
Cells(j, 10).Interior.ColorIndex = 4
End If

If Cells(j, 11) < 0 Then
Cells(j, 11).Interior.ColorIndex = 3
Else
Cells(j, 11).Interior.ColorIndex = 4
End If

Next j

Next i

End Sub

