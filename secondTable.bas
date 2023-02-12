Attribute VB_Name = "Module2"
Sub secondTable()

'Naming Table columns

Range("o2").Value = "Greatest % increase"
Range("o3").Value = "Greatest % decrease"
Range("o4").Value = "Greatest total volume"
Range("p1").Value = "Ticker"
Range("q1").Value = "Value"

'Cells with dates also return a value, and get covered for determining smallest value. Percentages will convert and return numerics.


Dim MinPercent As Double
Dim MaxPercent As Double
Dim PercentCol As String
Dim TickerInc As String
Dim TickerDec As String
Dim TickerVol As String

'Set range from which to determine smallest value
PercentCol = "k2:k3001"
Set rng1 = Range(PercentCol)

'Find Min % Value
MinPercent = Application.WorksheetFunction.Min(rng1)
Range("q3").Value = MinPercent

'Find Max % Value
MaxPercent = Application.WorksheetFunction.Max(rng1)
Range("q2").Value = MaxPercent

'Format As %
Range("q2:q3").NumberFormat = "0.00%"

'Volume Range
VolCol = "l2:l3001"
Set rng2 = Range(VolCol)

'Find Volume Value
MaxVolume = Application.WorksheetFunction.Max(rng2)
Range("q4").Value = MaxVolume

'Loop to find the Ticker name for the value
For i = 2 To 3001

'Ticker for Max %
If Cells(i, 11) = MaxPercent Then
TickerInc = Cells(i, 9)
Range("p2").Value = TickerInc
End If

'Ticker for Min %
If Cells(i, 11) = MinPercent Then
TickerDec = Cells(i, 9)
Range("p3").Value = TickerDec
End If

'Ticker for Volume
If Cells(i, 12) = MaxVolume Then
TickerVol = Cells(i, 9)
Range("p4").Value = TickerVol
End If

Next i

End Sub
