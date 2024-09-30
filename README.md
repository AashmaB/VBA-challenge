# VBA-challenge
VBA-challenge assignment
Sub multiple_year_stock_data()

Dim ws As Worksheet
For Each ws In Worksheets

Dim i As Long
Dim ticker As String
Dim rocount As Double

Dim total_stock_volume As Long

total_stock_volume = 0
Dim percent_change As Double
Dim quarterly_change As Double


Dim summary_table_row As Long
summary_table_row = 2



Cells(1, 9).Value = "ticker"
Cells(1, 10).Value = "quarterly_change"
Cells(1, 11).Value = "percentage_change"
Cells(1, 12).Value = "total_stock_volume"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"



For i = 2 To 93001

Cells(i, 10).Value = Cells(i, 6) - Cells(i, 3)

Set r1 = Range("J" & i)
If r1.Value < 0 Then r1.Interior.ColorIndex = 3
If r1.Value > 0 Then r1.Interior.ColorIndex = 4

Cells(i, 11).Value = Cells(i, 10) / Cells(i, 3)
 
Cells(i, 11).Value = FormatPercent(Cells(i, 11))
 
 
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
ticker = Cells(i, 1).Value

total_stock_volume = total_stock_volume + Cells(i, 7).Value

Range("I" & summary_table_row).Value = ticker

Range("L" & summary_table_row).Value = total_stock_volume

summary_table_row = summary_table_row + 1

total_stock_volume = 0


Else
total_stock_volume = total_stock_volume + Cells(i, 7).Value


End If
Next i

Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & RowCount)) * 100
Range("Q3") = "%" & worksheeetfunction.Min(Range("K2:K" & RowCount)) * 100
Range("Q4") = WorksheetFunction.Max(Range("L2:L" & RowCount))

increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & RowCount)), Range("L2:L" & RowCount), 0)


Range("P2") = Cells(increase_number + 1, 9)
    Range("P3") = Cells(decrease_number + 1, 9)
    Range("P4") = Cells(volume_number + 1, 9)


End Sub
