Attribute VB_Name = "Module1"
Sub JK()

Dim open_price As Double
Dim close_price As Double
Dim ticker As Integer
ticker = 2
Dim percent_change As Double
Dim yearly_change As Double
Dim total_stock As Double
total_stock = 0

Cells(1, 9).Value = "ticker"
Cells(1, 10).Value = "yearly change"
Cells(1, 11).Value = "percent change"
Cells(1, 12).Value = "total stock volume"

open_price_index = 2

For i = 2 To 705646

If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
Cells(ticker, 9).Value = Cells(i, 1).Value

If open_price = 0 Then

For j = open_price_index To i
If Cells(j, 3).Value <> 0 Then
open_price = Cells(j, 3).Value
Exit For
End If

Next j

End If
open_price_index = i + 1
close_price = Cells(i, 6).Value
If open_price = 0 Then
percent_change = 0

Else
percent_change = (close_price / open_price) - 1
Cells(ticker, 11).Value = percent_change
Cells(ticker, 11).NumberFormat = "0.00%"

End If



yearly_change = close_price - open_price

Cells(ticker, 10).Value = yearly_change

' Wherever you put this one, it's fine
total_stock = total_stock + Cells(i, 7).Value
' Without this total_stock increment here in the If-statement, the summary table is not gonna calculate the last volume of every ticker
' Total stock increment has to be ahead of this equation, or else it's not going to calculate the last volume
Cells(ticker, 12).Value = total_stock
' this is in the last row too so that the total volumn restart calculating for another ticker
total_stock = 0





If yearly_change > 0 Or yearly_change = 0 Then
Cells(ticker, 10).Interior.ColorIndex = 8

Else
Cells(ticker, 10).Interior.ColorIndex = 4


End If
' put it after percent_change so that the next percent_change can be calculated
open_price = Cells(i + 1, 3)

' you always make sure you put the ticker + 1 in the last row so the summary table doesn't slide down a row
ticker = ticker + 1

Else
' If you don't put it here, then not only is it not gonna add all the volumes inside each ticker, but it's gonna
' input only the last volume of each ticker
total_stock = total_stock + Cells(i, 7).Value

End If


Next i

For j = 2 To 2835

If Cells(j, 11).Value = Application.WorksheetFunction.Max(Range("K2:K2835")) Then
                Cells(2, 16).Value = Cells(j, 9).Value
                Cells(2, 17).Value = Cells(j, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(j, 11).Value = Application.WorksheetFunction.Min(Range("K2:K2835")) Then
                Cells(3, 16).Value = Cells(j, 9).Value
                Cells(3, 17).Value = Cells(j, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(j, 12).Value = Application.WorksheetFunction.Max(Range("L2:L2835")) Then
                Cells(4, 16).Value = Cells(j, 9).Value
                Cells(4, 17).Value = Cells(j, 12).Value
                Cells(4, 17).NumberFormat = "0.0000E+00"
            End If

Next j

End Sub

