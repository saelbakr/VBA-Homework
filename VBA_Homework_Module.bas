Attribute VB_Name = "Module1"
Sub Main()
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet


For Each ws In ThisWorkbook.Worksheets
ws.Activate

Call VBA
Call MaxIncrease
Call MaxDecrease
Call TotalVolume

Next ws

End Sub


Sub VBA()
rownum = Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox (rownum)



    Dim ticker_value As String
    Dim yearly_change As Double
    Dim percent_change As Double
    yearly_change = 0
    Dim volume As Double
    volume = 0
    Dim tick As Integer
    tick = 2
    Dim openingprice As Double
    openingprice = 0
    Dim closingprice As Double
    closingprice = 0



    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
'We define the opening price of A of the first stock
    openingprice = Cells(2, 3).Value


For i = 2 To rownum
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    'Here we extract the last value before Ticker symbol changed
    closingprice = Cells(i, 6).Value
    yearly_change = (closingprice - openingprice)
    Range("J" & tick).Value = yearly_change
    

        If openingprice = 0 Then
            Range("K" & tick).Value = 0
        Else:
            percent_change = (closingprice - openingprice) / openingprice
            Range("K" & tick).Value = FormatPercent(percent_change)
        End If
        
            If Range("K" & tick).Value >= 0 Then
                Range("K" & tick).Interior.Color = vbGreen
            Else:
                Range("K" & tick).Interior.Color = vbRed
            End If
            

    'Here we found the String of each ticker
    ticker_value = Cells(i, 1).Value
    Range("I" & tick).Value = ticker_value

        volume = volume + Cells(i, 7).Value
      Range("L" & tick).Value = volume

    tick = tick + 1

    openingprice = Cells(i + 1, 3).Value

 'Reset volume here to start fresh

volume = 0

Else:

    volume = volume + Cells(i, 7).Value



    End If
Next i



End Sub

Sub MaxIncrease()





Cells(1, 16).Value = "Value"
Cells(1, 15).Value = "Ticker"
Cells(2, 14).Value = "Greatest % Increase"

greatest_change = 0
Dim greatest_ticker As String

myrows = Cells(Rows.Count, 11).End(xlUp).Row
    For i = 2 To myrows
        If greatest_change < Cells(i + 1, 11).Value Then
            greatest_change = Cells(i + 1, 11).Value
            greatest_ticker = Cells(i + 1, 9).Value
        
    End If

Next i
Range("P" & 2).Value = FormatPercent(greatest_change)
Range("O" & 2).Value = greatest_ticker


End Sub

Sub MaxDecrease()




Cells(3, 14).Value = "Greatest % Decrease"

greatest_change = 0
Dim greatest_ticker As String

myrows = Cells(Rows.Count, 11).End(xlUp).Row
    For i = 2 To myrows
        If greatest_change > Cells(i + 1, 11).Value Then
            greatest_change = Cells(i + 1, 11).Value
            greatest_ticker = Cells(i + 1, 9).Value
        
    End If

Next i
Range("O" & 3).Value = greatest_ticker
Range("P" & 3).Value = FormatPercent(greatest_change)


End Sub

Sub TotalVolume()



Cells(4, 14).Value = "Greatest Total Volume"

greatest_volume = 0
Dim greatest_ticker As String

myrows = Cells(Rows.Count, 12).End(xlUp).Row
    For i = 2 To myrows
        If greatest_volume < Cells(i + 1, 12).Value Then
            greatest_volume = Cells(i + 1, 12).Value
            greatest_ticker = Cells(i + 1, 9).Value
        
    End If

Next i
Range("O" & 4).Value = greatest_ticker
Range("P" & 4).Value = Format(greatest_volume, "scientific")


End Sub

