Attribute VB_Name = "Module2"
Sub Stocks()
'define variables
    Dim ws As Worksheet
    Dim i As Long
    Dim lastrow As Long
    Dim k As Integer
        k = 2
    Dim l As Integer
        l = 0
    Dim total_volume As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Percent_Change As Double
    Dim YearlyChange As Double

'loop over each ws
    For Each ws In Worksheets

'make/name columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"

      'determining the last row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'looping through each row
'<> if two values are equal then true
        For i = 2 To lastrow
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                Open_Price = ws.Cells(i, 3).Value
            End If
                total_volume = total_volume + ws.Cells(i, 7)
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ws.Cells(k, 9) = ws.Cells(i, 1).Value
                ws.Cells(k, 12) = total_volume
                Close_Price = ws.Cells(i, 6).Value
                
'calc yearly change
            If Open_Price <> 0 Then
                PercentChange = ((Close_Price - Open_Price) / Open_Price)
                YearlyChange = Close_Price - Open_Price
             Else
                 PercentChange = 0
                 YearlyChange = 0
              End If

'print percent change and yearly change as percents
                ws.Cells(k, 11) = PercentChange
                ws.Cells(k, 11).NumberFormat = "0.00%"
                ws.Cells(k, 10) = YearlyChange

 'conditionals
                    If ws.Cells(k, 10).Value > 0 Then
                        ws.Cells(k, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(k, 10).Interior.ColorIndex = 3
                    End If
                total_volume = 0
                k = k + 1
                l = 0
            End If

        Next i
    k = 2
    Next ws                                            'repeat'

End Sub
