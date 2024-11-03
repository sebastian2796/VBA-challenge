Attribute VB_Name = "Module1"
Sub StockAnalysis()
'declare variables
Dim Qchange As Double
Dim Pchange As Double
Dim Totalvolume As Double
Dim i As Long
Dim Lastrow As Long
Dim j As Integer
Dim Start As Long

'set column headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Quarterly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
'set initial values
Qchange = 0
Totalvolume = 0
j = 0
Start = 2
'get last row number of data
Lastrow = Cells(Rows.Count, "A").End(xlUp).Row
'set for loop(set conditional for ticker change, calculations, print results, set colors)
For i = 2 To Lastrow
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Totalvolume = Totalvolume + Cells(i, 7).Value
If Totalvolume = 0 Then
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("L" & 2 + j).Value = 0
Else

                If Cells(Start, 3) = 0 Then
                    For find_value = Start To i
                        If Cells(find_value, 3).Value <> 0 Then
                            Start = find_value
                            Exit For
                        End If
                     Next find_value
                End If
                Start = i + 1
Range("L" & j + 2).Value = Totalvolume
Range("I" & j + 2).Value = Cells(i, 1).Value
End If


Qchange = 0
Totalvolume = 0

j = j + 1
Else
Totalvolume = Totalvolume + Cells(i, 7).Value

End If


Next i
'end for loop

'find max and min of values

End Sub