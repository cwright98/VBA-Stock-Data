Attribute VB_Name = "Module3"
Sub WonkyStonky():


Dim i As Integer
Dim x As Integer
Dim cond1 As FormatCondition, cond2 As FormatCondition, cond3 As FormatCondition
i = 2
x = 2
Cells(1, 10).Value = "Ticker"
Cells(1, 11).Value = "Yearly Change"
Cells(1, 12).Value = "Percentage Change"
Cells(1, 13).Value = "Total Volume"
While Cells(i, 1).Value <> ""

    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        Cells(x, 10).Value = Cells(i, 1).Value
        Cells(x, 11).Value = Cells(i, 6).Value - Cells(i - i + 2, 3).Value
        Cells(x, 12).Value = Cells(x, 11) / Cells(i, 3)
        x = x + 1
        i = i + 1
    ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value Then
        i = i + 1
        Cells(x, 13).Value = Cells(x, 13).Value + Cells(i, 7).Value
         
    End If
Wend

Set rg = Range("K2", Range("K2").End(xlDown))

Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, "0")
Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, "0")
Set cond3 = rg.FormatConditions.Add(xlCellValue, xlEqual, "0")

With cond1
.Interior.Color = vbGreen
.Font.Color = vbBlack
End With

With cond2
.Interior.Color = vbRed
.Font.Color = vbWhite
End With

'With cond3
'.Interior.Color = ""
'.Font.Color = vbBlack
'End With

'loop through again and find max/min values
rn = 2
Max = Cells(rn, 11).Value
Min = Cells(rn, 11).Value
While Cells(rn, 11).Value <> ""
    If Cells(rn, 11).Value > Max Then
        Max = Cells(rn, 11).Value
        Cells(2, 16).Value = Cells(rn, 10).Value
        Cells(2, 17).Value = Max
        rn = rn + 1
    ElseIf Cells(rn, 11).Value < Min Then
        Min = Cells(rn, 11).Value
        Cells(3, 16).Value = Cells(rn, 10).Value
        Cells(3, 17).Value = Min
        rn = rn + 1
    End If
Wend
'reset cells to start back at row 2
rn = 2
Max_total_volume = Cells(2, 13).Value
While Cells(rn, 13).Value <> ""
    If Cells(rn, 13).Value > Max_total_volume Then
        Max_total_volume = Cells(rn, 13).Value
        Cells(4, 16) = Max_total_volume
        Cells(4, 17) = Cells(rn, 10).Value
        rn = rn + 1
    Else
        rn = rn + 1
    End If
Wend

        
        
End Sub
