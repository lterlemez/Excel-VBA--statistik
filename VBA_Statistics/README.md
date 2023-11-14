# Some Statistical Calculation Examples For Grouped Frequency Series in Excel
Excel has a large function library, including statistical ones varying from arithmetic mean to Gamma probability function. But, one of the problems is that most of these functions accepts data as simple series. And, sometimes, even statisticians may have to work with other series like frequency or grouped frequency series/tables/distributions. Excel also has some functions in its library that can help to do the math like **SUMPRODUCT** but again, you have to tell Excel how to do!
So, here, I have some simple code samples to calculate more easily in Excel.

## Some Central Tendency Measures
This small function code can calculate arithmetic (metot=1 ,default) , geometric (2) and harmonic mean (3) for grouped frequency distribution entered  as  below in Excel spreadsheet. Of course all possible situations must be checked, this function do not have yet.

<img src="https://github.com/lterlemez/lterlemez/blob/main/VBA_istatistik/grup_seri.PNG" width="400">

```vba
Function GOrtalama(veri As Range, Optional metot As Integer = 1)
    'Metot=1 Aritmetik Ortalama ve varsayılan
    'Metot=2 Geometrik Ortalama
    'Metot=3 Harmonik Ortalama
    toplam = 0
    If metot = 1 Then
        For i = 1 To veri.Rows.Count
            toplam = toplam + WorksheetFunction.Average(Range(veri.Cells(i, 1), veri.Cells(i, 2))) * veri.Cells(i, 3)
        Next i
        ort = toplam / WorksheetFunction.Sum(veri.Columns(3))
    ElseIf metot = 2 Then
        For i = 1 To veri.Rows.Count
            toplam = toplam + WorksheetFunction.Log(WorksheetFunction.Average(Range(veri.Cells(i, 1), veri.Cells(i, 2)))) * veri.Cells(i, 3)
        Next i
        ort = WorksheetFunction.Power(10, (toplam / WorksheetFunction.Sum(veri.Columns(3))))
    ElseIf metot = 3 Then
        For i = 1 To veri.Rows.Count
            toplam = toplam + veri.Cells(i, 3) / (WorksheetFunction.Average(Range(veri.Cells(i, 1), veri.Cells(i, 2))))
        Next i
        ort = WorksheetFunction.Sum(veri.Columns(3)) / toplam
    ElseIf metot = 4 Then
        For i = 1 To veri.Rows.Count
            toplam = toplam + WorksheetFunction.Power(WorksheetFunction.Average(Range(veri.Cells(i, 1), veri.Cells(i, 2))), 2) * veri.Cells(i, 3)
        Next i
        ort = Sqr(toplam / WorksheetFunction.Sum(veri.Columns(3)))
    End If
    GOrtalama = ort
End Function
```

## Raw Moments of a Distribution

If series' column count is 1 then it is assumed as simple series, if it is 2 then is assumed as frequency ditribution series, and if it is 3 then is assumed as grouped frequency distribution series and otherwise en error message will be shown.

<img src="https://github.com/lterlemez/lterlemez/blob/main/VBA_istatistik/moment_raw.png" width="400" >

```vba
Function moment_raw(seri As Range, Optional r As Integer = 1)
    Dim t As Single
    t = 0
    Select Case seri.Columns.Count
        Case 1
            For Each i In seri
                t = t + i.Value ^ r
            Next i
           moment_raw = t / seri.Rows.Count
        Case 2
            For Each i In seri.Rows
                t = t + (i.Columns(1).Value ^ r) * i.Columns(2).Value
            Next i
            moment_raw = t / WorksheetFunction.Sum(seri.Columns(2))
            
        Case 3
            For Each i In seri.Rows
                t = t + WorksheetFunction.Average(i.Columns(1).Value, i.Columns(2).Value) ^ r * i.Columns(3).Value
            Next i
            moment_raw = t / WorksheetFunction.Sum(seri.Columns(3))
        Case Else
           moment_raw = "#N/A!"
    End Select
End Function
```
## Central Moments of a Distribution
This code is consist of conversition formulas from raw moments, but it will have classic formula calculations, too.
<img src="https://github.com/lterlemez/lterlemez/blob/main/VBA_istatistik/moment_cent.png" width="400" > </br>
<img src="https://github.com/lterlemez/lterlemez/blob/main/VBA_istatistik/moment_cent_org.png" width="400" >


```vba
Function moment_cent(moments As Range, Optional convert As Boolean = True, Optional r As Integer = 1, Optional mean As Single = 0)
    Dim t As Single
    t = 0
    With moments
        Debug.Print "Row Count: " & .Rows.Count
        Select Case convert
            Case True
                For j = 0 To .Rows.Count - 1
                    t = t + WorksheetFunction.Combin(.Rows.Count - 1, j) * (-1) ^ (.Rows.Count - 1 - j) * .Rows(j + 1) * .Rows(2) ^ (.Rows.Count - 1 - j)
                    Debug.Print "j= " & .Rows(j + 1) & " " & t
                Next j
                moment_cent = t
            Case False
                Select Case .Columns.Count
                    Case 1
                        For Each i In moments
                            t = t + (i.Value - mean) ^ r
                        Next i
                        moment_cent = t / .Rows.Count
                    Case 2
                        For Each i In moments
                            t = t + (i.Columns(1).Value - mean) ^ r * i.Columns(2).Value
                        Next i
                        moment_cent = t / WorksheetFunction.Sum(.Columns(2))
                    Case 3
                    
                    Case Else
                        moment_cent = "#N/A!"
                End Select
            Case Else
                moment_cent = "#N/A!"
        End Select
    End With
End Function
```



