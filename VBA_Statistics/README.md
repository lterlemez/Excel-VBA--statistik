# Some Statistical Calculation Examples For Grouped Frequency Series in Excel


Excel has a large function library, including statistical ones varying from arithmetic mean to Gamma probability function. But, one of the problems is that most of these functions accepts data as simple series. And, sometimes, even statisticians may have to work with other series like frequency or grouped frequency series/tables/distributions. Excel also has some functions in its library that can help to do the math like **SUMPRODUCT** but again, you have to tell Excel how to do! So, here, I have some simple code samples to calculate more easily in Excel. 

## Some Rules to Select Number or Width of Bins for Histogram and Grouping Data

There is no **best** choice of ***number*** or ***width*** of **bins** for _histogram_ or _grouping data_, but there are some _suggested rules_ that can be used for choosing. This small function is calculating number or width of bins for a given simple series.

### Square-root Rule
&nbsp;&nbsp;&nbsp;&nbsp; $k=\lceil\text{ } \sqrt{n}\text{ } \rceil$

### Sturges' Rule

&nbsp;&nbsp;&nbsp;&nbsp; $k=1+\lceil log_2 n \rceil$

### Rice Rule

&nbsp;&nbsp;&nbsp;&nbsp; $k=\lceil\text{ } 2 \sqrt[3]{n} \text{ } \rceil$

### Doane's Rule

&nbsp;&nbsp;&nbsp;&nbsp; $k=1+\lceil log_2 n +log_2(1+\frac{\left | g_1 \right |}{\sigma_{g_1}})\rceil$; where $g_1$ is the estimate of skewness of the distribution and

&nbsp;&nbsp;&nbsp;&nbsp; $\sigma_{g_1}=\sqrt{\frac{6(n-1)}{(n+1)(n+3)}}$

### Scott's Rule

&nbsp;&nbsp;&nbsp;&nbsp; $h=\frac{3.49 \hat{\sigma}}{\sqrt[3]{n}}$; where $\hat{\sigma}$ is the sample standart deviation.

### Freedman-Diaconis's (FD) Rule

&nbsp;&nbsp;&nbsp;&nbsp; $h=2\frac{IQR(x)}{\sqrt[3]{n}}$

```vba
'While two declarations were made in general declaration section of the module in use
'Kullanılan modülün genel bildirimler bölümünde iki bildirim yapılmış iken
Enum etiketler
        Karekok = 1
        Sturges = 2
        Rice = 3
        Doane = 4
        Scott = 5
        FD = 6
End Enum
Dim yuvarla As Boolean

Function grupla(veri As Range, Optional metot As etiketler = Sturges, Optional yuvarla = False)
    Dim x As Integer
    If veri.Columns.Count = 1 Then
        n = veri.Rows.Count
    ElseIf veri.Rows.Count = 1 Then
        n = veri.Columns.Count
    Else
        MsgBox "Veriniz satır veya sütun şeklinde olmalı!" '/"Your data must be in rows or columns!"
    End If
    With WorksheetFunction
        Select Case metot
            Case Karekok
                'Grup sayısı döndürür/Returns number of bins
                k = WorksheetFunction.Ceiling(Sqr(n), 1)
                grupla = k
            Case Sturges
                'Grup sayısı döndürür/Returns number of bins
                k = .Ceiling(.Log(n, 2), 1) + 1
                grupla = k
            Case Rice
                'Grup sayısı döndürür/Returns number of bins
                k = .Ceiling(2 * .Power(n, 1 / 3), 1)
                grupla = k
            Case Doane
                'Grup sayısı döndürür/Returns number of bins
                sd = Sqr((6 * (n - 2)) / ((n + 1) * (n + 3)))
                k = 1 + .Ceiling(.Log(n, 2) + .Log(1 + Abs(.Skew(veri)) / sd, 2), 1)
                grupla = k
            Case Scott
                'Grup aralığı döndürür/Returns width of bins
                h = 3.5 * .StDev_S(veri) / .Power(n, 1 / 3)
                grupla = h
            Case FD
                'Grup aralığı döndürür/Returns width of bins
                h = 2 * (.Quartile_Exc(veri, 3) - .Quartile_Exc(veri, 1)) / .Power(n, 1 / 3)
                grupla = h
        End Select
        If yuvarla Then
            grupla = .Round(grupla, 2)
        Else
            grupla = grupla
            End If
    End With
End Function
```

## Some Central Tendency Measures

This small function code can calculate **arithmetic mean** (metot=1 ,default) , **geometric mean** (2), **harmonic mean** (3) and **root mean square** (4) for ___grouped frequency distribution___ entered as below in Excel spreadsheet. Of course all possible situations must be checked, this function do not have yet.

<img src="https://github.com/lterlemez/Excel-VBA-Istatistik/blob/main/VBA_Statistics/media/grup_seri.PNG" width="400"/>

***Figure 1:*** *Calculation example for grouped frequency distribution*

``` vba
Function GOrtalama(veri As Range, Optional metot As Integer = 1)
    'Metot=1 Aritmetik Ortalama ve varsayılan (Arithmetic Mean and set as default)
    'Metot=2 Geometrik Ortalama (Geometric Mean)
    'Metot=3 Harmonik Ortalama (Harmonic Mean)
    'Metot=4 Kareli Ortalama (Root Mean Square)
    toplam = 0
    If metot = 1 Then
        For i = 1 To veri.Rows.Count
            toplam = toplam + WorksheetFunction.Average(veri.Cells(i, 1), veri.Cells(i, 2)) * veri.Cells(i, 3)
        Next i
        ort = toplam / WorksheetFunction.Sum(veri.Columns(3))
    ElseIf metot = 2 Then
        For i = 1 To veri.Rows.Count
            toplam = toplam + WorksheetFunction.Log(WorksheetFunction.Average(Range(veri.Cells(i, 1), veri.Cells(i, 2)))) * veri.Cells(i, 3)
        Next i
        ort = WorksheetFunction.Power(10, (toplam / WorksheetFunction.Sum(veri.Columns(3))))
    ElseIf metot = 3 Then
        For i = 1 To veri.Rows.Count
            toplam = toplam + veri.Cells(i, 3) / (WorksheetFunction.Average(veri.Cells(i, 1), veri.Cells(i, 2)))
        Next i
        ort = WorksheetFunction.Sum(veri.Columns(3)) / toplam
    ElseIf metot = 4 Then
        For i = 1 To veri.Rows.Count
            toplam = toplam + WorksheetFunction.Power(WorksheetFunction.Average(veri.Cells(i, 1), veri.Cells(i, 2)), 2) * veri.Cells(i, 3)
        Next i
        ort = Sqr(toplam / WorksheetFunction.Sum(veri.Columns(3)))
    End If
    GOrtalama = ort
End Function
```

## Raw Moments of a Distribution

If series' column count is 1 then it is assumed as simple series, if it is 2 then is assumed as frequency ditribution series, and if it is 3 then is assumed as grouped frequency distribution series and otherwise en error message will be shown.

Flowchart of calculation algorith for simple series:

```mermaid
%%{ init: { 'flowchart': { 'curve': 'monotoneXY' } } }%%
flowchart LR
A([seri, r]) --> B
B{seri.Cols.Count}
B--1--> E{{"i=1;seri.Rows.Count"}}
E-->F("t=t+seri(i)^r")-->G((i)) 
G-->E
E----Z1((" ")):::hidden--->H("moment_raw=t/seri.Rows.Count")-->Z((" "))
B--2-->K{{"i=1;seri.Rows.Count"}}
K-->M("t=t+seri(i,1)^r * seri(i,2)")-->N((i))
N-->K
K----Z2((" ")):::hidden--->O("moment_raw=t/SUM(seri(,2))")-->Z((" "))-->J([moment_raw])
B--3-->L{{"i=1;seri.Rows.Count"}}
L-->P("t=t+(AVERAGE(seri(i,1),seri(i,2))^r * seri(i,3)")-->Q((i))
Q-->L
L----Z3((" ")):::hidden--->R("moment_raw=t/SUM(seri(,3))")-->Z((" "))
linkStyle 5 stroke:black,stroke-width:2px,color;
classDef hidden display: none;
```

<img src="https://github.com/lterlemez/Excel-VBA-Istatistik/blob/main/VBA_Statistics/media/moment_raw.png" width="400"/>

``` vba
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

This code is consist of conversition formulas from raw moments, but it will have classic formula calculations, too. <img src="https://github.com/lterlemez/Excel-VBA-Istatistik/blob/main/VBA_Statistics/media/moment_cent.png" width="400"/> </br> <img src="https://github.com/lterlemez/Excel-VBA-Istatistik/blob/main/VBA_Statistics/media/moment_cent_org.png" width="400"/>

``` vba
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
                        For Each i In seri.Rows
                              t = t + (WorksheetFunction.Average(i.Columns(1).Value, i.Columns(2).Value)-mean) ^ r * i.Columns(3).Value
                        Next i
                        moment_raw = t / WorksheetFunction.Sum(seri.Columns(3))
                    Case Else
                        moment_cent = "#N/A!"
                End Select
            Case Else
                moment_cent = "#N/A!"
        End Select
    End With
End Function
```
