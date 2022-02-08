... Dergisi makalesine ilişkin örnek kodları barındırmaktadır.

## Kontenjans Tablosu Oluşturma
```vba
Private Sub CommandButton1_Click()
    Dim veri As Range, cikti As Range
    Set veri = Range(RefEdit1.Text)
    Set cikti = Range(RefEdit2.Text)
    If WorksheetFunction.CountBlank(veri) = 0 Then
        Dim Col As New Collection, Row As New Collection
        Dim i As Long
        Dim CellVal As Variant
        For i = 1 To veri.Rows.Count
            CellVal = veri.Cells(i, 1).Value
            On Error Resume Next
            Row.Add CellVal, Chr(34) & CellVal & Chr(34)
            On Error GoTo 0
            CellVal = veri.Cells(i, 2).Value
            On Error Resume Next
            Col.Add CellVal, Chr(34) & CellVal & Chr(34)
            On Error GoTo 0
        Next i
        For i = 1 To Row.Count
            cikti.Offset(i, 0).Value = Row.Item(i)
        Next i
        cikti.Offset(Row.Count + 1, 0).Value = "Toplam"
        For i = 1 To Col.Count
            cikti.Offset(0, i).Value = Col.Item(i)
        Next i
        cikti.Offset(0, Col.Count + 1).Value = "Toplam"
        For i = 1 To Row.Count
            t = 0
            For j = 1 To Col.Count
                cikti.Offset(i, j).Value = WorksheetFunction.CountIfs(veri.Columns(1), Row.Item(i), veri.Columns(2), Col.Item(j))
                t = t + WorksheetFunction.CountIfs(veri.Columns(1), Row.Item(i), veri.Columns(2), Col.Item(j))
            Next
            cikti.Offset(i, Col.Count + 1).Value = t
        Next
        t = 0
        For i = 1 To Col.Count
            cikti.Offset(Row.Count + 1, i).Value = WorksheetFunction.Sum(Range(cikti.Offset(1, i).Address & ":" & cikti.Offset(Row.Count, i).Address))
            t = t + WorksheetFunction.Sum(Range(cikti.Offset(1, i).Address & ":" & cikti.Offset(Row.Count, i).Address))
        Next
        cikti.Offset(Row.Count + 1, Col.Count + 1).Value = t
    Else
        MsgBox "Veri setinde " & WorksheetFunction.CountBlank(veri) & " adet eksik veri var," & vbCrLf & "Makro çalıştırılmayacak!"
    End If
End Sub
```

## Beklenen Frekanslar Tablosu Oluşturma
```vba
Function ctablo(gozlem As Range)
  sa = gozlem.Rows.Count
  su = gozlem.Columns.Count
  ReDim bekle(sa-1,su-1)
  With gozlem
    For i = 1 To sa - 1
      For j = 1 To su - 1
        bekle(i-1,j-1)=.Cells(sa,j).Value*.Cells(i,su).Value/.Cells(sa,su).Value
      Next j
    Next i
  End With
  ctablo = bekle
End Function
```
## Ki-Kare Test İstatistiği ve Olasılığını Hesaplama
```vba
Function kikare_testi(gozlem As Range)
  sa = gozlem.Rows.Count
  su = gozlem.Columns.Count
  ReDim bekle(sa - 1, su - 1)
  With gozlem
    For i = 1 To sa - 1
      For j = 1 To su - 1
        bekle(i - 1, j - 1) = .Cells(sa, j).Value * .Cells(i, su).Value / .Cells(sa, su).Value
      Next j
    Next i
  End With
  With gozlem
    For i = 1 To sa - 1
      For j = 1 To su - 1
        kikare = kikare + (bekle(i - 1, j - 1) - .Cells(i, j).Value) ^ 2 / bekle(i - 1, j - 1)
      Next j
    Next i
  End With
  Dim ki(2) As Single
  ki(0) = kikare: ki(1) = 1 - WorksheetFunction.ChiSq_Dist(kikare, (sa - 2) * (su - 2), True)
  If Selection.Columns.Count = 2 Then
    kikare_testi = ki
  Else
    kikare_testi = WorksheetFunction.Transpose(ki)
  End If
End Function
```
## Fisher'in Kesin Ki-Kare Testi Olasılığının Hesaplanması
```vba
Function kikare_kesin(gozlem As Range)
    Dim veri, deneme(1), p1 As Double, p2 As Double, p As Double
    veri = gozlem
    Do
    deneme(0) = veri(2, 1): deneme(1) = veri(2, 2)
    veri(2, 1) = veri(1, 1): veri(1, 1) = veri(1, 2)
    veri(1, 2) = deneme(1): veri(2, 2) = deneme(0)
    Loop Until (veri(1, 1) = WorksheetFunction.Min(gozlem))
    veri(1, 3) = WorksheetFunction.Sum(veri(1, 1), veri(1, 2))
    veri(2, 3) = WorksheetFunction.Sum(veri(2, 1), veri(2, 2))
    veri(3, 1) = WorksheetFunction.Sum(veri(1, 1), veri(2, 1))
    veri(3, 2) = WorksheetFunction.Sum(veri(1, 2), veri(2, 2))
    veri(3, 3) = WorksheetFunction.Sum(veri(1, 3), veri(2, 3))
    p = 0: p1 = 0: p2 = 0
    For i = 0 To veri(1, 1)
        With WorksheetFunction
            p1 = .Fact(veri(1, 3)) * .Fact(veri(2, 3)) * .Fact(veri(3, 1)) * .Fact(veri(3, 2))
            p2 = .Fact(veri(3, 3)) * .Fact(veri(1, 1) - i) * .Fact(veri(1, 2) + i) * .Fact(veri(2, 1) + i) * .Fact(veri(2, 2) - i)
            p = p + (p1 / p2)
        End With
    Next i
    kikare_kesin = p
End Function
```
