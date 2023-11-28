# Excel'de Gruplandırılmış Frekans Serileri İçin Bazı İstatistik Hesaplama Örnekleri

Excel, aritmetik ortalamadan Gamma olasılık fonksiyonuna kadar değişen istatistiksel olan fonlsyionlar da dahil olmak üzere geniş bir fonksiyon kütüphanesine sahiptir. Ancak sorunlardan biri, bu fonksiyonların çoğunun verileri basit seriler olarak kabul etmesidir. Bazen istatistikçiler bile frekans ya da gruplandırılmış frekans serileri/tabloları/dağılımları gibi diğer serilerle çalışmak zorunda kalabilirler. Excel'in kütüphanesinde **TOPLA.ÇARPIM** (SUMPRODUCT) gibi hesaplamaları yapmanıza yardımcı olabilecek bazı fonksiyonlar da vardır, ancak yine de Excel'e nasıl yapılacağını söylemenizi gerektiriyor! Bu yüzden, burada Excel'de daha kolay hesaplama yapmak için bazı basit kod örneklerim var.

## Bazı Merkezi Eğilim Ölçüleri

Bu küçük fonksiyon kodu, Excel elektronik tablosuna, aşağıdaki gibi girilen ___grupladırılnmış frekans dağılımı___ için **aritmetik ortalama** (metot=1, varsayılan), **geometrik ortalama** (2), **harmonik ortalama** (3) ve **kareli ortalama** (4) hesaplayabilir. Tabii ki tüm olası durumlar kontrol edilmelidir, bu işlev henüz yoktur.

  <img src="https://github.com/lterlemez/Excel-VBA-Istatistik/blob/main/VBA_Statistics/media/grup_seri.PNG" width="400"/>
 
 ***Şekil 1:*** *Gruplandırılmış Seri için hesaplama örneği*

``` vba
Function GOrtalama(veri As Range, Optional metot As Integer = 1)
    'Metot=1 Aritmetik Ortalama ve varsayılan (Arithmetic Mean and setted as default)
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
