VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} txtBirimFiyat 
   Caption         =   "Satýn Alma Kayýt Formu"
   ClientHeight    =   10056
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8436.001
   OleObjectBlob   =   "txtBirimFiyat.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "txtBirimFiyat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAnaliz_Click()

    Dim ws As Worksheet
    Dim sonSatir As Long
    Dim i As Long
    Dim secilenTedarikci As String
    Dim tedarikci As String
    Dim kalemNo As String
    Dim toplamMiktar As Double
    Dim toplamTutar As Double
    Dim farkliKalemSayisi As Long
    Dim dict As Object
    
    Set ws = ThisWorkbook.Worksheets("SATIN ALMA")
    Set dict = CreateObject("Scripting.Dictionary")
    
    secilenTedarikci = Trim(cmbRaporTedarikci.Value)
    
    If secilenTedarikci = "" Then
        MsgBox "Lütfen analiz edilecek tedarikçiyi seçiniz.", _
               vbExclamation, "Eksik Bilgi"
        Exit Sub
    End If
    
    sonSatir = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    toplamMiktar = 0
    toplamTutar = 0
    
    For i = 2 To sonSatir
        
        tedarikci = Trim(ws.Cells(i, "B").Value)
        
        If StrComp(tedarikci, secilenTedarikci, vbTextCompare) = 0 Then
            
            kalemNo = Trim(ws.Cells(i, "A").Value)
            
            If kalemNo <> "" Then
                If Not dict.Exists(kalemNo) Then
                    dict.Add kalemNo, True
                End If
            End If
            
            If IsNumeric(ws.Cells(i, "C").Value) Then
                toplamMiktar = toplamMiktar + CDbl(ws.Cells(i, "C").Value)
            End If
            
            If IsNumeric(ws.Cells(i, "F").Value) And ws.Cells(i, "F").Value <> "" Then
    
                toplamTutar = toplamTutar + CDbl(ws.Cells(i, "F").Value)

            ElseIf IsNumeric(ws.Cells(i, "C").Value) And _
                IsNumeric(ws.Cells(i, "D").Value) Then
    
                toplamTutar = toplamTutar + _
                           (CDbl(ws.Cells(i, "C").Value) * _
                            CDbl(ws.Cells(i, "D").Value))

            End If
            
        End If
        
    Next i
    
    farkliKalemSayisi = dict.Count
    
    lblKalemSayisi.Caption = _
        "Farklý Kalem Sayýsý: " & farkliKalemSayisi
    
    lblToplamMiktar.Caption = _
        "Toplam Miktar: " & Format(toplamMiktar, "#,##0.##")
    
    lblToplamTutarRapor.Caption = _
        "Toplam Tutar: " & Format(toplamTutar, "#,##0.00") & " TL"
    
    If toplamMiktar > 0 Then
        lblOrtalamaFiyat.Caption = _
            "Ortalama Birim Fiyat: " & _
            Format(toplamTutar / toplamMiktar, "#,##0.00") & " TL"
    Else
        lblOrtalamaFiyat.Caption = _
            "Ortalama Birim Fiyat: 0 TL"
    End If
    
    MsgBox "Tedarikçi analizi tamamlandý.", _
           vbInformation, "Analiz Tamamlandý"

End Sub

Private Sub btnAra_Click()

    Dim wbKaynak As Workbook
    Dim ws As Worksheet
    Dim bulunan As Range
    Dim kalemNo As String

    If Trim(txtKalemNo.Value) = "" Then
        MsgBox "Lütfen aranacak Kalem No giriniz.", vbExclamation, "Eksik Bilgi"
        Exit Sub
    End If

    ' Þirket dosyasýný kontrol et
    On Error Resume Next
    Set wbKaynak = Workbooks("ARAÇ BAZLI MALZEME PLANLAMA_2026.xlsx")
    On Error GoTo 0

    If wbKaynak Is Nothing Then
        MsgBox "ARAÇ BAZLI MALZEME PLANLAMA_2026.xlsx dosyasý açýk deðil." & vbCrLf & _
               "Önce þirket dosyasýný açýnýz.", vbExclamation, "Dosya Bulunamadý"
        Exit Sub
    End If

    Set ws = wbKaynak.Worksheets("KALEM AYRINTI")

    kalemNo = Trim(txtKalemNo.Value)

    ' Kalem numarasýný C sütununda ara
    Set bulunan = ws.Columns("C").Find( _
        What:=kalemNo, _
        LookIn:=xlValues, _
        LookAt:=xlWhole)

    If bulunan Is Nothing Then
        MsgBox "Bu Kalem No KALEM AYRINTI sayfasýnda bulunamadý.", _
               vbExclamation, "Sonuç Yok"
        Exit Sub
    End If

    ' Þirket dosyasýndaki bilgiler
    cmbTedarikci.Value = ws.Cells(bulunan.Row, "R").Value
    txtBirimFiyat.Value = ws.Cells(bulunan.Row, "K").Value
    txtTarih.Value = ws.Cells(bulunan.Row, "H").Value
    
    ' KALEM AYRINTI'da miktar alaný olmadýðý için boþ býrak
    txtMiktar.Value = ""

    MsgBox "Kayýt þirket dosyasýndan bulundu.", vbInformation, "Arama Sonucu"

End Sub

Private Sub btnGuncelle_Click()

    Dim ws As Worksheet
    Dim bulunan As Range

    Set ws = ThisWorkbook.Worksheets("SATIN ALMA")

    If txtKalemNo.Value = "" Then
        MsgBox "Lütfen güncellenecek Kalem No giriniz.", vbExclamation, "Eksik Bilgi"
        Exit Sub
    End If

    Set bulunan = ws.Range("A:A").Find(What:=txtKalemNo.Value, LookAt:=xlWhole)

    If bulunan Is Nothing Then
        MsgBox "Güncellenecek kayýt bulunamadý.", vbExclamation, "Sonuç Yok"
        Exit Sub
    End If

    If cmbTedarikci.Value = "" Then
        MsgBox "Lütfen Tedarikçi giriniz.", vbExclamation, "Eksik Bilgi"
        Exit Sub
    End If

    If txtMiktar.Value = "" Or txtBirimFiyat.Value = "" Then
        MsgBox "Lütfen Miktar ve Birim Fiyat bilgilerini giriniz.", vbExclamation, "Eksik Bilgi"
        Exit Sub
    End If

    If Not IsNumeric(txtMiktar.Value) Or Not IsNumeric(txtBirimFiyat.Value) Then
        MsgBox "Miktar ve Birim Fiyat sayý olmalýdýr.", vbExclamation, "Hatalý Veri"
        Exit Sub
    End If

    If CDbl(txtMiktar.Value) <= 0 Or CDbl(txtBirimFiyat.Value) <= 0 Then
        MsgBox "Miktar ve Birim Fiyat 0'dan büyük olmalýdýr.", vbExclamation, "Hatalý Veri"
        Exit Sub
    End If

    If txtTarih.Value = "" Or Not IsDate(txtTarih.Value) Then
        MsgBox "Lütfen geçerli bir tarih giriniz.", vbExclamation, "Hatalý Tarih"
        Exit Sub
    End If

    ws.Cells(bulunan.Row, "B").Value = cmbTedarikci.Value
    ws.Cells(bulunan.Row, "C").Value = CDbl(txtMiktar.Value)
    ws.Cells(bulunan.Row, "D").Value = CDbl(txtBirimFiyat.Value)
    ws.Cells(bulunan.Row, "E").Value = CDate(txtTarih.Value)
    ws.Cells(bulunan.Row, "E").NumberFormat = "dd.mm.yyyy"
    
    ws.Cells(bulunan.Row, "F").Value = _
        CDbl(txtMiktar.Value) * CDbl(txtBirimFiyat.Value)

    ws.Cells(bulunan.Row, "F").NumberFormat = "#,##0.00"

    MsgBox "Kayýt baþarýyla güncellendi.", vbInformation, "Baþarýlý"

End Sub

Private Sub btnKalemAnaliz_Click()

    Dim wbKaynak As Workbook
    Dim ws As Worksheet
    Dim wsCalisma As Worksheet
    Dim kalemNo As String
    Dim bulunan As Range
    Dim sonSatir As Long
    Dim i As Long
    Dim satir As Long

    ' Kaynak dosyayý bul
    On Error Resume Next
    Set wbKaynak = Workbooks("ARAÇ BAZLI MALZEME PLANLAMA_2026.xlsx")
    On Error GoTo 0

    If wbKaynak Is Nothing Then
        MsgBox "ARAÇ BAZLI MALZEME PLANLAMA_2026.xlsx dosyasý açýk deðil.", _
               vbExclamation, "Kaynak Dosya Bulunamadý"
        Exit Sub
    End If

    ' KALEM AYRINTI sayfasý
    Set ws = wbKaynak.Worksheets("KALEM AYRINTI")

    ' CALISMA sayfasý
    Set wsCalisma = wbKaynak.Worksheets("CALISMA ")

    ' Kalem numarasý
    kalemNo = Trim(txtKalemNo.Value)

    If kalemNo = "" Then
        MsgBox "Lütfen Kalem No giriniz.", _
               vbExclamation, "Eksik Bilgi"
        Exit Sub
    End If

    '====================================================
    ' KALEM AYRINTI BÝLGÝLERÝ
    '====================================================
    Set bulunan = ws.Columns("C").Find( _
        What:=kalemNo, _
        LookIn:=xlValues, _
        LookAt:=xlWhole)
    If bulunan Is Nothing Then

        MsgBox "Bu kalem KALEM AYRINTI sayfasýnda bulunamadý.", _
               vbExclamation, "Kalem Bulunamadý"

    Else

        txtSonAlimTarihi.Value = Format( _
            ws.Cells(bulunan.Row, "H").Value, "dd.mm.yyyy")

        txtSonFiyat.Value = ws.Cells(bulunan.Row, "I").Value
        txtGuncelFiyat.Value = ws.Cells(bulunan.Row, "K").Value
        txtAnalizTedarikci.Value = ws.Cells(bulunan.Row, "R").Value
        If IsNumeric(ws.Cells(bulunan.Row, "I").Value) And _
           IsNumeric(ws.Cells(bulunan.Row, "K").Value) Then

            If CDbl(ws.Cells(bulunan.Row, "I").Value) <> 0 Then

                txtFiyatDegisimi.Value = _
                    Format( _
                    ((CDbl(ws.Cells(bulunan.Row, "K").Value) - _
                    CDbl(ws.Cells(bulunan.Row, "I").Value)) / _
                    CDbl(ws.Cells(bulunan.Row, "I").Value)) * 100, _
                    "0.00") & "%"

            Else
                txtFiyatDegisimi.Value = ""
            End If

        Else
            txtFiyatDegisimi.Value = ""
        End If
    End If

    '====================================================
    ' CALISMA BÝLGÝLERÝ
    '====================================================

    Set bulunan = wsCalisma.Columns("B").Find( _
        What:=kalemNo, _
        LookIn:=xlValues, _
        LookAt:=xlWhole)

    If bulunan Is Nothing Then

        MsgBox "Kalem CALISMA sayfasýnda bulunamadý.", _
               vbExclamation, "Kalem Bulunamadý"

        Exit Sub

    End If

    ' B = Kalem Kodu
    ' C = Kalem Tanýmý
    ' D = Tedarikçi
    ' E = En Son Alým Yapan Firma
    ' BG = Stok
    ' BP = Açýk SAS
    ' BQ = Kullaným Adedi
    ' BZ = En Son Alým Tarihi

    txtKalemTanimi.Value = wsCalisma.Cells(bulunan.Row, "C").Value

    txtStok.Value = wsCalisma.Cells(bulunan.Row, "BG").Value

    txtAcikSAS.Value = wsCalisma.Cells(bulunan.Row, "BP").Value

    txtKullanim.Value = wsCalisma.Cells(bulunan.Row, "BQ").Value

    ' Ýhtiyaç sütununu henüz belirlemediðimiz için boþ býrakýyoruz
    txtIhtiyac.Value = ""

    '====================================================
    ' TEDARÝKÇÝ KARÞILAÞTIRMA
    '====================================================

    lstTedarikciKarsilastirma.Clear

    lstTedarikciKarsilastirma.AddItem "Tedarikçi"

    lstTedarikciKarsilastirma.List(0, 1) = "Son Fiyat"

    lstTedarikciKarsilastirma.List(0, 2) = "Para Birimi"

    sonSatir = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    For i = 2 To sonSatir

        If Trim(ws.Cells(i, "C").Value) = kalemNo Then

            lstTedarikciKarsilastirma.AddItem ws.Cells(i, "R").Value

            satir = lstTedarikciKarsilastirma.ListCount - 1

            lstTedarikciKarsilastirma.List(satir, 1) = _
                ws.Cells(i, "I").Value

            lstTedarikciKarsilastirma.List(satir, 2) = _
                ws.Cells(i, "J").Value

        End If

    Next i
    '====================================================
    ' SATIN ALMA KARAR DESTEK ANALÝZÝ
    '====================================================

    Dim stok As Double
    Dim acikSAS As Double
    Dim fiyatDegisimi As Double

    stok = Val(txtStok.Value)
    acikSAS = Val(txtAcikSAS.Value)

    fiyatDegisimi = 0

    If IsNumeric(txtFiyatDegisimi.Value) Then
        fiyatDegisimi = Val(Replace(txtFiyatDegisimi.Value, "%", ""))
    End If

    ' Önce alaný temizle
    txtSatinAlmaOnerisi.Value = ""

    ' Stok ve açýk SAS deðerlendirmesi
    If stok <= 0 And acikSAS <= 0 Then

        txtSatinAlmaOnerisi.Value = _
            "ACÝL: Stok bulunmuyor ve açýk SAS yok." & vbCrLf & _
            "Satýn alma iþlemi deðerlendirilmelidir."

    ElseIf stok <= 0 And acikSAS > 0 Then

        txtSatinAlmaOnerisi.Value = _
            "DÝKKAT: Stok bulunmuyor ancak açýk SAS mevcut." & vbCrLf & _
            "Mevcut satýn alma sipariþleri kontrol edilmelidir."

    ElseIf stok > 0 And acikSAS <= 0 Then

        txtSatinAlmaOnerisi.Value = _
            "Stok mevcut ancak açýk SAS bulunmuyor." & vbCrLf & _
            "Yeni satýn alma ihtiyacý stok durumuna göre deðerlendirilmelidir."

    Else

        txtSatinAlmaOnerisi.Value = _
            "Stok ve açýk SAS mevcut." & vbCrLf & _
            "Yeni satýn alma öncesinde mevcut sipariþler ve stok seviyesi kontrol edilmelidir."

    End If

    ' Fiyat deðiþimini ayrýca deðerlendir
    If fiyatDegisimi < 0 Then

        txtSatinAlmaOnerisi.Value = _
            txtSatinAlmaOnerisi.Value & vbCrLf & _
            "Fiyat avantajý: Güncel fiyat önceki fiyata göre daha düþük."

    ElseIf fiyatDegisimi > 0 Then

        txtSatinAlmaOnerisi.Value = _
            txtSatinAlmaOnerisi.Value & vbCrLf & _
            "UYARI: Güncel fiyat önceki fiyata göre daha yüksek."

    Else

        txtSatinAlmaOnerisi.Value = _
            txtSatinAlmaOnerisi.Value & vbCrLf & _
            "Fiyat deðiþimi bulunmuyor."

    End If

End Sub

Private Sub btnKaydet_Click()

    Dim wbKaynak As Workbook
    Dim ws As Worksheet
    Dim wsCalisma As Worksheet
    Dim bulunan As Range
    Dim bulunanCalisma As Range
    Dim kalemNo As String
    Dim sonSatir As Long
    Dim sonSatirCalisma As Long

    '-----------------------------
    ' ZORUNLU ALAN KONTROLLERÝ
    '-----------------------------

    If Trim(txtKalemNo.Value) = "" Then
        MsgBox "Lütfen Kalem No giriniz.", vbExclamation, "Eksik Bilgi"
        Exit Sub
    End If

    If Trim(cmbTedarikci.Value) = "" Then
        MsgBox "Lütfen Tedarikçi giriniz.", vbExclamation, "Eksik Bilgi"
        Exit Sub
    End If

    If Trim(txtBirimFiyat.Value) = "" Then
        MsgBox "Lütfen Birim Fiyat giriniz.", vbExclamation, "Eksik Bilgi"
        Exit Sub
    End If

    If Not IsNumeric(txtBirimFiyat.Value) Then
        MsgBox "Birim Fiyat sadece sayý olmalýdýr.", vbExclamation, "Hatalý Veri"
        Exit Sub
    End If

    If CDbl(txtBirimFiyat.Value) <= 0 Then
        MsgBox "Birim Fiyat 0'dan büyük olmalýdýr.", vbExclamation, "Hatalý Veri"
        Exit Sub
    End If

    If Trim(txtTarih.Value) = "" Then
        MsgBox "Lütfen Satýn Alma Tarihi giriniz.", vbExclamation, "Eksik Bilgi"
        Exit Sub
    End If

    If Not IsDate(txtTarih.Value) Then
        MsgBox "Lütfen geçerli bir tarih giriniz.", vbExclamation, "Hatalý Tarih"
        Exit Sub
    End If

    '-----------------------------
    ' ÞÝRKET DOSYASINI BUL
    '-----------------------------

    On Error Resume Next
    Set wbKaynak = Workbooks("ARAÇ BAZLI MALZEME PLANLAMA_2026.xlsx")
    On Error GoTo 0

    If wbKaynak Is Nothing Then
        MsgBox "ARAÇ BAZLI MALZEME PLANLAMA_2026.xlsx açýk deðil." & vbCrLf & _
               "Önce þirket dosyasýný açýnýz.", _
               vbExclamation, "Dosya Bulunamadý"
        Exit Sub
    End If

    Set ws = wbKaynak.Worksheets("KALEM AYRINTI")
    Set wsCalisma = wbKaynak.Worksheets("CALISMA ")

    kalemNo = Trim(txtKalemNo.Value)

    '-----------------------------
    ' KALEM AYRINTI'DA KALEMÝ ARA
    '-----------------------------

    Set bulunan = ws.Columns("C").Find( _
        What:=kalemNo, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False)

    '-----------------------------
    ' KALEM VARSA GÜNCELLE
    '-----------------------------

    If Not bulunan Is Nothing Then

        ws.Cells(bulunan.Row, "H").Value = CDate(txtTarih.Value)
        ws.Cells(bulunan.Row, "H").NumberFormat = "dd.mm.yyyy"

        ws.Cells(bulunan.Row, "K").Value = CDbl(txtBirimFiyat.Value)

        ws.Cells(bulunan.Row, "R").Value = cmbTedarikci.Value

    '-----------------------------
    ' KALEM YOKSA YENÝ EKLE
    '-----------------------------

    Else

        sonSatir = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row + 1

        ws.Cells(sonSatir, "C").Value = kalemNo

        ws.Cells(sonSatir, "H").Value = CDate(txtTarih.Value)
        ws.Cells(sonSatir, "H").NumberFormat = "dd.mm.yyyy"

        ws.Cells(sonSatir, "K").Value = CDbl(txtBirimFiyat.Value)

        ws.Cells(sonSatir, "R").Value = cmbTedarikci.Value

    End If

    '--------------------------------
    ' CALISMA'YA SADECE KALEM KODU
    ' EKLE
    '--------------------------------

    Set bulunanCalisma = wsCalisma.Columns("B").Find( _
        What:=kalemNo, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False)

    ' Kalem CALISMA'da yoksa B sütununa ekle
    If bulunanCalisma Is Nothing Then

        sonSatirCalisma = wsCalisma.Cells( _
            wsCalisma.Rows.Count, "B").End(xlUp).Row + 1

        wsCalisma.Cells(sonSatirCalisma, "B").Value = kalemNo

    End If

    '-----------------------------
    ' ÞÝRKET DOSYASINI KAYDET
    '-----------------------------

    wbKaynak.Save

    MsgBox "Kayýt baþarýyla tamamlandý." & vbCrLf & _
           "KALEM AYRINTI güncellendi ve CALISMA kontrol edildi.", _
           vbInformation, "Baþarýlý"

End Sub
Private Sub btnSil_Click()

    Dim ws As Worksheet
    Dim bulunan As Range
    Dim cevap As VbMsgBoxResult

    Set ws = ThisWorkbook.Worksheets("SATIN ALMA")

    If txtKalemNo.Value = "" Then
        MsgBox "Lütfen silinecek Kalem No giriniz.", vbExclamation, "Eksik Bilgi"
        Exit Sub
    End If

    Set bulunan = ws.Range("A:A").Find( _
        What:=txtKalemNo.Value, _
        LookAt:=xlWhole)

    If bulunan Is Nothing Then
        MsgBox "Silinecek kayýt bulunamadý.", vbExclamation, "Sonuç Yok"
        Exit Sub
    End If

    cevap = MsgBox( _
        "Bu kaydý silmek istediðinize emin misiniz?", _
        vbYesNo + vbQuestion, _
        "Kayýt Silme")

    If cevap = vbNo Then
        Exit Sub
    End If

    ws.Rows(bulunan.Row).Delete

    MsgBox "Kayýt baþarýyla silindi.", vbInformation, "Baþarýlý"

    txtKalemNo.Value = ""
    cmbTedarikci.Value = ""
    txtMiktar.Value = ""
    txtBirimFiyat.Value = ""
    txtTarih.Value = ""

    lblToplamTutar.Caption = "Toplam Tutar: 0 TL"

End Sub

Private Sub btnTemizle_Click()
    txtKalemNo.Value = ""
    cmbTedarikci.Value = ""
    txtMiktar.Value = ""
    txtBirimFiyat.Value = ""
    txtTarih.Value = ""
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub txtBirimFiyat_Change()
    If IsNumeric(txtMiktar.Value) And IsNumeric(txtBirimFiyat.Value) Then

        lblToplamTutar.Caption = _
            "Toplam Tutar: " & Format(CDbl(txtMiktar.Value) * CDbl(txtBirimFiyat.Value), "#,##0.00") & " TL"

    Else

        lblToplamTutar.Caption = "Toplam Tutar: 0 TL"

    End If

End Sub

Private Sub txtMiktar_Change()
    If IsNumeric(txtMiktar.Value) And IsNumeric(txtBirimFiyat.Value) Then

        lblToplamTutar.Caption = _
            "Toplam Tutar: " & Format(CDbl(txtMiktar.Value) * CDbl(txtBirimFiyat.Value), "#,##0.00") & " TL"

    Else

        lblToplamTutar.Caption = "Toplam Tutar: 0 TL"

    End If

End Sub

Private Sub UserForm_Initialize()

    Dim ws As Worksheet
    Dim sonSatir As Long
    Dim i As Long
    Dim tedarikci As String
    Dim dict As Object

    Set ws = ThisWorkbook.Worksheets("SATIN ALMA")
    Set dict = CreateObject("Scripting.Dictionary")

    sonSatir = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row

    For i = 2 To sonSatir

        tedarikci = Trim(ws.Cells(i, "G").Value)

        If tedarikci <> "" Then

            If Not dict.Exists(tedarikci) Then
                dict.Add tedarikci, True
                cmbTedarikci.AddItem tedarikci
                cmbRaporTedarikci.AddItem tedarikci
            End If

        End If

    Next i

End Sub
Private Function HeaderColumn(ws As Worksheet, headerName As String) As Long

    Dim r As Long
    Dim c As Long
    Dim sonSutun As Long
    Dim hucreDegeri As String

    For r = 1 To 20

        sonSutun = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column

        For c = 1 To sonSutun

            If Not IsError(ws.Cells(r, c).Value) Then

                hucreDegeri = CStr(ws.Cells(r, c).Value)

                If Trim$(LCase$(hucreDegeri)) = Trim$(LCase$(headerName)) Then

                    HeaderColumn = c
                    Exit Function

                End If

            End If

        Next c

    Next r

    HeaderColumn = 0

End Function
Private Function GetCellByHeader(ws As Worksheet, satir As Long, headerName As String) As Variant

    Dim sutun As Long

    sutun = HeaderColumn(ws, headerName)

    If sutun = 0 Then
        GetCellByHeader = ""
    Else
        GetCellByHeader = ws.Cells(satir, sutun).Value
    End If

End Function
