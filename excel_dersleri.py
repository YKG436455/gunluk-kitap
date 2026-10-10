# -*- coding: utf-8 -*-
"""
📊 GÜNLÜK EXCEL DERSİ SİSTEMİ
45 ders: Excel'e giriş → Formüller → Biçimlendirme → Mantık → Pivot Tablo → Dashboard
"""

import os
import smtplib
import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

MAIL_USER = os.environ.get("MAIL_USER")
MAIL_PASS = os.environ.get("MAIL_PASS")
MAIL_TO = os.environ.get("MAIL_TO")

# ============================================================
# 45 DERSLİK EXCEL PROGRAMI
# ============================================================
EXCEL_DERSLERI = [
    # === HAFTA 1: EXCEL'E GİRİŞ ===
    {"baslik": "Ders 1: Excel Ekranı ve İlk Tablon",
     "ozet": "Excel'in temel yapı taşları: hücre, satır, sütun, sayfa, çalışma kitabı. Bunları anlarsan gerisi kolay.",
     "uygulama": "Boş Excel aç. A1'den başlayarak şu tabloyu gir:\nMağaza\tCiro\tMüşteri\nSaadet\t120000\t500\nÇayırlı\t95000\t420\nHastanebaşı\t150000\t610\nDosyayı 'Excel_Derslerim.xlsx' olarak kaydet.",
     "videolar": ["Excel'e Giriş Türkçe Ders 1", "Excel for Beginners Full Course"],
     "ipucu": "Hücre adresi = Sütun harfi + Satır numarası. Örnek: A1, B3, D10.",
     "sik_hata": "Verileri rastgele hücrelere yazmak. Her zaman A1'den başla."},

    {"baslik": "Ders 2: Hücre ve Aralık Kavramları",
     "ozet": "Hücre adreslerini ve A1:C4 gibi aralıkları öğreniyoruz.",
     "uygulama": "Ders 1'deki tablonda A1:C4 aralığını seç. Ad Kutusu'nda ne yazdığına bak. Sonra A1 hücresine tıkla, sadece A1'in seçildiğini gör.",
     "videolar": ["Excel hücre adresleri aralık seçme", "Excel Cell References"],
     "ipucu": "Aralık seçmek için başlangıç hücresinden bitiş hücresine sürükle.",
     "sik_hata": "Aralığı elle yazarken yanlış yazmak. Ad Kutusu'na 'A1:C4' yazıp Enter'a basmak en garantili yoldur."},

    {"baslik": "Ders 3: Veri Türleri ve Sayı Biçimleri",
     "ozet": "Excel metin, sayı, tarih türlerini otomatik algılar. Türkçe Excel'de ondalık ayırıcı virgüldür (,).",
     "uygulama": "Tablona 'Tarih' sütunu ekle. Bugünün tarihini '10.10.2026' olarak gir. Excel'in bunu tarih olarak algıladığını sağa yaslanmasından anla.",
     "videolar": ["Excel veri türleri sayı biçimleri", "Excel Data Types"],
     "ipucu": "Sayıları sade yaz (500000). Biçimi sonra ver (500.000).",
     "sik_hata": "Sayıyı '500.000' diye yazmak. Excel bunu tarih veya metin sanabilir."},

    {"baslik": "Ders 4: Sütun Genişlikleri ve Otomatik Doldurma",
     "ozet": "Sütunları içeriğe göre genişletmeyi ve Excel'in otomatik doldurma özelliğini öğreniyoruz.",
     "uygulama": "Sütun başlıkları arasındaki çizgiye çift tıkla. A2'ye 'Mağaza 1' yaz, sağ alt köşedeki küçük kareden aşağı sürükle. Excel 'Mağaza 2, Mağaza 3...' diye devam eder.",
     "videolar": ["Excel sütun genişliği otomatik doldurma", "Excel Autofill"],
     "ipucu": "Sayıları binlik ayırıcıyla göstermek için C2:C4 seç, 'Binlik Ayırıcı' düğmesine (000) bas.",
     "sik_hata": "Otomatik doldurma karesini görmeyip elle yazmak. Sağ alt köşedeki küçük kareyi kullan."},

    {"baslik": "Ders 5: İlk Formülün ve Kaydetme",
     "ozet": "Formüller her zaman '=' ile başlar. İlk formülünü yazıp dosyanı kaydediyoruz.",
     "uygulama": "E1'e 'Fark' yaz. E2'ye '=D2-C2' yaz. Enter'a bas. E2'yi E4'e kadar sürükle. Ctrl+S ile kaydet.",
     "videolar": ["Excel ilk formül toplama çıkarma", "Excel First Formula"],
     "ipucu": "Hata yaparsan Ctrl+Z ile geri al.",
     "sik_hata": "'=' işaretini unutmak. Formül değil, metin olur."},

    # === HAFTA 2: TEMEL FORMÜLLER ===
    {"baslik": "Ders 6: TOPLA ve ORTALAMA Fonksiyonları",
     "ozet": "Tek tek toplamak yerine hazır fonksiyonlar kullanıyoruz.",
     "uygulama": "B5'e '=TOPLA(B2:B4)' yaz. C5'e '=ORTALAMA(C2:C4)' yaz. Sonuçları kontrol et.",
     "videolar": ["Excel TOPLA ORTALAMA fonksiyonları", "Excel SUM AVERAGE"],
     "ipucu": "Alt + = kısayolu otomatik TOPLA formülü ekler.",
     "sik_hata": "TOPLA yerine tek tek hücreleri toplamak: =B2+B3+B4. Fonksiyon daha kısa ve hatasızdır."},

    {"baslik": "Ders 7: Yüzde Hesaplama ve Hedef Gerçekleşme",
     "ozet": "Bir sayının diğerine oranını (yüzde) hesaplıyoruz.",
     "uygulama": "Tabloya 'Hedef' ve 'Gerçekleşen' sütunları ekle. 'Gerçekleşme %' sütununa '=Gerçekleşen/Hedef' yaz. Sonucu % biçiminde göster.",
     "videolar": ["Excel yüzde hesaplama hedef gerçekleşme", "Excel Percentage"],
     "ipucu": "Yüzde biçimi için hücreleri seç, Giriş sekmesindeki '%' düğmesine bas.",
     "sik_hata": "Yüzdelerin ortalamasını almak. Doğrusu: toplam gerçekleşme / toplam hedef."},

    {"baslik": "Ders 8: Mutlak Başvuru ve $ İşareti",
     "ozet": "Formülü aşağı sürüklerken bir hücrenin sabit kalmasını sağlar.",
     "uygulama": "G2'ye '=B2/$B$5' yaz. Aşağı sürükle. Her mağazanın toplam ciro içindeki payını gör.",
     "videolar": ["Excel mutlak başvuru dolar işareti", "Excel Absolute Reference"],
     "ipucu": "Hücre adresini yazdıktan sonra F4 tuşuna basarak $ işaretlerini ekleyebilirsin.",
     "sik_hata": "$ işaretini unutmak. Formül kayar, yanlış sonuç verir (#BÖL/0! hatası)."},

    {"baslik": "Ders 9: MİN, MAK ve SAY Fonksiyonları",
     "ozet": "En küçük, en büyük değeri bulmayı ve sayı saymayı öğreniyoruz.",
     "uygulama": "En düşük ve en yüksek ciroyu MİN ve MAK ile bul. Kaç mağaza olduğunu SAY ile hesapla.",
     "videolar": ["Excel MİN MAK SAY fonksiyonları", "Excel MIN MAX COUNT"],
     "ipucu": "Formül yazarken '=MİN(' yazıp fareyle aralığı seçebilirsin.",
     "sik_hata": "SAY yerine DOLUSAY kullanmak. SAY sadece sayıları sayar, DOLUSAY tüm dolu hücreleri sayar."},

    {"baslik": "Ders 10: Formül Hatalarını Anlama",
     "ozet": "Sık karşılaşılan hata kodları ve nedenleri.",
     "uygulama": "'=B2/0' yaz, #BÖL/0! hatasını gör. '=B2/\"\"' yaz, #DEĞER! hatasını gör.",
     "videolar": ["Excel formül hataları #BÖL/0 #DEĞER", "Excel Formula Errors"],
     "ipucu": "Hata veren hücreyi seçip formül çubuğuna bak, sorunu görürsün.",
     "sik_hata": "Hata kodunu görmezden gelmek. Hatanın nedenini anlamadan düzeltmeye çalışmak zaman kaybıdır."},

    # === HAFTA 3: BİÇİMLENDİRME ===
    {"baslik": "Ders 11: Başlıkları Biçimlendirme ve Tablo",
     "ozet": "Başlık satırını kalın yapma, renk verme ve Ctrl+T ile Excel Tablosu oluşturma.",
     "uygulama": "A1:C1'i seç, kalın yap, dolgu rengi ver. Ctrl+T ile Excel Tablosu'na dönüştür.",
     "videolar": ["Excel tablo oluşturma biçimlendirme", "Excel Format as Table"],
     "ipucu": "Excel Tablosu'nda her sütun başlığında otomatik filtre ok belirir.",
     "sik_hata": "Başlık satırı olmadan tablo oluşturmak. Başlık yoksa Excel sütunları karıştırır."},

    {"baslik": "Ders 12: Sıralama ve Filtreleme",
     "ozet": "Verileri büyükten küçüğe sıralama ve kriterlere göre filtreleme.",
     "uygulama": "Ciro sütununu büyükten küçüğe sırala. Filtre oklarını kullanarak sadece 'İlkadım' bölgesini göster.",
     "videolar": ["Excel sıralama filtreleme", "Excel Sort Filter"],
     "ipucu": "Filtreyi temizlemek için filtre okuna tıkla, 'Filtreyi Temizle' seç.",
     "sik_hata": "Tabloyu seçmeden sıralama yapmak. Sadece bir sütun sıralanır, diğerleri karışır."},

    {"baslik": "Ders 13: Koşullu Biçimlendirme",
     "ozet": "Belirli koşulu sağlayan hücreleri otomatik renklendirme.",
     "uygulama": "Ciro sütununu seç. Koşullu Biçimlendirme → Küçüktür → 100000 → Kırmızı. 100.000 altı kırmızı olur.",
     "videolar": ["Excel koşullu biçimlendirme", "Excel Conditional Formatting"],
     "ipucu": "Renk yerine veri çubukları, renk ölçekleri veya simge kümeleri de ekleyebilirsin.",
     "sik_hata": "Aralığı seçmeden koşullu biçimlendirme uygulamak. Yanlış hücreler renklenir."},

    {"baslik": "Ders 14: Kenarlıklar ve Yazı Tipi",
     "ozet": "Tabloyu profesyonel göstermek için kenarlıklar ve yazı tipi ayarları.",
     "uygulama": "Tüm hücreleri seç. Kenarlıklar → Tüm Kenarlıklar. Yazı tipini Calibri 11 yap.",
     "videolar": ["Excel kenarlıklar yazı tipi ayarları", "Excel Borders Fonts"],
     "ipucu": "Ctrl+1 ile tüm biçimlendirme ayarlarına tek yerden ulaşabilirsin.",
     "sik_hata": "Rastgele renk ve yazı tipi kullanmak. Sade ve tutarlı olmak daha profesyonel görünür."},

    {"baslik": "Ders 15: Hafta Sonu Mini Proje",
     "ozet": "Bu hafta öğrendiklerini birleştirerek mağaza performans raporu hazırla.",
     "uygulama": "10 mağazalık tablo oluştur. Biçimlendir. Ciroya göre sırala. 100.000 altı kırmızı yap. Filtrele.",
     "videolar": ["Excel mini proje mağaza performans raporu", "Excel Mini Project"],
     "ipucu": "Bu proje, bir yönetici için temel rapor hazırlama becerisini kazandırır.",
     "sik_hata": "Adımları atlamak. Her adımı sırayla yap, sonucu kontrol et."},

    # === HAFTA 4: MANTIK FONKSİYONLARI ===
    {"baslik": "Ders 16: EĞER Fonksiyonu",
     "ozet": "Koşula göre iki farklı sonuç döndürme.",
     "uygulama": "Tabloya 'Durum' sütunu ekle. '=EĞER(B2>=100000;\"Başarılı\";\"Başarısız\")' yaz.",
     "videolar": ["Excel EĞER fonksiyonu", "Excel IF Function"],
     "ipucu": "Türkçe Excel'de formül parçaları noktalı virgül (;) ile ayrılır.",
     "sik_hata": "Virgül (,) kullanmak. Türkçe Excel'de hata verir."},

    {"baslik": "Ders 17: VE / YADA ile Çoklu Koşul",
     "ozet": "Birden fazla koşulu aynı anda kontrol etme.",
     "uygulama": "'=EĞER(VE(B2>=100000;C2>=500);\"Tam Başarı\";\"Kısmi\")' yaz.",
     "videolar": ["Excel EĞER VE YADA fonksiyonları", "Excel IF AND OR"],
     "ipucu": "VE: Tüm koşullar doğru olmalı. YADA: En az biri doğru olmalı.",
     "sik_hata": "VE ile YADA'yı karıştırmak. Yanlış sonuç verir."},

    {"baslik": "Ders 18: EĞERHATA Fonksiyonu",
     "ozet": "Hata mesajları yerine anlamlı bir şey gösterme.",
     "uygulama": "'=EĞERHATA(B2/0;\"Hesaplanamadı\")' yaz.",
     "videolar": ["Excel EĞERHATA fonksiyonu", "Excel IFERROR"],
     "ipucu": "Özellikle DÜŞEYARA'da aranan değer bulunamadığında (#YOK) çok işe yarar.",
     "sik_hata": "EĞERHATA yerine EĞER kullanmak. Hatayı yakalayamazsın."},

    {"baslik": "Ders 19: İç İçe EĞER",
     "ozet": "Birden fazla koşulu sırayla kontrol etme.",
     "uygulama": "'=EĞER(B2>=150000;\"Çok İyi\";EĞER(B2>=100000;\"İyi\";\"Geliştirilmeli\"))' yaz.",
     "videolar": ["Excel iç içe EĞER kullanımı", "Excel Nested IF"],
     "ipucu": "Her yeni EĞER bir parantez daha gerektirir. Parantez sayısına dikkat et.",
     "sik_hata": "Parantez eksikliği. #AD? hatası verir."},

    {"baslik": "Ders 20: Hafta Sonu Mantık Uygulaması",
     "ozet": "Bu hafta öğrendiklerini birleştirerek otomatik değerlendirme yap.",
     "uygulama": "Mağazaları ciro ve müşteri sayısına göre 'Yıldız', 'İyi', 'Geliştirilmeli' olarak sınıflandır.",
     "videolar": ["Excel mantık fonksiyonları uygulama", "Excel Logic Functions"],
     "ipucu": "Bu hafta öğrendiklerin, raporlama ve karar vermede en çok kullanacağın becerilerdir.",
     "sik_hata": "Sınıflandırma kriterlerini net belirlememek. Önce kuralları yaz, sonra formülü."},

    # === HAFTA 5: KOŞULLU TOPLAMA ===
    {"baslik": "Ders 21: ETOPLA ile Koşullu Toplama",
     "ozet": "Sadece belirli koşulu sağlayan hücreleri toplama.",
     "uygulama": "'=ETOPLA(Bölge_Sütunu;\"Kuzey\";Ciro_Sütunu)' yaz.",
     "videolar": ["Excel ETOPLA fonksiyonu", "Excel SUMIF"],
     "ipucu": "ETOPLA'nın İngilizce karşılığı SUMIF'tir.",
     "sik_hata": "Koşul aralığı ile toplam aralığını karıştırmak."},

    {"baslik": "Ders 22: ÇOKETOPLA ile Çoklu Koşullu Toplama",
     "ozet": "Birden fazla koşulu sağlayan hücreleri toplama.",
     "uygulama": "'=ÇOKETOPLA(Ciro;Bölge;\"Kuzey\";Ciro;\">100000\")' yaz.",
     "videolar": ["Excel ÇOKETOPLA fonksiyonu", "Excel SUMIFS"],
     "ipucu": "Koşul çiftlerini istediğin kadar çoğaltabilirsin.",
     "sik_hata": "Koşul sırasını karıştırmak. Önce aralık, sonra koşul yazılır."},

    {"baslik": "Ders 23: EĞERSAY ve ÇOKEĞERSAY",
     "ozet": "Belirli koşulu sağlayan hücreleri sayma.",
     "uygulama": "'=EĞERSAY(Durum;\"Başarılı\")' yaz. Kaç mağaza başarılı oldu?",
     "videolar": ["Excel EĞERSAY ÇOKEĞERSAY", "Excel COUNTIF COUNTIFS"],
     "ipucu": "Metin, sayı veya tarih içeren hücreleri sayabilir.",
     "sik_hata": "Tırnak işareti koymayı unutmak. Metin koşullar tırnak içinde olmalı."},

    {"baslik": "Ders 24: Bölge Bazında Özet Tablo",
     "ozet": "Bölge bazında toplam ciro ve mağaza sayısını özetleyen tablo.",
     "uygulama": "Bölgeleri listele. Her bölge için ETOPLA ve EĞERSAY kullan.",
     "videolar": ["Excel bölge bazında özet tablo", "Excel Regional Summary"],
     "ipucu": "Bu tablo, yöneticilerin en sık ihtiyaç duyduğu rapordur.",
     "sik_hata": "Bölge isimlerini elle yazmak. Büyük-küçük harf tutarsızlığı hata verir."},

    {"baslik": "Ders 25: Hafta Sonu Raporlama",
     "ozet": "Her bölgenin toplam cirosunu, ortalamasını ve başarı oranını hesapla.",
     "uygulama": "Bölge bazında kapsamlı bir rapor oluştur.",
     "videolar": ["Excel koşullu toplama raporu", "Excel Summary Report"],
     "ipucu": "Bu raporu her ay güncelleyerek bölge performansını takip edebilirsin.",
     "sik_hata": "Formülleri sabit değerle karıştırmak. Formüller veri değişince otomatik güncellenir."},

    # === HAFTA 6: VERİ ARAMA ===
    {"baslik": "Ders 26: DÜŞEYARA ile Veri Arama",
     "ozet": "Bir değeri kullanarak başka tablodan bilgi getirme.",
     "uygulama": "'=DÜŞEYARA(Aranan_Kod;Bilgi_Tablosu;2;YANLIŞ)' yaz.",
     "videolar": ["Excel DÜŞEYARA VLOOKUP", "Excel VLOOKUP Tutorial"],
     "ipucu": "Son parametre YANLIŞ (FALSE) tam eşleşme için kullanılır.",
     "sik_hata": "Son parametreyi yazmamak. Excel yaklaşık eşleşme yapar, yanlış sonuç verir."},

    {"baslik": "Ders 27: ÇAPRAZARA (XLOOKUP)",
     "ozet": "DÜŞEYARA'nın daha esnek hali. Excel 2021 ve 365'te var.",
     "uygulama": "'=ÇAPRAZARA(Aranan_Deger;Arama_Dizisi;Sonuç_Dizisi)' yaz.",
     "videolar": ["Excel ÇAPRAZARA XLOOKUP", "Excel XLOOKUP Tutorial"],
     "ipucu": "ÇAPRAZARA sola doğru da arama yapabilir.",
     "sik_hata": "Eski Excel'de ÇAPRAZARA yoksa DÜŞEYARA kullan."},

    {"baslik": "Ders 28: Veri Doğrulama",
     "ozet": "Hücrelere sadece belirli veri girilmesini sağlama.",
     "uygulama": "Bölge sütunu için açılır liste oluştur (Veri → Veri Doğrulama → Liste).",
     "videolar": ["Excel veri doğrulama açılır liste", "Excel Data Validation"],
     "ipucu": "Paylaşılan tablolarda veri tutarlılığı için şarttır.",
     "sik_hata": "Listeyi manuel yazmak yerine başka hücreden almak. Kaynak değişince liste güncellenmez."},

    {"baslik": "Ders 29: İNDİS ve KAÇINCI",
     "ozet": "DÜŞEYARA'nın esnek alternatifi.",
     "uygulama": "'=İNDİS(Sonuç_Aralığı;KAÇINCI(Aranan;Arama_Aralığı;0))' yaz.",
     "videolar": ["Excel İNDİS KAÇINCI", "Excel INDEX MATCH"],
     "ipucu": "İNDİS+KAÇINCI, DÜŞEYARA'dan daha esnektir ama öğrenmesi biraz zordur.",
     "sik_hata": "KAÇINCI'nın son parametresini 0 yazmamak. Yaklaşık eşleşme yapar."},

    {"baslik": "Ders 30: Hafta Sonu Arama Uygulaması",
     "ozet": "Farklı tablolardaki verileri ilişkilendirme.",
     "uygulama": "Mağaza, ürün, satış listesi oluştur. DÜŞEYARA ile mağaza ve ürün adlarını getir.",
     "videolar": ["Excel DÜŞEYARA uygulama örneği", "Excel VLOOKUP Practice"],
     "ipucu": "Bu sistem, gerçek iş hayatında veri tabanları arasında ilişki kurar.",
     "sik_hata": "Arama sütununun en solda olmaması. DÜŞEYARA hep soldan sağa arar."},

    # === HAFTA 7: PIVOT TABLO (GENİŞLETİLMİŞ) ===
    {"baslik": "Ders 31: Pivot Tablo Nedir?",
     "ozet": "Excel'in en güçlü aracı. Büyük veriyi saniyeler içinde özetler. Sürükle-bırak ile çalışır.",
     "uygulama": "Mağaza tablondan Pivot Tablo oluştur: Ekle → Özet Tablo → Tamam. 'Bölge'yi Satırlar'a, 'Ciro'yu Değerler'e sürükle.",
     "videolar": ["Excel Pivot Tablo nedir Türkçe", "Excel Pivot Table for Beginners"],
     "ipucu": "Pivot Tablo orijinal veriyi değiştirmez, sadece özetini gösterir.",
     "sik_hata": "Boş hücreli veriyle Pivot Tablo oluşturmak. Hata verir."},

    {"baslik": "Ders 32: Pivot Tablo Alanları",
     "ozet": "4 kutu: Filtreler, Satırlar, Sütunlar, Değerler. Her birinin işlevi farklı.",
     "uygulama": "'Ay' alanını Sütunlar'a sürükle. Bölge × Ay çapraz tablosu oluşur.",
     "videolar": ["Excel Pivot Tablo alanları", "Excel Pivot Table Fields"],
     "ipucu": "Aynı alanı birden fazla kutuya sürükleyebilirsin.",
     "sik_hata": "Metin alanını Değerler'e koymak. Excel onu sayar."},

    {"baslik": "Ders 33: Değer Alanı Ayarları",
     "ozet": "Toplam yerine ortalama, yüzde, sayı gibi hesaplamalar yapma.",
     "uygulama": "Değer hücresine sağ tıkla → Değer Alanı Ayarları → Ortalama seç.",
     "videolar": ["Excel Pivot Tablo ortalama yüzde", "Excel Pivot Table Value Settings"],
     "ipucu": "'Değerleri Farklı Göster' ile yüzdesel pay gösterebilirsin.",
     "sik_hata": "Değer alanı ayarlarını bulamamak. Sağ tık menüsünde."},

    {"baslik": "Ders 34: Çok Seviyeli Gruplama",
     "ozet": "Bölge → Mağaza şeklinde iç içe listeleme.",
     "uygulama": "'Mağaza' alanını Satırlar'a sürükle (Bölge'nin altına).",
     "videolar": ["Excel Pivot Tablo çok seviyeli", "Excel Pivot Table Grouping"],
     "ipucu": "Sol taraftaki [+] işaretiyle grupları açıp kapatabilirsin.",
     "sik_hata": "Grupları kapatmayı unutmak. Uzun listelerde karmaşık görünür."},

    {"baslik": "Ders 35: Tarih Gruplama",
     "ozet": "Tarihleri yıl, çeyrek, ay bazında otomatik gruplama.",
     "uygulama": "'Tarih' alanını Satırlar'a sürükle. Excel otomatik gruplar.",
     "videolar": ["Excel Pivot Tablo tarih gruplama", "Excel Pivot Date Grouping"],
     "ipucu": "İstediğin seviyeye kadar açıp kapatabilirsin.",
     "sik_hata": "Tarih sütununun metin olarak girilmesi. Tarih türü olmalı."},

    {"baslik": "Ders 36: Dilimleyici (Slicer)",
     "ozet": "Butonlarla interaktif filtreleme.",
     "uygulama": "Pivot Tablo Analiz → Dilimleyici Ekle → Bölge seç.",
     "videolar": ["Excel Pivot Tablo dilimleyici slicer", "Excel Slicer Tutorial"],
     "ipucu": "Bir dilimleyiciyi birden fazla Pivot Tablo'ya bağlayabilirsin.",
     "sik_hata": "Her tabloya ayrı dilimleyici eklemek. Yer kaplar."},

    {"baslik": "Ders 37: Pivot Grafik",
     "ozet": "Pivot Tablo'yu grafiğe dönüştürme.",
     "uygulama": "Pivot Tablo Analiz → Pivot Grafik → Sütun Grafiği seç.",
     "videolar": ["Excel Pivot Grafik", "Excel Pivot Chart"],
     "ipucu": "Tablo değişince grafik otomatik güncellenir.",
     "sik_hata": "Grafik türünü yanlış seçmek. Verine uygun tür önemli."},

    {"baslik": "Ders 38: Mağaza Performans Analizi",
     "ozet": "Pivot Tablo ile gerçek iş sorularına cevap verme.",
     "uygulama": "Hangi bölge en yüksek? Hangi mağaza hedefi tuttu? Aylık trend? Sorularını cevapla.",
     "videolar": ["Excel Pivot Tablo mağaza performans", "Excel Retail Pivot"],
     "ipucu": "Dilimleyici ile interaktif hale getir.",
     "sik_hata": "Tek boyutlu bakmak. Çok boyutlu düşün."},

    {"baslik": "Ders 39: Pivot Tablo Dashboard",
     "ozet": "Birden fazla Pivot Tablo ve grafiği tek sayfada birleştirme.",
     "uygulama": "3 Pivot Tablo + 1 dilimleyici + 2 grafik ile dashboard oluştur.",
     "videolar": ["Excel Pivot Tablo dashboard", "Excel Dashboard Tutorial"],
     "ipucu": "Aynı dilimleyiciyi tüm tablolara bağla, tek tıkla hepsi filtrelensin.",
     "sik_hata": "Dashboard'a gereksiz bilgi doldurmak. Sade ve net ol."},

    {"baslik": "Ders 40: Hafta Sonu Pivot Projesi",
     "ozet": "Öğrendiklerini birleştirerek profesyonel rapor hazırla.",
     "uygulama": "62 mağazalık veri ile bölge × ay × mağaza analizi yap. Dashboard hazırla.",
     "videolar": ["Excel Pivot Tablo proje", "Excel Pivot Project"],
     "ipucu": "Bu proje bittiğinde 5 dakikada rapor hazırlayabilirsin.",
     "sik_hata": "Adımları atlamak. Sırayla git."},

    # === HAFTA 8: VERİ TEMİZLEME ===
    {"baslik": "Ders 41: Metin Fonksiyonları",
     "ozet": "KIRP, SOLDAN, SAĞDAN, UZUNLUK ile metin temizleme.",
     "uygulama": "'=KIRP(A1)' ile boşlukları temizle. '=SOLDAN(A1;5)' ile ilk 5 karakteri al.",
     "videolar": ["Excel metin fonksiyonları KIRP SOLDAN", "Excel Text Functions"],
     "ipucu": "Başka sistemden kopyaladığın verilerdeki görünmez boşlukları temizler.",
     "sik_hata": "KIRP yerine elle boşluk silmek. Zaman kaybı."},

    {"baslik": "Ders 42: Tarih Fonksiyonları",
     "ozet": "BUGÜN, YIL, AY, GÜN ile tarihlerle çalışma.",
     "uygulama": "'=BUGÜN()' yaz. '=YIL(A1)' ile yılı çıkar.",
     "videolar": ["Excel tarih fonksiyonları BUGÜN", "Excel Date Functions"],
     "ipucu": "Tarih fonksiyonları dosya her açıldığında güncellenir.",
     "sik_hata": "Tarihleri metin olarak girmek. Hesaplanamaz."},

    {"baslik": "Ders 43: METNEÇEVİR Fonksiyonu",
     "ozet": "Sayı ve tarihleri özel formatlı metne çevirme.",
     "uygulama": "'=METNEÇEVİR(A1;\"0000\")' yaz. 123 → 0123 olur.",
     "videolar": ["Excel METNEÇEVİR TEXT", "Excel TEXT Function"],
     "ipucu": "Rapor başlıklarında tarih ve sayı birleştirmek için kullanılır.",
     "sik_hata": "Türkçe format kodlarını İngilizce yazmak. Farklıdır."},

    {"baslik": "Ders 44: Haftalık Performans Paneli",
     "ozet": "Otomatik güncellenen haftalık rapor hazırlama.",
     "uygulama": "BUGÜN ve ETOPLA ile otomatik güncellenen panel oluştur.",
     "videolar": ["Excel haftalık performans paneli", "Excel Weekly Dashboard"],
     "ipucu": "Her hafta güncel veriyi gir, rapor otomatik güncellensin.",
     "sik_hata": "Manuel güncelleme. Formüller otomatik yapar."},

    {"baslik": "Ders 45: Genel Tekrar ve Sonraki Adımlar",
     "ozet": "45 dersin özeti ve ileri seviye konular (Power Query, Makro, Power BI).",
     "uygulama": "Tüm öğrendiklerini kullanarak kapsamlı bir rapor hazırla ve sakla.",
     "videolar": ["Excel ileri seviye Power Query", "Excel Advanced Power BI"],
     "ipucu": "Tebrikler! Artık Excel'in günlük işlerinin %90'ını yapabilirsin.",
     "sik_hata": "Öğrenmeyi bırakmak. Her gün 15 dakika pratik yap."},
]

# ============================================================
# İLERLEME TAKİBİ
# ============================================================
STATE_FILE = "excel_state.json"

def load_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return {"son_ders_index": 0, "toplam_ders": 0}

def save_state(state):
    with open(STATE_FILE, 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

def gunun_dersi():
    state = load_state()
    index = state["son_ders_index"] % len(EXCEL_DERSLERI)
    ders = EXCEL_DERSLERI[index]
    state["son_ders_index"] = index + 1
    state["toplam_ders"] = state.get("toplam_ders", 0) + 1
    save_state(state)
    return ders

# ============================================================
# E-POSTA GÖNDER
# ============================================================
def excel_mail_gonder():
    tarih = datetime.now().strftime('%d.%m.%Y %A')
    ders = gunun_dersi()
    
    # Video linkleri
    video_html = ""
    for video in ders["videolar"]:
        youtube_link = f"https://www.youtube.com/results?search_query={video.replace(' ', '+')}"
        video_html += f'<li style="padding: 8px 0; font-size: 14px;"><a href="{youtube_link}" style="color: #e74c3c; text-decoration: none;">🎥 {video}</a></li>'
    
    html_body = f'''
    <html><body style="font-family: Arial, sans-serif; max-width: 900px; margin: auto; background: #ecf0f1; padding: 20px;">
    <div style="background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); color: white; padding: 25px; border-radius: 12px; text-align: center;">
        <h1 style="margin: 0;">📊 GÜNLÜK EXCEL DERSİ</h1>
        <p style="margin: 8px 0 0 0;">{tarih}</p>
        <p style="margin: 5px 0 0 0; opacity: 0.8; font-size: 12px;">45 Ders · 8 Hafta · Sıfırdan Pivot Tablo'ya</p>
    </div>
    <div style="background: white; padding: 25px; margin-top: 15px; border-radius: 12px;">
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 10px; margin-bottom: 20px;">
            <h2 style="margin: 0; font-size: 22px;">{ders['baslik']}</h2>
        </div>
        <div style="background: #f8f9fa; padding: 15px; border-left: 5px solid #667eea; margin-bottom: 20px;">
            <h3 style="margin: 0 0 10px 0; color: #2c3e50;">📖 Ders Özeti</h3>
            <p style="line-height: 1.7; color: #333;">{ders['ozet']}</p>
        </div>
        <div style="background: #e8f8f5; padding: 15px; border-left: 5px solid #27ae60; margin-bottom: 20px;">
            <h3 style="margin: 0 0 10px 0; color: #2c3e50;">✅ Günün Uygulaması</h3>
            <p style="line-height: 1.7; color: #333; white-space: pre-wrap;">{ders['uygulama']}</p>
        </div>
        <div style="background: #fff9e6; padding: 15px; border-left: 5px solid #f39c12; margin-bottom: 20px;">
            <h3 style="margin: 0 0 10px 0; color: #2c3e50;">🎥 İlgili Videolar</h3>
            <ul style="margin: 0; padding-left: 20px;">{video_html}</ul>
        </div>
        <div style="background: #f5eef8; padding: 15px; border-left: 5px solid #9b59b6; margin-bottom: 20px;">
            <h3 style="margin: 0 0 10px 0; color: #2c3e50;">💡 Günün İpucu</h3>
            <p style="margin: 0; color: #333;">{ders['ipucu']}</p>
        </div>
        <div style="background: #fdedec; padding: 15px; border-left: 5px solid #e74c3c;">
            <h3 style="margin: 0 0 10px 0; color: #2c3e50;">⚠️ Sık Yapılan Hata</h3>
            <p style="margin: 0; color: #333;">{ders['sik_hata']}</p>
        </div>
    </div>
    <div style="background: #34495e; color: white; padding: 15px; margin-top: 15px; border-radius: 12px; font-size: 11px; text-align: center;">
    Bu bülten otomatik oluşturulmuştur. Başarılar!<br>
    📊 Ders {ders['baslik'].split(':')[0].replace('Ders ', '')}/{len(EXCEL_DERSLERI)} · İyi çalışmalar!
    </div>
    </body></html>
    '''
    
    msg = MIMEMultipart('alternative')
    msg['Subject'] = f"📊 Excel Dersi: {ders['baslik']}"
    msg['From'] = MAIL_USER
    msg['To'] = MAIL_TO
    msg.attach(MIMEText(html_body, 'html', 'utf-8'))
    
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(MAIL_USER, MAIL_PASS)
            server.sendmail(MAIL_USER, MAIL_TO, msg.as_string())
        print(f"✅ Excel dersi gönderildi: {ders['baslik']}")
    except Exception as e:
        print(f"❌ Hata: {e}")

if __name__ == "__main__":
    print("📊 Günlük Excel dersi hazırlanıyor...")
    excel_mail_gonder()
