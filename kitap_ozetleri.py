# -*- coding: utf-8 -*-
"""
📚 GÜNLÜK KİTAP ÖZETİ SİSTEMİ v2
+ gTTS ile sesli özet (MP3 ek)
46 kitap havuzu - her gün farklı kitap + AI yorumu + ses dosyası
"""

import os
import smtplib
import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime

try:
    from gtts import gTTS
    GTTS_AKTIF = True
except ImportError:
    GTTS_AKTIF = False
    print("⚠️ gTTS yüklü değil, ses dosyası oluşturulamayacak.")

MAIL_USER = os.environ.get("MAIL_USER")
MAIL_PASS = os.environ.get("MAIL_PASS")
MAIL_TO = os.environ.get("MAIL_TO")
GROQ_API_KEY = os.environ.get("GROQ_API_KEY")

# ============================================================
# KİTAP HAVUZU (46 Kitap)
# ============================================================
KITAPLAR = [
    {
        "baslik": "Limit Sizsiniz",
        "yazar": "Mümin Sekman",
        "kategori": "Kişisel Gelişim",
        "ozet": "İnsanın kendi sınırlarını kendisinin çizdiğini ve bu sınırları kırdığında çok daha fazlasını yapabileceğini anlatır. Başarı, 'doğru düşünme, doğru planlama ve doğru eylem' üçlemi üzerine kuruludur.",
        "ana_fikirler": [
            "Dışımızdaki limitler, içimizdekiler kadar büyür ya da küçülür.",
            "Kafesten çıkınca değil, kafesi içimizden çıkarınca özgürleşiriz.",
            "Büyük başarı kalpten gelir, beyinde büyür, ellerden hayata akar.",
            "Kendi sınırlarını kendin çiziyorsun — onları silmek de senin elinde.",
            "Başarı tesadüf değil; doğru düşünce + doğru plan + doğru eylem."
        ],
        "sorular": [
            "Hayatımda 'yapamam' dediğim ama aslında yapabileceğim ne var?",
            "Kendi kendime koyduğum en büyük sınır nedir?",
            "Bu sınırı kaldırsam ilk ne yapardım?"
        ]
    },
    {
        "baslik": "Her Şey Seninle Başlar",
        "yazar": "Mümin Sekman",
        "kategori": "Kişisel Gelişim",
        "ozet": "Kişisel kurtuluş savaşını başlatmak isteyenler için bir rehber. İnsanın kendi hayatının sorumluluğunu alması ve içsel dönüşümü başlatması gerektiğini anlatır.",
        "ana_fikirler": [
            "Değişim dışarıdan değil, içeriden başlar.",
            "Kurban rolünden çık, sorumluluk al.",
            "Bahane üretmek, başarının en büyük düşmanıdır.",
            "Küçük bir içsel karar, büyük bir dışsal değişimi tetikler.",
            "Kendine yatırım yapmak, en kârlı yatırımdır."
        ],
        "sorular": [
            "Hayatımda değiştirmek istediğim şey için ilk adımı attım mı?",
            "Hâlâ hangi bahaneye sığınıyorum?",
            "Bugün kendim için ne yapabilirim?"
        ]
    },
    {
        "baslik": "Ya Bir Yol Bul, Ya Bir Yol Aç, Ya da Yoldan Çekil",
        "yazar": "Mümin Sekman",
        "kategori": "Kişisel Gelişim",
        "ozet": "Mümin Sekman'ın 21 yaşında yazdığı ilk kitabı. Hayatta karşılaşılan engeller karşısında üç seçenek sunar: yol bulmak, yol açmak veya yoldan çekilmek.",
        "ana_fikirler": [
            "Engel, çözümün başlangıcıdır.",
            "Yol bulamıyorsan, yol açmayı dene.",
            "Yoldan çekilmek de bir karardır — ama en son seçenek olmalı.",
            "Pasiflik, en büyük zaman kaybıdır.",
            "Her engel, yeni bir kapı açar."
        ],
        "sorular": [
            "Şu an hayatımda hangi engele takıldım?",
            "Bu engel için 'yol bulma' seçeneğim ne?",
            "'Yol açma' seçeneğim ne olabilir?"
        ]
    },
    {
        "baslik": "Türk Usulü Başarı",
        "yazar": "Mümin Sekman",
        "kategori": "Kişisel Gelişim",
        "ozet": "Türkiye'de 1980'lerden sonra oluşan yeni başarı kültürünü analiz eder. Türk insanının başarıya bakışını ve başarısızlık nedenlerini ele alır.",
        "ana_fikirler": [
            "Türkiye'de başarı, Batı'dakinden farklı dinamiklere sahiptir.",
            "Kültürel faktörler başarıyı etkiler.",
            "Türk insanının güçlü ve zayıf yönlerini bilmek önemlidir.",
            "Başarı taklit edilemez, uyarlanabilir.",
            "Yerel bağlamı anlamadan başarı stratejisi kurulamaz."
        ],
        "sorular": [
            "Türkiye'de başarılı olmak için hangi yerel dinamikleri bilmeliyim?",
            "Kültürel faktörler benim başarımı nasıl etkiliyor?",
            "Başarı stratejimi yerel gerçeklere göre uyarladım mı?"
        ]
    },
    {
        "baslik": "Kesintisiz Öğrenme",
        "yazar": "Mümin Sekman",
        "kategori": "Öğrenme",
        "ozet": "Daha kalıcı, hızlı ve kolay öğrenmenin yollarını anlatır. Öğrenmenin sadece okulda değil, hayat boyu devam eden bir süreç olduğunu vurgular.",
        "ana_fikirler": [
            "Öğrenme, hayat boyu devam eden bir süreçtir.",
            "Öğrenmeyi öğrenmek, en değerli beceridir.",
            "Hızlı öğrenme teknikleri öğrenilebilir.",
            "Tekrar, öğrenmenin anahtarıdır.",
            "Öğrendiğini uygulamak, kalıcılığı artırır."
        ],
        "sorular": [
            "En son ne öğrendim ve ne kadarını hatırlıyorum?",
            "Öğrenme yöntemim verimli mi?",
            "Öğrendiğimi günlük hayatımda nasıl uygularım?"
        ]
    },
    {
        "baslik": "Kişisel Ataleti Yenmek",
        "yazar": "Mümin Sekman",
        "kategori": "Kişisel Gelişim",
        "ozet": "Erteleme, tembellik ve harekete geçememe sorununu ele alır. Ataletin nedenlerini ve çözümlerini pratik yöntemlerle anlatır.",
        "ana_fikirler": [
            "Atalet, harekete geçememe halidir.",
            "Erteleme, mükemmeliyetçilikten beslenir.",
            "Küçük başlangıçlar, büyük eylemleri tetikler.",
            "Disiplin, motivasyondan daha güvenilirdir.",
            "Harekete geçmek için ilham beklemek, ataletin tuzağıdır."
        ],
        "sorular": [
            "Şu an ertelediğim en önemli şey ne?",
            "Erteleme nedenim gerçekten geçerli mi?",
            "Bu işin 'en küçük ilk adımı' ne olabilir?"
        ]
    },
    {
        "baslik": "Başarı Üniversitesi",
        "yazar": "Mümin Sekman",
        "kategori": "Kişisel Gelişim",
        "ozet": "Başarının okulda öğretilmeyen yönlerini ele alır. Hayatta başarılı olmak için gereken becerileri, zihniyet ve stratejileri anlatır.",
        "ana_fikirler": [
            "Okul, hayatın tamamını öğretmez.",
            "Başarı için farklı beceriler gerekir.",
            "Kendini geliştirmek, mezun olduktan sonra da devam eder.",
            "Başarı, öğrenilebilir bir beceridir.",
            "Hayat, en iyi öğretmendir."
        ],
        "sorular": [
            "Okulda öğrenmediğim ama hayatta gerekli olan ne var?",
            "Kendimi hangi alanda geliştirmeliyim?",
            "Hayatın bana öğrettiği en önemli ders nedir?"
        ]
    },
    {
        "baslik": "Başarı Düşünürü",
        "yazar": "Mümin Sekman",
        "kategori": "Kişisel Gelişim",
        "ozet": "Mümin Sekman'ın bireysel gelişim ve başarı ideolojisini zaman, zekâ ve strateji boyutlarıyla ele aldığı eseri.",
        "ana_fikirler": [
            "Başarı, bütüncül bir yaklaşım gerektirir.",
            "Zaman yönetimi, başarının temelidir.",
            "Zekâ tek başına yeterli değildir.",
            "Strateji, başarıyı sistematik hale getirir.",
            "Başarı, sürekli öğrenme ve uyum sağlama ile mümkündür."
        ],
        "sorular": [
            "Başarı modelim bütüncül mü?",
            "Zamanımı nasıl yönetiyorum?",
            "Stratejim net mi?"
        ]
    },
    {
        "baslik": "Etkili İnsanların 7 Alışkanlığı",
        "yazar": "Stephen R. Covey",
        "kategori": "Liderlik",
        "ozet": "Kişisel ve profesyonel etkinliği artırmak için 7 temel alışkanlık sunar: proaktif olmak, sonunu düşünmek, önemli işlere öncelik vermek, kazan/kazan düşünmek, önce anlamak, sinerji yaratmak ve baltayı bilemek.",
        "ana_fikirler": [
            "Proaktif ol: Kendi hayatının sorumluluğunu al.",
            "Sonunu düşün: Hedefini netleştir.",
            "Önemli işlere öncelik ver: Acil olan değil, önemli olan.",
            "Kazan/kazan: Herkesin kazandığı çözümler ara.",
            "Önce anla: Empati kur, sonra konuş.",
            "Sinerji: Farklılıklardan güç al.",
            "Baltayı bile: Kendini sürekli yenile."
        ],
        "sorular": [
            "Hangi alışkanlığım en zayıf?",
            "Bugün hangi alışkanlığı uygulayabilirim?",
            "Bu alışkanlıklar hayatıma nasıl entegre olur?"
        ]
    },
    {
        "baslik": "İnsanın Anlam Arayışı",
        "yazar": "Viktor E. Frankl",
        "kategori": "Psikoloji",
        "ozet": "Nazi toplama kamplarında hayatta kalan psikiyatr Viktor Frankl'ın deneyimleri üzerinden logoterapiyi anlatır. İnsanın temel motivasyonunun haz değil, anlam olduğunu savunur.",
        "ana_fikirler": [
            "İnsanın temel motivasyonu anlam arayışıdır.",
            "En zor koşullarda bile anlam bulunabilir.",
            "Acı, anlamlandırıldığında katlanılabilir hale gelir.",
            "Özgürlük, koşullara verdiğimiz tepkiyi seçme özgürlüğüdür.",
            "Yaşamın anlamı kişiye ve duruma göre değişir."
        ],
        "sorular": [
            "Hayatımın anlamı nedir?",
            "Zorluklar karşısında nasıl bir tutum sergiliyorum?",
            "Bugün neyi anlamlandırabilirim?"
        ]
    },
    {
        "baslik": "Dost Kazanma ve İnsanları Etkileme Sanatı",
        "yazar": "Dale Carnegie",
        "kategori": "İletişim",
        "ozet": "İnsan ilişkilerinin temel prensiplerini pratik örneklerle anlatır. İnsanları eleştirmek yerine anlamak, samimi ilgi göstermek ve dinlemenin güçlü ilişkilerin anahtarı olduğunu vurgular.",
        "ana_fikirler": [
            "Eleştirme, suçlama ve şikâyet etme.",
            "Samimi ilgi göster.",
            "İsmini hatırla ve kullan.",
            "İyi bir dinleyici ol.",
            "Karşındakinin bakış açısını anla."
        ],
        "sorular": [
            "İnsanlarla ilişkilerimde hangi prensibi uygulamıyorum?",
            "Bugün kime samimi ilgi gösterebilirim?",
            "Dinleme becerim ne durumda?"
        ]
    },
    {
        "baslik": "Düşün ve Zengin Ol",
        "yazar": "Napoleon Hill",
        "kategori": "Başarı",
        "ozet": "1937'de yayımlanan bu klasik, zenginlik ve başarıya ulaşmak için 13 ilke sunar: arzu, inanç, uzmanlaşmış bilgi, hayal gücü, planlama, kararlılık ve azim.",
        "ana_fikirler": [
            "Arzu, başarının başlangıç noktasıdır.",
            "İnanç, düşünceyi güçlendirir.",
            "Uzmanlaşmış bilgi, güçtür.",
            "Hayal gücü, fırsatları görünür kılar.",
            "Kararlılık ve azim, engelleri aşar."
        ],
        "sorular": [
            "Gerçekten ne istiyorum?",
            "Bu isteğe ne kadar inanıyorum?",
            "Bunu başarmak için planım ne?"
        ]
    },
    {
        "baslik": "Simyacı",
        "yazar": "Paulo Coelho",
        "kategori": "Roman",
        "ozet": "Endülüslü çoban Santiago'nun Mısır piramitlerinde hazine arayışını anlatan felsefi roman. 'Kişisel Menkıbe'ni takip et' mesajı verir.",
        "ana_fikirler": [
            "Kişisel menkıbeni takip et.",
            "Hayallerinin peşinden gitmekten korkma.",
            "Evren, hayalini gerçekleştirmek isteyene yardım eder.",
            "Hazine, gittiğin yolda değil, döndüğün yerde olabilir.",
            "Yolculuk, varıştan daha önemlidir."
        ],
        "sorular": [
            "Benim kişisel menkıbem ne?",
            "Hayalimin peşinden gidiyor muyum?",
            "Yolculuğumda ne öğrendim?"
        ]
    },
    {
        "baslik": "Atomik Alışkanlıklar",
        "yazar": "James Clear",
        "kategori": "Alışkanlık",
        "ozet": "Küçük, sürekli değişikliklerin büyük sonuçlar doğurduğunu anlatır. %1 iyileşmenin zamanla birikerek büyük farklar yarattığını vurgular.",
        "ana_fikirler": [
            "Her gün %1 iyileşme, yılda 37 kat gelişme demektir.",
            "Kimliğine uygun alışkanlıklar seç.",
            "Sistemler, hedeflerden daha önemlidir.",
            "Alışkanlık döngüsü: işaret → istek → tepki → ödül.",
            "Kötü alışkanlıkları görünmez, zor ve tatminsiz kıl."
        ],
        "sorular": [
            "Hangi küçük alışkanlığı eklemek istiyorum?",
            "Hangi kötü alışkanlığı çıkarmalıyım?",
            "Ortamım alışkanlıklarımı destekliyor mu?"
        ]
    },
    {
        "baslik": "Ustalık Gerektiren Kafaya Takmama Sanatı",
        "yazar": "Mark Manson",
        "kategori": "Kişisel Gelişim",
        "ozet": "Her şeye iyi tarafından bakma baskısını eleştirir. Hayatın bazen berbat olduğunu ve yapılacak en sağlıklı şeyin bunu kabul etmek olduğunu savunur.",
        "ana_fikirler": [
            "Her şeye iyi tarafından bakmak zorunda değilsin.",
            "Hayat bazen berbattır; bunu kabul et.",
            "Önemli olan neyi önemsediğindir.",
            "Daha iyi problemler seç.",
            "Mutluluk, problemlerin yokluğu değil, iyi problemlerin varlığıdır."
        ],
        "sorular": [
            "Ne için endişeleniyorum ve bu gerçekten önemli mi?",
            "Hangi problemlerle uğraşmak istiyorum?",
            "Neyi daha az önemsemeliyim?"
        ]
    },
    {
        "baslik": "Bir Ömür Nasıl Yaşanır?",
        "yazar": "İlber Ortaylı",
        "kategori": "Yaşam",
        "ozet": "Tarihçi İlber Ortaylı'nın hayat tecrübelerini ve tavsiyelerini içerir. Hayatı dört döneme ayırır ve her dönem için öneriler sunar.",
        "ana_fikirler": [
            "Cesur ol, konfor alanının dışına çık.",
            "Farklı dünyalarla tanışmak için çaba göster.",
            "Miskinlik, insanın bittiği andır.",
            "Kitap oku, dil öğren, seyahat et.",
            "Her dönemin kendine özgü fırsatları vardır."
        ],
        "sorular": [
            "Hangi hayat dönemindeyim ve bu dönemin fırsatları ne?",
            "Konfor alanımın dışına çıkıyor muyum?",
            "Bugün öğrenmek için ne yaptım?"
        ]
    },
    {
        "baslik": "Ikigai: Japonların Uzun ve Mutlu Yaşam Sırrı",
        "yazar": "Hector Garcia & Francesc Miralles",
        "kategori": "Yaşam Felsefesi",
        "ozet": "Japonların 'ikigai' kavramını anlatır: Her sabah yataktan kalkmak için bir sebep. İkigai, sevdiğin, iyi olduğun, dünyanın ihtiyaç duyduğu ve para kazanabileceğin şeylerin kesişimidir.",
        "ana_fikirler": [
            "Herkesin bir ikigaisi vardır.",
            "İkigai, hayata anlam katar.",
            "Uzun ve mutlu yaşamın sırrı ikigaidir.",
            "Küçük şeylerden mutlu ol.",
            "Şimdiki anda yaşa."
        ],
        "sorular": [
            "Benim ikigaim ne?",
            "Sevdiğim, iyi olduğum ve dünyanın ihtiyaç duyduğu şey ne?",
            "Her sabah kalkma sebebim ne?"
        ]
    },
    {
        "baslik": "Şu Hortumlu Dünyada Fil Yalnız Bir Hayvandır",
        "yazar": "Ahmet Şerif İzgören",
        "kategori": "Motivasyon",
        "ozet": "İletişim, başarı ve hayat üzerine pratik motivasyon kitabı. Mizahi bir dille hayat dersleri verir.",
        "ana_fikirler": [
            "Hayat, iletişimle anlam kazanır.",
            "Kendin ol, başkaları gibi olmaya çalışma.",
            "Mizah, hayatın zorluklarıyla başa çıkmanın yoludur.",
            "Başarı, küçük adımlarla gelir.",
            "Farkında olmak, değişimin başlangıcıdır."
        ],
        "sorular": [
            "Kendim olmaktan ne zaman vazgeçtim?",
            "Hayatıma mizah katıyor muyum?",
            "Bugün fark ettiğim bir şey ne oldu?"
        ]
    },
    {
        "baslik": "Savaşçı",
        "yazar": "Doğan Cüceloğlu",
        "kategori": "Psikoloji",
        "ozet": "Bireyin kendi içsel yolculuğunu ve yaşam mücadelesini ele alır. 'Savaşçı', hayatın zorluklarına göğüs geren, kendi değerlerinin farkında olan kişidir.",
        "ana_fikirler": [
            "Savaşçı, içsel bütünlüğünü koruyan kişidir.",
            "Hayat, bir mücadeledir ve savaşçı bu mücadeleyi kabul eder.",
            "Kendi değerlerinin farkında olmak önemlidir.",
            "İçsel yolculuk, dışsal başarıdan önce gelir.",
            "Savaşçı, korkularıyla yüzleşir."
        ],
        "sorular": [
            "Ben bir savaşçı mıyım?",
            "Hangi değerlerim için mücadele ediyorum?",
            "Korkularımla yüzleşiyor muyum?"
        ]
    },
    {
        "baslik": "Var Mısın?",
        "yazar": "Doğan Cüceloğlu",
        "kategori": "Kişisel Gelişim",
        "ozet": "Bireyin kendine sorması gereken temel soruları ele alır: Hayatta neyi başarmak istiyorsun? Kendine güveniyor musun? Değişime hazır mısın?",
        "ana_fikirler": [
            "Kendine 'var mısın?' sorusunu sor.",
            "Değişim için cesaret gerekir.",
            "Kendine güven, başarının temelidir.",
            "Hayatta anlam bulmak için çaba gerekir.",
            "İçsel motivasyon, dışsal motivasyondan güçlüdür."
        ],
        "sorular": [
            "Hayatta neyi başarmak istiyorum?",
            "Değişime var mıyım?",
            "Kendime ne kadar güveniyorum?"
        ]
    },
    {
        "baslik": "Akış",
        "yazar": "Mihaly Csikszentmihalyi",
        "kategori": "Psikoloji",
        "ozet": "Mutluluk ve yaratıcılığın zirvesi olan 'akış' halini anlatır. Akış, bir işe tamamen dalmış olma halidir. Yüksek beceri ve yüksek zorluk dengesinde ortaya çıkar.",
        "ana_fikirler": [
            "Akış, tamamen dalmış olma halidir.",
            "Mutluluk, akış halinde ortaya çıkar.",
            "Beceri ve zorluk dengesi akışı sağlar.",
            "Akış, yaratıcılığı ve verimliliği artırır.",
            "Herkes akış halini deneyimleyebilir."
        ],
        "sorular": [
            "En son ne zaman akış halini yaşadım?",
            "Hangi iş beni akışa sokar?",
            "Akış halini nasıl daha sık yaşayabilirim?"
        ]
    },
    {
        "baslik": "Mindset",
        "yazar": "Carol Dweck",
        "kategori": "Psikoloji",
        "ozet": "Sabit ve gelişim zihniyeti arasındaki farkı anlatır. Sabit zihniyet, yeteneklerin doğuştan geldiğine inanır. Gelişim zihniyeti ise yeteneklerin çabayla geliştirilebileceğini savunur.",
        "ana_fikirler": [
            "İki tür zihniyet vardır: sabit ve gelişim.",
            "Gelişim zihniyeti, başarıyı artırır.",
            "Yetenekler çabayla geliştirilebilir.",
            "Başarısızlık, öğrenme fırsatıdır.",
            "Zihniyet değiştirilebilir."
        ],
        "sorular": [
            "Hangi zihniyete sahibim?",
            "Başarısızlıkları nasıl yorumluyorum?",
            "Zihniyetimi nasıl geliştirebilirim?"
        ]
    },
    {
        "baslik": "Sapiens",
        "yazar": "Yuval Noah Harari",
        "kategori": "Tarih",
        "ozet": "İnsanlık tarihini üç devrimle anlatır: Bilişsel Devrim, Tarım Devrimi ve Bilimsel Devrim. İnsanın nasıl dünyaya hâkim olduğunu ve geleceğini sorgular.",
        "ana_fikirler": [
            "İnsanlık üç büyük devrim yaşadı.",
            "Bilişsel Devrim, dil ve hayal gücünü getirdi.",
            "Tarım Devrimi, yerleşik hayatı başlattı.",
            "Bilimsel Devrim, günümüzü şekillendirdi.",
            "İnsan, dünyanın en güçlü ama en tehlikeli türüdür."
        ],
        "sorular": [
            "İnsanlık nereye gidiyor?",
            "Teknoloji hayatımızı nasıl dönüştürüyor?",
            "Gelecekte bizi ne bekliyor?"
        ]
    },
    {
        "baslik": "Outliers",
        "yazar": "Malcolm Gladwell",
        "kategori": "Başarı",
        "ozet": "Başarının sadece yetenek ve çabayla değil, fırsatlar, kültürel faktörler ve zamanlama ile de ilgili olduğunu savunur.",
        "ana_fikirler": [
            "Başarı, sadece yetenek değildir.",
            "Fırsatlar ve zamanlama önemlidir.",
            "10.000 saat kuralı: Ustalık için çok çalışmak gerekir.",
            "Kültürel faktörler başarıyı etkiler.",
            "Şans, başarının gizli faktörüdür."
        ],
        "sorular": [
            "Hangi fırsatları kaçırdım?",
            "Hangi alanda 10.000 saat harcadım?",
            "Başarımda şansın rolü ne?"
        ]
    },
    {
        "baslik": "Cesur Yeni Dünya",
        "yazar": "Aldous Huxley",
        "kategori": "Roman",
        "ozet": "Distopik bir gelecekte, teknoloji ve biyoloji ile kontrol edilen bir toplumu anlatır. Bireysel özgürlüğün nasıl yok edildiğini ve konforun nasıl bir tuzak haline geldiğini gösterir.",
        "ana_fikirler": [
            "Konfor, özgürlüğü yok edebilir.",
            "Teknoloji, kontrol aracı olabilir.",
            "Bireysellik, toplum tarafından bastırılabilir.",
            "Mutluluk, dayatıldığında anlamını yitirir.",
            "Distopya, sessizce gelebilir."
        ],
        "sorular": [
            "Özgürlüğümü ne kadar koruyorum?",
            "Konfor beni esir alıyor mu?",
            "Teknolojiyi nasıl kullanıyorum?"
        ]
    },
    {
        "baslik": "İrade Terbiyesi",
        "yazar": "Jules Payot (çev. Ethem Bakar)",
        "kategori": "İrade",
        "ozet": "İradenin eğitilebilir bir güç olduğunu savunur. Disiplin, dikkat ve azim gibi zihinsel yeteneklerin nasıl geliştirileceğini bilimsel temellerle anlatır.",
        "ana_fikirler": [
            "İrade doğuştan değil, eğitimle geliştirilebilir.",
            "Dikkat, iradenin temelidir.",
            "Küçük fedakârlıklar, büyük irade gücü yaratır.",
            "Zihinsel disiplin, düzenli egzersizle güçlenir.",
            "Duyguları yönetmek, irade eğitiminin özüdür."
        ],
        "sorular": [
            "İrademi en çok hangi durumlarda kaybediyorum?",
            "Dikkatimi geliştirmek için ne yapabilirim?",
            "Bugün hangi küçük fedakârlığı yapabilirim?"
        ]
    },
    {
        "baslik": "İyi Hissetmek",
        "yazar": "David Burns",
        "kategori": "Psikoloji",
        "ozet": "Bilişsel davranışçı terapi temelli bir rehberdir. Depresyon, kaygı ve olumsuz düşünce kalıplarını tanıyıp değiştirmenin yollarını öğretir.",
        "ana_fikirler": [
            "Duygular, düşüncelerden doğar.",
            "Olumsuz düşünce kalıpları tanınabilir ve değiştirilebilir.",
            "Bilişsel çarpıtmalar yaygındır.",
            "Kendine karşı şefkat, iyileşmenin başlangıcıdır.",
            "Küçük günlük pratikler, büyük değişim yaratır."
        ],
        "sorular": [
            "Bugün hangi olumsuz düşünceyi fark ettim?",
            "Bu düşünce gerçekten doğru mu?",
            "Daha gerçekçi bir bakış açısı nasıl olurdu?"
        ]
    },
    {
        "baslik": "Beden Dili",
        "yazar": "Joe Navarro",
        "kategori": "İletişim",
        "ozet": "Eski FBI ajanı Joe Navarro, sözsüz iletişimin sırlarını anlatır. İnsanların beden hareketlerinin arkasındaki gerçek anlamları okumayı öğretir.",
        "ana_fikirler": [
            "Beden dili, sözcüklerden önce gelir.",
            "Bacaklar ve ayaklar, gerçek duyguları en çok ele veren uzuvlardır.",
            "Rahatlık ve rahatsızlık, iki temel sinyaldir.",
            "Beden dili tek başına değil, bağlamla okunmalıdır.",
            "Kendi beden dilini kontrol etmek, güven verir."
        ],
        "sorular": [
            "Karşımdakinin beden dili ne anlatıyor?",
            "Benim bedenim hangi mesajı veriyor?",
            "Bu durumda nasıl daha güven verici durabilirim?"
        ]
    },
    {
        "baslik": "İnsan İnsana",
        "yazar": "Doğan Cüceloğlu",
        "kategori": "İletişim",
        "ozet": "Türkiye'de iletişim ve insan psikolojisinin dinamiklerini en yalın anlatan kaynaklardan biridir.",
        "ana_fikirler": [
            "İletişim, hayatın temelidir.",
            "Kendini tanımak, başkalarını anlamanın başlangıcıdır.",
            "Dinlemek, konuşmaktan daha önemlidir.",
            "Kültürel farklar, iletişimi etkiler.",
            "Empati, sağlıklı ilişkilerin temelidir."
        ],
        "sorular": [
            "Bugün kimseyi gerçekten dinledim mi?",
            "İletişimimde hangi hataları yapıyorum?",
            "Empati kurmakta zorlandığım durumlar neler?"
        ]
    },
    {
        "baslik": "Ted Gibi Konuş",
        "yazar": "Carmine Gallo",
        "kategori": "İletişim",
        "ozet": "Dünyanın en etkili konuşmacılarının sırlarını analiz eder. Etkileyici sunum yapmanın 9 temel prensibini örneklerle anlatır.",
        "ana_fikirler": [
            "İyi konuşmalar, hikayelerle başlar.",
            "Basit ve net ol, karmaşıklıktan kaç.",
            "Benzetme ve metafor kullan.",
            "Tutku, konuşmanın itici gücüdür.",
            "Beden dili ve ses tonu, içerikten önce gelir."
        ],
        "sorular": [
            "Sunum yaparken hangi prensipleri kullanıyorum?",
            "Hikaye anlatma becerim nasıl?",
            "Bu konuşmayı nasıl daha etkili yapabilirim?"
        ]
    },
    {
        "baslik": "Neden ile Başla (Start with Why)",
        "yazar": "Simon Sinek",
        "kategori": "Liderlik",
        "ozet": "Gerçek liderlerin ekiplerine 'nasıl' ve 'ne'den önce 'neden' sorusuyla ilham verdiğini açıklar. Altın Çember modeli: Neden → Nasıl → Ne.",
        "ana_fikirler": [
            "İnsanlar ne yaptığın için değil, neden yaptığın için harekete geçer.",
            "Altın Çember: Neden → Nasıl → Ne.",
            "İlham vermek, manipüle etmekten güçlüdür.",
            "Net bir 'neden', sadakat ve bağlılık yaratır.",
            "Liderler, bir amaç etrafında toplar."
        ],
        "sorular": [
            "Benim 'neden'im ne?",
            "Bu işi neden yapıyorum?",
            "Ekibimi bir amaç etrafında nasıl toplarım?"
        ]
    },
    {
        "baslik": "Bir CEO'nun Günlüğü",
        "yazar": "Steven Bartlett",
        "kategori": "Liderlik",
        "ozet": "Modern iş dünyasının liderlik dinamiklerini ve güncel iş kurallarını inceler. Genç bir girişimcinin gözünden başarı ve başarısızlık üzerine samimi bir günlüktür.",
        "ana_fikirler": [
            "Başarısızlık, başarının bir parçasıdır.",
            "Genç yaşta liderlik, önyargıları kırar.",
            "Modern iş dünyası, sürekli uyum gerektirir.",
            "Deneyim, en değerli öğretmendir.",
            "Şeffaflık, liderlikte güç yaratır."
        ],
        "sorular": [
            "Başarısızlıklarımdan ne öğrendim?",
            "Liderlik tarzım nasıl evriliyor?",
            "Bugün hangi kararı cesaretle alabilirim?"
        ]
    },
    {
        "baslik": "İktidar - Güç Sahibi Olmanın 48 Yasası",
        "yazar": "Robert Greene",
        "kategori": "Strateji",
        "ozet": "Tarih boyunca güç sahibi olmuş kişilerin stratejilerini 48 yasada toplar. İktidar, strateji ve iknanın sırlarını anlatır.",
        "ana_fikirler": [
            "Güç, elde edilir ve korunur.",
            "Görünmez kalmak, bazen en güçlü stratejidir.",
            "Bağımlılık yaratmak, kontrol sağlar.",
            "Duyguları yönetmek, gücün temelidir.",
            "İtibar, gerçek güç kadar önemlidir."
        ],
        "sorular": [
            "Gücü nasıl kullanıyorum?",
            "Hangi durumlarda gücümü kaybediyorum?",
            "Bu yasalar etik sınırlar içinde nasıl uygulanır?"
        ]
    },
    {
        "baslik": "Değer Katan Ekip Çalışması",
        "yazar": "Prof. Dr. Acar Baltaş",
        "kategori": "Yönetim",
        "ozet": "Kurumsal ekiplerin uyumunu ve verimliliğini artırmaya yönelik temel bir rehberdir. Ekip oluşturma ve güven inşası üzerine pratik bilgiler sunar.",
        "ana_fikirler": [
            "Ekip, bireylerin toplamından fazlasıdır.",
            "Güven, ekibin temelidir.",
            "Ortak hedef, ekibi bir arada tutar.",
            "Roller net olmalıdır.",
            "İletişim, ekip verimliliğinin anahtarıdır."
        ],
        "sorular": [
            "Ekibimde güven var mı?",
            "Roller net mi?",
            "Ortak hedefimiz herkesçe biliniyor mu?"
        ]
    },
    {
        "baslik": "Pürüzsüz (Grit)",
        "yazar": "Angela Duckworth",
        "kategori": "Kariyer",
        "ozet": "Kariyerde başarının zekadan ziyade tutku ve azim (grit) ile şekillendiğini kanıtlar. Azmin öğrenilebilir bir beceri olduğunu savunur.",
        "ana_fikirler": [
            "Başarı, tutku + azim = grit.",
            "Zekâ, tek başına yeterli değildir.",
            "Azim, öğrenilebilir bir beceridir.",
            "Uzun vadeli hedeflere bağlılık önemlidir.",
            "Zorluklar, büyümenin fırsatlarıdır."
        ],
        "sorular": [
            "Neye tutkuyla bağlıyım?",
            "Zorluklar karşısında pes mi ediyorum, direniyor muyum?",
            "Azmimi nasıl geliştirebilirim?"
        ]
    },
    {
        "baslik": "Derin Çalışma (Deep Work)",
        "yazar": "Cal Newport",
        "kategori": "Verimlilik",
        "ozet": "Dikkat dağıtıcı unsurlarla dolu modern iş dünyasında, odaklanarak nasıl yüksek katma değerli işler üretebileceğinizi öğretir.",
        "ana_fikirler": [
            "Derin çalışma, 21. yüzyılın süper gücüdür.",
            "Sığ işler bizi tüketir.",
            "Odaklanma, öğrenilebilir bir beceridir.",
            "Dijital detoks, üretkenliği artırır.",
            "Derin çalışma, uzun vadeli başarı getirir."
        ],
        "sorular": [
            "Günde kaç saat derin çalışıyorum?",
            "Dikkatimi ne dağıtıyor?",
            "Derin çalışma için nasıl zaman ayırabilirim?"
        ]
    },
    {
        "baslik": "Ustalık (Mastery)",
        "yazar": "Robert Greene",
        "kategori": "Kariyer",
        "ozet": "Seçtiğiniz alanda sıradanlıktan sıyrılıp nasıl bir 'usta' haline geleceğinizin tarihsel ve stratejik haritasını sunar.",
        "ana_fikirler": [
            "Ustalık, hayat boyu süren bir yolculuktur.",
            "Kendi yolunu bulmak, ustalığın başlangıcıdır.",
            "Çıraklık, ustalığın zorunlu aşamasıdır.",
            "Sosyal zekâ, ustalığı hızlandırır.",
            "Yaratıcılık, ustalığın zirvesidir."
        ],
        "sorular": [
            "Hangi alanda usta olmak istiyorum?",
            "Bu alanda ne kadar yol kat ettim?",
            "Ustalık yolculuğumda hangi aşamadayım?"
        ]
    },
    {
        "baslik": "Kendi Kutup Yıldızını Bul",
        "yazar": "Nüvit Osmay",
        "kategori": "Vizyon",
        "ozet": "Kariyer yolculuğunda kişisel vizyon oluşturma ve potansiyeli doğru yönlendirme üzerine rehberlik eder.",
        "ana_fikirler": [
            "Herkesin bir kutup yıldızı olmalıdır.",
            "Vizyon, yolculuğun başlangıcıdır.",
            "Kendini tanımak, doğru yönü bulmanın yoludur.",
            "Hedefe odaklanmak, ilerlemeyi hızlandırır.",
            "Küçük adımlar, büyük yolculukların başlangıcıdır."
        ],
        "sorular": [
            "Benim kutup yıldızım ne?",
            "Hayatımda nereye gidiyorum?",
            "Bu yön doğru mu?"
        ]
    },
    {
        "baslik": "Ekibin Beş Temel Aksaklığı",
        "yazar": "Patrick Lencioni",
        "kategori": "Ekip",
        "ozet": "Ekiplerin başarısız olma nedenlerini sürükleyici bir iş hikayesi üzerinden anlatır. Beş aksaklık: güven eksikliği, çatışma korkusu, bağlılık eksikliği, sorumluluk kaçınma, sonuçlara odaklanmama.",
        "ana_fikirler": [
            "Ekiplerin başarısızlığı beş temel sebepten kaynaklanır.",
            "Güven, tüm aksaklıkların çözümüdür.",
            "Yapıcı çatışma, sağlıklı ekibin işaretidir.",
            "Bağlılık, kararların arkasında durmaktır.",
            "Sorumluluk, ekibin ortak yüküdür."
        ],
        "sorular": [
            "Ekibimde güven var mı?",
            "Çatışmadan mı kaçınıyoruz?",
            "Kararların arkasında duruyor muyuz?"
        ]
    },
    {
        "baslik": "Radikal Candor",
        "yazar": "Kim Scott",
        "kategori": "Liderlik",
        "ozet": "Bir lider olarak çalışanları kırmadan ama tamamen açık ve dürüst geri bildirimlerle nasıl yönetebileceğinizi açıklar.",
        "ana_fikirler": [
            "Radikal candor = özen + dürüstlük.",
            "Kişisel özen göstermeden dürüstlük, zulüm olur.",
            "Dürüstlük olmadan özen, sahtekârlık olur.",
            "Geri bildirim, sürekli ve doğrudan olmalıdır.",
            "Liderlik, başkalarının başarısını kolaylaştırmaktır."
        ],
        "sorular": [
            "Ekibime yeterince dürüst müyüm?",
            "Onlara yeterince özen gösteriyor muyum?",
            "Geri bildirimlerimi nasıl daha etkili yapabilirim?"
        ]
    },
    {
        "baslik": "Liderler En Son Yer",
        "yazar": "Simon Sinek",
        "kategori": "Liderlik",
        "ozet": "Başarılı ekiplerin arkasında, liderlerin çalışanları için nasıl bir güven ve koruma çemberi oluşturduğunu inceler.",
        "ana_fikirler": [
            "Liderler, ekibini korumak için en sona geçer.",
            "Güven, güvenlik hissinden doğar.",
            "Kriz anlarında liderler öne çıkar.",
            "Ekip, liderin karakterini yansıtır.",
            "Fedakârlık, gerçek liderliğin işaretidir."
        ],
        "sorular": [
            "Ekibim için ne kadar fedakârlık yapıyorum?",
            "Onlara güvenli bir ortam sağlıyor muyum?",
            "Kriz anlarında nasıl davranıyorum?"
        ]
    },
    {
        "baslik": "Kültür Kodu",
        "yazar": "Daniel Coyle",
        "kategori": "Kültür",
        "ozet": "Dünyanın en başarılı organizasyonlarının arkasındaki ekip kültürünü ve ortak aidiyet duygusunun nasıl inşa edildiğini çözümler.",
        "ana_fikirler": [
            "Kültür, başarının en güçlü itici gücüdür.",
            "Güvenlik hissi, öğrenmeyi kolaylaştırır.",
            "Ortak kırılganlık, bağları güçlendirir.",
            "Amaç, kültürün en derin katmanıdır.",
            "Küçük davranışlar, büyük kültür yaratır."
        ],
        "sorular": [
            "Ekibimin kültürü nasıl?",
            "İnsanlar kendini güvende hissediyor mu?",
            "Kültürümüzü nasıl güçlendirebilirim?"
        ]
    },
    {
        "baslik": "Takım Oyunu",
        "yazar": "Acar Baltaş",
        "kategori": "Ekip",
        "ozet": "Türkiye'deki kurumsal kültür üzerinden ortak bir hedef etrafında kilitlenmeyi ve yüksek performanslı ekipler yaratmayı anlatır.",
        "ana_fikirler": [
            "Ekip başarısı, bireysel başarıdan önce gelir.",
            "Ortak hedef, ekibi birleştirir.",
            "Güven, takım oyununun temelidir.",
            "Türk kültürü, ekip dinamiklerini etkiler.",
            "Yüksek performans, sistemli çalışmayla gelir."
        ],
        "sorular": [
            "Ekibim ortak hedefe kilitlenmiş mi?",
            "Güven inşa etmek için ne yapıyorum?",
            "Takım oyununu nasıl güçlendirebilirim?"
        ]
    },
    {
        "baslik": "Drive (Nasıl Motive Oluruz?)",
        "yazar": "Daniel H. Pink",
        "kategori": "Motivasyon",
        "ozet": "Klasik ödül-ceza yöntemlerinin artık işe yaramadığını; modern motivasyonun özerklik, ustalık ve amaç olduğunu bilimsel verilerle kanıtlar.",
        "ana_fikirler": [
            "Ödül-ceza sistemi, yaratıcılığı öldürür.",
            "Gerçek motivasyon üç şeyden gelir: özerklik, ustalık, amaç.",
            "Özerklik, insanın kendi kararlarını vermesidir.",
            "Ustalık, sürekli gelişme arzusudur.",
            "Amaç, bir şeyin büyük resme hizmet etmesidir."
        ],
        "sorular": [
            "Beni gerçekten ne motive ediyor?",
            "İşimde özerkliğim var mı?",
            "Yaptığım işin amacı ne?"
        ]
    },
    {
        "baslik": "Kelimelerle İkna Et",
        "yazar": "Phil M. Jones",
        "kategori": "İkna",
        "ozet": "Günlük ve profesyonel konuşmalarda, karşınızdaki kişiyi yönlendirmek ve direnç göstermeden ikna etmek için kullanabileceğiniz sihirli soru ve kelime kalıplarını sunar.",
        "ana_fikirler": [
            "Doğru kelimeler, direnci kırar.",
            "Soru sormak, iknanın en güçlü aracıdır.",
            "Küçük kelime değişiklikleri, büyük etki yaratır.",
            "İkna, manipülasyon değil, anlayıştır.",
            "Sessizlik, bazen en güçlü cevaptır."
        ],
        "sorular": [
            "İkna ederken hangi hataları yapıyorum?",
            "Karşımdakini anlamak için yeterince soru soruyor muyum?",
            "Kelimelerimi nasıl daha etkili kullanabilirim?"
        ]
    },
    {
        "baslik": "Dürtme (Nudge)",
        "yazar": "Richard H. Thaler & Cass R. Sunstein",
        "kategori": "Davranış",
        "ozet": "Davranışsal iktisat ve psikoteknik ilkelerini kullanarak, insanların kararlarını özgürlüklerini kısıtlamadan küçük mimari dokunuşlarla nasıl doğru yöne çekebileceğinizi açıklar.",
        "ana_fikirler": [
            "İnsanlar her zaman rasyonel karar vermez.",
            "Küçük dürtmeler, büyük davranış değişiklikleri yaratır.",
            "Seçim mimarisi, kararları etkiler.",
            "Özgürlük korunarak yönlendirme yapılabilir.",
            "Varsayılan seçenekler, güçlü bir dürtmedir."
        ],
        "sorular": [
            "Günlük kararlarımı ne etkiliyor?",
            "Beni doğru yöne iten dürtmeler neler?",
            "Başkalarına nasıl etik dürtmeler yapabilirim?"
        ]
    }
]

# ============================================================
# İLERLEME TAKİBİ
# ============================================================
STATE_FILE = "kitap_state.json"

def load_kitap_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return {
        "son_kitap_index": 0,
        "toplam_okunan": 0,
        "baslangic_tarihi": str(datetime.now().date())
    }

def save_kitap_state(state):
    with open(STATE_FILE, 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

def gunun_kitabi():
    state = load_kitap_state()
    index = state["son_kitap_index"] % len(KITAPLAR)
    kitap = KITAPLAR[index]
    state["son_kitap_index"] = index + 1
    state["toplam_okunan"] = state.get("toplam_okunan", 0) + 1
    save_kitap_state(state)
    return kitap

# ============================================================
# SESLİ ÖZET OLUŞTURMA (gTTS)
# ============================================================
def sesli_ozet_olustur(kitap):
    """Kitap özetini MP3 dosyasına çevirir."""
    if not GTTS_AKTIF:
        return None
    try:
        # Seslendirilecek metni hazırla
        metin = f"{kitap['baslik']}. Yazar: {kitap['yazar']}. "
        metin += f"Kitap özeti: {kitap['ozet']} "
        metin += "Ana fikirler: "
        for i, fikir in enumerate(kitap['ana_fikirler'], 1):
            metin += f"{i}. {fikir} "
        metin += "Düşünce soruları: "
        for i, soru in enumerate(kitap['sorular'], 1):
            metin += f"{i}. {soru} "
        
        # MP3 dosyası oluştur
        dosya_adi = f"günün_kitabı.mp3"
        tts = gTTS(text=metin, lang='tr', slow=False)
        tts.save(dosya_adi)
        print(f"✅ Ses dosyası oluşturuldu: {dosya_adi}")
        return dosya_adi
    except Exception as e:
        print(f"❌ Ses dosyası hatası: {e}")
        return None

# ============================================================
# AI YORUMU
# ============================================================
def ai_yorum(kitap):
    if not GROQ_API_KEY:
        return "AI yorumu için API anahtarı gerekli."
    try:
        from groq import Groq
        client = Groq(api_key=GROQ_API_KEY)
        prompt = f"""
        Sen bir kişisel gelişim koçusun. Aşağıdaki kitap özetini okuyup, 
        okuyucuya kişisel bir yorum ve öneri yaz.
        
        Kitap: {kitap['baslik']}
        Yazar: {kitap['yazar']}
        Özet: {kitap['ozet']}
        Ana Fikirler: {', '.join(kitap['ana_fikirler'])}
        
        Şu formatta yaz (en fazla 200 kelime):
        
        💬 KİŞİSEL YORUM:
        [Bu kitabın sana ne anlattığını, hangi fikrinin seni en çok etkilediğini yaz]
        
        🎯 SANA ÖNERİM:
        [Bu kitaptaki hangi fikri hayatına nasıl uygulayabileceğini yaz]
        
        Türkçe, samimi ve motive edici bir dille yaz.
        """
        chat = client.chat.completions.create(
            messages=[{"role": "user", "content": prompt}],
            model="openai/gpt-oss-20b",
            max_tokens=600,
        )
        return chat.choices[0].message.content
    except Exception as e:
        print(f"AI hatası: {e}")
        return "🤖 AI yorumu şu an oluşturulamadı."

# ============================================================
# E-POSTA GÖNDER
# ============================================================
def kitap_mail_gonder():
    tarih = datetime.now().strftime('%d.%m.%Y %A')
    kitap = gunun_kitabi()
    yorum = ai_yorum(kitap)
    
    # Sesli özet oluştur
    ses_dosyasi = sesli_ozet_olustur(kitap)
    
    fikirler_html = ""
    for i, fikir in enumerate(kitap["ana_fikirler"], 1):
        fikirler_html += f'<li style="padding: 6px 0; border-bottom: 1px solid #eee;"><strong>{i}.</strong> {fikir}</li>'
    
    sorular_html = ""
    for i, soru in enumerate(kitap["sorular"], 1):
        sorular_html += f'<li style="padding: 8px 0;"><strong>{i}.</strong> {soru}</li>'
    
    ses_bilgi = ""
    if ses_dosyasi:
        ses_bilgi = '''
        <div style="background: #e7f3ff; padding: 15px; border-left: 5px solid #3498db; margin-bottom: 20px; border-radius: 8px;">
            <h3 style="margin: 0 0 10px 0; color: #2c3e50;">🎧 SESLİ DİNLE</h3>
            <p style="margin: 0; color: #333;">Bu dersin sesli özeti e-posta ektedir. 
            <strong>günün_kitabı.mp3</strong> dosyasını indirip istediğiniz zaman dinleyebilirsiniz.</p>
        </div>
        '''
    
    html_body = f'''
    <html><body style="font-family: Arial, sans-serif; max-width: 950px; margin: auto; background: #ecf0f1; padding: 20px;">
    <div style="background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); color: white; padding: 25px; border-radius: 12px; text-align: center;">
        <h1 style="margin: 0;">📚 GÜNLÜK KİTAP ÖZETİ</h1>
        <p style="margin: 8px 0 0 0;">{tarih}</p>
        <p style="margin: 5px 0 0 0; opacity: 0.8; font-size: 12px;">🎧 Sesli dinleme aktif · {len(KITAPLAR)} kitap havuzu</p>
    </div>
    <div style="background: white; padding: 25px; margin-top: 15px; border-radius: 12px;">
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 10px; margin-bottom: 20px;">
            <h2 style="margin: 0; font-size: 24px;">{kitap["baslik"]}</h2>
            <p style="margin: 8px 0 0 0; font-size: 14px; opacity: 0.9;">✍️ {kitap["yazar"]} · 📂 {kitap["kategori"]}</p>
        </div>
        {ses_bilgi}
        <div style="background: #f8f9fa; padding: 15px; border-left: 5px solid #667eea; margin-bottom: 20px;">
            <h3 style="margin: 0 0 10px 0; color: #2c3e50;">📖 Kitap Özeti</h3>
            <p style="line-height: 1.7; color: #333;">{kitap["ozet"]}</p>
        </div>
        <div style="background: #e8f8f5; padding: 15px; border-left: 5px solid #27ae60; margin-bottom: 20px;">
            <h3 style="margin: 0 0 10px 0; color: #2c3e50;">🎯 Ana Fikirler</h3>
            <ul style="margin: 0; padding-left: 20px;">{fikirler_html}</ul>
        </div>
        <div style="background: #fef5e7; padding: 15px; border-left: 5px solid #f39c12; margin-bottom: 20px;">
            <h3 style="margin: 0 0 10px 0; color: #2c3e50;">❓ Düşünce Soruları</h3>
            <ul style="margin: 0; padding-left: 20px;">{sorular_html}</ul>
        </div>
        <div style="background: #f5eef8; padding: 15px; border-left: 5px solid #9b59b6; margin-bottom: 20px;">
            <h3 style="margin: 0 0 10px 0; color: #2c3e50;">💬 Kişisel Yorum ve Öneri</h3>
            <div style="white-space: pre-wrap; line-height: 1.7;">{yorum}</div>
        </div>
    </div>
    <div style="background: #34495e; color: white; padding: 15px; margin-top: 15px; border-radius: 12px; font-size: 11px; text-align: center;">
    Bu bülten otomatik oluşturulmuştur.<br>
    📚 {len(KITAPLAR)} kitap · 🎧 Sesli dinleme aktif
    </div>
    </body></html>
    '''
    
    msg = MIMEMultipart('mixed')
    msg['Subject'] = f"📚 Günün Kitabı: {kitap['baslik']} - {kitap['yazar']}"
    msg['From'] = MAIL_USER
    msg['To'] = MAIL_TO
    
    # HTML gövde
    msg_alt = MIMEMultipart('alternative')
    msg_alt.attach(MIMEText(html_body, 'html', 'utf-8'))
    msg.attach(msg_alt)
    
    # MP3 eki
    if ses_dosyasi and os.path.exists(ses_dosyasi):
        try:
            with open(ses_dosyasi, 'rb') as f:
                ses_ek = MIMEApplication(f.read(), _subtype='mp3')
                ses_ek.add_header('Content-Disposition', 'attachment', 
                                  filename=f"gunun_kitabi_{datetime.now().strftime('%Y%m%d')}.mp3")
                msg.attach(ses_ek)
            print("✅ MP3 eki hazırlandı.")
        except Exception as e:
            print(f"MP3 eki hatası: {e}")
    
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(MAIL_USER, MAIL_PASS)
            server.sendmail(MAIL_USER, MAIL_TO, msg.as_string())
        print(f"✅ Kitap e-postası gönderildi: {kitap['baslik']}")
    except Exception as e:
        print(f"❌ Hata: {e}")

if __name__ == "__main__":
    print("📚 Günlük kitap özeti hazırlanıyor...")
    kitap_mail_gonder()
