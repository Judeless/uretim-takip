# -*- coding: utf-8 -*-
"""
tel_proses.py — TK1 tel üretiminin proses adımı mantığı (2026-08-04).

Neden ayrı modül: bu kuralı app.py, mail_raporu.py ve oee.py'nin üçü de kullanıyor.
mail_raporu/oee, app.py'yi import EDEMEZ (Flask uygulamasını ayağa kaldırır, döngüsel
bağımlılık olur) — ortak mantık burada durur, üçü de buradan alır.

TEMEL KURAL
-----------
Bir tel referansı sırayla birden çok HATTAN geçer (halat kesme → yarı/tam otomat →
kapama → son montaj) ve her adımı FARKLI operatör kendi vardiyasında kaydeder.
Adımlar aynı kodla toplanırsa aynı 100 adetlik iş 4 kez sayılır.

ÇÖZÜM: her adım kendi EKİYLE kaydedilir — '93.TK.464 KESIM', '93.TK.464 KAPAMA'…
Böylece adımlar ayrı kodlarda toplanır; ne çoklu sayım olur ne de kimin ne iş
yaptığı kaybolur.

NEDEN "SON ADIM" TANIMI YOK (2026-08-04 revizyonu):
Kapaması ya da son montajı dışarıda yapılacak ürünler SABİT DEĞİL — üretim
yoğunluğuna ve tedarikçi durumuna göre değişiyor. Bu yüzden bir referansın son
adımı önceden tanımlanamaz; 'son adım base kodu alır' kuralı terk edildi.
Ayrıca TK1'de üretim teyidini KALİTE departmanı verdiği için base kodun ERP
article'ıyla birebir kalma zorunluluğu da yok.
"""

import re

# Üretim sırasında proses adımları.
# 'Otomatik Hazırlık' 2026-08-19'da eklendi (kullanıcı): otomatik hazırlık yapma
# makinesi. KESİMDEN SONRA gelir (kullanıcı) → kesim ile otomat adımlarının
# arasında. Bu tuple RAPOR KOLON SIRASINI belirler; mantık TEL_ADIM_SIRA
# üyeliğine bakar, sayısal değere değil.
# 'Soyma' 2026-08-20'de eklendi (kullanıcı): "kesim yapılan ürünlerde soyma
# yapılır, HER ÜRÜNE YAPILMAZ. Kesim bölümünün altında görünsün fakat kodun
# yanında SOYMA olarak gelsin raporda." Bu yüzden AYRI adımdır (kendi kod eki +
# kendi rapor sütunu) ama sırada kesimin hemen ardından gelir; operatör
# listesinde de kesim makinelerinin yanında görünür (bkz. /api/kayit_hatlari).
# KESIM ile aynı eki ALAMAZDI: soyma her ürüne yapılmadığı için aynı kodda
# toplansalardı, soyulan ürünün kesimi ile soyması tek koda yığılır ve aynı 100
# parça 200 üretim gibi görünürdü — adım ayrımının varlık sebebi tam olarak bu.
# 'Halat Kesme' → 'Halat/Spiral Kesme' (kullanıcı 2026-08-26): adım altında
# hem halat kesme hem spiral kesme makineleri var, adın ikisini de kapsaması
# istendi. KOD EKİ 'KESIM' DEĞİŞMEDİ — değişseydi bu ekle yazılmış tüm geçmiş
# kayıtlar rapor kırılımından düşerdi (bkz. TEL_ADIM_EKI_ESKI notu).
# 'Ön Hazırlık' 2026-08-27'de eklendi (kullanıcı): TK1 telde 2 tezgâh.
# KENDİ ADIMI — 'Otomatik Hazırlık'ın makinesi DEĞİL (kullanıcı kararı): kod eki
# ve rapor sütunu ayrıdır, aynı ürünün ön hazırlığı ile otomatik hazırlığı ayrı
# ayrı sayılır. Aynı eke bağlansalardı iki farklı operasyon tek kodda toplanır
# ve 100 parçalık iş 200 üretim gibi görünürdü (Soyma'nın kesimden ayrılma
# gerekçesinin aynısı). Sırada SOYMA'dan sonra, Otomatik Hazırlık'tan ÖNCE.
TEL_ADIMLARI = ('Halat/Spiral Kesme', 'Soyma', 'Ön Hazırlık', 'Otomatik Hazırlık',
                'Yarı Otomatik', 'Tam Otomatik', 'Kapama', 'Son Montaj')

# Sıra numarası: yarı ve tam otomatik AYNI adım sayılır — biri diğerinin
# alternatifi, ikisi arka arkaya uygulanmaz.
# NOT: sayısal değer HİÇBİR YERDE kullanılmıyor (mantık ÜYELİĞE bakar); burada
# prosesi okunur kılmak için duruyor. Bu yüzden araya adım eklemek güvenlidir.
TEL_ADIM_SIRA = {'Halat/Spiral Kesme': 1, 'Soyma': 2,
                 'Ön Hazırlık': 3, 'Otomatik Hazırlık': 4,
                 'Yarı Otomatik': 5, 'Tam Otomatik': 5,
                 'Kapama': 6, 'Son Montaj': 7}

# HAT ADI → ADIM İSTİSNASI (kullanıcı 2026-08-20). Normalde hat adı adımın
# kendisidir ya da sonuna numara alır ('Kapama 30' → 'Kapama'). 'Manuel Kesim'
# bu kalıba UYMAZ: operatör listede makinenin gerçek adını görmeli ama sistem
# bunu Halat Kesme saymalı (kod eki KESIM, raporda kesim sütunu) — kesim üretimi
# iki ayrı sütuna bölünmesin.
TEL_HAT_ADIM_ISTISNA = {
    'Manuel Kesim': 'Halat/Spiral Kesme',
    # MAKINE ADI ADIM ADINDAN BAGIMSIZ (2026-08-26): kullanici makine adlarini
    # sahadaki etiketlere gore degistirdi ama ADIM adlari (dolayisiyla kod ekleri
    # ve rapor sutunlari) AYNI kaldi. Her yeni makine adi buraya bir satir ekler.
    #   'Otomatik Pres Makinesi' -> adim 'Tam Otomatik'  (ek: OTOMATİK MAKİNE HATTI)
    #   'Hidrolik Pres N'        -> adim 'Kapama'        (ek: KAPAMA)
    #   'Halat Kesme N'          -> adim 'Halat/Spiral Kesme' (ek: KESIM)
    'Otomatik Pres Makinesi': 'Tam Otomatik',
    # ESKI AD — YALNIZ OKUMA: bu adla yazilmis vardiyalarin (robot_no) adimi
    # cozulmeye devam etsin, yoksa gecmis kayitlar rapor kiriliminden duser.
    'Otomatik Kesim Makinesi': 'Tam Otomatik',
    'Hidrolik Pres': 'Kapama',
    'Kapama': 'Kapama',              # eski hat adi ('Kapama 5') — yalniz okuma
    'Halat Kesme': 'Halat/Spiral Kesme',
    # SPIRAL KESME (kullanici 2026-08-21): TK1'de halat kesmenin YANINDA 3 otomatik
    # + 1 manuel spiral kesme makinesi var. Kullanici karari: AYRI ADIM DEGIL,
    # kesim adimina DAHIL -> kod eki KESIM, raporda tek kesim sutunu.
    # Operator listede makinenin gercek adini gorur ('Manuel Kesim' kalibinin ayni).
    # Numarali adlar da cozulur: tel_hat_adimi son numarayi atip burada tekrar arar.
    'Otomatik Spiral Kesme': 'Halat/Spiral Kesme',
    'Manuel Spiral Kesme': 'Halat/Spiral Kesme',
}

# FİZİKSEL HATLAR (kullanıcı 2026-08-04): kesim 3 · yarı otomatik 2 · tam otomatik 1
# · kapama 12 · son montaj 4 = 22 hat. Tek makineli adımlarda numara YOK.
# KAPAMA presleri 2026-08-18'de SAHA KODLARINA çevrildi — aşağıdaki nota bakın.
# YARI OTOMATİK 1 → 2 (kullanıcı 2026-08-18): sahada iki makine olduğu ortaya çıktı.
# İkisinin de sayacı ESP32 DEĞİL, SVP test cihazı (app.TEST_CIHAZ_ESLEME:
# 'Yarı Otomatik 1'→41, 'Yarı Otomatik 2'→43) — bu hatlara saha modülü takılmayacak.
# KAPAMA 4 → 12 (2026-08-07): kapama hattında 12 pres var; sayaç modülleri
# takılırken gerçek makine sayısı ortaya çıktı (eski 4 rakamı eksikti). Presler
# üçerli ortak panolarda; 4 ESP32 modülü 3'er presi okur (bkz. pilot/firmware/
# _templates/tel_kapama.ino.tpl). Hat adları firmware'deki robot_no ile BİREBİR.
# ── KAPAMA PRESLERİ: SAHA KODU (kullanıcı 2026-08-18) ───────────────────────
# Modüller takılınca sistemdeki sıra numarası ile presin ÜRETİM SAHASINDAKİ kodu
# tutmadığı görüldü: sistemde 'Kapama 1' olan modülün bağlı olduğu presin saha
# kodu 5. Operatör sahada gördüğü kodu seçmeli → hat adları SAHA KODUNA çevrildi.
#
# FIRMWARE YENİDEN YÜKLENMEDİ: modüller hâlâ robot_no='Kapama 1..6' gönderiyor.
# Hat adı → cihaz eşlemesi app.HAT_SAYAC_CIHAZI'nda (LF-LFP→YF1 ile aynı kalıp).
#   MODÜL 1-2 (2026-08-18)          MODÜL 3-4 (2026-08-19)
#   saha 5  ← cihaz Kapama 1        saha 21 ← cihaz Kapama 7
#   saha 12 ← cihaz Kapama 2        saha 19 ← cihaz Kapama 8
#   saha 4  ← cihaz Kapama 3        saha 20 ← cihaz Kapama 9
#   saha 30 ← cihaz Kapama 4        saha 27 ← cihaz Kapama 10
#   saha 28 ← cihaz Kapama 5        saha 25 ← cihaz Kapama 11
#   saha 29 ← cihaz Kapama 6        saha 26 ← cihaz Kapama 12
#
# 12 presin TAMAMI adlandırıldı — geçici 'Kapama 9xx' adları kalktı, gizli hat
# kalmadı. Kural aynı kalıyor: kod gelince YALNIZ AD değişir, SIRA ASLA.
#
# SIRA DEĞİŞMEZ — YALNIZ AD DEĞİŞTİ: pozisyon = uretim_kayitlari.istasyon.
# Liste kısaltılsaydı Son Montaj hatları 19..22'den 13..16'ya kayar ve yazılmış
# kayıtlar başka hattı göstermeye başlardı. Yeni hat HEP SONA eklenir.
TEL_HATLARI = (
    ['Halat Kesme %d' % i for i in range(1, 4)]
    + ['Yarı Otomatik %d' % i for i in range(1, 3)]
    # AD DEGISTI, POZISYON AYNI (kullanici 2026-08-21): operator makinenin gercek
    # adini gormeli. ADIM 'Tam Otomatik' KALIR (TEL_HAT_ADIM_ISTISNA) — adim adi
    # degisseydi kod eki de degisirdi ('... TAM OTOMAT') ve o ekle yazilmis TUM
    # gecmis kayitlar rapor kirilimindan duserdi.
    # AD DEGISTI, POZISYON AYNI (kullanici 2026-08-26): 'Otomatik Kesim
    # Makinesi' -> 'Otomatik Pres Makinesi'. Adim 'Tam Otomatik' KALIR.
    + ['Otomatik Pres Makinesi']
    # AD DEGISTI, SIRA AYNI (kullanici 2026-08-26): "kapama makinelerinin
    # adlarini makine etiket numaralari kalacak sekilde hidrolik pres olarak
    # degistirelim." Saha kodu (5, 12, 4 ...) KORUNDU; yalniz basindaki kelime
    # degisti. Pozisyon = uretim_kayitlari.istasyon -> asla kaymaz.
    # Cihaz adlari (firmware 'Kapama 1..12') DEGISMEDI, yeniden flash GEREKMEZ;
    # esleme app._KAPAMA_SAHA_CIHAZ'da. Yan fayda: hat ve cihaz adlari artik
    # birbirine benzemiyor, eski 'Kapama 12 hem hat hem cihaz' karisikligi bitti.
    + ['Hidrolik Pres 5', 'Hidrolik Pres 12', 'Hidrolik Pres 4',
       'Hidrolik Pres 30', 'Hidrolik Pres 28', 'Hidrolik Pres 29']
    + ['Hidrolik Pres 21', 'Hidrolik Pres 19', 'Hidrolik Pres 20',
       'Hidrolik Pres 27', 'Hidrolik Pres 25', 'Hidrolik Pres 26']
    # SON MONTAJ 2 HATTA DÜŞTÜ (kullanıcı 2026-08-20): sahada bu iş 4 ve 5 numaralı
    # buton modülleriyle yapılıyor (eski YF modülü son montaja taşındı) → hat adı
    # BUTON NUMARASINI taşır; operatör masasının üstünde yazan numarayı seçer
    # (kapama preslerindeki saha kodu mantığının aynısı).
    # 903/904 = kullanılmayan iki slot: SİLİNMEDİ çünkü liste kısalsaydı sonraki
    # hat ('Otomatik Hazırlık') 23'ten 21'e kayar ve yazılmış kayıtlar başka hattı
    # gösterirdi. Gizli tutulur (TEL_GIZLI_HATLAR), gerekirse adı değiştirilip açılır.
    + ['Son Montaj 4', 'Son Montaj 5', 'Son Montaj 903', 'Son Montaj 904']
    # YENİ HAT HEP SONA (2026-08-19: otomatik hazırlık makinesi) — araya girmek
    # yazılmış kayıtların istasyon numaralarını başka hatta kaydırır.
    # Operatöre Türkçe görünür; sayaç cihazının robot_no'su ASCII 'Otomatik Hazirlik'
    # (firmware/klasör adı kuralı) — eşleme app.HAT_SAYAC_CIHAZI'nda.
    + ['Otomatik Hazırlık']
    # 2026-08-20 (kullanıcı): TK1'de bir MANUEL KESİM makinesi ve bir SOYMA
    # makinesi var. Listenin SONUNA eklendiler — araya girmek yazılmış kayıtların
    # istasyon numaralarını başka hatta kaydırırdı. Operatör listesinde yine de
    # kendi adımlarının yanında görünürler: /api/kayit_hatlari listeyi PROSES
    # SIRASINA göre sıralayıp gönderir, gönderilen değer POZİSYON numarasıdır.
    #   'Manuel Kesim' → adım 'Halat Kesme' (TEL_HAT_ADIM_ISTISNA), kod eki KESIM
    #   'Soyma'        → kendi adımı,                                kod eki SOYMA
    # İkisinin de sayaç modülü YOK → adet elle girilir.
    + ['Manuel Kesim', 'Soyma']
    # 2026-08-21 (kullanici): "halat kesme haricinde 3 adet otomatik spiral kesme
    # ve 1 adet manuel spiral kesme makinesi". Yine SONA eklendi (pozisyon = istasyon
    # no, asla kaymaz); operator listesinde kesim makinelerinin yaninda gorunurler
    # cunku /api/kayit_hatlari listeyi PROSES SIRASINA gore siralar ve bu dordunun
    # adimi 'Halat Kesme'dir (TEL_HAT_ADIM_ISTISNA).
    # Kod eki KESIM -> raporda ayri sutun ACILMAZ, kesim uretimine eklenir.
    # Ikisinin de sayac modulu YOK -> adet elle girilir.
    + ['Otomatik Spiral Kesme %d' % i for i in range(1, 4)]
    + ['Manuel Spiral Kesme']
    # 2026-08-27 (kullanici): TK1 telde 2 adet ON HAZIRLIK tezgahi. Yine SONA
    # eklendi (pozisyon = istasyon no, asla kaymaz). Operator listesinde kendi
    # adimlarinin yaninda gorunurler; /api/kayit_hatlari listeyi PROSES SIRASINA
    # gore siralar ve bu ikisinin adimi 'On Hazirlik'tir.
    # Adim adi hat adinin kokune esit oldugu icin ISTISNA GEREKMEZ:
    # tel_hat_adimi('On Hazirlik 1') -> son numarayi atar -> 'On Hazirlik'.
    # Sayac modulu YOK -> adet elle girilir.
    + ['Ön Hazırlık %d' % i for i in range(1, 3)]
)

# Kullanılmayan / kodu bekleyen hatlar — operatör listesinde gösterilmez ama
# POZİSYONLARI listede DURUR (istasyon numaraları kaymasın). Adı verilince
# yalnız ADI değiştirilip buradan çıkarılır.
# 2026-08-20: son montaj 4 slottan 2'ye düştü, kalan iki slot burada.
TEL_GIZLI_HATLAR = frozenset({'Son Montaj 903', 'Son Montaj 904'})

_SON_NUMARA = re.compile(r'\s*\d+$')


def tel_hat_adimi(robot_no):
    """Hat adı → PROSES ADIMI. 'Kapama 3' → 'Kapama', 'Yarı Otomatik' → kendisi.

    Son adım karşılaştırması hat adına DEĞİL adım tipine bakar; aksi hâlde
    'Son Montaj 2'de çalışan operatörün kaydı 'Son Montaj' ile eşleşmez ve o
    referansın üretimi hiç sayılmazdı. Tanınmayan hat → None."""
    ad = str(robot_no or '').strip()
    if not ad:
        return None
    if ad in TEL_HAT_ADIM_ISTISNA:      # 'Manuel Kesim' → 'Halat Kesme'
        return TEL_HAT_ADIM_ISTISNA[ad]
    if ad in TEL_ADIM_SIRA:
        return ad
    kok = _SON_NUMARA.sub('', ad).strip()
    # NUMARA ATILDIKTAN SONRA DA ISTISNAYA BAK (2026-08-21): 'Otomatik Spiral
    # Kesme 1' gibi HEM numarali HEM istisnali hat adlari icin sart. Eskiden
    # istisna yalnizca TAM ADLA aranirdi; numarali bir istisna hattinin adimi
    # sessizce None dondu ve o hattin uretimi rapor kiriliminden duserdi.
    if kok in TEL_HAT_ADIM_ISTISNA:
        return TEL_HAT_ADIM_ISTISNA[kok]
    return kok if kok in TEL_ADIM_SIRA else None


# ── PROSES ADIMI EKİ (kullanıcı 2026-08-04) ─────────────────────────────────
# "Kesim hattındayken seçilen referansın sonuna KESIM yazılır; böylece son
#  operasyon olmadığı anlaşılır ve diğer hatlardaki işlemlerle karışmaz."
#
# HER ADIM EK ALIR — 'son adım ek almaz' kuralı YOK (2026-08-04 revizyonu).
# Sebep (kullanıcı): "Kapaması ya da son montajı dışarıda olacak ürünler belli
# değil; üretim yoğunluğuna ve tedarikçi durumuna göre değişiyor." Yani bir
# referansın SON adımı sabit değildir, referans bazında tanımlanamaz.
# İkinci sebep: TK1'de üretim teyidini KALİTE departmanı veriyor → base kodun
# ERP'deki article ile birebir kalması gerekmiyor, ek serbestçe kullanılabilir.
#
#   93.TK.464 @ Halat Kesme 1 → '93.TK.464 KESIM'
#   93.TK.464 @ Kapama 2      → '93.TK.464 KAPAMA'
#   93.TK.464 @ Son Montaj 3  → '93.TK.464 SON MONTAJ'
#
# Sonuç: her adım kendi kodunda toplanır, hatlar birbirine karışmaz, aynı iş
# birden çok kez "üretim" sayılmaz ve hangi ürünün nerede bittiği raporda
# kendiliğinden görünür — sabit bir proses tanımına ihtiyaç kalmaz.
TEL_ADIM_EKI = {
    # DİKKAT — 'ÖN HAZIRLIK' ile 'HAZIRLIK' iç içe geçiyor: '… ÖN HAZIRLIK'
    # kodunda desen 'HAZIRLIK' ekini de görebilir. _ek_deseni_kur() ekleri
    # UZUNDAN KISAYA sıralayıp alternasyona koyduğu için uzun olan kazanır.
    # Bu bilinçli bir bağımlılık: sıralama kaldırılırsa ön hazırlık kayıtları
    # sessizce 'Otomatik Hazırlık' adımına düşer (bkz. testler).
    'Ön Hazırlık': 'ÖN HAZIRLIK',
    'Otomatik Hazırlık': 'HAZIRLIK',
    'Halat/Spiral Kesme': 'KESIM',
    'Soyma':         'SOYMA',
    # Otomat adımlarının ekleri makine adını taşır (kullanıcı 2026-08-21):
    # hat adı 'Otomatik Kesim Makinesi' olunca kodun yanında 'TAM OTOMAT'
    # yazması anlamsız kalıyordu.
    'Yarı Otomatik': 'YARI OTO. MAKİNE HATTI',
    'Tam Otomatik':  'OTOMATİK MAKİNE HATTI',
    'Kapama':        'KAPAMA',
    'Son Montaj':    'SON MONTAJ',
}

# ── ESKİ EKLER — YALNIZ OKUMA (2026-08-21) ──────────────────────────────────
# Kod eki değiştiğinde ESKİ ekle yazılmış kayıtlar ortada kalır: onların adımı
# çözülemezse rapor kırılımından sessizce düşerler (aynı ürünün üretimi
# "bilinmeyen adım"a gider). Bu yüzden eski ekler tanınmaya DEVAM eder; yalnız
# artık YENİ kayıt üretmezler.
# YENİ EK EKLERKEN: eskisini buraya taşı, silme.
TEL_ADIM_EKI_ESKI = {
    'TAM OTOMAT':  'Tam Otomatik',
    'YARI OTOMAT': 'Yarı Otomatik',
}

# Ek zaten yazılmış mı? (operatör elle yazdıysa iki kez eklenmesin)
# DESEN ELLE YAZILMAZ — eklerden ÜRETİLİR. Elle yazıldığında yeni bir ek
# eklenip desene yazılmayı unutmak sessiz hataya yol açıyordu: o ek ayıklanmaz
# ve kod her kayıtta bir ek daha alırdı ('… SOYMA SOYMA').
#
# SIRA KRİTİK: uzun ekler ÖNCE denenmeli. Alternasyon ilk eşleşeni alır;
# 'YARI OTO. MAKİNE HATTI' kısa bir alternatifle çakışırsa yanlış adım çözülür.
#
# TÜRKÇE BÜYÜK-İ: yeni ekler 'İ' içeriyor ve IGNORECASE Python'da dotted-İ ile
# 'i' eşleşmesini güvenilir yapmaz. Bu yüzden desende hem yazıldığı hâli hem
# casefold hâli bulunur; eşleşme her iki yazımda da çalışır.
def _ek_anahtar(ek):
    """Ek karşılaştırma anahtarı — TÜRKÇE-DUYARSIZ.

    str.upper() Türkçede güvenilmez: 'otomatik'.upper() → 'OTOMATIK' (NOKTASIZ I),
    oysa yazılan ek 'OTOMATİK' (noktalı İ). İkisi eşleşmez ve kodun adımı
    çözülemez. Bu yüzden Türkçe harfler ASCII karşılığına indirgenir; hem harita
    anahtarları hem aranan değer AYNI fonksiyondan geçer."""
    t = str(ek or '').strip()
    for a, b in (('İ', 'I'), ('ı', 'I'), ('Ş', 'S'), ('ş', 'S'), ('Ğ', 'G'),
                 ('ğ', 'G'), ('Ü', 'U'), ('ü', 'U'), ('Ö', 'O'), ('ö', 'O'),
                 ('Ç', 'C'), ('ç', 'C')):
        t = t.replace(a, b)
    return ' '.join(t.upper().split())


def _ek_deseni_kur():
    ekler = list(TEL_ADIM_EKI.values()) + list(TEL_ADIM_EKI_ESKI.keys())
    varyant = set()
    for e in ekler:
        varyant.add(e)
        varyant.add(e.casefold())
        # ASCII-KATLANMIŞ HÂL DE DESENDE (2026-08-27): operatör kodu elle
        # yazdığında Türkçe harfleri kullanmayabiliyor ('ÖN HAZIRLIK' yerine
        # 'ON HAZIRLIK'). Katlanmış hâl desende yoksa uzun ek eşleşmez ve kod
        # İÇİNDEKİ KISA EKE düşer: '… ON HAZIRLIK' sessizce 'HAZIRLIK' sayılıp
        # Otomatik Hazırlık adımına yazılırdı — yanlış adımda üretim demek.
        # (Ters harita _EK_ADIM zaten _ek_anahtar ile normalize arıyor.)
        _a = _ek_anahtar(e)
        varyant.add(_a)
        varyant.add(_a.casefold())
    # Boşluklar esnek (çift boşlukla yazılmış kodlar da tanınsın)
    parcalar = [re.escape(v).replace(r'\ ', r'\s+') for v in
                sorted(varyant, key=len, reverse=True)]
    return re.compile(r'\s+(' + '|'.join(parcalar) + r')\s*$', re.IGNORECASE)


_EK_DESEN = _ek_deseni_kur()


def tel_ek_ayikla(referans_kodu):
    """Koddaki ara-operasyon ek(ler)ini atar: '93.TK.464 KESIM' → '93.TK.464'.

    ÜST ÜSTE EKLERİ DE TEMİZLER (2026-08-20): desen sonda tek eşleşme attığı için
    '… YARI OTOMAT YARI OTOMAT' gibi çiftlenmiş bir kod tek geçişte düzelmiyordu
    ve kendini ASLA toparlamıyordu. Normal akışta böyle bir kod oluşmaz (kayıtta
    ek eklenmeden önce ayıklanır) ama elle yazımla girebilir; girdiğinde de
    sessizce yaşamasın."""
    kod = str(referans_kodu or '').strip()
    for _ in range(8):          # > adım sayısı; sonsuz döngü olamaz
        yeni = _EK_DESEN.sub('', kod).strip()
        if yeni == kod:
            return kod
        kod = yeni
    return kod


# Ek → adım ters haritası. Rapor katmanı adımı KODDAN çözer (hat/istasyon
# gerekmez): mail_raporu.py app.py'yi import EDEMEZ ama kodu görebilir.
# Ek → adım ters haritası. ESKİ ekler de burada: o ekle yazılmış kayıtların
# adımı çözülmeye devam etsin (yoksa raporda 'bilinmeyen adım'a düşerler).
_EK_ADIM = {_ek_anahtar(ek): adim for adim, ek in TEL_ADIM_EKI.items()}
_EK_ADIM.update({_ek_anahtar(ek): adim for ek, adim in TEL_ADIM_EKI_ESKI.items()})


def tel_koddan_adim(referans_kodu):
    """'93.TK.464 KAPAMA' → 'Kapama'. Ek yoksa None (tel dışı ya da eksiz kayıt).

    Adım bazlı rapor kırılımının temeli: aynı ürünün farklı adımları AYRI kodlarda
    durduğu için, base kod + adım ikilisi 'aynı 100 parça 4 adımdan geçti'yi
    '400 üretim' gibi göstermeden çözer."""
    m = _EK_DESEN.search(str(referans_kodu or ''))
    if not m:
        return None
    # Ek yazımı esnek olabilir ('SON  MONTAJ', küçük harf, Türkçe İ/I karışımı)
    return _EK_ADIM.get(_ek_anahtar(m.group(1)))


def tel_referans_kodu(referans_kodu, robot_no):
    """Bu hatta kaydedilecek NİHAİ referans kodunu üretir: '93.TK.464 KAPAMA'.

    Her proses adımı kendi ekini alır (bkz. TEL_ADIM_EKI açıklaması). Operatör
    eki elle yazdıysa tekrarlanmaz — önce ayıklanır, sonra hattın doğru eki
    konur; yanlış hatta yazılmış ek de böylece kendiliğinden düzelir.
    Hat tanınmıyorsa (tel dışı bir robot_no) koda DOKUNULMAZ."""
    base = tel_ek_ayikla(referans_kodu)
    if not base:
        return str(referans_kodu or '').strip()
    ek = TEL_ADIM_EKI.get(tel_hat_adimi(robot_no))
    return f'{base} {ek}' if ek else base


def tel_adim_etiketi(robot_no):
    """Hat → rapor/panel için okunabilir adım etiketi ('Kapama 3' → 'Kapama').
    Tel dışı hatlarda '' döner."""
    return tel_hat_adimi(robot_no) or ''
