# -*- coding: utf-8 -*-
"""
data/uretim_verileri.xlsx dosyasından referansları ve operatörleri
uretim.db veritabanına aktarır.

Excel yapısı (bölüm başına referans + operatör sayfası):
  - Kaynak Referans      | Kaynak Operator
  - Montaj Referans      | Montaj Operator
  - Metal Referans       | Metal Operator
  - İşleme Referans      | İşleme Operatör
  - Lazer Referans       | Lazer Operatör
  - Pres Abkant Referans | Pres Abkant Operatör

Her referans sayfası: 1. sütun kod, 2. sütun cycle time (sn).
Her operator sayfası: 2. sütun operatör adı (1. sütun No).
"""
import openpyxl
import sqlite3
import io
import os

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_YOL   = os.path.join(PROJECT_DIR, 'data', 'uretim_verileri.xlsx')
DB_PATH     = os.path.join(PROJECT_DIR, 'uretim.db')

BOLUM_SAYFA = {
    'kaynak': {'ref': 'Kaynak Referans', 'op': 'Kaynak Operator'},
    'montaj': {'ref': 'Montaj Referans', 'op': 'Montaj Operator'},
    'metal':  {'ref': 'Metal Referans',  'op': 'Metal Operator'},
    # Yeni TK2 bölümleri (2026-07): iş takibi/sayaç/andon YOK — sadece vardiya + üretim girişi.
    # Sayfa adları Excel'dekiyle BAYT-BAYT aynı olmalı (yenilerde Türkçe 'Operatör', eskilerde ASCII 'Operator').
    'isleme': {'ref': 'İşleme Referans',      'op': 'İşleme Operatör'},
    'lazer':  {'ref': 'Lazer Referans',       'op': 'Lazer Operatör'},
    'pres':   {'ref': 'Pres Abkant Referans', 'op': 'Pres Abkant Operatör'},
}

# Bölüm bazlı duruş sebepleri sayfaları.
# Yeni bölümlerin sayfası Excel'de henüz yok → durus_sebepleri_yukle sayfa bulunamazsa
# 'Montaj Duruş Listesi'ne düşer (genel sebepler). Kendi sayfası eklenirse otomatik kullanılır.
BOLUM_DURUS_SAYFA = {
    'kaynak': 'Robotik Kaynak Duruş Listesi',
    'montaj': 'Montaj Duruş Listesi',
    'metal':  'Metal Enjeksiyon Duruş Listesi',
    'isleme': 'İşleme Duruş Listesi',
    'lazer':  'Lazer Duruş Listesi',
    'pres':   'Pres Abkant Duruş Listesi',
}

# ── BÖLÜME ÖZEL EK DURUŞ SEBEPLERİ (kullanıcı 2026-08-18) ────────────────────
# pres / lazer / işleme bölümlerinin Excel'de KENDİ duruş sayfası yok → hepsi
# 'Montaj Duruş Listesi'ne düşüyor. Montaj bir MASA bölümü olduğu için o listede
# makine ayarı sebebi yok; bu üç bölümde ise ayar en sık duruşlardan biri ve
# operatör mecburen 'Setup Süresi' / 'Bakım Çalışması' gibi yanlış sebep seçiyordu.
# Sebep KODDA duruyor ki Excel'e sayfa eklenene kadar sahada seçilebilsin.
# Excel'e aynı adla eklenirse TEKRARLANMAZ (bkz. _ek_durus_uygula dedupe) —
# yani bu tablo Excel'i ezmez, yalnızca eksiği tamamlar.
# Tip 'planli': ayar planlı bir hazırlık duruşu (Metal Enjeksiyon sayfasındaki
# 'Makine Ayar Yapımı' da Planlı) → OEE'de plansız arıza gibi cezalandırılmaz.
BOLUM_EK_DURUS = {
    'isleme': [{'sebep': 'Makine Ayarı', 'tip': 'planli'}],
    'lazer':  [{'sebep': 'Makine Ayarı', 'tip': 'planli'}],
    # 'Malzeme Toplama' YALNIZ pres (kullanıcı 2026-08-20): abkantta operatör
    # sıradaki işin sacını depodan kendisi topluyor — montaj listesindeki
    # 'Malzeme Teslim' (birinin getirmesi) ve 'Malzeme Bekleme' (plansız, iş
    # duruyor) bunu karşılamıyor. Planlı: işin bilinen bir parçası.
    'pres':   [{'sebep': 'Makine Ayarı',    'tip': 'planli'},
               {'sebep': 'Malzeme Toplama', 'tip': 'planli'}],
}


def _durus_anahtar(sebep):
    """Duruş adı karşılaştırma anahtarı — boşluk/büyük-küçük harf farkına dayanıklı."""
    return ' '.join(str(sebep or '').split()).casefold()


def _ek_durus_uygula(bolum, lokasyon, sonuc):
    """BOLUM_EK_DURUS'taki sebepleri listeye ekler; Excel'de zaten varsa eklemez.
    TK1'de uygulanmaz: TK1'in duruş listesi bölümden bağımsız TEK listedir."""
    if (lokasyon or 'TK2').upper() == 'TK1':
        return sonuc
    ekler = BOLUM_EK_DURUS.get(bolum)
    if not ekler:
        return sonuc
    var_olan = {_durus_anahtar(s.get('sebep')) for s in sonuc}
    for e in ekler:
        if _durus_anahtar(e['sebep']) not in var_olan:
            sonuc.append(dict(e))
    return sonuc


# Ek sayfalar (kaynak bölümüne özel — diğer bölümler için gerekmez)
ROBOT_PROGRAM_SAYFA = 'Robot Program Listesi'
FIKSTUR_RAF_SAYFA   = 'Fikstür Raf Listesi'

# ── TK1 (yan tesis) — ayrı Excel, montaj mantığı, lokasyon='TK1' ──
TK1_EXCEL_YOL = os.path.join(PROJECT_DIR, 'data', 'Tk1 Veriler.xlsx')
TK1_SAYFA = {
    'ref':   'Refernaslar',     # tek kolon 'Ürün Kodu' (sayfa adı Excel'de bu typo ile)
    'op':    'Operatörler',     # BAŞLIK SATIRI YOK — row 0 = ilk operatör
    'durus': 'Duruş Listesi',   # No | Duruş Listesi | Planlı/Plansız
    # ── Plastik enjeksiyon (2026-08-04) ──
    # TK1'de montajdan AYRI bölüm (bolum='plastik'): makineleri 320T / 407T /
    # Yapistirma / Sizdirmazlik Test. Yapıştırma AYRI BÖLÜM DEĞİL, bu bölümün bir
    # makinesi — kodları (sonu 'G' ile biten) aynı referans sayfasında durur.
    # Sayfa adları Excel'dekiyle BİREBİR — ikisinde de yazım hatası var
    # ('Enjeksyion'), düzeltmeyin: Excel'de düzeltilirse burası da güncellenmeli.
    'plastik_ref': 'Plastik Enjeksiyon Referanslar',
    'plastik_op':  'Plastik Enjeksyion Operatörler',
    # ── Tel üretimi (2026-08-04) ──
    # TK1'de ayrı bölüm (bolum='tel'). Referans sayfası ÇOK KOLONLU: 1. kolon ürün
    # kodu, sonraki kolonlar PROSES ADIMLARI (başlık satırından okunur, işaretli
    # olanlar referansın adımları olur). Bkz. TEL_ADIM_BASLIK.
    'tel_ref': 'Tel Referanslar',
    'tel_op':  'Tel Operatörler',
}

# Tel referans sayfasının adım kolonları. Excel başlığı (küçük harfe indirgenmiş)
# → sistemdeki adım adı. Başlık yazımı esnek olsun diye birkaç varyant tanınır;
# tanınmayan kolon sessizce YOK SAYILIR (yeni kolon eklenirse buraya da eklenmeli).
TEL_ADIM_BASLIK = {
    # ADIM ADI 2026-08-26'da 'Halat/Spiral Kesme' oldu; Excel başlığı hangisi
    # yazılırsa yazılsın AYNI adıma düşer (eski dosyalar bozulmasın).
    'halat/spiral kesme': 'Halat/Spiral Kesme', 'halat spiral kesme': 'Halat/Spiral Kesme',
    'halat kesme': 'Halat/Spiral Kesme', 'kesim': 'Halat/Spiral Kesme',
    'halat kesim': 'Halat/Spiral Kesme',
    'manuel kesim': 'Halat/Spiral Kesme',   # makine adı ayrı, adım aynı (2026-08-20)
    # Spiral kesme (2026-08-21): 3 otomatik + 1 manuel makine, adım kesim
    # (kullanıcı kararı: ayrı sütun değil, kesime dahil).
    'spiral kesme': 'Halat/Spiral Kesme', 'spiral kesim': 'Halat/Spiral Kesme',
    'otomatik spiral kesme': 'Halat/Spiral Kesme',
    'manuel spiral kesme': 'Halat/Spiral Kesme',
    'soyma': 'Soyma',                # kesim yapılan her ürüne yapılmaz → AYRI adım
    # Ön hazırlık (2026-08-27) — 'otomatik hazırlık'tan AYRI adım, başlık da ayrı
    'ön hazırlık': 'Ön Hazırlık', 'on hazirlik': 'Ön Hazırlık',
    'ön hazirlik': 'Ön Hazırlık', 'on hazırlık': 'Ön Hazırlık',
    'otomatik hazırlık': 'Otomatik Hazırlık', 'otomatik hazirlik': 'Otomatik Hazırlık',
    'hazırlık': 'Otomatik Hazırlık', 'hazirlik': 'Otomatik Hazırlık',
    'yarı otomatik': 'Yarı Otomatik', 'yari otomatik': 'Yarı Otomatik',
    'yarı otomat': 'Yarı Otomatik', 'yari otomat': 'Yarı Otomatik',
    'tam otomatik': 'Tam Otomatik', 'tam otomat': 'Tam Otomatik',
    'otomatik pres': 'Tam Otomatik', 'otomatik pres makinesi': 'Tam Otomatik',
    'kapama': 'Kapama', 'hidrolik pres': 'Kapama',   # makine adı 2026-08-26
    'son montaj': 'Son Montaj', 'montaj': 'Son Montaj',
}
# Hücre "bu adım var" mı diyor? X / x / 1 / EVET / VAR / ✓ kabul edilir.
_TEL_ISARET = {'x', '1', 'evet', 'var', '✓', '✔', 'e', 'yes', 'true'}


# Üretim sırası — tel_adimlar bu sıraya göre yazılır (app.TEL_ADIMLARI ile aynı).
# Sıralı tutmak "son adım" hesabını kolaylaştırır ve panelde okunaklı gösterir.
# 2026-08-20: 'Soyma' + 'Otomatik Hazırlık' eklendi — tel_proses.TEL_ADIMLARI ile
# BİREBİR olmalı, yoksa Excel'de işaretlenen adım panelde sırasız/eksik görünür.
TEL_ADIM_SIRASI = ('Halat/Spiral Kesme', 'Soyma', 'Ön Hazırlık', 'Otomatik Hazırlık',
                   'Yarı Otomatik', 'Tam Otomatik', 'Kapama', 'Son Montaj')


def _tel_isaretli_mi(hucre):
    if hucre is None:
        return False
    s = str(hucre).strip().lower()
    if not s:
        return False
    if s in _TEL_ISARET:
        return True
    # Sayısal 1 (openpyxl int/float döndürebilir)
    try:
        return float(s) == 1.0
    except ValueError:
        return False

# Başlık satırı tespiti (2026-08-04): TK1 sayfalarının bazısında başlık VAR
# ('Refernaslar' → 'Ürün Kodu'), bazısında YOK ('Operatörler' → row 0 = ilk kişi).
# Yeni plastik sayfalarında hangisinin olacağı garanti değil; sabit "ilk satırı atla"
# kuralı başlıksız sayfada İLK KAYDI YUTAR. Bu yüzden içeriğe bakıyoruz.
_TK1_BASLIKLAR = {
    'ürün kodu', 'urun kodu', 'ürün kod', 'parça kodu', 'parca kodu', 'parça kod',
    'kod', 'kodu', 'referans', 'referans kodu', 'referanslar',
    'ad', 'adi', 'adı', 'isim', 'ad soyad', 'operatör', 'operator',
    'operatörler', 'operatorler', 'no',
}


def _tk1_baslik_mi(deger):
    """Hücre bir başlık mı, veri mi? (TK1 sayfalarında başlık satırı tutarsız)"""
    return str(deger or '').strip().lower() in _TK1_BASLIKLAR


# ── TK1 OPERATÖR DAĞILIMI (kullanıcı 2026-08-04) ────────────────────────────
# Excel'deki tek 'Operatörler' sayfası TK1'in TÜM operatörlerini tutuyor; bölüm
# ayrımı Excel'de yok, kural burada:
#   montaj  → yalnız aşağıdaki iki kişi
#   tel     → kalan HERKES
#   plastik → kendi sayfasından gelir (Mustafa Kaya, Musa Kolip)
# Ad karşılaştırması Türkçe karakter ve boşluk farklarına DAYANIKLI (_ad_normal):
# Excel'de 'BİRCAN KILIÇ' / 'Bircan Kilic' / çift boşluk hepsi eşleşir.
TK1_MONTAJ_OPERATORLERI = ('BİRCAN KILIÇ', 'OSMAN İMAT')


def _ad_normal(ad):
    """Ad eşleştirme anahtarı: Türkçe karakterler sadeleşir, boşluklar tekilleşir."""
    s = str(ad or '').strip().upper()
    for a, b in (('İ', 'I'), ('Ş', 'S'), ('Ğ', 'G'), ('Ü', 'U'), ('Ö', 'O'), ('Ç', 'C')):
        s = s.replace(a, b)
    return ' '.join(s.split())


_TK1_MONTAJ_NORMAL = {_ad_normal(a) for a in TK1_MONTAJ_OPERATORLERI}


def tk1_operator_bolumu(ad):
    """TK1 ana operatör sayfasındaki bir isim hangi bölüme ait? 'montaj' | 'tel'"""
    return 'montaj' if _ad_normal(ad) in _TK1_MONTAJ_NORMAL else 'tel'


# ── SON BİLİNEN İYİ KOPYA (kullanıcı 2026-08-26 arıza bildirimi) ─────────
# Sunucudaki data/uretim_verileri.xlsx bozuldu (zip açılamıyor:
# "Error -3 while decompressing data") ve TK2'nin BÜTÜN bölümlerinde operatör
# ekranı "duruş listesi sunucudan gelmedi" deyip 6 maddelik yedek listeye
# düştü. Operatörler o listeyle duruş kaydetmeye devam ediyor → veri sessizce
# bozuluyor (tam olarak 2026-08-18'de önlenmeye çalışılan durum, bu kez
# dosyanın kendisi bozulduğu için).
#
# ÇÖZÜM: Excel her BAŞARILI okunduğunda dosyanın bir kopyası yedeklenir.
# Ana dosya açılamazsa liste YEDEKTEN okunur — operatör tam listeyi görmeye
# devam eder, panel/API ise durumu AÇIKÇA uyarı olarak bildirir.
# Yedek asla otomatik geri yazılmaz: bozuk dosyayı sessizce onarmak, kimsenin
# fark etmediği eski bir listeyle çalışmak demek olurdu.
def _yedek_yolu(excel_yol):
    kok, uzanti = os.path.splitext(excel_yol)
    return kok + '.sonbilinen' + uzanti


def _yedegi_tazele(excel_yol, ham):
    """Başarıyla okunan dosyayı yedekle (aynıysa yazma)."""
    try:
        y = _yedek_yolu(excel_yol)
        if os.path.exists(y) and os.path.getsize(y) == len(ham):
            return
        gecici = y + '.tmp'
        with open(gecici, 'wb') as f:
            f.write(ham)
        os.replace(gecici, y)      # atomik: yarım yedek oluşmaz
    except Exception as e:
        print(f'[durus_sebepleri] yedek yazilamadi: {e}')


def _yedekten_oku(excel_yol):
    """(workbook, yedek_tarihi) ya da (None, '')."""
    y = _yedek_yolu(excel_yol)
    if not os.path.exists(y):
        return None, ''
    try:
        with open(y, 'rb') as f:
            wb = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True)
        from datetime import datetime as _dt
        return wb, _dt.fromtimestamp(os.path.getmtime(y)).strftime('%d.%m.%Y %H:%M')
    except Exception as e:
        print(f'[durus_sebepleri] yedek de okunamadi: {e}')
        return None, ''


# Son okumanın yedekten mi geldiği — /api/durus_sebepleri uyarı metnine koyar.
SON_OKUMA_YEDEKTEN = {}

# OKUMA ÖNBELLEĞİ (2026-10-05): duruş listesi operatör ekranında HER istekte okunuyor;
# Ana Veri tek dosyada ~3.400 referans satırı taşıdığı için her seferinde baştan açmak
# pahalı. Dosya değişmediği sürece (mtime + boyut) aynı salt-okunur kitap kullanılır.
_WB_ONBELLEK = {}


def _wb_onbellekten(yol):
    """data_only=True kitabı; dosya değişince yeniden okunur. Başarılı okuma yedeklenir."""
    st = os.stat(yol)
    anahtar = (st.st_mtime_ns, st.st_size)
    kayit = _WB_ONBELLEK.get(yol)
    if kayit and kayit[0] == anahtar:
        return kayit[1]
    with open(yol, 'rb') as fh:
        veri = fh.read()
    wb = openpyxl.load_workbook(io.BytesIO(veri), data_only=True)
    _yedegi_tazele(yol, veri)
    _WB_ONBELLEK[yol] = (anahtar, wb)
    return wb


def durus_sebepleri_yukle(bolum, lokasyon='TK2'):
    """Bölüme (TK2) veya lokasyona (TK1) ait duruş sebeplerini Excel'den okur.

    Sayfa formatı: No | Duruş Listesi | Planlı/Plansız
    Döner: [{'sebep': str, 'tip': 'planli'|'plansiz'}, ...]

    TK1 → data/Tk1 Veriler.xlsx 'Duruş Listesi' (bolum'dan bağımsız, tek liste).
    TK2 (default) → data/AnaVeri.xlsx (varsa) ya da data/uretim_verileri.xlsx,
    bolum-spesifik sayfa. Excel/sayfa yoksa boş liste.
    """
    if (lokasyon or 'TK2').upper() == 'TK1':
        excel_yol = TK1_EXCEL_YOL
        sayfa_adi = TK1_SAYFA['durus']
    else:
        if bolum not in BOLUM_DURUS_SAYFA:
            return []
        excel_yol = tk2_excel_yolu()           # Ana Veri varsa o (tek dosya)
        sayfa_adi = BOLUM_DURUS_SAYFA[bolum]
        if excel_yol != EXCEL_YOL and os.path.exists(EXCEL_YOL):
            # Ana Veri'de bölümün sayfası yoksa (elle konmuş eksik dosya) önce eski
            # dosyadaki AYNI sayfa, o da yoksa genel (montaj) liste — operatör ASLA
            # boş listeyle ya da yanlış bölümün listesiyle kalmasın.
            for yol, dus in ((excel_yol, False), (EXCEL_YOL, False), (excel_yol, True), (EXCEL_YOL, True)):
                sonuc = _durus_oku(yol, sayfa_adi, bolum, lokasyon, montaj_dus=dus)
                if sonuc is not None:
                    return sonuc
            return []
    return _durus_oku(excel_yol, sayfa_adi, bolum, lokasyon) or []


def _durus_sayfa_oku(ws):
    """Duruş sayfası (No | Duruş Listesi | Planlı/Plansız) → [{'sebep', 'tip'}]."""
    sonuc = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue  # Başlık
        if not row or len(row) < 2 or row[1] is None:
            continue
        sebep = str(row[1]).strip()
        if not sebep:
            continue
        # Türkçe ı/i karakter farkı: 'Plansız' lower → 'plansız' (dotless ı),
        # ama 'siz' substring'i (regular i) bulunmaz. Bu yüzden açık string match.
        tip_raw = str((row[2] if len(row) > 2 else '') or '').strip().lower()
        if 'plansız' in tip_raw or 'plansiz' in tip_raw:
            tip = 'plansiz'
        elif 'planlı' in tip_raw or 'planli' in tip_raw:
            tip = 'planli'
        else:
            tip = 'plansiz'  # boş/bilinmeyen → güvenli yan: plansız
        sonuc.append({'sebep': sebep, 'tip': tip})
    return sonuc


def _durus_oku(excel_yol, sayfa_adi, bolum, lokasyon, montaj_dus=True):
    """Liste; dosya/sayfa YOKSA None (çağıran başka dosyaya düşebilsin).
    montaj_dus: bölümün sayfası yoksa genel 'Montaj Duruş Listesi'ne düşülsün mü."""
    if not os.path.exists(excel_yol):
        return None

    try:
        # DOSYAYI BELLEGE OKU, openpyxl"e BytesIO ver (2026-08-18).
        # NEDEN: bozuk bir xlsx"te openpyxl uye okurken BadZipFile firlatiyor ve
        # arsivi ACIK birakiyor -> her istekte bir tutamak sizip dosya KILITLENIYOR.
        # Sahada yasandi: bozuk dosya ne tasinabildi ne degistirilebildi
        # ("baska bir islem tarafindan kullaniliyor"). BytesIO ile OS tutamagi
        # "with" bitince kapanir; bozuk dosya bile kendini kilitlemez.
        SON_OKUMA_YEDEKTEN.pop(excel_yol, None)
        try:
            wb = _wb_onbellekten(excel_yol)
        except Exception as _ex:
            # ANA DOSYA AÇILAMADI → son bilinen iyi kopyayı dene. Operatörün
            # tam listeyi görmesi, 6 maddelik yedekle yanlış sebep kaydetmesinden
            # iyidir; durum uyarı olarak yukarı taşınır.
            wb, _ytar = _yedekten_oku(excel_yol)
            if wb is None:
                raise
            SON_OKUMA_YEDEKTEN[excel_yol] = _ytar
            print(f'[durus_sebepleri] ANA DOSYA BOZUK ({_ex}) — yedekten okundu ({_ytar})')
        if sayfa_adi not in wb.sheetnames:
            # Bölümün kendi duruş sayfası yoksa genel listeye düş (yeni bölümler:
            # işleme/lazer/pres — Excel'e kendi sayfaları eklenince otomatik geçilir).
            # TK1'in tek listesi var — montaj düşüşü yalnız TK2'de anlamlı.
            sayfa_adi = (BOLUM_DURUS_SAYFA.get('montaj', '')
                         if montaj_dus and (lokasyon or 'TK2').upper() != 'TK1' else '')
            if sayfa_adi not in wb.sheetnames:
                return None
        sonuc = _durus_sayfa_oku(wb[sayfa_adi])
        # Bölüme özel ek sebepler — YALNIZ Excel'den gerçek bir liste okunduysa.
        # Boş listeye eklemek olmaz: '/api/durus_sebepleri' boş listeyi "Excel
        # okunamadı" tanısı için kullanıyor, tek maddelik liste o tanıyı susturur
        # ve operatör 1 sebeple baş başa kalırdı (mobil yedek listesi de devreye girmez).
        if sonuc:
            _ek_durus_uygula(bolum, lokasyon, sonuc)
        return sonuc
    except Exception as e:
        print(f"[durus_sebepleri] Hata: {e}")
        return []


# ── ANA VERİ (kullanıcı 2026-10-05) ────────────────────────────────────────
# TK2 referanslarının TEK listesi: planlamanın Anaveri'si (PRIORIT. · A/P · TK ·
# Makine) ile Forge'un süreleri aynı sayfada; bölüm E sütunundan okunur
# (Montaj / Kaynak / Metal Enjeksiyon / Lazer Kesim / Büküm / İşleme).
# data/AnaVeri.xlsx VARSA altı TK2 bölümünün referansları YALNIZ buradan okunur ve
# buraya yazılır; uretim_verileri.xlsx'in '<Bölüm> Referans' sayfaları kullanılmaz
# (operatör, duruş, robot program, fikstür sayfaları yine oradan okunur).
# Dosya YOKSA eski davranış birebir sürer → kod, dosya sunucuya konmadan önce de
# güvenle deploy edilir.
#
# Eski bölüm sayfalarından BİLİNÇLİ farklar:
#   · BOŞ HÜCRE = DOKUNMA (eskiden 0 yazılırdı); süreyi silmek için 0 yazılır.
#     Dosya artık panelden indirilip düzenlenip geri yükleniyor — boş kalmış bir
#     hücre panelde girilmiş süreyi sessizce silmemeli.
#   · Açıklama (K) her bölümde okunur/yazılır (eskiden yalnız pres).
#   · Kalıp göz (metal) Excel'den de yönetilir; panelle aynı kural: değişince açık
#     otomatik sayaç kayıtlarına da uygulanır.
ANA_VERI_YOL = os.path.join(PROJECT_DIR, 'data', 'AnaVeri.xlsx')
ANA_VERI_SAYFA = 'Ana Veri'
ANA_VERI_YEDEK_KLASOR = os.path.join(PROJECT_DIR, 'data', 'ana_veri_yedek')
ANA_VERI_ETIKET = {'montaj': 'Montaj', 'kaynak': 'Kaynak', 'metal': 'Metal Enjeksiyon',
                   'lazer': 'Lazer Kesim', 'pres': 'Büküm', 'isleme': 'İşleme'}
ANA_VERI_KALIP_GOZ_UST = 64          # app.KALIP_GOZ_UST ile aynı olmalı
# Başlık → alan. Başlık küçük harfe (TR) indirilir; aday ile BAŞLAYAN ilk sütun alınır.
_ANA_SUTUN = (
    ('kod', ('cd art', 'referans kodu', 'kod')),
    ('bolum', ('bölüm', 'bolum', 'hat')),
    ('ap', ('a/p',)),
    ('tk', ('tk-1/2', 'tk 1/2', 'tk')),
    ('cevrim', ('çevrim', 'cevrim', 'cycle', 'söktak')),
    ('kaynak', ('kaynak süre', 'kaynak sure')),
    ('goz', ('kalıp göz', 'kalip goz')),
    ('bukum', ('büküm', 'bukum')),
    ('aciklama', ('açıklama', 'aciklama')),
    ('teyit', ('süre teyit', 'sure teyit', 'teyit')),
    ('not', ('not',)),
)
ANA_VERI_INDIRILEN_KLASOR = os.path.join(PROJECT_DIR, 'data', 'ana_veri_indirilen')
ANA_VERI_DAMGA_ADI = 'forge_indirme'       # Excel özel belge özelliği (Dosya > Bilgi > Özellikler)
_ANA_YENI_BASLIK = {'cevrim': 'Çevrim süresi / söktak (sn)', 'kaynak': 'Kaynak süresi (sn)',
                    'goz': 'Kalıp göz', 'bukum': 'Büküm op.', 'aciklama': 'Açıklama',
                    'teyit': 'Süre teyit'}


def ana_veri_aktif():
    return os.path.exists(ANA_VERI_YOL)


def tk2_excel_yolu():
    """TK2'nin TEK Excel'i: Ana Veri varsa o, yoksa eski uretim_verileri.xlsx.
    Referanslar, operatörler, duruş listeleri, robot program ve fikstür buradan okunur
    (Ana Veri'de olmayan sayfa eski dosyadan — bkz. _SayfaKumesi)."""
    return ANA_VERI_YOL if ana_veri_aktif() else EXCEL_YOL


def ana_veri_ek_sayfalar():
    """Ana Veri'nin referans dışı sayfaları (eski uretim_verileri.xlsx'teki adlarıyla AYNI —
    okuyan kod sayfayı adıyla bulur)."""
    ad = [BOLUM_SAYFA[b]['op'] for b in BOLUM_SAYFA]
    for s in BOLUM_DURUS_SAYFA.values():
        if s not in ad:
            ad.append(s)
    return ad + [ROBOT_PROGRAM_SAYFA, FIKSTUR_RAF_SAYFA]


class _SayfaKumesi:
    """Birden çok kitaptan sayfa seçen salt-okunur görünüm: sayfa İLK hangi kitapta varsa
    oradan gelir (Ana Veri önce, eski dosya sonra). openpyxl kitabı gibi kullanılır
    (sheetnames + [ad])."""

    def __init__(self, *kitaplar):
        self._kitaplar = [k for k in kitaplar if k is not None]

    @property
    def sheetnames(self):
        adlar = []
        for k in self._kitaplar:
            adlar.extend(s for s in k.sheetnames if s not in adlar)
        return adlar

    def __getitem__(self, ad):
        for k in self._kitaplar:
            if ad in k.sheetnames:
                return k[ad]
        raise KeyError(ad)


def _tk2_okuma_kumesi():
    """Ana Veri (varsa) + eski dosya (varsa) okuma görünümü; ikisi de yoksa None."""
    kitaplar = []
    if ana_veri_aktif():
        kitaplar.append(openpyxl.load_workbook(ANA_VERI_YOL, data_only=True))
    if os.path.exists(EXCEL_YOL):
        kitaplar.append(openpyxl.load_workbook(EXCEL_YOL, data_only=True))
    if not kitaplar:
        return None
    return _SayfaKumesi(*kitaplar) if len(kitaplar) > 1 else kitaplar[0]


def _sayfa_kopyala(ws_kaynak, wb_hedef, ad):
    """Değerler + birleştirilmiş hücreler + sütun genişlikleri (biçim değil)."""
    ws = wb_hedef.create_sheet(ad)
    for row in ws_kaynak.iter_rows(values_only=True):
        ws.append(list(row))
    for rng in getattr(ws_kaynak, 'merged_cells', None) and ws_kaynak.merged_cells.ranges or ():
        ws.merge_cells(str(rng))
    for harf, boyut in getattr(ws_kaynak, 'column_dimensions', {}).items():
        if boyut.width:
            ws.column_dimensions[harf].width = boyut.width
    return ws


# ── İNDİRME DAMGASI / ÜÇ YÖNLÜ KARŞILAŞTIRMA (kullanıcı 2026-10-05) ───────────────
# OLAY: sabah hazırlanan dosya öğleden sonra yüklenince önizleme, o arada operatörün
# açtığı 2 kodu SİLMEK ve mobilden 1→2 yapılmış büküm op.'u GERİ ALMAK istedi.
# Dosya-↔-veritabanı iki yönlü karşılaştırmada "kullanıcı mı değiştirdi, yoksa dosya
# mı eski?" ayırt edilemez. Çözüm: 'Ana Veri İndir' dosyaya bir damga koyar ve İNDİRİLEN
# HÂLİ sunucuda saklar. Yüklemede taban = o hâl:
#   · hücre tabandakiyle AYNIYSA kullanıcı dokunmamıştır → veritabanı değeri korunur
#   · tabanda olmayan referans (indirildikten SONRA açılmış) silinmez
#   · tabanda olup veritabanından silinmiş referans (panelden silinmiş) geri eklenmez
# Damgasız dosyada (elle hazırlanmış) eski iki yönlü karşılaştırma geçerlidir.


def ana_veri_damgali_indir():
    """İndirilecek hâl: özel belge özelliğine damga yazılır, aynı hâl sunucuda saklanır
    (ana_veri_indirilen/, son 40). → (bytes, damga)"""
    from datetime import datetime as _dt
    from openpyxl.packaging.custom import StringProperty
    import secrets
    damga = 'AV-' + _dt.now().strftime('%Y%m%d-%H%M%S') + '-' + secrets.token_hex(3)
    with open(ANA_VERI_YOL, 'rb') as fh:
        wb = openpyxl.load_workbook(io.BytesIO(fh.read()))
    if ANA_VERI_DAMGA_ADI in wb.custom_doc_props.names:
        del wb.custom_doc_props[ANA_VERI_DAMGA_ADI]
    wb.custom_doc_props.append(StringProperty(name=ANA_VERI_DAMGA_ADI, value=damga))
    bio = io.BytesIO()
    wb.save(bio)
    veri = bio.getvalue()
    os.makedirs(ANA_VERI_INDIRILEN_KLASOR, exist_ok=True)
    with open(os.path.join(ANA_VERI_INDIRILEN_KLASOR, damga + '.xlsx'), 'wb') as fh:
        fh.write(veri)
    eskiler = sorted(f for f in os.listdir(ANA_VERI_INDIRILEN_KLASOR) if f.startswith('AV-'))
    for f in eskiler[:-40]:
        try:
            os.remove(os.path.join(ANA_VERI_INDIRILEN_KLASOR, f))
        except OSError:
            pass
    return veri, damga


def _damga_oku(wb):
    try:
        if ANA_VERI_DAMGA_ADI in wb.custom_doc_props.names:
            return str(wb.custom_doc_props[ANA_VERI_DAMGA_ADI].value or '').strip() or None
    except Exception:
        pass
    return None


def _damga_zamani(damga):
    """'AV-20261005-103122-ab12cd' → '05.10.2026 10:31'"""
    try:
        from datetime import datetime as _dt
        return _dt.strptime(damga[3:18], '%Y%m%d-%H%M%S').strftime('%d.%m.%Y %H:%M')
    except Exception:
        return ''


def _taban_oku(damga):
    """Damganın indirilen hâli → {(norm, bolum): satır} | None (saklanan kopya yoksa)."""
    if not damga or not all(ch.isalnum() or ch == '-' for ch in damga):
        return None
    yol = os.path.join(ANA_VERI_INDIRILEN_KLASOR, damga + '.xlsx')
    if not os.path.exists(yol):
        return None
    satirlar, _r = ana_veri_oku(openpyxl.load_workbook(yol, data_only=True))
    return {(_norm_kod(s['kod']), b): s for b, liste in satirlar.items() for s in liste}


def _ayni_deger(a, b):
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) < 0.001
    return str(a).strip() == str(b).strip()


def ana_veri_hazirla(ham):
    """Yüklenen dosyayı hazırlar: referans satırları + rapor + ek sayfaların GEÇERLİ görünümü.
    Dosyada OLMAYAN ek sayfa (operatör / duruş / robot / fikstür) şu an kullanılan
    kaynaktan (mevcut Ana Veri, yoksa uretim_verileri.xlsx) KOPYALANIR — yalnız
    referans sayfası yüklense bile hiçbir liste kaybolmaz. → dict"""
    wb_yeni = openpyxl.load_workbook(io.BytesIO(ham), data_only=True)
    satirlar, rapor = ana_veri_oku(wb_yeni)
    damga = _damga_oku(wb_yeni)
    taban = _taban_oku(damga) if damga else None
    mevcut = _tk2_okuma_kumesi()
    eksik = [s for s in ana_veri_ek_sayfalar()
             if s not in wb_yeni.sheetnames and mevcut is not None and s in mevcut.sheetnames]
    kaydedilecek = ham
    if eksik:
        wb_yaz = openpyxl.load_workbook(io.BytesIO(ham))       # formüller korunsun
        for s in eksik:
            _sayfa_kopyala(mevcut[s], wb_yaz, s)
        bio = io.BytesIO()
        wb_yaz.save(bio)
        kaydedilecek = bio.getvalue()
    return {'satirlar': satirlar, 'rapor': rapor, 'eksik': eksik, 'kaydedilecek': kaydedilecek,
            'kume': _SayfaKumesi(wb_yeni, mevcut), 'mevcut': mevcut, 'taban': taban,
            'damga': {'var': bool(damga), 'zaman': _damga_zamani(damga) if damga else '',
                      'taban': taban is not None}}


def ana_veri_durus_farki(kume, mevcut):
    """Dosyadaki duruş sayfaları ↔ şu an kullanılan listeler (duruş listesi veritabanına
    aktarılmaz, her istekte okunur — yüklenince HEMEN geçerli olur)."""
    out = []
    for sayfa in dict.fromkeys(BOLUM_DURUS_SAYFA.values()):
        if sayfa not in kume.sheetnames:
            continue
        yeni = {x['sebep']: x['tip'] for x in _durus_sayfa_oku(kume[sayfa])}
        eski = ({x['sebep']: x['tip'] for x in _durus_sayfa_oku(mevcut[sayfa])}
                if mevcut is not None and sayfa in mevcut.sheetnames else {})
        out.append({'sayfa': sayfa, 'dosyada': len(yeni), 'mevcut': len(eski),
                    'eklenen': [s for s in yeni if s not in eski],
                    'cikan': [s for s in eski if s not in yeni],
                    'tip_degisen': [s for s in yeni if s in eski and yeni[s] != eski[s]]})
    return out


def _tr_kucuk(s):
    return str(s or '').strip().replace('İ', 'i').replace('I', 'ı').lower()


def _norm_kod(kod):
    return str(kod or '').strip().upper().replace(' ', '')


def ana_veri_bolum(deger):
    """E sütunundaki metin → bölüm anahtarı (tanınmazsa None)."""
    s = _tr_kucuk(deger)
    if not s:
        return None
    if 'montaj' in s:
        return 'montaj'
    if 'kaynak' in s:
        return 'kaynak'
    if 'metal' in s:
        return 'metal'
    if 'lazer' in s:
        return 'lazer'
    if any(k in s for k in ('büküm', 'bukum', 'pres', 'abkant')):
        return 'pres'
    if 'işleme' in s or 'isleme' in s:
        return 'isleme'
    return None


def _ana_sutunlar(baslik_satiri):
    basliklar = [_tr_kucuk(x) for x in (baslik_satiri or ())]
    kol = {}
    for alan, adaylar in _ANA_SUTUN:
        for i, h in enumerate(basliklar):
            if h and i not in kol.values() and any(h.startswith(a) for a in adaylar):
                kol[alan] = i
                break
    return kol


def _ana_sayfa_bul(wb):
    """('Ana Veri' sayfası ya da kod + bölüm başlığı olan ilk sayfa, sütun haritası)."""
    adaylar = ([wb[ANA_VERI_SAYFA]] if ANA_VERI_SAYFA in wb.sheetnames else []) + list(wb.worksheets)
    for ws in adaylar:
        try:
            ilk = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        except StopIteration:
            continue
        kol = _ana_sutunlar(ilk)
        if 'kod' in kol and 'bolum' in kol:
            return ws, kol
    return None, {}


def _ana_sayi(v):
    """Hücre → float; boş = None (dokunma); sayı değilse ValueError."""
    if isinstance(v, bool):
        raise ValueError('sayı değil')
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(',', '.')
    if s in ('', '-'):
        return None
    return float(s)


def ana_veri_oku(wb):
    """Ana Veri sayfasını okur → ({bolum: [satır]}, rapor). Satır alanı None = boş hücre.
    rapor: sayfa, satir, bolumsuz[], tekrar[], hatali[], tel_atlanan."""
    ws, kol = _ana_sayfa_bul(wb)
    if ws is None:
        raise ValueError("'Ana Veri' sayfası bulunamadı — 1. satırda en az 'CD ART.' (kod) ve "
                         "'Bölüm' başlıklı sütun olmalı")
    satirlar = {b: [] for b in ANA_VERI_ETIKET}
    rapor = {'sayfa': ws.title, 'satir': 0, 'bolumsuz': [], 'tekrar': [], 'hatali': [], 'tel_atlanan': 0}
    gorulen = set()

    def al(row, alan):
        i = kol.get(alan)
        return row[i] if (i is not None and i < len(row)) else None

    for no, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row:
            continue
        kod = str(al(row, 'kod') or '').strip()
        if len(kod) < 2:
            continue
        rapor['satir'] += 1
        b = ana_veri_bolum(al(row, 'bolum'))
        if not b:
            rapor['bolumsuz'].append({'satir': no, 'kod': kod, 'deger': str(al(row, 'bolum') or '')})
            continue
        if kod.startswith('93.'):            # TK1 tel kodu TK2 listesine girmez (bkz. _bolum_import)
            rapor['tel_atlanan'] += 1
            continue
        anahtar = (_norm_kod(kod), b)
        if anahtar in gorulen:
            rapor['tekrar'].append({'satir': no, 'kod': kod, 'bolum': ANA_VERI_ETIKET[b]})
            continue
        try:
            d = {'kod': kod, 'satir': no, 'cevrim': _ana_sayi(al(row, 'cevrim')),
                 'kaynak': _ana_sayi(al(row, 'kaynak')), 'goz': _ana_sayi(al(row, 'goz')),
                 'bukum': _ana_sayi(al(row, 'bukum'))}
        except (TypeError, ValueError):
            rapor['hatali'].append({'satir': no, 'kod': kod, 'sebep': 'süre / kalıp göz / büküm sayı değil'})
            continue
        if any((d[k] or 0) < 0 for k in ('cevrim', 'kaynak', 'goz', 'bukum')):
            rapor['hatali'].append({'satir': no, 'kod': kod, 'sebep': 'negatif değer'})
            continue
        gorulen.add(anahtar)
        ack = al(row, 'aciklama')
        d['aciklama'] = (str(ack).strip() or None) if ack is not None else None
        satirlar[b].append(d)
    return satirlar, rapor


def _ana_veri_bolum_uygula(conn, bolum, satirlar, ornek_n=12, taban=None):
    """Bir bölümün Ana Veri satırlarını referans_listesi'ne uygular (TK2) — commit ETMEZ.
    Ayna senkronu: dosyada olmayan referans bu bölümden silinir (dosyada bölümün HİÇ
    satırı yoksa silme yapılmaz). taban (indirilen hâl, bkz. İNDİRME DAMGASI) verilirse
    üç yönlü: yalnız kullanıcının değiştirdiği hücre uygulanır, indirildikten sonra
    açılan referans silinmez, panelden silinmiş referans geri eklenmez.
    Döner: sayılar + örnekler (önizleme için)."""
    son = {'referanslar_eklenen': 0, 'referanslar_guncellenen': 0, 'referanslar_ayni': 0,
           'referanslar_silinen': 0, 'referanslar_korunan': 0, 'referanslar_geri_eklenmeyen': 0,
           'kalip_acik_kayit': 0, 'dosyada': len(satirlar),
           'ornek_eklenen': [], 'ornek_degisen': [], 'ornek_silinen': [], 'ornek_korunan': []}
    son['mevcut'] = conn.execute(
        "SELECT COUNT(*) FROM referans_listesi WHERE COALESCE(bolum,'kaynak')=? "
        "AND COALESCE(lokasyon,'TK2')='TK2'", (bolum,)).fetchone()[0]
    excel_norm = set()
    for s in satirlar:
        kod = s['kod']
        excel_norm.add(_norm_kod(kod))
        t = taban.get((_norm_kod(kod), bolum)) if taban is not None else None
        if t is not None:
            # Tabandakiyle aynı hücre = kullanıcı dokunmamış → veritabanındaki değer kalsın
            s = dict(s)
            for alan in ('cevrim', 'kaynak', 'goz', 'bukum', 'aciklama'):
                if s[alan] is not None and t.get(alan) is not None and _ayni_deger(s[alan], t[alan]):
                    s[alan] = None
        m = conn.execute(
            "SELECT id, referans_kodu, COALESCE(hedef_cycle_time_sn,0), COALESCE(kaynak_suresi_sn,0), "
            "COALESCE(soktak_suresi_sn,0), COALESCE(aciklama,''), COALESCE(bukum_operasyon,1), "
            "COALESCE(kalip_goz,1) FROM referans_listesi "
            "WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ','')) "
            "AND COALESCE(bolum,'kaynak')=? AND COALESCE(lokasyon,'TK2')='TK2' "
            "ORDER BY (referans_kodu = ?) DESC, id LIMIT 1", (kod, bolum, kod)).fetchone()
        yeni = {}
        if bolum == 'kaynak':
            # G = söktak, H = kaynak; çevrim = toplam. Biri boşsa mevcut değeri korunur.
            if s['kaynak'] is not None or s['cevrim'] is not None:
                ks = s['kaynak'] if s['kaynak'] is not None else (float(m[3]) if m else 0.0)
                ss = s['cevrim'] if s['cevrim'] is not None else (float(m[4]) if m else 0.0)
                yeni.update(kaynak_suresi_sn=ks, soktak_suresi_sn=ss, hedef_cycle_time_sn=round(ks + ss, 2))
        elif s['cevrim'] is not None:
            yeni['hedef_cycle_time_sn'] = s['cevrim']
        if bolum == 'pres' and s['bukum'] is not None:
            yeni['bukum_operasyon'] = max(1, min(99, int(s['bukum'])))
        if bolum == 'metal' and s['goz'] is not None:
            yeni['kalip_goz'] = max(1, min(ANA_VERI_KALIP_GOZ_UST, int(s['goz'])))
        if s['aciklama'] is not None:
            yeni['aciklama'] = s['aciklama']

        if m:
            eski = {'hedef_cycle_time_sn': m[2], 'kaynak_suresi_sn': m[3], 'soktak_suresi_sn': m[4],
                    'aciklama': m[5], 'bukum_operasyon': m[6], 'kalip_goz': m[7]}
            fark = {}
            for k, v in yeni.items():
                if k == 'aciklama':
                    if (eski[k] or '') != v:
                        fark[k] = v
                elif abs(float(eski[k] or 0) - float(v)) > 0.001:
                    fark[k] = v
            if m[1] != kod:                  # yazım Excel'deki hâline çekilir (eski import gibi)
                fark['referans_kodu'] = kod
            if not fark:
                son['referanslar_ayni'] += 1
                continue
            conn.execute(f"UPDATE referans_listesi SET {', '.join(k + '=?' for k in fark)} WHERE id=?",
                         list(fark.values()) + [m[0]])
            son['referanslar_guncellenen'] += 1
            if len(son['ornek_degisen']) < ornek_n:
                son['ornek_degisen'].append({'kod': kod, 'alanlar': {
                    k: [m[1] if k == 'referans_kodu' else eski.get(k), v] for k, v in fark.items()}})
        elif t is not None:
            # İndirilen dosyada vardı ama veritabanında yok → o arada panelden silinmiş;
            # kullanıcı satıra dokunmadıysa geri getirme.
            son['referanslar_geri_eklenmeyen'] += 1
            continue
        else:
            conn.execute(
                "INSERT INTO referans_listesi (referans_kodu, hedef_cycle_time_sn, kaynak_suresi_sn, "
                "soktak_suresi_sn, aciklama, bolum, lokasyon, bukum_operasyon, kalip_goz) "
                "VALUES (?, ?, ?, ?, ?, ?, 'TK2', ?, ?)",
                (kod, yeni.get('hedef_cycle_time_sn', 0), yeni.get('kaynak_suresi_sn', 0),
                 yeni.get('soktak_suresi_sn', 0), yeni.get('aciklama', ''), bolum,
                 yeni.get('bukum_operasyon', 1), yeni.get('kalip_goz', 1)))
            son['referanslar_eklenen'] += 1
            if len(son['ornek_eklenen']) < ornek_n:
                son['ornek_eklenen'].append(kod)
            fark = yeni
        ct = fark.get('hedef_cycle_time_sn')
        if ct and ct > 0:                     # geçmiş üretim kayıtlarının cycle'ı (yalnız bu bölüm, TK2)
            conn.execute(
                "UPDATE uretim_kayitlari SET cycle_time_sn = ? "
                "WHERE UPPER(REPLACE(referans_kodu, ' ', '')) = UPPER(REPLACE(?, ' ', '')) "
                "AND vardiya_id IN (SELECT id FROM vardiyalar WHERE COALESCE(lokasyon, 'TK2') = 'TK2' "
                "AND COALESCE(bolum, 'kaynak') = ?)", (ct, kod, bolum))
        if 'kalip_goz' in fark:               # panelle aynı: açık otomatik sayaç kayıtları yeni göz çarpanıyla
            son['kalip_acik_kayit'] += conn.execute(
                "UPDATE uretim_kayitlari SET paket_adedi=? WHERE sayac_otomatik=1 "
                "AND UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ','')) "
                "AND vardiya_id IN (SELECT id FROM vardiyalar WHERE COALESCE(lokasyon,'TK2')='TK2' "
                "AND COALESCE(bolum,'kaynak')='metal')", (fark['kalip_goz'], kod)).rowcount

    if excel_norm:
        for rid, rk in conn.execute(
                "SELECT id, referans_kodu FROM referans_listesi "
                "WHERE COALESCE(bolum,'kaynak')=? AND COALESCE(lokasyon,'TK2')='TK2'", (bolum,)).fetchall():
            n = _norm_kod(rk)
            if n and n not in excel_norm:
                if taban is not None and (n, bolum) not in taban:
                    # İndirildikten SONRA açılmış (operatör girişi / panel) — kullanıcı
                    # bu satırı hiç görmedi, silmesi söz konusu değil.
                    son['referanslar_korunan'] += 1
                    if len(son['ornek_korunan']) < ornek_n:
                        son['ornek_korunan'].append(rk)
                    continue
                conn.execute('DELETE FROM referans_listesi WHERE id = ?', (rid,))
                son['referanslar_silinen'] += 1
                if len(son['ornek_silinen']) < ornek_n:
                    son['ornek_silinen'].append(rk)
    else:
        son['uyari'] = "dosyada bu bölümün satırı yok — bu bölümde silme yapılmadı"
    print(f"  [ANA VERİ/{bolum}] {son['referanslar_eklenen']} eklendi, "
          f"{son['referanslar_guncellenen']} güncellendi, {son['referanslar_silinen']} silindi")
    return son


def _liste_farki(conn, sql, yeni):
    """DB kümesi ↔ dosya kümesi → (eklenecek, çıkacak) — sıralı listeler."""
    eski = {tuple(r) for r in conn.execute(sql).fetchall()}
    yeni = set(yeni)
    return sorted(yeni - eski), sorted(eski - yeni)


def ana_veri_uygula(satirlar, uygula=False, commit_oncesi=None, kume=None, taban=None):
    """Altı bölümün referansları + (kume verilirse) operatörler, robot program, fikstür —
    TEK işlemde. uygula=False → ÖNİZLEME (her şey geri alınır).
    commit_oncesi: commit'ten hemen önce çağrılır (dosyayı yerine koymak için);
    hata verirse veritabanı değişikliği geri alınır.
    Döner: {'bolumler': {bolum: referans özeti}, 'operatorler': {bolum: [eklenen ad]},
            'robot': {...}, 'fikstur': {...}}"""
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    try:
        sonuc = {'bolumler': {b: _ana_veri_bolum_uygula(conn, b, satirlar.get(b, []), taban=taban)
                              for b in ANA_VERI_ETIKET},
                 'operatorler': {}, 'robot': None, 'fikstur': None}
        if kume is not None:
            for b in ANA_VERI_ETIKET:
                ad = []
                _operator_import(conn, kume, b, eklenenler=ad)
                sonuc['operatorler'][b] = ad
            # Robot program / fikstür: Excel master (sil + yaz) — ama YALNIZ fark varsa;
            # aynıysa tablolara dokunulmaz (güncelleme tarihleri boşuna değişmesin).
            if ROBOT_PROGRAM_SAYFA in kume.sheetnames:
                k = _program_listesi_oku(kume[ROBOT_PROGRAM_SAYFA])
                if k is not None:
                    ek, cik = _liste_farki(conn, "SELECT robot_no, istasyon, referans_kodu FROM robot_programlari", k)
                    sonuc['robot'] = {'dosyada': len(k), 'eklenen': len(ek), 'cikan': len(cik),
                                      'ornek_eklenen': [' · '.join(map(str, x)) for x in ek[:10]],
                                      'ornek_cikan': [' · '.join(map(str, x)) for x in cik[:10]]}
                    if ek or cik:
                        _program_listesi_import(conn, kume)
            if FIKSTUR_RAF_SAYFA in kume.sheetnames:
                k = _fikstur_raf_oku(kume[FIKSTUR_RAF_SAYFA])
                if k is not None:
                    ek, cik = _liste_farki(conn, "SELECT referans_kodu, raf_no FROM fikstur_raf", k)
                    sonuc['fikstur'] = {'dosyada': len(k), 'eklenen': len(ek), 'cikan': len(cik),
                                        'ornek_eklenen': [' · '.join(x) for x in ek[:10]],
                                        'ornek_cikan': [' · '.join(x) for x in cik[:10]]}
                    if ek or cik:
                        _fikstur_raf_import(conn, kume)
        if uygula:
            if commit_oncesi:
                commit_oncesi()
            conn.commit()
        else:
            conn.rollback()
        return sonuc
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def ana_veri_operator_ekle(ad, bolum):
    """Panelden eklenen operatörü Ana Veri'nin '<Bölüm> Operator' sayfasına da yazar
    (eskiden yalnız veritabanına giriyordu, Excel ile liste ayrışıyordu). → hata | None"""
    if not ana_veri_aktif() or _ana_veri_yazma_engeli() or bolum not in BOLUM_SAYFA:
        return None
    try:
        with open(ANA_VERI_YOL, 'rb') as fh:
            wb = openpyxl.load_workbook(io.BytesIO(fh.read()))
    except Exception as e:
        return f'AnaVeri.xlsx açılamadı: {e}'
    sayfa = BOLUM_SAYFA[bolum]['op']
    if sayfa in wb.sheetnames:
        ws = wb[sayfa]
    else:
        ws = wb.create_sheet(sayfa)
        ws.append(['No', 'Operatör İsmi'])
    son_no, hedef = 0, _ad_normal(ad)
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0 or not row:
            continue
        if len(row) > 1 and row[1] is not None and _ad_normal(row[1]) == hedef:
            return None                                   # zaten listede
        try:
            son_no = max(son_no, int(row[0]))
        except (TypeError, ValueError):
            pass
    ws.append([son_no + 1, ad])
    return _ana_veri_kaydet(wb)


def ana_veri_dosya_kaydet(ham):
    """Yüklenen dosyayı data/AnaVeri.xlsx yapar; öncekini ana_veri_yedek/'e alır (son 20)."""
    from datetime import datetime as _dt
    import shutil
    os.makedirs(os.path.dirname(ANA_VERI_YOL), exist_ok=True)
    if os.path.exists(ANA_VERI_YOL):
        os.makedirs(ANA_VERI_YEDEK_KLASOR, exist_ok=True)
        shutil.copy2(ANA_VERI_YOL, os.path.join(
            ANA_VERI_YEDEK_KLASOR, 'AnaVeri-' + _dt.now().strftime('%Y%m%d-%H%M%S') + '.xlsx'))
        eskiler = sorted(f for f in os.listdir(ANA_VERI_YEDEK_KLASOR) if f.startswith('AnaVeri-'))
        for f in eskiler[:-20]:
            try:
                os.remove(os.path.join(ANA_VERI_YEDEK_KLASOR, f))
            except OSError:
                pass
    gecici = ANA_VERI_YOL + '.yukleniyor'
    with open(gecici, 'wb') as fh:
        fh.write(ham)
    os.replace(gecici, ANA_VERI_YOL)


def _ana_veri_yazma_engeli():
    """Ana Veri'ye yazılmaması gereken durum varsa sebebi (geliştirme kopyası / senkron kapalı)."""
    if os.path.exists(os.path.join(PROJECT_DIR, 'data', 'GELISTIRME_KOPYASI.json')):
        return 'geliştirme kopyası — Excel yalnız canlı sunucuda güncellenir'
    try:
        import kurulum as _kur
        if not _kur.modul('excel_senkron'):
            return 'Excel senkronu bu kurulumda kapalı'
    except Exception:
        pass
    return ''


def _ana_veri_yazmak_icin_ac():
    """(wb, ws, kol) — formüller korunsun diye data_only=False; bellekten okunur (dosya kilitlenmez)."""
    with open(ANA_VERI_YOL, 'rb') as fh:
        wb = openpyxl.load_workbook(io.BytesIO(fh.read()))
    ws, kol = _ana_sayfa_bul(wb)
    if ws is None:
        raise ValueError("AnaVeri.xlsx içinde 'Ana Veri' sayfası / başlıkları bulunamadı")
    return wb, ws, kol


def _ana_veri_kaydet(wb):
    try:
        wb.save(ANA_VERI_YOL)
        return None
    except PermissionError:
        return 'AnaVeri.xlsx şu an açık — kapatıp tekrar deneyin.'


def _sayi_hucre(v):
    """Tam sayı → int; ondalık TAM HASSASİYETLE (yuvarlanırsa sonraki yüklemede
    '439,0244 → 439,02' sahte değişiklik olarak görünür ve geçmiş cycle'lar oynar)."""
    v = float(v or 0)
    return int(v) if v == int(v) else v


def _ana_veri_export(conn, bolum_listesi, zorla_ekle=None):
    """DB → Ana Veri. Mevcut satırın süre/göz/büküm/açıklama/teyit hücreleri güncellenir;
    DB'de olup listede olmayan HER referans sona eklenir — SÜRESİZ OLANLAR DA (kullanıcı
    2026-10-05: liste veritabanının kendisi; operatörün üretim girişinde açtığı kod
    listede görünmezse bir sonraki yükleme onu "listede yok" diye silerdi — oysa aranan
    tam da bu tanımsız kodlar). Süresiz satır 'Not' sütununda işaretlenir.
    Süresi 0 olan referansın süre hücresine dokunulmaz (Excel'deki değer kalır).
    zorla_ekle: eski imza uyumu (artık herkes eklendiği için etkisiz)."""
    bolum_listesi = [b for b in bolum_listesi if b in ANA_VERI_ETIKET]
    try:
        wb, ws, kol = _ana_veri_yazmak_icin_ac()
    except Exception as e:
        return {'basarili': False, 'hata': f'AnaVeri.xlsx açılamadı: {e}'}
    son_kol = ws.max_column
    for alan in ('cevrim', 'kaynak', 'goz', 'bukum', 'aciklama', 'teyit'):
        if alan not in kol:                   # kullanıcı sütunu silmişse sona yeniden açılır
            son_kol += 1
            ws.cell(row=1, column=son_kol, value=_ANA_YENI_BASLIK[alan])
            kol[alan] = son_kol - 1
    harita = {}
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        kod = row[kol['kod']] if kol['kod'] < len(row) else None
        b = ana_veri_bolum(row[kol['bolum']] if kol['bolum'] < len(row) else None)
        if kod is None or not b:
            continue
        harita.setdefault((_norm_kod(kod), b), (i, str(kod).strip()))

    def yaz(ri, alan, deger):
        ws.cell(row=ri, column=kol[alan] + 1, value=deger)

    toplam = 0
    for b in bolum_listesi:
        rows = conn.execute(
            "SELECT referans_kodu, COALESCE(hedef_cycle_time_sn,0) ct, COALESCE(kaynak_suresi_sn,0) ks, "
            "COALESCE(soktak_suresi_sn,0) ss, COALESCE(sure_teyit,0) teyit, COALESCE(aciklama,'') aciklama, "
            "COALESCE(bukum_operasyon,1) bukum, COALESCE(kalip_goz,1) goz FROM referans_listesi "
            "WHERE COALESCE(bolum,'kaynak')=? AND COALESCE(lokasyon,'TK2')='TK2' ORDER BY referans_kodu",
            (b,)).fetchall()
        yazilan_norm = {}
        for r in rows:
            kod = (r['referans_kodu'] or '').strip()
            norm = _norm_kod(kod)
            if not norm:
                continue
            if (norm, b) in harita:
                ri, sayfa_kod = harita[(norm, b)]
                tam_es = (kod == sayfa_kod)
            else:
                ri = ws.max_row + 1
                yaz(ri, 'kod', kod)
                yaz(ri, 'bolum', ANA_VERI_ETIKET[b])
                if 'ap' in kol:
                    yaz(ri, 'ap', 'P')
                if 'tk' in kol:
                    yaz(ri, 'tk', 'TK-2')
                if 'not' in kol:
                    yaz(ri, 'not', "Forge'da açıldı — süre bekliyor" if r['ct'] <= 0 else "Forge'da açıldı")
                harita[(norm, b)] = (ri, kod)
                tam_es = True
            if yazilan_norm.get(norm) and not tam_es:
                continue                       # yazım varyantı tam-eşin değerini ezmesin
            if r['ct'] > 0:
                if b == 'kaynak':
                    yaz(ri, 'kaynak', _sayi_hucre(r['ks']))
                    yaz(ri, 'cevrim', _sayi_hucre(r['ss']))
                else:
                    yaz(ri, 'cevrim', _sayi_hucre(r['ct']))
            if b == 'kaynak':
                yaz(ri, 'teyit', 'EVET' if r['teyit'] else None)
            if b == 'metal':
                yaz(ri, 'goz', int(r['goz'] or 1))
            if b == 'pres':
                yaz(ri, 'bukum', int(r['bukum'] or 1))
            if (r['aciklama'] or '').strip():
                yaz(ri, 'aciklama', r['aciklama'].strip())
            yazilan_norm[norm] = True
            toplam += 1
    hata = _ana_veri_kaydet(wb)
    if hata:
        return {'basarili': False, 'hata': hata}
    return {'basarili': True, 'yazilan': toplam, 'dosya': ANA_VERI_YOL}


def ana_veri_satir_sil(kod, bolum=None):
    """Panelden silinen referansın satırını Ana Veri'den de kaldırır (yoksa sonraki
    yüklemede geri gelirdi). bolum=None → koddaki tüm satırlar. → hata metni | None"""
    if not ana_veri_aktif() or _ana_veri_yazma_engeli():
        return None
    try:
        wb, ws, kol = _ana_veri_yazmak_icin_ac()
    except Exception as e:
        return f'AnaVeri.xlsx açılamadı: {e}'
    hedef = _norm_kod(kod)
    silinecek = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        k = row[kol['kod']] if kol['kod'] < len(row) else None
        if k is None or _norm_kod(k) != hedef:
            continue
        if bolum and ana_veri_bolum(row[kol['bolum']] if kol['bolum'] < len(row) else None) != bolum:
            continue
        silinecek.append(i)
    if not silinecek:
        return None
    for i in reversed(silinecek):
        ws.delete_rows(i)
    return _ana_veri_kaydet(wb)


def ana_veri_kod_degistir(degistir):
    """Kod yeniden adlandırması (app._guvenli_kod_replace) Ana Veri'nin kod sütununa da
    uygulanır; yoksa sonraki yüklemede eski kod geri gelir, yenisi silinirdi.
    Yeni kod aynı bölümde zaten satır olarak varsa adı değişen satır kaldırılır
    (veritabanındaki birleştirme kuralıyla aynı). → hata metni | None"""
    if not ana_veri_aktif() or _ana_veri_yazma_engeli():
        return None
    try:
        wb, ws, kol = _ana_veri_yazmak_icin_ac()
    except Exception as e:
        return f'AnaVeri.xlsx açılamadı: {e}'
    mevcut, degisen = set(), []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        k = row[kol['kod']] if kol['kod'] < len(row) else None
        if k is None:
            continue
        b = ana_veri_bolum(row[kol['bolum']] if kol['bolum'] < len(row) else None)
        yeni = degistir(str(k))
        if yeni != str(k):
            degisen.append((i, yeni, b))
        else:
            mevcut.add((_norm_kod(k), b))
    sil = []
    for i, yeni, b in degisen:
        if (_norm_kod(yeni), b) in mevcut:
            sil.append(i)
        else:
            ws.cell(row=i, column=kol['kod'] + 1, value=yeni)
            mevcut.add((_norm_kod(yeni), b))
    for i in reversed(sil):
        ws.delete_rows(i)
    # Robot Program / Fikstür sayfaları da: rename robot_programlari ve fikstur_raf
    # tablolarını da değiştiriyor; sayfa eski kodda kalırsa sonraki yükleme geri alırdı
    # (2026-10-05'te eski Excel'de tam bu olmuştu: 10.300.4199 ↔ DB'de 10.300.4199W).
    ek = 0
    for sayfa, ilk_satir, kolonlar in ((ROBOT_PROGRAM_SAYFA, 3, None), (FIKSTUR_RAF_SAYFA, 2, 3)):
        if sayfa not in wb.sheetnames:
            continue
        s = wb[sayfa]
        for row in s.iter_rows(min_row=ilk_satir):
            for hucre in (row if kolonlar else row[:1]):
                if kolonlar and (hucre.column - 1) % kolonlar:
                    continue                   # fikstür: her 3 sütunun ilki kod
                if isinstance(hucre.value, str):
                    yeni = degistir(hucre.value)
                    if yeni != hucre.value:
                        hucre.value = yeni
                        ek += 1
    if not degisen and not ek:
        return None
    return _ana_veri_kaydet(wb)


def _bolum_import(conn, wb, bolum):
    """Tek bölümün referans + operatör verisini import eder."""
    sayfalar = BOLUM_SAYFA[bolum]
    c = conn.cursor()

    # ── Referans sayfası ──
    if sayfalar['ref'] not in wb.sheetnames:
        return {
            'referanslar_eklenen': 0,
            'referanslar_guncellenen': 0,
            'referanslar_silinen': 0,
            'operatorler_eklenen': 0,
            'hata': f"'{sayfalar['ref']}' sayfası bulunamadı"
        }
    ref_sayfa = wb[sayfalar['ref']]
    print(f"  [{bolum.upper()}] Referans sayfası: '{ref_sayfa.title}'")

    ref_sayisi = 0
    ref_guncellenen = 0
    tel_kodu_atlanan = 0        # TK2'ye yazılmayan 93.* (TK1 tel) kodları
    excel_kodlari_norm = set()

    # Kaynak bölümü için Excel artık 4 kolon: Kod | Kaynak Süresi | Söktak Süresi | Toplam Cycle
    # Pres Abkant 3 kolon: Kod | Açıklama | Süre (operatör mobilde açıklamayı görür)
    # Montaj/Metal/diğerleri eski 2 kolon: Kod | Cycle Time (tek değer)
    kaynak_modu = (bolum == 'kaynak')
    pres_modu = (bolum == 'pres')

    for i, row in enumerate(ref_sayfa.iter_rows(values_only=True)):
        if i == 0:
            continue  # Başlık
        if not row or row[0] is None:
            continue
        kod = str(row[0]).strip()
        if not kod or len(kod) < 2:
            continue

        # Süreleri parse et
        kaynak_sn = 0.0
        soktak_sn = 0.0
        cycle = 0.0
        aciklama = ''
        bukum_op = 1   # yalnız pres modunda D kolonundan okunur; diğerlerinde 1 = bölme yok
        if kaynak_modu:
            # B = Kaynak Süresi, C = Söktak Süresi, D = Toplam (formül)
            try:
                kaynak_sn = float(row[1]) if (len(row) > 1 and row[1] is not None) else 0.0
            except (ValueError, TypeError):
                kaynak_sn = 0.0
            try:
                soktak_sn = float(row[2]) if (len(row) > 2 and row[2] is not None) else 0.0
            except (ValueError, TypeError):
                soktak_sn = 0.0
            cycle = round(kaynak_sn + soktak_sn, 2)
        elif pres_modu:
            # B = Açıklama (metin), C = Süre, D = Büküm Op. (2026-07-29)
            aciklama = str(row[1]).strip() if (len(row) > 1 and row[1] is not None) else ''
            try:
                cycle = float(row[2]) if (len(row) > 2 and row[2] is not None) else 0.0
            except (ValueError, TypeError):
                cycle = 0.0
            # Büküm operasyon sayısı: bir parçaya sırayla uygulanan büküm adedi.
            # Sayaç bölücüsü olarak kullanılır (3 op → 3. sinyal = 1 parça).
            # Boş/bozuk = 1 (bölme yok). D kolonu eski dosyalarda hiç yok → len kontrolü şart.
            try:
                bukum_op = int(row[3]) if (len(row) > 3 and row[3] is not None) else 1
            except (ValueError, TypeError):
                bukum_op = 1
            bukum_op = max(1, min(99, bukum_op))
        else:
            # Montaj/Metal/diğerleri — tek değer
            try:
                cycle = float(row[1]) if (len(row) > 1 and row[1] is not None) else 0.0
            except (ValueError, TypeError):
                cycle = 0.0

        # ── TK1 TEL KODU TK2'YE YAZILMAZ (kullanıcı 2026-08-06) ──────────────
        # Kural: "93. ile başlayan bütün referanslar tel referansıdır" → tel
        # üretimi TK1'de yapılır, bu kodların TK2 listesinde işi yok.
        # OLAY: TK2 Excel'inin 'Montaj Referans' sayfasında 1460 adet 93.* kod
        # duruyordu; hepsi süresiz olduğu için TK2 montajın "Süre Tanımı Bekleyen
        # Referanslar" panelini şişiriyor (2384 satırın 1459'u bunlardı) ve gerçek
        # eksikler arasında kayboluyordu.
        # Koda ATLAMA yeterli: bu kodlar excel_kodlari_norm'a girmediği için
        # aşağıdaki mirror-sync onları DB'den de temizler — Excel'i elle
        # düzeltmeye gerek kalmaz, sonraki içe aktarımlarda geri gelmezler.
        # Ölçüm (2026-08-06): TK2'de 93.* kodların HİÇBİRİNİN süresi tanımlı
        # değildi, yani silinen bir bilgi yok.
        if kod.strip().startswith('93.'):
            tel_kodu_atlanan += 1
            continue

        excel_kodlari_norm.add(kod.upper().replace(' ', ''))

        # LOKASYON GÜVENLİĞİ: _bolum_import YALNIZCA TK2 Excel'ini (EXCEL_YOL) işler.
        # Tüm referans_listesi okuma/yazma işlemleri lokasyon='TK2' ile kapatılır ki
        # TK2 import'u TK1 (yan tesis) kayıtlarını eşleştirip ezmesin/silmesin.
        # BÖLÜM KAPSAMASI: UNIQUE(referans_kodu, bolum, lokasyon) sonrası aynı kod birden
        # fazla bölümde farklı süreyle yaşayabilir (16 kod metal+işleme'de ortak) —
        # eşleştirme ve cycle geri-yayılımı bu bölümün satırı/vardiyalarıyla sınırlı,
        # yoksa bölümler birbirinin kaydını çalar/cycle'ını ezer.
        # TAM-EŞ TERCİHİ: DB'de aynı koda normalize olan birden fazla yazım varyantı
        # olabilir (eski auto-create kalıntısı: '94.LTK.10' + '94.ltk.10'). Excel'deki
        # yazımla birebir eşleşen satır varsa ONU güncelle — yoksa varyantı Excel
        # yazımına çevirirken tam-eş satırla UNIQUE çakışması oluşur.
        mevcut = c.execute(
            "SELECT id, referans_kodu FROM referans_listesi "
            "WHERE UPPER(REPLACE(referans_kodu, ' ', '')) = UPPER(REPLACE(?, ' ', '')) "
            "AND COALESCE(bolum, 'kaynak') = ? AND COALESCE(lokasyon, 'TK2') = 'TK2' "
            "ORDER BY (referans_kodu = ?) DESC, id LIMIT 1",
            (kod, bolum, kod)
        ).fetchone()

        if mevcut:
            if pres_modu:
                # Pres'te açıklama VE büküm operasyon sayısı Excel'den yönetilir —
                # import her seferinde günceller (operatörün mobilden girdiği değer de
                # zaten Excel'e yazılıyor, iki yön aynı hücreye bakar).
                c.execute(
                    'UPDATE referans_listesi SET hedef_cycle_time_sn = ?, kaynak_suresi_sn = ?, soktak_suresi_sn = ?, referans_kodu = ?, aciklama = ?, bukum_operasyon = ? WHERE id = ?',
                    (cycle, kaynak_sn, soktak_sn, kod, aciklama, bukum_op, mevcut[0])
                )
            else:
                # Diğer bölümlerde aciklama'ya DOKUNMA (dashboard'dan girilmiş olabilir)
                c.execute(
                    'UPDATE referans_listesi SET hedef_cycle_time_sn = ?, kaynak_suresi_sn = ?, soktak_suresi_sn = ?, referans_kodu = ? WHERE id = ?',
                    (cycle, kaynak_sn, soktak_sn, kod, mevcut[0])
                )
            if cycle > 0:
                c.execute(
                    "UPDATE uretim_kayitlari SET cycle_time_sn = ? "
                    "WHERE UPPER(REPLACE(referans_kodu, ' ', '')) = UPPER(REPLACE(?, ' ', '')) "
                    "AND vardiya_id IN (SELECT id FROM vardiyalar WHERE COALESCE(lokasyon, 'TK2') = 'TK2' AND COALESCE(bolum, 'kaynak') = ?)",
                    (cycle, kod, bolum)
                )
            ref_guncellenen += 1
        else:
            c.execute(
                "INSERT INTO referans_listesi (referans_kodu, hedef_cycle_time_sn, kaynak_suresi_sn, soktak_suresi_sn, aciklama, bolum, lokasyon, bukum_operasyon) VALUES (?, ?, ?, ?, ?, ?, 'TK2', ?)",
                (kod, cycle, kaynak_sn, soktak_sn, aciklama, bolum, bukum_op)
            )
            if cycle > 0:
                c.execute(
                    "UPDATE uretim_kayitlari SET cycle_time_sn = ? "
                    "WHERE UPPER(REPLACE(referans_kodu, ' ', '')) = UPPER(REPLACE(?, ' ', '')) "
                    "AND vardiya_id IN (SELECT id FROM vardiyalar WHERE COALESCE(lokasyon, 'TK2') = 'TK2' AND COALESCE(bolum, 'kaynak') = ?)",
                    (cycle, kod, bolum)
                )
            ref_sayisi += 1

    print(f"  Referanslar: {ref_sayisi} eklendi, {ref_guncellenen} güncellendi"
          + (f", {tel_kodu_atlanan} adet '93.*' TK1 tel kodu atlandı" if tel_kodu_atlanan else ""))

    # ── MIRROR SYNC: Excel'de olmayan referansları bu bölümden temizle ──
    ref_silinen = 0
    if excel_kodlari_norm:
        bolum_refs = c.execute(
            "SELECT id, referans_kodu FROM referans_listesi "
            "WHERE COALESCE(bolum, 'kaynak') = ? AND COALESCE(lokasyon, 'TK2') = 'TK2'",
            (bolum,)
        ).fetchall()
        for ref_row in bolum_refs:
            ref_norm = str(ref_row[1] or '').upper().replace(' ', '')
            if ref_norm and ref_norm not in excel_kodlari_norm:
                c.execute('DELETE FROM referans_listesi WHERE id = ?', (ref_row[0],))
                ref_silinen += 1
        if ref_silinen:
            print(f"  Referanslar: {ref_silinen} adet (Excel'de olmayan) silindi")
    else:
        print("  UYARI: Excel'den hiçbir referans okunamadı, silme atlandı")

    return {
        'referanslar_eklenen': ref_sayisi,
        'referanslar_guncellenen': ref_guncellenen,
        'referanslar_silinen': ref_silinen,
        'operatorler_eklenen': _operator_import(conn, wb, bolum)
    }


def _operator_import(conn, wb, bolum, eklenenler=None):
    """'<Bölüm> Operator' sayfasındaki yeni kişileri ekler (kimseyi silmez) → eklenen sayısı.
    wb: kitap ya da _SayfaKumesi (Ana Veri + eski dosya). eklenenler: liste verilirse
    eklenen adlar oraya yazılır (Ana Veri önizlemesi için)."""
    sayfalar = BOLUM_SAYFA[bolum]
    c = conn.cursor()
    op_sayisi = 0
    if sayfalar['op'] in wb.sheetnames:
        op_sayfa = wb[sayfalar['op']]
        print(f"  [{bolum.upper()}] Operatör sayfası: '{op_sayfa.title}'")
        # TÜRKÇE HARFE DUYARSIZ EŞLEŞME (2026-10-05): SQLite UPPER() 'i'yi 'I' yapar ama
        # 'İ'ye dokunmaz → 'İbrahim Nak' ile 'İBRAHİM NAK' farklı sayılıp AYNI KİŞİ İKİ
        # KEZ eklenmişti (sunucuda montajda iki kayıt). Panel ekleme (_ad_esitle) ile aynı kural.
        tk2_kisiler = {}                       # _ad_normal(ad) → [(bolum, pin)]
        for ad_db, bolum_db, pin_db in c.execute(
                "SELECT ad, bolum, pin FROM operatorler WHERE COALESCE(lokasyon,'TK2')='TK2'").fetchall():
            tk2_kisiler.setdefault(_ad_normal(ad_db), []).append((bolum_db, pin_db))
        for i, row in enumerate(op_sayfa.iter_rows(values_only=True)):
            if i == 0:
                continue
            if not row or len(row) < 2 or row[1] is None:
                continue
            ad = str(row[1]).strip()
            if not ad:
                continue
            try:
                # LOKASYON: TK2 import'u yalnız TK2 operatörlerine bakar/yazar —
                # TK1'de aynı ad varsa TK2 operatörü sessizce atlanmasın (ayrı fabrika).
                # ÇOKLU BÖLÜM: aynı kişi birden fazla bölümde çalışabilir → bölüm başına
                # satır (UNIQUE(ad, bolum, lokasyon)). Kişi başka bölümde zaten varsa
                # yeni bölüm satırı ONUN PIN'iyle açılır (bir kişi = tek PIN).
                kisi = tk2_kisiler.get(_ad_normal(ad), [])
                if not any(b == bolum for b, _p in kisi):
                    pin = (kisi[0][1] or '0000') if kisi else '0000'
                    c.execute("INSERT INTO operatorler (ad, bolum, pin, lokasyon) VALUES (?, ?, ?, 'TK2')", (ad, bolum, pin))
                    tk2_kisiler.setdefault(_ad_normal(ad), []).append((bolum, pin))
                    op_sayisi += 1
                    if eklenenler is not None:
                        eklenenler.append(ad)
            except Exception as e:
                print(f"  Operatör eklenemedi ({ad}): {e}")
        print(f"  Operatörler: {op_sayisi} eklendi")
    return op_sayisi


def import_tk1(conn=None):
    """TK1 (yan tesis) referans + operatör verisini import eder — lokasyon='TK1', bolum='montaj'.

    _bolum_import'tan AYRI (mirror-sync YOK → TK2 verisini silmez). referans_kodu global
    UNIQUE olduğu için INSERT OR IGNORE (TK2 ile çakışan ~47 kod atlanır; kod yine elle
    girilebilir). TK1 'Operatörler' sayfasında BAŞLIK SATIRI YOK → row 0 dahil. PIN default
    '0000' (panelden değiştirilir). Idempotent — tekrar çalıştırılabilir.
    """
    kapat = False
    if conn is None:
        conn = sqlite3.connect(DB_PATH, timeout=20.0)
        kapat = True
    c = conn.cursor()
    if not os.path.exists(TK1_EXCEL_YOL):
        return {'basarili': False, 'hata': f"'{TK1_EXCEL_YOL}' bulunamadı"}
    wb = openpyxl.load_workbook(TK1_EXCEL_YOL, data_only=True)

    def _sayfa_aktar(sayfa_adi, bolum, tur):
        """Tek kolonlu TK1 sayfasını aktarır. tur: 'ref' | 'op'. Döner: eklenen satır.

        Başlık satırı SABİT İNDEKSLE DEĞİL İÇERİKLE atlanır (bkz. _tk1_baslik_mi):
        montaj referans sayfasında başlık var, operatör sayfasında yok; plastik
        sayfalarında hangisinin olacağı belli değil ve sabit kural ya ilk kaydı
        yutar ya da başlığı referans/operatör olarak veritabanına yazar."""
        if sayfa_adi not in wb.sheetnames:
            return 0
        eklenen = 0
        ilk_veri = True
        for row in wb[sayfa_adi].iter_rows(values_only=True):
            if not row or row[0] is None:
                continue
            deger = str(row[0]).strip()
            if not deger:
                continue
            # Başlık YALNIZ ilk dolu satırda aranır. Kelime listesini her satıra
            # uygulamak MEŞRU kodları elerdi: TK1 listesinde 'TEST', 'ÖZEL ÜRÜN',
            # 'HALAT TAHRIBAT' gibi rakamsız ama GERÇEK referanslar var (aynı
            # sebeple "kod rakam içermeli" kuralı da kullanılamaz).
            if ilk_veri:
                ilk_veri = False
                if _tk1_baslik_mi(deger):
                    continue
            if tur == 'ref':
                if len(deger) < 2:
                    continue
                # 93.* KODLARI DOĞRUDAN TELE YAZILIR (kullanıcı 2026-08-04).
                # Ana sayfa montaj + tel kodlarını birlikte tutuyor. Önce montaja
                # yazıp sonra UPDATE ile taşımak İKİ KAYIT üretiyordu (tel'e taşınan
                # eski satır + ana sayfadan tekrar eklenen montaj satırı) ve ikinci
                # içe aktarımda UNIQUE(kod,bolum,lokasyon) hatası veriyordu.
                _ref_bolum = ('tel' if (bolum == 'montaj' and deger.startswith('93.'))
                              else bolum)
                cur = c.execute(
                    "INSERT OR IGNORE INTO referans_listesi "
                    "(referans_kodu, hedef_cycle_time_sn, bolum, lokasyon) VALUES (?, 0, ?, 'TK1')",
                    (deger, _ref_bolum))
            else:
                # Ana operatör sayfasında bölüm İSME göre belirlenir (montaj = 2 kişi,
                # kalan herkes tel). Plastik/tel kendi sayfalarından gelirse bolum sabit.
                _op_bolum = tk1_operator_bolumu(deger) if bolum == 'montaj' else bolum
                cur = c.execute(
                    "INSERT OR IGNORE INTO operatorler (ad, pin, bolum, lokasyon) "
                    "VALUES (?, '0000', ?, 'TK1')",
                    (deger, _op_bolum))
            eklenen += cur.rowcount
        return eklenen

    # ── ÖNCE MEVCUT VERİYİ DÜZELT, SONRA EKLE ────────────────────────────────
    # Sıra kritik: yeni kayıtlar isme/koda göre doğru bölümle ekleniyor. Eğer önce
    # eklersek, aynı kişi/kod hem eski 'montaj' satırı hem yeni 'tel' satırı olarak
    # bulunur ve taşıma UNIQUE(ad,bolum,lokasyon) kısıtına takılır.
    #
    # 1) Kod kuralı — 93.* → tel. İDEMPOTENT ve GERİYE DÖNÜK. Tersi YAPILMAZ
    #    (tel sayfasından gelen 93.* olmayan kod tel kalır — orada bilinçli tanım var).
    # Satır satır: aynı kod tel'de zaten varsa (eski veriden kalma çift kayıt)
    # toplu UPDATE tüm işlemi patlatırdı — o satır silinir, tel kaydı korunur.
    tel_tasinan = 0
    for _rr in c.execute(
            "SELECT id, referans_kodu FROM referans_listesi "
            "WHERE COALESCE(lokasyon,'TK2')='TK1' AND COALESCE(bolum,'')='montaj' "
            "AND TRIM(referans_kodu) LIKE '93.%'").fetchall():
        try:
            c.execute("UPDATE referans_listesi SET bolum='tel' WHERE id=?", (_rr[0],))
        except sqlite3.IntegrityError:
            c.execute("DELETE FROM referans_listesi WHERE id=?", (_rr[0],))
        tel_tasinan += 1

    # 2) Operatör dağılımı — montaj listesinde OLMAYAN herkes tele taşınır.
    #    'Admin' (her bölümde görünen özel kullanıcı) dışarıda bırakılır.
    #    Çakışma olursa (kişi zaten tel'de kayıtlı) eski montaj satırı silinir ama
    #    PIN'i korunur — operatör kendi PIN'iyle girmeye devam etsin.
    op_tel_tasinan = 0
    for _r in c.execute(
            "SELECT id, ad, COALESCE(pin,'0000') FROM operatorler "
            "WHERE COALESCE(lokasyon,'TK2')='TK1' AND COALESCE(bolum,'')='montaj' "
            "AND ad != 'Admin'").fetchall():
        if tk1_operator_bolumu(_r[1]) != 'tel':
            continue
        try:
            c.execute("UPDATE operatorler SET bolum='tel' WHERE id=?", (_r[0],))
        except sqlite3.IntegrityError:
            # Aynı kişi tel'de zaten var → PIN'i oraya taşı, montaj satırını sil
            c.execute("UPDATE operatorler SET pin=? WHERE ad=? AND bolum='tel' "
                      "AND COALESCE(lokasyon,'TK2')='TK1' AND COALESCE(pin,'0000')='0000'",
                      (_r[2], _r[1]))
            c.execute("DELETE FROM operatorler WHERE id=?", (_r[0],))
        op_tel_tasinan += 1

    # ── Montaj (mevcut akış) ──
    # TEL KURALI (kullanıcı 2026-08-04): TK1'de "93." ile başlayan HER referans
    # tel referansıdır. Ana referans sayfası montaj + tel kodlarını birlikte
    # tutuyor; bölüm koda bakılarak ayrılır (ayrı sayfa tutmaya gerek yok).
    ref_eklenen = _sayfa_aktar(TK1_SAYFA['ref'], 'montaj', 'ref')
    op_eklenen  = _sayfa_aktar(TK1_SAYFA['op'],  'montaj', 'op')

    # ── Plastik enjeksiyon (2026-08-04) — TK1'de AYRI bölüm ──
    # Operatör ve referansları montajdan bağımsız: plastik operatörü montaj
    # referanslarını görmemeli, mobilde bölüm seçilince kendi listesi gelmeli.
    p_ref_eklenen = _sayfa_aktar(TK1_SAYFA['plastik_ref'], 'plastik', 'ref')
    p_op_eklenen  = _sayfa_aktar(TK1_SAYFA['plastik_op'],  'plastik', 'op')

    # ── Tel üretimi (2026-08-04) — OPSİYONEL sayfa ──
    # Tel referansları ARTIK BU SAYFADAN GELMİYOR: TK1'de '93.' ile başlayan her
    # kod tel referansıdır (yukarıdaki kural) ve operatörler ana sayfadan isimle
    # ayrışır. Bu blok yalnız GERİYE UYUM için duruyor — sayfa varsa okunur,
    # yoksa hiçbir şey olmaz (normal durum budur).
    # tel_adimlar (proses tanımı) artık KULLANILMIYOR: kapaması/son montajı
    # dışarıda yapılacak ürünler sabit olmadığı için referans bazlı akış tanımı
    # terk edildi; her adım kendi referans ekiyle kaydediliyor (bkz. tel_proses.py).
    t_ref_eklenen = t_adim_guncel = 0
    if TK1_SAYFA['tel_ref'] in wb.sheetnames:
        ws = wb[TK1_SAYFA['tel_ref']]
        sutun_adim = {}          # kolon indeksi → adım adı
        baslik_okundu = False
        for row in ws.iter_rows(values_only=True):
            if not row or row[0] is None or not str(row[0]).strip():
                continue
            ilk = str(row[0]).strip()
            if not baslik_okundu:
                baslik_okundu = True
                if _tk1_baslik_mi(ilk):
                    # Başlık satırı: adım kolonlarını buradan öğren
                    for i, h in enumerate(row[1:], start=1):
                        ad = TEL_ADIM_BASLIK.get(str(h or '').strip().lower())
                        if ad:
                            sutun_adim[i] = ad
                    continue      # başlık satırı veri değil
            kod = ilk
            if len(kod) < 2:
                continue
            # İşaretli adımları ÜRETİM SIRASINDA topla (kolon sırası değil —
            # Excel'de kolonlar karışık dizilse bile sıra doğru kalsın).
            secili = [sutun_adim[i] for i in sorted(sutun_adim)
                      if i < len(row) and _tel_isaretli_mi(row[i])]
            sirali = [a for a in TEL_ADIM_SIRASI if a in secili]
            adim_metni = ','.join(sirali)
            cur = c.execute(
                "INSERT OR IGNORE INTO referans_listesi "
                "(referans_kodu, hedef_cycle_time_sn, bolum, lokasyon, tel_adimlar) "
                "VALUES (?, 0, 'tel', 'TK1', ?)", (kod, adim_metni))
            t_ref_eklenen += cur.rowcount
            if adim_metni:
                cur2 = c.execute(
                    "UPDATE referans_listesi SET tel_adimlar=? "
                    "WHERE referans_kodu=? AND COALESCE(bolum,'')='tel' "
                    "AND COALESCE(lokasyon,'TK2')='TK1' AND COALESCE(tel_adimlar,'')!=?",
                    (adim_metni, kod, adim_metni))
                t_adim_guncel += cur2.rowcount
    t_op_eklenen = _sayfa_aktar(TK1_SAYFA['tel_op'], 'tel', 'op')

    conn.commit()
    if kapat:
        conn.close()
    print(f"[TK1] montaj: {ref_eklenen} referans + {op_eklenen} operatör | "
          f"plastik: {p_ref_eklenen} referans + {p_op_eklenen} operatör | "
          f"tel: {t_ref_eklenen} referans + {t_op_eklenen} operatör "
          f"({t_adim_guncel} proses tanımı güncellendi, {tel_tasinan} kod '93.*' kuralıyla, "
          f"{op_tel_tasinan} operatör tele taşındı) (lokasyon=TK1)")
    ref_eklenen += p_ref_eklenen + t_ref_eklenen
    op_eklenen  += p_op_eklenen + t_op_eklenen
    # NOT: referans_kodu GLOBAL UNIQUE → TK2 ile çakışan TK1 kodları INSERT OR IGNORE ile
    # atlanır (atlanan = bu kodlar TK2'de var). Tam ayrışma için UNIQUE(referans_kodu,lokasyon)
    # migration gerekir (bkz. lokasyon denetimi). 'referanslar_guncellenen' UI mesajı için 0.
    # ── TK1 PLASTİK DEPO VARSAYILANI 01/01 (kullanıcı 2026-08-25) ─────────
    # database.py migration'ı da aynı işi yapar ama O YALNIZ AÇILIŞTA çalışır;
    # aktarımla YENİ gelen referanslar bir sonraki yeniden başlatmaya kadar boş
    # kalırdı ve o aralıkta AS400 teyidi verilemezdi. Boş olanlara dokunur,
    # DOLU kodu ASLA ezmez.
    try:
        c.execute("UPDATE referans_listesi SET depo_kodu='01' "
                  "WHERE COALESCE(lokasyon,'TK2')='TK1' AND COALESCE(bolum,'kaynak')='plastik' "
                  "  AND COALESCE(depo_kodu,'')=''")
        c.execute("UPDATE referans_listesi SET karsi_depo_kodu='01' "
                  "WHERE COALESCE(lokasyon,'TK2')='TK1' AND COALESCE(bolum,'kaynak')='plastik' "
                  "  AND COALESCE(karsi_depo_kodu,'')=''")
        conn.commit()
    except Exception as _e:
        print(f'[TK1] plastik depo varsayilani atlandi: {_e}')

    # ── OKUNAN DOSYANIN KİMLİĞİ (kullanıcı 2026-08-25) ─────────────────────
    # "Laptopumdaki Excel'de değişiklik yapmıştım, bundan dolayı olabilir mi?"
    # EVET: aktarım UYGULAMANIN YANINDAKİ data/ klasörünü okur. Uygulama sunucuda
    # koştuğu için okunan dosya SUNUCUDAKİ kopyadır; laptoptaki düzenleme oraya
    # kendiliğinden GİTMEZ (data/ gitignore'da ve Guncelle.bat sunucudaki data/'yı
    # bilerek korur). Hangi dosyanın okunduğu ve NE ZAMAN değiştiği artık yanıtta
    # dönüyor — "aktardım ama değişmedi" bir daha teşhis gerektirmesin.
    _dosya_bilgi = {'yol': TK1_EXCEL_YOL}
    try:
        import datetime as _dt
        _st = os.stat(TK1_EXCEL_YOL)
        _dosya_bilgi['degisme'] = _dt.datetime.fromtimestamp(_st.st_mtime).strftime('%Y-%m-%d %H:%M')
        _dosya_bilgi['boyut'] = _st.st_size
    except OSError:
        pass
    return {'basarili': True, 'referanslar_eklenen': ref_eklenen,
            'referanslar_guncellenen': 0, 'referanslar_silinen': 0,
            'operatorler_eklenen': op_eklenen,
            'kaynak_dosya': _dosya_bilgi}


def import_data(bolum=None):
    """Excel'den verileri import eder.

    bolum=None  → tüm bölümler (BOLUM_SAYFA anahtarları)
    bolum='kaynak' / 'montaj' / 'metal' / 'isleme' / 'lazer' / 'pres' → sadece o bölüm
    """
    if bolum and bolum not in BOLUM_SAYFA:
        return {'basarili': False, 'hata': f"Geçersiz bölüm: {bolum}"}

    if not os.path.exists(EXCEL_YOL) and not ana_veri_aktif():
        return {'basarili': False, 'hata': f'Excel dosyası bulunamadı: {EXCEL_YOL}'}

    # ANA VERİ varsa referanslar VE operatörler oradan (operatör sayfası Ana Veri'de
    # yoksa eski uretim_verileri.xlsx'ten — _SayfaKumesi)
    ana_satirlar, ana_rapor, wb_ana = None, None, None
    if ana_veri_aktif():
        try:
            with open(ANA_VERI_YOL, 'rb') as fh:
                ham = fh.read()
            wb_ana = openpyxl.load_workbook(io.BytesIO(ham), data_only=True)
            ana_satirlar, ana_rapor = ana_veri_oku(wb_ana)
            _yedegi_tazele(ANA_VERI_YOL, ham)
        except Exception as e:
            return {'basarili': False, 'hata': f'AnaVeri.xlsx okunamadı: {e}'}

    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    c = conn.cursor()

    c.execute('''
        CREATE TABLE IF NOT EXISTS operatorler (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ad TEXT UNIQUE NOT NULL,
            bolum TEXT DEFAULT 'kaynak'
        )
    ''')
    try:
        c.execute("ALTER TABLE operatorler ADD COLUMN bolum TEXT DEFAULT 'kaynak'")
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE referans_listesi ADD COLUMN bolum TEXT DEFAULT 'kaynak'")
    except Exception:
        pass

    wb = openpyxl.load_workbook(EXCEL_YOL, data_only=True) if os.path.exists(EXCEL_YOL) else None
    if wb_ana is not None:
        wb = _SayfaKumesi(wb_ana, wb)

    sonuclar = {}
    toplam = {'referanslar_eklenen': 0, 'referanslar_guncellenen': 0,
              'referanslar_silinen': 0, 'operatorler_eklenen': 0}

    islenen = [bolum] if bolum else list(BOLUM_SAYFA.keys())

    for b in islenen:
        print(f"\n{'='*50}")
        print(f"  {b.upper()} import başlıyor...")
        print(f"{'='*50}")
        if ana_satirlar is not None:
            sonuc = _ana_veri_bolum_uygula(conn, b, ana_satirlar.get(b, []))
            sonuc['operatorler_eklenen'] = _operator_import(conn, wb, b)
        else:
            sonuc = _bolum_import(conn, wb, b)
        sonuclar[b] = sonuc
        for k in toplam:
            toplam[k] += sonuc.get(k, 0)

    conn.commit()
    conn.close()

    sonuc = {'basarili': True, **toplam, 'detay': sonuclar}
    if ana_rapor is not None:
        from datetime import datetime as _dt
        sonuc['kaynak_dosya'] = {'yol': ANA_VERI_YOL, 'degisme': _dt.fromtimestamp(
            os.path.getmtime(ANA_VERI_YOL)).strftime('%d.%m.%Y %H:%M')}
        sonuc['bolumsuz_satir'] = len(ana_rapor['bolumsuz'])
    return sonuc


def _program_listesi_import(conn, wb):
    """Robot Program Listesi sayfası → robot_programlari tablosu.
    Sayfa formatı (matrix):
       Satır 0: ROBOT | RAFNO | ABB-1 | ABB-1 | ABB-2 | ABB-2 | ... (her robot 2 kez)
       Satır 1: İSTASYON | <boş> | İST-1 | İST-2 | İST-1 | İST-2 | ...
       Satır 2+: <referans_kodu> | <raf_no> | √ | <boş> | √ | √ | ...
    Bu matrix'i düzleştirip her √ işareti için bir satır INSERT eder.
    """
    if wb is None or ROBOT_PROGRAM_SAYFA not in wb.sheetnames:
        return {'eklenen': 0, 'silinen': 0, 'hata': 'sayfa yok'}

    kayitlar = _program_listesi_oku(wb[ROBOT_PROGRAM_SAYFA])
    if kayitlar is None:
        return {'eklenen': 0, 'silinen': 0, 'hata': 'yetersiz satır'}

    c = conn.cursor()
    # Mevcut programları temizle (Excel master)
    c.execute('DELETE FROM robot_programlari')
    for robot_no, ist, ref in kayitlar:
        c.execute(
            'INSERT INTO robot_programlari (robot_no, istasyon, referans_kodu, guncelleyen) VALUES (?, ?, ?, ?)',
            (robot_no, ist, ref, 'Excel İçe Aktar')
        )
    print(f"  Robot Program: {len(kayitlar)} satır eklendi")
    return {'eklenen': len(kayitlar)}


def _program_listesi_oku(ws):
    """Robot Program Listesi matrisini düzleştirir → [(robot_no, istasyon, referans)]
    (yetersiz satır → None). Veritabanına dokunmaz."""
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 3:
        return None

    # Satır 0: robot adları (MERGED — birden çok kolonu kapsar)
    # Satır 1: istasyon. Kolon 0=referans, 1=raf
    robot_satiri = rows[0]
    istasyon_satiri = rows[1]
    # Merged cell mantığı: None olan kolonda son görülen robot devam eder
    robot_son = ''
    kolon_eslesme = []  # [(col_idx, robot_no, istasyon), ...]
    for j in range(2, max(len(robot_satiri), len(istasyon_satiri))):
        # Robot adı — yeni değer varsa al, yoksa son görüleni kullan (merged)
        r_raw = robot_satiri[j] if j < len(robot_satiri) and robot_satiri[j] is not None else None
        if r_raw is not None:
            robot_son = str(r_raw).strip()
        if not robot_son:
            continue
        # İstasyon
        i_raw = istasyon_satiri[j] if j < len(istasyon_satiri) and istasyon_satiri[j] is not None else None
        if i_raw is None:
            continue
        i_str = str(i_raw).strip()
        robot_no = robot_son.replace('-', '').replace(' ', '')  # ABB-1 → ABB1
        # İstasyon: "İST-1" → 1
        ist = 0
        if '1' in i_str: ist = 1
        elif '2' in i_str: ist = 2
        elif '3' in i_str: ist = 3
        if ist > 0:
            kolon_eslesme.append((j, robot_no, ist))

    kayitlar = []
    for r in rows[2:]:
        if not r or r[0] is None: continue
        ref = str(r[0]).strip()
        if not ref or len(ref) < 2: continue
        for col_idx, robot_no, ist in kolon_eslesme:
            if col_idx >= len(r): continue
            val = str(r[col_idx] or '').strip()
            if val and val != '':  # √ veya başka bir işaret varsa
                kayitlar.append((robot_no, ist, ref))
    return kayitlar


def _fikstur_raf_import(conn, wb):
    """Fikstür Raf Listesi sayfası → fikstur_raf tablosu.
    Sayfa formatı: 3 raf yan yana (A, B, C). Her raf 2 kolon: kod | raf_no.
    Aralarda boş kolon olabilir.
    """
    if wb is None or FIKSTUR_RAF_SAYFA not in wb.sheetnames:
        return {'eklenen': 0, 'hata': 'sayfa yok'}

    kayitlar = _fikstur_raf_oku(wb[FIKSTUR_RAF_SAYFA])
    if kayitlar is None:
        return {'eklenen': 0, 'hata': 'yetersiz satır'}

    c = conn.cursor()
    c.execute('DELETE FROM fikstur_raf')
    for kod, raf in kayitlar:
        c.execute('INSERT INTO fikstur_raf (referans_kodu, raf_no) VALUES (?, ?)', (kod, raf))
    print(f"  Fikstür Raf: {len(kayitlar)} satır eklendi")
    return {'eklenen': len(kayitlar)}


def _fikstur_raf_oku(ws):
    """Fikstür Raf Listesi → [(kod, raf_no)] (yetersiz satır → None). Veritabanına dokunmaz."""
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        return None
    kayitlar = []
    # Her satırdaki tüm (kod, raf_no) çiftlerini topla
    for r in rows[1:]:  # Başlık satırını atla
        if not r: continue
        # Her 3 kolonda bir grup: (kod_col, raf_col, boş)
        i = 0
        while i < len(r):
            kod = str(r[i] or '').strip() if i < len(r) else ''
            raf = str(r[i+1] or '').strip() if i+1 < len(r) else ''
            if kod and raf:
                kayitlar.append((kod, raf))
            i += 3  # Sonraki grup
    return kayitlar


def kaynak_ek_import():
    """Robot Program Listesi + Fikstür Raf sayfalarını içe alır — kaynak alanının ek
    sayfaları. Eski 'Toplu Veri Yönetimi' panelinden taşındı: dashboard'da kaynak bölümü
    için 'Excel'den Aktar' artık bunları da kapsar (tek buton, tek akış)."""
    wb = _tk2_okuma_kumesi()
    if wb is None:
        return {'program_eklenen': 0, 'fikstur_eklenen': 0}
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    try:
        sonuc = {}
        try:
            p = _program_listesi_import(conn, wb)
            sonuc['program_eklenen'] = p.get('eklenen', 0)
        except Exception as e:
            print(f"  Robot Program HATA: {e}")
            sonuc['program_eklenen'] = 0
        try:
            f = _fikstur_raf_import(conn, wb)
            sonuc['fikstur_eklenen'] = f.get('eklenen', 0)
        except Exception as e:
            print(f"  Fikstür HATA: {e}")
            sonuc['fikstur_eklenen'] = 0
        conn.commit()
        return sonuc
    finally:
        conn.close()


def import_tum(yedek_al=False):
    """Excel'deki TÜM sayfaları okuyup DB'yi günceller:
       - Tüm bölüm referansları (cycle time)
       - Tüm bölüm operatörleri
       - Robot programları (kaynak için matrix)
       - Fikstür raf listesi
       - Duruş sebepleri (validasyon — read on-demand)
    """
    if not os.path.exists(EXCEL_YOL) and not ana_veri_aktif():
        return {'basarili': False, 'hata': f'Excel bulunamadı: {EXCEL_YOL}'}

    # Önce normal referans+operator (mevcut)
    sonuc = import_data()

    # Sonra ek sayfalar (Ana Veri varsa oradan, sayfa yoksa eski dosyadan)
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    try:
        wb = _tk2_okuma_kumesi()

        # Robot Program
        try:
            prog_sonuc = _program_listesi_import(conn, wb)
            sonuc['program_eklenen'] = prog_sonuc.get('eklenen', 0)
        except Exception as e:
            print(f"  Robot Program HATA: {e}")
            sonuc['program_eklenen'] = 0

        # Fikstür Raf
        try:
            fik_sonuc = _fikstur_raf_import(conn, wb)
            sonuc['fikstur_eklenen'] = fik_sonuc.get('eklenen', 0)
        except Exception as e:
            print(f"  Fikstür HATA: {e}")
            sonuc['fikstur_eklenen'] = 0

        conn.commit()
    finally:
        conn.close()

    # Duruş sebepleri validasyonu — her bölüm için sayıyı raporla
    # (gerçek DB import yok, read-on-demand; ama kullanıcı sayıyı görsün)
    durus_ozet = {}
    for b in BOLUM_DURUS_SAYFA.keys():
        durus_ozet[b] = len(durus_sebepleri_yukle(b))
    sonuc['durus_sebepleri'] = durus_ozet

    return sonuc


def export_referans_cycle_times(bolum=None, lokasyon='TK2', zorla_ekle=None):
    """DB'deki cycle_time'ları Excel'in <Bolum> Referans sayfa(lar)ına yazar (SADECE TK2).
    Diğer veriler (operatör, duruş, program, fikstür) korunur.
    data/AnaVeri.xlsx varsa oraya yazar (bkz. _ana_veri_export); zorla_ekle = süresi
    olmasa da listeye eklenecek kodlar (panelden bilerek açılan referans).

    lokasyon='TK1' ise ATLA: TK1 referansları cycle time kullanmaz ve ayrı dosyadadır
    (data/Tk1 Veriler.xlsx, tek kolon) — TK1 verisi TK2 Excel'ine sızmamalı.
    """
    if (lokasyon or 'TK2').upper() == 'TK1':
        return {'basarili': True, 'atlandi': 'TK1 (cycle time export yok)', 'yazilan': 0}
    # GELİŞTİRME KOPYASI (2026-09-30): buradaki Excel sunucudan indirilmiş SALT OKUNUR
    # bir kopyadır. Yazmaya çalışmak 'dosya açık' hatası verir ve daha kötüsü, iki
    # ayrı Excel'in ayrışmasını yeniden başlatır. Tek kaynak sunucudaki dosya.
    if os.path.exists(os.path.join(PROJECT_DIR, 'data', 'GELISTIRME_KOPYASI.json')):
        return {'basarili': True, 'yazilan': 0,
                'atlandi': 'geliştirme kopyası — Excel yalnız canlı sunucuda güncellenir'}
    # KURULUM PROFİLİ: Excel senkronu kapalı kurulumda (referanslar yalnız veritabanında)
    # her süre teyidinde 'Excel bulunamadı' uyarısı çıkmasın.
    try:
        import kurulum as _kur
        if not _kur.modul('excel_senkron'):
            return {'basarili': True, 'yazilan': 0, 'atlandi': 'Excel senkronu bu kurulumda kapalı'}
    except Exception:
        pass
    if ana_veri_aktif():
        conn = sqlite3.connect(DB_PATH, timeout=20.0)
        conn.row_factory = sqlite3.Row
        try:
            return _ana_veri_export(conn, [bolum] if bolum else list(ANA_VERI_ETIKET), zorla_ekle)
        finally:
            conn.close()
    if not os.path.exists(EXCEL_YOL):
        return {'basarili': False, 'hata': f'Excel bulunamadı: {EXCEL_YOL}'}

    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    conn.row_factory = sqlite3.Row
    bolum_listesi = [bolum] if bolum else list(BOLUM_SAYFA.keys())
    toplam_yazilan = 0
    # BOZUK DOSYA KURTARMA (kullanıcı 2026-09-16: "Excel'e Yaz → Error -3 while
    # decompressing data: invalid distance too far back"): xlsx bir zip'tir, bozulunca
    # openpyxl zlib hatası verir. Okuma yolu zaten son bilinen iyi kopyaya düşüyordu,
    # yazma yolu çöküyordu. Artık yedekten devam edilir; bozuk dosya SİLİNMEZ,
    # '.bozuk-<zaman>' adıyla saklanır ve hangi kopyadan devam edildiği DÖNER —
    # sessiz kurtarma, kimsenin fark etmediği eski bir dosyaya yazmak olurdu.
    # NOT: yedek burada data_only=False okunur — _yedekten_oku (data_only=True)
    # kullanılsaydı kaydederken Excel'deki FORMÜLLER (kaynak Toplam kolonu) silinirdi.
    kurtarma = ''
    try:
        with open(EXCEL_YOL, 'rb') as _fh:
            wb = openpyxl.load_workbook(io.BytesIO(_fh.read()))
    except Exception as _ex:
        _yol = _yedek_yolu(EXCEL_YOL)
        _wb2 = None
        if os.path.exists(_yol):
            try:
                from datetime import datetime as _dt
                with open(_yol, 'rb') as _f2:
                    _wb2 = openpyxl.load_workbook(io.BytesIO(_f2.read()))
                _ytar = _dt.fromtimestamp(os.path.getmtime(_yol)).strftime('%d.%m.%Y %H:%M')
            except Exception as _e3:
                print(f'[export] yedek de okunamadi: {_e3}')
                _wb2 = None
        if _wb2 is None:
            conn.close()
            return {'basarili': False,
                    'hata': f'Excel dosyası BOZUK ve kullanılabilir yedek yok ({_ex}). '
                            f'Dosyayı sağlam bir kopyayla değiştirin: {EXCEL_YOL}'}
        try:
            from datetime import datetime as _dt2
            _bz = EXCEL_YOL + '.bozuk-' + _dt2.now().strftime('%Y%m%d-%H%M%S')
            os.replace(EXCEL_YOL, _bz)
        except Exception as _e2:
            conn.close()
            return {'basarili': False,
                    'hata': f'Excel BOZUK, bozuk dosya kenara alınamadı ({_e2}). '
                            f'Dosya açık olabilir — kapatıp tekrar deneyin: {EXCEL_YOL}'}
        wb = _wb2
        kurtarma = (f'ANA DOSYA BOZUKTU ({_ex}) — {_ytar} tarihli son bilinen iyi kopyadan '
                    f'devam edildi. Bozuk dosya: {os.path.basename(_bz)}. Yedekten sonraki '
                    f'Excel düzenlemeleri KAYBOLMUŞ olabilir, listeyi gözden geçirin.')
        print('[export_referans_cycle_times] ' + kurtarma)

    for b in bolum_listesi:
        sayfa_adi = BOLUM_SAYFA[b]['ref']
        if sayfa_adi not in wb.sheetnames:
            continue
        ws = wb[sayfa_adi]
        kaynak_modu = (b == 'kaynak')
        # Mevcut Excel satırlarını oku, kod → (satır, sayfadaki ham yazım) map'i
        kod_satir = {}
        for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if i == 1: continue  # Başlık
            if not row or row[0] is None: continue
            kod = str(row[0]).strip()
            kod_satir[kod.upper().replace(' ', '')] = (i, kod)

        # DB'deki bu bölüme ait referansları çek — SADECE TK2 (TK1 ayrı dosyada/cycle'sız).
        # SÜRESİZLER (ct=0/NULL) HARİÇ: Excel resmî/tanımlı referans listesidir; operatör
        # üretim girişinden auto-create ile doğan "Süre Tanımı Bekleyen" kodlar Excel'e
        # YAZILMAZ (yeni satır olarak eklenmez, Excel'de zaten varsa değeri de ezilmez).
        # Süre tanımlanınca (ct>0) bir sonraki sync'te Excel'e girer.
        db_rows = conn.execute(
            "SELECT referans_kodu, hedef_cycle_time_sn, kaynak_suresi_sn, soktak_suresi_sn, "
            "COALESCE(sure_teyit,0) as sure_teyit, COALESCE(aciklama,'') as aciklama, "
            "COALESCE(bukum_operasyon,1) as bukum_operasyon FROM referans_listesi "
            "WHERE COALESCE(bolum,'kaynak')=? AND COALESCE(lokasyon,'TK2')='TK2' "
            "AND COALESCE(hedef_cycle_time_sn,0) > 0 ORDER BY referans_kodu",
            (b,)
        ).fetchall()

        # Kaynak sayfası 4+1 kolonlu: Kod | Kaynak(B) | Söktak(C) | Toplam(D formül) | Süre Teyit(E)
        # (ESKİ BUG: her bölümde B'ye toplam cycle yazılıyordu — kaynakta B=Kaynak Süresi
        #  kolonunu eziyordu. Montaj/metal 2 kolonlu: Kod | Cycle(B) — o davranış doğru.)
        if kaynak_modu and (ws.cell(row=1, column=5).value or '') == '':
            ws.cell(row=1, column=5, value='Süre Teyit')
        # Pres sayfası 3 kolonluydu — D başlığı yoksa aç (2026-07-29 büküm operasyonu)
        if b == 'pres' and (ws.cell(row=1, column=4).value or '') == '':
            ws.cell(row=1, column=4, value='Büküm Op.')

        yazilan = 0
        # Aynı koda normalize olan birden fazla DB satırı (eski yazım varyantları:
        # '94.LTK.10' + '94.ltk.10') aynı Excel satırına düşer — sayfadaki yazımla
        # birebir eşleşen satır ÖNCELİKLİDİR; bayat varyantın taze değeri ezmesine izin verme.
        yazilan_norm = {}
        for r in db_rows:
            kod = (r['referans_kodu'] or '').strip()
            norm = kod.upper().replace(' ', '')
            if norm in kod_satir:
                ri, sayfa_kod = kod_satir[norm]
                tam_es = (kod == sayfa_kod)
            else:
                ri = ws.max_row + 1
                ws.cell(row=ri, column=1, value=kod)
                kod_satir[norm] = (ri, kod)
                tam_es = True
            if yazilan_norm.get(norm) and not tam_es:
                continue  # bu koda tam-eş (veya ilk) yazım zaten yazıldı — varyantla ezme
            if kaynak_modu:
                ws.cell(row=ri, column=2, value=r['kaynak_suresi_sn'] or 0)
                ws.cell(row=ri, column=3, value=r['soktak_suresi_sn'] or 0)
                ws.cell(row=ri, column=4, value=f'=B{ri}+C{ri}')
                ws.cell(row=ri, column=5, value='EVET' if r['sure_teyit'] else '')
            elif b == 'pres':
                # Pres 4 kolonlu: B=Açıklama, C=Süre, D=Büküm Op.
                # (süreyi B'ye yazmak açıklamayı ezerdi)
                ws.cell(row=ri, column=2, value=r['aciklama'] or '')
                ws.cell(row=ri, column=3, value=r['hedef_cycle_time_sn'] or 0)
                ws.cell(row=ri, column=4, value=int(r['bukum_operasyon'] or 1))
            else:
                ws.cell(row=ri, column=2, value=r['hedef_cycle_time_sn'] or 0)
            yazilan_norm[norm] = True
            yazilan += 1
        toplam_yazilan += yazilan
        print(f"  {b}: {yazilan} satır Excel'e yazıldı")

    conn.close()
    try:
        wb.save(EXCEL_YOL)
    except PermissionError:
        # Dosya Excel'de/OneDrive'da AÇIK — kullanıcı sessiz kayıp yaşamasın, net mesaj dön
        return {'basarili': False,
                'hata': 'Excel dosyası şu an açık (uretim_verileri.xlsx) — kapatıp tekrar deneyin.'}
    sonuc = {'basarili': True, 'yazilan': toplam_yazilan, 'dosya': EXCEL_YOL}
    if kurtarma:
        sonuc['uyari'] = kurtarma      # panel BU metni göstermeli (sessiz kurtarma yok)
    return sonuc


if __name__ == '__main__':
    res = import_tum()
    print("\n✅ Import tamamlandı:", res)
