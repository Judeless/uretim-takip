# -*- coding: utf-8 -*-
"""Kaynak planındaki referansların ÜRETİLEBİLİR ADEDİNİ ERP'den hesaplar (2026-07-30).

NEDEN: Planlama her hafta bir "Kaynak ihtiyaçları" dosyası çıkarıyor; referanslar
önem sırasına dizili. Launch talimatı vermek için her kod AS400'de 07>10>01
ekranında tek tek açılıp 9+Enter ile ürün ağacı görüntüleniyor ve 1. seviye alt
parçaların 01D/CF2 stoğu gözle kontrol ediliyor. 230 satır için bu iş saatler
sürüyor ve hata kaldırmıyor.

BU SCRIPT aynı kontrolü toplu yapar:
  1. Plan dosyasındaki "Kaynak kodu" sütununu SIRASIYLA okur (sıra = öncelik)
  2. Her kod için ERP'den 1. SEVİYE ürün ağacını çeker      (TKC0301F.BSPEF2)
  3. Alt parçaların depo stoklarını çeker                    (TKC0300F.BSMAF0)
  4. Üretilebilir adet = min(alt parça stoğu / birim miktar) — en kısıtlayan parça
  5. Excel raporu yazar: özet + parça kırılımı

STOK KURALI: kullanıcının elle yaptığı kontrolle aynı — yalnız 01D ve CF2.
Diğer depolar (01W, REP, MDT...) rapora bilgi olarak yazılır ama hesaba GİRMEZ.

NEGATİF STOK: ERP'de eksi bakiye olabiliyor (girilmemiş hareket, ters kayıt).
Eksi bakiye "o kadar üretilebilir" demek değildir → 0 sayılır ve işaretlenir.
İSTİSNA yalnız CF2 (2026-09-30): 01D'deki eksi CF2'deki stoğu götürmez; diğer
depolar (MK2, MT2, CF) 01D ile NETLEŞİR — bkz. hesapla().

KULLANIM:
    python kaynak_plan_kontrol.py
    python kaynak_plan_kontrol.py --plan "Q:\\...\\Kaynak ihtiyaçları 260729.xlsb"
    python kaynak_plan_kontrol.py --limit 20        (ilk 20 kod — hızlı deneme)
"""
import argparse
import math
import os
import re
import sys
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

KOK = os.path.dirname(os.path.abspath(__file__))
VARSAYILAN_PLAN = r"Q:\UretimPlanlama\Yarımamul Üretim Planları\Kaynak ihtiyaçları 260729.xlsb"
SAYILAN_DEPOLAR = ('01D', 'CF2')          # kullanıcının kuralı
# 01D eksisiyle NETLEŞMEYEN depolar — yalnız CF2 (bkz. hesapla).
AYRI_SAYILAN_DEPOLAR = ('CF2',)
# Stoğu SAYILMAYAN ama EKSİ bakiyesi serbest stoktan DÜŞÜLEN depolar (bkz. hesapla).
EKSI_DUSULEN_DEPOLAR = ('REP', '02')
GOSTERILEN_DEPOLAR = ('01D', 'CF2', '01W', 'REP', '02', 'MDT', 'MK2', 'MT2')
GUVENLI_KOD = re.compile(r'^[A-Za-z0-9./\- ]{3,21}$')
# Acik uretim emirleri gorunumu (as400_config.KAYNAK_TABLO ile ayni kaynak)
OPR_TABLO = 'tkc0301F.XPRO90'
IHTIYAC_UFUK_GUN = 42          # 6 hafta — kullanicinin 1845W ornegiyle uyumlu


def _satirlari_getir(yol):
    """Dosyayı uzantısına göre açar ve satırları {kolon_index: değer} olarak verir.

    .xlsb  → pyxlsb (planlamanın çıkardığı özgün biçim)
    .xlsx  → openpyxl (zaten kurulu; pyxlsb kurulamadığında kaçış yolu —
             kullanıcı dosyayı Excel'de "Farklı Kaydet > .xlsx" ile verebilir)
    pyxlsb yoksa hata mesajı NE YAPILACAĞINI söyler (2026-07-31: sunucuda modül
    kurulu olmadığı için ham "No module named 'pyxlsb'" görünüyordu)."""
    uzanti = os.path.splitext(yol)[1].lower()
    if uzanti == '.xlsb':
        try:
            from pyxlsb import open_workbook
        except ImportError:
            raise RuntimeError(
                "Bu sunucuda .xlsb okuyucusu (pyxlsb) kurulu değil. İki çözüm: "
                "(1) sunucuda 'pip install pyxlsb' çalıştırıp servisi yeniden başlatın, "
                "(2) ya da dosyayı Excel'de 'Farklı Kaydet > Excel Çalışma Kitabı (.xlsx)' "
                "ile kaydedip .xlsx olarak yükleyin — o biçim ek kurulum istemez.")
        with open_workbook(yol) as wb:
            sayfa = 'Kaynak' if 'Kaynak' in wb.sheets else wb.sheets[0]
            with wb.get_sheet(sayfa) as s:
                for satir in s.rows():
                    yield {c.c: c.v for c in satir}
        return
    if uzanti in ('.xlsx', '.xlsm'):
        import openpyxl
        wb = openpyxl.load_workbook(yol, read_only=True, data_only=True)
        try:
            ad = 'Kaynak' if 'Kaynak' in wb.sheetnames else wb.sheetnames[0]
            for satir in wb[ad].iter_rows(values_only=True):
                yield {i: v for i, v in enumerate(satir)}
        finally:
            wb.close()
        return
    raise RuntimeError(f'Desteklenmeyen dosya biçimi: {uzanti or "(uzantısız)"} — .xlsb veya .xlsx olmalı')


def plan_oku(yol, limit=None):
    """Plan dosyasının 'Kaynak' sayfasını okur. [{sira, kaynak_kod, urun, ustler}]
    Satır sırası = öncelik; korunur."""
    satirlar, gorulen = [], set()

    def _say(h, idx):
        try:
            return float(h.get(idx))
        except (TypeError, ValueError):
            return 0.0

    for i, h in enumerate(_satirlari_getir(yol)):
        if i == 0:
            continue                              # başlık satırı
        kod = str(h.get(5) or '').strip()         # F = Kaynak kodu
        if not kod or kod in gorulen:
            continue
        gorulen.add(kod)
        satirlar.append({
            'sira': len(satirlar) + 1,
            'kaynak_kod': kod,
            'urun': str(h.get(1) or '').strip(),
            'ustler': [str(h.get(k) or '').strip() for k in (2, 3, 4) if str(h.get(k) or '').strip()],
            # ── İHTİYAÇ ÖLÇÜTÜ (kullanıcı 2026-07-31, sütunları kendisi verdi) ──
            #   BO (66) "Toplam 6 haft iht"  = ihtiyaç adedi
            #   BV (73) "Launch"             = mevcutta üretime alınan adet
            #   BQ (68) "toplam stok"
            # Kural: üretime alınan ihtiyaçtan azsa KALANI üretmek gerekir ve
            # eldeki stok da ihtiyaçtan düşülür:
            #     GEREKEN = max(0, BO − BV − toplam stok)
            # NOT: dosyadaki hazır "6 hafta için fark" (BT/71) sütunu KULLANILMIYOR;
            # o sütun negatif stoğu yok sayıyor ve 12 satırda bu formülden
            # ayrışıyor (ör. BO=100, stok=−44 → BT 100 der, bu kural 144).
            'acik_launch': _say(h, 73),          # BV — üretime alınan
            'bakiye': _say(h, 15),
            'toplam_stok': _say(h, 68),          # BQ
            'iht_2h': _say(h, 20),
            'iht_6h': _say(h, 66),               # BO — 6 haftalık ihtiyaç
            'kontrol6': (None if h.get(74) is None else bool(h.get(74))),
            'gun_stok': _say(h, 75),
            # Planlamanın elle yazdığı alt parçalar — ERP ağacıyla kıyaslanır
            'plan_parcalar': [str(h.get(k) or '').strip()
                              for k in range(6, 13) if str(h.get(k) or '').strip()],
        })
        son = satirlar[-1]
        son['gereken'] = max(0.0, (son['iht_6h'] or 0) - (son['acik_launch'] or 0)
                            - (son['toplam_stok'] or 0))
        son['stok_eksi'] = (son['toplam_stok'] or 0) < 0
        if limit and len(satirlar) >= limit:
            break
    return satirlar


def erp_baglan():
    d = os.path.join(KOK, 'as400')
    if d not in sys.path:
        sys.path.insert(0, d)
    import launch_cek
    return launch_cek.baglan(timeout=60)


def urun_agaci(cn, kodlar):
    """1. seviye açılım: {ust_kod: [(alt_kod, birim_miktar, birim)]}"""
    agac = {}
    kodlar = [k for k in kodlar if GUVENLI_KOD.match(k)]
    cu = cn.cursor()
    for i in range(0, len(kodlar), 50):
        grup = kodlar[i:i + 50]
        yer = ','.join('?' * len(grup))
        cu.execute(
            f"SELECT TRIM(Y2ARTI), TRIM(Y2RICD), Y2IVQT, Y2IVUM "
            f"FROM TKC0301F.BSPEF2 WHERE TRIM(Y2ARTI) IN ({yer})", grup)
        for ust, alt, miktar, um in cu.fetchall():
            if not (ust and alt):
                continue
            agac.setdefault(ust, []).append((alt, float(miktar or 0), str(um or '').strip()))
    return agac


HAYALI_BAYRAK = '2'       # BARTF0.A0PROD — ekranda 'Product.Explos. 2 = fictit. assembly'


def hayali_kodlar(cn, kodlar):
    """kodlar içinden HAYALİ MONTAJ (fictitious assembly) olanların kümesi.

    Kaynak: BARTF0.A0PROD = '2'. Hayali parça stokta TUTULMAZ — yalnızca ürün
    ağacında bir gruplama düğümüdür; stoğuna bakmak her zaman 0 verir ve ürünü
    'malzeme yok' gösterir. Doğrusu kendi alt parçalarına inmektir."""
    kodlar = [k for k in kodlar if GUVENLI_KOD.match(k)]
    bulunan = set()
    cu = cn.cursor()
    for i in range(0, len(kodlar), 50):
        grup = kodlar[i:i + 50]
        yer = ','.join('?' * len(grup))
        cu.execute(
            f"SELECT TRIM(A0ARTI) FROM TKC0301F.BARTF0 "
            f"WHERE A0PROD='{HAYALI_BAYRAK}' AND A0ARTI IN ({yer})", grup)
        bulunan.update(r[0] for r in cu.fetchall())
    return bulunan


def dokum_kodlari(cn, kodlar, hammadde_onek=('21.',), azami_seviye=4):
    """ENJEKSİYONDA BASILAN KOD — {kod: [döküm kodları]}.

    Kullanıcı 2026-10-01: "Enjeksiyon kodlarında en alt kademedeki 21.AL.150 ya da
    21.AL.036 hammaddesinin 1 kademe üstündeki referans bizim enjeksiyonda bastığımız
    kod oluyor; diğer üst kodlar sonraki işleme adımları için." Örnek: Forge'daki
    10.300.1369 bir İŞLEME kodu (fason); enjeksiyonda basılan 10.300.1369W.

    Ağaçta aşağı inilir; doğrudan alt parçası hammadde (hammadde_onek) olan kod
    DÖKÜM KODUDUR. Hayali (fictitious) ara kod ŞEFFAFTIR: altındaki hammadde
    ebeveynine sayılır — hayali kod stoklanmaz, emri ebeveyne açılır
    (10.130.3778 → hayali 10.130.3778W → 21.AL.036: döküm kodu 10.130.3778).
    Birden çok kademe olabilir: 10.300.4708 → 4708C → 4708W → 21.AL.150.
    Hammaddeye inilemeyen kodun listesi BOŞ döner (ağacı yok / başka malzeme)."""
    hammadde_onek = tuple(hammadde_onek)
    agac, hayali = {}, set()
    sinir = [x for x in dict.fromkeys(kodlar) if x]
    for _ in range(azami_seviye + 1):
        sinir = [x for x in sinir if x not in agac]
        if not sinir:
            break
        agac.update({x: [] for x in sinir})
        agac.update(urun_agaci(cn, sinir))
        cocuk = sorted({a for x in sinir for a, _, _ in agac[x] if not a.startswith(hammadde_onek)})
        if cocuk:
            hayali |= hayali_kodlar(cn, cocuk)
        sinir = cocuk

    def etkin(x, d=0):
        for a, _, _ in agac.get(x, []):
            if a in hayali and d < azami_seviye:
                yield from etkin(a, d + 1)
            else:
                yield a

    def bul(x, d=0):
        alt = list(etkin(x))
        if any(a.startswith(hammadde_onek) for a in alt):
            return [x]
        if d >= azami_seviye:
            return []
        return sorted({y for a in alt if not a.startswith(hammadde_onek) for y in bul(a, d + 1)})
    return {x: bul(x) for x in dict.fromkeys(kodlar) if x}


def urun_agaci_hayali(cn, kodlar, azami_seviye=3, acilmayan_onekler=()):
    """Ürün ağacı — HAYALİ alt parçalar kendi alt parçalarına AÇILIR. (agac, iz)

    Kullanıcı 2026-09-30 (TK2 montaj planı): "1. seviye alt parçalara bakacağız;
    1. seviye parça fictitious tanımlıysa onun 2. seviye alt parçalarına, o da
    fictitious ise 3. seviye alt parçalarına bakacağız."

    agac: {ust: [(alt, birim, um)]} — hesapla() ile AYNI biçim. Miktarlar yol
          boyunca ÇARPILIR (üstte 2 adet hayali × altında 3 adet parça = 6); aynı
          yaprak birden çok yoldan geliyorsa TOPLANIR.
    iz:   {ust: {alt: {'seviye', 'yol', 'hayali'}}} — panelde parçanın hangi hayali
          düğümün altından geldiği; 'hayali' True ise azami seviyede HÂLÂ hayali
          (ya da ağacı yok) — o parça stoksuz görünür, elle bakılmalı.

    acilmayan_onekler: bu ön ekle başlayan alt parça hayali OLSA BİLE açılmaz
          (montaj: '93.' = tel; TK1'de ya da fasonda üretilir, alt parçaları
          TK2'nin kısıtı değildir — kullanıcı 2026-09-30)."""
    acilmayan_onekler = tuple(acilmayan_onekler or ())
    duz = urun_agaci(cn, kodlar)
    calisma = {u: [(a, q, um, 1, '') for a, q, um in lst] for u, lst in duz.items()}
    hayali_kalan = set()
    for seviye in range(1, max(1, int(azami_seviye)) + 1):
        adaylar = sorted({a for lst in calisma.values() for a, _q, _u, sv, _y in lst if sv == seviye})
        if not adaylar:
            break
        hayali = hayali_kodlar(cn, adaylar)
        if acilmayan_onekler:
            muaf = {h for h in hayali if h.startswith(acilmayan_onekler)}
            hayali_kalan |= muaf            # açılmaz; kısıt da sayılmaz
            hayali -= muaf
        if not hayali:
            break
        if seviye >= azami_seviye:
            hayali_kalan |= hayali          # daha derine inilmez; işaretle
            break
        alt_agac = urun_agaci(cn, sorted(hayali))
        hayali_kalan |= {h for h in hayali if not alt_agac.get(h)}   # ağaçsız hayali
        for u, lst in calisma.items():
            yeni = []
            for a, q, um, sv, yol in lst:
                if sv == seviye and a in hayali and alt_agac.get(a):
                    for a2, q2, um2 in alt_agac[a]:
                        yeni.append((a2, q * q2, um2, sv + 1, (yol + ' › ' if yol else '') + a))
                else:
                    yeni.append((a, q, um, sv, yol))
            calisma[u] = yeni
    agac, iz = {}, {}
    for u, lst in calisma.items():
        toplam = {}
        for a, q, um, sv, yol in lst:
            d = toplam.setdefault(a, {'q': 0.0, 'um': um, 'seviye': sv, 'yollar': []})
            d['q'] += q
            d['seviye'] = min(d['seviye'], sv)
            if yol and yol not in d['yollar']:
                d['yollar'].append(yol)
        agac[u] = [(a, d['q'], d['um']) for a, d in toplam.items()]
        iz[u] = {a: {'seviye': d['seviye'], 'yol': ' | '.join(d['yollar']),
                     'hayali': a in hayali_kalan} for a, d in toplam.items()}
    return agac, iz

def stoklar(cn, kodlar):
    """{artikel: {depo: stok}} — tüm depolar (hesapta yalnız SAYILAN_DEPOLAR)."""
    st = {}
    kodlar = [k for k in kodlar if GUVENLI_KOD.match(k)]
    cu = cn.cursor()
    for i in range(0, len(kodlar), 50):
        grup = kodlar[i:i + 50]
        yer = ','.join('?' * len(grup))
        cu.execute(
            f"SELECT TRIM(S0ARTI), TRIM(S0MGCD), SUM(S0GIAD) FROM TKC0300F.BSMAF0 "
            f"WHERE S0SOCI='01' AND S0ARTI IN ({yer}) GROUP BY TRIM(S0ARTI), TRIM(S0MGCD)", grup)
        for art, depo, q in cu.fetchall():
            st.setdefault(art, {})[depo] = float(q or 0)
    return st


def opr_ihtiyaclari(cn, kodlar, ufuk_gun=42):
    """AÇIK ÜRETİM EMİRLERİNDEN (OPR) ihtiyaç — plan dosyasına gerek yok.

    Kullanıcı 2026-07-31: "referansın ihtiyaç adedi, 10 07 01 ekranında kodun
    içine girildiğinde 2 ile tekrar içine girildiğinde O PR numaralarından
    anlayabiliriz... opr tarihi en eskiden itibaren sıraladığımızda aciliyet
    sırasına göre sıralamış oluruz."

    Kaynak: tkc0301F.XPRO90 (yalnız AÇIK emirler görünümü).
      Q0AVAN 10 = OPR var, launch alınmamış → kalan = sipariş adedi
      Q0AVAN 40/45/50 = launch açık        → kalan = sipariş − teyit verilen
      Q0FPD* = üretim bitiş (teslim) tarihi → aciliyet sırası ve ufuk filtresi

    ufuk_gun: bugünden itibaren kaç gün ilerisi sayılsın (varsayılan 42 = 6 hafta).
    Kullanıcının 10.300.1845W örneği bununla birebir tutuyor: 29/06'dan kalan 86
    + 27/07 tarihli 270 + 31/08 tarihli 732 = 1088; 19/10 tarihli 4. emir ufkun
    dışında kaldığı için sayılmıyor.

    Döner: {kod: {'ihtiyac', 'en_eski', 'opr_sayisi', 'gecikmis', 'satirlar'[]}}
    """
    from datetime import date, timedelta
    kodlar = [k for k in kodlar if GUVENLI_KOD.match(k)]
    if not kodlar:
        return {}
    bugun = date.today()
    sinir = bugun + timedelta(days=ufuk_gun)
    sonuc = {}
    cu = cn.cursor()
    for i in range(0, len(kodlar), 50):
        grup = kodlar[i:i + 50]
        yer = ','.join('?' * len(grup))
        cu.execute(
            f"SELECT TRIM(Q0ARTI), TRIM(Q0AVAN), Q0QTOR, Q0QTRI, Q0RED2, Q0RENU, "
            f"Q0FPD1, Q0FPD2, Q0FPD3, Q0FPD4 "
            f"FROM {OPR_TABLO} WHERE TRIM(Q0ARTI) IN ({yer}) "
            f"AND TRIM(Q0AVAN) IN ('10','40','45','50')", grup)
        for r in cu.fetchall():
            kod, durum = r[0], (r[1] or '').strip()
            adet, teyit = float(r[2] or 0), float(r[3] or 0)
            # durum 10'da henüz launch yok → teyit de olamaz; 40/45/50'de kalan düşülür
            kalan = adet if durum == '10' else max(0.0, adet - teyit)
            if kalan <= 0:
                continue
            try:
                t = date(int(r[6]) * 100 + int(r[7]), int(r[8]), int(r[9]))
            except (TypeError, ValueError):
                t = None
            d = sonuc.setdefault(kod, {'ihtiyac': 0.0, 'en_eski': None, 'opr_sayisi': 0,
                                       'gecikmis': 0.0, 'satirlar': [],
                                       'launch_adet': 0.0, 'launch_sayisi': 0,
                                       'launch_durum': {}})
            # LAUNCH ALINMIŞ MI? (kullanıcı 2026-09-30): durum 40/45/50 = üretim
            # emri açılmış. Ufuktan BAĞIMSIZ sayılır — açık launch açık launch'tır;
            # amaç aynı parçaya ikinci kez emir açmamak.
            if durum != '10':
                d['launch_adet'] += kalan
                d['launch_sayisi'] += 1
                d['launch_durum'][durum] = d['launch_durum'].get(durum, 0) + 1
            d['satirlar'].append({'opr': f"{str(r[4]).strip()}-{str(r[5]).strip()}",
                                  'durum': durum, 'adet': adet, 'teyit': teyit,
                                  'kalan': kalan, 'tarih': t.isoformat() if t else '',
                                  'ufukta': bool(t and t <= sinir)})
            if t and t <= sinir:
                d['ihtiyac'] += kalan
                d['opr_sayisi'] += 1
                if t < bugun:
                    d['gecikmis'] += kalan            # teslim tarihi geçmiş
                if d['en_eski'] is None or t.isoformat() < d['en_eski']:
                    d['en_eski'] = t.isoformat()
    for d in sonuc.values():
        d['satirlar'].sort(key=lambda x: (x['tarih'] or '9999'))
        # '40×2 · 45×1' — panelde kodun yanındaki rozet
        d['launch_ozet'] = ' · '.join(f'{k}×{v}' for k, v in sorted(d['launch_durum'].items()))
    return sonuc


def oncelik_puani(opr_satirlar, stok, ufuk_gun=IHTIYAC_UFUK_GUN, bugun=None):
    """Aciliyet puanı: TARİH ile ADEDİ birlikte tartar. (puan, acik_adet, en_gec_gun)

    NEDEN (kullanıcı 2026-09-30): sıralama yalnız 'en eski OPR tarihi'ne göreydi.
    120 gün gecikmiş 2 adetlik bir kalıntı, 20 gün gecikmiş 300 adetlik işin
    önüne geçiyordu — oysa hattı durduracak olan ikincisi.

    HESAP: eldeki stok emirlere TARİH SIRASIYLA dağıtılır (en eski önce); stokla
    kapanmayan her emir için  açık adet × gün ağırlığı  toplanır.
        gün ağırlığı = ufuk + gecikme günü   (en az 1)
    Bugün teslim edilecek emir 'ufuk' kadar, 30 gün gecikmiş olan ufuk+30 kadar,
    ufkun sonundaki emir 1 kadar ağırlık alır — gecikme arttıkça ve adet
    büyüdükçe puan yükselir, ikisi birbirini dengeleyebilir.
        120 gün gecikmiş 2 adet   → 2 × 162  =    324
         20 gün gecikmiş 300 adet → 300 × 62 = 18.600  (öne geçer)

    en_gec_gun: stokla kapanmayan en eski emrin gecikme günü (+ = gecikmiş)."""
    from datetime import date
    bugun = bugun or date.today()
    kalan_stok = max(0.0, float(stok or 0))
    puan, acik_toplam, en_gec = 0.0, 0.0, None
    for x in sorted(opr_satirlar or [], key=lambda r: (r.get('tarih') or '9999')):
        if not x.get('ufukta') or not x.get('tarih'):
            continue
        kalan = float(x.get('kalan') or 0)
        karsilanan = min(kalan_stok, kalan)
        kalan_stok -= karsilanan
        acik = kalan - karsilanan
        if acik <= 0:
            continue
        try:
            gun = (bugun - date.fromisoformat(x['tarih'])).days
        except (TypeError, ValueError):
            continue
        puan += acik * max(1, int(ufuk_gun) + gun)
        acik_toplam += acik
        if en_gec is None or gun > en_gec:
            en_gec = gun
    return round(puan, 1), acik_toplam, en_gec


def stok_ggi(cn, kodlar):
    """Referansın ERP ekranında görünen stoğu (G GI) = 01D + MDT.

    Kullanıcı 2026-07-31: "07 10 01 ekranında kodun içine girip 2 ile içine
    girdiğimizde en üstte G GI kısmında görebiliriz, 300.1845w için 764".
    Ölçüldü: 01D 308 + MDT 456 = 764 — birebir. MDT fasondaki malzeme; ekranda
    stoğa dahil ediliyor, 01W dahil DEĞİL."""
    ham = stoklar(cn, kodlar)
    return {k: {'ggi': v.get('01D', 0) + v.get('MDT', 0), 'depolar': {a: b for a, b in v.items() if b}}
            for k, v in ham.items()}


def hesapla(satirlar, agac, stok, ref_stok=None, sayilan=None, gosterilen=None, iz=None,
            kisit_disi=None, eksi_dusulen=None):
    """Her plan satırına üretilebilir adet ve kısıtlayan parçayı ekler.

    ref_stok verilirse (kullanıcı 2026-07-31: "referans stoklarını excelden değil
    as400'den kontrol edelim") referansın KENDİ stoğu ERP'den alınır ve GEREKEN
    yeniden hesaplanır. Plan dosyasındaki stok sütunu haftalık — dosya çekildiği
    günün fotoğrafı; ölçüldü: 230 satırın yalnız 11'i ERP ile birebir tutuyor.
    Depo kuralı alt parçalarla AYNI: 01D + CF2.

    sayilan / gosterilen: alt parça stoğunda SAYILAN ve kırılımda GÖSTERİLEN depolar.
    Verilmezse kaynak planının kuralı (01D+CF2). Montaj planı 01D+CF2+MK2+MT2
    sayar — kaynaktan gelen yarı mamul transit depolarda (MK2/MT2) bekler."""
    sayilan = tuple(sayilan or SAYILAN_DEPOLAR)
    eksi_dusulen = tuple(EKSI_DUSULEN_DEPOLAR if eksi_dusulen is None else eksi_dusulen)
    gosterilen = tuple(gosterilen or GOSTERILEN_DEPOLAR)
    # iz: urun_agaci_hayali'nin ikinci çıktısı. 'hayali' işaretli parça (ağacı
    # olmayan ya da azami seviyede hâlâ hayali) stokta TUTULMAZ — stoğu hep 0
    # görünür. Onu kısıt saymak ürünü sonsuza dek 'malzeme yok' gösterirdi
    # (canlıda 50.002.109: 103 üründe, ağacı yok, her depoda 0). Listede
    # GÖSTERİLİR ama üretilebilir adedi KISITLAMAZ.
    # Kullanıcı teyidi 2026-09-30: "50.002.109 kablo ve hayali tanımlı; hayali
    # tanımlı referansta stok hareketi yapılmaz, launch alınırken de hesaba
    # katılmaz — sadece görüntü olarak reçetede durur."
    iz = iz or {}
    # kisit_disi: stoğu kontrol EDİLMEYECEK parçalar (etiket, sarf…). Kullanıcı
    # 2026-09-30: "50.010.700 etiket olduğu için kontrolden çıkartalım, bakmaya
    # gerek yok" — tek başına 25 mekanizmayı kilitliyordu (bakiyesi yalnız kayıp
    # depo 01W'de). Kod tam eşleşir; sonu '*' ise ön ek ('50.010.*').
    kisit_disi = tuple(str(k).strip() for k in (kisit_disi or ()) if str(k).strip())

    def _kisit_disi_mi(kod):
        return any(kod == k or (k.endswith('*') and kod.startswith(k[:-1])) for k in kisit_disi)
    for s in satirlar:
        if ref_stok is not None:
            d = ref_stok.get(s['kaynak_kod'], {})
            # plan_stok'a DOKUNMA: o, plan dosyasındaki değer ve yalnızca
            # yüklemede yazılır. Burada ezilirse ikinci ölçümde "kıyas" değeri
            # bir önceki ERP okumasına dönüşür ve karşılaştırma anlamını yitirir.
            s['toplam_stok'] = d.get('01D', 0) + d.get('CF2', 0)
            s['stok_kaynak'] = 'ERP'
            s['ref_depolar'] = {k: v for k, v in d.items() if v}
            s['gereken'] = max(0.0, (s.get('iht_6h') or 0) - (s.get('acik_launch') or 0)
                               - s['toplam_stok'])
        parcalar = agac.get(s['kaynak_kod']) or []
        s['parcalar'] = []
        if not parcalar:
            s['durum'] = 'AGAC YOK'
            s['uretilebilir'] = None
            s['kisitlayan'] = ''
            continue
        kapasiteler = []
        for alt, birim, um in sorted(parcalar):
            depolar = stok.get(alt, {})
            ham = sum(depolar.get(d, 0) for d in sayilan)        # net bakiye (bilgi)
            # CF2 AYRI SAYILIR, DİĞERLERİ 01D İLE NETLEŞİR (kullanıcı 2026-09-30):
            #   · "Stok CF2'de varsa ve 01D eksideyse, ürünü CF2'ye taşımışlar fakat alt
            #     koddan üst kodun 01D stoğuna henüz aktarım yapmamış demektir; CF2
            #     stoğunu doğru sayabiliriz." → CF2, 01D eksisinden ETKİLENMEZ.
            #   · "Ürün gelmiş, kullanılmış ve 01D eksiye düşmüş; sadece MK2'den 01D'ye
            #     aktarım bekliyor olabilir. SADECE CF2 için bu durum geçerli." → MK2 /
            #     MT2 / CF'deki artı, 01D'deki eksiyle AYNI malın iki kaydıdır: o stok
            #     çoktan tüketilmiş, serbest değildir. Bunlar toplanıp sonra sıfırlanır.
            # İlk denemede kural bütün depolara genellenmişti (depo bazında sıfırlama);
            # MK2'deki tüketilmiş stoğu 'var' gösteriyordu — kullanıcı düzeltti.
            # REP / 02 EKSİSİ SERBEST STOKTAN DÜŞÜLÜR (kullanıcı 2026-09-30, 10.300.4982G/10):
            # 01D 179 görünüyor ama REP −126 ve 02 −50: "REP ve 02 stoklarındaki eksi
            # miktarlara bakınca 01D adediyle örtüştüğünü görüyoruz; yani 01D'deki ürün
            # KULLANILMIŞ, stokta bulunmamakta. Bu kod launch alınabilir gibi
            # görünmemeli." Tüketim REP/02'den düşülmüş, mal kaydı hâlâ 01D'de. Bu
            # depoların ARTI bakiyesi sayılmaz (REP artısı başka emre rezerve).
            tuketilmis = sum(min(0.0, depolar.get(d, 0)) for d in eksi_dusulen
                             if d not in sayilan)
            eldeki = (max(0.0, sum(depolar.get(d, 0) for d in sayilan
                               if d not in AYRI_SAYILAN_DEPOLAR) + tuketilmis)
                      + sum(max(0.0, depolar.get(d, 0)) for d in sayilan
                            if d in AYRI_SAYILAN_DEPOLAR))
            _iz = (iz.get(s['kaynak_kod']) or {}).get(alt) or {}
            _muaf = _kisit_disi_mi(alt)
            if _iz.get('hayali') or _muaf:
                kap = None                  # hayali / kısıt dışı: gösterilir, sayılmaz
            elif birim > 0:
                kap = int(math.floor(eldeki / birim))
                kapasiteler.append((kap, alt))
            else:
                kap = None                  # birim tanımsız → kısıt sayma, işaretle
            s['parcalar'].append({
                'kod': alt, 'birim': birim, 'um': um,
                # stok_sayilan = SAYILAN stok (CF2 ayrı, diğerleri netleşmiş); net bakiye
                # ayrıca saklanır ki panel 'eksi var' uyarısını gösterebilsin.
                'stok_sayilan': eldeki, 'stok_net': ham,
                'eksi_bakiye': any(depolar.get(d, 0) < 0 for d in sayilan),
                'kapasite': kap,
                'depolar': {d: depolar[d] for d in gosterilen if depolar.get(d)},
                'seviye': _iz.get('seviye', 1), 'yol': _iz.get('yol', ''),
                'hayali': bool(_iz.get('hayali')), 'muaf': _muaf,
            })
        if not kapasiteler:
            s['durum'] = 'BIRIM YOK'
            s['uretilebilir'] = None
            s['kisitlayan'] = ''
            continue
        en_dusuk = min(kapasiteler)
        s['uretilebilir'] = en_dusuk[0]
        s['kisitlayan'] = en_dusuk[1]
        s['durum'] = 'URETILEBILIR' if en_dusuk[0] > 0 else 'STOK YOK'
        s['eksi_var'] = any(p['eksi_bakiye'] for p in s['parcalar'])

    # ── KARAR: planlamanın launch ihtiyacı + malzeme durumu ──────────────
    # Kullanıcının elle yaptığı iş tam olarak bu: "launch gerekiyor mu?" (plan
    # dosyasındaki Launch sütunu) ile "malzemesi var mı?" (AS400'de ağaç+stok)
    # sorularını birleştirip talimat vermek. İkisi burada birleşir.
    for s in satirlar:
        ihtiyac = s.get('gereken') or 0
        u = s.get('uretilebilir')
        # ERP ağacı planlamanın listesini kapsıyor mu? (kapsamıyorsa ağaç
        # değişmiş olabilir — karar elle doğrulanmalı)
        erp_kodlar = {p['kod'] for p in s.get('parcalar', [])}
        s['agac_farki'] = sorted(set(s.get('plan_parcalar') or []) - erp_kodlar)
        # ── İKİ AYRI SORU (kullanıcı 2026-07-31) ─────────────────────────
        # Tek "karar" sütunu iki soruyu karıştırıyordu. Kullanıcının sorduğu
        # sıra: (1) bu ürünü kaynatmam gerekiyor mu? (2) gerekiyorsa malzemem
        # var mı? Örnek 10.300.1845W: ihtiyaç 2287 − stok 639 − launch 646 =
        # 1002 KAYNATILMALI; ayrıca alt parçası (10.300.1845A) 01D+CF2'de 0
        # olduğu için malzeme YOK. İkisi ayrı bilgi, ayrı sütun.
        s['kaynatilmali'] = ihtiyac > 0
        if not s['kaynatilmali']:
            s['malzeme'] = ''
            s['karar'] = 'GEREK YOK'
        elif u is None:
            s['malzeme'] = 'BILINMIYOR'
            s['karar'] = 'ELLE BAK'
        elif u >= ihtiyac:
            s['malzeme'] = 'TAM'
            s['karar'] = 'TALIMAT VER'
        elif u > 0:
            s['malzeme'] = 'KISMI'
            s['karar'] = 'KISMI'
        else:
            s['malzeme'] = 'YOK'
            s['karar'] = 'MALZEME YOK'
    return satirlar


def excel_yaz(satirlar, cikti):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = openpyxl.Workbook()

    ince = Side(style='thin', color='D0D7E2')
    kenar = Border(left=ince, right=ince, top=ince, bottom=ince)
    bas_fill = PatternFill('solid', fgColor='6D28D9')
    bas_font = Font(bold=True, color='FFFFFF', size=11)

    def basliklar(ws, kolonlar, genislikler):
        ws.append(kolonlar)
        for i, g in enumerate(genislikler, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = g
            c = ws.cell(row=1, column=i)
            c.fill, c.font, c.border = bas_fill, bas_font, kenar
            c.alignment = Alignment(horizontal='center', vertical='center')
        ws.freeze_panes = 'A2'

    KARAR_RENK = {'TALIMAT VER': 'D6EBDD', 'KISMI': 'F7EFDD',
                  'MALZEME YOK': 'F7E6E4', 'ELLE BAK': 'EDEAF7', 'GEREK YOK': 'F3F3F8'}

    def _satir_yaz(ws, s):
        kis_stok = ''
        for p in s.get('parcalar', []):
            if p['kod'] == s.get('kisitlayan'):
                kis_stok = p['stok_sayilan']
                break
        notlar = []
        if s.get('agac_farki'):
            notlar.append('ERP ağacı plandakinden farklı: ' + ', '.join(s['agac_farki'][:3]))
        if s.get('eksi_var'):
            notlar.append('bazı parçalarda EKSİ bakiye (0 sayıldı)')
        if s['durum'] == 'AGAC YOK':
            notlar.append("ERP'de ürün ağacı yok — satın alma parçası olabilir")
        if s['durum'] == 'BIRIM YOK':
            notlar.append('ağaçta birim miktar tanımsız')
        ws.append([s['sira'], s['kaynak_kod'], s['urun'],
                   s.get('iht_6h') or 0, s.get('acik_launch') or 0, s.get('toplam_stok') or 0,
                   s.get('gereken') or 0,
                   s['uretilebilir'] if s['uretilebilir'] is not None else '—',
                   s.get('karar', ''), s.get('kisitlayan', ''), kis_stok,
                   len(s.get('parcalar', [])), ' · '.join(notlar)])
        r = ws.max_row
        f = KARAR_RENK.get(s.get('karar'))
        for c in range(1, 14):
            ws.cell(row=r, column=c).border = kenar
            if f:
                ws.cell(row=r, column=c).fill = PatternFill('solid', fgColor=f)
        ws.cell(row=r, column=7).font = Font(bold=True)
        ws.cell(row=r, column=8).font = Font(bold=True)
        ws.cell(row=r, column=9).font = Font(bold=True)

    BASLIK = ['Sıra', 'Kaynak Kodu', 'Ürün', '6 Hf İhtiyaç (BO)', 'Üretimde (BV)',
              'Toplam Stok', 'GEREKEN', 'Üretilebilir', 'KARAR', 'Kısıtlayan Parça',
              'Kısıtlayan Stok', 'Parça Sayısı', 'Not']
    GENIS = [6, 22, 18, 15, 13, 12, 11, 13, 15, 22, 15, 12, 46]

    # ── 1. SAYFA: yalnız AKSİYON gerekenler (launch ihtiyacı > 0) ──
    ws = wb.active
    ws.title = 'Aksiyon'
    basliklar(ws, BASLIK, GENIS)
    sira_onem = {'TALIMAT VER': 0, 'KISMI': 1, 'MALZEME YOK': 2, 'ELLE BAK': 3}
    aksiyon = [s for s in satirlar if (s.get('gereken') or 0) > 0]
    for s in sorted(aksiyon, key=lambda x: (sira_onem.get(x.get('karar'), 9), x['sira'])):
        _satir_yaz(ws, s)

    # ── 2. SAYFA: planın tamamı, öncelik sırası korunmuş ──
    ws_t = wb.create_sheet('Tüm Plan')
    basliklar(ws_t, BASLIK, GENIS)
    for s in satirlar:
        _satir_yaz(ws_t, s)

    # ── PARÇA KIRILIMI ──
    ws2 = wb.create_sheet('Parça Kırılımı')
    basliklar(ws2, ['Sıra', 'Kaynak Kodu', 'Alt Parça', 'Birim Miktar', 'Birim',
                    '01D', 'CF2', '01D+CF2', 'Bu parçadan kaç adet', 'Kısıt mı?', 'Diğer depolar'],
              [6, 22, 22, 12, 7, 11, 11, 11, 20, 10, 30])
    for s in satirlar:
        for p in s.get('parcalar', []):
            diger = ' '.join(f'{d}:{v:g}' for d, v in p['depolar'].items()
                             if d not in SAYILAN_DEPOLAR)
            ws2.append([s['sira'], s['kaynak_kod'], p['kod'], p['birim'], p['um'],
                        p['depolar'].get('01D', 0), p['depolar'].get('CF2', 0),
                        p['stok_sayilan'],
                        p['kapasite'] if p['kapasite'] is not None else '—',
                        'KISIT' if p['kod'] == s.get('kisitlayan') else '',
                        diger])
            r = ws2.max_row
            for c in range(1, 12):
                ws2.cell(row=r, column=c).border = kenar
            if p['eksi_bakiye']:
                for c in (6, 7, 8):
                    ws2.cell(row=r, column=c).fill = PatternFill('solid', fgColor='F7E6E4')
            if p['kod'] == s.get('kisitlayan'):
                ws2.cell(row=r, column=10).font = Font(bold=True, color='B4231C')

    # ── AÇIKLAMA ──
    ws3 = wb.create_sheet('Nasıl okunur')
    for satir in [
        ['Kaynak plan kontrolü — otomatik üretilebilirlik hesabı'],
        [f'Üretim tarihi: {datetime.now().strftime("%d.%m.%Y %H:%M")}'],
        [],
        ['Bu rapor, AS400 07>10>01 ekranında her kodu tek tek açıp (9+Enter) 1. seviye'],
        ['alt parçaların stoğuna bakma işini otomatikleştirir.'],
        [],
        ['Üretilebilir', 'En kısıtlayan alt parçanın izin verdiği adet:'],
        ['', 'min( alt parça 01D+CF2 stoğu / ürün ağacındaki birim miktar )'],
        ['Stok kuralı', 'Yalnız 01D ve CF2 depoları sayılır (elle yapılan kontrolle aynı).'],
        ['', 'Diğer depolar (01W, REP, MDT...) bilgi olarak gösterilir, hesaba girmez.'],
        ['Eksi bakiye', "ERP'de eksi stok olabiliyor; 0 sayılır ve satır kırmızı işaretlenir."],
        ['AGAC YOK', "ERP'de 1. seviye ürün ağacı bulunamadı — satın alma parçası olabilir,"],
        ['', 'bu satırı elle kontrol edin.'],
        ['Sıra', 'Plan dosyasındaki öncelik sırası korunmuştur.'],
        [],
        ['DİKKAT: hesap yalnız STOK kısıtına bakar. Açık sipariş, rezervasyon, kalite'],
        ['bloğu ve kapasite dikkate alınmaz — launch kararında bunları ayrıca değerlendirin.'],
    ]:
        ws3.append(satir)
    ws3.column_dimensions['A'].width = 16
    ws3.column_dimensions['B'].width = 92
    ws3['A1'].font = Font(bold=True, size=13, color='6D28D9')

    wb.save(cikti)
    return cikti


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--plan', default=VARSAYILAN_PLAN)
    ap.add_argument('--limit', type=int, default=None)
    ap.add_argument('--cikti', default=None)
    a = ap.parse_args()

    if not os.path.exists(a.plan):
        print(f'HATA: plan dosyası bulunamadı: {a.plan}')
        return 1
    print(f'Plan okunuyor: {os.path.basename(a.plan)}')
    satirlar = plan_oku(a.plan, a.limit)
    print(f'  {len(satirlar)} kaynak kodu (öncelik sırasıyla)')

    print('ERP bağlanıyor...')
    cn = erp_baglan()
    try:
        print('  ürün ağaçları çekiliyor...')
        agac = urun_agaci(cn, [s['kaynak_kod'] for s in satirlar])
        alt_kodlar = sorted({alt for lst in agac.values() for alt, _, _ in lst})
        print(f'  {len(agac)} kodun ağacı bulundu · {len(alt_kodlar)} farklı alt parça')
        print('  stoklar çekiliyor...')
        stok = stoklar(cn, alt_kodlar)
        print(f'  {len(stok)} parçanın stok kaydı var')
        # Referansın KENDİ stoğu da ERP'den (plan dosyasındaki sütun haftalık)
        ref_stok = stoklar(cn, [s['kaynak_kod'] for s in satirlar])
    finally:
        cn.close()

    hesapla(satirlar, agac, stok, ref_stok)
    say = {}
    for s in satirlar:
        say[s['karar']] = say.get(s['karar'], 0) + 1
    aksiyon = [s for s in satirlar if (s.get('gereken') or 0) > 0]
    print(f'\n6 haftalık açığı olan: {len(aksiyon)} / {len(satirlar)} satır')
    print('KARAR: ' + ' · '.join(f'{k}: {v}' for k, v in sorted(say.items())))

    ver = [s for s in aksiyon if s['karar'] == 'TALIMAT VER']
    print(f'\n>>> TALİMAT VERİLEBİLİR ({len(ver)}) — malzemesi tam:')
    for s in ver:
        print(f"   {s['sira']:>3}. {s['kaynak_kod']:<20} gereken {s['gereken']:>7.0f} · "
              f"üretilebilir {s['uretilebilir']:>7}")
    kis = [s for s in aksiyon if s['karar'] == 'KISMI']
    if kis:
        print(f'\n>>> KISMİ ({len(kis)}) — malzeme yetmiyor:')
        for s in kis[:12]:
            print(f"   {s['sira']:>3}. {s['kaynak_kod']:<20} gereken {s['gereken']:>7.0f} · "
                  f"üretilebilir {s['uretilebilir']:>7}  (kısıt: {s['kisitlayan']})")
    yok = [s for s in aksiyon if s['karar'] == 'MALZEME YOK']
    if yok:
        print(f'\n>>> MALZEME YOK ({len(yok)}):')
        for s in yok[:12]:
            print(f"   {s['sira']:>3}. {s['kaynak_kod']:<20} gereken {s['gereken']:>7.0f} · "
                  f"kısıt: {s['kisitlayan']}")
    fark = [s for s in aksiyon if s.get('agac_farki')]
    if fark:
        print(f'\n>>> ERP AĞACI PLANDAN FARKLI ({len(fark)}) — kararı elle doğrulayın:')
        for s in fark[:8]:
            print(f"   {s['sira']:>3}. {s['kaynak_kod']:<20} planda var, ERP ağacında yok: "
                  f"{', '.join(s['agac_farki'][:3])}")

    cikti = a.cikti or os.path.join(
        os.path.expanduser('~'), 'OneDrive', 'Masaüstü',
        f'Kaynak_Uretilebilirlik_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx')
    try:
        excel_yaz(satirlar, cikti)
        print(f'\nRapor: {cikti}')
    except PermissionError:
        print(f'\nHATA: dosya açık olabilir → {cikti}')
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
