# -*- coding: utf-8 -*-
"""
as400_teyit_import.py — ÜRETİM EMRİ TEYİDİNİ (rientro) staging tablosuyla yazar.
Cofle S.p.A. IT · Simone Rota, 2026-09-17 maili "Cofle Forge - Warehouse movements".

NEDEN: teyit bugün PCOMM 5250 EKRAN ROBOTUYLA veriliyor (oturum kilidi, ECL37110
takılması, kurtarma .bat'ları hep bunun yüzünden). Bu tablo ile teyit de CFI gibi
DOĞRUDAN TABLOYA YAZIM olur; ekran yok, robot yok.

SÖZLEŞME (Simone'nin maili + AS400 kolon açıklamaları):
  Kütüphane COFLEFORGE, tablo BPRCF0I ("Consunt. ord. prod. Testata" kopyası).
  BİZİM DOLDURDUKLARIMIZ
    J0RED1  yüzyıl (20)        ⎫ üretim emri kaydı — XPRO90'daki
    J0RED2  yıl    (26)        ⎬ Q0RED1/Q0RED2/Q0RENU ile birebir aynı
    J0RENU  numara (200385)    ⎭
    J0QTRI  teyit edilen adet
    J0FLSA  'A' | 'S'  → DİKKAT: kolon açıklaması "Flag Acconto/saldo".
            A = ARA teyit (acconto), S = KAPANIŞ teyidi (saldo). Scrap DEĞİLDİR.
            Yanlış 'S' emri erken kapatır; varsayılanımız bu yüzden 'A'.
    J0STE2  bizim anahtarımız (15 karakter) — BMMAF0I'daki MGSTE2'nin karşılığı
  PROGRAMIN GERİ YAZDIKLARI
    J0CNPR  teyit sıra no · J0MGSS/J0MGAA/J0MGNU/J0MGPG depo hareketi
    J0TIME  damga · J0STAT 1=OK 2=HATA 3=UYARI · J0NOTE açıklama
  J0NOTE kodları: 01E hatalı emir · 02W maliyet eksik (UYARI, teyit YAPILDI) ·
                  03E emir işlenemedi · 04E hatalı adet · 05E yalnız A/S kabul edilir

ANAHTAR: 'T' + yymmddHHMMSS + 2 haneli sayaç = 15 karakter. CFI tarafı 'F' ile
başlar; baş harf ayrı olduğu için hangi satırın teyitten geldiği tek bakışta belli
(deneme satırını elle düzeltmek gerekirse J0STE2 LIKE 'T%' yeter).

TEST ORTAMI (2026-09-17): import programı COFLETKPR'de koşuyor. Tablo tek olduğu
için yazdığımız satırın hangi ortamda işlendiğini Simone'ye doğrulatıyoruz; ilk
deneme bu yüzden 'A' (ara teyit) ve küçük adetle yapılır, satır J0STE2 ile geri
bulunur.

DOĞRULANDI (2026-09-21): sahte emir 26-999999 ile yazılan satır ~2 SANİYEDE işlendi;
J0STAT=2, J0NOTE='[01E-Incorrect order]'. Program tablo üzerinde canlı çalışıyor ve
bizim satırlarımızı görüyor. J0TIME İTALYA SAATİNDE damgalanır (TR'den 1 saat geri) —
gecikmeyi J0TIME ile yerel saati çıkararak hesaplama.

ENGEL (2026-09-21): CANLI açık emir 26-202269 (XPRO90'da var, durum 40, kalan 25) de
'01E-Incorrect order' aldı; durum 45'teki 26-202046 da aynı cevabı verdi (emir durumu
kuralı DEĞİL). Program canlı üretim emirlerini GÖRMÜYOR — test ortamına
(COFLETKPR) bakıyor. Canlıya almadan önce İtalya'nın programı canlı ortama yöneltmesi
gerekiyor; o güne kadar teyit PCOMM ekran robotuyla verilmeye devam eder.

ODBC: as400_config ile aynı bağlantı; INSERT yetkisi yalnız COFLEFORGE'da.
"""
import re
import threading
import time
from datetime import datetime, timedelta

import as400_config as CFG

KUTUPHANE = 'COFLEFORGE'
TABLO = 'BPRCF0I'
REFERANS_KOLONU = 'J0STE2'        # bizim kayit anahtarimiz (CFI'daki MGSTE2'nin esi)
ANAHTAR_ONEK = 'T'                # teyit; CFI 'F' kullanir
ZORUNLU = ('J0RED1', 'J0RED2', 'J0RENU', 'J0QTRI', 'J0FLSA')
DURUM_KOLONLARI = ('J0STAT', 'J0NOTE', 'J0TIME', 'J0CNPR',
                   'J0MGSS', 'J0MGAA', 'J0MGNU', 'J0MGPG')
MUKERRER_PENCERE_DK = 30          # islenmis ayni satir YALNIZ bu kadar yeniyse engel
FLSA_DEGERLERI = ('A', 'S')

# J0NOTE kodlari (Simone'nin listesi) — panelde Turkce gosterilir
NOT_KODLARI = {
    '01E': 'Hatalı üretim emri (emir numarası ERP’de bulunamadı ya da uygun değil)',
    '02W': 'Maliyet eksik — teyit YAPILDI, uyarı olarak işaretlendi',
    '03E': 'Emir işlenemedi',
    '04E': 'Hatalı adet',
    '05E': 'Yalnız A/S kabul edilir (J0FLSA)',
}

_AD = re.compile(r'^[A-Z][A-Z0-9_]{0,9}$')
_KILIT = threading.Lock()
_KOLON_ONBELLEK = {}
_ANAHTAR_SAYAC = [0]


def anahtar_uret(onek=ANAHTAR_ONEK):
    """J0STE2 anahtarı: 'T' + yymmddHHMMSS + 2 haneli sayaç → 15 karakter."""
    _ANAHTAR_SAYAC[0] = (_ANAHTAR_SAYAC[0] + 1) % 100
    return f"{onek[:1]}{datetime.now().strftime('%y%m%d%H%M%S')}{_ANAHTAR_SAYAC[0]:02d}"


def _ad(x, ne):
    x = str(x or '').strip().upper()
    if not _AD.match(x):
        raise ValueError(f'geçersiz {ne} adı: {x!r}')
    return x


def baglan(timeout=20, kullanici=None):
    return CFG.baglan(timeout=timeout, kullanici=kullanici)


def kolonlar(kutuphane=KUTUPHANE, tablo=TABLO, cn=None, tazele=False, kullanici=None):
    """Tablonun kolon adları (önbellekli) — beklenen alanlar var mı diye bakarız."""
    k, t = _ad(kutuphane, 'kütüphane'), _ad(tablo, 'tablo')
    if not tazele and (k, t) in _KOLON_ONBELLEK:
        return _KOLON_ONBELLEK[(k, t)]
    kapat = cn is None
    cn = cn or baglan(kullanici=kullanici)
    try:
        rows = cn.cursor().execute(
            "SELECT COLUMN_NAME, DATA_TYPE, LENGTH FROM QSYS2.SYSCOLUMNS "
            "WHERE TABLE_SCHEMA=? AND TABLE_NAME=? ORDER BY ORDINAL_POSITION", (k, t)).fetchall()
        out = [{'ad': str(r[0]).strip(), 'tip': str(r[1]).strip(), 'uzunluk': r[2]} for r in rows]
        _KOLON_ONBELLEK[(k, t)] = out
        return out
    finally:
        if kapat:
            cn.close()


def emir_parcala(emir):
    """'26-200385' | '20-26-200385' | (20, 26, 200385) → (yüzyıl, yıl, numara).

    XPRO90'da emir üç parçadır: Q0RED1 yüzyıl (20), Q0RED2 yıl (26), Q0RENU numara.
    Ekranda "Ord.No. 26 200385" diye yalnız yıl+numara görünür; yüzyıl verilmezse 20."""
    if isinstance(emir, (tuple, list)) and len(emir) == 3:
        p = [str(x).strip() for x in emir]
    else:
        p = [x for x in re.split(r'[-/ ]+', str(emir or '').strip()) if x]
        if len(p) == 2:
            p = ['20'] + p
    if len(p) != 3 or not all(x.isdigit() for x in p):
        raise ValueError(f'geçersiz emir numarası: {emir!r} (örn. 26-200385)')
    d1, d2, nu = int(p[0]), int(p[1]), int(p[2])
    if not (0 <= d1 <= 99 and 0 <= d2 <= 99 and 0 < nu <= 9999999):
        raise ValueError(f'emir numarası aralık dışı: {emir!r}')
    return d1, d2, nu


def not_cevir(notu):
    """J0NOTE metnini Türkçeleştirir: '[02W-Missing cost]' → kod + açıklama."""
    n = str(notu or '').strip()
    if not n:
        return ''
    kod = re.search(r'\[(\d{2}[EW])', n)
    if kod and kod.group(1) in NOT_KODLARI:
        return f'{kod.group(1)} — {NOT_KODLARI[kod.group(1)]} ({n})'
    return n


def _satir_oku(cn, k, t, rrn):
    mevcut = {c['ad'] for c in kolonlar(k, t, cn=cn)}
    secim = [c for c in DURUM_KOLONLARI if c in mevcut]
    r = cn.cursor().execute(
        f"SELECT {', '.join(secim)} FROM {k}.{t} x WHERE RRN(x)=?", (int(rrn),)).fetchone()
    if not r:
        return None
    d = {secim[i].lower(): ('' if r[i] is None else str(r[i]).strip()) for i in range(len(secim))}
    hareket = '-'.join(x for x in (d.get('j0mgss'), d.get('j0mgaa'), d.get('j0mgnu'),
                                   d.get('j0mgpg')) if x)
    return {'j0stat': d.get('j0stat', ''), 'not': not_cevir(d.get('j0note', '')),
            'ham_not': d.get('j0note', ''), 'j0time': d.get('j0time', ''),
            'teyit_sira': d.get('j0cnpr', ''), 'hareket_no': hareket}


def teyit_yaz(emir, adet, flsa='A', referans=None, kutuphane=KUTUPHANE, tablo=TABLO,
              bekleme_sn=60, yokla_sn=3, zorla=False, cn=None, sadece_yaz=False,
              kullanici=None, ekler=None):
    """Bir üretim emri teyidi yazar ve programın işlemesini bekler.

    emir : '26-200385' (yıl-numara) ya da (20, 26, 200385)
    adet : teyit edilecek adet (J0QTRI)
    flsa : 'A' ara teyit (VARSAYILAN) · 'S' kapanış teyidi — 'S' emri KAPATIR
    ekler: isteğe bağlı kolonlar {'J0ARTI': referans kodu, 'J0CRCD': rientro
           neden kodu, 'J0MGPR': ana depo, 'J0COMM': iş emri} — Simone'nin
           listesinde yoklar ama tabloda varlar; '01E' alırsak sırayla denenir

    Döner: {ok, durum, teyit_sira, hareket_no, not, rrn, anahtar, alanlar, sql}
      islendi     : J0STAT='1'
      uyarili     : J0STAT='3' → teyit YAPILDI ama uyarı var (bugün: maliyet eksik)
      reddedildi  : J0STAT='2' → J0NOTE'ta sebep
      zaman_asimi : bekleme_sn içinde J0STAT boş (program durmuş olabilir); satır
                    tabloda DURUR, rrn ile sonra bakılır (teyit_durum)
      mevcut      : aynı emir/adet için bekleyen ya da son MUKERRER_PENCERE_DK
                    dakikada bizim yazdığımız işlenmiş satır var (zorla=False)
    """
    k, t = _ad(kutuphane, 'kütüphane'), _ad(tablo, 'tablo')
    try:
        d1, d2, nu = emir_parcala(emir)
    except ValueError as e:
        return {'ok': False, 'durum': 'hata', 'not': str(e)}
    try:
        adet_f = float(adet)
    except (TypeError, ValueError):
        adet_f = 0.0
    if adet_f <= 0 or adet_f > 9999999:
        return {'ok': False, 'durum': 'hata', 'not': f'geçersiz adet: {adet!r}'}
    flsa = str(flsa or 'A').strip().upper()[:1]
    if flsa not in FLSA_DEGERLERI:
        return {'ok': False, 'durum': 'hata',
                'not': f"geçersiz J0FLSA: {flsa!r} — 'A' (ara teyit) ya da 'S' (kapanış)"}

    kapat = cn is None
    cn = cn or baglan(kullanici=kullanici)
    try:
        mevcut_kolonlar = {c['ad'] for c in kolonlar(k, t, cn=cn)}
        eksik = [c for c in ZORUNLU + ('J0STAT',) if c not in mevcut_kolonlar]
        if eksik:
            return {'ok': False, 'durum': 'hata',
                    'not': f'{k}.{t} beklenen kolonları taşımıyor: {eksik}'}
        alanlar = {'J0RED1': d1, 'J0RED2': d2, 'J0RENU': nu, 'J0QTRI': adet_f, 'J0FLSA': flsa}
        ref = str(referans or '').strip() or anahtar_uret()
        if REFERANS_KOLONU in mevcut_kolonlar:
            uz = next((c.get('uzunluk') for c in kolonlar(k, t, cn=cn)
                       if c['ad'] == REFERANS_KOLONU), None)
            try:
                uz = int(uz or 0)
            except (TypeError, ValueError):
                uz = 0
            alanlar[REFERANS_KOLONU] = ref[:uz] if uz else ref

        for _k, _v in (ekler or {}).items():
            _k = str(_k).strip().upper()
            if _v is None or str(_v).strip() == '':
                continue
            if _k not in mevcut_kolonlar:
                return {'ok': False, 'durum': 'hata',
                        'not': f'{k}.{t} içinde {_k} kolonu yok'}
            alanlar[_k] = str(_v).strip() if isinstance(_v, str) else _v

        with _KILIT:
            cur = cn.cursor()
            # MÜKERRER FRENİ (CFI tarafındaki kuralın aynısı): bekleyen satır her
            # zaman engeller; işlenmiş satır yalnız BİZİM anahtarımızla son
            # MUKERRER_PENCERE_DK dakikada yazıldıysa engeller. Aksi hâlde aynı
            # emre ertesi gün verilen gerçek teyit "mevcut" diye reddedilirdi.
            if not zorla:
                kosul = ("J0RED1=? AND J0RED2=? AND J0RENU=? AND J0QTRI=? "
                         "AND (J0STAT='' OR J0STAT IS NULL")
                par = [d1, d2, nu, adet_f]
                if REFERANS_KOLONU in mevcut_kolonlar:
                    esik = ANAHTAR_ONEK + (datetime.now() - timedelta(minutes=MUKERRER_PENCERE_DK)
                                           ).strftime('%y%m%d%H%M%S')
                    kosul += (f" OR (J0STAT IN ('1','3') AND {REFERANS_KOLONU} LIKE "
                              f"'{ANAHTAR_ONEK}%' AND {REFERANS_KOLONU} >= ?))")
                    par.append(esik)
                else:
                    kosul += " OR J0STAT IN ('1','3'))"
                var = cur.execute(f"SELECT RRN(x), J0STAT FROM {k}.{t} x WHERE {kosul} "
                                  f"ORDER BY RRN(x) DESC FETCH FIRST 1 ROW ONLY", par).fetchone()
                if var:
                    d = _satir_oku(cn, k, t, int(var[0])) or {}
                    return {'ok': d.get('j0stat') in ('1', '3'), 'durum': 'mevcut',
                            'rrn': int(var[0]), 'alanlar': alanlar, **d,
                            'mesaj': ('aynı teyit zaten işlenmiş' if d.get('j0stat') in ('1', '3')
                                      else 'aynı teyit kuyrukta bekliyor')}
            kol = ', '.join(alanlar)
            yer = ', '.join('?' * len(alanlar))
            sql = f"INSERT INTO {k}.{t} ({kol}) VALUES ({yer})"
            cur.execute(sql, list(alanlar.values()))
            kosul = ' AND '.join(f'{c}=?' for c in alanlar)
            r = cur.execute(f"SELECT RRN(x) FROM {k}.{t} x WHERE {kosul} "
                            f"ORDER BY RRN(x) DESC FETCH FIRST 1 ROW ONLY",
                            list(alanlar.values())).fetchone()
            rrn = int(r[0]) if r else None
        anahtar = alanlar.get(REFERANS_KOLONU, '')
        if rrn is None:
            return {'ok': False, 'durum': 'hata', 'not': 'INSERT sonrası satır geri bulunamadı',
                    'sql': sql, 'alanlar': alanlar, 'anahtar': anahtar}
        if sadece_yaz:
            return {'ok': True, 'durum': 'yazildi', 'rrn': rrn, 'sql': sql,
                    'alanlar': alanlar, 'anahtar': anahtar}
        son = teyit_durum(rrn, k, t, cn=cn, bekleme_sn=bekleme_sn, yokla_sn=yokla_sn)
        son.update({'sql': sql, 'alanlar': alanlar, 'anahtar': anahtar})
        return son
    finally:
        if kapat:
            cn.close()


def teyit_durum(rrn, kutuphane=KUTUPHANE, tablo=TABLO, cn=None, bekleme_sn=0, yokla_sn=3,
                kullanici=None):
    """RRN ile satırın durumunu okur; bekleme_sn>0 ise J0STAT dolana kadar yoklar.
    J0STAT 3 (UYARI) BAŞARIDIR: teyit yapılmıştır, not'ta sebebi yazar."""
    k, t = _ad(kutuphane, 'kütüphane'), _ad(tablo, 'tablo')
    kapat = cn is None
    cn = cn or baglan(kullanici=kullanici)
    try:
        bitis = time.time() + max(0, float(bekleme_sn))
        while True:
            d = _satir_oku(cn, k, t, int(rrn))
            if d is None:
                return {'ok': False, 'durum': 'hata', 'rrn': rrn, 'not': f'RRN {rrn} bulunamadı'}
            if d['j0stat'] == '1':
                return {'ok': True, 'durum': 'islendi', 'rrn': rrn, **d}
            if d['j0stat'] == '3':
                return {'ok': True, 'durum': 'uyarili', 'rrn': rrn, **d}
            if d['j0stat'] == '2':
                return {'ok': False, 'durum': 'reddedildi', 'rrn': rrn, **d}
            if time.time() >= bitis:
                return {'ok': False, 'durum': 'zaman_asimi', 'rrn': rrn, **d,
                        'not': d['not'] or (f'{int(bekleme_sn)} sn içinde işlenmedi '
                                            f'(J0STAT boş) — program çalışıyor mu?')}
            time.sleep(max(1.0, float(yokla_sn)))
    finally:
        if kapat:
            cn.close()


def son_satirlar(n=20, kutuphane=KUTUPHANE, tablo=TABLO, cn=None, kullanici=None,
                 yalniz_bizim=False):
    """Panel/CLI için: tablodaki son n satır (ne yazdık, program ne dedi).
    yalniz_bizim=True → sadece bizim anahtarımızla yazılanlar (J0STE2 LIKE 'T%')."""
    k, t = _ad(kutuphane, 'kütüphane'), _ad(tablo, 'tablo')
    kapat = cn is None
    cn = cn or baglan(kullanici=kullanici)
    try:
        mevcut = {c['ad'] for c in kolonlar(k, t, cn=cn)}
        secim = [c for c in ('J0RED1', 'J0RED2', 'J0RENU', 'J0QTRI', 'J0FLSA', REFERANS_KOLONU,
                             'J0CNPR', 'J0STAT', 'J0NOTE', 'J0TIME', 'J0MGSS', 'J0MGAA',
                             'J0MGNU', 'J0MGPG') if c in mevcut]
        nere = (f"WHERE {REFERANS_KOLONU} LIKE '{ANAHTAR_ONEK}%' "
                if yalniz_bizim and REFERANS_KOLONU in mevcut else '')
        rows = cn.cursor().execute(
            f"SELECT RRN(x), {', '.join(secim)} FROM {k}.{t} x {nere}"
            f"ORDER BY RRN(x) DESC FETCH FIRST {int(n)} ROWS ONLY").fetchall()
        out = []
        for r in rows:
            d = {'rrn': int(r[0])}
            for i, c in enumerate(secim, start=1):
                d[c.lower()] = '' if r[i] is None else str(r[i]).strip()
            d['emir'] = '-'.join(x for x in (d.get('j0red2'), d.get('j0renu')) if x)
            d['hareket_no'] = '-'.join(x for x in (d.get('j0mgss'), d.get('j0mgaa'),
                                                   d.get('j0mgnu'), d.get('j0mgpg')) if x)
            d['not'] = not_cevir(d.get('j0note', ''))
            out.append(d)
        return out
    finally:
        if kapat:
            cn.close()


def acik_emirler(article=None, n=20, cn=None, kullanici=None):
    """Teyit denemesi için açık üretim emirleri (XPRO90). Emir no + kalan adet.
    Teyit edilecek emri elle aramak yerine buradan seçilir."""
    kapat = cn is None
    cn = cn or baglan(kullanici=kullanici)
    try:
        sql = ("SELECT Q0RED1, Q0RED2, Q0RENU, Q0ARTI, Q0AVAN, Q0QTOR, Q0QTRI "
               "FROM TKC0301F.XPRO90 WHERE Q0AVAN IN ('40','45','50')")
        par = []
        if article:
            sql += " AND Q0ARTI = ?"
            par.append(str(article).strip())
        sql += f" ORDER BY Q0RENU DESC FETCH FIRST {int(n)} ROWS ONLY"
        out = []
        for r in cn.cursor().execute(sql, par).fetchall():
            d1, d2, nu = int(r[0] or 0), int(r[1] or 0), int(r[2] or 0)
            adet, teyit = float(r[5] or 0), float(r[6] or 0)
            out.append({'emir': f'{d2:02d}-{nu}', 'parcali': (d1, d2, nu),
                        'article': str(r[3] or '').strip(), 'durum': str(r[4] or '').strip(),
                        'adet': adet, 'teyit': teyit, 'kalan': max(0.0, adet - teyit)})
        return out
    finally:
        if kapat:
            cn.close()
