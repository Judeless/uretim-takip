# -*- coding: utf-8 -*-
"""
as400_import.py — Depo hareketini AS/400'e EKRAN ROBOTU YERİNE staging
tablosuyla yazar (Cofle S.p.A. IT, Simone Rota, 2026-09-07).

SÖZLEŞME (IT'nin maili):
  · Kütüphane COFLEFORGE, tablo BMMAF0I (BMMAF0'ın "input" kopyası).
  · Zorunlu: MGARCD (kod) · MGCACD (causal: CFI / COP) · MGMGCD (depo) ·
    MGMCCD (karşı depo; causal isterse) · MGQTA (adet).
  · İsteğe bağlı: MGDSSO/MGDAAO/MGDMMO/MGDGGO (hareket tarihi: yüzyıl 20,
    yıl 26, ay, gün — BMMAF0'da okuduğumuz biçimin aynısı). Verilmezse bugün.
  · Program 10 sn'de bir MGSTAT BOŞ satırları alır:
        MGSTAT='1' işlendi → MGSERE/MGANRE/MGNURE/MGPRRE hareket numarası
        MGSTAT='2' hata    → MGNOTE ("[01-Item code error]" gibi)
    Her iki hâlde MGTIME işlem damgası.
  · ŞİMDİLİK TEST VERİTABANI: hareket COFLETKPR'ye düşer, üretim BMMAF0'da
    GÖRÜNMEZ. Doğrulama = MGSTAT + hareket no.

SATIRI GERİ BULMA (Simone 2026-09-08): kendi kaydımız **MGSTE2** alanına
yazılır. Alan 15 KARAKTER (saha 2026-09-08: 'TEST-260908093640-…' →
'TEST-2609081036' kesildi) — kod içeren bir anahtar sığmaz, o yüzden anahtar
BURADA üretilir: 'F' + yymmddHHMMSS + 2 haneli sayaç = tam 15 karakter,
süreç içinde benzersiz (kilit + sayaç). MGSTE2 satırı geri bulmak ve izlemek
içindir; MÜKERRER FRENİ yalnız kod/causal/adet/tarih bileşimine bakar
(anahtar her denemede yeni olduğu için ona bakılsaydı fren çalışmazdı).
Tabloda MGSTE2 yoksa eski yol: alanlar + RRN.

ODBC: mevcut okuma bağlantısıyla aynı (as400_config), tek fark INSERT yetkisi
— COFLEFORGE kütüphanesine IT verdi. Başka hiçbir tabloya yazılmaz.
"""
import re
import threading
import time
from datetime import date, datetime

import as400_config as CFG

KUTUPHANE = 'COFLEFORGE'
TABLO = 'BMMAF0I'
TARIH_KOLONLARI = ('MGDSSO', 'MGDAAO', 'MGDMMO', 'MGDGGO')
REFERANS_KOLONU = 'MGSTE2'       # bizim kayit anahtarimiz (Simone Rota, 2026-09-08)
DURUM_KOLONLARI = ('MGSTAT', 'MGNOTE', 'MGTIME', 'MGSERE', 'MGANRE', 'MGNURE', 'MGPRRE')
_AD = re.compile(r'^[A-Z][A-Z0-9_]{0,9}$')       # kütüphane/tablo/kolon adı — SQL'e ham giriyor
_KILIT = threading.Lock()
_KOLON_ONBELLEK = {}
_ANAHTAR_SAYAC = [0]


def anahtar_uret(onek='F'):
    """MGSTE2 anahtarı: 'F' + yymmddHHMMSS + 2 haneli sayaç → 15 karakter."""
    _ANAHTAR_SAYAC[0] = (_ANAHTAR_SAYAC[0] + 1) % 100
    return f"{onek[:1]}{datetime.now().strftime('%y%m%d%H%M%S')}{_ANAHTAR_SAYAC[0]:02d}"


def _ad(x, ne):
    x = str(x or '').strip().upper()
    if not _AD.match(x):
        raise ValueError(f'geçersiz {ne} adı: {x!r}')
    return x


def baglan(timeout=20):
    """Okuma bağlantısıyla aynı kimlik; autocommit (tek satırlık INSERT)."""
    import pyodbc
    pw = CFG.sifre_al()
    if not pw:
        raise RuntimeError('AS400 şifresi kasada yok')
    return pyodbc.connect(CFG.baglanti_dizesi(pw), timeout=timeout, autocommit=True)


def kolonlar(kutuphane=KUTUPHANE, tablo=TABLO, cn=None, tazele=False):
    """QSYS2.SYSCOLUMNS'tan kolon listesi: [{ad, tip, uzunluk, ondalik, aciklama}].
    Salt okunur; bir kez çekilip önbelleğe alınır (INSERT'te hangi isteğe bağlı
    kolonların VAR olduğunu bilmek için)."""
    k, t = _ad(kutuphane, 'kütüphane'), _ad(tablo, 'tablo')
    if not tazele and (k, t) in _KOLON_ONBELLEK:
        return _KOLON_ONBELLEK[(k, t)]
    kapat = cn is None
    cn = cn or baglan()
    try:
        rows = cn.cursor().execute(
            "SELECT COLUMN_NAME, DATA_TYPE, LENGTH, NUMERIC_SCALE, COLUMN_TEXT, ORDINAL_POSITION "
            "FROM QSYS2.SYSCOLUMNS WHERE TABLE_SCHEMA=? AND TABLE_NAME=? ORDER BY ORDINAL_POSITION",
            (k, t)).fetchall()
    finally:
        if kapat:
            cn.close()
    out = [{'ad': str(r[0]).strip(), 'tip': str(r[1]).strip(), 'uzunluk': r[2],
            'ondalik': r[3], 'aciklama': (str(r[4]).strip() if r[4] else '')} for r in rows]
    _KOLON_ONBELLEK[(k, t)] = out
    return out


def _tarih_parcala(uretim_tarihi):
    """'2026-09-08' → (20, 26, 9, 8). Boş/bozuksa None (program bugünü alır)."""
    if not uretim_tarihi:
        return None
    try:
        d = datetime.strptime(str(uretim_tarihi)[:10], '%Y-%m-%d').date()
    except ValueError:
        return None
    return d.year // 100, d.year % 100, d.month, d.day


def _hareket_no(r):
    """MGSERE/MGANRE/MGNURE/MGPRRE → ekrandaki gibi 'YY/NNN r.P'.
    COFLETKPR 'Visualizzazione movimento' (2026-09-08): Nr Registrazione = 26 / 34,
    Riga attuale = 10 → MGANRE=26 (yıl), MGNURE=34 (kayıt no), MGPRRE=10 (satır);
    MGSERE=20 yüzyıl, gösterilmez."""
    sere, anre, nure, prre = [(str(x).strip() if x is not None else '') for x in r]
    def sayi(x):
        try:
            return str(int(float(x)))
        except (TypeError, ValueError):
            return x
    if not nure or sayi(nure) == '0':
        return ''
    return f"{sayi(anre)}/{sayi(nure)}" + (f" r.{sayi(prre)}" if prre and sayi(prre) != '0' else '')


def _satir_oku(cn, k, t, rrn):
    cur = cn.cursor()
    sql = (f"SELECT MGSTAT, MGNOTE, MGSERE, MGANRE, MGNURE, MGPRRE, MGTIME "
           f"FROM {k}.{t} x WHERE RRN(x)=?")
    r = cur.execute(sql, (rrn,)).fetchone()
    if not r:
        return None
    return {'mgstat': (str(r[0]).strip() if r[0] is not None else ''),
            'not': (str(r[1]).strip() if r[1] is not None else ''),
            'hareket_no': _hareket_no(r[2:6]),
            'hareket_alanlari': {'MGSERE': r[2], 'MGANRE': r[3], 'MGNURE': r[4], 'MGPRRE': r[5]},
            'mgtime': (str(r[6]).strip() if r[6] is not None else '')}


def hareket_yaz(article, adet, causal='CFI', wh='01D', cp='01D', uretim_tarihi=None,
                referans=None, kutuphane=KUTUPHANE, tablo=TABLO, bekleme_sn=60,
                yokla_sn=3, zorla=False, cn=None, sadece_yaz=False):
    """Bir depo hareketi satırı yazar ve programın işlemesini bekler.

    Döner: {ok, durum: 'islendi'|'reddedildi'|'zaman_asimi'|'mevcut'|'hata',
            hareket_no, not, rrn, mgtime, sql, alanlar}
      · islendi     : MGSTAT='1', hareket_no dolu
      · reddedildi  : MGSTAT='2', not = MGNOTE
      · zaman_asimi : bekleme_sn içinde MGSTAT hâlâ boş (program durmuş olabilir);
                      satır tabloda DURUR, rrn ile sonra bakılır (hareket_durum)
      · mevcut      : aynı kod/causal/adet/tarih için bekleyen ya da işlenmiş
                      satır zaten var (zorla=False) — tekrar YAZILMADI
    sadece_yaz=True → bekleme yok (rrn döner)."""
    k, t = _ad(kutuphane, 'kütüphane'), _ad(tablo, 'tablo')
    article = str(article or '').strip()
    causal = str(causal or 'CFI').strip().upper()
    wh = str(wh or '').strip().upper()
    cp = str(cp or '').strip().upper()
    try:
        adet_f = float(adet)
    except (TypeError, ValueError):
        adet_f = 0.0
    if not article or len(article) > 21 or not article.isascii():
        return {'ok': False, 'durum': 'hata', 'not': f'geçersiz article: {article!r}'}
    if adet_f <= 0 or adet_f > 99999:
        return {'ok': False, 'durum': 'hata', 'not': f'geçersiz adet: {adet!r}'}
    if causal not in ('CFI', 'COP') or not wh:
        return {'ok': False, 'durum': 'hata', 'not': f'geçersiz causal/depo: {causal}/{wh}'}

    tarih = _tarih_parcala(uretim_tarihi)
    kapat = cn is None
    cn = cn or baglan()
    try:
        mevcut_kolonlar = {c['ad'] for c in kolonlar(k, t, cn=cn)}
        eksik = [c for c in ('MGARCD', 'MGCACD', 'MGMGCD', 'MGMCCD', 'MGQTA', 'MGSTAT') if c not in mevcut_kolonlar]
        if eksik:
            return {'ok': False, 'durum': 'hata', 'not': f'{k}.{t} beklenen kolonları taşımıyor: {eksik}'}
        alanlar = {'MGARCD': article, 'MGCACD': causal, 'MGMGCD': wh, 'MGMCCD': cp, 'MGQTA': adet_f}
        if tarih and all(c in mevcut_kolonlar for c in TARIH_KOLONLARI):
            alanlar.update(dict(zip(TARIH_KOLONLARI, tarih)))
        # Kendi kaydımız (MGSTE2): satırı geri bulmak + izleme. Verilmezse 15
        # karakterlik anahtar üretilir; kolon uzunluğunu aşarsa kırpılır.
        ref = str(referans or '').strip()
        if not ref and REFERANS_KOLONU in mevcut_kolonlar:
            ref = anahtar_uret()
        if ref and REFERANS_KOLONU in mevcut_kolonlar:
            uz = next((c.get('uzunluk') for c in kolonlar(k, t, cn=cn) if c['ad'] == REFERANS_KOLONU), None)
            try:
                uz = int(uz or 0)
            except (TypeError, ValueError):
                uz = 0
            alanlar[REFERANS_KOLONU] = ref[:uz] if uz else ref

        with _KILIT:
            cur = cn.cursor()
            # MÜKERRER FRENİ: aynı kod/causal/adet(/tarih) için bekleyen ('') ya da
            # işlenmiş ('1') satır varsa yeniden yazma — ağ kopması sonrası tekrar
            # denemede ERP'ye ikinci hareket girmesin. Reddedilmiş ('2') satır
            # engel değil: düzeltilip tekrar denenebilir.
            if not zorla:
                kosul = "MGARCD=? AND MGCACD=? AND MGQTA=? AND (MGSTAT='1' OR MGSTAT='' OR MGSTAT IS NULL)"
                par = [article, causal, adet_f]
                if 'MGDGGO' in alanlar:
                    kosul += " AND MGDSSO=? AND MGDAAO=? AND MGDMMO=? AND MGDGGO=?"
                    par += list(tarih)
                # MGSTE2 koşula GİRMEZ: anahtar her denemede yeni; bileşik alanlar
                # aynı olduğu sürece bu aynı gönderimdir ve yeniden YAZILMAZ.
                var = cur.execute(f"SELECT RRN(x), MGSTAT FROM {k}.{t} x WHERE {kosul} "
                                  f"ORDER BY RRN(x) DESC FETCH FIRST 1 ROW ONLY", par).fetchone()
                if var:
                    d = _satir_oku(cn, k, t, int(var[0])) or {}
                    return {'ok': (d.get('mgstat') == '1'), 'durum': 'mevcut', 'rrn': int(var[0]),
                            'hareket_no': d.get('hareket_no', ''), 'not': d.get('not', ''),
                            'mgtime': d.get('mgtime', ''), 'alanlar': alanlar,
                            'mesaj': ('aynı satır zaten işlenmiş' if d.get('mgstat') == '1'
                                      else 'aynı satır kuyrukta bekliyor')}
            kol = ', '.join(alanlar)
            yer = ', '.join('?' * len(alanlar))
            sql = f"INSERT INTO {k}.{t} ({kol}) VALUES ({yer})"
            cur.execute(sql, list(alanlar.values()))
            # Satırımızı geri bul: aynı değerlerle EN SON RRN (kilit altındayız)
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
            return {'ok': True, 'durum': 'yazildi', 'rrn': rrn, 'sql': sql, 'alanlar': alanlar, 'anahtar': anahtar}
        son = hareket_durum(rrn, k, t, cn=cn, bekleme_sn=bekleme_sn, yokla_sn=yokla_sn)
        son.update({'sql': sql, 'alanlar': alanlar, 'anahtar': anahtar})
        return son
    finally:
        if kapat:
            cn.close()


def hareket_durum(rrn, kutuphane=KUTUPHANE, tablo=TABLO, cn=None, bekleme_sn=0, yokla_sn=3):
    """RRN ile satırın durumunu okur; bekleme_sn>0 ise MGSTAT dolana kadar yoklar."""
    k, t = _ad(kutuphane, 'kütüphane'), _ad(tablo, 'tablo')
    kapat = cn is None
    cn = cn or baglan()
    try:
        bitis = time.time() + max(0, float(bekleme_sn))
        while True:
            d = _satir_oku(cn, k, t, int(rrn))
            if d is None:
                return {'ok': False, 'durum': 'hata', 'rrn': rrn, 'not': f'RRN {rrn} bulunamadı'}
            if d['mgstat'] == '1':
                return {'ok': True, 'durum': 'islendi', 'rrn': rrn, **d}
            if d['mgstat'] == '2':
                return {'ok': False, 'durum': 'reddedildi', 'rrn': rrn, **d}
            if time.time() >= bitis:
                return {'ok': False, 'durum': 'zaman_asimi', 'rrn': rrn, **d,
                        'not': d['not'] or f'{int(bekleme_sn)} sn içinde işlenmedi (MGSTAT boş) — program çalışıyor mu?'}
            time.sleep(max(1.0, float(yokla_sn)))
    finally:
        if kapat:
            cn.close()


def son_satirlar(n=20, kutuphane=KUTUPHANE, tablo=TABLO, cn=None):
    """Panel için: tablodaki son n satır (ne yazdık, program ne dedi)."""
    k, t = _ad(kutuphane, 'kütüphane'), _ad(tablo, 'tablo')
    kapat = cn is None
    cn = cn or baglan()
    try:
        mevcut = {c['ad'] for c in kolonlar(k, t, cn=cn)}
        secim = [c for c in ('MGARCD', 'MGCACD', 'MGMGCD', 'MGMCCD', 'MGQTA', 'MGDSSO', 'MGDAAO',
                             'MGDMMO', 'MGDGGO', REFERANS_KOLONU, 'MGSTAT', 'MGNOTE', 'MGSERE',
                             'MGANRE', 'MGNURE', 'MGPRRE', 'MGTIME') if c in mevcut]
        rows = cn.cursor().execute(
            f"SELECT RRN(x), {', '.join(secim)} FROM {k}.{t} x ORDER BY RRN(x) DESC "
            f"FETCH FIRST {int(n)} ROWS ONLY").fetchall()
        out = []
        for r in rows:
            d = {'rrn': int(r[0])}
            for i, c in enumerate(secim, 1):
                v = r[i]
                d[c] = (str(v).strip() if isinstance(v, str) else v)
            out.append(d)
        return out
    finally:
        if kapat:
            cn.close()
