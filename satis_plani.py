# -*- coding: utf-8 -*-
"""
SATIŞ PLANI ARŞİVİ + SİPARİŞ FARKLARI (kullanıcı 2026-10-05, yeniden 2026-10-09)

KAYNAK — DÜZELTİLDİ (2026-10-09)
AS400 10-05-03-06 "Orders plan" F6 ile çalıştırılınca sonuç **TKC0301F.XWPVF0**'a
(satırlar, XWQUAL='C') ve **TKC0301F.XWPVF1**'e (hafta başlıkları, XWQUA1='C') yazılır.
Planlamanın 'satış planı.xls' şablonundaki ODBC sorgusu birebir:
    SELECT XWQUAL … XWUBI2 FROM S650B9F2.tkc0301F.XWPVF0 WHERE XWQUAL='C'
'S650B9F2' AS400'ün SİSTEM (RDB) adıdır — üç parçalı adın katalog kısmı. İlk sürüm
bunu bir dosya adı sanıp 'S650B9' tablosunu aradı ve her denemede SQL0204 (bulunamadı)
aldı; modül hiç çekim yapamadı.

XWPVF0 ORTAK BİR ÇALIŞMA DOSYASIDIR: kim 10-05-03-06'yı hangi parametreyle çalıştırırsa
içerik odur. Kullanıcının çektiği iki dosya bunu gösterdi: 05.10 çekimi yalnız TK2 (01D)
kodları (402 satır), 09.10 çekimi TK1 + TK2 + satınalma (1.576 satır). Bu yüzden her
çekimin KAPSAMI (Anaveri / Forge'a göre TK1-TK2 kod oranı) saklanır ve kapsamı farklı
iki çekim kıyaslanırken yalnız ORTAK kapsam (TK2) karşılaştırılır.

DOSYADAN YÜKLEME: kullanıcının kendi çektiği 'satış planı.xls' (şablonun 'Finale'
sayfası) ya da .xlsx sürümü tarihiyle yüklenir. .xls için sunucuda xlrd gerekir.

SİPARİŞ FARKLARI — PLANLAMANIN MANTIĞI (Q:\\UretimPlanlama\\EMRE\\Yeni klasör\\Siparişler\\
Bakiye-Sipariş farkı\\Sipariş farkları GG.AA-GG.AA-GG.AA.xlsx): ürün bazında (müşteriler
toplanır) her çekimin Stok · Bakiye (gecikmiş) · takvim haftasına HİZALI haftalık talebi;
  'İLK 3 HAFTA İÇİNE EKLENEN MİKTAR' = Σ(yeni: bakiye + ilk 3 hafta)
                                       − Σ(eski: bakiye + aynı bitiş haftasına kadar)
  (eski çekimde geçmiş haftalar eski bakiyeye katılır — hafta dönünce kıyas kaymaz)
  'FARK' = toplam(yeni) − toplam(eski), ortak son haftaya kadar.
Pozitif = yakın haftalara giren talep (araya giren sipariş); negatif = sevk, iptal ya
da erteleme (satış planı hangisi olduğunu söylemez — stok değişimi ipucu verir).
"""
import os
import re
import json
import hashlib
from datetime import datetime

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
KLASOR = os.path.join(PROJECT_DIR, 'data', 'satis_plani')
KAYNAK_TABLO = 'TKC0301F.XWPVF0'
BASLIK_TABLO = 'TKC0301F.XWPVF1'
QUAL = 'C'                      # şablonun süzgeci (XWQUAL / XWQUA1)
SORGU_ZAMAN_ASIMI = 45          # sn — tek sorgu
GUN_SAKLA = 400                 # veritabanında tutulan gün (dosyalar silinmez)
YAKIN_HAFTA = 2                 # müşteri bazlı farkta "yakın termin" = gecikmiş + ilk 2 hafta
TK1_ESIK = 0.05                 # TK1 kod payı bunun altındaysa çekim 'yalnız TK2' sayılır

# XWPVF0 alanları — şablondaki sorguyla AYNI sıra (XWQT01 = gecikmiş, 02..13 = 12 hafta)
SUTUNLAR = (['XWQUAL', 'XWCFCD', 'XWDECD', 'XWRGSC', 'XWCCC2', 'XWARCD', 'XWPRTE', 'XWPROV',
             'XWC1CD', 'XWC2CD', 'XWPNCD', 'XWCLCO', 'XWUMMA', 'XWQTGI']
            + ['XWQT%02d' % i for i in range(1, 14)]
            + ['XWQTOT', 'XWTART', 'XWMGUC', 'XWUBI1', 'XWUBI2'])
BASLIK = (['C', 'CLIENTE', 'DEST.', 'RAG.SOC.', 'AM/OEM C', 'CD ART.', 'PRIORIT.', 'PROVEN.',
           'CLASSE', 'MARCA', 'PUSHPULL', 'AM/OEM A', 'U.M.', 'GIACENZA', 'ARRETR.'])
_HAFTA_RE = re.compile(r'^\s*(\d{1,2})\s*/\s*(\d{4})\s*$')
_TR = str.maketrans('"£$§', 'ÜÖİŞ')     # EBCDIC(1026)→ODBC bozulması (app._plan_siparisler ile aynı)


def tablolari_kur(conn):
    conn.execute("""CREATE TABLE IF NOT EXISTS satis_plani_cekim (
        id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, tarih TEXT, imza TEXT,
        satir INTEGER, kod INTEGER, haftalar TEXT, depo TEXT, kutuphane TEXT,
        dosya TEXT, kullanici TEXT)""")
    kolonlar = {r[1] for r in conn.execute("PRAGMA table_info(satis_plani_cekim)")}
    for ad, tip in (('kaynak', "TEXT DEFAULT 'as400'"), ('kapsam', "TEXT DEFAULT ''")):
        if ad not in kolonlar:
            conn.execute(f"ALTER TABLE satis_plani_cekim ADD COLUMN {ad} {tip}")
    conn.execute("""CREATE TABLE IF NOT EXISTS satis_plani_satir (
        cekim_id INTEGER, musteri_kodu TEXT, dest TEXT, musteri TEXT, article TEXT,
        prov TEXT, priorit TEXT, stok REAL, gecikmis REAL, haftalar TEXT, toplam REAL,
        tipo TEXT, depo TEXT, ham TEXT)""")
    conn.execute("CREATE INDEX IF NOT EXISTS ix_sps_cekim ON satis_plani_satir(cekim_id)")
    conn.commit()


def _t(v):
    return '' if v is None else str(v).strip()


def _f(v):
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def _kod_metin(v):
    """Müşteri kodu sayı gelebilir (10201.0) → '10201'."""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return _t(v)


def _haftalar_baslik(satir):
    """XWPVF1 / şablon başlık satırından hafta etiketleri ('41/2026'), soldan sağa."""
    out = []
    for v in (satir or ()):
        m = _HAFTA_RE.match(_t(v))
        if m:
            out.append(f'{int(m.group(1)):02d}/{m.group(2)}')
    return out


def _hafta_sira(etiket):
    """'40/2026' → 202640 (sıralama/kıyas için)."""
    try:
        h, y = etiket.split('/')
        return int(y) * 100 + int(h)
    except Exception:
        return 0


def _satirlara(veri, basliklar):
    """XWPVF0 sırasındaki ham satırlar (liste) → satır sözlükleri."""
    if len(basliklar) != 12:
        basliklar = [f'H{i}' for i in range(1, 13)]
    satirlar = []
    for r in veri:
        r = list(r) + [None] * max(0, 32 - len(r))
        if _t(r[0]).upper() != QUAL:
            continue
        art = _t(r[5])
        if not art:
            continue
        haftalar = {basliklar[i]: _f(r[15 + i]) for i in range(12)}
        satirlar.append({
            'qual': _t(r[0]), 'musteri_kodu': _kod_metin(r[1]), 'dest': _kod_metin(r[2]),
            'musteri': _t(r[3]).translate(_TR), 'article': art, 'prov': _t(r[7]).upper(),
            'priorit': _kod_metin(r[6]), 'stok': _f(r[13]), 'gecikmis': _f(r[14]), 'haftalar': haftalar,
            'toplam': _f(r[27]), 'tipo': _t(r[28]), 'depo': _t(r[29]),
            'ham': [_kod_metin(x) if isinstance(x, float) and i in (1, 2) else
                    (x.strip() if isinstance(x, str) else x) for i, x in enumerate(r[:32])],
        })
    return satirlar, basliklar


def degisim_zamani(cur):
    """XWPVF0'ın son yazılma anı (= son F6) ve satır sayısı — AS400 kataloğundan."""
    try:
        cur.execute("SELECT LAST_CHANGE_TIMESTAMP, NUMBER_ROWS FROM QSYS2.SYSPARTITIONSTAT "
                    "WHERE TABLE_SCHEMA='TKC0301F' AND TABLE_NAME='XWPVF0'")
        r = cur.fetchone()
        if r:
            return {'degisim': r[0].strftime('%Y-%m-%d %H:%M:%S') if hasattr(r[0], 'strftime') else (str(r[0]) if r[0] else None),
                    'satir': int(r[1]) if r[1] is not None else None}
    except Exception as e:
        print(f'[SATIŞ PLANI] SYSPARTITIONSTAT okunamadı: {e}')
    return {'degisim': None, 'satir': None}


def as400_oku(baglan=None):
    """XWPVF0 + XWPVF1 → (satırlar, hafta etiketleri, kaynak, [konum]). Yalnız OKUR."""
    if baglan is None:
        import sys
        sys.path.insert(0, os.path.join(PROJECT_DIR, 'as400'))
        import as400_config as CFG
        baglan = CFG.baglan
    cn = baglan(timeout=30)
    try:
        try:
            cn.timeout = SORGU_ZAMAN_ASIMI      # F6 işi dosyayı tutuyorsa sonsuz beklemesin
        except Exception:
            pass
        cur = cn.cursor()
        konum = degisim_zamani(cur)
        veri, hata = None, ''
        # WITH UR: yazan işin satır kilidini bekleme (salt okuma). Sürüm bu eki tanımazsa
        # (SQL0104) eksiz — şablondaki sorguyla aynı — denenir.
        for ek in (' WITH UR', ''):
            try:
                cur.execute(f"SELECT {', '.join(SUTUNLAR)} FROM {KAYNAK_TABLO} WHERE XWQUAL=?{ek}", (QUAL,))
                veri = cur.fetchall()
                break
            except Exception as e:
                hata = str(e)
                if ek and 'SQL0104' in hata:
                    continue
                break
        if veri is None:
            raise RuntimeError(f'{KAYNAK_TABLO} okunamadı — {hata[:300]}')
        basliklar = []
        for ek in (' WITH UR', ''):
            try:
                cur.execute(f"SELECT * FROM {BASLIK_TABLO} WHERE XWQUA1=?{ek}", (QUAL,))
                basliklar = _haftalar_baslik(cur.fetchone())
                break
            except Exception:
                continue
    finally:
        try:
            cn.close()
        except Exception:
            pass
    satirlar, basliklar = _satirlara(veri, basliklar)
    return satirlar, basliklar, KAYNAK_TABLO, [dict(konum, kutuphane=KAYNAK_TABLO)]


# ── DOSYADAN (kullanıcının kendi çektiği satış planı.xls / .xlsx) ───────────
def _sayfalar(yol):
    """{sayfa adı: satır listesi} — .xlsx/.xlsm openpyxl, .xls xlrd (sunucuda kurulu olmalı)."""
    if yol.lower().endswith('.xls'):
        try:
            import xlrd
        except ImportError:
            raise ValueError("Sunucuda .xls okuyucu (xlrd) kurulu değil — dosyayı Excel'de "
                             "'Farklı Kaydet → .xlsx' ile kaydedip yükleyin ya da sunucuda "
                             "'pip install xlrd==2.0.1' çalıştırın")
        kitap = xlrd.open_workbook(yol, on_demand=True)
        try:
            return {ad: [kitap.sheet_by_name(ad).row_values(i) for i in range(kitap.sheet_by_name(ad).nrows)]
                    for ad in kitap.sheet_names()}
        finally:
            kitap.release_resources()
    from openpyxl import load_workbook
    wb = load_workbook(yol, read_only=True, data_only=True)
    try:
        return {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}
    finally:
        wb.close()


def dosyadan_oku(yol):
    """Şablonun 'Finale' sayfası (yoksa 'CD ART.' başlıklı ilk sayfa) → (satırlar, haftalar).
    Başlık satırı: F sütunu 'CD ART.'; hafta etiketleri aynı satırda ('  41/2026')."""
    sayfalar = _sayfalar(yol)
    sira = sorted(sayfalar, key=lambda a: (a.strip().lower() != 'finale', a))
    for ad in sira:
        satirlar = sayfalar[ad]
        for i, r in enumerate(satirlar[:60]):
            if len(r) > 5 and _t(r[5]).upper() == 'CD ART.' and _t(r[0]).upper() == QUAL:
                basliklar = _haftalar_baslik(r)
                veri = [x for x in satirlar[i + 1:] if x and _t(x[0]).upper() == QUAL and len(x) > 5 and _t(x[5])]
                if veri:
                    return _satirlara(veri, basliklar)
    raise ValueError("Dosyada satış planı bulunamadı — 'Finale' sayfasında 'C' · 'CD ART.' "
                     "başlık satırı ve altında satırlar olmalı (planlamanın 'satış planı.xls' şablonu)")


# ── KAPSAM (bu çekim yalnız TK2 mi, TK1+TK2 mi?) ───────────────────────────
def tesis_haritasi(conn):
    """{KOD: 'TK1'|'TK2'} — önce planlamanın Anaveri'si (TK-1/2), yoksa Forge referans listesi
    (tek tesiste tanımlıysa), 93.* (tel) TK1."""
    import anaveri as AV
    out = {}
    for r in conn.execute("SELECT UPPER(REPLACE(referans_kodu,' ','')), GROUP_CONCAT(DISTINCT COALESCE(lokasyon,'TK2')) "
                          "FROM referans_listesi GROUP BY 1"):
        if r[1] in ('TK1', 'TK2'):
            out[r[0]] = r[1]
    for k, v in AV.haritasi(conn).items():
        t = AV.TESIS.get((v.get('tk') or '').upper())
        if t:
            out[k] = t
    return out


def tesis_bul(kod, harita):
    k = re.sub(r'\s', '', kod or '').upper()
    if k in harita:
        return harita[k]
    if k.startswith('93.'):
        return 'TK1'
    return ''


def kapsam(conn, satirlar):
    """{'etiket': 'TK2'|'TK1+TK2'|'?', 'tk1', 'tk2', 'diger'} — kod sayıları."""
    h = tesis_haritasi(conn)
    say = {'TK1': 0, 'TK2': 0, '': 0}
    for k in {s['article'].upper() for s in satirlar}:
        say[tesis_bul(k, h)] += 1
    t = say['TK1'] + say['TK2']
    etiket = '?' if not t else ('TK2' if say['TK1'] / t < TK1_ESIK else
                                ('TK1' if say['TK2'] / t < TK1_ESIK else 'TK1+TK2'))
    return {'etiket': etiket, 'tk1': say['TK1'], 'tk2': say['TK2'], 'diger': say['']}


def _imza(satirlar, basliklar):
    ozet = json.dumps([basliklar] + sorted(
        [s['musteri_kodu'], s['dest'], s['article'], s['gecikmis'], sorted(s['haftalar'].items()), s['toplam'], s['stok']]
        for s in satirlar), ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.sha256(ozet.encode('utf-8')).hexdigest()


def excel_yaz(satirlar, basliklar, ts, depo, yol, kaynak_metni='AS400 10-05-03-06 (XWPVF0)'):
    """Günün dosyası: şablonun 'Finale' düzeni (C … TOTALE, TIPO ART, MAG., UBIC1, UBIC2)."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    ws = wb.active
    ws.title = 'Satış Planı'
    ws.append([f'SATIŞ PLANI — {kaynak_metni} · çekim {ts} · depo {depo or "?"} · {len(satirlar)} satır'])
    ws['A1'].font = Font(bold=True, size=12)
    ws.append([])
    ws.append(BASLIK + basliklar + ['TOTALE', 'TIPO ART', 'MAG.', 'UBIC1', 'UBIC2'])
    for c in ws[3]:
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor='1F3864')
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for s in satirlar:
        h = list(s['ham']) + [None] * max(0, 32 - len(s['ham']))
        ws.append([h[0], h[1], h[2], s['musteri'], h[4], s['article'], h[6], s['prov'], h[8], h[9],
                   h[10], h[11], h[12], s['stok'], s['gecikmis']] + [s['haftalar'].get(b, 0) for b in basliklar]
                  + [s['toplam'], s['tipo'], s['depo'], h[30], h[31]])
    for i, w in enumerate([4, 9, 8, 34, 8, 18, 7, 7, 7, 7, 8, 8, 5, 9, 9] + [8] * len(basliklar) + [9, 6, 6, 7, 7], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'G4'
    ws.auto_filter.ref = f'A3:{get_column_letter(ws.max_column)}{ws.max_row}'
    os.makedirs(os.path.dirname(yol), exist_ok=True)
    gecici = yol + '.tmp'
    wb.save(gecici)
    try:
        os.replace(gecici, yol)
    except PermissionError:
        os.remove(gecici)
        raise


def kaydet(conn, satirlar, basliklar, kullanici, kaynak='as400', tarih=None, kutuphane='', zorla=False):
    """Çekimi sakla + günün Excel'ini yaz. tarih verilmezse bugün. Aynı günün önceki
    çekimlerinin SATIRLARI silinir (gün = son hâli). İçerik son çekimle aynıysa (ve zorla
    değilse) hiçbir şey yazılmaz → {'degisti': False, ...}."""
    tablolari_kur(conn)
    if not satirlar:
        raise ValueError('Satış planında satır yok')
    imza = _imza(satirlar, basliklar)
    depolar = sorted({s['depo'] for s in satirlar if s['depo']})
    depo = ','.join(depolar)
    kod = len({s['article'].upper() for s in satirlar})
    simdi = datetime.now()
    tarih = tarih or simdi.strftime('%Y-%m-%d')
    son = conn.execute("SELECT id, imza, ts, tarih FROM satis_plani_cekim WHERE tarih=? ORDER BY id DESC LIMIT 1"
                       if kaynak == 'dosya' else
                       "SELECT id, imza, ts, tarih FROM satis_plani_cekim ORDER BY id DESC LIMIT 1",
                       (tarih,) if kaynak == 'dosya' else ()).fetchone()
    if son and son[1] == imza and not zorla:
        return {'degisti': False, 'cekim_id': son[0], 'satir': len(satirlar), 'kod': kod,
                'son_degisim': son[2], 'depo': depo, 'haftalar': basliklar, 'kutuphane': kutuphane}
    ks = kapsam(conn, satirlar)
    ts = f"{tarih} {simdi.strftime('%H:%M:%S')}"
    ek = '_dosya' if kaynak == 'dosya' else ''
    dosya = os.path.join(KLASOR, f'Satis_Plani_{tarih}{ek}.xlsx')
    metin = 'kullanıcının yüklediği dosya' if kaynak == 'dosya' else 'AS400 10-05-03-06 (XWPVF0)'
    try:
        excel_yaz(satirlar, basliklar, ts, depo, dosya, metin)
    except PermissionError:
        # Günün dosyası biri tarafından açık/kilitli — çekim KAYBOLMASIN, saatli adla yaz
        dosya = os.path.join(KLASOR, f'Satis_Plani_{tarih}{ek}_{simdi.strftime("%H%M")}.xlsx')
        excel_yaz(satirlar, basliklar, ts, depo, dosya, metin)
    cur = conn.execute(
        "INSERT INTO satis_plani_cekim (ts, tarih, imza, satir, kod, haftalar, depo, kutuphane, dosya, kullanici, "
        "kaynak, kapsam) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
        (ts, tarih, imza, len(satirlar), kod, json.dumps(basliklar), depo, kutuphane,
         os.path.basename(dosya), kullanici, kaynak, json.dumps(ks, ensure_ascii=False)))
    cid = cur.lastrowid
    conn.executemany(
        "INSERT INTO satis_plani_satir (cekim_id, musteri_kodu, dest, musteri, article, prov, priorit, "
        "stok, gecikmis, haftalar, toplam, tipo, depo, ham) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        [(cid, s['musteri_kodu'], s['dest'], s['musteri'], s['article'], s['prov'], s['priorit'],
          s['stok'], s['gecikmis'], json.dumps(s['haftalar']), s['toplam'], s['tipo'], s['depo'],
          json.dumps(s['ham'], ensure_ascii=False, default=str)) for s in satirlar])
    eski = [r[0] for r in conn.execute(
        "SELECT id FROM satis_plani_cekim WHERE tarih=? AND id<>?", (tarih, cid)).fetchall()]
    for e in eski:
        conn.execute("DELETE FROM satis_plani_satir WHERE cekim_id=?", (e,))
    sinir = conn.execute("SELECT date('now','localtime', ?)", (f'-{GUN_SAKLA} day',)).fetchone()[0]
    for e in conn.execute("SELECT id FROM satis_plani_cekim WHERE tarih < ?", (sinir,)).fetchall():
        conn.execute("DELETE FROM satis_plani_satir WHERE cekim_id=?", (e[0],))
    conn.commit()
    print(f'[SATIŞ PLANI] yeni çekim #{cid} ({kaynak}, {tarih}): {len(satirlar)} satır, {kod} kod, kapsam {ks["etiket"]}')
    return {'degisti': True, 'cekim_id': cid, 'satir': len(satirlar), 'kod': kod, 'tarih': tarih,
            'ts': ts, 'dosya': os.path.basename(dosya), 'depo': depo, 'haftalar': basliklar,
            'kutuphane': kutuphane, 'kapsam': ks}


def cek(conn, baglan=None, kullanici='otomatik', zorla=False):
    """AS400'den oku; içerik son çekimden FARKLIYSA sakla + günün Excel'ini yaz."""
    tablolari_kur(conn)
    satirlar, basliklar, kutuphane, konumlar = as400_oku(baglan)
    as400 = {'as400_degisim': (konumlar[0] if konumlar else {}).get('degisim'), 'kopyalar': []}
    return {**kaydet(conn, satirlar, basliklar, kullanici, 'as400', None, kutuphane, zorla), **as400}


def yukle(conn, yol, tarih, kullanici):
    """Kullanıcının çektiği dosyayı o günün çekimi olarak sakla."""
    if not re.match(r'^\d{4}-\d{2}-\d{2}$', str(tarih or '')):
        raise ValueError('Tarih YYYY-AA-GG olmalı')
    satirlar, basliklar = dosyadan_oku(yol)
    return kaydet(conn, satirlar, basliklar, kullanici, 'dosya', tarih, os.path.basename(yol), zorla=True)


def cekim_sil(conn, cekim_id):
    """Yanlış yüklenen çekimi siler (satırları + kaydı). Excel dosyası diskte kalır."""
    tablolari_kur(conn)
    n = conn.execute("DELETE FROM satis_plani_cekim WHERE id=?", (cekim_id,)).rowcount
    conn.execute("DELETE FROM satis_plani_satir WHERE cekim_id=?", (cekim_id,))
    conn.commit()
    return n


# ── ARKA PLAN ÇEKİMİ + SON DENEME (2026-10-05) ───────────────────────────────
# OLAY: panelde "Şimdi kontrol et" → "Unexpected token '<' … not valid JSON". İstek
# Cloudflare üzerinden gidiyor; AS400 okuması 100 sn'yi aşınca vekil isteği kesip HTML
# hata sayfası döndürüyor. Artık okuma ARKA PLANDA koşar, uç hemen döner, panel durumu
# sorar. Son denemenin sonucu (hata metni dahil) panelde görünür — sunucu logu
# açmadan teşhis için (otomatik 30 dk'lık denemeler de buraya yazar).
import threading as _threading
import time as _time
_DENEME = {'calisiyor': False, 'basladi': None, 'bitti': None, 'kim': '', 'hata': '',
           'sonuc': None, 'sure_sn': None}
_DENEME_KILIT = _threading.Lock()


def deneme_durumu():
    return dict(_DENEME)


def cek_kayitli(conn_ac, kullanici='otomatik', baglan=None):
    """cek() + sonucu _DENEME'ye yaz. Aynı anda ikinci çekim başlamaz."""
    if not _DENEME_KILIT.acquire(blocking=False):
        return {'calisiyor': True}
    t0 = _time.time()
    _DENEME.update(calisiyor=True, basladi=datetime.now().strftime('%Y-%m-%d %H:%M:%S'), bitti=None,
                   kim=kullanici, hata='', sonuc=None, sure_sn=None)
    conn = None
    try:
        conn = conn_ac()
        _DENEME['sonuc'] = cek(conn, baglan=baglan, kullanici=kullanici)
    except Exception as e:
        _DENEME['hata'] = f'{type(e).__name__}: {e}'[:2000]
        print(f'[SATIŞ PLANI] çekim başarısız ({kullanici}): {e}')
    finally:
        try:
            if conn is not None:
                conn.close()
        except Exception:
            pass
        _DENEME.update(calisiyor=False, bitti=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                       sure_sn=round(_time.time() - t0, 1))
        _DENEME_KILIT.release()
    return deneme_durumu()


def cek_arka_planda(conn_ac, kullanici, baglan=None):
    """Arka planda başlatır → True; zaten çalışıyorsa False."""
    if _DENEME['calisiyor']:
        return False
    # İşaret thread başlamadan konur: panel hemen sorarsa ESKİ denemenin sonucunu
    # yeni sonuç sanmasın
    _DENEME.update(calisiyor=True, basladi=datetime.now().strftime('%Y-%m-%d %H:%M:%S'), bitti=None,
                   kim=kullanici, hata='', sonuc=None, sure_sn=None)
    _threading.Thread(target=cek_kayitli, args=(conn_ac, kullanici, baglan), daemon=True,
                      name='satis-plani-cek').start()
    return True


# ── OKUMA ────────────────────────────────────────────────────────────────────
def gunluk_cekimler(conn, limit=120):
    """Her günün SON çekimi (satırları olanlar), yeniden eskiye."""
    tablolari_kur(conn)
    out = []
    for r in conn.execute(
            "SELECT c.id, c.ts, c.tarih, c.satir, c.kod, c.depo, c.dosya, c.haftalar, c.kaynak, c.kapsam, c.kullanici "
            "FROM satis_plani_cekim c "
            "WHERE c.id = (SELECT MAX(id) FROM satis_plani_cekim c2 WHERE c2.tarih = c.tarih) "
            "ORDER BY c.tarih DESC LIMIT ?", (limit,)).fetchall():
        g = dict(zip(('id', 'ts', 'tarih', 'satir', 'kod', 'depo', 'dosya', 'haftalar', 'kaynak', 'kapsam',
                      'kullanici'), r))
        try:
            g['kapsam'] = json.loads(g['kapsam'] or '{}')
        except ValueError:
            g['kapsam'] = {}
        g['kaynak'] = g['kaynak'] or 'as400'
        out.append(g)
    return out


def _satirlar(conn, cekim_id):
    out = []
    for r in conn.execute("SELECT musteri_kodu, dest, musteri, article, prov, priorit, stok, gecikmis, "
                          "haftalar, toplam FROM satis_plani_satir WHERE cekim_id=?", (cekim_id,)).fetchall():
        out.append({'musteri_kodu': r[0], 'dest': r[1], 'musteri': r[2], 'article': r[3], 'prov': r[4],
                    'priorit': r[5], 'stok': r[6] or 0, 'gecikmis': r[7] or 0,
                    'haftalar': json.loads(r[8] or '{}'), 'toplam': r[9] or 0})
    return out


def _kovalar(s, ilk_hafta):
    """Gecikmiş + ilk_hafta'dan önceki haftalar → 'GEC'; sonrası etiketiyle.
    Hafta dönünce dünün ilk haftası bugünün gecikmişine düşer — kıyas kaymasın."""
    k = {'GEC': s['gecikmis']}
    for etiket, adet in s['haftalar'].items():
        if _hafta_sira(etiket) and ilk_hafta and _hafta_sira(etiket) < ilk_hafta:
            k['GEC'] += adet
        else:
            k[etiket] = k.get(etiket, 0) + adet
    return k


def fark(conn, yeni_id, eski_id):
    """İki çekim arasında müşteri × dest × kod bazında talep değişimi.
    tur: 'yeni' (müşteride yeni kod) · 'one_cekildi' (yakın talep arttı, toplam artmadı)
         · 'yakin_artis' (yakın haftalara yeni talep — araya giren sipariş)
         · 'kalkti' (tüm talep kalktı: sevk ya da iptal) · 'ertelendi' (yakından ileriye)
         · 'azaldi' (toplam azaldı: sevk ya da iptal) · 'artti' (ileri haftalara ek talep)"""
    yeni, eski = _satirlar(conn, yeni_id), _satirlar(conn, eski_id)
    hb = json.loads((conn.execute("SELECT haftalar FROM satis_plani_cekim WHERE id=?", (yeni_id,)).fetchone() or ['[]'])[0] or '[]')
    ilk = _hafta_sira(hb[0]) if hb else 0
    yakin = ['GEC'] + hb[:YAKIN_HAFTA]
    anahtar = lambda s: (s['musteri_kodu'], s['dest'], s['article'].upper())    # noqa: E731
    y = {anahtar(s): s for s in yeni}
    e = {anahtar(s): s for s in eski}
    out = []
    for k in set(y) | set(e):
        sy, se = y.get(k), e.get(k)
        ky = _kovalar(sy, ilk) if sy else {}
        ke = _kovalar(se, ilk) if se else {}
        top_y, top_e = sum(ky.values()), sum(ke.values())
        yak_y = sum(ky.get(h, 0) for h in yakin)
        yak_e = sum(ke.get(h, 0) for h in yakin)
        if abs(top_y - top_e) < 0.001 and all(abs(ky.get(h, 0) - ke.get(h, 0)) < 0.001 for h in set(ky) | set(ke)):
            continue
        d_top, d_yak = top_y - top_e, yak_y - yak_e
        if not se:
            tur = 'yeni'
        elif not sy:
            tur = 'kalkti'
        elif d_yak > 0.001 and d_top <= d_yak * 0.5:
            tur = 'one_cekildi'
        elif d_yak > 0.001:
            tur = 'yakin_artis'
        elif d_yak < -0.001 and d_top > -0.001:
            tur = 'ertelendi'
        elif d_top < -0.001:
            tur = 'azaldi'
        else:
            tur = 'artti'
        s = sy or se
        haftalar = [h for h in (['GEC'] + hb)]
        out.append({'tur': tur, 'musteri': s['musteri'], 'musteri_kodu': s['musteri_kodu'], 'dest': s['dest'],
                    'article': s['article'], 'prov': s['prov'],
                    'onceki_toplam': round(top_e, 2), 'yeni_toplam': round(top_y, 2),
                    'onceki_yakin': round(yak_e, 2), 'yeni_yakin': round(yak_y, 2),
                    'degisim': [{'hafta': 'Gecikmiş' if h == 'GEC' else h, 'onceki': round(ke.get(h, 0), 2),
                                 'yeni': round(ky.get(h, 0), 2)}
                                for h in haftalar if abs(ky.get(h, 0) - ke.get(h, 0)) > 0.001]})
    sira = {'yakin_artis': 0, 'one_cekildi': 1, 'yeni': 2, 'kalkti': 3, 'ertelendi': 4, 'azaldi': 5, 'artti': 6}
    out.sort(key=lambda x: (sira.get(x['tur'], 9), -abs(x['yeni_yakin'] - x['onceki_yakin']), x['article']))
    return out


# ── SİPARİŞ FARKLARI — ÜRÜN BAZINDA (planlamanın 'Sipariş farkları' dosyası) ──
def _cekim_bilgi(conn, cid):
    r = conn.execute("SELECT id, ts, tarih, haftalar, kapsam, kaynak FROM satis_plani_cekim WHERE id=?", (cid,)).fetchone()
    if not r:
        return None
    try:
        ks = json.loads(r[4] or '{}')
    except ValueError:
        ks = {}
    return {'id': r[0], 'ts': r[1], 'tarih': r[2], 'haftalar': json.loads(r[3] or '[]'), 'kapsam': ks,
            'kaynak': r[5] or 'as400'}


def _urun_topla(satirlar):
    """Müşteriler toplanır. Stok (GIACENZA) her müşteri satırında AYNI değer — toplanmaz."""
    u = {}
    for s in satirlar:
        k = re.sub(r'\s', '', s['article']).upper()
        d = u.setdefault(k, {'kod': s['article'].strip(), 'prov': s['prov'], 'stok': s['stok'], 'gec': 0.0,
                             'h': {}, 'musteri': {}})
        d['gec'] += s['gecikmis']
        for h, a in s['haftalar'].items():
            d['h'][h] = d['h'].get(h, 0) + a
        m = d['musteri'].setdefault((s['musteri_kodu'], s['dest']), {'musteri': s['musteri'], 'gec': 0.0, 'h': {}})
        m['gec'] += s['gecikmis']
        for h, a in s['haftalar'].items():
            m['h'][h] = m['h'].get(h, 0) + a
    return u


def _hizala(d, ilk, son, ufuk):
    """(bakiye*, yakın, toplam) — bakiye*: gecikmiş + ilk haftadan ÖNCEKİ haftalar (geçmiş),
    yakın: bakiye* + ufuk haftaları, toplam: bakiye* + son haftaya kadar."""
    if not d:
        return 0.0, 0.0, 0.0
    bak = d['gec'] + sum(a for h, a in d['h'].items() if _hafta_sira(h) and _hafta_sira(h) < ilk)
    yak = bak + sum(d['h'].get(h, 0) for h in ufuk)
    top = bak + sum(a for h, a in d['h'].items() if _hafta_sira(h) and ilk <= _hafta_sira(h) <= son)
    return bak, yak, top


def urun_farki(conn, cekim_idler, ilk_n=3, kapsam_secimi='oto', yalniz_p=True, ara=''):
    """cekim_idler: [yeni, eski] ya da [yeni, orta, eski] (planlamanın üçlü dosyası).
    kapsam_secimi: 'oto' (kapsamlar farklıysa yalnız ortak tesis) · 'TK2' · 'TK1' · 'hepsi'
    · 'ortak' (iki çekimde de olan kodlar). → {'cekimler', 'haftalar', 'satirlar', 'ozet', 'notlar'}"""
    ck = [_cekim_bilgi(conn, c) for c in cekim_idler]
    if any(c is None for c in ck) or len(ck) < 2:
        raise ValueError('Karşılaştırılacak çekim bulunamadı')
    ilk_n = max(1, min(12, int(ilk_n or 3)))
    notlar = []
    etiketler = [c['kapsam'].get('etiket', '?') for c in ck]
    tesis = None
    if kapsam_secimi in ('TK1', 'TK2'):
        tesis = kapsam_secimi
    elif kapsam_secimi == 'oto' and len(set(etiketler)) > 1:
        ortak = [e for e in ('TK2', 'TK1') if all(x in (e, 'TK1+TK2') for x in etiketler)]
        tesis = ortak[0] if ortak else None
        notlar.append('Çekimlerin kapsamı farklı (' + ' · '.join(f"{c['tarih']}: {e}" for c, e in zip(ck, etiketler)) +
                      ')' + (f' — yalnız {tesis} kodları karşılaştırıldı.' if tesis else
                             ' — ortak tesis bulunamadı, tümü karşılaştırıldı; farkların bir kısmı kapsamdan.'))
    harita = tesis_haritasi(conn) if tesis else {}
    toplam = [_urun_topla(_satirlar(conn, c['id'])) for c in ck]
    yeni_hb = ck[0]['haftalar'] or []
    ilk = _hafta_sira(yeni_hb[0]) if yeni_hb else 0
    ufuk = yeni_hb[:ilk_n]
    son = min(_hafta_sira(c['haftalar'][-1]) for c in ck if c['haftalar']) if all(c['haftalar'] for c in ck) else 0
    kodlar = set().union(*[set(t) for t in toplam])
    if kapsam_secimi == 'ortak':
        kodlar = set.intersection(*[set(t) for t in toplam])
    ara = re.sub(r'\s', '', ara or '').upper()
    satirlar = []
    for k in kodlar:
        ds = [t.get(k) for t in toplam]
        ilk_d = next(d for d in ds if d)
        if yalniz_p and (ilk_d['prov'] or '').upper() != 'P':
            continue
        if tesis and tesis_bul(k, harita) != tesis:
            continue
        if ara and ara not in k:
            continue
        hiz = [_hizala(d, ilk, son, ufuk) for d in ds]
        yakin_fark = round(hiz[0][1] - hiz[1][1], 2)
        toplam_fark = round(hiz[0][2] - hiz[1][2], 2)
        onceki_fark = round(hiz[1][2] - hiz[2][2], 2) if len(ds) > 2 else None
        if abs(yakin_fark) < 0.01 and abs(toplam_fark) < 0.01 and not (onceki_fark and abs(onceki_fark) >= 0.01):
            continue
        if not ds[1]:
            tur = 'yeni_kod'
        elif not ds[0]:
            tur = 'kalkan_kod'
        elif yakin_fark > 0:
            tur = 'eklendi'
        elif yakin_fark < 0:
            tur = 'azaldi'
        else:
            tur = 'ileri'
        # Müşteri kırılımı: yakın talebi değişen müşteriler
        musteri = []
        mk = set((ds[0] or {}).get('musteri', {})) | set((ds[1] or {}).get('musteri', {}))
        for m in mk:
            my, me = (ds[0] or {}).get('musteri', {}).get(m), (ds[1] or {}).get('musteri', {}).get(m)
            hy, he = _hizala(my, ilk, son, ufuk), _hizala(me, ilk, son, ufuk)
            if abs(hy[1] - he[1]) >= 0.01 or abs(hy[2] - he[2]) >= 0.01:
                musteri.append({'musteri': (my or me)['musteri'], 'musteri_kodu': m[0], 'dest': m[1],
                                'yakin_eski': round(he[1], 2), 'yakin_yeni': round(hy[1], 2),
                                'toplam_eski': round(he[2], 2), 'toplam_yeni': round(hy[2], 2)})
        musteri.sort(key=lambda x: -abs(x['yakin_yeni'] - x['yakin_eski']))
        satirlar.append({
            'kod': ilk_d['kod'], 'prov': ilk_d['prov'], 'tur': tur, 'tesis': tesis_bul(k, harita) if tesis else '',
            'stok': [round(d['stok'], 2) if d else None for d in ds],
            'bakiye': [round(h[0], 2) for h in hiz],
            'haftalar': [[round((d or {}).get('h', {}).get(w, 0), 2) for d in ds] for w in ufuk],
            'yakin': [round(h[1], 2) for h in hiz], 'toplam': [round(h[2], 2) for h in hiz],
            'yakin_fark': yakin_fark, 'toplam_fark': toplam_fark, 'onceki_fark': onceki_fark,
            'musteriler': musteri[:25], 'musteri_sayisi': len(musteri),
        })
    sira = {'eklendi': 0, 'yeni_kod': 1, 'azaldi': 2, 'kalkan_kod': 3, 'ileri': 4}
    satirlar.sort(key=lambda s: (sira.get(s['tur'], 9), -abs(s['yakin_fark']), -abs(s['toplam_fark']), s['kod']))
    ozet = {}
    for s in satirlar:
        o = ozet.setdefault(s['tur'], {'urun': 0, 'yakin': 0.0, 'toplam': 0.0})
        o['urun'] += 1
        o['yakin'] = round(o['yakin'] + s['yakin_fark'], 2)
        o['toplam'] = round(o['toplam'] + s['toplam_fark'], 2)
    return {'cekimler': ck, 'haftalar': ufuk, 'ilk_n': ilk_n, 'tesis': tesis or '', 'satirlar': satirlar,
            'ozet': ozet, 'notlar': notlar,
            'son_hafta': next((h for c in ck for h in c['haftalar'] if _hafta_sira(h) == son), '')}


def urun_farki_excel(sonuc):
    """Planlamanın 'Sipariş farkları' düzeni: her çekim için Stok · Bakiye · ilk N hafta ·
    Toplam blokları yan yana + 'İLK N HAFTA İÇİNE EKLENEN MİKTAR' · FARK sütunları."""
    import io
    from openpyxl import Workbook
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    ck, ufuk, n = sonuc['cekimler'], sonuc['haftalar'], sonuc['ilk_n']
    wb = Workbook()
    ws = wb.active
    ws.title = 'sipariş farkları'
    blok = 2 + len(ufuk) + 1                              # stok, bakiye, haftalar, toplam
    bas1, bas2 = ['', '', ''], ['Ürün kodu', 'A/P', 'Tür']
    for c in ck:
        bas1 += [datetime.strptime(c['tarih'], '%Y-%m-%d').strftime('%d.%m.%Y') + f" ({c['kapsam'].get('etiket', '?')})"] + [''] * (blok - 1)
        bas2 += ['Stok', 'Bakiye'] + list(ufuk) + ['Toplam']
    g = lambda c: datetime.strptime(c['tarih'], '%Y-%m-%d').strftime('%d.%m')   # noqa: E731
    farklar = [f'İLK {n} HAFTA İÇİNE EKLENEN MİKTAR ({g(ck[0])} − {g(ck[1])})',
               f'FARK toplam ({g(ck[0])} − {g(ck[1])})']
    if len(ck) > 2:
        farklar.append(f'FARK toplam ({g(ck[1])} − {g(ck[2])})')
    bas1 += [''] * len(farklar) + ['']
    bas2 += farklar + ['Değişen müşteriler']
    ws.append(bas1)
    ws.append(bas2)
    tur_ad = {'eklendi': 'eklendi', 'yeni_kod': 'yeni kod', 'azaldi': 'azaldı', 'kalkan_kod': 'kalktı',
              'ileri': 'ileri haftalar'}
    for s in sonuc['satirlar']:
        r = [s['kod'], s['prov'], tur_ad.get(s['tur'], s['tur'])]
        for i in range(len(ck)):
            r += [s['stok'][i], s['bakiye'][i]] + [hw[i] for hw in s['haftalar']] + [s['toplam'][i]]
        r += [s['yakin_fark'], s['toplam_fark']] + ([s['onceki_fark']] if len(ck) > 2 else [])
        r.append(', '.join(f"{m['musteri']} ({m['yakin_eski']:g}→{m['yakin_yeni']:g})" for m in s['musteriler'][:5]))
        ws.append(r)
    koyu = PatternFill('solid', fgColor='1F3864')
    for i in range(1, ws.max_column + 1):
        for rr in (1, 2):
            c = ws.cell(row=rr, column=i)
            c.font = Font(bold=True, color='FFFFFF' if rr == 2 else '1F3864')
            if rr == 2:
                c.fill = koyu
            c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
    ilk_fark = 3 + blok * len(ck) + 1
    son_fark = ilk_fark + len(farklar) - 1
    alan = f'{get_column_letter(ilk_fark)}3:{get_column_letter(son_fark)}{max(3, ws.max_row)}'
    ws.conditional_formatting.add(alan, CellIsRule(operator='greaterThan', formula=['0'],
                                                   fill=PatternFill('solid', fgColor='F8CBAD')))
    ws.conditional_formatting.add(alan, CellIsRule(operator='lessThan', formula=['0'],
                                                   fill=PatternFill('solid', fgColor='C6EFCE')))
    for i in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(i)].width = 9
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['C'].width = 11
    for i in range(ilk_fark, son_fark + 1):
        ws.column_dimensions[get_column_letter(i)].width = 16
    ws.column_dimensions[get_column_letter(son_fark + 1)].width = 60
    ws.row_dimensions[2].height = 45
    ws.freeze_panes = 'D3'
    ws.auto_filter.ref = f'A2:{get_column_letter(ws.max_column)}{max(2, ws.max_row)}'
    if sonuc['notlar']:
        w2 = wb.create_sheet('Not')
        for x in sonuc['notlar']:
            w2.append([x])
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


# ── ANA VERİ'DE OLMAYAN KODLAR ───────────────────────────────────────────────
_ISLEM_EKI = ('W', '-S', 'S', 'GW', 'GRW', '-SW', 'SW')


def _kok(c):
    m = re.match(r'^(\d{2}\.[0-9A-Z]+\.[A-Z]*\d+)', c)
    return m.group(1) if m else c


def ana_veri_kodlari():
    """Ana Veri'deki TÜM kodlar (bölümü boş satırlar dahil) + TK2 referans listesi."""
    import import_excel as IE
    kodlar = set()
    if IE.ana_veri_aktif():
        wb = IE._wb_onbellekten(IE.ANA_VERI_YOL)
        ws, kol = IE._ana_sayfa_bul(wb)
        if ws is not None:
            for row in ws.iter_rows(min_row=2, values_only=True):
                v = row[kol['kod']] if kol['kod'] < len(row) else None
                if v not in (None, ''):
                    kodlar.add(IE._norm_kod(v))
    return kodlar


def eksik_kodlar(conn, cekim_id, yalniz_p=True):
    """Çekimdeki kodlardan Ana Veri'de (ve TK2 referans listesinde) olmayanlar.
    W / -S hâli kayıtlı olanlar (kaynak/lazer koda o hâliyle girilir) ayrı döner."""
    import import_excel as IE
    kayitli = ana_veri_kodlari()
    for r in conn.execute("SELECT referans_kodu FROM referans_listesi WHERE COALESCE(lokasyon,'TK2')='TK2'"):
        kayitli.add(IE._norm_kod(r[0]))
    kok_idx = {}
    for k in kayitli:
        kok_idx.setdefault(_kok(k), set()).add(k)
    toplam = {}
    for s in _satirlar(conn, cekim_id):
        if yalniz_p and s['prov'] != 'P':
            continue
        n = IE._norm_kod(s['article'])
        t = toplam.setdefault(n, {'kod': s['article'], 'prov': s['prov'], 'priorit': s['priorit'],
                                  'toplam': 0.0, 'gecikmis': 0.0, 'musteri': set(), 'ilk_hafta': ''})
        t['toplam'] += s['toplam']
        t['gecikmis'] += s['gecikmis']
        t['musteri'].add(s['musteri'])
        if not t['ilk_hafta']:
            for h, a in s['haftalar'].items():
                if a:
                    t['ilk_hafta'] = h
                    break
    eksik, islem_hali = [], []
    for n, t in toplam.items():
        if n in kayitli:
            continue
        turev = sorted(kok_idx.get(_kok(n), set()) - {n})
        kayit = {'kod': t['kod'], 'prov': t['prov'], 'priorit': t['priorit'], 'toplam': round(t['toplam'], 2),
                 'gecikmis': round(t['gecikmis'], 2), 'musteri_sayisi': len(t['musteri']),
                 'musteriler': sorted(t['musteri'])[:5], 'ilk_hafta': t['ilk_hafta'], 'benzer': turev[:5]}
        if any(x.startswith(n) and x[len(n):] in _ISLEM_EKI for x in turev):
            islem_hali.append(kayit)
        else:
            eksik.append(kayit)
    eksik.sort(key=lambda x: -x['toplam'])
    islem_hali.sort(key=lambda x: -x['toplam'])
    return {'eksik': eksik, 'islem_hali': islem_hali, 'kontrol_edilen': len(toplam)}


def eksik_excel(eksik, tarih):
    """Ana Veri sütun düzeninde (CD ART. … Not) — masaüstü Ana Veri'ye yapıştırmalık."""
    import io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = 'Ana Veri'
    ws.append(['CD ART.', 'PRIORIT.', 'A/P', 'TK-1/2', 'Bölüm', 'Makine', 'Çevrim süresi / söktak (sn)',
               'Kaynak süresi (sn)', 'Kalıp göz', 'Büküm op.', 'Açıklama', 'Süre teyit', 'Kapasite 1V/Adet', 'Not'])
    for c in ws[1]:
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor='1F3864')
    for x in eksik:
        ws.append([x['kod'], x['priorit'], x['prov'], 'TK-2', None, None, None, None, None, None, None, None, None,
                   f"Satış planı {tarih} — bölüm atanacak · açık {x['toplam']:g} ({x['musteri_sayisi']} müşteri)"])
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['N'].width = 60
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()
