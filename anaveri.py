# -*- coding: utf-8 -*-
"""
anaveri.py — planlamanın ANAVERİ sayfası (PLAN 26xxxx.xlsb), Forge'da ortak tablo.

Sütunlar: CD ART. · PRIORIT. · A/P · TK-1/2 · Hat · Makine · Hammadde · Kapasite 1V/Adet
· zor/kolay. TK-1/2 değerleri: TK-1 · TK-2 · Pandora (tel tedarikçisi) · Tabo (hortum) ·
Satınalma · Fason.

Kullananlar:
  · tel_plani  — telin tedarikçisi (Pandora / TK-1 / Satınalma), Pull kontrolü, kapasite
  · satis_plani — çekimin KAPSAMI (yalnız TK2 mi, TK1+TK2 mi) ve tesis süzgeci

Tohum: tohum/plan_anaveri.json (PLAN 260921, tüm satırlar) — tablo BOŞSA yazılır;
sonrası panelden 'Anaveri yükle' / 'Plan klasöründen oku'.
"""
import json
import os
import re
from datetime import datetime

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
TOHUM = os.path.join(PROJECT_DIR, 'tohum', 'plan_anaveri.json')
VARSAYILAN_PLAN_KLASORU = r'Q:\UretimPlanlama\EMRE\Yeni klasör\Plan'
# TK-1/2 değeri → tesis (satış planı kapsamı için). Pandora telleri ve Tabo hortumları
# TK1 ürünlerinde kullanılır; satınalma/fason kalemleri üretim tesisine sayılmaz.
TESIS = {'TK-1': 'TK1', 'TK1': 'TK1', 'PANDORA': 'TK1', 'TABO': 'TK1', 'TK-2': 'TK2', 'TK2': 'TK2'}

_BASLIK = {'kod': ('CD ART.', 'CD ART'), 'ap': ('A/P',), 'tk': ('TK-1/2', 'TK 1/2'), 'hat': ('HAT',),
           'makine': ('MAKINE', 'MAKİNE'), 'kapasite': ('KAPASITE 1V/ADET', 'KAPASİTE 1V/ADET', 'KAPASITE'),
           'zorluk': ('ZOR/KOLAY', 'ZORLUK')}


def tablolari_kur(conn):
    conn.execute("""CREATE TABLE IF NOT EXISTS plan_anaveri (
        kod TEXT PRIMARY KEY, ap TEXT DEFAULT '', tk TEXT DEFAULT '', hat TEXT DEFAULT '',
        makine TEXT DEFAULT '', kapasite REAL, zorluk TEXT DEFAULT '')""")
    conn.execute("""CREATE TABLE IF NOT EXISTS plan_anaveri_yukleme (
        id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT NOT NULL, dosya TEXT DEFAULT '',
        satir INTEGER DEFAULT 0, kullanici TEXT DEFAULT '')""")
    if conn.execute("SELECT COUNT(*) FROM plan_anaveri").fetchone()[0] == 0 and os.path.exists(TOHUM):
        try:
            with open(TOHUM, encoding='utf-8') as f:
                tohum = json.load(f)
            yaz(conn, tohum.get('satirlar') or [], tohum.get('dosya') or 'tohum', 'tohum', commit=False)
        except Exception as e:                    # tohum okunamazsa kullananlar yine çalışır
            print(f'[ANAVERİ] tohum okunamadı: {e}')
    conn.commit()


def _metin(v):
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return '' if s in ('0x2a', '#N/A', '#YOK') else s


def satirlari(satirlar):
    """Ham sayfa satırları → [{kod, ap, tk, hat, makine, kapasite, zorluk}] (tekil kod).
    Başlık satırı ilk 10 satırda 'CD ART' ile aranır; sütunlar ADLA bulunur."""
    satirlar = list(satirlar)
    bas_i, ix = None, {}
    for i, r in enumerate(satirlar[:10]):
        adlar = [_metin(v).upper() for v in r]
        if any(a.startswith('CD ART') for a in adlar):
            bas_i = i
            for alan, adaylar in _BASLIK.items():
                for j, a in enumerate(adlar):
                    if alan not in ix and a in adaylar:
                        ix[alan] = j
            break
    if bas_i is None or 'kod' not in ix or 'tk' not in ix:
        raise ValueError("Anaveri sayfasında 'CD ART.' ve 'TK-1/2' başlıkları bulunamadı")
    out, gorulen = [], set()
    for r in satirlar[bas_i + 1:]:
        def al(a):
            return _metin(r[ix[a]]) if a in ix and ix[a] < len(r) else ''
        kod = re.sub(r'\s+', '', al('kod')).upper()
        if not kod or kod in gorulen or not re.match(r'^[0-9A-Z]', kod):
            continue
        gorulen.add(kod)
        try:
            kap = float(al('kapasite')) if al('kapasite') else None
        except ValueError:
            kap = None
        out.append({'kod': kod, 'ap': al('ap'), 'tk': al('tk'), 'hat': al('hat'), 'makine': al('makine'),
                    'kapasite': kap if kap and kap > 0 else None, 'zorluk': al('zorluk')})
    return out


def dosyadan(yol):
    """PLAN 26xxxx.xlsb / .xlsx → satırlar. Sayfa 'Anaveri' (ya da 'Ana Veri')."""
    if yol.lower().endswith('.xlsb'):
        from pyxlsb import open_workbook
        with open_workbook(yol) as wb:
            ad = next((s for s in wb.sheets if s.replace(' ', '').lower() == 'anaveri'), None)
            if not ad:
                raise ValueError(f"Dosyada 'Anaveri' sayfası yok (sayfalar: {', '.join(wb.sheets)})")
            with wb.get_sheet(ad) as ws:
                return satirlari([c.v for c in r] for r in ws.rows())
    from openpyxl import load_workbook
    wb = load_workbook(yol, read_only=True, data_only=True)
    try:
        ws = next((w for w in wb.worksheets if w.title.replace(' ', '').lower() == 'anaveri'), None)
        if ws is None:
            raise ValueError(f"Dosyada 'Anaveri' sayfası yok (sayfalar: {', '.join(wb.sheetnames)})")
        return satirlari(ws.iter_rows(values_only=True))
    finally:
        wb.close()


def en_yeni_plan(klasor=VARSAYILAN_PLAN_KLASORU):
    """Klasördeki en yeni 'PLAN *.xlsb' (Excel'in kilit/geçici dosyaları hariç)."""
    if not os.path.isdir(klasor):
        raise FileNotFoundError(f'Klasöre erişilemiyor: {klasor}')
    adaylar = [os.path.join(klasor, f) for f in os.listdir(klasor)
               if f.upper().startswith('PLAN ') and f.lower().endswith(('.xlsb', '.xlsx')) and not f.startswith('~$')]
    if not adaylar:
        raise FileNotFoundError(f"'{klasor}' içinde PLAN dosyası yok")
    return max(adaylar, key=os.path.getmtime)


def yaz(conn, satirlar, dosya, kullanici, commit=True):
    if not satirlar:
        raise ValueError("Anaveri'de kod satırı bulunamadı")
    conn.execute("DELETE FROM plan_anaveri")
    conn.executemany("INSERT OR REPLACE INTO plan_anaveri (kod, ap, tk, hat, makine, kapasite, zorluk) "
                     "VALUES (?,?,?,?,?,?,?)",
                     [(s['kod'], s.get('ap', ''), s.get('tk', ''), s.get('hat', ''), s.get('makine', ''),
                       s.get('kapasite'), s.get('zorluk', '')) for s in satirlar])
    conn.execute("INSERT INTO plan_anaveri_yukleme (ts, dosya, satir, kullanici) VALUES (?,?,?,?)",
                 (datetime.now().strftime('%Y-%m-%d %H:%M'), os.path.basename(str(dosya)), len(satirlar), kullanici))
    if commit:
        conn.commit()
    return len(satirlar)


def haritasi(conn, onek=None):
    """{kod: {ap, tk, hat, makine, kapasite, zorluk}} — onek verilirse yalnız o ön ekliler."""
    tablolari_kur(conn)
    sql, par = "SELECT kod, ap, tk, hat, makine, kapasite, zorluk FROM plan_anaveri", ()
    if onek:
        sql, par = sql + " WHERE kod LIKE ?", (onek + '%',)
    return {r[0]: {'ap': r[1] or '', 'tk': r[2] or '', 'hat': r[3] or '', 'makine': r[4] or '',
                   'kapasite': r[5], 'zorluk': r[6] or ''} for r in conn.execute(sql, par)}


def durumu(conn):
    tablolari_kur(conn)
    r = conn.execute("SELECT ts, dosya, satir, kullanici FROM plan_anaveri_yukleme ORDER BY id DESC LIMIT 1").fetchone()
    return dict(zip(('ts', 'dosya', 'satir', 'kullanici'), r)) if r else None
