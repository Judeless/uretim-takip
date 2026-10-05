# -*- coding: utf-8 -*-
"""
SATIŞ PLANI ARŞİVİ + KARŞILAŞTIRMA (kullanıcı 2026-10-05)

AS400 10-05-03-06 "Orders plan" (Periyot S · Detay S · Wh. Dft DB 01D · I) F6 ile
çalıştırılınca sonuç QGPL/TKC0301F içindeki S650B9 iş dosyasına yazılır (hafta
başlıkları S650B9F2'de). Planlamanın 'satış planı.xls' şablonu da bu dosyayı ODBC
ile okuyor. Forge aynı dosyayı COFLEFORGE profiliyle okur ve:

  · içerik DEĞİŞTİYSE çekimi saklar + günün Excel'ini yazar
    (data/satis_plani/Satis_Plani_YYYY-MM-DD.xlsx — gün içinde yeni çalıştırma
    gelirse o günün dosyası en son hâliyle değişir),
  · bir önceki günle karşılaştırır: yeni giren / öne çekilen (yakın haftaya
    eklenen) / kalkan / azalan talep — "üretim geç kaldı" denen durumda araya
    giren siparişi göstermek için,
  · Ana Veri'de olmayan P kodlarını çıkarır (bölüm ataması için).

S650B9 YALNIZ biri 10-05-03-06'yı çalıştırınca tazelenir; Forge bunu tetiklemez
(ekran robotu yok). Çalıştırılmayan gün için yeni dosya oluşmaz, panel söyler.
"""
import os
import re
import json
import hashlib
from datetime import datetime

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
KLASOR = os.path.join(PROJECT_DIR, 'data', 'satis_plani')
KUTUPHANELER = ('QGPL', 'TKC0301F')
GUN_SAKLA = 400                 # veritabanında tutulan gün (dosyalar silinmez)
YAKIN_HAFTA = 2                 # "yakın termin" = gecikmiş + ilk 2 hafta

# S650B9 alanları — şablondaki sorguyla AYNI sıra (XWQT01 = gecikmiş, 02..13 = 12 hafta)
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
    """S650B9F2 tek satırından hafta etiketleri ('40/2026'), soldan sağa."""
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


def as400_oku(baglan=None):
    """S650B9 + S650B9F2 → (satırlar [dict], hafta etiketleri, kütüphane). Yalnız OKUR."""
    if baglan is None:
        import sys
        sys.path.insert(0, os.path.join(PROJECT_DIR, 'as400'))
        import as400_config as CFG
        baglan = CFG.baglan
    cn = baglan(timeout=60)
    try:
        cur = cn.cursor()
        veri, kutuphane, hatalar = None, None, []
        for lib in KUTUPHANELER:
            try:
                cur.execute(f"SELECT {', '.join(SUTUNLAR)} FROM {lib}.S650B9")
                veri, kutuphane = cur.fetchall(), lib
                break
            except Exception as e:
                hatalar.append(f'{lib}: {e}')
        if veri is None:
            raise RuntimeError('S650B9 okunamadı — ' + ' | '.join(h[:220] for h in hatalar))
        basliklar = []
        for lib in (kutuphane,) + tuple(k for k in KUTUPHANELER if k != kutuphane):
            try:
                cur.execute(f"SELECT * FROM {lib}.S650B9F2")
                basliklar = _haftalar_baslik(cur.fetchone())
                if basliklar:
                    break
            except Exception:
                continue
    finally:
        try:
            cn.close()
        except Exception:
            pass
    if len(basliklar) != 12:
        basliklar = [f'H{i}' for i in range(1, 13)]
    satirlar = []
    for r in veri:
        r = list(r)
        art = _t(r[5])
        if not art:
            continue
        haftalar = {basliklar[i]: _f(r[15 + i]) for i in range(12)}
        satirlar.append({
            'qual': _t(r[0]), 'musteri_kodu': _kod_metin(r[1]), 'dest': _kod_metin(r[2]),
            'musteri': _t(r[3]).translate(_TR), 'article': art, 'prov': _t(r[7]).upper(),
            'priorit': _t(r[6]), 'stok': _f(r[13]), 'gecikmis': _f(r[14]), 'haftalar': haftalar,
            'toplam': _f(r[27]), 'tipo': _t(r[28]), 'depo': _t(r[29]),
            'ham': [_kod_metin(x) if isinstance(x, float) and i in (1, 2) else
                    (x.strip() if isinstance(x, str) else x) for i, x in enumerate(r)],
        })
    return satirlar, basliklar, kutuphane


def _imza(satirlar, basliklar):
    ozet = json.dumps([basliklar] + sorted(
        [s['musteri_kodu'], s['dest'], s['article'], s['gecikmis'], sorted(s['haftalar'].items()), s['toplam'], s['stok']]
        for s in satirlar), ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.sha256(ozet.encode('utf-8')).hexdigest()


def excel_yaz(satirlar, basliklar, ts, depo, yol):
    """Günün dosyası: şablonun 'Finale' düzeni (C … TOTALE, TIPO ART, MAG., UBIC1, UBIC2)."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    ws = wb.active
    ws.title = 'Satış Planı'
    ws.append([f'SATIŞ PLANI — AS400 10-05-03-06 (S650B9) · çekim {ts} · depo {depo or "?"} · {len(satirlar)} satır'])
    ws['A1'].font = Font(bold=True, size=12)
    ws.append([])
    ws.append(BASLIK + basliklar + ['TOTALE', 'TIPO ART', 'MAG.', 'UBIC1', 'UBIC2'])
    for c in ws[3]:
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor='1F3864')
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for s in satirlar:
        h = s['ham']
        ws.append([h[0], h[1], h[2], s['musteri'], h[4], s['article'], h[6], s['prov'], h[8], h[9],
                   h[10], h[11], h[12], s['stok'], s['gecikmis']] + [s['haftalar'][b] for b in basliklar]
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


def cek(conn, baglan=None, kullanici='otomatik', zorla=False):
    """AS400'den oku; içerik son çekimden FARKLIYSA sakla + günün Excel'ini yaz.
    Döner: {'degisti', 'cekim_id', 'satir', 'kod', 'tarih', 'dosya', 'depo', 'haftalar', ...}"""
    tablolari_kur(conn)
    satirlar, basliklar, kutuphane = as400_oku(baglan)
    imza = _imza(satirlar, basliklar)
    son = conn.execute("SELECT id, imza, ts, tarih FROM satis_plani_cekim ORDER BY id DESC LIMIT 1").fetchone()
    depolar = sorted({s['depo'] for s in satirlar if s['depo']})
    depo = ','.join(depolar)
    kod = len({s['article'].upper() for s in satirlar})
    if son and son[1] == imza and not zorla:
        return {'degisti': False, 'cekim_id': son[0], 'satir': len(satirlar), 'kod': kod,
                'son_degisim': son[2], 'depo': depo, 'haftalar': basliklar, 'kutuphane': kutuphane}
    simdi = datetime.now()
    ts, tarih = simdi.strftime('%Y-%m-%d %H:%M:%S'), simdi.strftime('%Y-%m-%d')
    dosya = os.path.join(KLASOR, f'Satis_Plani_{tarih}.xlsx')
    try:
        excel_yaz(satirlar, basliklar, simdi.strftime('%d.%m.%Y %H:%M'), depo, dosya)
    except PermissionError:
        # Günün dosyası biri tarafından açık/kilitli — çekim KAYBOLMASIN, saatli adla yaz
        dosya = os.path.join(KLASOR, f'Satis_Plani_{tarih}_{simdi.strftime("%H%M")}.xlsx')
        excel_yaz(satirlar, basliklar, simdi.strftime('%d.%m.%Y %H:%M'), depo, dosya)
    cur = conn.execute(
        "INSERT INTO satis_plani_cekim (ts, tarih, imza, satir, kod, haftalar, depo, kutuphane, dosya, kullanici) "
        "VALUES (?,?,?,?,?,?,?,?,?,?)",
        (ts, tarih, imza, len(satirlar), kod, json.dumps(basliklar), depo, kutuphane,
         os.path.basename(dosya), kullanici))
    cid = cur.lastrowid
    conn.executemany(
        "INSERT INTO satis_plani_satir (cekim_id, musteri_kodu, dest, musteri, article, prov, priorit, "
        "stok, gecikmis, haftalar, toplam, tipo, depo, ham) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        [(cid, s['musteri_kodu'], s['dest'], s['musteri'], s['article'], s['prov'], s['priorit'],
          s['stok'], s['gecikmis'], json.dumps(s['haftalar']), s['toplam'], s['tipo'], s['depo'],
          json.dumps(s['ham'], ensure_ascii=False, default=str)) for s in satirlar])
    # Aynı günün önceki çekimleri: satırları silinir (gün = son hâli), kaydı kalır
    eski = [r[0] for r in conn.execute(
        "SELECT id FROM satis_plani_cekim WHERE tarih=? AND id<>?", (tarih, cid)).fetchall()]
    for e in eski:
        conn.execute("DELETE FROM satis_plani_satir WHERE cekim_id=?", (e,))
    # Çok eski günler (dosyalar diskte kalır)
    sinir = conn.execute("SELECT date('now','localtime', ?)", (f'-{GUN_SAKLA} day',)).fetchone()[0]
    for e in conn.execute("SELECT id FROM satis_plani_cekim WHERE tarih < ?", (sinir,)).fetchall():
        conn.execute("DELETE FROM satis_plani_satir WHERE cekim_id=?", (e[0],))
    conn.commit()
    print(f'[SATIŞ PLANI] yeni çekim #{cid}: {len(satirlar)} satır, {kod} kod, depo {depo or "?"} → {os.path.basename(dosya)}')
    return {'degisti': True, 'cekim_id': cid, 'satir': len(satirlar), 'kod': kod, 'tarih': tarih,
            'ts': ts, 'dosya': os.path.basename(dosya), 'depo': depo, 'haftalar': basliklar,
            'kutuphane': kutuphane}


# ── OKUMA ────────────────────────────────────────────────────────────────────
def gunluk_cekimler(conn, limit=120):
    """Her günün SON çekimi (satırları olanlar), yeniden eskiye."""
    tablolari_kur(conn)
    return [dict(zip(('id', 'ts', 'tarih', 'satir', 'kod', 'depo', 'dosya', 'haftalar'), r)) for r in conn.execute(
        "SELECT c.id, c.ts, c.tarih, c.satir, c.kod, c.depo, c.dosya, c.haftalar FROM satis_plani_cekim c "
        "WHERE c.id = (SELECT MAX(id) FROM satis_plani_cekim c2 WHERE c2.tarih = c.tarih) "
        "ORDER BY c.tarih DESC LIMIT ?", (limit,)).fetchall()]


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
    anahtar = lambda s: (s['musteri_kodu'], s['dest'], s['article'].upper())
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
