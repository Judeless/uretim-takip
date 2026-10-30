# -*- coding: utf-8 -*-
"""DENETİM — açılan / kapatılan launch (iş emri) listesi (kullanıcı 2026-10-30).

İç kontrol talebi madde 1: "Açılan ve kapatılan tüm launch / iş emirlerinin listesi: iş emri
no, ürün kodu, açılış ve kapanış tarihi, planlanan / üretilen / kapatılan adet, kapatan
kullanıcı, finalize edildiği yer (TK1 / TK2)" — 15–19.06.2026 ve 07–11.09.2026 (ADP istenmedi).

SALT OKUNUR — yalnız SELECT. SUNUCUDA, promanage (RDP) oturumunda çalışır (ODBC sürücüsü ve
şifre kasası orada; laptoptan AS400'e bağlanılmaz):
    sunucu\\Launch_Listesi.bat
    python as400\\launch_listesi.py [2026-06-15:2026-06-19 2026-09-07:2026-09-11]

NEDEN KOLON KEŞFİ: BPROF0'dan (OP ana dosyası, kapalı emirler dahil) bugüne kadar yalnız
numara / durum / adet / planlanan tarih (Q0IPD*, Q0FPD*) okundu; launch'ın AÇILIŞ ve KAPANIŞ
tarihi ile kullanıcı kolonları bilinmiyor. Betik kolonları QSYS2.SYSCOLUMNS açıklamalarından
(İtalyanca: 'lancio', 'chiusura', 'utente') seçer ve seçimini Excel'in 'Kolon seçimi'
sayfasına yazar. Bütün kolonlar 'Ham' sayfasında durur — seçim yanlışsa AS400'e yeniden
gitmeden düzeltilebilir. ERP tarihi 4 kolondur: SS (yüzyıl, 20) · AA (26) · AY · GÜN.
"""
import io
import os
import re
import sqlite3
import sys
from datetime import date, datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
AS = os.path.dirname(os.path.abspath(__file__))
KOK = os.path.dirname(AS)
sys.path.insert(0, AS)
sys.path.insert(0, KOK)
import as400_config as CFG  # noqa: E402

SEMA = 'TKC0301F'
DONEMLER = [('2026-06-15', '2026-06-19'), ('2026-09-07', '2026-09-11')]
SAYISAL = ('DECIMAL', 'NUMERIC', 'INTEGER', 'SMALLINT', 'BIGINT')
# Kolon açıklamasında aranan sözcükler (ERP İtalyanca) — ilk eşleşen kazanır
ACILIS_ANAHTAR = ('LANCIO', 'LANC', 'EMISS', 'APERT', 'IMMISS')
KAPANIS_ANAHTAR = ('CHIUSURA', 'CHIUS', 'CHIU', 'SALDO', 'EVAS', 'FINE EFF', 'CONSUNT')
KULLANICI_ANAHTAR = ('UTENTE', 'UTEN', 'USER', 'OPERAT')
DURUM_AD = {10: 'OPR (launch alınmamış)', 40: 'Launch açık', 45: 'Launch açık (TK1 akışı)',
            50: 'Launch açık (TK1 akışı)', 70: 'Kapalı'}


def _t(v):
    return '' if v is None else str(v).strip()


def _k(v):
    """Launch numara parçası anahtarı: Decimal('26') / 26 / '26' → '26'."""
    t = _t(v)
    return t[:-2] if t.endswith('.0') else t


_KONTROL = re.compile('[' + ''.join(chr(c) for c in list(range(0, 9)) + [11, 12] + list(range(14, 32))) + ']')


def _hucre(v):
    """Excel'e yazılabilir değer: Decimal → float, bytes → hex, kontrol karakteri atılır."""
    if hasattr(v, 'as_tuple'):
        return float(v)
    if isinstance(v, (bytes, bytearray)):
        return v.hex()
    if isinstance(v, str):
        return _KONTROL.sub('', v).strip()
    return v


def _sayi(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _gun(n):
    """20260615 → date | None."""
    try:
        n = int(n or 0)
        return date(n // 10000, n // 100 % 100, n % 100) if n else None
    except (TypeError, ValueError):
        return None


def kolonlar(cur, tablo):
    cur.execute("SELECT COLUMN_NAME, DATA_TYPE, LENGTH, NUMERIC_SCALE, COALESCE(COLUMN_TEXT, ''), "
                "COALESCE(COLUMN_HEADING, '') FROM QSYS2.SYSCOLUMNS "
                "WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ? ORDER BY ORDINAL_POSITION", (SEMA, tablo))
    return [{'ad': _t(r[0]), 'tip': _t(r[1]).upper(), 'uzunluk': r[2], 'olcek': r[3] or 0,
             'aciklama': _t(r[4]), 'baslik': re.sub(r'\s+', ' ', _t(r[5]))} for r in cur.fetchall()]


def _metin(*kol):
    parca = [x for k in kol for x in (k['aciklama'], k['baslik']) if x]
    return ' / '.join(dict.fromkeys(p.upper() for p in parca))


def tarih_gruplari(kol):
    """Tarih kolonları → [{ad, kolonlar, ifade, metin}]. ifade = YYYYMMDD sayısı veren SQL.
    (1) 4'lü SS/AA/AY/GÜN: aynı önek + 1..4 (Q0FPD1..4 gibi) sayısal kolonlar;
    (2) DATE / TIMESTAMP; (3) 8 haneli sayı (YYYYMMDD), açıklamasında tarih geçen."""
    adlar = {k['ad']: k for k in kol}
    out = []
    for k in kol:
        m = re.match(r'^(.+)1$', k['ad'])
        if not m:
            continue
        parca = [adlar.get(m.group(1) + str(i)) for i in range(1, 5)]
        if all(parca) and all(p['tip'] in SAYISAL and p['olcek'] == 0 and (p['uzunluk'] or 0) <= 4 for p in parca):
            c = [p['ad'] for p in parca]
            out.append({'ad': m.group(1), 'kolonlar': c, 'metin': _metin(*parca),
                        'ifade': f"({c[0]}*1000000 + {c[1]}*10000 + {c[2]}*100 + {c[3]})"})
    for k in kol:
        if k['tip'] in ('DATE', 'TIMESTAMP'):
            out.append({'ad': k['ad'], 'kolonlar': [k['ad']], 'metin': _metin(k),
                        'ifade': f"(YEAR({k['ad']})*10000 + MONTH({k['ad']})*100 + DAY({k['ad']}))"})
        elif (k['tip'] in SAYISAL and k['olcek'] == 0 and k['uzunluk'] == 8
              and re.search(r'\bDT|DAT', _metin(k))):
            out.append({'ad': k['ad'], 'kolonlar': [k['ad']], 'metin': _metin(k), 'ifade': f"({k['ad']})"})
    return out


def sec(liste, anahtarlar, haric=()):
    for a in anahtarlar:
        for g in liste:
            if a in g['metin'] and not any(h in g['metin'] for h in haric):
                return g
    return None


def kullanici_kolonlari(kol):
    return [k for k in kol if k['tip'] in ('CHAR', 'VARCHAR', 'GRAPHIC', 'VARGRAPHIC')
            and any(a in _metin(k) for a in KULLANICI_ANAHTAR)]


def tesis_haritasi():
    """Forge'un ürün → TK1/TK2 eşlemesi (Anaveri + referans listesi). Yoksa boş."""
    try:
        import satis_plani as SP
        c = sqlite3.connect(f"file:{os.path.join(KOK, 'uretim.db')}?mode=ro", uri=True)
        try:
            h = SP.tesis_haritasi(c)
        finally:
            c.close()
        return h, SP.tesis_bul
    except Exception as e:
        print(f'  (tesis eşlemesi okunamadı: {e} — TK1/TK2 yalnız ERP durumundan)')
        return {}, lambda kod, h: ''


def main(donemler):
    print('Açılan / kapatılan launch listesi — salt okunur')
    print('Dönemler: ' + ', '.join(f'{a} – {b}' for a, b in donemler))
    cn = CFG.baglan(timeout=30)
    cur = cn.cursor()
    print(f'AS400 bağlantısı OK (profil {CFG.DB_KULLANICI})')

    # ── 1) Kolon keşfi ──
    kol_p = kolonlar(cur, 'BPROF0')
    kol_m = kolonlar(cur, 'BMMAF0')
    print(f'BPROF0: {len(kol_p)} kolon · BMMAF0: {len(kol_m)} kolon')
    gruplar = tarih_gruplari(kol_p)
    # Planlanan üretim tarihleri (Q0IPD / Q0FPD — 'inizio / fine produzione') açılış/kapanış DEĞİL
    plan = ('INIZ', 'FINE PROD', 'FINE PRD', 'PREV', 'RICH', 'CONSEGN')
    g_ac = sec(gruplar, ACILIS_ANAHTAR, plan)
    g_kp = sec(gruplar, KAPANIS_ANAHTAR, plan)
    kul_p = kullanici_kolonlari(kol_p)
    kul_kapanis = next((k for k in kul_p if any(a in _metin(k) for a in ('CHIU', 'SALD'))), None) or \
        next((k for k in kul_p if any(a in _metin(k) for a in ('VAR', 'MOD', 'ULT', 'AGG'))), None) or \
        (kul_p[-1] if kul_p else None)
    qt_kapanis = next((k for k in kol_p if k['tip'] in SAYISAL and k['ad'].startswith('Q0QT')
                       and any(a in _metin(k) for a in ('CHIU', 'SALD'))), None)
    depo = next((k for k in kol_p if any(a in _metin(k) for a in ('MAGAZ', 'DEPOS', 'MAG.'))), None)
    kul_m = kullanici_kolonlari(kol_m)
    print(f"Açılış tarihi  : {g_ac['ad'] + ' — ' + g_ac['metin'][:60] if g_ac else 'BULUNAMADI'}")
    print(f"Kapanış tarihi : {g_kp['ad'] + ' — ' + g_kp['metin'][:60] if g_kp else 'BULUNAMADI (son teyit tarihi kullanılacak)'}")
    print(f"Kapatan kullanıcı: {kul_kapanis['ad'] + ' — ' + _metin(kul_kapanis)[:60] if kul_kapanis else 'BPROF0da yok (son teyidi giren kullanıcı)'}")

    # ── 2) Dönemlere değen emirler (HERHANGİ bir tarih grubu dönemde) ──
    araliklar = [(int(a.replace('-', '')), int(b.replace('-', ''))) for a, b in donemler]
    kosul = ' OR '.join(f"{g['ifade']} BETWEEN {a} AND {b}" for g in gruplar for a, b in araliklar) or '1=0'
    cur.execute(f"SELECT p.*, a.A0ARDS AS URUN_TANIMI FROM {SEMA}.BPROF0 p "
                f"LEFT JOIN {SEMA}.BARTF0 a ON a.A0ARTI = p.Q0ARTI WHERE {kosul}")
    adlar = [d[0].strip().upper() for d in cur.description]
    ham = [dict(zip(adlar, r)) for r in cur.fetchall()]
    anahtar = lambda red2, renu: (_k(red2), _k(renu))      # noqa: E731
    renu_sayisal = next((k['tip'] in SAYISAL for k in kol_m if k['ad'] == 'MGRENU'), True)
    if not g_kp:
        # Kapanış kolonu yok → kapanış = son teyit. Planlanan tarihi dönem dışında kalıp
        # dönemde teyidi yapılan (ve kapanan) emirler de aday: BMMAF0'dan bulunur.
        mg = '(MGDSSO*1000000 + MGDAAO*10000 + MGDMMO*100 + MGDGGO)'
        cur.execute(f"SELECT DISTINCT MGRED2, MGRENU FROM {SEMA}.BMMAF0 WHERE MGCACD = 'RPR' AND (" +
                    ' OR '.join(f'{mg} BETWEEN {a} AND {b}' for a, b in araliklar) + ')')
        var = {anahtar(r.get('Q0RED2'), r.get('Q0RENU')) for r in ham}
        ek = {anahtar(r[0], r[1]) for r in cur.fetchall()} - var
        renu_ek = sorted({e[1] for e in ek if e[1]})
        for i in range(0, len(renu_ek), 400):
            parca = renu_ek[i:i + 400]
            cur.execute(f"SELECT p.*, a.A0ARDS AS URUN_TANIMI FROM {SEMA}.BPROF0 p "
                        f"LEFT JOIN {SEMA}.BARTF0 a ON a.A0ARTI = p.Q0ARTI WHERE p.Q0RENU IN ({','.join('?' * len(parca))})",
                        [int(x) if x.isdigit() else x for x in parca])
            ham += [d for d in (dict(zip(adlar, r)) for r in cur.fetchall())
                    if anahtar(d.get('Q0RED2'), d.get('Q0RENU')) in ek]
        print(f'Dönemde teyit hareketi olan ek emir: {len(ek)}')
    print(f'Dönemlere değen emir satırı: {len(ham)}')

    def tarih(r, g):
        if not g:
            return None
        if len(g['kolonlar']) == 4:
            v = [int(_sayi(r.get(c)) or 0) for c in g['kolonlar']]
            return _gun(v[0] * 1000000 + v[1] * 10000 + v[2] * 100 + v[3])
        v = r.get(g['kolonlar'][0])
        if isinstance(v, (date, datetime)):
            return v.date() if isinstance(v, datetime) else v
        return _gun(_sayi(v))

    def donemde(d):
        if not d:
            return ''
        n = d.year * 10000 + d.month * 100 + d.day
        return next((f'{a[8:10]}.{a[5:7]}–{b[8:10]}.{b[5:7]}.{b[:4]}'
                     for (a, b), (x, y) in zip(donemler, araliklar) if x <= n <= y), '')

    # ── 3) Teyit hareketleri (BMMAF0 RPR) — üretilen + son teyit / kullanıcı ──
    hareket = {}
    launchlar = sorted({anahtar(r.get('Q0RED2'), r.get('Q0RENU')) for r in ham})
    kul_m_ad = kul_m[0]['ad'] if kul_m else None
    secim = ('MGRED2, MGRENU, MGDSSO, MGDAAO, MGDMMO, MGDGGO, MGQTA' + (f', {kul_m_ad}' if kul_m_ad else ''))
    renular = sorted({l[1] for l in launchlar if l[1]})
    for i in range(0, len(renular), 400):
        parca = renular[i:i + 400]
        cur.execute(f"SELECT {secim} FROM {SEMA}.BMMAF0 WHERE MGCACD = 'RPR' AND MGRENU IN "
                    f"({','.join('?' * len(parca))})",
                    [int(x) if renu_sayisal and x.isdigit() else x for x in parca])
        for r in cur.fetchall():
            d = _gun(int(_sayi(r[2]) or 0) * 1000000 + int(_sayi(r[3]) or 0) * 10000 +
                     int(_sayi(r[4]) or 0) * 100 + int(_sayi(r[5]) or 0))
            h = hareket.setdefault(anahtar(r[0], r[1]), {'adet': 0.0, 'sayi': 0, 'ilk': None, 'son': None,
                                                         'son_kul': '', 'donem_adet': {}})
            h['adet'] += _sayi(r[6]) or 0
            h['sayi'] += 1
            if d and (not h['ilk'] or d < h['ilk']):
                h['ilk'] = d
            if d and (not h['son'] or d >= h['son']):
                h['son'] = d
                h['son_kul'] = _t(r[7]) if kul_m_ad else ''
            if donemde(d):
                h['donem_adet'][donemde(d)] = h['donem_adet'].get(donemde(d), 0) + (_sayi(r[6]) or 0)
    cn.close()
    print(f'Teyit hareketi bulunan launch: {len(hareket)}')

    # ── 4) Liste ──
    harita, tesis_bul = tesis_haritasi()
    satirlar = []
    for r in ham:
        try:
            durum = int(_sayi(r.get('Q0AVAN')) or -1)
        except (TypeError, ValueError):
            durum = -1
        h = hareket.get(anahtar(r.get('Q0RED2'), r.get('Q0RENU')), {})
        ac = tarih(r, g_ac)
        kp = tarih(r, g_kp) if g_kp else (h.get('son') if durum == 70 else None)
        d_ac, d_kp = donemde(ac), donemde(kp)
        if not (d_ac or d_kp):
            continue                       # yalnız planlanan tarihi dönemde olan emir (açılmamış/kapanmamış)
        kod = _t(r.get('Q0ARTI'))
        tesis = 'TK1' if durum in (45, 50) else (tesis_bul(kod, harita) or '')
        planlanan = _sayi(r.get('Q0QTOR')) or 0
        uretilen = _sayi(r.get('Q0QTRI')) or 0
        kapatilan = (_sayi(r.get(qt_kapanis['ad'])) if qt_kapanis else uretilen) if durum == 70 else None
        kapatan = (_t(r.get(kul_kapanis['ad'])) if kul_kapanis else h.get('son_kul', '')) if durum == 70 else ''
        # Haziran'da açılıp Eylül'de kapanan emir İKİ dönemin listesinde de görünür (ayrı satır)
        for dn in dict.fromkeys(x for x in (d_ac, d_kp) if x):
            olay = ' + '.join(x for x, v in (('Açıldı', d_ac), ('Kapandı', d_kp)) if v == dn)
            satirlar.append([_hucre(x) for x in [
                dn, olay, f"{_k(r.get('Q0RED1'))}-{_k(r.get('Q0RED2'))}-{_k(r.get('Q0RENU'))}", kod,
                _t(r.get('URUN_TANIMI')), tesis or '?', ac, kp, f'{durum} · {DURUM_AD.get(durum, "?")}',
                planlanan, uretilen, kapatilan, kapatan, h.get('ilk'), h.get('son'), h.get('sayi') or 0,
                (h.get('donem_adet') or {}).get(dn, 0), _t(r.get(depo['ad'])) if depo else '']])
    satirlar.sort(key=lambda s: (s[0], s[6] or s[7] or date.min, s[2]))
    print(f'Listeye giren launch: {len(satirlar)}')

    # ── 5) Excel ──
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter as H
    wb = Workbook()
    ws = wb.active
    ws.title = 'Launch listesi'
    koyu = PatternFill('solid', fgColor='1F3864')

    def baslik(sayfa, satir_no=1):
        for c in sayfa[satir_no]:
            c.font, c.fill = Font(bold=True, color='FFFFFF'), koyu
            c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')

    ws.append(['Dönem', 'Dönemdeki olay', 'İş emri (launch) no', 'Ürün kodu', 'Ürün tanımı', 'Tesis',
               'Açılış tarihi', 'Kapanış tarihi', 'Durum (ERP)', 'Planlanan adet', 'Üretilen (teyit) adet',
               'Kapatılan adet', 'Kapatan kullanıcı', 'İlk teyit', 'Son teyit', 'Teyit sayısı',
               'Dönemde teyit edilen', 'Depo'])
    baslik(ws)
    for s in satirlar:
        ws.append(s)
    for i, gen in enumerate([17, 16, 16, 18, 34, 7, 12, 12, 24, 11, 11, 11, 16, 11, 11, 9, 11, 8], start=1):
        ws.column_dimensions[H(i)].width = gen
    for row in ws.iter_rows(min_row=2):
        for i in (6, 7, 13, 14):
            row[i].number_format = 'DD.MM.YYYY'
        for i in (9, 10, 11, 16):
            row[i].number_format = '#,##0'
    ws.freeze_panes = 'D2'
    ws.auto_filter.ref = f'A1:{H(ws.max_column)}{max(1, ws.max_row)}'
    ws.row_dimensions[1].height = 32

    oz = wb.create_sheet('Özet')
    oz.append(['Dönem', 'Tesis', 'Açılan launch', 'Kapanan launch', 'Kapanan — planlanan adet', 'Kapanan — üretilen adet'])
    baslik(oz)
    toplam = {}
    for s in satirlar:
        t = toplam.setdefault((s[0], s[5]), [0, 0, 0.0, 0.0])
        if 'Açıldı' in s[1]:
            t[0] += 1
        if 'Kapandı' in s[1]:
            t[1] += 1
            t[2] += s[9] or 0
            t[3] += s[10] or 0
    for (d, tesis), t in sorted(toplam.items()):
        oz.append([d, tesis] + t)
    for i, gen in enumerate([22, 8, 14, 14, 22, 22], start=1):
        oz.column_dimensions[H(i)].width = gen

    ks = wb.create_sheet('Kolon seçimi')
    ks.append(['Alan', 'BPROF0 kolonu', 'Kolon açıklaması (AS400)'])
    baslik(ks)
    for alan, k, met in (('Açılış tarihi', g_ac and ', '.join(g_ac['kolonlar']), g_ac and g_ac['metin']),
                         ('Kapanış tarihi', g_kp and ', '.join(g_kp['kolonlar']),
                          g_kp['metin'] if g_kp else 'BULUNAMADI — kapalı (70) emirde SON TEYİT (BMMAF0 RPR) tarihi'),
                         ('Kapatan kullanıcı', kul_kapanis and kul_kapanis['ad'],
                          _metin(kul_kapanis) if kul_kapanis else
                          (f'BPROF0da yok — son teyidi giren (BMMAF0.{kul_m_ad})' if kul_m_ad else 'BULUNAMADI')),
                         ('Kapatılan adet', qt_kapanis['ad'] if qt_kapanis else 'Q0QTRI',
                          _metin(qt_kapanis) if qt_kapanis else 'ayrı kolon yok — kapanışta teyit edilen (Q0QTRI)'),
                         ('Planlanan adet', 'Q0QTOR', ''), ('Üretilen adet', 'Q0QTRI', 'teyit (rientrata)'),
                         ('Depo', depo and depo['ad'], depo and _metin(depo))):
        ks.append([alan, k or '—', met or ''])
    ks.append([])
    ks.append(['Bütün tarih grupları', 'Kolonlar', 'Açıklama', 'Dönemdeki satır'])
    for g in gruplar:
        n = sum(1 for r in ham if donemde(tarih(r, g)))
        ks.append([g['ad'], ', '.join(g['kolonlar']), _hucre(g['metin']), n])
    ks.append([])
    ks.append(['Kullanıcı kolonları (BPROF0)', 'Kolon', 'Açıklama'])
    for k in kul_p:
        ks.append(['', k['ad'], _metin(k)])
    for i, gen in enumerate([26, 30, 70, 14], start=1):
        ks.column_dimensions[H(i)].width = gen

    for ad, kol in (('BPROF0 kolonları', kol_p), ('BMMAF0 kolonları', kol_m)):
        w = wb.create_sheet(ad)
        w.append(['Kolon', 'Tip', 'Uzunluk', 'Ondalık', 'Açıklama', 'Başlık'])
        baslik(w)
        for k in kol:
            w.append([_hucre(x) for x in (k['ad'], k['tip'], k['uzunluk'], k['olcek'], k['aciklama'], k['baslik'])])
        for i, gen in enumerate([12, 10, 8, 8, 40, 40], start=1):
            w.column_dimensions[H(i)].width = gen

    hw = wb.create_sheet('Ham')
    hw.append(adlar)
    baslik(hw)
    for r in ham:
        hw.append([_hucre(v) for v in r.values()])
    hw.freeze_panes = 'A2'

    ad = f"Launch_Listesi_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    klasor = os.path.join(KOK, 'data', 'denetim')
    os.makedirs(klasor, exist_ok=True)
    yol = os.path.join(klasor, ad)
    wb.save(yol)
    print(f'\nKaydedildi: {yol}')
    masa = os.path.join(os.path.expanduser('~'), 'Desktop')
    if os.path.isdir(masa):
        import shutil
        shutil.copy2(yol, os.path.join(masa, ad))
        print(f'Masaüstüne de kopyalandı: {os.path.join(masa, ad)}')
    return yol


if __name__ == '__main__':
    donem = []
    for a in sys.argv[1:]:
        m = re.match(r'^(\d{4}-\d{2}-\d{2}):(\d{4}-\d{2}-\d{2})$', a)
        if not m:
            print(f'Dönem biçimi YYYY-AA-GG:YYYY-AA-GG olmalı: {a}')
            sys.exit(2)
        donem.append((m.group(1), m.group(2)))
    try:
        main(donem or DONEMLER)
    except Exception as e:
        print(f'\nHATA: {e!r}')
        if getattr(CFG, 'kimlik_hatasi_mi', None) and CFG.kimlik_hatasi_mi(e):
            print('  -> KİMLİK HATASI: tekrar denemeyin (profil kilitlenir); İtalya IT ile kontrol edin.')
        sys.exit(1)
