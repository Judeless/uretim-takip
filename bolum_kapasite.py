# -*- coding: utf-8 -*-
"""BÖLÜM KAPASİTESİ — Ana Veri'ye göre 4 haftalık yük / kişi analizi (kullanıcı 2026-10-21).

Kullanıcı: "ana verimize göre bölüm bazlı kapasite analizi yapalım. 4 haftalık kapasite
bakalım. bölümde çalışacak operatör sayılarını elle düzeltebileyim anlık. mevcuttaki
modülü beğenmedim ve eski database ile yapmıştık. daha doğru ve kolay anlaşılabilir olan
bir şey yapalım sonrasında onu görselleştirelim de grafiklerle."

HESAP (tek ölçü: KİŞİ)
  İhtiyaç   AS400 açık iş emirleri (XPRO90: OPR + açık launch, kalan = adet − teyit) —
            kapasite.talep_senkron'un kapasite_talep tablosu. Emir bitiş tarihinin haftasına
            yazılır; bu haftanın pazartesisinden eskisi 'gecikmiş' (ayrı sütun).
  Bölüm     Ana Veri (data/AnaVeri.xlsx): D 'TK-1/2' + E 'Bölüm'. TK2 bölümleri Forge
            anahtarıyla (Kaynak, Lazer Kesim, Büküm, Metal Enjeksiyon, İşleme, Montaj), diğer
            tesisler E'deki adla (TK1 Tel, LF, Joystick…). Satın alma / ADP / Aspanda dışarıda.
  Süre      G 'Çevrim süresi / söktak' (sn/adet; kaynakta G + H 'Kaynak süresi' = OEE'deki
            hedef çevrim). G yoksa M 'Kapasite 1V/Adet' (bir kişinin bir vardiyada yaptığı adet).
  Kişi-saat sn'li ürün: adet × sn ÷ 3600 ÷ verim · vardiya adetli ürün: adet ÷ M × günlük saat
            (M zaten gerçekçi vardiya adedi → verim ikinci kez düşülmez)
  Gereken kişi = haftalık kişi-saat ÷ (gün × günlük saat). Doluluk = gereken ÷ mevcut kişi.
  Süresiz   Ana Veri'de süresi olmayan ürünün yükü bölümün ORTANCA birim süresiyle TAHMİN
            edilir ve ayrı gösterilir (gizlenmez, tanımlıya da karışmaz).

Varsayılanlar 'Aylık NçS' sayfasından: Kişi + günlük saat = NçS ÷ kişi ÷ o ayın iş günü.
Kişi / günlük saat / verim panelden bölüm başına değiştirilir (bolum_kapasite_ayar).
"""
import calendar
import os
import statistics
from datetime import date, datetime, timedelta

VARSAYILAN_GUN = 5
VARSAYILAN_SAAT = 8.5
VARSAYILAN_VERIM = 85.0
TK2_SIRA = ('kaynak', 'lazer', 'pres', 'metal', 'isleme', 'montaj')
# Bizim üretmediğimiz kaynaklar: kapasiteye girmez (fason / satın alma)
DISARIDA_TESIS = ('satınalma', 'satinalma', 'adp', 'aspanda', 'depo')
HARIC_ONEK = ('92.',)            # 92.* bizde üretilmiyor (feedback 2026-10)

_ONBELLEK = {'anahtar': None, 'veri': None}


def tablolari_kur(conn):
    conn.execute("""
        CREATE TABLE IF NOT EXISTS bolum_kapasite_ayar (
            anahtar TEXT PRIMARY KEY,
            kisi REAL,
            gunluk_saat REAL,
            verim REAL,
            guncelleyen TEXT DEFAULT '',
            updated_at TEXT DEFAULT ''
        )
    """)


def _tr(s):
    return str(s or '').strip().replace('İ', 'i').replace('I', 'ı').lower()


def _sayi(v):
    try:
        f = float(str(v).replace(',', '.'))
        return f if f > 0 else None
    except (TypeError, ValueError):
        return None


def _tesis(d):
    s = _tr(d).replace(' ', '').replace('-', '')
    if s in ('tk2', ''):
        return 'TK2' if s else ''
    if s == 'tk1':
        return 'TK1'
    if any(s.startswith(x) for x in DISARIDA_TESIS):
        return None                  # dışarıda
    return str(d).strip()            # Pandora …


def _bolum_ad(tesis, e):
    """(anahtar, görünen ad) — TK2'de Forge bölüm anahtarı, diğerlerinde E'deki ad."""
    import import_excel as IE
    if tesis == 'TK2':
        b = IE.ana_veri_bolum(e)
        return (f'TK2|{b}', IE.ANA_VERI_ETIKET[b]) if b else (None, None)
    ad = ' '.join(str(e or '').split())
    if not ad or any(x in _tr(ad) for x in ('bölüm yok', 'olabilir', 'fason')):
        return None, None
    ad = ad[:1].upper() + ad[1:]
    return f'{tesis}|{_tr(ad)}', ad


def ana_veri_oku(yol=None):
    """Ana Veri → {'urun': {KOD: {...}}, 'ncs': {anahtar: {...}}, 'mtime'} (mtime önbellekli).
    urun: anahtar, tesis, bolum, sn (adet başı), kap1v (vardiya adedi)."""
    import openpyxl
    import import_excel as IE
    yol = yol or IE.ANA_VERI_YOL
    mt = os.path.getmtime(yol)
    if _ONBELLEK['anahtar'] == (yol, mt):
        return _ONBELLEK['veri']
    wb = openpyxl.load_workbook(yol, read_only=True, data_only=True)
    try:
        ws, kol = IE._ana_sayfa_bul(wb)
        if ws is None:
            raise ValueError("Ana Veri sayfası bulunamadı")

        def al(r, a):
            i = kol.get(a)
            return r[i] if (i is not None and i < len(r)) else None
        urun = {}
        for r in ws.iter_rows(min_row=2, values_only=True):
            kod = IE._norm_kod(al(r, 'kod'))
            if len(kod) < 2 or kod in urun:
                continue
            tesis = _tesis(al(r, 'tk'))
            if tesis is None:
                urun[kod] = {'disarida': True}
                continue
            anahtar, ad = _bolum_ad(tesis or 'TK2', al(r, 'bolum'))
            if not anahtar:
                urun[kod] = {'bolumsuz': True}
                continue
            g, h = _sayi(al(r, 'cevrim')), _sayi(al(r, 'kaynak'))
            sn = ((g or 0) + (h or 0)) if anahtar == 'TK2|kaynak' and (g or h) else g
            urun[kod] = {'anahtar': anahtar, 'tesis': tesis or 'TK2', 'bolum': ad,
                         'sn': sn, 'kap1v': _sayi(al(r, 'kapasite')) if not sn else None}
        ncs = _ncs_oku(wb)
    finally:
        wb.close()
    veri = {'urun': urun, 'ncs': ncs, 'mtime': datetime.fromtimestamp(mt).strftime('%Y-%m-%d %H:%M')}
    _ONBELLEK.update(anahtar=(yol, mt), veri=veri)
    return veri


def _is_gunu(ay):
    """'2026-09' → hafta içi gün sayısı (resmî tatil bilinmiyor)."""
    try:
        y, m = (int(x) for x in str(ay).split('-')[:2])
    except ValueError:
        return 22
    return sum(1 for d in range(1, calendar.monthrange(y, m)[1] + 1) if date(y, m, d).weekday() < 5)


def _ncs_oku(wb):
    """'Aylık NçS' → {anahtar: {'kisi', 'gunluk_saat', 'ay'}} (varsayılan kişi + saat)."""
    if 'Aylık NçS' not in wb.sheetnames:
        return {}
    ay, out, bas = '', {}, None
    for r in wb['Aylık NçS'].iter_rows(values_only=True):
        r = list(r) + [None] * 6
        if _tr(r[0]) == 'ay':
            ay = str(r[1] or '').strip()[:7]
            continue
        if _tr(r[0]) == 'tesis':
            bas = True
            continue
        if not bas or not r[0] or not r[1]:
            continue
        tesis = _tesis(r[0]) or ''
        e = r[1]
        if tesis == 'TK1' and 'tel' in _tr(e):
            e = 'Tel'
        anahtar, _ad = _bolum_ad(tesis, e)
        kisi, saat = _sayi(r[4]), _sayi(r[2])
        if anahtar and kisi:
            gs = round(saat / kisi / _is_gunu(ay), 2) if saat else None
            out[anahtar] = {'kisi': kisi, 'gunluk_saat': gs if gs and 3 <= gs <= 12 else None, 'ay': ay}
    return out


def ayarlar(conn):
    return {r[0]: {'kisi': r[1], 'gunluk_saat': r[2], 'verim': r[3], 'guncelleyen': r[4], 'updated_at': r[5]}
            for r in conn.execute("SELECT anahtar, kisi, gunluk_saat, verim, guncelleyen, updated_at "
                                  "FROM bolum_kapasite_ayar")}


def ayar_kaydet(conn, anahtar, kisi=None, gunluk_saat=None, verim=None, kullanici=''):
    eski = ayarlar(conn).get(anahtar, {})

    def sec(yeni, ad, alt, ust):
        if yeni is None or yeni == '':
            return eski.get(ad)
        v = float(yeni)
        if not alt <= v <= ust:
            raise ValueError(f'{ad} {alt}–{ust} arasında olmalı')
        return v
    conn.execute(
        "INSERT INTO bolum_kapasite_ayar (anahtar, kisi, gunluk_saat, verim, guncelleyen, updated_at) "
        "VALUES (?,?,?,?,?,?) ON CONFLICT(anahtar) DO UPDATE SET kisi=excluded.kisi, "
        "gunluk_saat=excluded.gunluk_saat, verim=excluded.verim, guncelleyen=excluded.guncelleyen, "
        "updated_at=excluded.updated_at",
        (anahtar, sec(kisi, 'kisi', 0, 500), sec(gunluk_saat, 'gunluk_saat', 1, 24),
         sec(verim, 'verim', 10, 150), kullanici, datetime.now().strftime('%Y-%m-%d %H:%M')))
    conn.commit()


def _haftalar(bugun, n):
    pzt = bugun - timedelta(days=bugun.weekday())
    return [{'no': f'H{(pzt + timedelta(days=7 * i)).isocalendar()[1]}',
             'bas': (pzt + timedelta(days=7 * i)).isoformat(),
             'bit': (pzt + timedelta(days=7 * i + 6)).isoformat()} for i in range(n)]


def hesapla(conn, hafta=4, bugun=None, yol=None):
    """Bölüm × hafta yük tablosu. Her hücre: kişi-saat (tanımlı + tahmini), adet, kod sayısı.
    Ayar/varsayılan bölüm satırına eklenir; gereken kişi + doluluk İSTEMCİDE hesaplanır
    (kişi kutusu değişince sayfa sunucuya gitmeden güncellensin)."""
    bugun = bugun or date.today()
    av = ana_veri_oku(yol)
    urun = av['urun']
    hft = _haftalar(bugun, hafta)
    sinir = [date.fromisoformat(h['bas']) for h in hft] + [date.fromisoformat(hft[-1]['bit']) + timedelta(days=1)]
    ay = ayarlar(conn)
    bolumler, atanmamis, sayac = {}, {}, {'tarihsiz': 0, 'ufuk_disi': 0, 'disarida': 0}
    talep_ts = None

    def bolum(anahtar, ad, tesis):
        if anahtar not in bolumler:
            a, n = ay.get(anahtar, {}), av['ncs'].get(anahtar, {})
            bolumler[anahtar] = {
                'anahtar': anahtar, 'ad': ad, 'tesis': tesis,
                'kisi': a.get('kisi') if a.get('kisi') is not None else (n.get('kisi') or 0),
                'gunluk_saat': a.get('gunluk_saat') or n.get('gunluk_saat') or VARSAYILAN_SAAT,
                'verim': a.get('verim') or VARSAYILAN_VERIM,
                'kisi_kaynagi': 'elle' if a.get('kisi') is not None else ('NçS ' + n['ay'] if n.get('kisi') else ''),
                'guncelleyen': a.get('guncelleyen') or '', 'gun': VARSAYILAN_GUN,
                'hucre': {k: {'saat_sn': 0.0, 'vardiya': 0.0, 'adet': 0.0, 'adet_suresiz': 0.0, 'kod': set(),
                              'suresiz_kod': set()} for k in ['gecikmis'] + [h['no'] for h in hft]},
                '_sn': [], '_kap': []}
        return bolumler[anahtar]

    for r in conn.execute("SELECT kod, kalan, bitis, senk_at FROM kapasite_talep WHERE kalan > 0"):
        talep_ts = talep_ts or r[3]
        kod = str(r[0] or '').strip().upper().replace(' ', '')
        if not kod or kod.startswith(HARIC_ONEK):
            continue
        u = urun.get(kod)
        if not u or u.get('bolumsuz'):
            a = atanmamis.setdefault(kod, {'kod': kod, 'adet': 0.0, 'emir': 0,
                                           'neden': 'Ana Veri\'de bölümsüz' if u else 'Ana Veri\'de yok'})
            a['adet'] += r[1]
            a['emir'] += 1
            continue
        if u.get('disarida'):
            sayac['disarida'] += 1
            continue
        try:
            bt = date.fromisoformat(str(r[2])[:10])
        except ValueError:
            sayac['tarihsiz'] += 1
            continue
        if bt < sinir[0]:
            hk = 'gecikmis'
        elif bt >= sinir[-1]:
            sayac['ufuk_disi'] += 1
            continue
        else:
            hk = next(hft[i]['no'] for i in range(len(hft)) if sinir[i] <= bt < sinir[i + 1])
        b = bolum(u['anahtar'], u['bolum'], u['tesis'])
        h = b['hucre'][hk]
        h['adet'] += r[1]
        h['kod'].add(kod)
        if u.get('sn'):
            h['saat_sn'] += r[1] * u['sn'] / 3600.0
        elif u.get('kap1v'):
            h['vardiya'] += r[1] / u['kap1v']
        else:
            h['adet_suresiz'] += r[1]
            h['suresiz_kod'].add(kod)

    # Bölümün ortanca birim süresi (süresiz ürünü tahmin etmek için) — Ana Veri'deki
    # TÜM ürünlerinden, yalnız ihtiyacı olanlardan değil
    for u in urun.values():
        b = bolumler.get(u.get('anahtar'))
        if b is not None:
            (b['_sn'] if u.get('sn') else b['_kap'] if u.get('kap1v') else []).append(u.get('sn') or u.get('kap1v'))
    for b in bolumler.values():
        b['ortanca_sn'] = round(statistics.median(b['_sn']), 1) if b['_sn'] else None
        b['ortanca_kap1v'] = round(statistics.median(b['_kap']), 1) if b['_kap'] else None
        tanimli = suresiz = 0.0
        for h in b['hucre'].values():
            h['kod_sayisi'], h['suresiz_kod_sayisi'] = len(h.pop('kod')), len(h.pop('suresiz_kod'))
            tanimli += h['adet'] - h['adet_suresiz']
            suresiz += h['adet_suresiz']
            h['adet'], h['adet_suresiz'] = round(h['adet']), round(h['adet_suresiz'])
            h['saat_sn'], h['vardiya'] = round(h['saat_sn'], 2), round(h['vardiya'], 3)
        b['kapsama'] = round(100.0 * tanimli / (tanimli + suresiz), 1) if (tanimli + suresiz) else None
        del b['_sn'], b['_kap']
    sira = {f'TK2|{k}': i for i, k in enumerate(TK2_SIRA)}
    liste = sorted(bolumler.values(), key=lambda b: ({'TK2': 0, 'TK1': 1}.get(b['tesis'], 2),
                                                     sira.get(b['anahtar'], 99), b['ad']))
    # Ana Veri'de olmayan kodlardan yalnız ÜRETİM (P) olanlar uyarıdır — satın alma (A)
    # kodunun OPR'si bizim kapasitemizi ilgilendirmez (kapasite_urun = AS400 BARTF0)
    try:
        prov = {r[0]: r[1] for r in conn.execute("SELECT kod, prov FROM kapasite_urun")}
    except Exception:
        prov = {}
    at = sorted((a for a in atanmamis.values() if prov.get(a['kod'], 'P') == 'P'), key=lambda a: -a['adet'])
    return {'haftalar': hft, 'bolumler': liste, 'ana_veri': av['mtime'], 'talep_ts': talep_ts,
            'varsayilan': {'gun': VARSAYILAN_GUN, 'saat': VARSAYILAN_SAAT, 'verim': VARSAYILAN_VERIM},
            'atanmamis': {'kod': len(at), 'adet': round(sum(a['adet'] for a in at)),
                          'ilk': [{**a, 'adet': round(a['adet'])} for a in at[:40]]},
            'sayac': sayac}


def detay(conn, anahtar, hafta_no, hafta=4, bugun=None, yol=None, n=40):
    """Bir bölümün bir haftasındaki ürünler (yüke göre azalan): kod, adet, süre, kişi-saat."""
    bugun = bugun or date.today()
    urun = ana_veri_oku(yol)['urun']
    hft = _haftalar(bugun, hafta)
    bas0 = date.fromisoformat(hft[0]['bas'])
    if hafta_no == 'gecikmis':
        aralik = (date.min, bas0)
    else:
        h = next((x for x in hft if x['no'] == hafta_no), None)
        if not h:
            return []
        aralik = (date.fromisoformat(h['bas']), date.fromisoformat(h['bit']) + timedelta(days=1))
    out = {}
    for r in conn.execute("SELECT kod, kalan, bitis, durum, emir_no FROM kapasite_talep WHERE kalan > 0"):
        kod = str(r[0] or '').strip().upper().replace(' ', '')
        u = urun.get(kod)
        if not u or u.get('anahtar') != anahtar:
            continue
        try:
            bt = date.fromisoformat(str(r[2])[:10])
        except ValueError:
            continue
        if not aralik[0] <= bt < aralik[1]:
            continue
        d = out.setdefault(kod, {'kod': kod, 'adet': 0.0, 'emir': 0, 'launch': 0.0, 'en_eski': r[2],
                                 'sn': u.get('sn'), 'kap1v': u.get('kap1v')})
        d['adet'] += r[1]
        d['emir'] += 1
        if str(r[3]) != '10':
            d['launch'] += r[1]
        d['en_eski'] = min(d['en_eski'], r[2])
    for d in out.values():
        d['saat_sn'] = round(d['adet'] * d['sn'] / 3600.0, 2) if d['sn'] else None
        d['vardiya'] = round(d['adet'] / d['kap1v'], 2) if d['kap1v'] else None
        d['adet'], d['launch'] = round(d['adet']), round(d['launch'])
    # Yüke göre: süreli ürünler kişi-saatiyle, süresizler en sonda adetle
    return sorted(out.values(), key=lambda d: (
        -(d['saat_sn'] or (d['vardiya'] or 0) * VARSAYILAN_SAAT), -d['adet']))[:n]
