# -*- coding: utf-8 -*-
"""KOD SORGU (kullanıcı 2026-10-26) — AS400 07-10-01 ekranının telefondan karşılığı.

"Bilgisayar başında değilken AS400'deki verilere ulaşmak için bir modül yapalım.
Genelde AS400'ü 07 10 01 ekranı için, bir kod yazıp stoklarını görmek ve alt
parçada kısıtını görmek için kullanıyoruz."

Tek kod için CANLI ERP okuması (önbellek yok):
  · ürün kartı      BARTF0  — açıklama, P/A, hayali (A0PROD='2')
  · depo stokları   BSMAF0  — ekrandaki G GI = 01D + MDT (kaynak_plan_kontrol.stok_ggi)
  · açık emirler    XPRO90  — OPR (10) / launch (40/45/50), kalan, teslim tarihi
  · ürün ağacı      BSPEF2  — hayali alt parçalar 3 seviyeye kadar açılır; her parçanın
                              SAYILAN stoğu ve kaç ürün çıkardığı → üretilebilir adet
                              ve KISITLAYAN parça
Stok kuralları plan modülleriyle AYNI fonksiyondan (kaynak_plan_kontrol.hesapla):
hangi depoların sayıldığı ürünün bölümünün plan profilinden gelir (app.py'de seçilir).
Öneri listesi ERP'ye gitmez: kapasite_urun (BARTF0'ın yerel kopyası) üzerinden."""
import time

import kaynak_plan_kontrol as kp

DURUM_AD = {'10': 'OPR', '40': 'Launch', '45': 'Launch (TK1)', '50': 'Launch (TK1)'}
GGI_DEPOLAR = ('01D', 'MDT')          # 07-10-01 ekranındaki G GI (bkz. kp.stok_ggi)
HAYALI_SEVIYE = 3


def kod_normal(kod):
    return str(kod or '').strip().upper().replace(' ', '')


def urun_kartlari(cn, kodlar):
    """{kod: {'aciklama', 'prov', 'hayali'}} — BARTF0."""
    out = {}
    kodlar = [k for k in kodlar if kp.GUVENLI_KOD.match(k)]
    cu = cn.cursor()
    for i in range(0, len(kodlar), 50):
        grup = kodlar[i:i + 50]
        cu.execute("SELECT TRIM(A0ARTI), TRIM(A0ARDS), TRIM(A0PROV), TRIM(A0PROD) FROM TKC0301F.BARTF0 "
                   f"WHERE A0ARTI IN ({','.join('?' * len(grup))})", grup)
        for kod, ad, prov, prod in cu.fetchall():
            out[kod] = {'aciklama': ad or '', 'prov': prov or '',
                        'hayali': (prod or '') == kp.HAYALI_BAYRAK}
    return out


def oneriler(conn, q, limit=15):
    """Yazılan parçaya uyan kodlar (yerel kapasite_urun): tam > sonu > başı > içinde.
    Kısa yazım (LTK.685 → 94.LTK.685) ve açıklamada arama da çalışır."""
    q = kod_normal(q)
    if len(q) < 2:
        return []
    desen = '%' + q.replace('%', '').replace('_', '') + '%'
    try:
        satirlar = conn.execute(
            "SELECT kod, aciklama, prov FROM kapasite_urun "
            "WHERE COALESCE(iptal,0)=0 AND (UPPER(REPLACE(kod,' ','')) LIKE ? OR UPPER(aciklama) LIKE ?) "
            "LIMIT 400", (desen, desen)).fetchall()
    except Exception:
        return []

    def sira(r):
        k = kod_normal(r[0])
        return (0 if k == q else 1 if k.endswith('.' + q) or k.endswith(q) else
                2 if k.startswith(q) else 3 if q in k else 4, len(k), k)
    return [{'kod': r[0], 'aciklama': r[1] or '', 'prov': r[2] or ''}
            for r in sorted(satirlar, key=sira)[:limit]]


def sorgula(cn, kod, sayilan, gosterilen, eksi_dusulen, kisit_disi=(), acilmayan_onekler=()):
    """Tek kodun 07-10-01 görünümü. cn: açık ERP bağlantısı (kapatmak çağıranın işi)."""
    t0 = time.time()
    kod = kod_normal(kod)
    kart = urun_kartlari(cn, [kod]).get(kod)
    if not kart:
        return {'kod': kod, 'bulundu': False, 'sure_ms': int((time.time() - t0) * 1000)}

    depolar = kp.stoklar(cn, [kod]).get(kod, {})
    stok = {'ggi': sum(depolar.get(d, 0) for d in GGI_DEPOLAR),
            'depolar': {d: q for d, q in sorted(depolar.items()) if q}}

    o = kp.opr_ihtiyaclari(cn, [kod], ufuk_gun=3650).get(kod) or {}
    emirler = [{'no': s['opr'], 'durum': s['durum'], 'durum_ad': DURUM_AD.get(s['durum'], s['durum']),
                'adet': s['adet'], 'teyit': s['teyit'], 'kalan': s['kalan'], 'tarih': s['tarih']}
               for s in o.get('satirlar') or []]
    emir_ozet = {'opr': sum(e['kalan'] for e in emirler if e['durum'] == '10'),
                 'launch': sum(e['kalan'] for e in emirler if e['durum'] != '10'),
                 'gecikmis': o.get('gecikmis') or 0}

    agac, iz = kp.urun_agaci_hayali(cn, [kod], HAYALI_SEVIYE, tuple(acilmayan_onekler or ()))
    alt = sorted({a for a, _q, _u in agac.get(kod, [])})
    alt_stok = kp.stoklar(cn, alt) if alt else {}
    alt_kart = urun_kartlari(cn, alt) if alt else {}
    s = {'kaynak_kod': kod}
    kp.hesapla([s], agac, alt_stok, None, sayilan=sayilan, gosterilen=gosterilen, iz=iz,
               kisit_disi=kisit_disi, eksi_dusulen=eksi_dusulen)
    parcalar = []
    en_az = s.get('uretilebilir')
    for p in s.get('parcalar') or []:
        p = dict(p)
        p['aciklama'] = (alt_kart.get(p['kod']) or {}).get('aciklama', '')
        # Aynı en düşük adedi veren HER parça kısıttır (hesapla tek kod döndürür)
        p['kisit'] = en_az is not None and p['kapasite'] is not None and p['kapasite'] <= en_az
        parcalar.append(p)
    # Kısıt sırası: en az ürün çıkaran en üstte; sayılmayan (hayali / kısıt dışı) sonda
    parcalar.sort(key=lambda p: (p['kapasite'] is None, p['kapasite'] if p['kapasite'] is not None else 0, p['kod']))
    return {
        'kod': kod, 'bulundu': True, **kart,
        'stok': stok, 'emirler': emirler, 'emir_ozet': emir_ozet,
        'agac': {'durum': s.get('durum'), 'uretilebilir': s.get('uretilebilir'),
                 'kisitlayan': s.get('kisitlayan') or '', 'eksi_var': bool(s.get('eksi_var')),
                 'parcalar': parcalar},
        'sure_ms': int((time.time() - t0) * 1000),
    }
