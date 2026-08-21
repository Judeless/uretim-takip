# -*- coding: utf-8 -*-
"""
teyit_denetim.py — ÜRETİM ile VERİLEN TEYİT'i karşılaştırır. SALT OKUNUR.

Neden var (kullanıcı 2026-08-21): 94.LTK.215'te 19.08 + 20.08 toplam 200 adet
üretim var ama ERP'de 300 adet teyit görünüyor — 100 adet fazla. Bu ilk değil:
  · 94.LTK.340   → aynı kodun iki yazımı (büyük/küçük harf) adedi katlamıştı, 42 fazla
  · 94.LTK.487/10 → hareket başka güne atfedilip ikinci kez teyit gitmişti, 30 fazla
Her seferinde gözle fark edildi. Bu araç aynı kontrolü SİSTEMLİ yapar.

⚠ HİÇBİR ŞEY GÖNDERMEZ, HİÇBİR ŞEY YAZMAZ. Yalnız okur ve tablo basar.

KULLANIM (sunucuda, proje klasöründe):
    python as400\\teyit_denetim.py 94.LTK.215        → tek referansın dökümü
    python as400\\teyit_denetim.py --tara            → son 14 gün, fazla teyitleri sırala
    python as400\\teyit_denetim.py --tara 30         → pencereyi değiştir

YÖNTEM — GÜN BAZLI, PENCERE TOPLAMI DEĞİL. İlk sürüm pencere içindeki üretim
toplamı ile hareket toplamını karşılaştırıyordu; pencerenin BAŞINDAN önceki
üretimlerin teyidi içeri düşünce her referans "fazla" görünüyordu (ilk taramada
43 referans / 6683 adet SAHTE fazlalık çıktı — 10.130.3680GAW'da 22.07 tarihli
688'lik hareket 21.07 üretiminin teyidiydi, o üretim pencere dışındaydı).

Artık her ÜRETİM GÜNÜ için sistemin kendi kararı sorulur (launch_esle
._zaten_teyitli — regresyon testli, bkz. test_teyit.py):
    kesin → o gün tam teyitli        olasi → o güne FAZLA teyit
    kismi → günün bir kısmı teyitli  None  → o gün hiç teyit almamış
Geçmiş penceresi 20 gün geriden başlar; aksi hâlde önceki günlerin teyidi bizim
güne atfedilir ve aynı yanılgı geri gelir.
"""
import os
import sqlite3
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import launch_esle as le   # noqa: E402


def _uretim(bas, bit):
    """{kanonik referans: {tarih: adet}} — launch'lı bölümler, TK2, rework hariç.
    launch_esle.gun_uretimi ile AYNI kuralları uygular (rework filtresi dahil)."""
    conn = sqlite3.connect(le.URETIM_DB)
    conn.row_factory = sqlite3.Row
    rows = conn.execute(f"""
        SELECT v.tarih, u.referans_kodu ref, COALESCE(u.ok_adet,0) ok,
               COALESCE(u.nok_adet,0) nok, COALESCE(u.aciklama,'') aciklama,
               v.operator_adi, v.robot_no
        FROM uretim_kayitlari u JOIN vardiyalar v ON v.id = u.vardiya_id
        WHERE v.tarih >= ? AND v.tarih <= ?
          AND COALESCE(v.lokasyon,'TK2') = ?
          AND COALESCE(v.bolum,'kaynak') IN ({','.join('?' * len(le.LAUNCH_BOLUMLERI))})
          AND u.referans_kodu IS NOT NULL AND u.referans_kodu != ''
        ORDER BY v.tarih""",
        (bas, bit, le.LOKASYON) + le.LAUNCH_BOLUMLERI).fetchall()
    conn.close()
    out = defaultdict(lambda: defaultdict(float))
    kayit = defaultdict(list)
    for r in rows:
        if le.rework_mi(r['aciklama'], r['ref']):
            continue
        k = le.kanonik(le.bosluk_nokta(r['ref']))
        out[k][r['tarih']] += r['ok']
        kayit[k].append({'tarih': r['tarih'], 'adet': r['ok'], 'hurda': r['nok'],
                         'operator': r['operator_adi'], 'hat': r['robot_no']})
    return out, kayit


def _bizim_log(bas, bit):
    """Panelden GÖNDERDİĞİMİZ teyitler. AS400'deki hareketin bizden mi yoksa
    elle mi girildiğini ayırmanın tek yolu bu — fazlalığın kaynağını buradan
    anlarız."""
    conn = sqlite3.connect(le.URETIM_DB)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            "SELECT created_at, uretim_tarihi, yil, launch_no, referans, adet, sonuc, olusturan "
            "FROM as400_teyit_log WHERE uretim_tarihi >= ? AND uretim_tarihi <= ? "
            "ORDER BY id", (bas, bit)).fetchall()
    except sqlite3.OperationalError:
        rows = []
    conn.close()
    out = defaultdict(list)
    for r in rows:
        out[le.kanonik(r['referans'] or '')].append(dict(r))
    return out


def _hareketler(kodlar):
    """AS400 RPR/CFI hareketleri. Bağlantı yoksa boş döner (araç yine de üretim
    ve kendi log'umuzu gösterir — yarısı hiç yoktan iyidir)."""
    try:
        return le.teyit_hareketleri(sorted(kodlar))
    except Exception as e:
        print(f'  ! AS400 hareketleri okunamadı: {e}')
        print('    (üretim ve kendi gönderim log\'umuz yine de aşağıda)')
        return {}


def tek_referans(kod, gun=45):
    bit = date.today()
    bas = bit - timedelta(days=gun)
    bas_s, bit_s = bas.isoformat(), bit.isoformat()
    k = le.kanonik(le.bosluk_nokta(kod))
    uretim, kayitlar = _uretim(bas_s, bit_s)
    log = _bizim_log(bas_s, bit_s)
    hrk = _hareketler({kod, kod.upper(), le.bosluk_nokta(kod)})

    print('=' * 78)
    print(f'TEYİT DENETİMİ · {kod}   (son {gun} gün: {bas_s} → {bit_s})')
    print('=' * 78)

    u = uretim.get(k, {})
    u_top = sum(u.values())
    print(f'\n[1] ÜRETİM  (uretim.db · rework hariç)   TOPLAM {u_top:g} adet')
    if not u:
        print('    kayıt yok')
    for t in sorted(u):
        detay = [c for c in kayitlar[k] if c['tarih'] == t]
        print(f"    {t}  {u[t]:>7g} adet   " +
              ' · '.join(f"{c['adet']:g} ({c['operator']}/{c['hat']})" for c in detay))

    print('\n[2] BİZİM GÖNDERDİĞİMİZ  (as400_teyit_log)')
    kendi = [r for r in log.get(k, []) if r['sonuc'] == 'ok']
    if not kendi:
        print('    panelden hiç gönderim yok → ERP\'deki hareketler ELLE girilmiş')
    for r in kendi:
        print(f"    üretim {r['uretim_tarihi']}  →  {r['adet']:>6} adet  "
              f"launch {r['yil']}-{r['launch_no']}  ({r['created_at'][:16]}, {r['olusturan']})")
    kendi_top = sum(float(r['adet'] or 0) for r in kendi)

    print('\n[3] ERP HAREKETLERİ  (BMMAF0)')
    tum = []
    for anahtar, hs in hrk.items():
        for h in hs:
            if bas_s <= h['tarih'] <= (bit + timedelta(days=3)).isoformat():
                tum.append(h)
    tum.sort(key=lambda x: x['tarih'])
    h_top = sum(h['adet'] for h in tum)
    if not tum:
        print('    hareket yok / okunamadı')
    for h in tum:
        print(f"    {h['tarih']}  {h['adet']:>7g} adet   {h['tur']:<4} {h['launch']}")

    # [4] Sistemin gün bazlı kararı — panelin gördüğü şey budur
    if u:
        gecmis = le.uretim_gecmisi(
            (bas - timedelta(days=20)).isoformat(), bit_s).get(k, {})
        print('\n[4] SİSTEMİN KARARI  (panelde bu satır ne görünüyor)')
        for t in sorted(u):
            durum, ilgili = le._zaten_teyitli(tum, t, u[t], gecmis)
            kanit = sum(h['adet'] for h in ilgili)
            etiket = {'kesin': '✓ tam teyitli', 'olasi': '⚠ FAZLA teyit',
                      'kismi': '◐ kısmi teyitli', None: '· teyit almamış'}[durum]
            ek = ''
            if durum == 'kismi':
                ek = f'  → kalan {u[t] - kanit:g}'
            elif durum == 'olasi':
                ek = f'  → fazla {kanit - u[t]:g}'
            print(f'    {t}  üretim {u[t]:>7g} · kanıt {kanit:>7g}   {etiket}{ek}')

    print('\n' + '-' * 78)
    print(f'    ÜRETİM          : {u_top:g}')
    print(f'    ERP TEYİT       : {h_top:g}')
    fark = h_top - u_top
    print(f'    FARK            : {fark:+g}   ' +
          ('← FAZLA TEYİT' if fark > 0.001 else
           '← eksik teyit' if fark < -0.001 else '✓ eşit'))
    print(f'    bunun {kendi_top:g} adedi PANELDEN gönderilmiş, '
          f'{h_top - kendi_top:g} adedi elle girilmiş görünüyor')
    if abs(fark) > 0.001:
        print('\n    NOT: bu iki toplam KABA bir göstergedir — pencere başından önceki')
        print('    üretimlerin teyidi bu aralığa düşerse fark şişer. GERÇEK karar')
        print('    yukarıdaki [4] bölümündedir (gün bazlı, sistemin kendi mantığı).')
    print()


def _gun_karari(gun=14):
    """Her (referans, üretim günü) için sistemin teyit kararı.
    Döner: [(kod, tarih, adet, durum, kanit_toplam)]"""
    bit = date.today()
    bas = bit - timedelta(days=gun)
    bas_s, bit_s = bas.isoformat(), bit.isoformat()
    uretim, _ = _uretim(bas_s, bit_s)
    if not uretim:
        return [], bas_s, bit_s
    # GEÇMİŞ 20 GÜN GERİDEN: _zaten_teyitli önceki günlerin üretimini bilmezse
    # o günlerin teyidini BİZİM güne kanıt sayar ve fazlalık gizlenir/uydurulur.
    gecmis = le.uretim_gecmisi((bas - timedelta(days=20)).isoformat(), bit_s)
    hrk = _hareketler(set(uretim.keys()))
    out = []
    for k, gunler in uretim.items():
        hareketler = hrk.get(k, [])
        for t, adet in sorted(gunler.items()):
            if adet <= 0:
                continue
            durum, ilgili = le._zaten_teyitli(hareketler, t, adet, gecmis.get(k))
            out.append((k, t, adet, durum, sum(h['adet'] for h in ilgili)))
    return out, bas_s, bit_s


def tara(gun=14):
    kararlar, bas_s, bit_s = _gun_karari(gun)
    if not kararlar:
        print('Bu aralıkta üretim kaydı yok.')
        return
    fazla = [k for k in kararlar if k[3] == 'olasi']
    kismi = [k for k in kararlar if k[3] == 'kismi']
    yok   = [k for k in kararlar if k[3] is None]
    tam   = [k for k in kararlar if k[3] == 'kesin']

    print('=' * 78)
    print(f'TEYİT DENETİMİ · {bas_s} → {bit_s}   ({len(kararlar)} referans-gün)')
    print('=' * 78)
    print(f'  ✓ tam teyitli      : {len(tam)}')
    print(f'  ⚠ FAZLA teyit      : {len(fazla)}')
    print(f'  ◐ kısmi teyitli    : {len(kismi)}')
    print(f'  · teyit almamış    : {len(yok)}')

    if fazla:
        print('\n⚠ FAZLA TEYİT — o güne üretilenden çok teyit hareketi var')
        print(f'{"FAZLA":>8}  {"ÜRETİM":>8}  {"TEYİT":>8}  TARİH        REFERANS')
        print('-' * 78)
        for k, t, adet, _d, kanit in sorted(fazla, key=lambda x: -(x[4] - x[2])):
            print(f'{kanit - adet:>+8g}  {adet:>8g}  {kanit:>8g}  {t}   {k}')
        print('-' * 78)
        print(f'  toplam fazlalık {sum(k[4] - k[2] for k in fazla):g} adet')

    if kismi:
        print('\n◐ KISMİ TEYİTLİ — kalan adet henüz girilmemiş')
        print(f'{"KALAN":>8}  {"ÜRETİM":>8}  {"TEYİTLİ":>8}  TARİH        REFERANS')
        print('-' * 78)
        for k, t, adet, _d, kanit in sorted(kismi, key=lambda x: -(x[2] - x[4]))[:25]:
            print(f'{adet - kanit:>8g}  {adet:>8g}  {kanit:>8g}  {t}   {k}')

    print('\n  Ayrıntı için:  python as400\\teyit_denetim.py <REFERANS>')
    print()


if __name__ == '__main__':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    arg = sys.argv[1] if len(sys.argv) > 1 else '--tara'
    if arg == '--tara':
        tara(int(sys.argv[2]) if len(sys.argv) > 2 else 14)
    elif arg.startswith('-'):
        print(__doc__)
    else:
        tek_referans(arg, int(sys.argv[2]) if len(sys.argv) > 2 else 45)
