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

OKUMA NOTU — PENCERE KENARI: teyit normalde ÜRETİMİN ERTESİ GÜNÜ girilir. Bu
yüzden hareketler [baş, bit+3] aralığında toplanır. Pencerenin BAŞINDAN önceki
üretimlerin teyidi pencere içine düşerse fazla görünür — bu yüzden tek referans
dökümünde TÜM hareketler tarihiyle listelenir; kararı insan verir.
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

    print('\n' + '-' * 78)
    print(f'    ÜRETİM          : {u_top:g}')
    print(f'    ERP TEYİT       : {h_top:g}')
    fark = h_top - u_top
    print(f'    FARK            : {fark:+g}   ' +
          ('← FAZLA TEYİT' if fark > 0.001 else
           '← eksik teyit' if fark < -0.001 else '✓ eşit'))
    print(f'    bunun {kendi_top:g} adedi PANELDEN gönderilmiş, '
          f'{h_top - kendi_top:g} adedi elle girilmiş görünüyor')
    if fark > 0.001:
        print('\n    NOT: pencere başından önceki bir üretimin teyidi bu aralığa')
        print('    düşmüş olabilir. Yukarıdaki hareket tarihlerini üretim')
        print('    tarihleriyle eşleştirip karar verin; --tara ile geniş pencere deneyin.')
    print()


def tara(gun=14):
    bit = date.today()
    bas = bit - timedelta(days=gun)
    bas_s, bit_s = bas.isoformat(), bit.isoformat()
    uretim, _ = _uretim(bas_s, bit_s)
    if not uretim:
        print('Bu aralıkta üretim kaydı yok.')
        return
    hrk = _hareketler(set(uretim.keys()))
    sinir = (bit + timedelta(days=3)).isoformat()

    satir = []
    for k, gunler in uretim.items():
        u_top = sum(gunler.values())
        if u_top <= 0:
            continue
        h_top = sum(h['adet'] for h in hrk.get(k, [])
                    if bas_s <= h['tarih'] <= sinir)
        satir.append((h_top - u_top, k, u_top, h_top))
    satir.sort(reverse=True)

    print('=' * 78)
    print(f'FAZLA TEYİT TARAMASI · {bas_s} → {bit_s}   ({len(satir)} referans)')
    print('=' * 78)
    print(f'{"FARK":>8}  {"ÜRETİM":>8}  {"TEYİT":>8}   REFERANS')
    print('-' * 78)
    fazla = [s for s in satir if s[0] > 0.001]
    for fark, k, u, h in fazla[:25]:
        print(f'{fark:>+8g}  {u:>8g}  {h:>8g}   {k}')
    if not fazla:
        print('  ✓ fazla teyit görünmüyor')
    print('-' * 78)
    print(f'  {len(fazla)} referansta teyit üretimden fazla · '
          f'toplam fazlalık {sum(s[0] for s in fazla):g} adet')
    print('\n  Ayrıntı için:  python as400\\teyit_denetim.py <REFERANS>')
    print('  UYARI: pencere kenarı yanılgısı — pencereden ÖNCEKİ üretimin teyidi')
    print('  içeri düşerse fazla görünür. Tek referans dökümünde tarihlere bakın.')
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
