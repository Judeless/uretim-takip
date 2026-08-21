# -*- coding: utf-8 -*-
"""
test_teyit.py — _zaten_teyitli() regresyon testleri.  python as400\\test_teyit.py

NEDEN VAR: bu fonksiyon ERP stoğunu belirliyor. Yanlış 'kesin' derse gerçek
üretim hiç teyit edilmez (stok eksik); yanlış None derse aynı üretim ikinci kez
teyit edilir (stok fazla). İkisi de SESSİZ hatadır — kimse fark etmez.

Her vaka kodda yorumla belgelenmiş GERÇEK bir saha olayıdır. Bir tuzağı
kapatırken diğerini açmadığımızı ancak hepsi birden geçince biliriz.

AS400 bağlantısı GEREKTİRMEZ — saf fonksiyon, elle kurulmuş hareket listeleri.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import launch_esle as le   # noqa: E402


def H(tarih, adet, tur='RPR'):
    return {'tarih': tarih, 'adet': float(adet), 'tur': tur, 'launch': '—'}


VAKALAR = [
    # ── ad, hareketler, uretim_tarihi, adet, gecmis, beklenen_durum, beklenen_kalan
    (
        '1865W · aynı-gün hareket ÖNCEKİ günün teyidi (2026-07-20)\n'
        '        17.07 üretimi 30, aynı gün 68\'lik hareket var ama 68 = 16.07\'nin\n'
        '        üretimi → kanıt SAYILMAZ, bizim 30 hâlâ teyitsiz',
        [H('2026-07-17', 68)], '2026-07-17', 30,
        {'2026-07-16': {68}, '2026-07-17': {30}}, None, None,
    ),
    (
        'Aynı-gün İSTİSNASI · birebir adet + başka günle açıklanamıyor → kesin',
        [H('2026-07-17', 30)], '2026-07-17', 30,
        {'2026-07-17': {30}}, 'kesin', None,
    ),
    (
        '3680GAW · bölünmüş teyit, sonraki günlerde 1 + 275 = 276',
        [H('2026-07-18', 1), H('2026-07-19', 275)], '2026-07-17', 276,
        {'2026-07-17': {276}}, 'kesin', None,
    ),
    (
        '94.LTK.890 · bölünmüş teyit AYNI GÜN, 23 + 57 = 80 (2026-08-06)',
        [H('2026-08-06', 23), H('2026-08-06', 57)], '2026-08-06', 80,
        {'2026-08-06': {80}}, 'kesin', None,
    ),
    (
        '94.LTK.487/10 · İKİ GÜN AYNI ADET (2026-07-30 · ERP\'ye 30 FAZLA girmişti)\n'
        '        24.07 üretimi 30 → 27.07\'de teyit. 29.07 üretimi de 30 → 29.07\'de\n'
        '        teyit. İkinci hareket "24.07\'nin teyidi" sayılırsa satır teyitsiz\n'
        '        görünür ve ikinci kez gönderilir. BİR GÜN BİR KEZ karşılanmalı.',
        [H('2026-07-27', 30), H('2026-07-29', 30)], '2026-07-29', 30,
        {'2026-07-24': {30}, '2026-07-29': {30}}, 'kesin', None,
    ),
    (
        '94.LTK.215 · ÜÇ HAREKET DE AYNI GÜN (2026-08-21 · ERP\'ye 100 FAZLA girdi)\n'
        '        19.08 100 + 20.08 100 üretim; 20.08\'de 3×100 hareket.\n'
        '        Eski kod üçünü de "19.08\'in teyidi" sayıp eliyordu → 20.08 satırı\n'
        '        teyitsiz görünüyor ve tekrar gönderilebiliyordu.',
        [H('2026-08-20', 100), H('2026-08-20', 100), H('2026-08-20', 100, 'CFI')],
        '2026-08-20', 100, {'2026-08-19': {100}, '2026-08-20': {100}}, 'kesin', None,
    ),
    (
        '94.LTK.215 · aynı verinin 19.08 satırı — o gün de teyitli görünmeli',
        [H('2026-08-20', 100), H('2026-08-20', 100), H('2026-08-20', 100, 'CFI')],
        '2026-08-19', 100, {'2026-08-19': {100}, '2026-08-20': {100}}, 'kesin', None,
    ),
    (
        '94.LTK.09 · KISMİ TEYİT (2026-08-21)\n'
        '        20.08\'de iki kalem üretim 18 + 17 = 35; yalnız 18 teyit edilmiş.\n'
        '        Eski kod None diyordu → panel 35 öneriyordu → 18 FAZLA giderdi.\n'
        '        Doğrusu: kısmi, kalan 17.',
        [H('2026-08-21', 18)], '2026-08-20', 35,
        {'2026-08-20': {18, 17}}, 'kismi', 17,
    ),
    (
        'Sadece hurda · adet 0 ise teyit kanıtı da yok (2026-07-31)\n'
        '        Aksi hâlde "toplam >= adet" kuralı 0 ile DAİMA sağlanır ve satır\n'
        '        sahte "olası" işaretlenip Dikkat kovasını şişirirdi.',
        [H('2026-08-21', 50)], '2026-08-20', 0, {}, None, None,
    ),
    (
        'Fazla hareket · toplam adetten BÜYÜK → olası (elle kontrol)',
        [H('2026-08-21', 40), H('2026-08-21', 30)], '2026-08-20', 50,
        {'2026-08-20': {50}}, 'olasi', None,
    ),
    (
        'Hiç hareket yok → teyitsiz',
        [], '2026-08-20', 25, {'2026-08-20': {25}}, None, None,
    ),
    (
        'Önceki gün teyidi tek başına duruyor · bizim gün teyitsiz kalmalı\n'
        '        Pazartesi girilen 68 = cumanın 68\'lik üretimi; bizim 30 teyitsiz.',
        [H('2026-08-24', 68)], '2026-08-24', 30,
        {'2026-08-21': {68}, '2026-08-24': {30}}, None, None,
    ),
]


def calistir():
    gecti = basarisiz = 0
    for ad, hrk, gun, adet, gecmis, bek_durum, bek_kalan in VAKALAR:
        durum, ilgili = le._zaten_teyitli(hrk, gun, adet, gecmis)
        tamam = (durum == bek_durum)
        kalan = None
        if durum == 'kismi':
            kalan = adet - sum(h['adet'] for h in ilgili)
            if bek_kalan is not None and abs(kalan - bek_kalan) > 0.001:
                tamam = False
        bas = ad.split('\n')[0]
        if tamam:
            gecti += 1
            print(f'  ✓ {bas}')
        else:
            basarisiz += 1
            print(f'  ✗ {bas}')
            for satir in ad.split('\n')[1:]:
                print(f'    {satir.strip()}')
            print(f'      beklenen: {bek_durum!r}' +
                  (f' (kalan {bek_kalan})' if bek_kalan is not None else ''))
            print(f'      gelen   : {durum!r}' +
                  (f' (kalan {kalan})' if kalan is not None else '') +
                  f'  · kanıt {len(ilgili)} hareket')
    print()
    print(f'  {gecti} geçti · {basarisiz} başarısız  ({len(VAKALAR)} vaka)')
    return basarisiz


if __name__ == '__main__':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    print('=' * 74)
    print('TEYİT KANIT MANTIĞI · REGRESYON TESTLERİ')
    print('=' * 74)
    sys.exit(1 if calistir() else 0)
