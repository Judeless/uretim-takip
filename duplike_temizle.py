# -*- coding: utf-8 -*-
"""Duplike referans yazimlarini tekillestirir (2026-07-30).

NEDEN: Ayni urun birden fazla yazimla kayitliysa (94.LTK.858 / 94.ltk.858,
93.TK.469 / 93.TK.469.) iki somut zarar dogar:
  1) Cycle time bir yazimda tanimli, digerinde bos kalir -> operatorun yanlis
     adet girisine karsi kapasite freni O KAYITTA calismaz.
  2) Kapasite sorgusu bir zamanlar JOIN'di ve duplike yazim satiri COGALTIP
     adedi katliyordu (94.LTK.340'ta 42 -> 84 CFI). Sorgu artik skaler alt
     sorgu, ama duplike yazimin kendisi hala risk kaynagi.

TUTULACAK YAZIM — sirayla ilk uyan kural karar verir:
  1. Cycle time TAM OLARAK BIR yazimda tanimliysa      -> o yazim   (kullanici karari)
  2. ERP article master'da (BARTF0) TAM OLARAK BIR      -> o yazim
     yazim tanimliysa
  3. Yazim temizligi skoru dusuk olan                   -> o yazim
     (cift nokta / cift bosluk / sonda noktalama / bosluk / kucuk harf = ceza)
  4. Beraberlik: uretim kaydi cok olan, o da esitse dusuk id

Tutulan satirdaki BOS alanlar silinecek satirdan devralinir (sure, aciklama,
sure teyidi, bukum operasyonu) — bilgi kaybi olmaz.
Silinen yazimi kullanan GECMIS URETIM KAYITLARI tutulan yazima tasinir; yoksa
o kayitlar referanssiz kalir ve cycle time'lari kaybolur.

KULLANIM (sunucuda, C:\\cofle\\uretim_takip icinde):
    python duplike_temizle.py            # sadece rapor, hicbir sey degismez
    python duplike_temizle.py --uygula   # yedek alir ve uygular
"""
import io
import os
import re
import shutil
import sqlite3
import sys
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

KOK = os.path.dirname(os.path.abspath(__file__))
URETIM_DB = os.path.join(KOK, 'uretim.db')

# Kullanicinin 2026-07-30 karari: bu kodlar referans listesinde KALIR ama
# uretim teyidi verilmez. Ikisi de ERP article master'da tanimli degil, yani
# zaten teyit edilemezlerdi — isaret bunu teyit ekranindan da gizler.
GEREK_YOK = ['93.TK.791B', '10.300.0541G']


def norm(s):
    """Yazim farklarindan arindirilmis karsilastirma anahtari."""
    return re.sub(r'[\s\.\-_/]', '', (s or '').upper())


def temizlik_cezasi(s):
    """Yazim bozuklugu puani; DUSUK olan tercih edilir."""
    p = 0
    if '..' in s:
        p += 8                                    # 93.00..2799H
    if '  ' in s:
        p += 6                                    # cift bosluk
    if s != s.strip() or s.rstrip().endswith(('.', '-', '_')):
        p += 5                                    # 93.TK.469.
    if ' ' in s.strip():
        p += 3                                    # 10 300 4416
    if any(ch.islower() for ch in s):
        p += 2                                    # 94.ltk.858
    return p


def erp_yazimlari(yazimlar):
    """ERP article master'da TAM olarak var olan yazimlarin kumesi.
    Baglanti kurulamazsa None doner -> 2. kural tamamen atlanir (guvenli yon:
    bilinmeyeni karar gerekcesi yapma)."""
    try:
        sys.path.insert(0, os.path.join(KOK, 'as400'))
        import launch_cek
    except Exception as e:
        print(f'  ! AS400 modulu yuklenemedi ({e}) — ERP kurali atlanacak')
        return None
    # ERP alani EBCDIC; Turkce karakterli kod sorguya konamaz (DataError).
    sorgulanabilir = []
    for y in yazimlar:
        try:
            y.encode('ascii')
            sorgulanabilir.append(y)
        except UnicodeEncodeError:
            pass
    try:
        cn = launch_cek.baglan()
    except Exception as e:
        print(f'  ! ERP baglantisi kurulamadi ({str(e)[:90]}) — ERP kurali atlanacak')
        return None
    bulunan = set()
    try:
        cu = cn.cursor()
        for i in range(0, len(sorgulanabilir), 80):
            blok = sorgulanabilir[i:i + 80]
            ph = ','.join('?' * len(blok))
            cu.execute(f'SELECT A0ARTI FROM tkc0301F.BARTF0 WHERE A0ARTI IN ({ph})', blok)
            for (a,) in cu.fetchall():
                bulunan.add((a or '').strip())
    except Exception as e:
        print(f'  ! ERP sorgusu basarisiz ({str(e)[:90]}) — ERP kurali atlanacak')
        return None
    finally:
        try:
            cn.close()
        except Exception:
            pass
    return bulunan


def gruplari_cikar(conn):
    """(kanonik, bolum, lokasyon) ayni olan >1 satirlik gruplar."""
    gruplar = {}
    for r in conn.execute(
            'SELECT id, referans_kodu, aciklama, hedef_cycle_time_sn, bolum, '
            'kaynak_suresi_sn, soktak_suresi_sn, sure_teyit, sure_teyit_tarihi, '
            'lokasyon, bukum_operasyon FROM referans_listesi').fetchall():
        gruplar.setdefault((norm(r['referans_kodu']), r['bolum'], r['lokasyon']), []).append(r)
    return {k: v for k, v in gruplar.items() if len(v) > 1}


def uretim_sayisi(conn, kod, bolum, lokasyon):
    return conn.execute(
        'SELECT COUNT(*) FROM uretim_kayitlari u JOIN vardiyalar v ON v.id = u.vardiya_id '
        "WHERE UPPER(TRIM(u.referans_kodu)) = UPPER(TRIM(?)) "
        "AND COALESCE(v.bolum,'kaynak') = COALESCE(?,'kaynak') "
        "AND COALESCE(v.lokasyon,'TK2') = COALESCE(?,'TK2')",
        (kod, bolum, lokasyon)).fetchone()[0]


def tutulacagi_sec(conn, satirlar, bolum, lokasyon, erp):
    """(tutulan_satir, kural_adi) doner."""
    ctli = [r for r in satirlar if r['hedef_cycle_time_sn']]
    if len(ctli) == 1:
        return ctli[0], 'cycle time tanimli'
    if erp is not None:
        e = [r for r in satirlar if r['referans_kodu'].strip() in erp]
        if len(e) == 1:
            return e[0], 'ERPde tanimli'
    puanli = sorted(satirlar, key=lambda r: temizlik_cezasi(r['referans_kodu']))
    en_dusuk = temizlik_cezasi(puanli[0]['referans_kodu'])
    esitler = [r for r in puanli if temizlik_cezasi(r['referans_kodu']) == en_dusuk]
    if len(esitler) == 1:
        return esitler[0], 'yazim temizligi'
    esitler.sort(key=lambda r: (-uretim_sayisi(conn, r['referans_kodu'], bolum, lokasyon), r['id']))
    return esitler[0], 'kullanim sikligi'


DEVRALINACAK = ('aciklama', 'hedef_cycle_time_sn', 'kaynak_suresi_sn',
                'soktak_suresi_sn', 'sure_teyit', 'sure_teyit_tarihi', 'bukum_operasyon')


def main():
    uygula = '--uygula' in sys.argv
    if not os.path.exists(URETIM_DB):
        print(f'HATA: {URETIM_DB} bulunamadi'); return 1

    conn = sqlite3.connect(URETIM_DB)
    conn.row_factory = sqlite3.Row
    gruplar = gruplari_cikar(conn)
    if not gruplar:
        print('Duplike referans yok — yapilacak bir sey yok.'); return 0

    yazimlar = sorted({r['referans_kodu'].strip() for v in gruplar.values() for r in v})
    print(f'{len(gruplar)} duplike grup, {len(yazimlar)} farkli yazim bulundu.')
    print('ERP article master sorgulaniyor...')
    erp = erp_yazimlari(yazimlar)
    if erp is None:
        # ERP kurali 67 grubun 41'ini tek basina cozuyor. Onsuz karar yazim
        # temizligine duser ve en az bir grupta ERPde TANIMSIZ yazimi tutar
        # (93.001483). Yanlis yazimi kalicilastirmaktansa durmak dogru.
        print('\n>>> DURDURULDU: ERP article masterina ulasilamadi.')
        print('>>> Karar kalitesi buna bagli — baglantiyi duzeltip tekrar calistirin.')
        print('>>> Yine de ERPsiz devam etmek icin: --erpsiz (onerilmez)')
        if '--erpsiz' not in sys.argv:
            return 1
        print('>>> --erpsiz verildi, yazim temizligi ile devam ediliyor.')
    else:
        print(f'  ERPde tanimli yazim: {len(erp)}/{len(yazimlar)}')

    plan, kural_sayaci, tasima_toplam = [], {}, 0
    for (kan, bolum, lokasyon), satirlar in sorted(gruplar.items()):
        tut, kural = tutulacagi_sec(conn, satirlar, bolum, lokasyon, erp)
        silinecek = [r for r in satirlar if r['id'] != tut['id']]
        devral = {}
        for alan in DEVRALINACAK:
            if tut[alan] in (None, ''):
                for s in silinecek:
                    if s[alan] not in (None, ''):
                        devral[alan] = s[alan]; break
        tasima = [(s, uretim_sayisi(conn, s['referans_kodu'], bolum, lokasyon)) for s in silinecek]
        tasima_toplam += sum(n for _, n in tasima)
        kural_sayaci[kural] = kural_sayaci.get(kural, 0) + 1
        plan.append((bolum, lokasyon, tut, tasima, devral, kural))

    print('\n' + '=' * 78)
    for bolum, lokasyon, tut, tasima, devral, kural in plan:
        print(f'[{kural:<18}] {bolum}/{lokasyon}')
        ct = tut['hedef_cycle_time_sn'] or devral.get('hedef_cycle_time_sn')
        print(f'   TUT  {tut["referans_kodu"]:<26} ct={ct or "-"}'
              + (f'  (devralinan: {", ".join(devral)})' if devral else ''))
        for s, n in tasima:
            print(f'   SIL  {s["referans_kodu"]:<26}'
                  + (f'  -> {n} uretim kaydi tasinacak' if n else '  (uretim kaydi yok)'))
    print('=' * 78)
    print('kural dagilimi: ' + ' | '.join(f'{k}: {v}' for k, v in sorted(kural_sayaci.items())))
    print(f'silinecek referans satiri: {sum(len(t) for _, _, _, t, _, _ in plan)}')
    print(f'tasinacak uretim kaydi   : {tasima_toplam}')

    print(f'\nKALICI "teyit gerekmez" isareti eklenecek: {", ".join(GEREK_YOK)}')

    if not uygula:
        print('\n>>> BU BIR ONIZLEME. Hicbir sey degismedi.')
        print('>>> Uygulamak icin: python duplike_temizle.py --uygula')
        return 0

    yedek = URETIM_DB + '.duplike_yedek_' + datetime.now().strftime('%Y%m%d_%H%M%S')
    shutil.copy2(URETIM_DB, yedek)
    print(f'\nYedek alindi: {os.path.basename(yedek)}')

    tasinan = silinen = 0
    try:
        with conn:
            for bolum, lokasyon, tut, tasima, devral, _ in plan:
                if devral:
                    conn.execute(
                        'UPDATE referans_listesi SET '
                        + ', '.join(f'{a}=?' for a in devral) + ' WHERE id=?',
                        list(devral.values()) + [tut['id']])
                for s, _n in tasima:
                    # Sadece bu bolum+lokasyondaki kayitlar; ayni yazim baska
                    # bolumde farkli bir gruba ait olabilir.
                    cur = conn.execute(
                        'UPDATE uretim_kayitlari SET referans_kodu=? '
                        'WHERE UPPER(TRIM(referans_kodu))=UPPER(TRIM(?)) AND vardiya_id IN '
                        "(SELECT id FROM vardiyalar WHERE COALESCE(bolum,'kaynak')=COALESCE(?,'kaynak') "
                        "AND COALESCE(lokasyon,'TK2')=COALESCE(?,'TK2'))",
                        (tut['referans_kodu'], s['referans_kodu'], bolum, lokasyon))
                    tasinan += cur.rowcount
                    conn.execute('DELETE FROM referans_listesi WHERE id=?', (s['id'],))
                    silinen += 1
            for kod in GEREK_YOK:
                var = conn.execute(
                    "SELECT id FROM as400_teyit_isaret WHERE kapsam='kalici' AND durum='gerek_yok' "
                    "AND UPPER(REPLACE(referans,' ',''))=UPPER(REPLACE(?,' ',''))", (kod,)).fetchone()
                if var:
                    print(f'  isaret zaten var: {kod} (#{var["id"]})')
                    continue
                conn.execute(
                    'INSERT INTO as400_teyit_isaret (kapsam, uretim_tarihi, bolum, referans, '
                    'durum, aciklama, olusturan) VALUES (?,?,?,?,?,?,?)',
                    ('kalici', None, None, kod, 'gerek_yok',
                     'Kullanici karari 2026-07-30: referans listesinde kalir, uretim teyidi verilmez '
                     '(ERP article masterda tanimli degil).', 'duplike_temizle.py'))
                print(f'  isaret eklendi : {kod}')
    except Exception as e:
        print(f'\nHATA — degisiklikler geri alindi: {e}')
        print(f'Yedek: {yedek}')
        return 1

    kalan = len(gruplari_cikar(conn))
    conn.close()
    print(f'\nTAMAM: {silinen} duplike referans silindi, {tasinan} uretim kaydi tasindi.')
    print(f'Kalan duplike grup: {kalan}')
    print('\n>>> Panelin guncel veriyi gormesi icin servisi yeniden baslatin.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
