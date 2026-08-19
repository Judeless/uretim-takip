# -*- coding: utf-8 -*-
"""
eoq_kesif.py — EOQ'nun AS400 article master'inda HANGI KOLONDA durdugunu bulur.

NEDEN: EOQ'yu ekran kazimayla (07/10/01, urun basina ~3 sn) cekmek 323 referansta
15-25 dakika suruyor ve AS400 oturumunu mesgul ediyor. Sistemde ZATEN calisan bir
ODBC baglantisi var (as400_config + pyodbc, launch_esle/teyit sorgulari) — dogru
kolon bilinirse tum EOQ'lar TEK SORGUDA saniyeler icinde gelir.

Kolon adini bilmiyoruz (ekranda yalniz "EOQ . . . 850" yaziyor). Bu betik iki
yoldan arar:
  1) SYSCOLUMNS'ta adi/aciklamasi EOQ'ya benzeyen kolonlar
  2) Bilinen bir urunun satirini dokup DEGERI beklenen EOQ'ya esit kolonlari isaretler
     (kullanici 2026-08-19: 10.300.3059W -> EOQ 850)

SALT OKUNUR: yalnizca SELECT. Hicbir veri degistirmez.

KULLANIM (sunucuda — iSeries ODBC surucusu + keyring sifresi orada):
  1) Kolonu bul:
       python as400\\eoq_kesif.py 10.300.3059W 850
  2) Bulunan kolonla TUM kaynak referanslarinin EOQ'sunu cek ve panele yaz:
       python as400\\eoq_kesif.py --cek A0XXXX http://192.168.20.210:5000
     (once yazmadan gormek icin sona --deneme ekleyin)
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

KUTUPHANE = 'TKC0301F'
TABLO = 'BARTF0'        # article master (launch_esle.article_tanimli: BARTF0.A0ARTI)
KOD_KOLON = 'A0ARTI'


def _baglan():
    import pyodbc
    import as400_config as cfg
    return pyodbc.connect(cfg.baglanti_dizesi(cfg.sifre_al()), timeout=20, autocommit=True)


def eoq_benzeri_kolonlar(cn):
    """Adi ya da aciklamasi EOQ/lotto/riordino cagristiran kolonlar."""
    sql = f"""
        SELECT COLUMN_NAME, DATA_TYPE, LENGTH, COALESCE(COLUMN_TEXT,'')
        FROM QSYS2.SYSCOLUMNS
        WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ?
        ORDER BY ORDINAL_POSITION
    """
    kolonlar = cn.cursor().execute(sql, (KUTUPHANE, TABLO)).fetchall()
    anahtar = ('EOQ', 'LOTT', 'LOTM', 'RIORD', 'MINQ', 'QORD', 'ORDQ', 'ECON')
    bulunan = []
    for k in kolonlar:
        ad = (k[0] or '').strip().upper()
        aciklama = (k[3] or '').strip().upper()
        if any(a in ad or a in aciklama for a in anahtar):
            bulunan.append((ad, k[1], k[2], (k[3] or '').strip()))
    return kolonlar, bulunan


def satir_dok(cn, article, beklenen=None):
    """Bir urunun TUM kolonlarini doker; beklenen degere esit olanlari isaretler."""
    cur = cn.cursor()
    cur.execute(
        f"SELECT * FROM {KUTUPHANE}.{TABLO} WHERE TRIM({KOD_KOLON}) = ?",
        (article.strip(),))
    satir = cur.fetchone()
    if satir is None:
        print(f'  [!] {article} article master\'da BULUNAMADI '
              f'(kod bosluk dolgulu olabilir; TRIM ile arandi)')
        return []
    adlar = [d[0] for d in cur.description]
    eslesenler = []
    print(f'  --- {article} satiri ({len(adlar)} kolon) ---')
    for ad, deger in zip(adlar, satir):
        try:
            metin = str(deger).strip()
        except Exception:
            metin = repr(deger)
        if not metin or metin in ('0', '0.0', '0.00', 'None'):
            continue      # bos/sifir kolonlari gosterme (ekran cok uzuyor)
        isaret = ''
        if beklenen is not None:
            try:
                if abs(float(deger) - float(beklenen)) < 0.001:
                    isaret = '   <<<<<< BEKLENEN EOQ'
                    eslesenler.append(ad)
            except (TypeError, ValueError):
                pass
        print(f'    {ad:<12} = {metin}{isaret}')
    return eslesenler


def toplu_cek(cn, kolon, sunucu, deneme=False):
    """Sunucudaki kaynak referanslarinin EOQ'sunu TEK SORGUDA cekip panele yazar."""
    import json
    import urllib.request

    with urllib.request.urlopen(
            sunucu.rstrip('/') + '/api/kaynak_eoq?bolum=kaynak&lokasyon=TK2',
            timeout=30) as r:
        veri = json.loads(r.read().decode('utf-8'))
    kodlar = [x['referans_kodu'] for x in veri.get('referanslar', [])]
    print(f'  {len(kodlar)} kaynak referansi icin EOQ sorgulanacak...')
    if not kodlar:
        return 0

    # Bloklu IN sorgusu (DB2 parametre siniri) — ekran kazimaya gore saniyeler surer
    cur = cn.cursor()
    kayitlar, bulunamayan = [], []
    for i in range(0, len(kodlar), 400):
        blok = kodlar[i:i + 400]
        yer = ','.join('?' * len(blok))
        sql = (f"SELECT TRIM({KOD_KOLON}), {kolon} FROM {KUTUPHANE}.{TABLO} "
               f"WHERE TRIM({KOD_KOLON}) IN ({yer})")
        bulunan = {}
        for satir in cur.execute(sql, blok).fetchall():
            try:
                bulunan[str(satir[0]).strip().upper()] = int(float(satir[1] or 0))
            except (TypeError, ValueError):
                pass
        for kod in blok:
            deger = bulunan.get(kod.strip().upper())
            if deger is None:
                bulunamayan.append(kod)
            else:
                kayitlar.append({'referans_kodu': kod, 'eoq': deger})

    dolu = sum(1 for k in kayitlar if k['eoq'] > 0)
    print(f"  okunan={len(kayitlar)} (EOQ>0: {dolu})  ERP'de bulunamayan={len(bulunamayan)}")
    if bulunamayan[:10]:
        print('  bulunamayan ornek: ' + ', '.join(bulunamayan[:10]))
    if deneme:
        print('  [deneme] Panele YAZILMADI. Ilk 15 satir:')
        for k in kayitlar[:15]:
            print(f"    {k['referans_kodu']:<20} {k['eoq']}")
        return 0

    govde = json.dumps({'kaynak': 'as400', 'lokasyon': 'TK2',
                        'kayitlar': kayitlar}).encode('utf-8')
    istek = urllib.request.Request(
        sunucu.rstrip('/') + '/api/kaynak_eoq', data=govde,
        headers={'Content-Type': 'application/json'}, method='POST')
    with urllib.request.urlopen(istek, timeout=60) as r:
        print('  panel yaniti: ' + r.read().decode('utf-8')[:300])
    return len(kayitlar)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1

    # --cek <KOLON> <SUNUCU> [--deneme]  → kolon bilindikten SONRA toplu cekim
    if sys.argv[1] == '--cek':
        if len(sys.argv) < 4:
            print('Kullanim: eoq_kesif.py --cek <KOLON_ADI> <SUNUCU_URL> [--deneme]')
            return 1
        kolon, sunucu = sys.argv[2].strip().upper(), sys.argv[3]
        deneme = '--deneme' in sys.argv
        try:
            cn = _baglan()
        except Exception as e:
            print(f'[HATA] AS400 baglantisi kurulamadi: {e}')
            return 2
        try:
            print(f'=== EOQ cekiliyor: {KUTUPHANE}.{TABLO}.{kolon} -> {sunucu} ===')
            toplu_cek(cn, kolon, sunucu, deneme)
        except Exception as e:
            print(f'[HATA] {e}')
            return 3
        finally:
            cn.close()
        return 0

    article = sys.argv[1]
    beklenen = sys.argv[2] if len(sys.argv) > 2 else None

    try:
        cn = _baglan()
    except Exception as e:
        print(f'[HATA] AS400 baglantisi kurulamadi: {e}')
        print('       Bu betik SUNUCUDA calismali (iSeries ODBC surucusu + keyring sifresi).')
        return 2

    try:
        print(f'\n=== 1) {KUTUPHANE}.{TABLO} icinde EOQ benzeri kolon adlari ===')
        try:
            tum, benzer = eoq_benzeri_kolonlar(cn)
            print(f'  (tabloda toplam {len(tum)} kolon)')
            if benzer:
                for ad, tip, uzunluk, aciklama in benzer:
                    print(f'    {ad:<12} {tip}({uzunluk})  {aciklama}')
            else:
                print('    (ada gore aday bulunamadi — asagidaki deger eslesmesine bakin)')
        except Exception as e:
            print(f'  [!] SYSCOLUMNS okunamadi: {e}')

        print(f'\n=== 2) {article} satirindaki dolu kolonlar ===')
        if beklenen:
            print(f'  (beklenen EOQ = {beklenen} — esit kolonlar isaretlenecek)')
        eslesen = satir_dok(cn, article, beklenen)

        print('\n=== SONUC ===')
        if eslesen:
            print('  EOQ su kolon(lar)da olabilir: ' + ', '.join(eslesen))
            print('  Birden fazlaysa ikinci bir urunle tekrar calistirin — yanlis')
            print('  aday elenir (ornegin stok/siparis miktari tesadufen esit olmus olabilir).')
            print('')
            print("  SONRAKI ADIM — tum EOQ'lari cekip panele yazmak icin:")
            print(f'    python as400\\eoq_kesif.py --cek {eslesen[0]} '
                  f'http://192.168.20.210:5000 --deneme')
            print("    (cikti dogruysa --deneme'yi kaldirip tekrar calistirin)")
        else:
            print('  Deger eslesmesi bulunamadi. Ekrandaki EOQ baska bir dosyada olabilir')
            print('  (article master degil, depo/planlama dosyasi). Ciktiyi paylasin.')
    finally:
        cn.close()
    return 0


if __name__ == '__main__':
    sys.exit(main())
