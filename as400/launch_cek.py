# -*- coding: utf-8 -*-
"""AS400'den uretim emri (launch) verisini pyodbc ile ceker.
Sifre Windows Kimlik Bilgileri Yoneticisi'nden (keyring) runtime'da okunur —
hicbir yerde duz metin durmaz. Excel/F6 adimi YOK.

Kullanim (test):  python as400/launch_cek.py
"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pyodbc
import as400_config as CFG


def _sifre():
    """Sifre: once kasa, olmazsa TEYIT-AGENT (as400_config.sifre_al).

    2026-09-16 saha: Kaynak Plani 'AS400 baglantisi kurulamadi: Sifre kasada yok'
    veriyordu — burada keyring DOGRUDAN okunuyordu; sunucuda cofle-app NSSM
    servisi LocalSystem'da kosar ve promanage kasasini GOREMEZ. sifre_al agent'a
    (127.0.0.1:5010) duser, o da kasayi gorur.
    """
    p = CFG.sifre_al()
    if not p:
        raise RuntimeError(
            f"AS400 sifresi alinamadi ({CFG.DB_KULLANICI}) — kasada yok ya da teyit-agent kapali. "
            f"Sunucuda: python as400/kaydet_sifre.py {CFG.DB_KULLANICI} ve teyit-agent'i baslatin.")
    return p


def baglan(timeout=15):
    """AS400 baglantisi — tek merkez (kilit + sifre kasa/agent): as400_config.baglan."""
    return CFG.baglan(timeout=timeout)


def satir_cek(limit=None, sorgu=None):
    """Uretim emri satirlarini dondurur. limit verilirse ilk N satir."""
    q = sorgu or CFG.SORGU
    if limit:
        q += f" FETCH FIRST {int(limit)} ROWS ONLY"
    with baglan() as conn:
        cur = conn.cursor()
        cur.execute(q)
        kolonlar = [d[0] for d in cur.description]
        satirlar = [dict(zip(kolonlar, r)) for r in cur.fetchall()]
    return satirlar


if __name__ == '__main__':
    print('AS400 baglanti testi ...', flush=True)
    try:
        with baglan() as conn:
            print('  BAGLANDI:', conn.getinfo(pyodbc.SQL_DBMS_NAME),
                  conn.getinfo(pyodbc.SQL_DBMS_VER))
    except Exception as e:
        print('  BAGLANTI HATASI:', e)
        sys.exit(1)

    # Ornek satirlar
    try:
        ornek = satir_cek(limit=8)
    except Exception as e:
        print('  SORGU HATASI:', e)
        sys.exit(1)

    print(f'\n  Ornek {len(ornek)} satir (launch = SS-YY-Nr):')
    for r in ornek:
        launch = f"{r['XWRED1']}-{r['XWRED2']}-{r['XWRENU']}"
        art = (r['XWARTI'] or '').strip()
        print(f"    launch {launch:>16} | article {art:<22} | "
              f"qty {r['XWQTOR']} kalan {r['XWQTRE']} durum {r['XWAVAN']}")

    # Toplam satir + acik (durum<70) sayisi
    try:
        with baglan() as conn:
            cur = conn.cursor()
            top = cur.execute(f"SELECT COUNT(*) FROM {CFG.KAYNAK_TABLO}").fetchval()
            acik = cur.execute(
                f"SELECT COUNT(*) FROM {CFG.KAYNAK_TABLO} WHERE XWAVAN < 70"
            ).fetchval()
        print(f'\n  Toplam satir: {top} | durumu <70 (acik olabilir): {acik}')
    except Exception as e:
        print('  Sayim hatasi (onemsiz):', e)
    print('\nTAMAM.')
