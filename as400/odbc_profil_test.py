# -*- coding: utf-8 -*-
"""AS400 ODBC profili yetki testi (2026-09-10). Kullanim:
    python as400/odbc_profil_test.py COFLEFORGE      (varsayilan: as400_config.DB_KULLANICI)
Verilen profilin sifresini kasadan alir, sistemin kullandigi HER tabloya SELECT
dener ve tablo tablo OK/HATA yazar. VERI YAZMAZ, veri GOSTERMEZ (yalniz satir sayisi).
Sunucuda promanage (RDP) oturumunda calistirin — kasa oradadir.
Hepsi OK ise as400/odbc_config.json'a {"kullanici": "<profil>"} yazip cofle-app ve
teyit-agent'i yeniden baslatin."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, __file__.rsplit('\\', 1)[0])
import as400_config as CFG

ku = (sys.argv[1] if len(sys.argv) > 1 else CFG.DB_KULLANICI).strip().upper()
if ku not in CFG.KULLANICILAR:
    print(f'Bilinmeyen profil: {ku}. Gecerli: {", ".join(CFG.KULLANICILAR)}'); sys.exit(1)
pw = CFG.sifre_al(ku)
if not pw:
    print(f'{ku} sifresi kasada YOK — once: python as400/kaydet_sifre.py {ku}'); sys.exit(1)
try:
    import pyodbc
    cn = pyodbc.connect(CFG.baglanti_dizesi(pw, ku), timeout=20, autocommit=True)
except Exception as e:
    print(f'BAGLANTI HATASI ({ku}): {e}'); sys.exit(1)
print(f'Baglanti OK — profil {ku}, host {CFG.HOST}')

# (aciklama, sql) — sistemin gercekten kullandigi okumalar
SORGULAR = [
    ('XPRO90 acik emirler (launch listesi)',      'SELECT COUNT(*) FROM tkc0301F.XPRO90'),
    ('BMMAF0 depo hareketleri (teyit kontrolu)',  'SELECT COUNT(*) FROM tkc0301F.BMMAF0 WHERE MGCACD IN (\'RPR\',\'CFI\',\'COP\')'),
    ('BPROF0 OP ana dosyasi (launch durumu)',     'SELECT COUNT(*) FROM tkc0301F.BPROF0'),
    ('BARTF0 article master (kod dogrulama)',     'SELECT COUNT(*) FROM tkc0301F.BARTF0'),
    ('QSYS2.SYSCOLUMNS (import kolonlari)',       'SELECT COUNT(*) FROM QSYS2.SYSCOLUMNS WHERE TABLE_SCHEMA=\'COFLEFORGE\' AND TABLE_NAME=\'BMMAF0I\''),
    ('COFLEFORGE.BMMAF0I staging (import)',       'SELECT COUNT(*) FROM COFLEFORGE.BMMAF0I'),
]
hata = 0
for ad, sql in SORGULAR:
    try:
        n = cn.cursor().execute(sql).fetchone()[0]
        print(f'  OK    {ad}: {n} satir')
    except Exception as e:
        hata += 1
        print(f'  HATA  {ad}: {str(e)[:160]}')
cn.close()
print()
if hata:
    print(f'{hata} sorgu basarisiz — bu profil ODBC okumalari icin YETERSIZ. IT ye (Simone) yukaridaki tablolar icin SELECT yetkisi sorun.')
    sys.exit(2)
print(f'HEPSI OK — {ku} tum okumalar icin kullanilabilir. Sonraki adim: as400/odbc_config.json -> {{"kullanici": "{ku}"}} ve restart.')
