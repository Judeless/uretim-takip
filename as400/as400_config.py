# -*- coding: utf-8 -*-
"""AS400 (S.I. COFLETK) baglanti ayarlari — ORTAK sabitler.
ONEMLI: Bu dosyada ASLA sifre YOKTUR. Sifre Windows Kimlik Bilgileri
Yoneticisi'nde (keyring / DPAPI sifreli) durur; runtime'da okunur.
Sifreyi kaydetmek icin: python as400/kaydet_sifre.py
"""

# Windows Kimlik Bilgileri Yoneticisi anahtarlari
KEYRING_SERVICE = 'cofle_as400'
DB_KULLANICI    = 'EMREDTK'          # AS400 kullanici adi (sifre DEGIL)

# Baglanti
HOST   = '192.168.1.1'
DRIVER = 'iSeries Access ODBC Driver'   # 64-bit (Client Access 32-bit DEGIL)

# CANLI uretim emri kaynagi (2026-07-17 katalog kesfiyle bulundu).
# BPROF0 = OP ana dosyasi (~160k; durum 70=kapali cogunluk). XPRO90 = yalniz
# ACIK emirler gorunumu (durum 10/40/45/50, ~5k satir) — hizli ve daima guncel.
# Eski kaynak XB530W0 work-file'i (10>05>03>29+F6 raporu) yalniz GECIKMIS
# launch'lari icerdigi icin TERK EDILDI (F6/Excel adimi artik gereksiz).
KAYNAK_TABLO = 'tkc0301F.XPRO90'

# Kolonlar (BPROF0/XPRO90 ortak):
#   Q0RED1/Q0RED2/Q0RENU = OP no bilesenleri (SS-YY-Nr → launch numarasi)
#   Q0ARTI = article (referans), Q0AVAN = durum:
#       10 = OPR var, launch ALINMAMIS; 40 = launch acik (teyit verilebilir);
#       45/50 = TK1 akisi; 70 = kapali (yalniz BPROF0'da)
#   Q0QTOR = launch adedi, Q0QTRI = simdiye kadar teyit verilen (rientrata)
SORGU = (
    "SELECT Q0RED1, Q0RED2, Q0RENU, Q0ARTI, Q0AVAN, Q0QTOR, Q0QTRI "
    "FROM " + KAYNAK_TABLO
)

# ── TEYIT (rientro) HAREKETLERI ──
# Depo hareket ana dosyasi. MGCACD='RPR' = rientro produzione = uretim teyidi.
# Ekranda gorunen "Re-entry summary" satirlarinin kaynagi (dogrulandi 2026-07-17).
# ONEMLI: hareket tarihi = teyidin GIRILDIGI tarih (dunun uretimi bugun girilir),
# uretim tarihi DEGIL. Mukerrer kontrolu bunu hesaba katmalidir.
HAREKET_TABLO = 'tkc0301F.BMMAF0'
# ARTICLE bazli sorgu (launch bazli DEGIL): operator ayni urunu BASKA bir launch'a
# girip S ile kapatmis olabilir (o launch durum 70 → acik listede gorunmez).
# Mukerrer kontrolu bu yuzden article uzerinden yapilmali (2026-07-17 dersi:
# 1847A → operator 140636'ya girdi+kapatti; bizim listede zombi 17809 vardi).
# RPR = launch teyidi; CFI = launch'siz uretimin depo girisi (07>01>F1, kullanici
# 2026-07-20 ogretti) — ikisi de "bu uretim ERP'ye islendi" kaniti sayilir.
HAREKET_SORGU = (
    "SELECT MGARCD, MGDSSO, MGDAAO, MGDMMO, MGDGGO, MGQTA, MGRED2, MGRENU, MGCACD "
    "FROM " + HAREKET_TABLO + " WHERE MGCACD IN ('RPR','CFI') AND MGARCD IN ({yer}) "
)


def baglanti_dizesi(sifre):
    """pyodbc connection string uretir. sifre yalnizca runtime'da,
    bellekte kullanilir — hicbir yere yazilmaz."""
    return (
        f"DRIVER={{{DRIVER}}};"
        f"SYSTEM={HOST};"
        f"UID={DB_KULLANICI};"
        f"PWD={sifre};"
        # Metin cevirisi: AS400 EBCDIC → Unicode (Turkce alanlar dogru gelsin)
        "CCSID=1208;TRANSLATE=1;"
    )
