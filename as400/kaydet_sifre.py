# -*- coding: utf-8 -*-
"""AS400 sifresini Windows Kimlik Bilgileri Yoneticisi'ne (keyring/DPAPI) kaydeder.
TEK SEFER calistirilir (sifre degisince tekrar). Sifre gizli girilir (ekranda
gorunmez), hicbir dosyaya yazilmaz, buluta senkron olmaz.

Kullanim:  python as400/kaydet_sifre.py              (EMREDTK — okuma + ekran robotu)
           python as400/kaydet_sifre.py COFLEFORGE   (import profili, Italya IT 2026-09-09)
Sunucuda promanage (RDP) oturumunda calistirilir: cofle-app servisi sifreyi
teyit-agent uzerinden o oturumun kasasindan alir.
"""
import sys, io, getpass
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

try:
    import keyring
except ImportError:
    print('keyring kurulu degil. Once:  python -m pip install keyring')
    sys.exit(1)

sys.path.insert(0, __file__.rsplit('\\', 1)[0])
from as400_config import KEYRING_SERVICE, DB_KULLANICI, KULLANICILAR

kullanici = (sys.argv[1] if len(sys.argv) > 1 else DB_KULLANICI).strip().upper()
if kullanici not in KULLANICILAR:
    print(f'Bilinmeyen kullanici: {kullanici}. Gecerli: {", ".join(KULLANICILAR)}')
    sys.exit(1)

print(f'AS400 sifresi kaydediliyor — kullanici: {kullanici}')
print('(Sifre yaziyorken EKRANDA GORUNMEZ, bu normaldir.)')
p1 = getpass.getpass('Sifre        : ')
p2 = getpass.getpass('Sifre (tekrar): ')

if not p1:
    print('Bos sifre — vazgecildi.'); sys.exit(1)
if p1 != p2:
    print('Sifreler uyusmadi — tekrar calistir.'); sys.exit(1)

keyring.set_password(KEYRING_SERVICE, kullanici, p1)

# Geri okuyup DOGRULA (sadece uzunluk gosterilir, sifrenin kendisi DEGIL)
geri = keyring.get_password(KEYRING_SERVICE, kullanici)
if geri == p1:
    print(f'OK — {kullanici} sifresi kasaya kaydedildi ({len(geri)} karakter). Bu pencere kapatilabilir.')
else:
    print('HATA — kayit dogrulanamadi.')
    sys.exit(1)
