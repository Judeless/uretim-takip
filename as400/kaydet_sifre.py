# -*- coding: utf-8 -*-
"""AS400 sifresini Windows Kimlik Bilgileri Yoneticisi'ne (keyring/DPAPI) kaydeder.
TEK SEFER calistirilir (sifre degisince tekrar). Sifre gizli girilir (ekranda
gorunmez), hicbir dosyaya yazilmaz, buluta senkron olmaz.

Kullanim:  python as400/kaydet_sifre.py
"""
import sys, io, getpass
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

try:
    import keyring
except ImportError:
    print('keyring kurulu degil. Once:  python -m pip install keyring')
    sys.exit(1)

sys.path.insert(0, __file__.rsplit('\\', 1)[0])
from as400_config import KEYRING_SERVICE, DB_KULLANICI

print(f'AS400 sifresi kaydediliyor — kullanici: {DB_KULLANICI}')
print('(Sifre yaziyorken EKRANDA GORUNMEZ, bu normaldir.)')
p1 = getpass.getpass('Sifre        : ')
p2 = getpass.getpass('Sifre (tekrar): ')

if not p1:
    print('Bos sifre — vazgecildi.'); sys.exit(1)
if p1 != p2:
    print('Sifreler uyusmadi — tekrar calistir.'); sys.exit(1)

keyring.set_password(KEYRING_SERVICE, DB_KULLANICI, p1)

# Geri okuyup DOGRULA (sadece uzunluk gosterilir, sifrenin kendisi DEGIL)
geri = keyring.get_password(KEYRING_SERVICE, DB_KULLANICI)
if geri == p1:
    print(f'OK — sifre kasaya kaydedildi ({len(geri)} karakter). Bu pencere kapatilabilir.')
else:
    print('HATA — kayit dogrulanamadi.')
    sys.exit(1)
