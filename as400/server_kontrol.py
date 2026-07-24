# -*- coding: utf-8 -*-
"""
server_kontrol.py — Server AS400 teyit kurulumunun sağlık kontrolü (tek komut).

KULLANIM: PCOMM oturumlarının açık olduğu RDP oturumunda (promanage):
    cd C:\\cofle\\uretim_takip\\as400
    python server_kontrol.py

Her adımı [OK]/[HATA] ile raporlar; HATA satırı ne yapılacağını söyler.
Laptopta da çalışır (orada agent kapalı = normaldir).
"""
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

SONUC = {'ok': 0, 'hata': 0}


def basla(ad):
    print(f'\n-- {ad} --')


def ok(m):
    SONUC['ok'] += 1
    print(f'  [OK]   {m}')


def hata(m):
    SONUC['hata'] += 1
    print(f'  [HATA] {m}')


print('=' * 60)
print('  COFLE SERVER AS400 KONTROL')
print(f'  Makine: {os.environ.get("COMPUTERNAME", "?")} · Hesap: {os.environ.get("USERNAME", "?")}')
print('=' * 60)

# 1) Python paketleri
basla('Python paketleri')
for mod in ('flask', 'requests', 'pyodbc', 'keyring'):
    try:
        __import__(mod)
        ok(mod)
    except ImportError:
        hata(f'{mod} YOK -> pip install {mod}')

# 2) ODBC surucusu
basla('iSeries ODBC surucusu')
try:
    import pyodbc
    dr = [d for d in pyodbc.drivers() if 'iSeries' in d or 'IBM i' in d]
    if dr:
        ok('bulundu: ' + ', '.join(dr))
    else:
        hata('iSeries Access ODBC Driver YOK (iSeries Access kurulumunda ODBC secili miydi?). '
             'Mevcut: ' + ', '.join(pyodbc.drivers()))
except Exception as e:
    hata(f'pyodbc yuklenemedi: {e}')

# 3) AS400 sifresi — BU hesabin keyring kasasinda
basla(f'AS400 sifresi (keyring, hesap={os.environ.get("USERNAME", "?")})')
try:
    import keyring
    import as400_config as cfg
    pw = keyring.get_password(cfg.KEYRING_SERVICE, cfg.DB_KULLANICI)
    if pw:
        ok(f'{cfg.DB_KULLANICI} sifresi bu hesapta kayitli')
    else:
        hata('sifre YOK -> BU hesapta calistir: python kaydet_sifre.py')
except Exception as e:
    hata(f'{e}')

# 4) AS400 canli baglanti (ODBC + sifre uctan uca)
basla('AS400 canli sorgu (pyodbc -> BPROF0)')
try:
    import keyring
    import pyodbc
    import as400_config as cfg
    pw = keyring.get_password(cfg.KEYRING_SERVICE, cfg.DB_KULLANICI)
    if not pw:
        hata('sifre yok — 3. adimi tamamla')
    else:
        cn = pyodbc.connect(cfg.baglanti_dizesi(pw), timeout=10)
        r = cn.cursor().execute('SELECT COUNT(*) FROM tkc0301F.BPROF0').fetchone()
        cn.close()
        ok(f'BPROF0 okundu ({r[0]} satir) — liste/dogrulama sorgulari calisir')
except Exception as e:
    hata(f'{e}')

# 5) Teyit-agent /sifre — servis (LocalSystem) sifreyi buradan alir (Plan B)
# Artik NSSM servis hesabini degistirmeye GEREK YOK: cofle-app LocalSystem'de kalir,
# AS400 sifresini agent'tan (promanage oturumu, keyring gorunur) localhost ile alir.
basla('Teyit-agent /sifre (servis sifre koprusu)')
try:
    import requests
    r = requests.get('http://127.0.0.1:5010/sifre', timeout=2)
    if r.status_code == 200 and (r.json() or {}).get('sifre'):
        ok('agent sifreyi veriyor — LocalSystem servisi bunu kullanabilir')
    else:
        hata('agent /sifre BOS dondu — bu (promanage) oturumda AS400 sifresi kayitli mi? '
             '(kaydet_sifre.py bu oturumda calistirilmali)')
except Exception:
    hata('agent /sifre yanit vermedi — agent KAPALI (Teyit_Agent_Baslat.bat calistir)')

# 6) Teyit-agent (PCOMM oturumu koprusu)
basla('Teyit-agent (127.0.0.1:5010)')
try:
    import requests
    r = requests.get('http://127.0.0.1:5010/durum', timeout=2)
    d = r.json() or {}
    if d.get('agent') == 'cofle-teyit':
        p = d.get('pcomm')
        if isinstance(p, int) and p >= 2:
            ok(f'agent ayakta · PCOMM pencere: {p} (A+B gorunuyor)')
        else:
            hata(f'agent ayakta ama PCOMM pencere: {p} — A+B acik mi? '
                 'Agent, PCOMM ile AYNI RDP oturumunda mi calisiyor?')
    else:
        hata('5010 yanit verdi ama teyit-agent degil (port cakismasi?)')
except Exception:
    hata('agent KAPALI -> PCOMM acik olan RDP oturumunda Teyit_Agent_Baslat.bat calistir '
         '(+ shell:startup kisayolu)')

# 7) Ana app
basla('cofle-app web (127.0.0.1:5000)')
try:
    import requests
    r = requests.get('http://127.0.0.1:5000/', timeout=4)
    ok(f'HTTP {r.status_code}')
except Exception as e:
    hata(f'{e} — NSSM cofle-app calisiyor mu?')

print('\n' + '=' * 60)
print(f'  SONUC: {SONUC["ok"]} OK · {SONUC["hata"]} HATA')
if SONUC['hata'] == 0:
    print('  Kurulum TAM. Son test: dashboard -> AS400 Teyit -> Yenile,')
    print('  sonra TEK launch sec -> Gonder (Session B yi RDP den izle).')
else:
    print('  Yukaridaki [HATA] satirlarini sirayla duzeltip tekrar calistir.')
print('=' * 60)
