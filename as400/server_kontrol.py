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

# 5) NSSM servis hesabi — keyring tuzagi
basla('cofle-app servis hesabi (keyring gorunurlugu)')
try:
    out = subprocess.run(['sc', 'qc', 'cofle-app'], capture_output=True, timeout=10
                         ).stdout.decode(errors='replace')
    hesap = ''
    for satir in out.splitlines():
        if 'SERVICE_START_NAME' in satir:
            hesap = satir.split(':', 1)[-1].strip()
    if not hesap:
        hata('cofle-app servisi bulunamadi (bu makine server degil mi?)')
    elif hesap.lower() in ('localsystem', 'nt authority\\system'):
        hata(f'servis {hesap} olarak kosuyor -> promanage kasasindaki sifreyi GOREMEZ. '
             'Duzelt: C:\\cofle\\nssm.exe set cofle-app ObjectName .\\promanage <sifre> '
             've C:\\cofle\\nssm.exe restart cofle-app')
    elif 'promanage' in hesap.lower() or hesap.lower().startswith('.\\'):
        ok(f'servis hesabi: {hesap}')
    else:
        hata(f'servis hesabi: {hesap} — sifreyi kaydeden hesapla AYNI olmali')
except Exception as e:
    hata(f'sc qc calismadi: {e}')

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
