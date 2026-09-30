# -*- coding: utf-8 -*-
"""
Sunucudan_Veri_Cek.py — CANLI sunucunun verisini bu makineye (geliştirme kopyası) indirir.

NEDEN: canlı sistem sunucuda çalışıyor; laptopta aynı kodun AYRI veritabanı ve AYRI
Excel'i var. İkisi ayrışınca "hangisi gerçek" karışıyor (2026-09-30: montaj planı
laptoptaki eski referans listesiyle ölçüldü, sunucuyla tutmadı).

KURAL: tek kaynak SUNUCU. Veri TEK YÖNDE akar: sunucu → bu makine.
Bu betik sunucuya HİÇBİR ŞEY YAZMAZ; yalnız okur ve yerel kopyayı değiştirir.

YAPTIĞI
  1. Panel yöneticisi olarak sunucuya girer (şifre ekranda sorulur, SAKLANMAZ).
  2. uretim.db'nin tutarlı kopyasını ve data\\ altındaki Excel'leri indirir.
  3. Yereldeki eskileri yedekler (uretim.db.bak-gelistirme, *.xlsx.bak-gelistirme).
  4. Excel'leri SALT OKUNUR yapar — yanlışlıkla burada düzenlenmesin.
  5. data\\GELISTIRME_KOPYASI.json işaretini yazar: panelde şerit çıkar, zamanlanmış
     işler başlamaz, AS400'e yazan uçlar kapanır.

KULLANIM
  python Sunucudan_Veri_Cek.py                       # sorarak indirir
  python Sunucudan_Veri_Cek.py --sunucu http://192.168.20.210:5000
  python Sunucudan_Veri_Cek.py --bilgi               # yalnız sunucu özetini göster
  python Sunucudan_Veri_Cek.py --isaretle            # indirmeden yalnız işareti yaz

Bu makineyi yeniden CANLI yapmak gerekirse: data\\GELISTIRME_KOPYASI.json dosyasını silin.
"""
import argparse
import getpass
import json
import os
import shutil
import socket
import sqlite3
import stat
import sys
from datetime import datetime

KOK = os.path.dirname(os.path.abspath(__file__))
VERI = os.path.join(KOK, 'data')
DB = os.path.join(KOK, 'uretim.db')
ISARET = os.path.join(VERI, 'GELISTIRME_KOPYASI.json')
VARSAYILAN_SUNUCU = 'http://192.168.20.210:5000'
EXCELLER = {'excel': 'uretim_verileri.xlsx', 'excel_tk1': 'Tk1 Veriler.xlsx'}

for _a in (sys.stdout, sys.stderr):
    try:
        _a.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass


def isaret_yaz(sunucu='', bilgi=None, indirildi=False):
    os.makedirs(VERI, exist_ok=True)
    d = {'sunucu': sunucu, 'cekildi': datetime.now().strftime('%Y-%m-%d %H:%M') if indirildi else '',
         'sunucu_host': (bilgi or {}).get('host', ''), 'referans': (bilgi or {}).get('referans'),
         'son_vardiya': (bilgi or {}).get('son_vardiya'), 'bu_makine': socket.gethostname(),
         'not': 'Bu dosya varken kurulum GELISTIRME KOPYASIDIR. Canli yapmak icin silin.'}
    with open(ISARET, 'w', encoding='utf-8') as f:
        json.dump(d, f, ensure_ascii=False, indent=2)
    return d


def yerel_uygulama_calisiyor_mu():
    s = socket.socket()
    s.settimeout(0.6)
    try:
        return s.connect_ex(('127.0.0.1', 5000)) == 0
    finally:
        s.close()


def salt_okunur(yol, ac):
    try:
        os.chmod(yol, (stat.S_IREAD if ac else stat.S_IWRITE | stat.S_IREAD))
    except OSError:
        pass


def indir(oturum, sunucu, ne, hedef):
    r = oturum.get(sunucu + '/api/yedek/indir', params={'ne': ne}, stream=True, timeout=300)
    if r.status_code != 200:
        try:
            mesaj = r.json().get('hata', '')
        except Exception:
            mesaj = r.text[:200]
        raise RuntimeError(f'{ne}: HTTP {r.status_code} — {mesaj}')
    boyut = 0
    with open(hedef, 'wb') as f:
        for parca in r.iter_content(1 << 20):
            f.write(parca)
            boyut += len(parca)
    return boyut


def db_dogrula(yol):
    cn = sqlite3.connect(yol)
    try:
        durum = cn.execute('PRAGMA integrity_check').fetchone()[0]
        if durum != 'ok':
            raise RuntimeError(f'indirilen veritabanı bozuk: {durum}')
        return {'referans': cn.execute('SELECT COUNT(*) FROM referans_listesi').fetchone()[0],
                'vardiya': cn.execute('SELECT COUNT(*) FROM vardiyalar').fetchone()[0],
                'son_vardiya': cn.execute('SELECT MAX(tarih) FROM vardiyalar').fetchone()[0]}
    finally:
        cn.close()


def main():
    ap = argparse.ArgumentParser(description='Canlı sunucunun verisini geliştirme kopyasına indirir')
    ap.add_argument('--sunucu', default=VARSAYILAN_SUNUCU, help='sunucu adresi (yerel ağ)')
    ap.add_argument('--kullanici', help='panel yönetici kullanıcı adı')
    ap.add_argument('--bilgi', action='store_true', help='yalnız sunucu özetini göster, indirme')
    ap.add_argument('--isaretle', action='store_true', help='indirmeden yalnız geliştirme işaretini yaz')
    a = ap.parse_args()

    if a.isaretle:
        d = isaret_yaz()
        print('Geliştirme işareti yazıldı:', ISARET)
        print('  Panelde şerit çıkar, zamanlanmış işler ve AS400 gönderimi bu makinede kapalıdır.')
        return 0

    import requests
    sunucu = a.sunucu.rstrip('/')
    oturum = requests.Session()
    kullanici = a.kullanici or input('Panel yönetici kullanıcı adı: ').strip()
    sifre = getpass.getpass('Şifre (ekranda görünmez): ')
    try:
        r = oturum.post(sunucu + '/api/panel/giris', json={'kullanici_adi': kullanici, 'sifre': sifre},
                        timeout=20)
    except requests.RequestException as e:
        print(f'HATA: {sunucu} adresine ulaşılamadı — {e}')
        print('      Şirket ağında mısınız? Adres farklıysa --sunucu ile verin.')
        return 2
    finally:
        sifre = None
    if r.status_code != 200:
        print('HATA: giriş başarısız —', (r.json().get('hata') if r.headers.get('content-type', '')
                                          .startswith('application/json') else r.status_code))
        return 2

    r = oturum.get(sunucu + '/api/yedek/bilgi', timeout=30)
    if r.status_code != 200:
        print('HATA: sunucu özeti alınamadı (HTTP %s). Sunucu güncel mi, kullanıcı yönetici mi?'
              % r.status_code)
        return 2
    bilgi = r.json()
    print()
    print('SUNUCU   : %s (%s)' % (bilgi.get('host'), sunucu))
    print('  veri   : %s referans · %s vardiya · son vardiya %s · db %.1f MB'
          % (bilgi.get('referans'), bilgi.get('vardiya'), bilgi.get('son_vardiya'),
             (bilgi.get('db_boyut') or 0) / 1e6))
    if bilgi.get('host', '').upper() == socket.gethostname().upper():
        print('DURDURULDU: bu betik SUNUCUNUN KENDİSİNDE çalıştırılmaz (veri kendi üzerine yazılır).')
        return 3
    if bilgi.get('gelistirme_kopyasi'):
        print('DURDURULDU: bağlanılan adres de bir GELİŞTİRME KOPYASI — canlı sunucu değil.')
        return 3
    if os.path.exists(DB):
        try:
            y = db_dogrula(DB)
            print('BU MAKİNE: %s referans · %s vardiya · son vardiya %s'
                  % (y['referans'], y['vardiya'], y['son_vardiya']))
        except Exception as e:
            print('BU MAKİNE: yerel veritabanı okunamadı (%s)' % e)
    if a.bilgi:
        return 0

    if yerel_uygulama_calisiyor_mu():
        print()
        print('DURDURULDU: bu makinede uygulama çalışıyor (port 5000). Veritabanı açıkken')
        print('            değiştirilemez — önce Sistemi_Durdur.bat, sonra tekrar deneyin.')
        return 3

    print()
    print('İndiriliyor…')
    gecici = DB + '.indirilen'
    boyut = indir(oturum, sunucu, 'db', gecici)
    ozet = db_dogrula(gecici)
    print('  veritabanı  %.1f MB · %s referans · bütünlük tamam' % (boyut / 1e6, ozet['referans']))
    if os.path.exists(DB):
        shutil.copy2(DB, DB + '.bak-gelistirme')
    for ek in ('-wal', '-shm'):
        try:
            os.remove(DB + ek)
        except OSError:
            pass
    os.replace(gecici, DB)

    os.makedirs(VERI, exist_ok=True)
    for ne, ad in EXCELLER.items():
        if not (bilgi.get('dosyalar') or {}).get(ne):
            print('  %-24s sunucuda yok — atlandı' % ad)
            continue
        hedef = os.path.join(VERI, ad)
        b = indir(oturum, sunucu, ne, hedef + '.indirilen')
        if os.path.exists(hedef):
            salt_okunur(hedef, False)
            shutil.copy2(hedef, hedef + '.bak-gelistirme')
        os.replace(hedef + '.indirilen', hedef)
        salt_okunur(hedef, True)
        print('  %-24s %.0f KB · salt okunur yapıldı' % (ad, b / 1024))

    isaret_yaz(sunucu, bilgi, indirildi=True)
    print()
    print('TAMAM — bu makine artık sunucunun %s tarihli kopyasıyla çalışıyor.'
          % datetime.now().strftime('%d.%m.%Y %H:%M'))
    print('Önceki yerel veriler: uretim.db.bak-gelistirme ve data\\*.bak-gelistirme')
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print('\nİptal edildi.')
        sys.exit(1)
    except Exception as e:
        print('HATA:', e)
        sys.exit(2)
