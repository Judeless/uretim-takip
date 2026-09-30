# -*- coding: utf-8 -*-
"""
Sunucudan_Veri_Cek.py — CANLI sunucunun verisini bu makineye (geliştirme kopyası) indirir.

NEDEN: canlı sistem sunucuda çalışıyor; laptopta aynı kodun AYRI veritabanı ve AYRI
Excel'i var. İkisi ayrışınca "hangisi gerçek" karışıyor (2026-09-30: montaj planı
laptoptaki eski referans listesiyle ölçüldü, sunucuyla tutmadı).

KURAL: tek kaynak SUNUCU. Veri TEK YÖNDE akar: sunucu → bu makine.
Bu betik sunucunun verisine HİÇBİR ŞEY YAZMAZ; yalnız okur ve yerel kopyayı değiştirir.

YAPTIĞI
  1. Sunucuya bağlanır (yedek anahtarı varsa onunla, yoksa yönetici şifresini sorarak).
  2. uretim.db'nin tutarlı kopyasını ve data\\ altındaki Excel'leri indirir.
  3. Yereldeki eskileri yedekler (uretim.db.bak-gelistirme, *.xlsx.bak-gelistirme).
  4. Excel'leri SALT OKUNUR yapar — yanlışlıkla burada düzenlenmesin.
  5. data\\GELISTIRME_KOPYASI.json işaretini yazar: panelde şerit çıkar, zamanlanmış
     işler başlamaz, AS400'e yazan uçlar kapanır.

KULLANIM
  python Sunucudan_Veri_Cek.py --kur            # BİR KEZ: yedek anahtarını kur + indir
  python Sunucudan_Veri_Cek.py                  # indir (anahtar varsa şifre sormaz)
  python Sunucudan_Veri_Cek.py --otomatik --eskiyse 4
                                                # hiç soru sormaz; veri 4 saatten yeniyse dokunmaz
  python Sunucudan_Veri_Cek.py --bilgi          # yalnız sunucu özetini göster
  python Sunucudan_Veri_Cek.py --isaretle       # indirmeden yalnız işareti yaz
  python Sunucudan_Veri_Cek.py --anahtar-sil    # bu makinedeki yedek anahtarını sil

YEDEK ANAHTARI (--kur): yönetici şifresi bir kez sorulur; sunucu YALNIZ yedek indirmeye
yarayan bir anahtar üretir, anahtar doğrudan Windows Kimlik Kasası'na yazılır (ekrana
basılmaz, dosyaya yazılmaz). Sonraki çekimler şifresizdir. Sunucuda yalnız özeti durur;
yalnız yerel ağdan geçerlidir; --kur yeniden çalıştırılınca eskisi iptal olur.

ÇIKIŞ KODLARI: 0 tamam/güncel · 2 bağlantı-giriş hatası · 3 güvenlik freni ·
               4 otomatik kipte anahtar yok

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
KASA_SERVIS = 'cofle-forge-yedek'          # Windows Kimlik Kasası (keyring) servis adı
ANAHTAR_BASLIGI = 'X-Yedek-Anahtari'

for _a in (sys.stdout, sys.stderr):
    try:
        _a.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass


# ── Kimlik Kasası ─────────────────────────────────────────────────────────
def kasa():
    try:
        import keyring
        return keyring
    except ImportError:
        return None


def anahtar_oku(sunucu):
    k = kasa()
    try:
        return (k.get_password(KASA_SERVIS, sunucu) or '') if k else ''
    except Exception:
        return ''


def anahtar_yaz(sunucu, anahtar):
    k = kasa()
    if not k:
        raise RuntimeError("'keyring' modülü kurulu değil — 'pip install keyring'")
    k.set_password(KASA_SERVIS, sunucu, anahtar)


def anahtar_sil(sunucu):
    k = kasa()
    try:
        if k:
            k.delete_password(KASA_SERVIS, sunucu)
        return True
    except Exception:
        return False


# ── İşaret dosyası ────────────────────────────────────────────────────────
def isaret_oku():
    try:
        with open(ISARET, encoding='utf-8-sig') as f:
            return json.load(f)
    except Exception:
        return None


def isaret_yaz(sunucu='', bilgi=None, indirildi=False):
    os.makedirs(VERI, exist_ok=True)
    onceki = isaret_oku() or {}
    d = {'sunucu': sunucu or onceki.get('sunucu', ''),
         'cekildi': datetime.now().strftime('%Y-%m-%d %H:%M') if indirildi else onceki.get('cekildi', ''),
         'sunucu_host': (bilgi or {}).get('host', onceki.get('sunucu_host', '')),
         'referans': (bilgi or {}).get('referans', onceki.get('referans')),
         'son_vardiya': (bilgi or {}).get('son_vardiya', onceki.get('son_vardiya')),
         'bu_makine': socket.gethostname(),
         'not': 'Bu dosya varken kurulum GELISTIRME KOPYASIDIR. Canli yapmak icin silin.'}
    with open(ISARET, 'w', encoding='utf-8') as f:
        json.dump(d, f, ensure_ascii=False, indent=2)
    return d


def veri_yasi_saat():
    """Son çekimden bu yana geçen saat; hiç çekilmediyse None."""
    d = isaret_oku() or {}
    try:
        t = datetime.strptime(d.get('cekildi') or '', '%Y-%m-%d %H:%M')
    except ValueError:
        return None
    return (datetime.now() - t).total_seconds() / 3600.0


# ── Yardımcılar ───────────────────────────────────────────────────────────
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


def hata_metni(r):
    try:
        return r.json().get('hata', '') or str(r.status_code)
    except Exception:
        return 'HTTP %s' % r.status_code


def indir(oturum, sunucu, ne, hedef, basliklar):
    r = oturum.get(sunucu + '/api/yedek/indir', params={'ne': ne}, headers=basliklar,
                   stream=True, timeout=300)
    if r.status_code != 200:
        raise RuntimeError('%s indirilemedi: %s' % (ne, hata_metni(r)))
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
            raise RuntimeError('indirilen veritabanı bozuk: %s' % durum)
        return {'referans': cn.execute('SELECT COUNT(*) FROM referans_listesi').fetchone()[0],
                'vardiya': cn.execute('SELECT COUNT(*) FROM vardiyalar').fetchone()[0],
                'son_vardiya': cn.execute('SELECT MAX(tarih) FROM vardiyalar').fetchone()[0]}
    finally:
        cn.close()


def yonetici_girisi(oturum, sunucu, kullanici=None):
    """Yönetici oturumu açar (şifre sorulur, saklanmaz). Hata → RuntimeError."""
    kullanici = kullanici or input('Panel yönetici kullanıcı adı: ').strip()
    sifre = getpass.getpass('Şifre (ekranda görünmez): ')
    try:
        r = oturum.post(sunucu + '/api/panel/giris',
                        json={'kullanici_adi': kullanici, 'sifre': sifre}, timeout=20)
    finally:
        sifre = None
    if r.status_code != 200:
        raise RuntimeError('giriş başarısız — ' + hata_metni(r))


def main():
    ap = argparse.ArgumentParser(description='Canlı sunucunun verisini geliştirme kopyasına indirir')
    ap.add_argument('--sunucu', default=VARSAYILAN_SUNUCU, help='sunucu adresi (yerel ağ)')
    ap.add_argument('--kullanici', help='panel yönetici kullanıcı adı')
    ap.add_argument('--kur', action='store_true', help='yedek anahtarını kur (bir kez) ve indir')
    ap.add_argument('--otomatik', action='store_true', help='hiç soru sorma; anahtar yoksa çık')
    ap.add_argument('--eskiyse', type=float, metavar='SAAT',
                    help='yerel veri bu kadar saatten yeniyse indirme')
    ap.add_argument('--bilgi', action='store_true', help='yalnız sunucu özetini göster, indirme')
    ap.add_argument('--isaretle', action='store_true', help='indirmeden yalnız geliştirme işaretini yaz')
    ap.add_argument('--anahtar-sil', action='store_true', help='bu makinedeki yedek anahtarını sil')
    a = ap.parse_args()
    sunucu = a.sunucu.rstrip('/')

    if a.isaretle:
        isaret_yaz()
        print('Geliştirme işareti yazıldı:', ISARET)
        print('  Panelde şerit çıkar; zamanlanmış işler ve AS400 gönderimi bu makinede kapalıdır.')
        return 0
    if a.anahtar_sil:
        print('Yedek anahtarı bu makineden silindi.' if anahtar_sil(sunucu)
              else 'Silinecek anahtar bulunamadı.')
        return 0

    yas = veri_yasi_saat()
    if a.eskiyse is not None and yas is not None and yas < a.eskiyse and not a.kur:
        print('GÜNCEL: yerel veri %.1f saat önce çekilmiş (eşik %.1f saat) — indirilmedi.'
              % (yas, a.eskiyse))
        return 0

    import requests
    oturum = requests.Session()
    anahtar = '' if a.kur else anahtar_oku(sunucu)
    basliklar = {ANAHTAR_BASLIGI: anahtar} if anahtar else {}
    try:
        if a.kur:
            yonetici_girisi(oturum, sunucu, a.kullanici)
            r = oturum.post(sunucu + '/api/yedek/anahtar_uret', timeout=30)
            if r.status_code != 200:
                raise RuntimeError('yedek anahtarı üretilemedi — ' + hata_metni(r))
            anahtar_yaz(sunucu, r.json()['anahtar'])
            basliklar = {ANAHTAR_BASLIGI: anahtar_oku(sunucu)}
            print('Yedek anahtarı kuruldu (Windows Kimlik Kasası). Bundan sonra şifre sorulmaz.')
        elif not anahtar:
            if a.otomatik:
                print('ANAHTAR YOK: otomatik çekim için bir kez  Sunucudan_Veri_Cek.bat --kur  çalıştırın.')
                return 4
            yonetici_girisi(oturum, sunucu, a.kullanici)
        r = oturum.get(sunucu + '/api/yedek/bilgi', headers=basliklar, timeout=30)
    except requests.RequestException as e:
        print('HATA: %s adresine ulaşılamadı — %s' % (sunucu, e))
        print('      Şirket ağında mısınız? Adres farklıysa --sunucu ile verin.')
        return 2
    except RuntimeError as e:
        print('HATA:', e)
        return 2
    if r.status_code == 401 and anahtar:
        print('HATA: yedek anahtarı sunucuda geçersiz (yeniden üretilmiş ya da iptal edilmiş).')
        print('      Sunucudan_Veri_Cek.bat --kur  ile yeniden kurun.')
        return 4 if a.otomatik else 2
    if r.status_code != 200:
        print('HATA: sunucu özeti alınamadı — %s. Sunucu güncel mi?' % hata_metni(r))
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
            print('BU MAKİNE: %s referans · %s vardiya · son vardiya %s%s'
                  % (y['referans'], y['vardiya'], y['son_vardiya'],
                     (' · %.1f saat önce çekildi' % yas) if yas is not None else ' · hiç çekilmemiş'))
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
    boyut = indir(oturum, sunucu, 'db', gecici, basliklar)
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
        b = indir(oturum, sunucu, ne, hedef + '.indirilen', basliklar)
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
