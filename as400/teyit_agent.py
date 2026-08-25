# -*- coding: utf-8 -*-
"""
teyit_agent.py — AS400 robot çalıştırıcı mini servis (SERVER için).

NEDEN VAR: Server'da cofle-app NSSM SERVİSİ olarak Session 0'da koşar; PCOMM
emülatörü (pcsws.exe A/B oturumları) ise promanage'ın RDP oturumundadır. Windows
oturum izolasyonu yüzünden Session 0'dan başlatılan cscript, PCOMM COM
oturumlarını GÖREMEZ. Bu agent RDP oturumunda (PCOMM'un yanında) çalışır;
ana app robot çağrılarını buraya HTTP ile devreder.

ÇALIŞTIRMA: promanage RDP oturumunda Teyit_Agent_Baslat.bat (Başlangıç
klasörüne kısayol koy → logonda otomatik kalkar). Konsol penceresi açık kalır —
her robot koşusu ekranda görünür.

GÜVENLİK: Yalnız 127.0.0.1'e bağlanır (ağdan ERİŞİLEMEZ; tek istemci aynı
makinedeki cofle-app). Script beyaz-listeli (yalnız teyit_gir.js / cfi_gir.js) —
keyfi komut çalıştırılamaz.

LAPTOP'ta bu agent GEREKMEZ: app konsol uygulaması olarak zaten PCOMM ile aynı
oturumda; app.py agent'ı bulamayınca (127.0.0.1:5010 kapalı) yerel subprocess'e
düşer — davranış değişmez.
"""
import os
import subprocess
import sys
import threading
import time

from flask import Flask, request, jsonify

KOK      = os.path.dirname(os.path.abspath(__file__))     # .../as400
CSCRIPT  = r'C:\Windows\SysWOW64\cscript.exe'             # 32-bit (PCOMM COM 32-bit)
IZINLI   = ('teyit_gir.js', 'cfi_gir.js', 'transfer_iptal.js',
            'oturum_ac.js')                                       # beyaz liste — başka script ÇALIŞMAZ
# Şifreyi ortam değişkeniyle alan scriptler (argümanla ASLA geçilmez — argümanlar
# bu konsola ve Windows süreç listesine düşer). Bunların args'ı da loglanmaz.
SIFRE_ISTEYEN = ('oturum_ac.js',)
PORT     = 5010

app = Flask(__name__)
_KILIT = threading.Lock()   # Session B tek — koşular sıralanır (app zaten sıralar; ek emniyet)


@app.route('/durum')
def durum():
    """Sağlık — app.py yönlendirme kararını buradan verir (agent=cofle-teyit).
    HIÇBIR dış komut (tasklist vb.) ÇAĞIRMAZ — anında döner (bazı sunucularda
    tasklist asılıp agent'ı startta kilitliyordu; PCOMM sayacı kaldırıldı)."""
    return jsonify({'agent': 'cofle-teyit', 'ok': True, 'gozcu': dict(_GOZCU)})


@app.route('/sifre')
def sifre():
    """AS400 keyring şifresini yerel app'e verir (YALNIZ 127.0.0.1). Server'da
    cofle-app NSSM servisi LocalSystem'da koşar → promanage keyring'ini göremez;
    bu agent promanage oturumunda olduğu için görür. Şifre keyring'de zaten
    şifreli durur — burada yalnız bellekte, loopback üzerinden aktarılır."""
    try:
        import keyring
        import as400_config as cfg
        pw = keyring.get_password(cfg.KEYRING_SERVICE, cfg.DB_KULLANICI)
        return jsonify({'sifre': pw or ''})
    except Exception as e:
        return jsonify({'sifre': '', 'hata': str(e)}), 500


@app.route('/calistir', methods=['POST'])
def calistir():
    """Robot cscript'ini bu oturumda çalıştır.
    Body: {script: 'teyit_gir.js'|'cfi_gir.js', args: [str...], timeout: sn}
    Döner: 200 {rc, cikti} | 400 geçersiz istek | 504 timeout."""
    d = request.get_json(silent=True) or {}
    script = str(d.get('script') or '')
    args = d.get('args') or []
    try:
        timeout = min(300, max(10, int(d.get('timeout') or 150)))
    except (TypeError, ValueError):
        timeout = 150
    if script not in IZINLI:
        return jsonify({'hata': f'izinsiz script: {script}'}), 400
    if not isinstance(args, list) or len(args) > 8:
        return jsonify({'hata': 'geçersiz args'}), 400
    temiz = []
    for a in args:
        a = str(a)
        # Kontrol karakteri yok, makul uzunluk. '/' serbest (article kodlarında var:
        # 10.300.1941/20W). Shell KULLANILMIYOR (liste-form subprocess) → injection yok.
        if len(a) > 64 or any(ch in a for ch in '\r\n\x00'):
            return jsonify({'hata': f'geçersiz arg: {a[:30]}'}), 400
        temiz.append(a)
    yol = os.path.join(KOK, script)
    # Şifre isteyen scriptlerde ARGÜMANLARI YAZDIRMA (konsol + kayıt hijyeni)
    if script in SIFRE_ISTEYEN:
        print(f'[AGENT] {script} (args gizlendi) (timeout={timeout}s)')
    else:
        print(f'[AGENT] {script} {" ".join(temiz)} (timeout={timeout}s)')
    # Şifre YALNIZCA çocuk sürecin ortamına konur: argümana, dosyaya, loga girmez.
    ortam = None
    if script in SIFRE_ISTEYEN:
        try:
            import as400_config as _cfg
            _pw = _cfg.sifre_al()
        except Exception as _e:
            _pw = None
            print(f'[AGENT]   -> UYARI: kasadan şifre okunamadı: {_e}')
        if not _pw:
            return jsonify({'hata': 'AS400 şifresi kasada yok (kaydet_sifre.py çalıştırın)'}), 503
        ortam = dict(os.environ)
        ortam['COFLE_AS400_PW'] = _pw
    with _KILIT:
        try:
            pr = subprocess.run([CSCRIPT, '//nologo', yol] + temiz,
                                capture_output=True, timeout=timeout, env=ortam)
            cikti = (pr.stdout or b'').decode('cp1254', errors='replace')
            # stderr DE geri verilir (2026-08-17): PCOMM kapaliyken teyit_gir.js
            # SetConnectionByName("B")'de firliyor, JScript hatasi STDERR'e gidiyor,
            # stdout BOS kaliyordu. App bunu gormedigi icin operatore bombos
            # "Robot iptal:" yaziliyordu — 14 satirin 14'u sebepsiz.
            hata_cikti = (pr.stderr or b'').decode('cp1254', errors='replace')
            son = [l for l in cikti.splitlines() if l.strip()]
            hson = ' '.join(hata_cikti.split())[:200]
            print(f'[AGENT]   -> rc={pr.returncode} | {son[-1] if son else "(bos)"}'
                  + (f' | STDERR: {hson}' if hson else ''))
            return jsonify({'rc': pr.returncode, 'cikti': cikti, 'hata_cikti': hata_cikti})
        except subprocess.TimeoutExpired:
            print(f'[AGENT]   -> TIMEOUT ({timeout}s) — cscript kill edildi')
            return jsonify({'hata': 'timeout'}), 504



# ════════════════════════════════════════════════════════════════════
#   OTURUM GOZCUSU (2026-08-17, kullanici izinde iken sistem ayakta kalsin)
# ════════════════════════════════════════════════════════════════════
# Sunucudaki Session B uzun sure islem yapilmayinca dusuyor (AS400 QINACTITV ya
# da baglanti kopmasi) ve teyit robotu bir daha calisamiyor. Elle sign-on yapacak
# kimse olmadiginda sistem sessizce durur.
#
# BU, KURULUM KILAVUZUNDAKI "sign-on INSAN isi" KURALINI BILEREK GEVSETIR —
# kullanicinin acik istegi (2026-08-17). Guvenlik dengesi:
#   · sifre YINE kasada (Windows Credential Manager / DPAPI), kodda-dosyada DEGIL
#   · sifre cocuk surece YALNIZ ortam degiskeniyle gecer (arguman/log/disk YOK)
#   · varsayilan KAPALI — oturum_config.json'da etkin:true yapilmadan calismaz
#   · robot yanlis sifre DENEMEZ; sign-on reddedilirse durur (AS400 profil
#     kilitlenmesi riskine girilmez) ve bir sonraki turda tekrar dener
OTURUM_CFG = os.path.join(KOK, 'oturum_config.json')

# Gozcunun son bilinen durumu — /durum bunu OKUR (dosya/komut erisimi YAPMAZ).
# NEDEN (2026-08-17): "agent yeniden basladi ama gozcu acildi mi?" sorusunun
# cevabi yalnizca konsol penceresine bakmak olmamali; disaridan sorulabilmeli.
_GOZCU = {'etkin': False, 'kullanici': '', 'aralik_sn': 0,
          'son_durum': None, 'son_zaman': None, 'son_detay': '',
          # dongu_zaman = thread'in HER turda guncelledigi nabiz. son_zaman yalniz
          # kontrol BITINCE dolar; ikisini ayirmak 'thread olmus' ile 'kontrol
          # suruyor'u ayirt ettirir. thread_hata = donguyu oldurmus istisna.
          'dongu_zaman': None, 'thread_hata': ''}


def _oturum_ayar():
    """oturum_config.json'u HER TURDA taze oku (dosya degisince agent restart
    gerekmesin). Dosya yoksa/bozuksa gozcu KAPALI sayilir."""
    try:
        import json
        with open(OTURUM_CFG, 'r', encoding='utf-8-sig') as f:
            d = json.load(f) or {}
        return {
            'etkin': bool(d.get('etkin')),
            'kullanici': str(d.get('kullanici') or '').strip(),
            'aralik_sn': max(60, int(d.get('aralik_sn') or 300)),
        }
    except FileNotFoundError:
        return {'etkin': False, 'kullanici': '', 'aralik_sn': 300}
    except Exception as e:
        print(f'[GOZCU] config okunamadi ({e}) — kapali sayiliyor')
        return {'etkin': False, 'kullanici': '', 'aralik_sn': 300}


def _sessiz_sebep(rc, hata_cikti):
    """oturum_ac.js HIC stdout uretmedi -> sebebi anlasilir yaz (2026-08-20).

    NEDEN: burada stderr YAKALANIYOR ama HIC KULLANILMIYORDU (robot tarafinda
    duzeltilmisti, gozcude atlanmis). Sonuc: /durum "BILINMEYEN, rc=0 | " diyordu
    ve gercek hata gorunmuyordu. Sahada yasandi (2026-08-20): AS400 alt sistemi
    kapatilinca Session B oldu, gozcu her 5 dk deniyor ama NEDEN basarisiz oldugu
    hicbir yerde yazmiyordu.

    oturum_ac.js kendi bildigi hatalarda 'SONUC=IPTAL' BASAR. Hic cikti yoksa
    script daha ActiveXObject satirinda olmus demektir -> PCOMM yok/kapali."""
    st = ' '.join((hata_cikti or '').split())[:300]
    if 'ECL37110' in st or 'emulazione' in st.lower() or 'emulation interface' in st.lower():
        # PENCERELER ACIK OLMASI YETERLI DEGIL (2026-08-17 olayi, bkz.
        # SERVER_AS400_KURULUM.md): iki pcsws.exe ayni oturumda calisiyor ve sign-on
        # yapilmisken autECLConnList 0 baglanti goruyordu — pencereler otomasyon
        # katmanina KAYITLI DEGILDI. Bu yuzden mesaj OLCULEBILIR dogrulamayi soyler.
        return (f'rc={rc} | PCOMM emulasyon arayuzu YOK (ECL37110): PCOMM kurulu ama '
                f'agentin kostugu Windows oturumunda otomasyona KAYITLI emulator '
                f'oturumu yok. DOGRULA: as400 klasorunde Robot_Tani.bat calistir — '
                f'otomasyonun gordugu baglanti sayisi 2, adlar [A] ve [B] olmali. '
                f'0 ise: TUM PCOMM pencerelerini kapat, promanage oturumunda '
                f'YONETICI OLMADAN once A sonra B ac, sign-on yap, taniyi tekrarla. '
                f'RDP oturumunu Disconnect et, LOGOFF ETME. | {st}')
    if (not st) or any(k in st for k in ('ActiveX', 'Automation', '80040154', '800401F3',
                                         'PCOMM', 'autECL', 'SetConnectionByName', 'sunucu')):
        return (f'rc={rc} | Session B ACILAMADI: PCOMM penceresi kapali ya da COM nesnesi '
                f'olusturulamiyor. Sunucuda PCOMM A+B pencerelerini acip sign-on yapin '
                f'(as400\\Robot_Tani.bat teshis eder). | '
                + (st or 'stderr bos — cscript hic calismamis olabilir'))
    return f'rc={rc} | {st}'


def _oturum_kontrol_et(kullanici):
    """oturum_ac.js'i calistir. Doner: (sonuc_etiketi, son_satir).
    _KILIT ile korunur → teyit robotu kosarken ASLA araya girmez."""
    try:
        import as400_config as _cfg
        pw = _cfg.sifre_al()
    except Exception as e:
        return ('KASA-HATA', f'sifre okunamadi: {e}')
    if not pw:
        return ('KASA-BOS', 'AS400 sifresi kasada yok (kaydet_sifre.py)')
    ortam = dict(os.environ)
    ortam['COFLE_AS400_PW'] = pw
    with _KILIT:
        try:
            pr = subprocess.run([CSCRIPT, '//nologo', os.path.join(KOK, 'oturum_ac.js'), kullanici],
                                capture_output=True, timeout=180, env=ortam)
        except subprocess.TimeoutExpired:
            return ('TIMEOUT', 'oturum_ac.js 180 sn icinde bitmedi')
    cikti = (pr.stdout or b'').decode('cp1254', errors='replace')
    hata_cikti = (pr.stderr or b'').decode('cp1254', errors='replace')
    satirlar = [l.strip() for l in cikti.splitlines() if l.strip()]
    for etiket in ('SONUC=OK', 'SONUC=ZATEN', 'SONUC=IPTAL'):
        if etiket in cikti:
            return (etiket.split('=')[1], ' | '.join(satirlar[-3:]))
    # Buraya dusuldiyse SONUC satiri hic basilmamis -> sebep stderr'de.
    # rc'ye GUVENME: cscript, JScript calisma hatasinda bile 0 ile cikabiliyor.
    if not satirlar:
        return ('BILINMEYEN', _sessiz_sebep(pr.returncode, hata_cikti))
    return ('BILINMEYEN', f'rc={pr.returncode} | ' + ' | '.join(satirlar[-3:])
            + (f' | stderr: {" ".join(hata_cikti.split())[:200]}' if hata_cikti.strip() else ''))


def _gozcu_dongusu():
    """Arka plan: periyodik olarak Session B'yi kontrol eder, dusmusse geri getirir.
    SESSIZ CALISIR — yalniz DURUM DEGISINCE ekrana yazar (konsol sismesin)."""
    son_durum = None
    while True:
      try:
        _GOZCU['dongu_zaman'] = time.strftime('%Y-%m-%d %H:%M:%S')
        ayar = _oturum_ayar()
        _GOZCU['etkin'] = ayar['etkin']
        _GOZCU['kullanici'] = ayar['kullanici']
        _GOZCU['aralik_sn'] = ayar['aralik_sn']
        if not ayar['etkin'] or not ayar['kullanici']:
            if son_durum != 'kapali':
                print('[GOZCU] kapali (oturum_config.json etkin:false ya da kullanici bos)')
                son_durum = 'kapali'
            time.sleep(60)
            continue
        durum, detay = _oturum_kontrol_et(ayar['kullanici'])
        _GOZCU['son_durum'] = durum
        _GOZCU['son_detay'] = detay[:300]
        _GOZCU['son_zaman'] = time.strftime('%Y-%m-%d %H:%M:%S')
        if durum == 'OK':
            print(f'[GOZCU] Session B DUSMUSTU -> yeniden giris yapildi. {detay}')
        elif durum in ('IPTAL', 'TIMEOUT', 'KASA-BOS', 'KASA-HATA', 'BILINMEYEN'):
            print(f'[GOZCU] SORUN ({durum}): {detay}')
        elif durum != son_durum:
            print(f'[GOZCU] Session B saglikli (kontrol araligi {ayar["aralik_sn"]} sn)')
        son_durum = durum
        _GOZCU['thread_hata'] = ''
        time.sleep(ayar['aralik_sn'])
      except Exception as _e:
        # DAEMON THREAD SESSIZCE OLMESIN (2026-08-17): buradaki bir istisna
        # gozcuyu KALICI durdurur, /durum ise 'etkin:true' gostermeye devam
        # ederdi — en kotu yanilgi. Hatayi kaydet, bekle, DEVAM ET.
        _GOZCU['thread_hata'] = f'{type(_e).__name__}: {_e}'[:200]
        print(f'[GOZCU] DONGU HATASI (devam ediliyor): {_GOZCU["thread_hata"]}')
        time.sleep(60)


def _quickedit_kapat():
    """Windows konsol QuickEdit/Mark modunu KAPAT — pencereye tiklaninca agent
    DONMASIN (tiklama process'i durduruyor; kritik altyapi icin kabul edilemez)."""
    try:
        import ctypes
        k = ctypes.windll.kernel32
        h = k.GetStdHandle(-10)   # STD_INPUT_HANDLE
        mode = ctypes.c_uint()
        if k.GetConsoleMode(h, ctypes.byref(mode)):
            ENABLE_QUICK_EDIT, ENABLE_EXTENDED_FLAGS = 0x0040, 0x0080
            k.SetConsoleMode(h, (mode.value & ~ENABLE_QUICK_EDIT) | ENABLE_EXTENDED_FLAGS)
    except Exception:
        pass


if __name__ == '__main__':
    # STDOUT'U UTF-8'E SABITLE (2026-08-17): banner ve gozcu satirlari kutu/Turkce
    # karakter iceriyor. Konsol kod sayfasi cp1254 ise (ornegin cikti bir dosyaya
    # yonlendirilirse) print UnicodeEncodeError firlatiyor ve AGENT DAHA ACILIRKEN
    # OLUYOR. Gozetimsiz kosacak bir kopru icin kabul edilemez; errors='replace'
    # ile en kotu ihtimalle karakter bozulur, servis AYAKTA KALIR.
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass
    _quickedit_kapat()
    print('╔══════════════════════════════════════════════╗')
    print('║  COFLE TEYIT AGENT — PCOMM oturumu köprüsü   ║')
    print('╚══════════════════════════════════════════════╝')
    print(f'Dinleme: 127.0.0.1:{PORT} (yalniz yerel) — agent HAZIR')
    # ── OTURUM GOZCUSU ──
    # DIKKAT: bu blok bir onceki denemede sessizce UYGULANMAMISTI (metin degisimi
    # eslesmedi, kontrol edilmedi) — gozcu fonksiyonu dosyada vardi ama THREAD HIC
    # BASLAMIYORDU. /durum'da aralik_sn=0 gorulmesi bunun izidir (config bir kez
    # bile okunsa 60/300 olurdu). Degisiklikler artik dogrulanarak yapiliyor.
    _a = _oturum_ayar()
    _GOZCU.update({'etkin': _a['etkin'], 'kullanici': _a['kullanici'],
                   'aralik_sn': _a['aralik_sn']})
    if _a['etkin'] and _a['kullanici']:
        print(f"Oturum gözcüsü AÇIK — kullanıcı {_a['kullanici']}, her {_a['aralik_sn']} sn kontrol")
    else:
        print('Oturum gözcüsü KAPALI (as400/oturum_config.json → etkin:true ile açılır)')
    threading.Thread(target=_gozcu_dongusu, daemon=True).start()
    print('Bu pencereyi KAPATMA — robot koşuları burada görünür.' + chr(10))
    app.run(host='127.0.0.1', port=PORT, threaded=True)
