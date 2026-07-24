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
import threading

from flask import Flask, request, jsonify

KOK      = os.path.dirname(os.path.abspath(__file__))     # .../as400
CSCRIPT  = r'C:\Windows\SysWOW64\cscript.exe'             # 32-bit (PCOMM COM 32-bit)
IZINLI   = ('teyit_gir.js', 'cfi_gir.js')                 # beyaz liste — başka script ÇALIŞMAZ
PORT     = 5010

app = Flask(__name__)
_KILIT = threading.Lock()   # Session B tek — koşular sıralanır (app zaten sıralar; ek emniyet)


def _pcomm_sayisi():
    """Bu oturumda görünen pcsws.exe (PCOMM pencere) sayısı — tanı için."""
    try:
        out = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq pcsws.exe'],
                             capture_output=True, timeout=5).stdout.decode(errors='replace')
        return out.lower().count('pcsws.exe')
    except Exception:
        return -1


@app.route('/durum')
def durum():
    """Sağlık + tanı. app.py yönlendirme kararını buradan verir (agent=cofle-teyit)."""
    return jsonify({'agent': 'cofle-teyit', 'ok': True, 'pcomm': _pcomm_sayisi()})


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
    print(f'[AGENT] {script} {" ".join(temiz)} (timeout={timeout}s)')
    with _KILIT:
        try:
            pr = subprocess.run([CSCRIPT, '//nologo', yol] + temiz,
                                capture_output=True, timeout=timeout)
            cikti = (pr.stdout or b'').decode('cp1254', errors='replace')
            son = [l for l in cikti.splitlines() if l.strip()]
            print(f'[AGENT]   -> rc={pr.returncode} | {son[-1] if son else "(bos)"}')
            return jsonify({'rc': pr.returncode, 'cikti': cikti})
        except subprocess.TimeoutExpired:
            print(f'[AGENT]   -> TIMEOUT ({timeout}s) — cscript kill edildi')
            return jsonify({'hata': 'timeout'}), 504


if __name__ == '__main__':
    print('╔══════════════════════════════════════════════╗')
    print('║  COFLE TEYIT AGENT — PCOMM oturumu köprüsü   ║')
    print('╚══════════════════════════════════════════════╝')
    print(f'Dinleme: 127.0.0.1:{PORT} (yalnız yerel) · PCOMM pencere: {_pcomm_sayisi()}')
    print('Bu pencereyi KAPATMA — robot koşuları burada görünür.\n')
    app.run(host='127.0.0.1', port=PORT, threaded=True)
