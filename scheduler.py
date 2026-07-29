# -*- coding: utf-8 -*-
"""
scheduler.py — Basit zamanlanmış görev runner.

Flask uygulaması başlarken bu modülün start_scheduler() fonksiyonu çağrılır;
arka planda bir thread döngüsü kurulur ve gün içinde belirlenmiş saatlerde
fonksiyonlar otomatik tetiklenir.

Şu an aktif görevler:
  • 18:00 — Günlük vardiya arşivi (data/arsiv/YYYY-MM-DD_UretimTakip.xlsx)

Avantajlar (external dep yok):
  - threading.Thread daemon=True → ana süreç kapanınca otomatik son bulur
  - Sleep tabanlı (CPU yormaz)
  - Hata güvenli (exception loglar, döngü durmaz)
"""
import os
import threading
import time
from datetime import datetime, timedelta

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
ARSIV_DIR   = os.path.join(PROJECT_DIR, 'data', 'arsiv')
ARSIV_GUN_LIMIT = 30   # En son 30 günlük yedek tutulur, eskiler silinir


def _arsiv_dosya_adi(tarih=None):
    """data/arsiv/2026-05-13_UretimTakip.xlsx gibi dosya adı üret."""
    if tarih is None:
        tarih = datetime.now()
    return os.path.join(ARSIV_DIR, tarih.strftime('%Y-%m-%d') + '_UretimTakip.xlsx')


def _eski_arsivleri_temizle():
    """ARSIV_GUN_LIMIT günden eski xlsx'leri sil."""
    if not os.path.isdir(ARSIV_DIR):
        return
    simdi = time.time()
    silinen = 0
    for ad in os.listdir(ARSIV_DIR):
        if not ad.endswith('.xlsx'):
            continue
        yol = os.path.join(ARSIV_DIR, ad)
        try:
            yas_gun = (simdi - os.path.getmtime(yol)) / 86400
            if yas_gun > ARSIV_GUN_LIMIT:
                os.remove(yol)
                silinen += 1
        except Exception:
            pass
    if silinen:
        print(f'[OTO-ARSIV] {silinen} eski dosya silindi (>{ARSIV_GUN_LIMIT} gün)')


def gunluk_arsiv_calistir():
    """18:00'da çalışır — vardiya verilerini tarih-damgalı xlsx'e yazar."""
    try:
        from export_excel import export_arsiv, DOSYA_YOLU
        os.makedirs(ARSIV_DIR, exist_ok=True)

        # Önce mevcut export_arsiv'i çalıştır (Masaüstü/UretimTakipArsiv.xlsx)
        sonuc = export_arsiv()
        if not sonuc.get('basarili'):
            print(f'[OTO-ARSIV] Hata: {sonuc.get("hata")}')
            return

        # Sonra dated arşive kopyala (data/arsiv/YYYY-MM-DD_UretimTakip.xlsx)
        import shutil
        hedef = _arsiv_dosya_adi()
        shutil.copy(DOSYA_YOLU, hedef)
        print(f'[OTO-ARSIV] {datetime.now().strftime("%Y-%m-%d %H:%M")} → {hedef}')
        print(f'           {sonuc.get("vardiya_sayisi", 0)} vardiya · '
              f'{sonuc.get("uretim_kayit", 0)} üretim · '
              f'{sonuc.get("durus_kayit", 0)} duruş')

        _eski_arsivleri_temizle()
    except Exception as e:
        print(f'[OTO-ARSIV] BEKLENMEDIK HATA: {e}')
        import traceback
        traceback.print_exc()


def _aralikli_dongu(fn, etiket, aralik_fn):
    """Bir thread olarak çalışır, her `aralik_fn()` saniyede bir fn() çağırır.
    Aralık her turda config'ten okunur (poll_aralik_sn canlı değiştirilebilsin)."""
    print(f'[SCHED] {etiket}: periyodik başlatıldı')
    while True:
        try:
            fn()
        except Exception as e:
            print(f'[SCHED] {etiket} döngü hatası: {e}')
        try:
            time.sleep(max(5, int(aralik_fn())))
        except Exception:
            time.sleep(20)


def test_cihaz_poll():
    """Cofle test cihazlarının yeni başarılı testlerini yerel tabloya çeker."""
    try:
        import cofle_test
        if cofle_test.etkin():
            cofle_test.poll()
    except Exception as e:
        print(f'[SCHED] test_cihaz_poll hata: {e}')


def gunluk_mail_calistir():
    """Gönderim saatinde çalışır — o günün üretim raporunu Excel olarak alıcılara e-postalar.
    SMTP config yoksa/etkin değilse sessizce atlar (özellik kapalı sayılır)."""
    try:
        import mail_raporu
        if not mail_raporu.etkin():
            return  # config yok/etkin değil — sessiz geç
        sonuc = mail_raporu.gunluk_mail_gonder()
        if not sonuc.get('basarili') and not sonuc.get('atlandi'):
            print(f'[OTO-MAIL] Gönderilemedi: {sonuc.get("mesaj")}')
    except Exception as e:
        print(f'[OTO-MAIL] BEKLENMEDIK HATA: {e}')
        import traceback
        traceback.print_exc()


def _planli_dongu(hour, minute, fn, etiket, kacirildi_mi=None):
    """Bir thread olarak çalışır, her gün belirtilen saatte fn() çağırır.

    TELAFİ (2026-07-29): eskiden bu döngü, planlı saat GEÇMİŞSE doğrudan ertesi
    günü bekliyordu. Uygulama o saatte ayakta değilse (restart, çökme, sunucu
    bakımı) o günün koşusu SESSİZCE kayboluyordu — üstelik iş hiç çalışmadığı
    için panele atlama kaydı bile düşmüyordu, yani "dün akşam neden olmadı"
    sorusunun izi kalmıyordu.

    kacirildi_mi: parametresiz, True/False dönen bir fonksiyon. Yalnız
    "bugünün planlı saati geçmiş" durumunda çağrılır ve "bu iş bugün HİÇ
    koşmadı mı?" sorusunu yanıtlar. True dönerse iş bir kez ÇALIŞTIRILIR.
    Verilmezse telafi YAPILMAZ (eski davranış) — bu bilinçli: her iş idempotent
    değil (örn. günlük rapor maili iki kez gönderilmemeli)."""
    print(f'[SCHED] {etiket}: her gün {hour:02d}:{minute:02d}\'da planlandı'
          + (' (telafili)' if kacirildi_mi else ''))
    # ── Açılışta tek seferlik telafi kontrolü ──
    if kacirildi_mi:
        try:
            now = datetime.now()
            if now.replace(hour=hour, minute=minute, second=0, microsecond=0) <= now and kacirildi_mi():
                print(f'[SCHED] {etiket}: bugünkü {hour:02d}:{minute:02d} KAÇIRILMIŞ '
                      f'(uygulama o saatte ayakta değildi) → telafi çalıştırılıyor')
                fn()
        except Exception as e:
            print(f'[SCHED] {etiket} telafi hatası: {e}')
    while True:
        try:
            now = datetime.now()
            hedef = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
            if hedef <= now:
                hedef += timedelta(days=1)
            bekle_sn = (hedef - now).total_seconds()
            time.sleep(bekle_sn)
            print(f'[SCHED] {etiket} tetiklendi ({datetime.now().strftime("%H:%M:%S")})')
            fn()
        except Exception as e:
            print(f'[SCHED] {etiket} döngü hatası: {e}, 60sn bekle')
            time.sleep(60)


_started = False  # Yalnızca tek thread spawn etmek için


def start_scheduler(ek_gorevler=None):
    """Flask başlangıcında çağrılır. Daemon thread'leri başlatır.
    ek_gorevler: [(saat, dakika, fn, etiket), ...] veya
                 [(saat, dakika, fn, etiket, kacirildi_mi), ...] — app.py'nin kendi
    planlı işleri (örn. AS400 oto koşuları) buradan verilir; circular import olmaz.
    5. eleman verilirse o iş için açılışta TELAFİ yapılır (bkz. _planli_dongu)."""
    global _started
    if _started:
        return
    _started = True

    for i, gorev in enumerate(ek_gorevler or []):
        h, m, fn, etiket = gorev[:4]
        kacirildi_mi = gorev[4] if len(gorev) > 4 else None
        te = threading.Thread(
            target=_planli_dongu, args=(h, m, fn, etiket, kacirildi_mi),
            daemon=True, name=f'scheduler-ek-{i}'
        )
        te.start()

    # 18:00 — Günlük arşiv
    t = threading.Thread(
        target=_planli_dongu, args=(18, 0, gunluk_arsiv_calistir, 'Günlük Vardiya Arşivi'),
        daemon=True, name='scheduler-arsiv'
    )
    t.start()

    # Günlük üretim raporu e-postası — saat mail_config.json'dan (varsayılan 17:00).
    # SMTP config yoksa thread yine kurulur ama gunluk_mail_calistir sessizce atlar
    # (config sonradan doldurulunca ertesi gün otomatik devreye girer).
    try:
        import mail_raporu
        _msaat, _mdk = mail_raporu.gonderim_saati()
    except Exception:
        _msaat, _mdk = 17, 0
    tm = threading.Thread(
        target=_planli_dongu, args=(_msaat, _mdk, gunluk_mail_calistir, 'Günlük Üretim Raporu Maili'),
        daemon=True, name='scheduler-mail'
    )
    tm.start()

    # Cofle test cihazı poller — sadece config etkinse thread aç
    try:
        import cofle_test
        if cofle_test.etkin():
            tt = threading.Thread(
                target=_aralikli_dongu,
                args=(test_cihaz_poll, 'Cofle Test Cihazı Sayaç', cofle_test.poll_aralik_sn),
                daemon=True, name='scheduler-testcihaz'
            )
            tt.start()
        else:
            print('[SCHED] Cofle test cihazı entegrasyonu KAPALI (cofle_test_config.json yok/etkin değil)')
    except Exception as e:
        print(f'[SCHED] Test cihazı poller başlatılamadı: {e}')

    print('[SCHED] Scheduler başlatıldı')
