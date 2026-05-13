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


def _planli_dongu(hour, minute, fn, etiket):
    """Bir thread olarak çalışır, her gün belirtilen saatte fn() çağırır."""
    print(f'[SCHED] {etiket}: her gün {hour:02d}:{minute:02d}\'da planlandı')
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


def start_scheduler():
    """Flask başlangıcında çağrılır. Daemon thread'leri başlatır."""
    global _started
    if _started:
        return
    _started = True

    # 18:00 — Günlük arşiv
    t = threading.Thread(
        target=_planli_dongu, args=(18, 0, gunluk_arsiv_calistir, 'Günlük Vardiya Arşivi'),
        daemon=True, name='scheduler-arsiv'
    )
    t.start()
    print('[SCHED] Scheduler başlatıldı')
