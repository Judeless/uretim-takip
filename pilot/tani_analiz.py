# -*- coding: utf-8 -*-
"""
tani_analiz.py — Sayaç modüllerinin PULSE GENİŞLİĞİ ve VURUŞ ARALIĞI dağılımı.
SALT OKUNUR.

Neden var (kullanıcı 2026-08-21): iki saha şikâyeti var ve ikisi de eşik
değerlerine bağlı —
  1. Eksantrik presler hızlı çevrimde röleyi anlık çekip bırakıyor, firmware
     bunu saymıyor  → gerçek pulse ne kadar KISA, onu bilmemiz gerek.
  2. Broş makinesi aşağı inerken de yukarı çıkarken de sayıyor (çift sayım)
     → iniş↔çıkış arasındaki süre ne kadar, onu bilmemiz gerek.

Eşiği tahminle koymak sessiz yanlış sayım demektir. Firmware zaten her sayımda
pulse genişliğini (low_ms) ve önceki sayımdan farkı (gap_ms) tanı kanalıyla
gönderiyor; bu betik o veriyi okunur hâle getirir.

KULLANIM (sunucuda, proje klasöründe):
    python pilot\\tani_analiz.py                 → son 7 gün, tüm makineler
    python pilot\\tani_analiz.py "Pres 1"        → tek makine, ayrıntılı
    python pilot\\tani_analiz.py Bros 14         → tek makine, son 14 gün

OKUMA NOTU:
  low_ms  = rölenin KAPALI kaldığı süre (pulse genişliği). Firmware bir vuruşu
            saymak için bunun eşiği aşmasını bekler — eşikten kısa pulse HİÇ
            sayılmaz ve tanı kanalında da GÖRÜNMEZ (sayılmadığı için olay yok).
            Bu yüzden 'PARAZIT' sayacı önemli: ham HIGH→LOW kenarlarını sayar.
            PARAZIT >> SAYILDI ise sayılmayan gerçek vuruşlar olabilir.
  gap_ms  = bir önceki SAYILAN vuruştan bu yana geçen süre. Çift sayımda
            (broş iniş/çıkış) dağılım İKİ TEPELİ çıkar: kısa gap = aynı
            çevrimin ikinci sinyali, uzun gap = bir sonraki gerçek çevrim.
"""
import os
import sqlite3
import sys
from datetime import datetime, timedelta

PILOT_DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pilot.db')


def _yuzdelik(degerler, p):
    if not degerler:
        return 0
    k = (len(degerler) - 1) * p / 100.0
    alt = int(k)
    ust = min(alt + 1, len(degerler) - 1)
    return degerler[alt] + (degerler[ust] - degerler[alt]) * (k - alt)


def _histogram(degerler, kova_sinirlari, etiketler):
    """Basit metin histogramı — dağılımın İKİ TEPELİ olup olmadığı görünsün."""
    sayac = [0] * len(etiketler)
    for d in degerler:
        for i, sinir in enumerate(kova_sinirlari):
            if d < sinir:
                sayac[i] += 1
                break
        else:
            sayac[-1] += 1
    en_cok = max(sayac) or 1
    for etiket, n in zip(etiketler, sayac):
        if n == 0:
            continue
        cubuk = '█' * max(1, int(n * 40 / en_cok))
        print(f'    {etiket:>12}  {n:>6}  {cubuk}')


def _baglan():
    if not os.path.exists(PILOT_DB):
        print(f'[HATA] pilot.db bulunamadi: {PILOT_DB}')
        sys.exit(1)
    c = sqlite3.connect(PILOT_DB)
    c.row_factory = sqlite3.Row
    return c


def ozet(gun=7):
    c = _baglan()
    bas = (datetime.now() - timedelta(days=gun)).strftime('%Y-%m-%d %H:%M:%S')
    print('=' * 76)
    print(f'SAYAÇ TANI ÖZETİ · son {gun} gün')
    print('=' * 76)
    print(f'{"MAKİNE":<18}{"SAYILDI":>8}{"PARAZİT":>9}{"ERKEN":>7}   '
          f'{"pulse ms (medyan/min)":<22}{"gap ms medyan":>13}')
    print('-' * 76)
    for r in c.execute(
            "SELECT robot_no, "
            "  SUM(CASE WHEN tip='SAYILDI' THEN 1 ELSE 0 END) sayildi, "
            "  SUM(CASE WHEN tip='PARAZIT' THEN 1 ELSE 0 END) parazit, "
            "  SUM(CASE WHEN tip='ERKEN'   THEN 1 ELSE 0 END) erken "
            "FROM tani_olaylari WHERE ts >= ? GROUP BY robot_no ORDER BY sayildi DESC",
            (bas,)):
        lows = sorted(x[0] for x in c.execute(
            "SELECT low_ms FROM tani_olaylari WHERE robot_no=? AND tip='SAYILDI' "
            "AND ts >= ? AND low_ms > 0", (r['robot_no'], bas)))
        gaps = sorted(x[0] for x in c.execute(
            "SELECT gap_ms FROM tani_olaylari WHERE robot_no=? AND tip='SAYILDI' "
            "AND ts >= ? AND gap_ms > 0 AND gap_ms < 600000", (r['robot_no'], bas)))
        pulse = (f"{_yuzdelik(lows, 50):.0f} / {lows[0]:.0f}" if lows else '—')
        gap = (f"{_yuzdelik(gaps, 50):.0f}" if gaps else '—')
        # PARAZIT >> SAYILDI: sayilmayan gercek vurus olabilir (ya da gurultu)
        isaret = ' ⚠' if (r['parazit'] or 0) > (r['sayildi'] or 0) * 3 and (r['sayildi'] or 0) > 0 else ''
        print(f"{r['robot_no']:<18}{r['sayildi'] or 0:>8}{r['parazit'] or 0:>9}"
              f"{r['erken'] or 0:>7}   {pulse:<22}{gap:>13}{isaret}")
    print('-' * 76)
    print('  ⚠ = ham kenar sayısı sayılandan çok fazla — sayılmayan vuruş olabilir')
    print('  Ayrıntı: python pilot\\tani_analiz.py "Pres 1"')
    c.close()


def detay(makine, gun=7):
    c = _baglan()
    bas = (datetime.now() - timedelta(days=gun)).strftime('%Y-%m-%d %H:%M:%S')
    print('=' * 76)
    print(f'TANI DETAYI · {makine} · son {gun} gün')
    print('=' * 76)

    tipler = {r['tip']: r['n'] for r in c.execute(
        "SELECT tip, COUNT(*) n FROM tani_olaylari WHERE robot_no=? AND ts >= ? GROUP BY tip",
        (makine, bas))}
    if not tipler:
        print('  Bu aralıkta tanı olayı yok. Makine adı doğru mu? (ör. "Pres 1", "Bros")')
        c.close()
        return
    # Kümülatif sayaçlar — KÖR PENCERE kaybı burada görünür (firmware v2.8+).
    # ISR kesme HTTP blokesi sırasında da çalışır; loop çalışmaz. İkisinin farkı
    # doğrudan "loop'un göremediği vuruş sayısı"dır.
    kk = c.execute("SELECT tani_parazit, isr_count_ist1, firmware_ver "
                   "FROM cihaz_kayitlari WHERE robot_no=?", (makine,)).fetchone()
    if kk and (kk['isr_count_ist1'] or 0) > 0:
        isr = kk['isr_count_ist1']
        loop_kenar = kk['tani_parazit'] or 0
        print(f"\n[0] KÖR PENCERE ÖLÇÜMÜ  (firmware {kk['firmware_ver']})")
        print(f"    ISR kenar (blokede de sayar) : {isr}")
        print(f"    Loop kenar (tani_parazit)    : {loop_kenar}")
        kacan = isr - loop_kenar
        if kacan > 0:
            print(f"    → loop {kacan} kenarı GÖRMEDİ "
                  f"(%{kacan * 100 / max(isr, 1):.1f}) — HTTP blokesinde kaçmış")
        else:
            print('    → loop hiç kenar kaçırmamış')

    print('\n[1] OLAY SAYILARI')
    for t, n in sorted(tipler.items(), key=lambda x: -x[1]):
        aciklama = {'SAYILDI': 'geçerli vuruş olarak sayıldı',
                    'PARAZIT': 'ham HIGH→LOW kenar (sayılmış olanlar dahil)',
                    'ERKEN':   'MIN_PULSE_GAP dolmadan geldi → SAYILMADI'}.get(t, '')
        print(f'    {t:<10} {n:>7}   {aciklama}')

    lows = sorted(x[0] for x in c.execute(
        "SELECT low_ms FROM tani_olaylari WHERE robot_no=? AND tip='SAYILDI' "
        "AND ts >= ? AND low_ms > 0", (makine, bas)))
    if lows:
        print('\n[2] PULSE GENİŞLİĞİ (low_ms) — rölenin kapalı kaldığı süre')
        print(f'    min {lows[0]:.0f} · %5 {_yuzdelik(lows,5):.0f} · medyan '
              f'{_yuzdelik(lows,50):.0f} · %95 {_yuzdelik(lows,95):.0f} · max {lows[-1]:.0f} ms')
        _histogram(lows, [50, 100, 200, 400, 800, 1600, 3200],
                   ['<50 ms', '50-100', '100-200', '200-400', '400-800',
                    '800-1600', '1600-3200', '>3200 ms'])
        print('    NOT: firmware eşiği ~75 ms. Bu dağılımın SOL UCU eşiğe yakınsa,')
        print('    eşiğin ALTINDA kalan vuruşlar hiç sayılmıyor demektir (burada görünmezler).')

    gaps = sorted(x[0] for x in c.execute(
        "SELECT gap_ms FROM tani_olaylari WHERE robot_no=? AND tip='SAYILDI' "
        "AND ts >= ? AND gap_ms > 0 AND gap_ms < 300000", (makine, bas)))
    if gaps:
        print('\n[3] VURUŞ ARALIĞI (gap_ms) — önceki sayımdan bu yana')
        print(f'    min {gaps[0]:.0f} · %5 {_yuzdelik(gaps,5):.0f} · medyan '
              f'{_yuzdelik(gaps,50):.0f} · %95 {_yuzdelik(gaps,95):.0f} · max {gaps[-1]:.0f} ms')
        _histogram(gaps, [600, 1000, 2000, 4000, 8000, 15000, 30000],
                   ['<600 ms', '600-1000', '1-2 sn', '2-4 sn', '4-8 sn',
                    '8-15 sn', '15-30 sn', '>30 sn'])
        print('    ÇİFT SAYIM ARANIYORSA: dağılım İKİ TEPELİ mi? Kısa tepe = aynı')
        print('    çevrimin ikinci sinyali (broş iniş/çıkış), uzun tepe = gerçek çevrim.')

    print('\n[4] SON 15 SAYIM (zaman sırasıyla)')
    for r in c.execute(
            "SELECT ts, low_ms, gap_ms FROM tani_olaylari WHERE robot_no=? AND tip='SAYILDI' "
            "AND ts >= ? ORDER BY ts DESC LIMIT 15", (makine, bas)):
        print(f"    {r['ts']}   pulse {r['low_ms'] or 0:>6.0f} ms   gap {r['gap_ms'] or 0:>8.0f} ms")
    print()
    c.close()


if __name__ == '__main__':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    args = [a for a in sys.argv[1:]]
    if args and not args[0].isdigit():
        detay(args[0], int(args[1]) if len(args) > 1 else 7)
    else:
        ozet(int(args[0]) if args else 7)
