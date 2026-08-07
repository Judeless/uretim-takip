"""
Cofle Pilot firmware generator
================================

Tum saha cihazlari icin Arduino IDE ile yuklenebilir .ino dosyalarini
3 ayri template'ten (kaynak/montaj/metal) uretir.

Kullanim:
    cd pilot/firmware
    python generate.py

Cikti:
    pilot/firmware/
        cofle_sayac_abb1/cofle_sayac_abb1.ino
        cofle_sayac_abb2/cofle_sayac_abb2.ino
        ...
        cofle_sayac_m1/cofle_sayac_m1.ino
        cofle_sayac_m2/cofle_sayac_m2.ino
        ...
        cofle_sayac_300t/cofle_sayac_300t.ino
        cofle_sayac_400t/cofle_sayac_400t.ino
        cofle_sayac_550t/cofle_sayac_550t.ino

Yeni cihaz eklemek icin:
    Asagidaki DEVICES sozlugune ekle, scripti tekrar calistir.
"""

import os
import shutil
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TPL_DIR = os.path.join(SCRIPT_DIR, '_templates')

# ============================================================
# Cihaz listesi — bolume gore. (cihaz_id, robot_no) ciftleri.
# ============================================================
DEVICES = {
    'kaynak': [
        ('ABB1-IO', 'ABB1'),
        ('ABB2-IO', 'ABB2'),
        ('ABB3-IO', 'ABB3'),
        ('ABB4-IO', 'ABB4'),
        ('ABB5-IO', 'ABB5'),
        ('ABB6-IO', 'ABB6'),
        ('ABB7-IO', 'ABB7'),
        ('ABB8-IO', 'ABB8'),
        ('ABB9-IO', 'ABB9'),
    ],
    'montaj': [
        ('MONTAJ-M1',  'M1'),
        ('MONTAJ-M2',  'M2'),
        ('MONTAJ-M3',  'M3'),
        ('MONTAJ-M4',  'M4'),
        ('MONTAJ-M5',  'M5'),
        ('MONTAJ-M6',  'M6'),
        ('MONTAJ-M7',  'M7'),
        ('MONTAJ-M8',  'M8'),
        ('MONTAJ-M9',  'M9'),
        ('MONTAJ-M10', 'M10'),
        ('MONTAJ-M11', 'M11'),
        ('MONTAJ-M12', 'M12'),
        ('MONTAJ-YF1', 'YF1'),   # TK1 yan tesis montaj deneme modülü (sahada robot_no=YF1 ile flash'li)
        # TK1 (yan tesis) montaj masalari — 2026-07-31, 7 adet. TK2'deki M1..M12 modelinin
        # aynisi: MASA = hat (operator mobilde masasini secer, sayac dogrudan eslesir).
        # YF1 sahada KALIR (LF-LFP esleme sürüyor) — bu 7 cihaz onun yanina eklenir.
        # robot_no global benzersiz olmali: 'TK1-M1' (TK2'nin 'M1'i ile karismasin).
        # BUZZER_AKTIF='false': TK1 masalarina alinan 3 pinli hazir modul (VCC/I-O/GND)
        # AKTIF-LOW cikti — sahada denendi, HIGH surulunce SUREKLI otuyordu (2026-07-31).
        # Modul 3V3'ten beslenir (5V'ta 3.3V HIGH girisi kapatamiyor) -> ses kisik;
        # telafi: CIFT bip x 300ms (kullanici istegi). TK2 M1..M12 + YF1 tek bip
        # 200ms aktif-HIGH kalir; bu yuzden ayarlarin HEPSI CIHAZ BAZLI.
        ('MONTAJ-TK1-M1', 'TK1-M1', {'BUZZER_AKTIF': 'false', 'BUZZER_BEEP_ADET': '2', 'BUZZER_BEEP_MS': '300'}),
        ('MONTAJ-TK1-M2', 'TK1-M2', {'BUZZER_AKTIF': 'false', 'BUZZER_BEEP_ADET': '2', 'BUZZER_BEEP_MS': '300'}),
        ('MONTAJ-TK1-M3', 'TK1-M3', {'BUZZER_AKTIF': 'false', 'BUZZER_BEEP_ADET': '2', 'BUZZER_BEEP_MS': '300'}),
        ('MONTAJ-TK1-M4', 'TK1-M4', {'BUZZER_AKTIF': 'false', 'BUZZER_BEEP_ADET': '2', 'BUZZER_BEEP_MS': '300'}),
        ('MONTAJ-TK1-M5', 'TK1-M5', {'BUZZER_AKTIF': 'false', 'BUZZER_BEEP_ADET': '2', 'BUZZER_BEEP_MS': '300'}),
        ('MONTAJ-TK1-M6', 'TK1-M6', {'BUZZER_AKTIF': 'false', 'BUZZER_BEEP_ADET': '2', 'BUZZER_BEEP_MS': '300'}),
        ('MONTAJ-TK1-M7', 'TK1-M7', {'BUZZER_AKTIF': 'false', 'BUZZER_BEEP_ADET': '2', 'BUZZER_BEEP_MS': '300'}),
    ],
    'metal': [
        ('300T-IO', '300T'),
        ('400T-IO', '400T'),
        ('550T-IO', '550T'),
    ],
    # Abkant — pres bolumunun makineleri (bolum='pres', role sinyali). Bkz abkant.ino.tpl
    'abkant': [
        ('ABKANT-A1', 'Abkant 1'),
        ('ABKANT-A2', 'Abkant 2'),
        ('ABKANT-A3', 'Abkant 3'),
    ],
    # Pres makineleri — makinenin KENDI SAYACINA takilan role (bolum='pres').
    # 2026-07-31: eski iki-el buton (AND) modeli birakildi, abkant deseni kullaniliyor.
    # 5 eksantrik + 1 hidrolik pres + 1 bros. Bkz pres.ino.tpl
    # robot_no ASCII olmali: klasor/dosya adi + app.py PRES_MAKINELERI ile birebir.
    'pres': [
        ('PRES-P1', 'Pres 1'),
        ('PRES-P2', 'Pres 2'),
        ('PRES-P3', 'Pres 3'),
        ('PRES-P4', 'Pres 4'),
        ('PRES-P5', 'Pres 5'),
        ('PRES-HIDROLIK', 'Hidrolik Pres'),
        ('PRES-BROS',     'Bros'),
    ],
    # Plastik enjeksiyon — TK1, metal mantigi (bolum='plastik'). Bkz plastik.ino.tpl
    'plastik': [
        ('PLASTIK-320T', '320T'),
        ('PLASTIK-407T', '407T'),
    ],
    # Yapistirma — plastik enjeksiyon HATTININ makinesi (bolum='plastik', role KISA
    # pulse — plastik enj. uzun sinyalinin AKSINE). Abkant deseni. Bkz yapistirma.ino.tpl
    'yapistirma': [
        ('PLASTIK-YAPISTIRMA', 'Yapistirma'),
    ],
    # ── TK1 KAPAMA PRESLERI — COK KANALLI (2026-08-07) ─────────────────
    # 12 pres, ORTAK PANO: her panoda 3 pres var -> 3 presi TEK ESP32 okur
    # (GPIO25/26/27). 12 makine = 4 modul. robot_no LISTE verilir; her kanal
    # kendi makine adiyla sinyal gonderir, kanal sirasi = GPIO sirasi.
    #   klasor adi CIHAZ_ID'den turer (robot_no tek degil).
    #   HB_ROBOT_NO = heartbeat/saglik kayitlarinda gorunen modul etiketi.
    # Bolum 'tel' (kapama, tel uretiminin proses adimi). Bkz tel_kapama.ino.tpl
    'tel_kapama': [
        ('TEL-KAPAMA-1', ['Kapama 1',  'Kapama 2',  'Kapama 3'],
         {'HB_ROBOT_NO': 'Kapama 1-3'}),
        ('TEL-KAPAMA-2', ['Kapama 4',  'Kapama 5',  'Kapama 6'],
         {'HB_ROBOT_NO': 'Kapama 4-6'}),
        ('TEL-KAPAMA-3', ['Kapama 7',  'Kapama 8',  'Kapama 9'],
         {'HB_ROBOT_NO': 'Kapama 7-9'}),
        ('TEL-KAPAMA-4', ['Kapama 10', 'Kapama 11', 'Kapama 12'],
         {'HB_ROBOT_NO': 'Kapama 10-12'}),
    ],
}

TEMPLATES = {
    'kaynak':     'kaynak.ino.tpl',
    'montaj':     'montaj.ino.tpl',
    'metal':      'metal.ino.tpl',
    'abkant':     'abkant.ino.tpl',
    'pres':       'pres.ino.tpl',
    'plastik':    'plastik.ino.tpl',
    'yapistirma': 'yapistirma.ino.tpl',
    'tel_kapama': 'tel_kapama.ino.tpl',
}


def klasor_adi(ad):
    """Klasor + .ino dosya adi turet. ABB1 -> cofle_sayac_abb1

    Tire de alt cizgiye cevrilir ('TK1-M1' -> cofle_sayac_tk1_m1): Arduino sketch
    adinda tire sorun cikarabiliyor. robot_no'nun KENDISI degismez (sunucu ile
    birebir eslesmeli) — yalniz klasor/dosya adi sadelesir.

    Cok kanalli modullerde robot_no tek degildir; ad olarak CIHAZ_ID gecilir
    ('TEL-KAPAMA-1' -> cofle_sayac_tel_kapama_1)."""
    sade = str(ad).lower().replace(' ', '_').replace('-', '_')
    return f'cofle_sayac_{sade}'


def template_uygula(template_metni, cihaz_id, robot_no, ekstra=None):
    """Template placeholder'lari doldur.

    robot_no: tek makine icin str; COK KANALLI modulde makine adlari LISTESI
    (kanal sirasi = GPIO sirasi). Listede __ROBOT_NO_1/2/3__ doldurulur ve
    __ROBOT_NO__ ilk kanala baglanir (tek-kanalli template'lerle uyum).

    ekstra: cihaza OZEL degerler (opsiyonel 3. tuple elemani):
      BUZZER_AKTIF     : 'true' aktif-HIGH (TK2 kurulumu) | 'false' aktif-LOW (TK1 modulu)
      BUZZER_BEEP_ADET : pes pese bip sayisi   (TK2 '1', TK1 '2' — 3V3'te ses kisik)
      BUZZER_BEEP_MS   : tek bip suresi ms     (TK2 '200', TK1 '300' — %50 uzun)
      HB_ROBOT_NO      : cok kanalli modulun heartbeat/telemetri etiketi
    Placeholder'i olmayan template'lerde replace no-op'tur."""
    ek = ekstra or {}
    coklu = isinstance(robot_no, (list, tuple))
    kanallar = list(robot_no) if coklu else [robot_no]
    # FW surum soneki: cok kanalli modulde makine adi tek degil -> CIHAZ_ID kullan.
    # Tek kanalda tire KORUNUR ('tk1-m1'): sahadaki cihazlarin bildirdigi surum
    # etiketi degismesin (yalniz kozmetik ama gecmis telemetri ile karsilastirilir).
    fw_suffix = (cihaz_id.lower().replace(' ', '_').replace('-', '_') if coklu
                 else robot_no.lower().replace(' ', '_'))

    metin = (template_metni
             .replace('__CIHAZ_ID__', cihaz_id)
             .replace('__FW_SUFFIX__', fw_suffix)
             .replace('__HB_ROBOT_NO__', ek.get('HB_ROBOT_NO', kanallar[0]))
             .replace('__BUZZER_AKTIF__', ek.get('BUZZER_AKTIF', 'true'))
             .replace('__BUZZER_BEEP_ADET__', ek.get('BUZZER_BEEP_ADET', '1'))
             .replace('__BUZZER_BEEP_MS__', ek.get('BUZZER_BEEP_MS', '200')))
    for i, ad in enumerate(kanallar, start=1):
        metin = metin.replace(f'__ROBOT_NO_{i}__', ad)
    # __ROBOT_NO__ EN SONA: '__ROBOT_NO_1__' bunu icerir, once o doldurulmali
    return metin.replace('__ROBOT_NO__', kanallar[0])


def main():
    if not os.path.isdir(TPL_DIR):
        print(f'[HATA] Template klasoru bulunamadi: {TPL_DIR}', file=sys.stderr)
        sys.exit(1)

    # Eski auto-generated klasorleri tespit etmek icin marker
    AUTO_MARKER = 'OTOMATIK URETILDI: generate.py'

    # Tum cihaz klasorlerini yeniden uret
    toplam_uretildi = 0
    eski_temizlenen = 0

    # Once mevcut klasorleri tara, auto-generated olanlari sil
    for entry in os.listdir(SCRIPT_DIR):
        full = os.path.join(SCRIPT_DIR, entry)
        if not os.path.isdir(full) or entry.startswith('_'):
            continue
        if not entry.startswith('cofle_sayac_'):
            continue
        # Iceride .ino var mi ve auto-marker iceriyor mu kontrol et
        for fname in os.listdir(full):
            if fname.endswith('.ino'):
                try:
                    with open(os.path.join(full, fname), 'r', encoding='utf-8') as f:
                        ilk_500 = f.read(500)
                    if AUTO_MARKER in ilk_500:
                        shutil.rmtree(full)
                        eski_temizlenen += 1
                        print(f'  [sil]  {entry}/ (eski auto-generated)')
                        break
                except (OSError, UnicodeDecodeError):
                    pass

    print()

    # Eski jenerik .ino dosyasi varsa sil (cofle_sayac.ino kok seviyede)
    eski_kok_ino = os.path.join(SCRIPT_DIR, 'cofle_sayac.ino')
    if os.path.isfile(eski_kok_ino):
        try:
            with open(eski_kok_ino, 'r', encoding='utf-8') as f:
                if AUTO_MARKER not in f.read(500):
                    # Manuel — yine de yedek alip silebiliriz
                    pass
        except Exception:
            pass
        # Bu jenerik dosya artik gereksiz — silinsin
        os.remove(eski_kok_ino)
        print(f'  [sil]  cofle_sayac.ino (kok seviye eski generic)')
        print()

    # Yeniden uret
    for bolum, cihazlar in DEVICES.items():
        tpl_path = os.path.join(TPL_DIR, TEMPLATES[bolum])
        if not os.path.isfile(tpl_path):
            print(f'[UYARI] Template eksik, atlandi: {tpl_path}', file=sys.stderr)
            continue
        with open(tpl_path, 'r', encoding='utf-8') as f:
            tpl = f.read()

        print(f'[{bolum.upper()}] {len(cihazlar)} cihaz uretiliyor')
        for cihaz in cihazlar:
            # (cihaz_id, robot_no) veya (cihaz_id, robot_no, {ekstra}) — cihaza ozel
            # ayar gerekmeyen satirlar 2'li kalir (mevcut liste degismedi).
            cihaz_id, robot_no = cihaz[0], cihaz[1]
            ekstra = cihaz[2] if len(cihaz) > 2 else None
            # Cok kanalli modulde (robot_no LISTE) klasor adi CIHAZ_ID'den turer —
            # tek bir makine adi yok. Tek kanalli cihazlarda eski davranis aynen kalir.
            coklu = isinstance(robot_no, (list, tuple))
            klasor = klasor_adi(cihaz_id if coklu else robot_no)
            klasor_yolu = os.path.join(SCRIPT_DIR, klasor)
            os.makedirs(klasor_yolu, exist_ok=True)

            ino_yolu = os.path.join(klasor_yolu, klasor + '.ino')
            icerik = template_uygula(tpl, cihaz_id, robot_no, ekstra)
            with open(ino_yolu, 'w', encoding='utf-8') as f:
                f.write(icerik)
            print(f'  [ok]   {klasor}/{klasor}.ino  ({cihaz_id})')
            toplam_uretildi += 1
        print()

    print(f'OZET: {toplam_uretildi} firmware uretildi, {eski_temizlenen} eski klasor temizlendi')
    print()
    print('SONRAKI ADIM:')
    print('  1. Arduino IDE\'de istediginiz cihazin .ino dosyasini ac')
    print('  2. ESP32-WROOM-32U board sec, USB porta tak')
    print('  3. Yukleme (CTRL+U)')
    print('  4. Cihazi sahaya gotur, USB veya VIN ile 5V besleme')


if __name__ == '__main__':
    main()
