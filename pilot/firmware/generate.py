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
    # Eksantrik pres — iki-el AND butonu (bolum='pres'). Bkz pres.ino.tpl
    'pres': [
        ('PRES-P1', 'Pres 1'),
        ('PRES-P2', 'Pres 2'),
        ('PRES-P3', 'Pres 3'),
        ('PRES-P4', 'Pres 4'),
        ('PRES-P5', 'Pres 5'),
    ],
    # Plastik enjeksiyon — TK1, metal mantigi (bolum='plastik'). Bkz plastik.ino.tpl
    'plastik': [
        ('PLASTIK-320T', '320T'),
        ('PLASTIK-407T', '407T'),
    ],
}

TEMPLATES = {
    'kaynak':  'kaynak.ino.tpl',
    'montaj':  'montaj.ino.tpl',
    'metal':   'metal.ino.tpl',
    'abkant':  'abkant.ino.tpl',
    'pres':    'pres.ino.tpl',
    'plastik': 'plastik.ino.tpl',
}


def klasor_adi(robot_no):
    """Klasor + .ino dosya adi turet. ABB1 -> cofle_sayac_abb1"""
    sade = robot_no.lower().replace(' ', '_')
    return f'cofle_sayac_{sade}'


def template_uygula(template_metni, cihaz_id, robot_no):
    """Template placeholder'lari doldur."""
    fw_suffix = robot_no.lower().replace(' ', '_')
    return (template_metni
            .replace('__CIHAZ_ID__', cihaz_id)
            .replace('__ROBOT_NO__', robot_no)
            .replace('__FW_SUFFIX__', fw_suffix))


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
        for cihaz_id, robot_no in cihazlar:
            klasor = klasor_adi(robot_no)
            klasor_yolu = os.path.join(SCRIPT_DIR, klasor)
            os.makedirs(klasor_yolu, exist_ok=True)

            ino_yolu = os.path.join(klasor_yolu, klasor + '.ino')
            icerik = template_uygula(tpl, cihaz_id, robot_no)
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
