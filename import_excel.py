# -*- coding: utf-8 -*-
"""
data/uretim_verileri.xlsx dosyasından referansları ve operatörleri
uretim.db veritabanına aktarır.

Excel yapısı (6 sayfa):
  - Kaynak Referans   | Kaynak Operator
  - Montaj Referans   | Montaj Operator
  - Metal Referans    | Metal Operator

Her referans sayfası: 1. sütun kod, 2. sütun cycle time (sn).
Her operator sayfası: 2. sütun operatör adı (1. sütun No).
"""
import openpyxl
import sqlite3
import os

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_YOL   = os.path.join(PROJECT_DIR, 'data', 'uretim_verileri.xlsx')
DB_PATH     = os.path.join(PROJECT_DIR, 'uretim.db')

BOLUM_SAYFA = {
    'kaynak': {'ref': 'Kaynak Referans', 'op': 'Kaynak Operator'},
    'montaj': {'ref': 'Montaj Referans', 'op': 'Montaj Operator'},
    'metal':  {'ref': 'Metal Referans',  'op': 'Metal Operator'},
}

# Bölüm bazlı duruş sebepleri sayfaları
BOLUM_DURUS_SAYFA = {
    'kaynak': 'Robotik Kaynak Duruş Listesi',
    'montaj': 'Montaj Duruş Listesi',
    'metal':  'Metal Enjeksiyon Duruş Listesi',
}

# Ek sayfalar (kaynak bölümüne özel — diğer bölümler için gerekmez)
ROBOT_PROGRAM_SAYFA = 'Robot Program Listesi'
FIKSTUR_RAF_SAYFA   = 'Fikstür Raf Listesi'


def durus_sebepleri_yukle(bolum):
    """data/uretim_verileri.xlsx içinden bölüme ait duruş sebeplerini okur.

    Sayfa formatı: No | Duruş Listesi | Planlı/Plansız
    Döner: [{'sebep': str, 'tip': 'planli'|'plansiz'}, ...]

    Excel veya sayfa yoksa boş liste döner (frontend fallback'e güveneceğinden değil,
    hatayı görsün diye).
    """
    if bolum not in BOLUM_DURUS_SAYFA:
        return []
    if not os.path.exists(EXCEL_YOL):
        return []

    sayfa_adi = BOLUM_DURUS_SAYFA[bolum]
    try:
        wb = openpyxl.load_workbook(EXCEL_YOL, data_only=True)
        if sayfa_adi not in wb.sheetnames:
            return []
        ws = wb[sayfa_adi]
        sonuc = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                continue  # Başlık
            if not row or row[1] is None:
                continue
            sebep = str(row[1]).strip()
            if not sebep:
                continue
            # Türkçe ı/i karakter farkı: 'Plansız' lower → 'plansız' (dotless ı),
            # ama 'siz' substring'i (regular i) bulunmaz. Bu yüzden açık string match.
            tip_raw = str(row[2] or '').strip().lower()
            if 'plansız' in tip_raw or 'plansiz' in tip_raw:
                tip = 'plansiz'
            elif 'planlı' in tip_raw or 'planli' in tip_raw:
                tip = 'planli'
            else:
                tip = 'plansiz'  # boş/bilinmeyen → güvenli yan: plansız
            sonuc.append({'sebep': sebep, 'tip': tip})
        return sonuc
    except Exception as e:
        print(f"[durus_sebepleri] Hata: {e}")
        return []


def _bolum_import(conn, wb, bolum):
    """Tek bölümün referans + operatör verisini import eder."""
    sayfalar = BOLUM_SAYFA[bolum]
    c = conn.cursor()

    # ── Referans sayfası ──
    if sayfalar['ref'] not in wb.sheetnames:
        return {
            'referanslar_eklenen': 0,
            'referanslar_guncellenen': 0,
            'referanslar_silinen': 0,
            'operatorler_eklenen': 0,
            'hata': f"'{sayfalar['ref']}' sayfası bulunamadı"
        }
    ref_sayfa = wb[sayfalar['ref']]
    print(f"  [{bolum.upper()}] Referans sayfası: '{ref_sayfa.title}'")

    ref_sayisi = 0
    ref_guncellenen = 0
    excel_kodlari_norm = set()

    for i, row in enumerate(ref_sayfa.iter_rows(values_only=True)):
        if i == 0:
            continue  # Başlık
        if not row or row[0] is None:
            continue
        kod = str(row[0]).strip()
        if not kod or len(kod) < 2:
            continue
        try:
            cycle = float(row[1]) if (len(row) > 1 and row[1] is not None) else 0.0
        except ValueError:
            cycle = 0.0

        excel_kodlari_norm.add(kod.upper().replace(' ', ''))

        mevcut = c.execute(
            "SELECT id, referans_kodu FROM referans_listesi "
            "WHERE UPPER(REPLACE(referans_kodu, ' ', '')) = UPPER(REPLACE(?, ' ', ''))",
            (kod,)
        ).fetchone()

        if mevcut:
            c.execute(
                'UPDATE referans_listesi SET hedef_cycle_time_sn = ?, referans_kodu = ?, bolum = ? WHERE id = ?',
                (cycle, kod, bolum, mevcut[0])
            )
            if cycle > 0:
                c.execute(
                    "UPDATE uretim_kayitlari SET cycle_time_sn = ? "
                    "WHERE UPPER(REPLACE(referans_kodu, ' ', '')) = UPPER(REPLACE(?, ' ', ''))",
                    (cycle, kod)
                )
            ref_guncellenen += 1
        else:
            c.execute(
                'INSERT INTO referans_listesi (referans_kodu, hedef_cycle_time_sn, bolum) VALUES (?, ?, ?)',
                (kod, cycle, bolum)
            )
            if cycle > 0:
                c.execute(
                    "UPDATE uretim_kayitlari SET cycle_time_sn = ? "
                    "WHERE UPPER(REPLACE(referans_kodu, ' ', '')) = UPPER(REPLACE(?, ' ', ''))",
                    (cycle, kod)
                )
            ref_sayisi += 1

    print(f"  Referanslar: {ref_sayisi} eklendi, {ref_guncellenen} güncellendi")

    # ── MIRROR SYNC: Excel'de olmayan referansları bu bölümden temizle ──
    ref_silinen = 0
    if excel_kodlari_norm:
        bolum_refs = c.execute(
            "SELECT id, referans_kodu FROM referans_listesi WHERE COALESCE(bolum, 'kaynak') = ?",
            (bolum,)
        ).fetchall()
        for ref_row in bolum_refs:
            ref_norm = str(ref_row[1] or '').upper().replace(' ', '')
            if ref_norm and ref_norm not in excel_kodlari_norm:
                c.execute('DELETE FROM referans_listesi WHERE id = ?', (ref_row[0],))
                ref_silinen += 1
        if ref_silinen:
            print(f"  Referanslar: {ref_silinen} adet (Excel'de olmayan) silindi")
    else:
        print("  UYARI: Excel'den hiçbir referans okunamadı, silme atlandı")

    # ── Operatör sayfası ──
    op_sayisi = 0
    if sayfalar['op'] in wb.sheetnames:
        op_sayfa = wb[sayfalar['op']]
        print(f"  [{bolum.upper()}] Operatör sayfası: '{op_sayfa.title}'")
        for i, row in enumerate(op_sayfa.iter_rows(values_only=True)):
            if i == 0:
                continue
            if not row or len(row) < 2 or row[1] is None:
                continue
            ad = str(row[1]).strip()
            if not ad:
                continue
            try:
                mevcut_op = c.execute(
                    "SELECT id FROM operatorler WHERE UPPER(ad) = UPPER(?) AND bolum = ?",
                    (ad, bolum)
                ).fetchone()
                if not mevcut_op:
                    ayni_isim = c.execute(
                        "SELECT id, bolum FROM operatorler WHERE UPPER(ad) = UPPER(?)",
                        (ad,)
                    ).fetchone()
                    if not ayni_isim:
                        c.execute('INSERT INTO operatorler (ad, bolum) VALUES (?, ?)', (ad, bolum))
                        op_sayisi += 1
            except Exception as e:
                print(f"  Operatör eklenemedi ({ad}): {e}")
        print(f"  Operatörler: {op_sayisi} eklendi")

    return {
        'referanslar_eklenen': ref_sayisi,
        'referanslar_guncellenen': ref_guncellenen,
        'referanslar_silinen': ref_silinen,
        'operatorler_eklenen': op_sayisi
    }


def import_data(bolum=None):
    """Excel'den verileri import eder.

    bolum=None  → tüm bölümler
    bolum='kaynak' / 'montaj' / 'metal' → sadece o bölüm
    """
    if bolum and bolum not in BOLUM_SAYFA:
        return {'basarili': False, 'hata': f"Geçersiz bölüm: {bolum}"}

    if not os.path.exists(EXCEL_YOL):
        return {'basarili': False, 'hata': f'Excel dosyası bulunamadı: {EXCEL_YOL}'}

    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    c.execute('''
        CREATE TABLE IF NOT EXISTS operatorler (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ad TEXT UNIQUE NOT NULL,
            bolum TEXT DEFAULT 'kaynak'
        )
    ''')
    try:
        c.execute("ALTER TABLE operatorler ADD COLUMN bolum TEXT DEFAULT 'kaynak'")
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE referans_listesi ADD COLUMN bolum TEXT DEFAULT 'kaynak'")
    except Exception:
        pass

    wb = openpyxl.load_workbook(EXCEL_YOL, data_only=True)

    sonuclar = {}
    toplam = {'referanslar_eklenen': 0, 'referanslar_guncellenen': 0,
              'referanslar_silinen': 0, 'operatorler_eklenen': 0}

    islenen = [bolum] if bolum else list(BOLUM_SAYFA.keys())

    for b in islenen:
        print(f"\n{'='*50}")
        print(f"  {b.upper()} import başlıyor...")
        print(f"{'='*50}")
        sonuc = _bolum_import(conn, wb, b)
        sonuclar[b] = sonuc
        for k in toplam:
            toplam[k] += sonuc.get(k, 0)

    conn.commit()
    conn.close()

    return {
        'basarili': True,
        **toplam,
        'detay': sonuclar
    }


def _program_listesi_import(conn, wb):
    """Robot Program Listesi sayfası → robot_programlari tablosu.
    Sayfa formatı (matrix):
       Satır 0: ROBOT | RAFNO | ABB-1 | ABB-1 | ABB-2 | ABB-2 | ... (her robot 2 kez)
       Satır 1: İSTASYON | <boş> | İST-1 | İST-2 | İST-1 | İST-2 | ...
       Satır 2+: <referans_kodu> | <raf_no> | √ | <boş> | √ | √ | ...
    Bu matrix'i düzleştirip her √ işareti için bir satır INSERT eder.
    """
    if wb is None or ROBOT_PROGRAM_SAYFA not in wb.sheetnames:
        return {'eklenen': 0, 'silinen': 0, 'hata': 'sayfa yok'}

    ws = wb[ROBOT_PROGRAM_SAYFA]
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 3:
        return {'eklenen': 0, 'silinen': 0, 'hata': 'yetersiz satır'}

    # Satır 0: robot adları (MERGED — birden çok kolonu kapsar)
    # Satır 1: istasyon. Kolon 0=referans, 1=raf
    robot_satiri = rows[0]
    istasyon_satiri = rows[1]
    # Merged cell mantığı: None olan kolonda son görülen robot devam eder
    robot_son = ''
    kolon_eslesme = []  # [(col_idx, robot_no, istasyon), ...]
    for j in range(2, max(len(robot_satiri), len(istasyon_satiri))):
        # Robot adı — yeni değer varsa al, yoksa son görüleni kullan (merged)
        r_raw = robot_satiri[j] if j < len(robot_satiri) and robot_satiri[j] is not None else None
        if r_raw is not None:
            robot_son = str(r_raw).strip()
        if not robot_son:
            continue
        # İstasyon
        i_raw = istasyon_satiri[j] if j < len(istasyon_satiri) and istasyon_satiri[j] is not None else None
        if i_raw is None:
            continue
        i_str = str(i_raw).strip()
        robot_no = robot_son.replace('-', '').replace(' ', '')  # ABB-1 → ABB1
        # İstasyon: "İST-1" → 1
        ist = 0
        if '1' in i_str: ist = 1
        elif '2' in i_str: ist = 2
        elif '3' in i_str: ist = 3
        if ist > 0:
            kolon_eslesme.append((j, robot_no, ist))

    c = conn.cursor()
    # Mevcut programları temizle (Excel master)
    c.execute('DELETE FROM robot_programlari')

    eklenen = 0
    for r in rows[2:]:
        if not r or r[0] is None: continue
        ref = str(r[0]).strip()
        if not ref or len(ref) < 2: continue
        for col_idx, robot_no, ist in kolon_eslesme:
            if col_idx >= len(r): continue
            val = str(r[col_idx] or '').strip()
            if val and val != '':  # √ veya başka bir işaret varsa
                c.execute(
                    'INSERT INTO robot_programlari (robot_no, istasyon, referans_kodu, guncelleyen) VALUES (?, ?, ?, ?)',
                    (robot_no, ist, ref, 'Excel İçe Aktar')
                )
                eklenen += 1

    print(f"  Robot Program: {eklenen} satır eklendi")
    return {'eklenen': eklenen}


def _fikstur_raf_import(conn, wb):
    """Fikstür Raf Listesi sayfası → fikstur_raf tablosu.
    Sayfa formatı: 3 raf yan yana (A, B, C). Her raf 2 kolon: kod | raf_no.
    Aralarda boş kolon olabilir.
    """
    if wb is None or FIKSTUR_RAF_SAYFA not in wb.sheetnames:
        return {'eklenen': 0, 'hata': 'sayfa yok'}

    ws = wb[FIKSTUR_RAF_SAYFA]
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        return {'eklenen': 0, 'hata': 'yetersiz satır'}

    c = conn.cursor()
    c.execute('DELETE FROM fikstur_raf')

    eklenen = 0
    # Her satırdaki tüm (kod, raf_no) çiftlerini topla
    for r in rows[1:]:  # Başlık satırını atla
        if not r: continue
        # Her 3 kolonda bir grup: (kod_col, raf_col, boş)
        i = 0
        while i < len(r):
            kod = str(r[i] or '').strip() if i < len(r) else ''
            raf = str(r[i+1] or '').strip() if i+1 < len(r) else ''
            if kod and raf:
                c.execute(
                    'INSERT INTO fikstur_raf (referans_kodu, raf_no) VALUES (?, ?)',
                    (kod, raf)
                )
                eklenen += 1
            i += 3  # Sonraki grup

    print(f"  Fikstür Raf: {eklenen} satır eklendi")
    return {'eklenen': eklenen}


def import_tum(yedek_al=False):
    """Excel'deki TÜM sayfaları okuyup DB'yi günceller:
       - Tüm bölüm referansları (cycle time)
       - Tüm bölüm operatörleri
       - Robot programları (kaynak için matrix)
       - Fikstür raf listesi
    Duruş sebepleri her API isteğinde okunduğu için import gerekmez.
    """
    if not os.path.exists(EXCEL_YOL):
        return {'basarili': False, 'hata': f'Excel bulunamadı: {EXCEL_YOL}'}

    # Önce normal referans+operator (mevcut)
    sonuc = import_data()

    # Sonra ek sayfalar
    conn = sqlite3.connect(DB_PATH)
    try:
        wb = openpyxl.load_workbook(EXCEL_YOL, data_only=True)

        # Robot Program
        try:
            prog_sonuc = _program_listesi_import(conn, wb)
            sonuc['program_eklenen'] = prog_sonuc.get('eklenen', 0)
        except Exception as e:
            print(f"  Robot Program HATA: {e}")
            sonuc['program_eklenen'] = 0

        # Fikstür Raf
        try:
            fik_sonuc = _fikstur_raf_import(conn, wb)
            sonuc['fikstur_eklenen'] = fik_sonuc.get('eklenen', 0)
        except Exception as e:
            print(f"  Fikstür HATA: {e}")
            sonuc['fikstur_eklenen'] = 0

        conn.commit()
    finally:
        conn.close()

    return sonuc


def export_referans_cycle_times(bolum=None):
    """DB'deki cycle_time'ları Excel'in <Bolum> Referans sayfa(lar)ına yazar.
    Diğer veriler (operatör, duruş, program, fikstür) korunur.
    """
    if not os.path.exists(EXCEL_YOL):
        return {'basarili': False, 'hata': f'Excel bulunamadı: {EXCEL_YOL}'}

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    bolum_listesi = [bolum] if bolum else list(BOLUM_SAYFA.keys())
    toplam_yazilan = 0
    wb = openpyxl.load_workbook(EXCEL_YOL)

    for b in bolum_listesi:
        sayfa_adi = BOLUM_SAYFA[b]['ref']
        if sayfa_adi not in wb.sheetnames:
            continue
        ws = wb[sayfa_adi]
        # Mevcut Excel satırlarını oku, kod → satır map'i
        kod_satir = {}
        for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if i == 1: continue  # Başlık
            if not row or row[0] is None: continue
            kod = str(row[0]).strip()
            kod_satir[kod.upper().replace(' ', '')] = i

        # DB'deki bu bölüme ait referansları çek
        db_rows = conn.execute(
            "SELECT referans_kodu, hedef_cycle_time_sn FROM referans_listesi WHERE COALESCE(bolum,'kaynak')=? ORDER BY referans_kodu",
            (b,)
        ).fetchall()

        yazilan = 0
        for r in db_rows:
            kod = (r['referans_kodu'] or '').strip()
            norm = kod.upper().replace(' ', '')
            if norm in kod_satir:
                # Mevcut satırın 2. kolonunu güncelle
                ws.cell(row=kod_satir[norm], column=2, value=r['hedef_cycle_time_sn'] or 0)
                yazilan += 1
            else:
                # Yeni satır — sonuna ekle
                yeni_r = ws.max_row + 1
                ws.cell(row=yeni_r, column=1, value=kod)
                ws.cell(row=yeni_r, column=2, value=r['hedef_cycle_time_sn'] or 0)
                yazilan += 1
        toplam_yazilan += yazilan
        print(f"  {b}: {yazilan} satır Excel'e yazıldı")

    conn.close()
    wb.save(EXCEL_YOL)
    return {'basarili': True, 'yazilan': toplam_yazilan, 'dosya': EXCEL_YOL}


if __name__ == '__main__':
    res = import_tum()
    print("\n✅ Import tamamlandı:", res)
