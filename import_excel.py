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


if __name__ == '__main__':
    res = import_data()
    print("\n✅ Import tamamlandı:", res)
