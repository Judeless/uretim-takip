# -*- coding: utf-8 -*-
"""
Kaynakhane.xlsx, Montaj.xlsx ve Metal Enjeksiyon.xlsx dosyalarından
referansları ve operatörleri uretim.db veritabanına aktarır.
"""
import openpyxl
import sqlite3
import os
import sys

# Excel dosya yolları
EXCEL_DOSYALARI = {
    'kaynak': {
        'yol': r'C:\Users\selcu\OneDrive\Masaüstü\Kaynakhane.xlsx',
        'ref_sayfa_anahtar': ['kaynak', 'süre', 'sure'],
        'op_sayfa_anahtar': ['operat'],
    },
    'montaj': {
        'yol': r'C:\Users\selcu\OneDrive\Masaüstü\Montaj.xlsx',
        'ref_sayfa_anahtar': ['montaj', 'süre', 'sure'],
        'op_sayfa_anahtar': ['operat'],
    },
    'metal': {
        'yol': r'C:\Users\selcu\OneDrive\Masaüstü\Metal Enjeksiyon.xlsx',
        'ref_sayfa_anahtar': ['bask', 'süre', 'sure', 'enjeksiyon'],
        'op_sayfa_anahtar': ['operat'],
    },
}

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uretim.db')


def _sayfa_bul(wb, anahtar_kelimeler):
    """Sayfa isimlerinde anahtar kelimeleri arayarak uygun sayfayı bulur."""
    for adi in wb.sheetnames:
        for anahtar in anahtar_kelimeler:
            if anahtar in adi.lower():
                return wb[adi]
    return None


def _dosya_import(conn, bolum, dosya_bilgi):
    """Tek bir Excel dosyasından referans ve operatörleri import eder."""
    yol = dosya_bilgi['yol']
    
    if not os.path.exists(yol):
        print(f"  UYARI: Excel dosyası bulunamadı: {yol}")
        return {
            'referanslar_eklenen': 0,
            'referanslar_guncellenen': 0,
            'operatorler_eklenen': 0,
            'hata': f'Dosya bulunamadı: {yol}'
        }

    wb = openpyxl.load_workbook(yol, data_only=True)
    sayfa_isimleri = wb.sheetnames
    print(f"  [{bolum.upper()}] Excel sayfaları: {sayfa_isimleri}")

    c = conn.cursor()

    # ── Referans sayfasını bul ──
    ref_sayfa = _sayfa_bul(wb, dosya_bilgi['ref_sayfa_anahtar'])
    if ref_sayfa is None:
        ref_sayfa = wb[sayfa_isimleri[0]]
        print(f"  Referans sayfası bulunamadı, ilk sayfa kullanılıyor: '{sayfa_isimleri[0]}'")
    else:
        print(f"  Referans sayfası: '{ref_sayfa.title}'")

    # ── Operatör sayfasını bul ──
    op_sayfa = _sayfa_bul(wb, dosya_bilgi['op_sayfa_anahtar'])
    if op_sayfa:
        print(f"  Operatör sayfası: '{op_sayfa.title}'")

    # Referansları içe aktar
    ref_sayisi = 0
    ref_guncellenen = 0
    for i, row in enumerate(ref_sayfa.iter_rows(values_only=True)):
        if i == 0:
            continue  # Başlık satırını atla
        if not row or row[0] is None:
            continue
        kod = str(row[0]).strip()
        if not kod or len(kod) < 2:
            continue
        try:
            cycle = float(row[1]) if (len(row) > 1 and row[1] is not None) else 0.0
        except ValueError:
            cycle = 0.0

        # Var olanı güncelle, yoksa ekle (Case-insensitive & Space-insensitive)
        mevcut = c.execute(
            "SELECT id, referans_kodu FROM referans_listesi WHERE UPPER(REPLACE(referans_kodu, ' ', '')) = UPPER(REPLACE(?, ' ', ''))", 
            (kod,)
        ).fetchone()
        
        if mevcut:
            # Eşleşme bulundu, süreyi, kodu ve bölümü güncelle
            c.execute('UPDATE referans_listesi SET hedef_cycle_time_sn = ?, referans_kodu = ?, bolum = ? WHERE id = ?', (cycle, kod, bolum, mevcut[0]))
            if cycle > 0:
                c.execute("UPDATE uretim_kayitlari SET cycle_time_sn = ? WHERE UPPER(REPLACE(referans_kodu, ' ', '')) = UPPER(REPLACE(?, ' ', ''))", (cycle, kod))
            ref_guncellenen += 1
        else:
            c.execute('INSERT INTO referans_listesi (referans_kodu, hedef_cycle_time_sn, bolum) VALUES (?, ?, ?)', (kod, cycle, bolum))
            if cycle > 0:
                c.execute("UPDATE uretim_kayitlari SET cycle_time_sn = ? WHERE UPPER(REPLACE(referans_kodu, ' ', '')) = UPPER(REPLACE(?, ' ', ''))", (cycle, kod))
            ref_sayisi += 1

    print(f"  Referanslar: {ref_sayisi} eklendi, {ref_guncellenen} güncellendi")

    # Operatörleri içe aktar
    op_sayisi = 0
    if op_sayfa:
        for i, row in enumerate(op_sayfa.iter_rows(values_only=True)):
            if i == 0:
                continue  # Başlık satırını atla
            if not row or row[1] is None:
                continue
            ad = str(row[1]).strip()
            if not ad:
                continue
            try:
                # Aynı isim ve bölüm kombinasyonu yoksa ekle
                mevcut_op = c.execute(
                    "SELECT id FROM operatorler WHERE UPPER(ad) = UPPER(?) AND bolum = ?",
                    (ad, bolum)
                ).fetchone()
                if not mevcut_op:
                    # Aynı isim farklı bölümde olabilir, bölüm ile birlikte ekle
                    # UNIQUE kısıtını aşmak için kontrol et
                    ayni_isim = c.execute("SELECT id, bolum FROM operatorler WHERE UPPER(ad) = UPPER(?)", (ad,)).fetchone()
                    if ayni_isim:
                        # Aynı isim var ama farklı bölümde, güncelleme yapmıyoruz
                        # (bir operatör birden fazla bölümde çalışabilir - edge case)
                        pass
                    else:
                        c.execute('INSERT INTO operatorler (ad, bolum) VALUES (?, ?)', (ad, bolum))
                    op_sayisi += 1
            except Exception as e:
                print(f"  Operatör eklenemedi ({ad}): {e}")

    print(f"  Operatörler: {op_sayisi} eklendi")

    return {
        'referanslar_eklenen': ref_sayisi,
        'referanslar_guncellenen': ref_guncellenen,
        'operatorler_eklenen': op_sayisi
    }


def import_data(bolum=None):
    """Excel dosyalarından verileri import eder.

    bolum=None  → tüm bölümler (kaynak + montaj + metal)
    bolum='kaynak' / 'montaj' / 'metal' → sadece ilgili dosya
    """
    if bolum and bolum not in EXCEL_DOSYALARI:
        return {'basarili': False, 'hata': f"Geçersiz bölüm: {bolum}"}

    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    # Operatörler tablosu oluştur (yoksa)
    c.execute('''
        CREATE TABLE IF NOT EXISTS operatorler (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ad TEXT UNIQUE NOT NULL,
            bolum TEXT DEFAULT 'kaynak'
        )
    ''')

    # bolum kolonu yoksa ekle (mevcut DB için)
    try:
        c.execute("ALTER TABLE operatorler ADD COLUMN bolum TEXT DEFAULT 'kaynak'")
    except Exception:
        pass

    # bolum kolonu referans_listesi'ne yoksa ekle
    try:
        c.execute("ALTER TABLE referans_listesi ADD COLUMN bolum TEXT DEFAULT 'kaynak'")
    except Exception:
        pass

    sonuclar = {}
    toplam = {'referanslar_eklenen': 0, 'referanslar_guncellenen': 0, 'operatorler_eklenen': 0}

    if bolum:
        islenen = [(bolum, EXCEL_DOSYALARI[bolum])]
    else:
        islenen = list(EXCEL_DOSYALARI.items())

    for b, dosya_bilgi in islenen:
        print(f"\n{'='*50}")
        print(f"  {b.upper()} import başlıyor...")
        print(f"{'='*50}")
        sonuc = _dosya_import(conn, b, dosya_bilgi)
        sonuclar[b] = sonuc
        toplam['referanslar_eklenen'] += sonuc['referanslar_eklenen']
        toplam['referanslar_guncellenen'] += sonuc['referanslar_guncellenen']
        toplam['operatorler_eklenen'] += sonuc['operatorler_eklenen']

    conn.commit()
    conn.close()

    return {
        'basarili': True,
        'referanslar_eklenen': toplam['referanslar_eklenen'],
        'referanslar_guncellenen': toplam['referanslar_guncellenen'],
        'operatorler_eklenen': toplam['operatorler_eklenen'],
        'detay': sonuclar
    }

if __name__ == '__main__':
    res = import_data()
    print("\n✅ Import tamamlandı:", res)
