import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), 'uretim.db')


def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA synchronous = NORMAL")
    return conn


def init_db():
    conn = get_db()
    c = conn.cursor()

    # Vardiyalar tablosu
    c.execute('''
        CREATE TABLE IF NOT EXISTS vardiyalar (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tarih TEXT NOT NULL,
            vardiya_turu TEXT NOT NULL,
            robot_no TEXT NOT NULL,
            operator_adi TEXT NOT NULL,
            baslangic_saati TEXT NOT NULL,
            bitis_saati TEXT NOT NULL,
            toplam_sure_dk INTEGER NOT NULL,
            notlar TEXT DEFAULT '',
            durum TEXT DEFAULT 'aktif',
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    ''')

    # durum kolonu yoksa ekle (mevcut DB icin migration)
    try:
        c.execute("ALTER TABLE vardiyalar ADD COLUMN durum TEXT DEFAULT 'aktif'")
        c.execute("UPDATE vardiyalar SET durum='kapali' WHERE durum IS NULL")
    except Exception:
        pass  # Kolon zaten var

    # Uretim kayitlari tablosu
    c.execute('''
        CREATE TABLE IF NOT EXISTS uretim_kayitlari (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vardiya_id INTEGER NOT NULL,
            referans_kodu TEXT NOT NULL,
            ok_adet INTEGER NOT NULL DEFAULT 0,
            nok_adet INTEGER NOT NULL DEFAULT 0,
            tamir_adet INTEGER NOT NULL DEFAULT 0,
            hedef_adet INTEGER NOT NULL DEFAULT 0,
            cycle_time_sn REAL DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now', 'localtime')),
            FOREIGN KEY (vardiya_id) REFERENCES vardiyalar(id) ON DELETE CASCADE
        )
    ''')

    # Duruslar tablosu
    c.execute('''
        CREATE TABLE IF NOT EXISTS duruslar (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vardiya_id INTEGER NOT NULL,
            durus_sebebi TEXT NOT NULL,
            aciklama TEXT DEFAULT '',
            sure_dk INTEGER NOT NULL DEFAULT 0,
            baslangic_saati TEXT DEFAULT '',
            durus_tipi TEXT DEFAULT 'plansiz',
            created_at TEXT DEFAULT (datetime('now', 'localtime')),
            FOREIGN KEY (vardiya_id) REFERENCES vardiyalar(id) ON DELETE CASCADE
        )
    ''')

    # durus_tipi kolonu yoksa ekle (mevcut DB icin migration)
    try:
        c.execute("ALTER TABLE duruslar ADD COLUMN durus_tipi TEXT DEFAULT 'plansiz'")
    except Exception:
        pass  # Kolon zaten var

    # istasyon kolonu yoksa ekle (mevcut DB icin migration)
    try:
        c.execute("ALTER TABLE uretim_kayitlari ADD COLUMN istasyon INTEGER DEFAULT 0")
    except Exception:
        pass  # Kolon zaten var

    # launch_adet kolonu yoksa ekle (mevcut DB icin migration)
    try:
        c.execute("ALTER TABLE uretim_kayitlari ADD COLUMN launch_adet INTEGER DEFAULT 0")
    except Exception:
        pass  # Kolon zaten var

    # tamamlandi kolonu yoksa ekle (mevcut DB icin migration)
    try:
        c.execute("ALTER TABLE uretim_kayitlari ADD COLUMN tamamlandi INTEGER DEFAULT 0")
    except Exception:
        pass  # Kolon zaten var

    # bolum kolonu - vardiyalar tablosu (montaj/metal enjeksiyon desteği)
    try:
        c.execute("ALTER TABLE vardiyalar ADD COLUMN bolum TEXT DEFAULT 'kaynak'")
    except Exception:
        pass  # Kolon zaten var

    # bolum kolonu - referans_listesi tablosu
    try:
        c.execute("ALTER TABLE referans_listesi ADD COLUMN bolum TEXT DEFAULT 'kaynak'")
    except Exception:
        pass  # Kolon zaten var

    # bolum kolonu - operatorler tablosu
    try:
        c.execute("ALTER TABLE operatorler ADD COLUMN bolum TEXT DEFAULT 'kaynak'")
    except Exception:
        pass  # Kolon zaten var

    # Migration: 'metal_enjeksiyon' -> 'metal' (bölüm değerlerini standartlaştır)
    for tbl in ('vardiyalar', 'referans_listesi', 'operatorler'):
        try:
            c.execute(f"UPDATE {tbl} SET bolum='metal' WHERE bolum='metal_enjeksiyon'")
        except Exception:
            pass

    # tamir_adet kolonu yoksa ekle (mevcut DB icin migration)
    try:
        c.execute("ALTER TABLE uretim_kayitlari ADD COLUMN tamir_adet INTEGER DEFAULT 0")
    except Exception:
        pass  # Kolon zaten var


    # Referans listesi
    c.execute('''
        CREATE TABLE IF NOT EXISTS referans_listesi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            referans_kodu TEXT UNIQUE NOT NULL,
            aciklama TEXT DEFAULT '',
            hedef_cycle_time_sn REAL DEFAULT 0
        )
    ''')

    # Operatorler tablosu
    c.execute('''
        CREATE TABLE IF NOT EXISTS operatorler (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ad TEXT UNIQUE NOT NULL
        )
    ''')

    # Varsayilan referanslar
    referanslar = [
        ('REF-001', 'On Panel', 45),
        ('REF-002', 'Arka Panel', 38),
        ('REF-003', 'Yan Kapak', 52),
        ('REF-004', 'Taban Plakasi', 60),
        ('REF-005', 'Ust Kapak', 41),
    ]
    c.executemany(
        'INSERT OR IGNORE INTO referans_listesi (referans_kodu, aciklama, hedef_cycle_time_sn) VALUES (?,?,?)',
        referanslar
    )

    # Fikstür Adresleri Tablosu
    c.execute('''
        CREATE TABLE IF NOT EXISTS fikstur_adresleri (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            referans_kodu TEXT UNIQUE NOT NULL,
            raf_adresi TEXT DEFAULT '',
            notlar TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    ''')

    # Referans Robot Uyumu Tablosu
    c.execute('''
        CREATE TABLE IF NOT EXISTS referans_robot_uyumu (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            referans_kodu TEXT NOT NULL,
            robot_no TEXT NOT NULL,
            istasyon INTEGER NOT NULL,
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    ''')

    # Robot Programlari Tablosu (coklu kayit destekli - UNIQUE kaldirildi)
    c.execute('''
        CREATE TABLE IF NOT EXISTS robot_programlari (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            robot_no TEXT NOT NULL,
            istasyon INTEGER NOT NULL,
            referans_kodu TEXT DEFAULT '',
            guncelleyen TEXT DEFAULT '',
            guncelleme_tarihi TEXT DEFAULT (datetime('now', 'localtime'))
        )
    ''')

    # Eski UNIQUE kısıtlı tabloyu kontrol et ve gerekirse migrate et
    try:
        # Tablodaki indeksleri kontrol et
        c.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='robot_programlari'")
        row = c.fetchone()
        if row and 'UNIQUE' in (row[0] or ''):
            # Eski tabloyu yedekle
            c.execute("ALTER TABLE robot_programlari RENAME TO robot_programlari_eski")
            # Yeni tabloyu oluştur (UNIQUE yok)
            c.execute('''
                CREATE TABLE robot_programlari (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    robot_no TEXT NOT NULL,
                    istasyon INTEGER NOT NULL,
                    referans_kodu TEXT DEFAULT '',
                    guncelleyen TEXT DEFAULT '',
                    guncelleme_tarihi TEXT DEFAULT (datetime('now', 'localtime'))
                )
            ''')
            # Eski verilerden sadece dolu/anlamli olanlari tası
            c.execute('''
                INSERT INTO robot_programlari (robot_no, istasyon, referans_kodu, guncelleyen, guncelleme_tarihi)
                SELECT robot_no, istasyon, referans_kodu, guncelleyen, guncelleme_tarihi
                FROM robot_programlari_eski
                WHERE referans_kodu != ''
            ''')
    except Exception as e:
        pass  # Tablo zaten yeni formatta

    # Robot İş Atamaları Tablosu
    c.execute('''
        CREATE TABLE IF NOT EXISTS robot_is_atamalari (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            robot_no TEXT NOT NULL,
            istasyon INTEGER NOT NULL DEFAULT 0,
            referans_kodu TEXT NOT NULL,
            aciklama TEXT DEFAULT '',
            atayan TEXT DEFAULT '',
            atama_tarihi TEXT DEFAULT (datetime('now', 'localtime')),
            durum TEXT DEFAULT 'bekliyor'
        )
    ''')

    # Referans Takip (Launch / İş Emri) Tablosu
    c.execute('''
        CREATE TABLE IF NOT EXISTS referans_takip (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            referans_kodu TEXT NOT NULL,
            hedef_adet INTEGER NOT NULL DEFAULT 0,
            aciklama TEXT DEFAULT '',
            durum TEXT NOT NULL DEFAULT 'launch_alinacak',
            olusturan TEXT DEFAULT '',
            robot_no TEXT DEFAULT '',
            istasyon INTEGER DEFAULT 0,
            olusturma_tarihi TEXT DEFAULT (datetime('now', 'localtime')),
            guncelleme_tarihi TEXT DEFAULT (datetime('now', 'localtime'))
        )
    ''')
    # Migration: mevcut DB'ye robot_no ve istasyon kolonları
    try:
        c.execute("ALTER TABLE referans_takip ADD COLUMN robot_no TEXT DEFAULT ''")
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE referans_takip ADD COLUMN istasyon INTEGER DEFAULT 0")
    except Exception:
        pass

    # Andon Robot Ayarları Tablosu
    c.execute('''
        CREATE TABLE IF NOT EXISTS andon_robot_ayarlari (
            robot_no TEXT PRIMARY KEY,
            goster INTEGER NOT NULL DEFAULT 1,
            sira INTEGER NOT NULL DEFAULT 0
        )
    ''')
    # Genel Ayarlar Tablosu
    c.execute('''
        CREATE TABLE IF NOT EXISTS genel_ayarlar (
            anahtar TEXT PRIMARY KEY,
            deger TEXT
        )
    ''')
    # Varsayılan ayarları ekle
    c.execute("INSERT OR IGNORE INTO genel_ayarlar (anahtar, deger) VALUES ('andon_font_size', '0.57')")

    # Varsayılan robot satırlarını ekle (mevcut değilse)

    for i, rno in enumerate(['ABB1','ABB2','ABB3','ABB4','ABB5','ABB6','ABB7','ABB8','ABB9']):
        goster = 0 if rno == 'ABB9' else 1
        c.execute(
            'INSERT OR IGNORE INTO andon_robot_ayarlari (robot_no, goster, sira) VALUES (?, ?, ?)',
            (rno, goster, i)
        )

    conn.commit()
    conn.close()



if __name__ == '__main__':
    init_db()
    print("Veritabani hazir:", DB_PATH)
