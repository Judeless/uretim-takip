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

    # aciklama kolonu — operatör üretim kaydında not bırakabilir
    try:
        c.execute("ALTER TABLE uretim_kayitlari ADD COLUMN aciklama TEXT DEFAULT ''")
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

    # bolum kolonu - referans_takip tablosu (Iş Yönetimi her bölüm için ayrı çalışsın)
    try:
        c.execute("ALTER TABLE referans_takip ADD COLUMN bolum TEXT DEFAULT 'kaynak'")
    except Exception:
        pass  # Kolon zaten var

    # Legacy data: robot_no='MONTAJ' kayıtları montaj, robot_no='ME' kayıtları metal,
    # diğerleri (ABB*) kaynak. Sadece bolum alanı boşsa/NULL ise yansıt.
    try:
        c.execute("UPDATE referans_takip SET bolum='montaj' WHERE robot_no='MONTAJ' AND (bolum IS NULL OR bolum='' OR bolum='kaynak')")
        c.execute("UPDATE referans_takip SET bolum='metal' WHERE robot_no='ME' AND (bolum IS NULL OR bolum='' OR bolum='kaynak')")
    except Exception:
        pass

    # oncelik kolonu — Montaj iş emirlerinde sıralama için (1, 2, 3 ...)
    # NULL = öncelik belirtilmemiş. Sadece montaj için anlamlı, diğer bölümler ignore eder.
    try:
        c.execute("ALTER TABLE referans_takip ADD COLUMN oncelik INTEGER")
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
    # Migration: durum='uretimde'ye gectigi an — pilot sayac sifirlama referansi
    try:
        c.execute("ALTER TABLE referans_takip ADD COLUMN uretime_baslama_ts TEXT")
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

    # Fikstür Raf Tablosu (KAYNAKHANE FİKSTÜR RAF LİSTESİ.ods'tan gelir)
    c.execute('''
        CREATE TABLE IF NOT EXISTS fikstur_raf (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            referans_kodu TEXT NOT NULL,
            raf_no TEXT NOT NULL,
            guncelleme_tarihi TEXT DEFAULT (datetime('now', 'localtime'))
        )
    ''')
    try:
        c.execute('CREATE INDEX IF NOT EXISTS idx_fikstur_ref ON fikstur_raf(referans_kodu)')
    except Exception:
        pass

    # Migration (2026-05-14): Süre teyit takibi (kaynak için).
    # Kullanıcı sahada her referansın gerçek kaynak/söktak süresini ölçüp
    # 'teyit ettim' diye işaretler. Teyitsiz referanslar andonda işaretli görünür.
    try:
        c.execute("ALTER TABLE referans_listesi ADD COLUMN sure_teyit INTEGER DEFAULT 0")
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE referans_listesi ADD COLUMN sure_teyit_tarihi TEXT")
    except Exception:
        pass

    # Migration (2026-05-14): Robot kaynakta cycle_time'ı kaynak+söktak olarak ayır.
    # İki istasyon paralel çalıştığı için pair cycle = max(K1,S2)+max(K2,S1)
    # formülü kullanılabilsin diye.
    # İlk geçişte mevcut cycle_time × 0.4 = kaynak, × 0.6 = söktak (sonradan
    # kullanıcı her referansı tek tek günceller).
    try:
        c.execute("ALTER TABLE referans_listesi ADD COLUMN kaynak_suresi_sn REAL DEFAULT 0")
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE referans_listesi ADD COLUMN soktak_suresi_sn REAL DEFAULT 0")
    except Exception:
        pass
    # Sadece KAYNAK bölümü için ve henüz değer girilmemiş referanslar için doldur
    # (mevcut çalışan değerlere dokunma — idempotent)
    try:
        c.execute('''
            UPDATE referans_listesi
            SET kaynak_suresi_sn = ROUND(hedef_cycle_time_sn * 0.4, 1),
                soktak_suresi_sn = ROUND(hedef_cycle_time_sn * 0.6, 1)
            WHERE COALESCE(bolum, 'kaynak') = 'kaynak'
              AND hedef_cycle_time_sn > 0
              AND (kaynak_suresi_sn IS NULL OR kaynak_suresi_sn = 0)
              AND (soktak_suresi_sn IS NULL OR soktak_suresi_sn = 0)
        ''')
    except Exception as e:
        print(f'[migration] kaynak/söktak split atlandı: {e}')

    # Migration (2026-05-13): Süresi tanımsız referansların doğru bölüme atanması.
    # Eski kod operatör yeni bir referans girdiğinde bolum bilgisi olmadan INSERT
    # ediyordu → default 'kaynak'a düşüyordu. Şimdi vardiyanın bölümüne göre tag-le.
    # Sadece: (hedef_cycle_time = 0/NULL VE üretim kaydı olan VE bolum='kaynak') olanları
    # yeniden değerlendir; manuel tanımlanmış (cycle_time>0) referanslara dokunma.
    try:
        c.execute('''
            UPDATE referans_listesi
            SET bolum = (
                SELECT COALESCE(v.bolum, 'kaynak')
                FROM uretim_kayitlari u
                JOIN vardiyalar v ON v.id = u.vardiya_id
                WHERE UPPER(REPLACE(u.referans_kodu,' ','')) = UPPER(REPLACE(referans_listesi.referans_kodu,' ',''))
                GROUP BY v.bolum
                ORDER BY COUNT(*) DESC
                LIMIT 1
            )
            WHERE (hedef_cycle_time_sn IS NULL OR hedef_cycle_time_sn = 0)
              AND COALESCE(bolum, 'kaynak') = 'kaynak'
              AND EXISTS (
                  SELECT 1 FROM uretim_kayitlari u
                  JOIN vardiyalar v ON v.id = u.vardiya_id
                  WHERE UPPER(REPLACE(u.referans_kodu,' ','')) = UPPER(REPLACE(referans_listesi.referans_kodu,' ',''))
                    AND COALESCE(v.bolum, 'kaynak') != 'kaynak'
              )
        ''')
    except Exception as e:
        print(f'[migration] tanımsız referans bölüm düzeltmesi atlandı: {e}')

    # Migration (2026-05-13): Vardiya 'robotla_calisiyor' bayrağı
    # Metal enjeksiyonda makine + robot tam otomasyon modunda çalışabilir;
    # operatör vardiya sırasında bu modu açıp kapayabilir.
    try:
        c.execute("ALTER TABLE vardiyalar ADD COLUMN robotla_calisiyor INTEGER DEFAULT 0")
    except Exception:
        pass

    # Migration (2026-05-13): Metal enjeksiyondaki "500T" makinesinin gerçek adı "550T"
    # Tüm robot_no=='500T' kayıtları güncellenir. Tek seferlik, idempotent (ikinci çalıştırmada hiçbir şey yapmaz).
    try:
        for tablo in ('vardiyalar', 'referans_takip', 'andon_robot_ayarlari',
                      'sayac_olaylari', 'cihaz_kayitlari', 'robot_programlari'):
            try:
                c.execute(f"UPDATE {tablo} SET robot_no='550T' WHERE robot_no='500T'")
            except Exception:
                pass  # Tablo yoksa veya kolonsuzsa sessizce geç
    except Exception:
        pass

    # ─────────────────────────────────────────────────────────────
    # SAYAÇ OLAYLARI (ESP32 / PLC / sahadan gelen üretim pulse'ları)
    # Her pulse bir satır — vardiya_id ile ilişkilendirilir.
    # idempotency_key cihazın kendisinden gelir (device_id + sayac_no);
    # ağ titremesinde tekrar gelse bile UNIQUE constraint çift sayım önler.
    # ─────────────────────────────────────────────────────────────
    c.execute('''
        CREATE TABLE IF NOT EXISTS sayac_olaylari (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT DEFAULT (datetime('now', 'localtime')),
            bolum TEXT NOT NULL DEFAULT 'kaynak',
            robot_no TEXT NOT NULL,
            istasyon INTEGER NOT NULL DEFAULT 0,
            cihaz_id TEXT DEFAULT '',
            kaynak_tip TEXT DEFAULT 'robot_io',
            idempotency_key TEXT NOT NULL,
            vardiya_id INTEGER,
            FOREIGN KEY (vardiya_id) REFERENCES vardiyalar(id) ON DELETE SET NULL
        )
    ''')
    # Idempotency: aynı key iki kere gelirse INSERT'ler reddedilir
    try:
        c.execute('CREATE UNIQUE INDEX IF NOT EXISTS idx_sayac_idempotency ON sayac_olaylari(idempotency_key)')
    except Exception:
        pass
    # Hızlı sorgular için ek indexler (canlı counter — vardiya bazlı)
    try:
        c.execute('CREATE INDEX IF NOT EXISTS idx_sayac_vardiya ON sayac_olaylari(vardiya_id)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_sayac_robot ON sayac_olaylari(bolum, robot_no, istasyon, ts)')
    except Exception:
        pass

    # ─────────────────────────────────────────────────────────────
    # CİHAZ KAYITLARI (ESP32 heartbeat — son görülme + sağlık)
    # Andon'da "ABB1 sinyal yok 12 dk" uyarısı için kullanılır.
    # ─────────────────────────────────────────────────────────────
    c.execute('''
        CREATE TABLE IF NOT EXISTS cihaz_kayitlari (
            cihaz_id TEXT PRIMARY KEY,
            bolum TEXT DEFAULT '',
            robot_no TEXT DEFAULT '',
            firmware_ver TEXT DEFAULT '',
            ip_adresi TEXT DEFAULT '',
            mac_adresi TEXT DEFAULT '',
            son_heartbeat TEXT DEFAULT (datetime('now', 'localtime')),
            son_sinyal TEXT,
            toplam_sinyal INTEGER DEFAULT 0,
            buffer_kuyruk INTEGER DEFAULT 0,
            uptime_sn INTEGER DEFAULT 0,
            notlar TEXT DEFAULT ''
        )
    ''')

    conn.commit()
    conn.close()



if __name__ == '__main__':
    init_db()
    print("Veritabani hazir:", DB_PATH)
