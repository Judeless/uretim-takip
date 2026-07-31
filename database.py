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


def _migrate_composite_unique(c):
    """TK1/TK2 ayrı fabrika: aynı referans kodu / operatör adı iki lokasyonda
    legitimce var olabilmeli. Eski şema referans_kodu / ad'ı GLOBAL UNIQUE yapıyordu
    → çakışan kodlar tek lokasyona sıkışıyordu (INSERT OR IGNORE diğerini atlıyordu).
    Bu migration tabloları composite UNIQUE(.., lokasyon) ile yeniden kurar
    (robot_programlari rebuild paterni). Idempotent: sadece eski global UNIQUE varsa çalışır."""
    # referans_listesi → UNIQUE(referans_kodu, lokasyon)
    try:
        rl = c.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='referans_listesi'").fetchone()
        if rl and 'referans_kodu TEXT UNIQUE' in (rl[0] or ''):
            c.execute("DROP TABLE IF EXISTS _ref_eski")
            c.execute("ALTER TABLE referans_listesi RENAME TO _ref_eski")
            # TAZE DB SİGORTASI: init_db'deki bolum/lokasyon/kaynak-söktak ALTER'ları
            # CREATE'ten ÖNCE koştuğu için sıfırdan kurulan DB'de bu kolonlar henüz yok —
            # aşağıdaki INSERT..SELECT 'no such column' ile patlar ve satırlar _ref_eski'de
            # mahsur kalırdı. Eksik kolonları eski tabloya ekle (varsa sessizce geç).
            for _sql in (
                "ALTER TABLE _ref_eski ADD COLUMN bolum TEXT DEFAULT 'kaynak'",
                "ALTER TABLE _ref_eski ADD COLUMN kaynak_suresi_sn REAL DEFAULT 0",
                "ALTER TABLE _ref_eski ADD COLUMN soktak_suresi_sn REAL DEFAULT 0",
                "ALTER TABLE _ref_eski ADD COLUMN sure_teyit INTEGER DEFAULT 0",
                "ALTER TABLE _ref_eski ADD COLUMN sure_teyit_tarihi TEXT",
                "ALTER TABLE _ref_eski ADD COLUMN lokasyon TEXT DEFAULT 'TK2'",
            ):
                try:
                    c.execute(_sql)
                except Exception:
                    pass
            c.execute('''CREATE TABLE referans_listesi (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                referans_kodu TEXT NOT NULL,
                aciklama TEXT DEFAULT '',
                hedef_cycle_time_sn REAL DEFAULT 0,
                bolum TEXT DEFAULT 'kaynak',
                kaynak_suresi_sn REAL DEFAULT 0,
                soktak_suresi_sn REAL DEFAULT 0,
                sure_teyit INTEGER DEFAULT 0,
                sure_teyit_tarihi TEXT,
                lokasyon TEXT DEFAULT 'TK2',
                UNIQUE(referans_kodu, lokasyon)
            )''')
            c.execute('''INSERT INTO referans_listesi
                (id, referans_kodu, aciklama, hedef_cycle_time_sn, bolum, kaynak_suresi_sn,
                 soktak_suresi_sn, sure_teyit, sure_teyit_tarihi, lokasyon)
                SELECT id, referans_kodu, aciklama, hedef_cycle_time_sn, COALESCE(bolum,'kaynak'),
                       COALESCE(kaynak_suresi_sn,0), COALESCE(soktak_suresi_sn,0),
                       COALESCE(sure_teyit,0), sure_teyit_tarihi, COALESCE(lokasyon,'TK2')
                FROM _ref_eski''')
            c.execute("DROP TABLE _ref_eski")
            print('[MIGRATION] referans_listesi -> UNIQUE(referans_kodu, lokasyon)')
    except Exception as e:
        print(f'[MIGRATION] referans_listesi hata: {e}')
    # operatorler → UNIQUE(ad, lokasyon)
    try:
        op = c.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='operatorler'").fetchone()
        if op and 'ad TEXT UNIQUE' in (op[0] or ''):
            c.execute("DROP TABLE IF EXISTS _op_eski")
            c.execute("ALTER TABLE operatorler RENAME TO _op_eski")
            # Taze DB sigortası (yukarıdaki _ref_eski notuyla aynı sebep)
            for _sql in (
                "ALTER TABLE _op_eski ADD COLUMN bolum TEXT DEFAULT 'kaynak'",
                "ALTER TABLE _op_eski ADD COLUMN pin TEXT DEFAULT '0000'",
                "ALTER TABLE _op_eski ADD COLUMN lokasyon TEXT DEFAULT 'TK2'",
            ):
                try:
                    c.execute(_sql)
                except Exception:
                    pass
            c.execute('''CREATE TABLE operatorler (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ad TEXT NOT NULL,
                bolum TEXT DEFAULT 'kaynak',
                pin TEXT DEFAULT '0000',
                lokasyon TEXT DEFAULT 'TK2',
                UNIQUE(ad, lokasyon)
            )''')
            c.execute('''INSERT INTO operatorler (id, ad, bolum, pin, lokasyon)
                SELECT id, ad, COALESCE(bolum,'kaynak'), COALESCE(pin,'0000'), COALESCE(lokasyon,'TK2')
                FROM _op_eski''')
            c.execute("DROP TABLE _op_eski")
            print('[MIGRATION] operatorler -> UNIQUE(ad, lokasyon)')
    except Exception as e:
        print(f'[MIGRATION] operatorler hata: {e}')


def _migrate_bolum_composite_unique(c):
    """Çoklu bölüm desteği: aynı operatör birden fazla bölümde çalışabilir (İbrahim Nak:
    metal + işleme) ve aynı referans kodu birden fazla bölümde farklı süreyle var olabilir
    (aynı parça lazer'de kesilir, pres'te bükülür, kaynak'ta kaynatılır — her operasyonun
    cycle'ı farklı). Eski UNIQUE(ad, lokasyon) / UNIQUE(referans_kodu, lokasyon) bunu
    engelliyordu: import aynı ismi ikinci bölüme SESSİZCE atlıyor, ortak kodlar bölümler
    arasında el değiştiriyordu. bolum anahtara eklenerek bölüm başına satır açılır.
    Idempotent: sadece eski 2'li composite UNIQUE varsa çalışır."""
    # referans_listesi → UNIQUE(referans_kodu, bolum, lokasyon)
    try:
        rl = c.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='referans_listesi'").fetchone()
        if rl and 'UNIQUE(referans_kodu, lokasyon)' in (rl[0] or ''):
            c.execute("DROP TABLE IF EXISTS _ref_eski2")
            c.execute("ALTER TABLE referans_listesi RENAME TO _ref_eski2")
            c.execute('''CREATE TABLE referans_listesi (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                referans_kodu TEXT NOT NULL,
                aciklama TEXT DEFAULT '',
                hedef_cycle_time_sn REAL DEFAULT 0,
                bolum TEXT DEFAULT 'kaynak',
                kaynak_suresi_sn REAL DEFAULT 0,
                soktak_suresi_sn REAL DEFAULT 0,
                sure_teyit INTEGER DEFAULT 0,
                sure_teyit_tarihi TEXT,
                lokasyon TEXT DEFAULT 'TK2',
                UNIQUE(referans_kodu, bolum, lokasyon)
            )''')
            c.execute('''INSERT INTO referans_listesi
                (id, referans_kodu, aciklama, hedef_cycle_time_sn, bolum, kaynak_suresi_sn,
                 soktak_suresi_sn, sure_teyit, sure_teyit_tarihi, lokasyon)
                SELECT id, referans_kodu, aciklama, hedef_cycle_time_sn, COALESCE(bolum,'kaynak'),
                       COALESCE(kaynak_suresi_sn,0), COALESCE(soktak_suresi_sn,0),
                       COALESCE(sure_teyit,0), sure_teyit_tarihi, COALESCE(lokasyon,'TK2')
                FROM _ref_eski2''')
            c.execute("DROP TABLE _ref_eski2")
            print('[MIGRATION] referans_listesi -> UNIQUE(referans_kodu, bolum, lokasyon)')
    except Exception as e:
        print(f'[MIGRATION] referans_listesi (bolum) hata: {e}')
    # operatorler → UNIQUE(ad, bolum, lokasyon)
    # NOT: aynı kişinin bölüm satırları TEK PIN paylaşır — import yeni bölüm satırını
    # mevcut satırın PIN'iyle açar, PIN güncelleme (app.py) tüm satırlara yazar.
    try:
        op = c.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='operatorler'").fetchone()
        if op and 'UNIQUE(ad, lokasyon)' in (op[0] or ''):
            c.execute("DROP TABLE IF EXISTS _op_eski2")
            c.execute("ALTER TABLE operatorler RENAME TO _op_eski2")
            c.execute('''CREATE TABLE operatorler (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ad TEXT NOT NULL,
                bolum TEXT DEFAULT 'kaynak',
                pin TEXT DEFAULT '0000',
                lokasyon TEXT DEFAULT 'TK2',
                UNIQUE(ad, bolum, lokasyon)
            )''')
            c.execute('''INSERT INTO operatorler (id, ad, bolum, pin, lokasyon)
                SELECT id, ad, COALESCE(bolum,'kaynak'), COALESCE(pin,'0000'), COALESCE(lokasyon,'TK2')
                FROM _op_eski2''')
            c.execute("DROP TABLE _op_eski2")
            print('[MIGRATION] operatorler -> UNIQUE(ad, bolum, lokasyon)')
    except Exception as e:
        print(f'[MIGRATION] operatorler (bolum) hata: {e}')


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

    # Sayaç entegrasyonu (2026-06): operatör yeni referans kaydı açınca sensör
    # sayacı o andan sıfırlanır. sayac_baslangic_ts = referansın üretime başlama
    # anı (pulse filtresi bu ts'den sonrasını sayar). sayac_otomatik=1 iken ok_adet
    # sensörden otomatik beslenir; operatör elle düzeltince 0 olur (donar).
    try:
        c.execute("ALTER TABLE uretim_kayitlari ADD COLUMN sayac_baslangic_ts TEXT")
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE uretim_kayitlari ADD COLUMN sayac_otomatik INTEGER DEFAULT 0")
    except Exception:
        pass

    # Test cihazı sayaç kaynağı (2026-06): operatör harici test cihazı (Cofle/wicow) ile
    # çalışıyorsa bu kolon o cihazın id'sini tutar. Dolu ise sayaç ESP32 pulse yerine
    # test_sonuclari'ndaki başarılı testlerden beslenir (bkz. cofle_test.py).
    try:
        c.execute("ALTER TABLE uretim_kayitlari ADD COLUMN test_cihaz_id INTEGER")
    except Exception:
        pass

    # BÜKÜM OPERASYON SAYISI (2026-07-29) — Pres Abkant'a özel sayaç bölücüsü.
    # Bükümde bir parçaya SIRAYLA tüm büküm operasyonları uygulanır, sonra sıradaki
    # parçaya geçilir. Yani 3 operasyonlu bir referansta abkant 3 kez sinyal üretir
    # ama ortaya 1 parça çıkar → üretim teyidi 3. sinyaldir (ok_adet = pulse // 3).
    #
    # İKİ YERDE tutulur ve bu BİLİNÇLİDİR:
    #   referans_listesi.bukum_operasyon → referansın HATIRLANAN varsayılanı; Excel'e
    #     yazılır, sonraki üretimlerde operatöre hazır gelir, operatör değiştirebilir.
    #   uretim_kayitlari.bukum_operasyon → O KAYIT için fiilen kullanılan değer.
    #     Referansın varsayılanı sonradan değişirse geçmiş kayıtların adedi
    #     retroaktif olarak kaymasın diye kayda kopyalanır.
    # Varsayılan 1 = bölme yok = mevcut davranış (tüm eski kayıtlar/bölümler etkilenmez).
    try:
        c.execute("ALTER TABLE uretim_kayitlari ADD COLUMN bukum_operasyon INTEGER DEFAULT 1")
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE referans_listesi ADD COLUMN bukum_operasyon INTEGER DEFAULT 1")
    except Exception:
        pass
    # NULL'ları 1'e çek (ALTER ile eklenen kolon eski satırlarda NULL kalabilir;
    # bölme işlemi NULL görürse sayım çöker).
    for _t in ('uretim_kayitlari', 'referans_listesi'):
        try:
            c.execute(f"UPDATE {_t} SET bukum_operasyon=1 "
                      f"WHERE bukum_operasyon IS NULL OR bukum_operasyon < 1")
        except Exception:
            pass

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

    # ── lokasyon kolonu (TK1/TK2 çok-tesis desteği) ──────────────────
    # Mevcut tüm veri default 'TK2' (mevcut fabrika). TK1 (yan tesis) verileri
    # lokasyon='TK1' ile eklenir. bolum paterninin aynısı; lokasyonsuz sorgu = TK2.
    # NOT: andon_robot_ayarlari'nın lokasyon ALTER'ı CREATE'inden SONRA (aşağıda).
    # sayac_olaylari/cihaz_kayitlari ATLANDI — sensör verisi pilot/pilot.db'de tutulur,
    # bu tablolar ana DB'de pratikte kullanılmıyor; lokasyon eşlemesi robot_no ile yapılır.
    for tbl in ('vardiyalar', 'referans_listesi', 'operatorler', 'referans_takip'):
        try:
            c.execute(f"ALTER TABLE {tbl} ADD COLUMN lokasyon TEXT DEFAULT 'TK2'")
        except Exception:
            pass  # Kolon zaten var
    # NOT: uniqueness artık _migrate_composite_unique + _migrate_bolum_composite_unique ile
    # operatorler=UNIQUE(ad, bolum, lokasyon), referans_listesi=UNIQUE(referans_kodu, bolum,
    # lokasyon) olarak rebuild edilir (aynı kişi/kod birden çok bölümde var olabilir).

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
            ad TEXT UNIQUE NOT NULL,
            pin TEXT DEFAULT '0000'
        )
    ''')
    # Migration: mevcut DB'lerde pin kolonu yoksa ekle (default '0000')
    try:
        c.execute("ALTER TABLE operatorler ADD COLUMN pin TEXT DEFAULT '0000'")
    except Exception:
        pass

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
    # Migration (2026-07-16): depo notu — depocu kit hazirlarken durumu (eksik parca vb.)
    # yazar. aciklama'dan AYRI: aciklama = planlayicinin is tanimi; depo_notu = deponun notu.
    try:
        c.execute("ALTER TABLE referans_takip ADD COLUMN depo_notu TEXT DEFAULT ''")
    except Exception:
        pass

    # Andon Robot Ayarları Tablosu — bolum kolonu PRIMARY KEY'in parçası
    # (M1 hat'i hem montaj'da hem kaynak'ta görünmesin)
    c.execute('''
        CREATE TABLE IF NOT EXISTS andon_robot_ayarlari (
            robot_no TEXT NOT NULL,
            bolum TEXT NOT NULL DEFAULT 'kaynak',
            goster INTEGER NOT NULL DEFAULT 1,
            sira INTEGER NOT NULL DEFAULT 0,
            PRIMARY KEY (robot_no, bolum)
        )
    ''')
    # Migration (2026-05-22): eski PK robot_no idi, yeni (robot_no, bolum) — bolum kolonu eklendi
    try:
        cols = [r[1] for r in c.execute("PRAGMA table_info(andon_robot_ayarlari)").fetchall()]
        if 'bolum' not in cols:
            c.execute("ALTER TABLE andon_robot_ayarlari ADD COLUMN bolum TEXT NOT NULL DEFAULT 'kaynak'")
    except Exception:
        pass
    # lokasyon kolonu — TK1/TK2 ayrımı. PK (robot_no, bolum) DEĞİŞMEZ (TK1 hat adları
    # M1-M12 ile çakışmaz); lokasyon sadece andon/dashboard filtresi içindir.
    try:
        c.execute("ALTER TABLE andon_robot_ayarlari ADD COLUMN lokasyon TEXT DEFAULT 'TK2'")
    except Exception:
        pass

    # Genel Ayarlar Tablosu
    c.execute('''
        CREATE TABLE IF NOT EXISTS genel_ayarlar (
            anahtar TEXT PRIMARY KEY,
            deger TEXT
        )
    ''')
    # Varsayılan ayarları ekle
    c.execute("INSERT OR IGNORE INTO genel_ayarlar (anahtar, deger) VALUES ('andon_font_size', '0.57')")

    # Varsayılan robot satırlarını ekle (mevcut değilse) — 3 bölüm için
    # KAYNAK: ABB1..ABB9 (ABB9 default gizli)
    for i, rno in enumerate(['ABB1','ABB2','ABB3','ABB4','ABB5','ABB6','ABB7','ABB8','ABB9']):
        goster = 0 if rno == 'ABB9' else 1
        c.execute(
            'INSERT OR IGNORE INTO andon_robot_ayarlari (robot_no, bolum, goster, sira) VALUES (?, ?, ?, ?)',
            (rno, 'kaynak', goster, i)
        )
    # MONTAJ: M1..M12 (default hepsi gizli — operatör vardiya açtığında dinamik gelir)
    for i, mno in enumerate([f'M{n}' for n in range(1, 13)]):
        c.execute(
            'INSERT OR IGNORE INTO andon_robot_ayarlari (robot_no, bolum, goster, sira) VALUES (?, ?, ?, ?)',
            (mno, 'montaj', 1, i)
        )
    # METAL: 300T, 400T, 550T (hepsi default görünür)
    for i, mno in enumerate(['300T','400T','550T']):
        c.execute(
            'INSERT OR IGNORE INTO andon_robot_ayarlari (robot_no, bolum, goster, sira) VALUES (?, ?, ?, ?)',
            (mno, 'metal', 1, i)
        )
    # TK1 (yan tesis) hatları — montaj mantığı, lokasyon='TK1'. PK (robot_no, bolum)
    # ile TK2 montaj M1-M12'den ayrı (farklı robot_no). Hepsi default görünür.
    for i, hat in enumerate(['Pull', 'Push-Pull', 'Iveco', 'LF-LFP']):
        c.execute(
            'INSERT OR IGNORE INTO andon_robot_ayarlari (robot_no, bolum, goster, sira, lokasyon) VALUES (?, ?, ?, ?, ?)',
            (hat, 'montaj', 1, i, 'TK1')
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
              -- Hedef bölümde aynı kod zaten varsa retag ETME: UNIQUE(kod, bolum, lokasyon)
              -- ile çakışır ve tek UPDATE atomik olduğundan TÜM düzeltmeler iptal olurdu.
              AND NOT EXISTS (
                  SELECT 1 FROM referans_listesi r2
                  WHERE r2.id != referans_listesi.id
                    AND UPPER(REPLACE(r2.referans_kodu,' ','')) = UPPER(REPLACE(referans_listesi.referans_kodu,' ',''))
                    AND COALESCE(r2.lokasyon,'TK2') = COALESCE(referans_listesi.lokasyon,'TK2')
                    AND COALESCE(r2.bolum,'kaynak') = (
                        SELECT COALESCE(v.bolum, 'kaynak')
                        FROM uretim_kayitlari u
                        JOIN vardiyalar v ON v.id = u.vardiya_id
                        WHERE UPPER(REPLACE(u.referans_kodu,' ','')) = UPPER(REPLACE(referans_listesi.referans_kodu,' ',''))
                        GROUP BY v.bolum
                        ORDER BY COUNT(*) DESC
                        LIMIT 1
                    )
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

    # Migration (2026-06-11): Otomatik mola takip bitmask'ı.
    # Çay 09:00(1)/13:00(2)/15:00(4) 10dk + Yemek 11:00(8) 30dk planlı duruş olarak
    # bir kez otomatik işlenir (aynı türden tek kayıtta birikir); hangi molanın
    # işlendiği bit bit burada tutulur (operatör kaydı silse de tekrar eklenmez).
    try:
        c.execute("ALTER TABLE vardiyalar ADD COLUMN cay_molasi_bayrak INTEGER DEFAULT 0")
    except Exception:
        pass

    # Migration (2026-06-12): Web Push abonelikleri — operatör telefonuna bildirim.
    # Operatör mobilde PIN ile giriş yapınca tarayıcı push aboneliği buraya yazılır;
    # launch/referans eklendiğinde ilgili operatörlerin cihazlarına bildirim gönderilir.
    c.execute('''
        CREATE TABLE IF NOT EXISTS push_abonelikler (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            operator_adi TEXT NOT NULL,
            endpoint TEXT NOT NULL UNIQUE,
            p256dh TEXT NOT NULL,
            auth TEXT NOT NULL,
            olusturma TEXT DEFAULT (datetime('now','localtime'))
        )
    ''')

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

    # ─────────────────────────────────────────────────────────────
    # MAIL ALICILARI — günlük üretim raporu e-posta alıcı listesi
    # (dashboard'dan yönetilir; SMTP bilgileri mail_config.json'da — DB'de değil)
    # ─────────────────────────────────────────────────────────────
    c.execute('''
        CREATE TABLE IF NOT EXISTS mail_alicilari (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT NOT NULL UNIQUE,
            ad TEXT DEFAULT '',
            aktif INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    ''')

    # ── Composite UNIQUE migration (TK1/TK2 tam ayrışma) ──
    _migrate_composite_unique(c)
    # ── Bölüm-composite UNIQUE migration (çoklu bölüm: işleme/lazer/pres genişlemesi) ──
    _migrate_bolum_composite_unique(c)

    # ── Yönetici kullanıcı (Admin, PIN 9999) — her iki tesiste ──
    # Bu adla PIN girişi yapan mobilde TÜM vardiyalara erişir (app.py ADMIN_ADI).
    # INSERT OR IGNORE: PIN panelden değiştirilmişse ezilmez. Migration'dan SONRA
    # (UNIQUE(ad,lokasyon) aktifken) çalışmalı ki iki lokasyon kaydı da eklenebilsin.
    for _lok in ('TK2', 'TK1'):
        try:
            c.execute("INSERT OR IGNORE INTO operatorler (ad, bolum, pin, lokasyon) VALUES ('Admin', 'montaj', '9999', ?)", (_lok,))
        except Exception:
            pass

    # ─────────────────────────────────────────────────────────────
    # TEST SONUÇLARI (Cofle/wicow harici test cihazı entegrasyonu)
    # cofle_test.py poller'ı API'den çeker. Bir test cihazı = sayaç kaynağı:
    # başarılı test (result 'PASSED…') = 1 pulse. test_id = API'nin benzersiz id'si
    # (poller bununla dedup eder). Sayaç: referans başlangıcından beri basarili=1 say.
    # ─────────────────────────────────────────────────────────────
    c.execute('''
        CREATE TABLE IF NOT EXISTS test_sonuclari (
            test_id INTEGER PRIMARY KEY,
            cihaz_id INTEGER NOT NULL,
            cihaz_adi TEXT DEFAULT '',
            referans_kodu TEXT DEFAULT '',
            referans_norm TEXT DEFAULT '',
            sonuc TEXT DEFAULT '',
            basarili INTEGER DEFAULT 0,
            ts_epoch INTEGER DEFAULT 0,
            operator TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    ''')
    try:
        c.execute('CREATE INDEX IF NOT EXISTS idx_test_sonuc_say ON test_sonuclari(cihaz_id, referans_norm, basarili, ts_epoch)')
    except Exception:
        pass

    # ─────────────────────────────────────────────────────────────
    # PANEL KULLANICILARI (Dashboard/yönetici paneli giriş + yetki)
    # Şifreler HASH'li saklanır (werkzeug). rol='admin' → tüm sayfalar +
    # kullanıcı yönetimi. rol='kullanici' → yalnız izinler listesindeki
    # sayfalar. izinler = JSON dizi, sayfa anahtarları (data-sayfa değerleri):
    #   ozet, bolum, kayitlar, is-yonetimi, fikstur, referanslar,
    #   operatorler, saha-cihazlari, sinyal-analizi, andon-ayarlari, raporlar
    # sifre_gecici=1 → kullanıcı ilk girişte şifresini değiştirmeye zorlanır.
    # ─────────────────────────────────────────────────────────────
    c.execute('''
        CREATE TABLE IF NOT EXISTS panel_kullanicilari (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kullanici_adi TEXT NOT NULL UNIQUE,
            ad_soyad TEXT DEFAULT '',
            sifre_hash TEXT NOT NULL,
            rol TEXT DEFAULT 'kullanici',
            izinler TEXT DEFAULT '[]',
            aktif INTEGER DEFAULT 1,
            sifre_gecici INTEGER DEFAULT 0,
            son_giris TEXT,
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    ''')
    # ─────────────────────────────────────────────────────────────
    # AS400 TEYIT GONDERIM GUNLUGU — robotun ERP'ye yazdigi her teyit.
    # Mukerrer korumasi: ayni (uretim_tarihi, launch) icin sonuc='ok' kayit
    # varsa ikinci gonderim reddedilir (zorla bayragi haric).
    # teyitli_once/sonra: gonderim oncesi/sonrasi BPROF0.Q0QTRI (bagimsiz dogrulama).
    # ─────────────────────────────────────────────────────────────
    c.execute('''
        CREATE TABLE IF NOT EXISTS as400_teyit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            uretim_tarihi TEXT NOT NULL,
            yil TEXT NOT NULL,
            launch_no TEXT NOT NULL,
            referans TEXT DEFAULT '',
            article TEXT DEFAULT '',
            adet REAL NOT NULL,
            bayrak TEXT DEFAULT 'A',
            sonuc TEXT NOT NULL,
            mesaj TEXT DEFAULT '',
            teyitli_once REAL,
            teyitli_sonra REAL,
            olusturan TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    ''')

    # ─────────────────────────────────────────────────────────────
    # AS400 teyit ekranı İŞARETLERİ (2026-07-20). İki kapsam:
    #  - kapsam='kalici': referans bazında SÜREKLİ 'gerek_yok' (örn 6343a ara ürün)
    #    → launch_esle bu referansları teyit dışı bırakır (her gün otomatik gizli).
    #  - kapsam='gun'   : o üretim gününe özel 'kontrol' (kullanıcı elle gözden geçirdi).
    # referans: normalize (UPPER + boşluksuz) saklanır. uretim_tarihi kalıcıda ''.
    # ─────────────────────────────────────────────────────────────
    c.execute('''
        CREATE TABLE IF NOT EXISTS as400_teyit_isaret (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kapsam TEXT NOT NULL DEFAULT 'gun',
            uretim_tarihi TEXT DEFAULT '',
            bolum TEXT DEFAULT '',
            referans TEXT NOT NULL,
            durum TEXT NOT NULL,
            aciklama TEXT DEFAULT '',
            olusturan TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now', 'localtime')),
            UNIQUE(kapsam, uretim_tarihi, bolum, referans)
        )
    ''')

    # ─────────────────────────────────────────────────────────────
    # KAYNAK PLANI (2026-07-31). Planlama her hafta bir "Kaynak ihtiyaçları"
    # dosyası çıkarıyor; referanslar öncelik sırasına dizili. Launch talimatı
    # için her kod AS400'de 07>10>01 ile açılıp 1. seviye alt parçaların
    # 01D/CF2 stoğu elle kontrol ediliyordu (230 satır).
    # Plan buraya alınır, ERP'den hesaplanır ve HAFTA BOYUNCA takip edilir:
    #   - not          : kullanıcının o referansa yazdığı serbest not (kalıcı)
    #   - uretilebilir : son ölçümde min(01D+CF2 / birim miktar)
    #   - onceki_*     : bir önceki ölçüm → "stoğa bir şey geldi mi?" farkı
    # Yenilemede plan satırı KORUNUR (not kaybolmasın), yalnız ölçüm güncellenir.
    # ─────────────────────────────────────────────────────────────
    c.execute('''
        CREATE TABLE IF NOT EXISTS kaynak_plan (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kaynak_kod TEXT NOT NULL UNIQUE,
            sira INTEGER DEFAULT 0,
            urun TEXT DEFAULT '',
            acik_launch REAL DEFAULT 0,
            gereken REAL DEFAULT 0,
            kontrol6 INTEGER,
            gun_stok REAL DEFAULT 0,
            bakiye REAL DEFAULT 0,
            toplam_stok REAL DEFAULT 0,
            iht_2h REAL DEFAULT 0,
            iht_6h REAL DEFAULT 0,
            plan_dosya TEXT DEFAULT '',
            plan_yuklendi TEXT DEFAULT '',
            uretilebilir INTEGER,
            kisitlayan TEXT DEFAULT '',
            kisit_stok REAL DEFAULT 0,
            parca_sayisi INTEGER DEFAULT 0,
            durum TEXT DEFAULT '',
            karar TEXT DEFAULT '',
            agac_farki TEXT DEFAULT '',
            eksi_var INTEGER DEFAULT 0,
            olculdu TEXT DEFAULT '',
            onceki_uretilebilir INTEGER,
            onceki_olculdu TEXT DEFAULT '',
            aciklama TEXT DEFAULT '',
            not_guncelleyen TEXT DEFAULT '',
            not_guncellendi TEXT DEFAULT '',
            aktif INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    ''')
    c.execute('CREATE INDEX IF NOT EXISTS ix_kaynak_plan_sira ON kaynak_plan(aktif, sira)')
    # Alt parça kırılımı — "hangi parça kısıtlıyor, stoğu ne" ekranda görünsün.
    # Her yenilemede o kodun satırları silinip yeniden yazılır.
    c.execute('''
        CREATE TABLE IF NOT EXISTS kaynak_plan_parca (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kaynak_kod TEXT NOT NULL,
            alt_kod TEXT NOT NULL,
            birim REAL DEFAULT 0,
            um TEXT DEFAULT '',
            stok_01d REAL DEFAULT 0,
            stok_cf2 REAL DEFAULT 0,
            stok_sayilan REAL DEFAULT 0,
            kapasite INTEGER,
            eksi_bakiye INTEGER DEFAULT 0,
            diger_depolar TEXT DEFAULT '',
            onceki_stok REAL,
            olculdu TEXT DEFAULT ''
        )
    ''')
    c.execute('CREATE INDEX IF NOT EXISTS ix_kaynak_plan_parca_kod ON kaynak_plan_parca(kaynak_kod)')

    # ─────────────────────────────────────────────────────────────
    # AS400 TRANSFER İPTAL log'u (2026-07-27). TK1 hayali stok transferleri
    # (TA 02→01) robotla iptal edilir (07>01 + satır başı 2 + Shift+F4).
    # kayit = AS400 kayıt no (MGANRE-MGNURE, örn '26-245458'), satir = MGPRRE.
    # sonuc: ok | hata | atlandi. Doğrulama: iptal sonrası BMMAF0'da kayıt
    # aranır — silinmişse (fiziksel silme) 'ok'.
    # ─────────────────────────────────────────────────────────────
    c.execute('''
        CREATE TABLE IF NOT EXISTS as400_transfer_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kayit TEXT NOT NULL,
            satir INTEGER DEFAULT 0,
            article TEXT DEFAULT '',
            adet REAL DEFAULT 0,
            hareket_tarihi TEXT DEFAULT '',
            sonuc TEXT NOT NULL,
            mesaj TEXT DEFAULT '',
            olusturan TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    ''')

    # ─────────────────────────────────────────────────────────────
    # AS400 OTOMATİK KOŞU raporları (2026-07-27). Zamanlanmış işler:
    #  16:45 transfer iptali (tur='transfer'), 17:10 tüm teyit kuyruğu (tur='teyit').
    # detay = JSON: {'sonuclar': [...], 'sabah_kontrol': [...]} — dashboard
    # "Sabah Kontrol" paneli buradan okur. ok/hata/atlandi/kontrol = hızlı sayaçlar.
    # ─────────────────────────────────────────────────────────────
    c.execute('''
        CREATE TABLE IF NOT EXISTS as400_oto_kosu (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tur TEXT NOT NULL,
            ozet TEXT DEFAULT '',
            detay TEXT DEFAULT '',
            ok INTEGER DEFAULT 0,
            hata INTEGER DEFAULT 0,
            atlandi INTEGER DEFAULT 0,
            kontrol INTEGER DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    ''')

    # Seed: ilk admin (emre.dogutekin). Geçici şifre 'cofle1234' — ilk girişte
    # değiştirilir. INSERT OR IGNORE: şifre/izin panelden değiştirilmişse ezilmez.
    try:
        from werkzeug.security import generate_password_hash
        c.execute(
            "INSERT OR IGNORE INTO panel_kullanicilari (kullanici_adi, ad_soyad, sifre_hash, rol, izinler, aktif, sifre_gecici) "
            "VALUES (?, ?, ?, 'admin', '[]', 1, 1)",
            ('emre.dogutekin', 'Emre Doğutekin', generate_password_hash('cofle1234'))
        )
    except Exception as _seed_err:
        print('Panel admin seed hatası:', _seed_err)

    conn.commit()
    conn.close()



if __name__ == '__main__':
    init_db()
    print("Veritabani hazir:", DB_PATH)
