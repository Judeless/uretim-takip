# -*- coding: utf-8 -*-
"""
COFLE MANAGE — PILOT SAYAÇ BACKEND'i
====================================

Bu dosya **ana sistemden tamamen bağımsız** çalışan Flask uygulamasıdır.
Pilot başarılı olduktan sonra bu kod ana app.py'a entegre edilir; o aşamaya
kadar burası kanaryadır — bozulsa ana sistemi etkilemez.

- Port: 5001 (ana sistem 5000)
- DB:   pilot.db (ana uretim.db'ye dokunmaz)
- ESP32 endpoint'leri: /api/sinyal, /api/sinyal/heartbeat
- Canlı sayaç UI: /
"""
from flask import Flask, request, jsonify, render_template, send_file
import sqlite3
import os
import time
from datetime import datetime, date

# ─── Yapılandırma ─────────────────────────────────────────────────
PORT = 5001
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pilot.db')
# ESP32'lerin POST'larında Authorization header'da göndereceği token.
# Saha güvenliği için sahip yumruk kuralı (LAN-only kullanımda yeterli).
API_TOKEN = 'cofle-pilot-2026'  # ESP32 firmware'inde de aynı yazacak

app = Flask(__name__, template_folder='templates')


# ─── DB başlatma ──────────────────────────────────────────────────
def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=15.0)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute('PRAGMA synchronous=NORMAL')
    return conn


def init_db():
    conn = get_db()
    c = conn.cursor()
    # Saha cihazlarından gelen her pulse bir satır.
    # idempotency_key UNIQUE — ağ titremesinde aynı key tekrar gelirse INSERT reddedilir,
    # böylece çift sayım önlenir.
    c.execute('''
        CREATE TABLE IF NOT EXISTS sayac_olaylari (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT DEFAULT (datetime('now', 'localtime')),
            cihaz_id TEXT NOT NULL,
            bolum TEXT DEFAULT 'kaynak',
            robot_no TEXT NOT NULL,
            istasyon INTEGER NOT NULL DEFAULT 0,
            kaynak_tip TEXT DEFAULT 'robot_io',
            idempotency_key TEXT NOT NULL UNIQUE
        )
    ''')
    c.execute('CREATE INDEX IF NOT EXISTS idx_sayac_cihaz_ts ON sayac_olaylari(cihaz_id, ts)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_sayac_robot_ts ON sayac_olaylari(bolum, robot_no, istasyon, ts)')

    # Cihaz heartbeat — her ESP32 30sn'de bir "ben hayattayım" der
    c.execute('''
        CREATE TABLE IF NOT EXISTS cihaz_kayitlari (
            cihaz_id TEXT PRIMARY KEY,
            bolum TEXT DEFAULT '',
            robot_no TEXT DEFAULT '',
            firmware_ver TEXT DEFAULT '',
            ip_adresi TEXT DEFAULT '',
            mac_adresi TEXT DEFAULT '',
            wifi_rssi INTEGER DEFAULT 0,
            son_heartbeat TEXT DEFAULT (datetime('now', 'localtime')),
            son_sinyal TEXT,
            toplam_sinyal INTEGER DEFAULT 0,
            buffer_kuyruk INTEGER DEFAULT 0,
            uptime_sn INTEGER DEFAULT 0,
            free_heap INTEGER DEFAULT 0,
            notlar TEXT DEFAULT '',
            robot_calisiyor INTEGER DEFAULT 0,
            robot_durum_zamani TEXT
        )
    ''')
    # Migration: mevcut DB'lerde yoksa ekle
    for col, ddl in [
        ('robot_calisiyor',     'INTEGER DEFAULT 0'),
        ('robot_durum_zamani',  'TEXT'),
        # Firmware v2.2+ interrupt-based — ISR raw count (loop seviyesinde filtre oncesi)
        # isr_count - toplam_sinyal = MIN_PULSE_GAP_MS ile reddedilen back-to-back pulse sayisi
        ('isr_count_ist1',      'INTEGER DEFAULT 0'),
        ('isr_count_ist2',      'INTEGER DEFAULT 0'),
        # Firmware v2.4+ tani — sayim filtresi kararlarinin kumulatif sayaclari
        # (eksik/fazla sayma teshisi: parazit=opto zayif, erken=cycle/gap)
        ('tani_sayildi',        'INTEGER DEFAULT 0'),
        ('tani_parazit',        'INTEGER DEFAULT 0'),
        ('tani_erken',          'INTEGER DEFAULT 0'),
        # Cihazin NVS'teki reset-korumali sayimi (heartbeat'te gelir). server'a ulasan
        # sayim (sayac_olaylari) ile karsilastirilip transit/reset kaybi tespit edilir.
        # base_ist* = ilk goruste kurulan baseline offset (server_count - pulse_ist).
        # kayip = max(0, pulse_ist + base - server_count) (okuma aninda hesaplanir).
        ('pulse_ist1',          'INTEGER DEFAULT 0'),
        ('pulse_ist2',          'INTEGER DEFAULT 0'),
        ('base_ist1',           'INTEGER'),   # NULL = henuz baseline kurulmadi
        ('base_ist2',           'INTEGER'),
        # Firmware v2.6+ self-heal/zombie teshisi (heartbeat ile gelir)
        ('radyo_yenile',        'INTEGER DEFAULT 0'),   # radyoYenileToplam — artiyor ama offline ise kurtaramayan tedavi
        ('http_fail',           'INTEGER DEFAULT 0'),   # httpFailCount — RSSI iyi + artiyorsa zombie WL_CONNECTED isareti
        ('disconnect_say',      'INTEGER DEFAULT 0'),   # disconnectCount (10dk penceresi)
    ]:
        try:
            c.execute(f"ALTER TABLE cihaz_kayitlari ADD COLUMN {col} {ddl}")
        except Exception:
            pass  # zaten var

    # ─── Cihaz sağlık zaman serisi (append-only) ───
    # cihaz_kayitlari TEK satır (her heartbeat'te ezilir) → "iyi RSSI ama offline" gibi
    # olayların post-mortem'i imkansızdı. Bu tablo ANOMALİ anlarını (reboot/radyo tazele/
    # buffer dolu/httpfail) + ~10dk'da bir baseline'ı KALICI loglar; 30 günden eski budanır.
    c.execute('''
        CREATE TABLE IF NOT EXISTS saglik_olaylari (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT DEFAULT (datetime('now','localtime')),
            cihaz_id TEXT NOT NULL,
            robot_no TEXT DEFAULT '',
            sebep TEXT DEFAULT '',
            wifi_rssi INTEGER,
            free_heap INTEGER,
            uptime_sn INTEGER,
            buffer_kuyruk INTEGER,
            radyo_yenile INTEGER,
            http_fail INTEGER,
            disconnect_say INTEGER
        )
    ''')
    c.execute('CREATE INDEX IF NOT EXISTS idx_saglik_cihaz_ts ON saglik_olaylari(cihaz_id, ts)')

    # Sayac reset noktalari — yonetici dashboard'dan "sayaci sifirla" tiklayinca
    # bu tabloya o cihaz/istasyon icin reset timestamp'i yazilir.
    # Saha cihazlari endpoint'i sayim sirasinda ts > reset_ts filtresini de uygular.
    # istasyon=0 → tum istasyonlari kapsayan reset (montaj/metal icin)
    # istasyon=1 veya 2 → sadece o istasyon (kaynak icin)
    c.execute('''
        CREATE TABLE IF NOT EXISTS sayac_reset_noktalari (
            cihaz_id   TEXT NOT NULL,
            istasyon   INTEGER NOT NULL DEFAULT 0,
            reset_ts   TEXT NOT NULL,
            yapan      TEXT DEFAULT '',
            PRIMARY KEY (cihaz_id, istasyon)
        )
    ''')

    # Tani olaylari — firmware'in sayim filtresi her karar verdiginde bir satir.
    # Heartbeat'e binip gelir (firmware v2.4+). Eksik/fazla sayma teshisi icin:
    #   tip='SAYILDI' → gecerli pulse sayildi (gap_ms=onceki sayimdan fark, low_ms=LOW suresi)
    #   tip='PARAZIT' → multi-sample basarisiz (low_ms=kopmadan once kac ms LOW kaldi); opto zayif belirtisi
    #   tip='ERKEN'   → gap < MIN_PULSE_GAP, pulse elendi (gap_ms); cycle gap'ten kisa demek
    # low_ms (SAYILDI): ~22sn → tek parca; ~2x → role ardisik parcalari birlestirdi (eksik sayim).
    # idem UNIQUE → ayni olay tekrar heartbeat ile gelirse INSERT reddedilir (cift kayit yok).
    c.execute('''
        CREATE TABLE IF NOT EXISTS tani_olaylari (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT DEFAULT (datetime('now', 'localtime')),
            cihaz_id TEXT NOT NULL,
            robot_no TEXT DEFAULT '',
            uptime_ms INTEGER DEFAULT 0,
            tip TEXT DEFAULT '',
            gap_ms INTEGER DEFAULT 0,
            low_ms INTEGER DEFAULT 0,
            dev_seq INTEGER DEFAULT 0,
            idem TEXT NOT NULL UNIQUE
        )
    ''')
    c.execute('CREATE INDEX IF NOT EXISTS idx_tani_cihaz_ts ON tani_olaylari(cihaz_id, ts)')
    conn.commit()
    conn.close()


# ─── Yetki kontrolü ───────────────────────────────────────────────
def _yetki_kontrol():
    """ESP32 cihazlarından gelen istekler Authorization: Bearer <token> ile."""
    auth = request.headers.get('Authorization', '')
    return auth == f'Bearer {API_TOKEN}'


# ─── Endpoints ────────────────────────────────────────────────────
@app.route('/')
def index():
    """Canlı sayaç / cihaz durumu UI'sı."""
    return render_template('pilot.html')


@app.route('/api/sinyal', methods=['POST'])
def sinyal_al():
    """ESP32 her parça çıkışında bunu çağırır.
    JSON: { cihaz_id, bolum, robot_no, istasyon, idempotency_key, kaynak_tip }
    """
    if not _yetki_kontrol():
        return jsonify({'hata': 'Yetkisiz'}), 401

    d = request.get_json() or {}
    cihaz_id = (d.get('cihaz_id') or '').strip()
    robot_no = (d.get('robot_no') or '').strip()
    idem_key = (d.get('idempotency_key') or '').strip()
    if not cihaz_id or not robot_no or not idem_key:
        return jsonify({'hata': 'cihaz_id, robot_no, idempotency_key zorunludur'}), 400

    bolum = (d.get('bolum') or 'kaynak').strip()
    istasyon = int(d.get('istasyon') or 0)
    kaynak_tip = (d.get('kaynak_tip') or 'robot_io').strip()

    conn = get_db()
    try:
        # idempotency_key UNIQUE — duplicate pulse otomatik reddedilir
        conn.execute('''
            INSERT OR IGNORE INTO sayac_olaylari (cihaz_id, bolum, robot_no, istasyon, kaynak_tip, idempotency_key)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (cihaz_id, bolum, robot_no, istasyon, kaynak_tip, idem_key))
        # Cihaz son_sinyal zamanını güncelle
        conn.execute('''
            INSERT INTO cihaz_kayitlari (cihaz_id, bolum, robot_no, son_sinyal, toplam_sinyal)
            VALUES (?, ?, ?, datetime('now','localtime'), 1)
            ON CONFLICT(cihaz_id) DO UPDATE SET
                son_sinyal = excluded.son_sinyal,
                toplam_sinyal = toplam_sinyal + 1,
                bolum = excluded.bolum,
                robot_no = excluded.robot_no
        ''', (cihaz_id, bolum, robot_no))
        conn.commit()
    finally:
        conn.close()

    return jsonify({'basarili': True, 'idempotency_key': idem_key}), 201


@app.route('/api/sinyal/heartbeat', methods=['POST'])
def heartbeat():
    """ESP32 30sn'de bir burayı çağırır — cihaz sağlığı."""
    if not _yetki_kontrol():
        return jsonify({'hata': 'Yetkisiz'}), 401

    d = request.get_json() or {}
    cihaz_id = (d.get('cihaz_id') or '').strip()
    if not cihaz_id:
        return jsonify({'hata': 'cihaz_id zorunludur'}), 400

    yeni_robot_calisiyor = 1 if d.get('robot_calisiyor') else 0

    conn = get_db()
    try:
        # Önceki robot durumunu + pulse baseline'larını + sağlık metriklerini öğren
        mevcut = conn.execute(
            "SELECT robot_calisiyor, pulse_ist1, pulse_ist2, base_ist1, base_ist2, "
            "uptime_sn, radyo_yenile FROM cihaz_kayitlari WHERE cihaz_id = ?",
            (cihaz_id,)
        ).fetchone()
        eski_durum = (mevcut['robot_calisiyor'] if mevcut else None)
        durum_degisti = (eski_durum is None) or (eski_durum != yeni_robot_calisiyor)

        # Sağlık/self-heal metrikleri (firmware v2.6+) — yoksa 0
        yeni_radyo    = int(d.get('radyo_yenile') or 0)
        yeni_httpfail = int(d.get('http_fail') or 0)
        yeni_disc     = int(d.get('disconnect') or 0)
        yeni_uptime   = int(d.get('uptime_sn') or 0)
        yeni_buffer   = int(d.get('buffer_kuyruk') or 0)

        # ─── Pulse mutabakatı: cihazın NVS sayımı vs server'a ulaşan ───
        # Cihaz pulse_ist'i NVS'te tutar (reboot'a dayanıklı). Server'a POST'lanan
        # sayım WiFi kopması/reset'te kaybolabilir. İlk görüşte baseline kur:
        #   base = server_count - pulse_ist  (o an kayıp 0 varsayımı)
        # Sonraki heartbeat'lerde okuma anında: kayip = max(0, pulse_ist + base - server_count)
        yeni_p1 = int(d.get('pulse_ist1') or 0)
        yeni_p2 = int(d.get('pulse_ist2') or 0)
        base_ist1 = (mevcut['base_ist1'] if mevcut else None)
        base_ist2 = (mevcut['base_ist2'] if mevcut else None)
        try:
            eski_p1 = (mevcut['pulse_ist1'] if mevcut else 0) or 0
            eski_p2 = (mevcut['pulse_ist2'] if mevcut else 0) or 0
            # base henüz yoksa ya da cihaz NVS sıfırlandıysa (pulse geriledi) → yeniden kur
            if base_ist1 is None or yeni_p1 < eski_p1:
                sc1 = conn.execute("SELECT COUNT(*) FROM sayac_olaylari WHERE cihaz_id=? AND istasyon=1", (cihaz_id,)).fetchone()[0]
                base_ist1 = sc1 - yeni_p1
            if base_ist2 is None or yeni_p2 < eski_p2:
                sc2 = conn.execute("SELECT COUNT(*) FROM sayac_olaylari WHERE cihaz_id=? AND istasyon=2", (cihaz_id,)).fetchone()[0]
                base_ist2 = sc2 - yeni_p2
        except Exception:
            pass  # mutabakat hatası heartbeat'i düşürmesin

        conn.execute('''
            INSERT INTO cihaz_kayitlari (cihaz_id, bolum, robot_no, firmware_ver, ip_adresi, mac_adresi,
                                          wifi_rssi, buffer_kuyruk, uptime_sn, free_heap,
                                          isr_count_ist1, isr_count_ist2,
                                          tani_sayildi, tani_parazit, tani_erken,
                                          pulse_ist1, pulse_ist2, base_ist1, base_ist2,
                                          radyo_yenile, http_fail, disconnect_say,
                                          robot_calisiyor, robot_durum_zamani, son_heartbeat)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now','localtime'), datetime('now','localtime'))
            ON CONFLICT(cihaz_id) DO UPDATE SET
                bolum = excluded.bolum,
                robot_no = excluded.robot_no,
                firmware_ver = excluded.firmware_ver,
                ip_adresi = excluded.ip_adresi,
                mac_adresi = excluded.mac_adresi,
                wifi_rssi = excluded.wifi_rssi,
                buffer_kuyruk = excluded.buffer_kuyruk,
                uptime_sn = excluded.uptime_sn,
                free_heap = excluded.free_heap,
                isr_count_ist1 = excluded.isr_count_ist1,
                isr_count_ist2 = excluded.isr_count_ist2,
                tani_sayildi = excluded.tani_sayildi,
                tani_parazit = excluded.tani_parazit,
                tani_erken = excluded.tani_erken,
                pulse_ist1 = excluded.pulse_ist1,
                pulse_ist2 = excluded.pulse_ist2,
                base_ist1 = excluded.base_ist1,
                base_ist2 = excluded.base_ist2,
                radyo_yenile = excluded.radyo_yenile,
                http_fail = excluded.http_fail,
                disconnect_say = excluded.disconnect_say,
                robot_calisiyor = excluded.robot_calisiyor,
                robot_durum_zamani = CASE
                    WHEN ? = 1 THEN datetime('now','localtime')
                    ELSE robot_durum_zamani
                END,
                son_heartbeat = excluded.son_heartbeat
        ''', (
            cihaz_id,
            (d.get('bolum') or '').strip(),
            (d.get('robot_no') or '').strip(),
            (d.get('firmware_ver') or '').strip(),
            (d.get('ip_adresi') or '').strip(),
            (d.get('mac_adresi') or '').strip(),
            int(d.get('wifi_rssi') or 0),
            yeni_buffer,
            yeni_uptime,
            int(d.get('free_heap') or 0),
            int(d.get('isr_count_ist1') or 0),
            int(d.get('isr_count_ist2') or 0),
            int(d.get('tani_sayildi') or 0),
            int(d.get('tani_parazit') or 0),
            int(d.get('tani_erken') or 0),
            yeni_p1, yeni_p2, base_ist1, base_ist2,
            yeni_radyo, yeni_httpfail, yeni_disc,
            yeni_robot_calisiyor,
            1 if durum_degisti else 0,
        ))

        # ─── Sağlık zaman serisi: ANOMALİ veya ~10dk baseline'ı KALICI logla ───
        # Amaç: "iyi RSSI ama offline" tipi olayların post-mortem'i (cihaz_kayitlari tek
        # satır olduğu için geçmiş eziliyordu). Sadece ilginç durumda satır açar → tablo şişmez.
        try:
            eski_uptime = (mevcut['uptime_sn'] if mevcut else None)
            eski_radyo  = ((mevcut['radyo_yenile'] if mevcut else 0) or 0)
            sebep = ''
            if eski_uptime is None or yeni_uptime < (eski_uptime or 0):
                sebep = 'reboot'        # uptime geriledi/ilk görüş = cihaz yeniden başladı
            elif yeni_radyo > eski_radyo:
                sebep = 'radyo'         # radyo tazeleme yapıldı (kurtarma denemesi)
            elif yeni_httpfail >= 3:
                sebep = 'httpfail'      # sunucuya ulaşamıyor (RSSI iyiyse zombie adayı)
            elif yeni_buffer > 0:
                sebep = 'buffer'        # kuyrukta gönderilemeyen pulse var
            else:
                son = conn.execute("SELECT MAX(ts) FROM saglik_olaylari WHERE cihaz_id=?",
                                   (cihaz_id,)).fetchone()[0]
                if not son:
                    sebep = 'baseline'
                else:
                    fark_dk = conn.execute(
                        "SELECT (julianday('now','localtime') - julianday(?)) * 1440", (son,)
                    ).fetchone()[0]
                    if fark_dk is None or fark_dk > 10:
                        sebep = 'baseline'
            if sebep:
                conn.execute('''
                    INSERT INTO saglik_olaylari
                        (cihaz_id, robot_no, sebep, wifi_rssi, free_heap, uptime_sn,
                         buffer_kuyruk, radyo_yenile, http_fail, disconnect_say)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (cihaz_id, (d.get('robot_no') or '').strip(), sebep,
                      int(d.get('wifi_rssi') or 0), int(d.get('free_heap') or 0), yeni_uptime,
                      yeni_buffer, yeni_radyo, yeni_httpfail, yeni_disc))
                # Retention: 30 günden eski sağlık kayıtlarını buda (tablo şişmesin)
                conn.execute("DELETE FROM saglik_olaylari WHERE ts < datetime('now','localtime','-30 days')")
        except Exception:
            pass  # sağlık logu heartbeat'i ASLA düşürmesin

        # Tani olay dizisi (firmware v2.4+) — varsa her birini idempotent ekle
        robot_no_hb = (d.get('robot_no') or '').strip()
        TIP_AD = {0: 'SAYILDI', 1: 'PARAZIT', 2: 'ERKEN'}
        for ev in (d.get('tani') or []):
            try:
                # NOT: 0 gecerli deger (tip 0=SAYILDI, gap/low 0) → `or` KULLANMA, .get(k, default)
                dev_seq = int(ev.get('s', 0))
                idem = f"{cihaz_id}_{int(ev.get('b', 0))}_{dev_seq}"
                conn.execute('''
                    INSERT OR IGNORE INTO tani_olaylari
                        (cihaz_id, robot_no, uptime_ms, tip, gap_ms, low_ms, dev_seq, idem)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    cihaz_id, robot_no_hb,
                    int(ev.get('u', 0)),
                    TIP_AD.get(int(ev.get('t', -1)), 'BILINMEYEN'),
                    int(ev.get('g', 0)),
                    int(ev.get('l', 0)),
                    dev_seq, idem,
                ))
            except (ValueError, TypeError):
                continue  # bozuk olay kaydini atla, heartbeat'i dusurme

        conn.commit()
    finally:
        conn.close()
    return jsonify({'basarili': True}), 200


@app.route('/api/sinyal/canli', methods=['GET'])
def canli_sayac():
    """Bugünün toplam pulse sayısı + robot/istasyon kırılımı."""
    bugun = date.today().isoformat()
    bolum = request.args.get('bolum', 'kaynak')
    conn = get_db()
    rows = conn.execute('''
        SELECT robot_no, istasyon, COUNT(*) as adet,
               MIN(ts) as ilk_pulse,
               MAX(ts) as son_pulse
        FROM sayac_olaylari
        WHERE date(ts) = ? AND bolum = ?
        GROUP BY robot_no, istasyon
        ORDER BY robot_no, istasyon
    ''', (bugun, bolum)).fetchall()

    toplam = sum(r['adet'] for r in rows)

    # Son 60 dakika hızı (parça/saat)
    son60 = conn.execute('''
        SELECT robot_no, istasyon, COUNT(*) as adet
        FROM sayac_olaylari
        WHERE bolum = ? AND ts >= datetime('now','localtime','-60 minutes')
        GROUP BY robot_no, istasyon
    ''', (bolum,)).fetchall()
    son60_map = {f"{r['robot_no']}_{r['istasyon']}": r['adet'] for r in son60}

    sonuc = []
    for r in rows:
        anahtar = f"{r['robot_no']}_{r['istasyon']}"
        sonuc.append({
            'robot_no': r['robot_no'],
            'istasyon': r['istasyon'],
            'adet': r['adet'],
            'son60_dk': son60_map.get(anahtar, 0),
            'ilk_pulse': r['ilk_pulse'],
            'son_pulse': r['son_pulse']
        })

    conn.close()
    return jsonify({'bolum': bolum, 'tarih': bugun, 'toplam': toplam, 'kirilim': sonuc})


@app.route('/api/saha_cihazlari', methods=['GET'])
def saha_cihazlari():
    """Tüm beklenen saha cihazlarının durumu + bugün ist1/ist2 sayımları.

    Aynı mantıkla app.py'da da var (dashboard için). Burada pilot UI
    (port 5001) doğrudan çağırabilsin diye duplike edildi.
    """
    # NOT: bu liste app.py'daki BEKLENEN'in ESKİ/DAR bir kopyasıdır (pres, plastik ve
    # TK1 cihazları burada yok). Dashboard app.py'ı kullanır; burası saha kurulumunda
    # bakılan pilot UI'sı. 'tel' 2026-08-07'de eklendi — kapama modülleri sahaya
    # takılırken bu ekrandan kontrol edilecek.
    BEKLENEN = {
        'kaynak': [{'cihaz_id': f'ABB{i}-IO', 'robot_no': f'ABB{i}'} for i in range(1, 10)],
        'montaj': [{'cihaz_id': f'MONTAJ-M{i}', 'robot_no': f'M{i}'} for i in range(1, 13)],
        'metal':  [
            {'cihaz_id': '300T-IO', 'robot_no': '300T'},
            {'cihaz_id': '400T-IO', 'robot_no': '400T'},
            {'cihaz_id': '550T-IO', 'robot_no': '550T'},
        ],
        # TK1 kapama presleri: 1 modül = 3 pres (GPIO25/26/27 -> istasyon 1/2/3).
        # Burada satır MODÜL bazlıdır (cihaz sağlığı ekranı); pres bazlı kırılım
        # bugun_ist1/2/3 alanlarında. Makine bazlı kart app.py tarafındadır.
        'tel': [
            {'cihaz_id': f'TEL-KAPAMA-{m}',
             'robot_no': f'Kapama {(m - 1) * 3 + 1}-{(m - 1) * 3 + 3}'}
            for m in range(1, 5)
        ],
    }

    conn = get_db()
    mevcut = {}
    ist_sayimlari = {}

    rows = conn.execute('''
        SELECT *,
               CAST((julianday('now','localtime') - julianday(son_heartbeat)) * 1440 AS INTEGER) as son_heartbeat_dk
        FROM cihaz_kayitlari
    ''').fetchall()
    for r in rows:
        d = dict(r)
        d['durum'] = 'offline' if (d.get('son_heartbeat_dk') or 0) > 2 else 'online'
        mevcut[d['cihaz_id']] = d

    for r in conn.execute('''
        SELECT cihaz_id, istasyon, COUNT(*) as adet
        FROM sayac_olaylari
        WHERE date(ts) = date('now','localtime')
        GROUP BY cihaz_id, istasyon
    ''').fetchall():
        cid = r['cihaz_id']
        if cid not in ist_sayimlari: ist_sayimlari[cid] = {}
        ist_sayimlari[cid][r['istasyon']] = r['adet']

    conn.close()

    sonuc = {}
    for bolum, cihazlar in BEKLENEN.items():
        liste = []
        for beklenen_cihaz in cihazlar:
            cid = beklenen_cihaz['cihaz_id']
            kayit = mevcut.get(cid)
            if kayit:
                ist_map = ist_sayimlari.get(cid, {})
                liste.append({
                    'cihaz_id':         cid,
                    'robot_no':         beklenen_cihaz['robot_no'],
                    'bolum':            bolum,
                    'durum':            kayit.get('durum', 'offline'),
                    'ip_adresi':        kayit.get('ip_adresi', ''),
                    'wifi_rssi':        kayit.get('wifi_rssi', 0),
                    'son_heartbeat':    kayit.get('son_heartbeat', ''),
                    'son_heartbeat_dk': kayit.get('son_heartbeat_dk', 0),
                    'firmware_ver':     kayit.get('firmware_ver', ''),
                    'toplam_sinyal':    kayit.get('toplam_sinyal', 0),
                    'buffer_kuyruk':    kayit.get('buffer_kuyruk', 0),
                    'bugun_ist1':       ist_map.get(1, 0),
                    'bugun_ist2':       ist_map.get(2, 0),
                    'bugun_ist3':       ist_map.get(3, 0),   # 3 kanalli modul (kapama presleri)
                    'bugun_toplam':     sum(ist_map.values()),
                    'robot_calisiyor':  1 if kayit.get('robot_calisiyor') else 0,
                    'kayitli':          True,
                })
            else:
                liste.append({
                    'cihaz_id':      cid,
                    'robot_no':      beklenen_cihaz['robot_no'],
                    'bolum':         bolum,
                    'durum':         'beklemede',
                    'bugun_ist1':    0,
                    'bugun_ist2':    0,
                    'bugun_ist3':    0,
                    'bugun_toplam':  0,
                    'robot_calisiyor': 0,
                    'kayitli':       False,
                })
        sonuc[bolum] = liste
    return jsonify(sonuc), 200


@app.route('/api/sinyal/cihazlar', methods=['GET'])
def cihaz_listesi():
    """Tüm ESP32 cihazlarının sağlık durumu."""
    conn = get_db()
    rows = conn.execute('''
        SELECT *,
               CAST((julianday('now','localtime') - julianday(son_heartbeat)) * 1440 AS INTEGER) as son_heartbeat_dk
        FROM cihaz_kayitlari
        ORDER BY cihaz_id
    ''').fetchall()
    conn.close()
    cihazlar = []
    for r in rows:
        d = dict(r)
        # 2 dakikadan fazla heartbeat almadıysak "OFFLINE"
        d['durum'] = 'offline' if (d.get('son_heartbeat_dk') or 0) > 2 else 'online'
        cihazlar.append(d)
    return jsonify(cihazlar)


@app.route('/api/sinyal/son', methods=['GET'])
def son_sinyaller():
    """Son N sayaç olayı (debug/canlı log için)."""
    n = min(int(request.args.get('limit', 50)), 200)
    conn = get_db()
    rows = conn.execute('''
        SELECT * FROM sayac_olaylari
        ORDER BY id DESC LIMIT ?
    ''', (n,)).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/api/sinyal/mutabakat', methods=['GET'])
def mutabakat():
    """Pulse mutabakatı — cihazın NVS sayımı (pulse_ist) vs server'a ulaşan sayım.
    kayip > 0 → o cihazda WiFi/reset kaynaklı transit kaybı var (artıyorsa anten/bağlantı bak).
    Anten düzeldikten sonra kayip ~0'da kalmalı.
    """
    conn = get_db()
    rows = conn.execute(
        "SELECT cihaz_id, bolum, robot_no, pulse_ist1, pulse_ist2, base_ist1, base_ist2, wifi_rssi "
        "FROM cihaz_kayitlari ORDER BY cihaz_id"
    ).fetchall()
    out = []
    for r in rows:
        cid = r['cihaz_id']
        sc1 = conn.execute("SELECT COUNT(*) FROM sayac_olaylari WHERE cihaz_id=? AND istasyon=1", (cid,)).fetchone()[0]
        sc2 = conn.execute("SELECT COUNT(*) FROM sayac_olaylari WHERE cihaz_id=? AND istasyon=2", (cid,)).fetchone()[0]
        b1 = r['base_ist1'] if r['base_ist1'] is not None else 0
        b2 = r['base_ist2'] if r['base_ist2'] is not None else 0
        kayip1 = max(0, (r['pulse_ist1'] or 0) + b1 - sc1)
        kayip2 = max(0, (r['pulse_ist2'] or 0) + b2 - sc2)
        out.append({
            'cihaz_id': cid, 'bolum': r['bolum'], 'robot_no': r['robot_no'], 'wifi_rssi': r['wifi_rssi'],
            'cihaz_sayim_ist1': r['pulse_ist1'] or 0, 'server_sayim_ist1': sc1, 'kayip_ist1': kayip1,
            'cihaz_sayim_ist2': r['pulse_ist2'] or 0, 'server_sayim_ist2': sc2, 'kayip_ist2': kayip2,
            'baseline_kuruldu': r['base_ist1'] is not None,
        })
    conn.close()
    return jsonify(out)


@app.route('/api/sinyal/tani', methods=['GET'])
def tani_oku():
    """Sayim filtresi tani olaylari — eksik/fazla sayma teshisi.

    ?cihaz_id=400T-IO&limit=200  → o cihazin son N olayi + kumulatif sayaclar.
    Ozet: parazit cok → opto zayif; erken cok → cycle/gap; SAYILDI low_ms ~2x → role
    ardisik parcalari birlestiriyor (eksik sayim).
    """
    cihaz_id = (request.args.get('cihaz_id') or '').strip()
    n = min(int(request.args.get('limit', 200)), 1000)
    conn = get_db()

    sayaclar = {}
    if cihaz_id:
        r = conn.execute('''
            SELECT tani_sayildi, tani_parazit, tani_erken, firmware_ver, toplam_sinyal
            FROM cihaz_kayitlari WHERE cihaz_id = ?
        ''', (cihaz_id,)).fetchone()
        if r:
            sayaclar = dict(r)

    if cihaz_id:
        rows = conn.execute('''
            SELECT * FROM tani_olaylari WHERE cihaz_id = ?
            ORDER BY id DESC LIMIT ?
        ''', (cihaz_id, n)).fetchall()
    else:
        rows = conn.execute('''
            SELECT * FROM tani_olaylari ORDER BY id DESC LIMIT ?
        ''', (n,)).fetchall()
    conn.close()

    olaylar = [dict(r) for r in rows]
    sayildi_low = [o['low_ms'] for o in olaylar if o['tip'] == 'SAYILDI' and o['low_ms'] > 0]
    ozet = {
        'sayildi': sum(1 for o in olaylar if o['tip'] == 'SAYILDI'),
        'parazit': sum(1 for o in olaylar if o['tip'] == 'PARAZIT'),
        'erken':   sum(1 for o in olaylar if o['tip'] == 'ERKEN'),
        'low_ms_min': min(sayildi_low) if sayildi_low else 0,
        'low_ms_max': max(sayildi_low) if sayildi_low else 0,
        'low_ms_ort': round(sum(sayildi_low) / len(sayildi_low)) if sayildi_low else 0,
    }
    return jsonify({
        'cihaz_id': cihaz_id,
        'kumulatif_sayaclar': sayaclar,
        'pencere_ozeti': ozet,
        'olaylar': olaylar,
    })


@app.route('/api/sinyal/saglik', methods=['GET'])
def saglik_oku():
    """Cihaz sağlık zaman serisi — "iyi RSSI ama offline" tipi olayların post-mortem'i.

    ?cihaz_id=MONTAJ-M7&limit=200  → o cihazın son N sağlık olayı (reboot/radyo/buffer/
    httpfail/baseline). free_heap düşüşü, uptime sıfırlanması (reboot), radyo_yenile artışı
    ve http_fail birikimi burada görülür — cihaz_kayitlari tek-satır ezdiği için tek kaynak budur.
    """
    cihaz_id = (request.args.get('cihaz_id') or '').strip()
    n = min(int(request.args.get('limit', 200)), 2000)
    conn = get_db()
    if cihaz_id:
        rows = conn.execute(
            'SELECT * FROM saglik_olaylari WHERE cihaz_id = ? ORDER BY id DESC LIMIT ?',
            (cihaz_id, n)
        ).fetchall()
    else:
        rows = conn.execute(
            'SELECT * FROM saglik_olaylari ORDER BY id DESC LIMIT ?', (n,)
        ).fetchall()
    conn.close()
    return jsonify({'cihaz_id': cihaz_id, 'olaylar': [dict(r) for r in rows]})


@app.route('/logo')
def logo_serve():
    """Cofle Control Systems logosu — ana projedeki static/ dosyalarından
    (2026-07-21: resmi logo static/logo_koyu.png koyu-tema, logo.png orijinal)."""
    ana_static = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'static')
    yollar = [
        os.path.join(ana_static, 'logo_koyu.png'),   # pilot konsolu koyu temali
        os.path.join(ana_static, 'logo.png'),
        os.path.join(os.path.expanduser('~'), 'OneDrive', 'Masaüstü', 'LOGO_COFLE ONLY.png'),
    ]
    for y in yollar:
        if os.path.exists(y):
            return send_file(y, mimetype='image/png')
    return '', 404


# ─── Başlatma ─────────────────────────────────────────────────────
if __name__ == '__main__':
    init_db()
    print("\n" + "=" * 55)
    print("  COFLE PILOT SAYAÇ BACKEND")
    print("=" * 55)
    print(f"  Port      : {PORT}")
    print(f"  DB        : {DB_PATH}")
    print(f"  API token : {API_TOKEN}  (ESP32 firmware'inde aynı yazmalı)")
    print(f"  Canlı UI  : http://<bu-pc-ip>:{PORT}/")
    print(f"  ESP32 URL : http://<bu-pc-ip>:{PORT}/api/sinyal")
    print("=" * 55 + "\n")
    # GUVENLIK: debug=True = uzaktan kod calistirma riski. Varsayilan KAPALI;
    # yerel gelistirmede: COFLE_DEBUG=1
    app.run(host='0.0.0.0', port=PORT, debug=(os.environ.get('COFLE_DEBUG') == '1'))
