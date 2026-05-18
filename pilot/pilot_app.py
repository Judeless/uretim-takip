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
    ]:
        try:
            c.execute(f"ALTER TABLE cihaz_kayitlari ADD COLUMN {col} {ddl}")
        except Exception:
            pass  # zaten var
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
        # Önceki robot durumunu öğren — değiştiyse durum_zamani'nı güncelleyeceğiz
        mevcut = conn.execute(
            "SELECT robot_calisiyor FROM cihaz_kayitlari WHERE cihaz_id = ?",
            (cihaz_id,)
        ).fetchone()
        eski_durum = (mevcut['robot_calisiyor'] if mevcut else None)
        durum_degisti = (eski_durum is None) or (eski_durum != yeni_robot_calisiyor)

        conn.execute('''
            INSERT INTO cihaz_kayitlari (cihaz_id, bolum, robot_no, firmware_ver, ip_adresi, mac_adresi,
                                          wifi_rssi, buffer_kuyruk, uptime_sn, free_heap,
                                          robot_calisiyor, robot_durum_zamani, son_heartbeat)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now','localtime'), datetime('now','localtime'))
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
            int(d.get('buffer_kuyruk') or 0),
            int(d.get('uptime_sn') or 0),
            int(d.get('free_heap') or 0),
            yeni_robot_calisiyor,
            1 if durum_degisti else 0,
        ))
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
    BEKLENEN = {
        'kaynak': [{'cihaz_id': f'ABB{i}-IO', 'robot_no': f'ABB{i}'} for i in range(1, 10)],
        'montaj': [{'cihaz_id': f'MONTAJ-M{i}', 'robot_no': f'M{i}'} for i in range(1, 13)],
        'metal':  [
            {'cihaz_id': '300T-IO', 'robot_no': '300T'},
            {'cihaz_id': '400T-IO', 'robot_no': '400T'},
            {'cihaz_id': '550T-IO', 'robot_no': '550T'},
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


@app.route('/logo')
def logo_serve():
    """Cofle logosu — ana sistemdeki ile aynı yol."""
    yollar = [
        os.path.join(os.path.expanduser('~'), 'OneDrive', 'Masaüstü', 'LOGO_COFLE ONLY.png'),
        os.path.join(os.path.expanduser('~'), 'Desktop', 'LOGO_COFLE ONLY.png'),
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
    app.run(host='0.0.0.0', port=PORT, debug=True)
