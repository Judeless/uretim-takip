from flask import Flask, request, jsonify, render_template, send_file, g
from flask_cors import CORS
from datetime import datetime, date
import json
import os
import traceback # For debugging

from database import get_db as db_connect, init_db
from oee import hesapla_oee, hesapla_oee_ozet
from import_excel import import_data
from export_excel import export_arsiv

# ODS dosyası yolu
FIKSTUR_ODS_YOLU = r'C:\Users\selcu\OneDrive\Masaüstü\KAYNAKHANE FİKSTÜR RAF LİSTESİ.ods'
UYUM_ODS_YOLU = r'C:\Users\selcu\OneDrive\Masaüstü\RAF ve PROGRAM..ods'

def _ods_uyum_oku():
    """ODS dosyasındaki referans-robot uyum bilgilerini döndürür. (ref, robot_no, ist_no) listesi"""
    try:
        from odf.opendocument import load as ods_load
        from odf.table import Table, TableRow, TableCell
        from odf.text import P
        
        doc = ods_load(UYUM_ODS_YOLU)
        sh = doc.spreadsheet.getElementsByType(Table)[0]
        rows = sh.getElementsByType(TableRow)
        
        col_map = {}
        for r_idx in range(1, 15):
            base = 2 + (r_idx - 1) * 2
            col_map[base] = (f'ABB-{r_idx}', 1)
            col_map[base + 1] = (f'ABB-{r_idx}', 2)
            
        kayitlar = []
        for idx, row in enumerate(rows):
            if idx < 2: continue # basliklari atla
            cells = row.getElementsByType(TableCell)
            vals = []
            for c in cells:
                ps = c.getElementsByType(P)
                text = str(ps[0]).strip() if ps else ''
                rep = c.getAttribute('numbercolumnsrepeated')
                cnt = min(int(rep) if rep else 1, 35 - len(vals))
                vals.extend([text] * cnt)
                if len(vals) >= 35: break
            while len(vals) < 35: vals.append('')
            
            ref = vals[0].strip()
            if not ref or len(ref) < 3 or ref == 'ROBOT' or ref == 'İSTASYON': continue
            
            for col_idx, (r_no, ist) in col_map.items():
                if col_idx < len(vals):
                    val = vals[col_idx].strip()
                    if val and len(val) > 0:
                        kayitlar.append((ref, r_no, ist))
        return kayitlar
    except Exception as e:
        print('ODS Uyum Okuma Hatası:', e)
        return []

def _ods_kayitlari_oku():
    """ODS dosyasından tüm fikstür kayıtlarını {referans_kodu: raf_adresi} dict olarak döndürür."""
    try:
        from odf.opendocument import load as ods_load
        from odf.table import Table, TableRow, TableCell
        from odf.text import P

        doc = ods_load(FIKSTUR_ODS_YOLU)
        sh = doc.spreadsheet.getElementsByType(Table)[0]
        rows = sh.getElementsByType(TableRow)
        kayitlar = {}
        for row in rows:
            cells = row.getElementsByType(TableCell)
            vals = []
            for c in cells:
                ps = c.getElementsByType(P)
                text = str(ps[0]).strip() if ps else ''
                rep = c.getAttribute('numbercolumnsrepeated')
                cnt = min(int(rep) if rep else 1, 12 - len(vals))
                vals.extend([text] * cnt)
                if len(vals) >= 12:
                    break
            while len(vals) < 12:
                vals.append('')
            # Her 3 sütun grubunda 0=ref, 1=raf
            skip = {'A RAFI','B RAFI','C RAFI','D RAFI','REFERANS','RAF',''}
            for start in [0, 3, 6, 9]:
                ref = vals[start].strip()
                raf = vals[start + 1].strip() if start + 1 < 12 else ''
                if ref and ref not in skip and len(ref) > 2:
                    kayitlar[ref] = raf
        return kayitlar
    except Exception:
        return {}


def _ods_guncelle(tum_kayitlar):
    """
    tum_kayitlar: list of (referans_kodu, raf_adresi, notlar) tuples
    ODS dosyasını tamamen yeniden yazar.
    """
    try:
        from odf.opendocument import OpenDocumentSpreadsheet
        from odf.table import Table, TableRow, TableCell
        from odf.text import P
        from odf.style import Style, TextProperties, TableColumnProperties

        doc = OpenDocumentSpreadsheet()
        sheet = Table(name='Fikstür Listesi')
        doc.spreadsheet.addElement(sheet)

        # Başlık satırı
        hrow = TableRow()
        sheet.addElement(hrow)
        for baslik in ['Referans Kodu', 'Raf Adresi', 'Notlar']:
            tc = TableCell(valuetype='string')
            tc.addElement(P(text=baslik))
            hrow.addElement(tc)

        # Veri satırları
        for ref, raf, notlar in tum_kayitlar:
            row = TableRow()
            sheet.addElement(row)
            for val in [ref, raf, notlar]:
                tc = TableCell(valuetype='string')
                tc.addElement(P(text=str(val) if val else ''))
                row.addElement(tc)

        doc.save(FIKSTUR_ODS_YOLU)
    except Exception as e:
        print(f'ODS güncelleme hatası: {e}')

app = Flask(__name__)
CORS(app)

# ── Veritabanı Bağlantı Yönetimi ──
def get_db():
    if 'db' not in g:
        g.db = db_connect()
    return g.db

@app.teardown_appcontext
def teardown_db(exception):
    db = g.pop('db', None)
    if db is not None:
        try:
            db.close()
        except Exception:
            pass

# ─────────────────────────────────────────────────────────────
# SAYFA ROUTES
# ─────────────────────────────────────────────────────────────

@app.route('/')
def operator_sayfasi():
    """Operatör mobil giriş sayfası (v2 — yeni tasarım canlıya alındı)."""
    return render_template('mobile_v2.html')


@app.route('/mobile_v2')
def operator_v2_preview():
    """Operatör v2 önizlemesi (canlı ile aynı şablon — direkt link için saklanıyor)."""
    return render_template('mobile_v2.html')


@app.route('/mobile_legacy')
def operator_legacy_sayfasi():
    """Eski operatör tasarımı (geri dönüş için saklanıyor)."""
    return render_template('mobile.html')


@app.route('/dashboard')
def dashboard_sayfasi():
    """Yönetici dashboard sayfası (yeni neon temada)."""
    return render_template('dashboard.html')


@app.route('/dashboard_legacy')
def dashboard_legacy_sayfasi():
    """Eski dashboard tasarımı (geri dönüş için saklanıyor)."""
    return render_template('dashboard_legacy.html')


@app.route('/logo')
def logo_serve():
    """Şirket logosunu masaüstünden serve et."""
    logo_yollar = [
        os.path.join(os.path.expanduser('~'), 'OneDrive', 'Masaüstü', 'LOGO_COFLE ONLY.png'),
        os.path.join(os.path.expanduser('~'), 'Desktop', 'LOGO_COFLE ONLY.png'),
        os.path.join(os.path.dirname(__file__), 'static', 'logo.png'),
    ]
    for yol in logo_yollar:
        if os.path.exists(yol):
            return send_file(yol, mimetype='image/png')
    return '', 404


# ───── Favicon / iOS / Android home-screen ikonları ─────
# Telefonda "Ana ekrana ekle" denildiğinde Cofle logosu görünür.
@app.route('/favicon.ico')
@app.route('/favicon.png')
@app.route('/apple-touch-icon.png')
@app.route('/apple-touch-icon-precomposed.png')
def favicon_serve():
    """Tüm tarayıcı/iOS/Android ikon istekleri /logo'ya yönlendirilir."""
    return logo_serve()


@app.route('/manifest.json')
def web_manifest():
    """Android PWA manifesti — ana ekrana eklendiğinde Cofle logosu + neon tema."""
    return jsonify({
        "name": "Cofle Manage",
        "short_name": "Cofle",
        "description": "Cofle Manage Üretim Takip Sistemi",
        "start_url": "/",
        "display": "standalone",
        "orientation": "any",
        "background_color": "#060414",
        "theme_color": "#1c1f3a",
        "lang": "tr",
        "icons": [
            {"src": "/logo", "sizes": "192x192", "type": "image/png", "purpose": "any"},
            {"src": "/logo", "sizes": "512x512", "type": "image/png", "purpose": "any"},
            {"src": "/logo", "sizes": "180x180", "type": "image/png", "purpose": "maskable"}
        ]
    })

@app.route('/andon')
def andon_sayfasi():
    """Andon TV pano sayfasi — Robot Kaynak (v5)."""
    return render_template('andon_v5.html',
                           bolum='kaynak',
                           bolum_ad='Robot Kaynak',
                           bolum_ikon='🔧',
                           panel_baslik='Robot Durum Paneli')


@app.route('/andon_legacy')
def andon_legacy_sayfasi():
    """Eski andon tasarımı (geri dönüş için saklanıyor)."""
    return render_template('andon.html')

@app.route('/andon_v2')
def andon_v2_preview():
    """Andon V2 onizleme sayfasi."""
    return render_template('andon_v2_preview.html')


@app.route('/rapor')
def rapor_sayfasi():
    """Operatörler için günlük üretim raporu paylaşım sayfası."""
    return render_template('rapor.html')


@app.route('/andon_v4')
def andon_v4_preview():
    """Andon v4 Tema Önizleme."""
    return render_template('andon_v4.html')


@app.route('/andon_v5')
def andon_v5_preview():
    """Andon v5 — komuta merkezi tema önizlemesi (Robot Kaynak)."""
    return render_template('andon_v5.html')


@app.route('/andon_montaj')
def andon_montaj_sayfasi():
    """Montaj Andon ekranı (v5 tasarım)."""
    return render_template('andon_v5.html',
                           bolum='montaj',
                           bolum_ad='Montaj',
                           bolum_ikon='🔩',
                           panel_baslik='Hat Durum Paneli')


@app.route('/andon_metal')
def andon_metal_sayfasi():
    """Metal Enjeksiyon Andon ekranı (v5 tasarım)."""
    return render_template('andon_v5.html',
                           bolum='metal',
                           bolum_ad='Metal Enjeksiyon',
                           bolum_ikon='🏭',
                           panel_baslik='Makine Durum Paneli')


@app.route('/andon_montaj_legacy')
def andon_montaj_legacy_sayfasi():
    """Eski Montaj andon tasarımı (geri dönüş için saklanıyor)."""
    return render_template('andon_montaj.html')


@app.route('/andon_metal_legacy')
def andon_metal_legacy_sayfasi():
    """Eski Metal Enjeksiyon andon tasarımı (geri dönüş için saklanıyor)."""
    return render_template('andon_metal.html')


@app.route('/dashboard_v3')
def dashboard_v3_preview():
    """Dashboard v3 Tema Önizleme."""
    return render_template('dashboard_v3.html')


# ─────────────────────────────────────────────────────────────
# VARDİYA API
# ─────────────────────────────────────────────────────────────

@app.route('/api/vardiya', methods=['GET'])
def vardiya_listesi():
    """Vardiya listesini döndür (filtrelenebilir)."""
    conn = get_db()
    c = conn.cursor()

    tarih = request.args.get('tarih')
    robot = request.args.get('robot')
    vardiya_turu = request.args.get('vardiya_turu')
    bolum = request.args.get('bolum')
    limit = int(request.args.get('limit', 100))

    query = 'SELECT * FROM vardiyalar WHERE 1=1'
    params = []

    if tarih:
        query += ' AND tarih = ?'
        params.append(tarih)
    if robot:
        query += ' AND robot_no = ?'
        params.append(robot)
    if vardiya_turu:
        query += ' AND vardiya_turu = ?'
        params.append(vardiya_turu)
    if bolum:
        query += " AND COALESCE(bolum, 'kaynak') = ?"
        params.append(bolum)

    query += ' ORDER BY tarih DESC, created_at DESC LIMIT ?'
    params.append(limit)

    rows = c.execute(query, params).fetchall()
    conn.close()

    return jsonify([dict(r) for r in rows])


@app.route('/api/vardiya', methods=['POST'])
def vardiya_ekle():
    """Yeni vardiya kaydı oluştur."""
    data = request.get_json()

    zorunlu = ['tarih', 'vardiya_turu', 'robot_no', 'operator_adi', 'baslangic_saati', 'bitis_saati']
    for alan in zorunlu:
        if not data.get(alan):
            return jsonify({'hata': f'"{alan}" alanı zorunludur'}), 400

    # Planlı süreyi saat farkından hesapla
    try:
        bas = datetime.strptime(data['baslangic_saati'], '%H:%M')
        bit = datetime.strptime(data['bitis_saati'], '%H:%M')
        fark_dk = int((bit - bas).total_seconds() / 60)
        if fark_dk < 0:
            fark_dk += 1440  # Gece yarısı geçen vardiyalar için
    except ValueError:
        return jsonify({'hata': 'Saat formatı HH:MM olmalıdır'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute('''
        INSERT INTO vardiyalar (tarih, vardiya_turu, robot_no, operator_adi, baslangic_saati, bitis_saati, toplam_sure_dk, notlar, bolum)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        data['tarih'],
        data['vardiya_turu'],
        data['robot_no'],
        data['operator_adi'],
        data['baslangic_saati'],
        data['bitis_saati'],
        fark_dk,
        data.get('notlar', ''),
        data.get('bolum', 'kaynak')
    ))
    vardiya_id = c.lastrowid
    conn.commit()
    conn.close()

    return jsonify({'basarili': True, 'vardiya_id': vardiya_id, 'sure_dk': fark_dk}), 201


@app.route('/api/vardiya/<int:vid>', methods=['GET'])
def vardiya_detay(vid):
    """Tek vardiya detayı (üretim + duruşlar dahil)."""
    conn = get_db()
    c = conn.cursor()

    vardiya = c.execute('SELECT * FROM vardiyalar WHERE id = ?', (vid,)).fetchone()
    if not vardiya:
        conn.close()
        return jsonify({'hata': 'Vardiya bulunamadı'}), 404

    uretim = c.execute('SELECT * FROM uretim_kayitlari WHERE vardiya_id = ?', (vid,)).fetchall()
    duruslar = c.execute('SELECT * FROM duruslar WHERE vardiya_id = ?', (vid,)).fetchall()
    conn.close()

    return jsonify({
        'vardiya': dict(vardiya),
        'uretim': [dict(r) for r in uretim],
        'duruslar': [dict(r) for r in duruslar]
    })


@app.route('/api/vardiya/<int:vid>', methods=['DELETE'])
def vardiya_sil(vid):
    """Vardiyayi sil (uretim ve duruslarla birlikte)."""
    conn = get_db()
    conn.execute('DELETE FROM vardiyalar WHERE id = ?', (vid,))
    conn.commit()
    conn.close()
    return jsonify({'basarili': True})


@app.route('/api/vardiya/bugun', methods=['GET'])
def vardiya_bugun():
    """Bugune ait tum vardiylari getirir (operatorn secim ekrani icin). ?bolum= ile filtrelenebilir."""
    from datetime import date as _date
    bugun = _date.today().isoformat()
    bolum = request.args.get('bolum')
    conn = get_db()
    if bolum:
        rows = conn.execute(
            "SELECT * FROM vardiyalar WHERE tarih = ? AND COALESCE(bolum, 'kaynak') = ? ORDER BY baslangic_saati DESC",
            (bugun, bolum)
        ).fetchall()
    else:
        rows = conn.execute(
            'SELECT * FROM vardiyalar WHERE tarih = ? ORDER BY baslangic_saati DESC',
            (bugun,)
        ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/api/vardiya/<int:vid>/kapat', methods=['PATCH'])
def vardiya_kapat(vid):
    """Vardiyayi kapat ve bitis saatini guncelle."""
    data = request.get_json() or {}
    bitis = data.get('bitis_saati', datetime.now().strftime('%H:%M'))
    conn = get_db()
    c = conn.cursor()
    vardiya = c.execute('SELECT * FROM vardiyalar WHERE id = ?', (vid,)).fetchone()
    if not vardiya:
        conn.close()
        return jsonify({'hata': 'Vardiya bulunamadi'}), 404
    bas = datetime.strptime(vardiya['baslangic_saati'], '%H:%M')
    bit = datetime.strptime(bitis, '%H:%M')
    fark_dk = int((bit - bas).total_seconds() / 60)
    if fark_dk < 0:
        fark_dk += 1440
    c.execute(
        "UPDATE vardiyalar SET durum='kapali', bitis_saati=?, toplam_sure_dk=? WHERE id=?",
        (bitis, fark_dk, vid)
    )
    conn.commit()
    conn.close()
    return jsonify({'basarili': True, 'sure_dk': fark_dk})


@app.route('/api/vardiya/<int:vid>/ac', methods=['PATCH'])
def vardiya_ac(vid):
    """Kapanmış vardiyayı yeniden aç."""
    conn = get_db()
    c = conn.cursor()
    vardiya = c.execute('SELECT * FROM vardiyalar WHERE id = ?', (vid,)).fetchone()
    if not vardiya:
        conn.close()
        return jsonify({'hata': 'Vardiya bulunamadi'}), 404
    if vardiya['durum'] != 'kapali':
        conn.close()
        return jsonify({'hata': 'Vardiya zaten açık'}), 400
    c.execute(
        "UPDATE vardiyalar SET durum='aktif', bitis_saati=NULL, toplam_sure_dk=NULL WHERE id=?",
        (vid,)
    )
    conn.commit()
    conn.close()
    return jsonify({'basarili': True})


@app.route('/api/vardiya/<int:vid>', methods=['PUT'])
def vardiya_guncelle(vid):
    """Vardiya saatlerini ve detaylarını güncelle."""
    data = request.get_json()
    conn = get_db()
    c = conn.cursor()
    mevcut = c.execute('SELECT * FROM vardiyalar WHERE id = ?', (vid,)).fetchone()
    if not mevcut:
        conn.close()
        return jsonify({'hata': 'Vardiya bulunamadi'}), 404
        
    bas = data.get('baslangic_saati', mevcut['baslangic_saati'])
    bit = data.get('bitis_saati', mevcut['bitis_saati'])
    fark_dk = data.get('toplam_sure_dk')
    operator_adi = data.get('operator_adi', mevcut['operator_adi'])
    tarih = data.get('tarih', mevcut['tarih'])
    vardiya_turu = data.get('vardiya_turu', mevcut['vardiya_turu'])
    
    if fark_dk is None:
        try:
            b1 = datetime.strptime(bas, '%H:%M')
            b2 = datetime.strptime(bit, '%H:%M')
            fark_dk = int((b2 - b1).total_seconds() / 60)
            if fark_dk < 0:
                fark_dk += 1440
        except Exception:
            fark_dk = mevcut['toplam_sure_dk']
            
    c.execute(
        "UPDATE vardiyalar SET baslangic_saati=?, bitis_saati=?, toplam_sure_dk=?, operator_adi=?, tarih=?, vardiya_turu=? WHERE id=?",
        (bas, bit, fark_dk, operator_adi, tarih, vardiya_turu, vid)
    )
    conn.commit()
    conn.close()
    return jsonify({'basarili': True, 'sure_dk': fark_dk})



# ─────────────────────────────────────────────────────────────
# ÜRETİM API
# ─────────────────────────────────────────────────────────────

@app.route('/api/uretim', methods=['POST'])
def uretim_ekle():
    """Üretim kaydı ekle."""
    data = request.get_json()

    if not data.get('vardiya_id') or not data.get('referans_kodu'):
        return jsonify({'hata': 'vardiya_id ve referans_kodu zorunludur'}), 400

    conn = get_db()
    c = conn.cursor()

    # Birden fazla kayıt gelebilir
    satirlar = data.get('satirlar', [data])

    eklenen = 0
    for satir in satirlar:
        ref = satir.get('referans_kodu', data.get('referans_kodu', '')).strip()
        ct_in = float(satir.get('cycle_time_sn', 0))
        # cycle_time gönderilmediyse ya da 0 ise referanslardan otomatik çek
        if ct_in <= 0 and ref:
            ref_row = c.execute(
                "SELECT hedef_cycle_time_sn FROM referans_listesi WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ',''))",
                (ref,)
            ).fetchone()
            if ref_row:
                ct_in = float(ref_row['hedef_cycle_time_sn'] or 0)

        c.execute('''
            INSERT INTO uretim_kayitlari (vardiya_id, referans_kodu, ok_adet, nok_adet, tamir_adet, hedef_adet, cycle_time_sn, istasyon, launch_adet)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            data['vardiya_id'],
            ref,
            int(satir.get('ok_adet', 0)),
            int(satir.get('nok_adet', 0)),
            int(satir.get('tamir_adet', 0)),
            int(satir.get('hedef_adet', 0)),
            ct_in,
            int(satir.get('istasyon', data.get('istasyon', 0))),
            int(satir.get('launch_adet', data.get('launch_adet', 0)))
        ))
        # Referansı listeye otomatik ekle
        c.execute('INSERT OR IGNORE INTO referans_listesi (referans_kodu) VALUES (?)', (ref,))
        eklenen += 1

    conn.commit()
    conn.close()
    return jsonify({'basarili': True, 'eklenen': eklenen}), 201


@app.route('/api/uretim/<int:uid>', methods=['DELETE'])
def uretim_sil(uid):
    conn = get_db()
    conn.execute('DELETE FROM uretim_kayitlari WHERE id = ?', (uid,))
    conn.commit()
    conn.close()
    return jsonify({'basarili': True})


@app.route('/api/uretim/<int:uid>', methods=['PUT'])
def uretim_guncelle(uid):
    """Uretim kaydini guncelle."""
    data = request.get_json()
    conn = get_db()
    c = conn.cursor()
    mevcut = c.execute('SELECT id FROM uretim_kayitlari WHERE id = ?', (uid,)).fetchone()
    if not mevcut:
        conn.close()
        return jsonify({'hata': 'Kayit bulunamadi'}), 404

    ref = (data.get('referans_kodu') or '').strip()
    ct = float(data.get('cycle_time_sn', 0))
    # Gelen CT 0 veya gönderilmediyse referanslardan çek
    if ct <= 0 and ref:
        ref_row = c.execute(
            "SELECT hedef_cycle_time_sn FROM referans_listesi WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ',''))",
            (ref,)
        ).fetchone()
        if ref_row:
            ct = float(ref_row['hedef_cycle_time_sn'] or 0)

    c.execute('''
        UPDATE uretim_kayitlari
        SET referans_kodu=?, ok_adet=?, nok_adet=?, tamir_adet=?, hedef_adet=?, cycle_time_sn=?, istasyon=?, launch_adet=?
        WHERE id=?
    ''', (
        ref,
        int(data.get('ok_adet', 0)),
        int(data.get('nok_adet', 0)),
        int(data.get('tamir_adet', 0)),
        int(data.get('hedef_adet', 0)),
        ct,
        int(data.get('istasyon', 0)),
        int(data.get('launch_adet', 0)),
        uid
    ))
    conn.commit()
    conn.close()
    return jsonify({'basarili': True})


@app.route('/api/uretim/ct_guncelle', methods=['POST'])
def uretim_ct_toplu_guncelle():
    """Tum uretim kayitlarinda cycle_time_sn eksik olanlari referans listesinden toplu guncelle."""
    conn = get_db()
    cur = conn.execute("""
        UPDATE uretim_kayitlari
        SET cycle_time_sn = (
            SELECT r.hedef_cycle_time_sn
            FROM referans_listesi r
            WHERE UPPER(REPLACE(r.referans_kodu,' ','')) = UPPER(REPLACE(uretim_kayitlari.referans_kodu,' ',''))
              AND r.hedef_cycle_time_sn > 0
            LIMIT 1
        )
        WHERE (cycle_time_sn IS NULL OR cycle_time_sn = 0)
          AND EXISTS (
            SELECT 1 FROM referans_listesi r
            WHERE UPPER(REPLACE(r.referans_kodu,' ','')) = UPPER(REPLACE(uretim_kayitlari.referans_kodu,' ',''))
              AND r.hedef_cycle_time_sn > 0
          )
    """)
    conn.commit()
    conn.close()
    return jsonify({'basarili': True, 'guncellenen': cur.rowcount})


# ─────────────────────────────────────────────────────────────
# DURUŞ API
# ─────────────────────────────────────────────────────────────

@app.route('/api/uretim_kayit/<int:uid>/tamamlandi', methods=['PATCH'])
def uretim_tamamlandi_guncelle(uid):
    """Üretim kaydını tamamlandı / tamamlanmadı olarak işaretle."""
    data = request.get_json() or {}
    deger = 1 if data.get('tamamlandi') else 0
    conn = get_db()
    conn.execute('UPDATE uretim_kayitlari SET tamamlandi=? WHERE id=?', (deger, uid))
    conn.commit()
    conn.close()
    return jsonify({'basarili': True})

@app.route('/api/durus', methods=['POST'])
def durus_ekle():
    """Duruş kaydı ekle."""
    data = request.get_json()

    if not data.get('vardiya_id') or not data.get('durus_sebebi'):
        return jsonify({'hata': 'vardiya_id ve durus_sebebi zorunludur'}), 400

    # Birden fazla duruş gelebilir
    satirlar = data.get('satirlar', [data])

    conn = get_db()
    c = conn.cursor()
    eklenen = 0
    for satir in satirlar:
        durus_tipi = satir.get('durus_tipi', data.get('durus_tipi', 'plansiz'))
        c.execute('''
            INSERT INTO duruslar (vardiya_id, durus_sebebi, aciklama, sure_dk, baslangic_saati, durus_tipi)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (
            data['vardiya_id'],
            satir.get('durus_sebebi', data.get('durus_sebebi', '')),
            satir.get('aciklama', ''),
            int(satir.get('sure_dk', 0)),
            satir.get('baslangic_saati', ''),
            durus_tipi
        ))
        eklenen += 1

    conn.commit()
    conn.close()
    return jsonify({'basarili': True, 'eklenen': eklenen}), 201


@app.route('/api/durus/<int:did>', methods=['DELETE'])
def durus_sil(did):
    conn = get_db()
    conn.execute('DELETE FROM duruslar WHERE id = ?', (did,))
    conn.commit()
    conn.close()
    return jsonify({'basarili': True})


@app.route('/api/durus/<int:did>', methods=['PUT'])
def durus_guncelle(did):
    """Durus kaydini guncelle."""
    data = request.get_json()
    conn = get_db()
    c = conn.cursor()
    mevcut = c.execute('SELECT id FROM duruslar WHERE id = ?', (did,)).fetchone()
    if not mevcut:
        conn.close()
        return jsonify({'hata': 'Kayit bulunamadi'}), 404
    c.execute('''
        UPDATE duruslar
        SET durus_sebebi=?, aciklama=?, sure_dk=?, baslangic_saati=?, durus_tipi=?
        WHERE id=?
    ''', (
        data.get('durus_sebebi'),
        data.get('aciklama', ''),
        int(data.get('sure_dk', 0)),
        data.get('baslangic_saati', ''),
        data.get('durus_tipi', 'plansiz'),
        did
    ))
    conn.commit()
    conn.close()
    return jsonify({'basarili': True})



# ─────────────────────────────────────────────────────────────
# OEE API
# ─────────────────────────────────────────────────────────────

@app.route('/api/oee/<int:vid>', methods=['GET'])
def oee_vardiya(vid):
    """Tek vardiya OEE hesapla."""
    sonuc = hesapla_oee(vid)
    if not sonuc:
        return jsonify({'hata': 'Vardiya bulunamadı'}), 404
    return jsonify(sonuc)


@app.route('/api/oee', methods=['GET'])
def oee_ozet():
    """Tarih aralığı, robot ve bölüm için OEE özeti."""
    tarih_bas = request.args.get('tarih_bas')
    tarih_bit = request.args.get('tarih_bit')
    robot = request.args.get('robot')
    bolum = request.args.get('bolum')

    # Varsayılan: bugün
    if not tarih_bas and not tarih_bit:
        bugun = date.today().isoformat()
        tarih_bas = bugun
        tarih_bit = bugun

    sonuc = hesapla_oee_ozet(tarih_bas, tarih_bit, robot, bolum)
    return jsonify(sonuc)


# ─────────────────────────────────────────────────────────────
# REFERANS & ROBOT API
# ─────────────────────────────────────────────────────────────

@app.route('/api/referanslar', methods=['GET'])
def referans_listesi():
    """Referans autocomplete listesi. ?bolum= ile filtrelenebilir."""
    q = request.args.get('q', '')
    bolum = request.args.get('bolum', '')
    conn = get_db()
    if bolum:
        rows = conn.execute(
            "SELECT referans_kodu, aciklama, hedef_cycle_time_sn FROM referans_listesi WHERE REPLACE(referans_kodu, ' ', '') LIKE REPLACE(?, ' ', '') AND bolum = ? ORDER BY referans_kodu",
            (f'%{q}%', bolum)
        ).fetchall()
    else:
        rows = conn.execute(
            "SELECT referans_kodu, aciklama, hedef_cycle_time_sn FROM referans_listesi WHERE REPLACE(referans_kodu, ' ', '') LIKE REPLACE(?, ' ', '') ORDER BY referans_kodu",
            (f'%{q}%',)
        ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/api/referanslar', methods=['POST'])
def referans_ekle():
    """Yeni referans tanımla veya güncelle (bolum opsiyonel — varsayılan 'kaynak')."""
    data = request.get_json()
    ref = data.get('referans_kodu', '').strip()
    if not ref:
        return jsonify({'hata': 'referans_kodu zorunludur'}), 400

    ct = float(data.get('hedef_cycle_time_sn', 0))
    desc = data.get('aciklama', '')
    bolum = (data.get('bolum') or 'kaynak').strip()
    if bolum not in ('kaynak', 'montaj', 'metal'):
        bolum = 'kaynak'

    conn = get_db()
    try:
        # INSERT OR REPLACE — Mevcutsa günceller (Süre tanımı için önemli)
        conn.execute(
            'INSERT OR REPLACE INTO referans_listesi (referans_kodu, aciklama, hedef_cycle_time_sn, bolum) VALUES (?, ?, ?, ?)',
            (ref, desc, ct, bolum)
        )
        
        # Geriye dönük güncelleme: Aynı referansa sahip geçmiş üretim kayıtlarının süresini de güncelle.
        if ct > 0:
            conn.execute("UPDATE uretim_kayitlari SET cycle_time_sn = ? WHERE UPPER(REPLACE(referans_kodu, ' ', '')) = UPPER(REPLACE(?, ' ', ''))", (ct, ref))
            
        conn.commit()
    except Exception as e:
        conn.close()
        return jsonify({'hata': str(e)}), 400
    conn.close()
    return jsonify({'basarili': True}), 201


@app.route('/api/referanslar/eksik', methods=['GET'])
def referans_eksik_listesi():
    """Süresi tanımlanmamış (0 veya NULL) referansları getirir. ?bolum= ile filtrelenebilir."""
    bolum = request.args.get('bolum')
    conn = get_db()
    c = conn.cursor()
    if bolum:
        rows = c.execute('''
            SELECT referans_kodu
            FROM referans_listesi
            WHERE (hedef_cycle_time_sn IS NULL OR hedef_cycle_time_sn = 0)
              AND COALESCE(bolum, 'kaynak') = ?
            ORDER BY referans_kodu
        ''', (bolum,)).fetchall()
    else:
        rows = c.execute('''
            SELECT referans_kodu
            FROM referans_listesi
            WHERE (hedef_cycle_time_sn IS NULL OR hedef_cycle_time_sn = 0)
            ORDER BY referans_kodu
        ''').fetchall()
    conn.close()
    return jsonify([r['referans_kodu'] for r in rows])


@app.route('/api/referanslar/export_excel', methods=['POST'])
def referans_excel_export():
    """Veritabanındaki referans listesini Kaynakhane.xlsx dosyasına doğrudan yazar."""
    excel_path = r'C:\Users\selcu\OneDrive\Masaüstü\Kaynakhane.xlsx'
    if not os.path.exists(excel_path):
        return jsonify({'hata': 'Kaynakhane.xlsx dosyası bulunamadı.'}), 404
        
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill
        
        conn = get_db()
        rows = conn.execute('SELECT referans_kodu, hedef_cycle_time_sn FROM referans_listesi ORDER BY referans_kodu').fetchall()
        conn.close()
        
        wb = openpyxl.load_workbook(excel_path)
        # Uygun sayfayı bul (Import logic'teki gibi)
        sayfa = None
        for adi in wb.sheetnames:
            if 'kaynak' in adi.lower() or 'süre' in adi.lower() or 'sure' in adi.lower():
                sayfa = wb[adi]
                break
        if not sayfa: sayfa = wb.active
        
        # Temizleyip baştan yazalım mı? 
        # Kullanıcının "doğrudan güncellesin" demesi mevcut verileri koruyup sadece bu tabloyu güncellemek olabilir.
        # En güvenlisi: Mevcut satırları tara, varsa güncelle yoksa sona ekle.
        
        ref_map = {r['referans_kodu']: r['hedef_cycle_time_sn'] for r in rows}
        
        # Mevcut satırları kontrol et
        max_r = sayfa.max_row
        existing_refs = {}
        for r in range(2, max_r + 1):
            cell_val = sayfa.cell(row=r, column=1).value
            if cell_val:
                existing_refs[str(cell_val).strip()] = r
                
        for ref, ct in ref_map.items():
            if ref in existing_refs:
                row_idx = existing_refs[ref]
                sayfa.cell(row=row_idx, column=2, value=ct)
            else:
                new_row = sayfa.max_row + 1
                sayfa.cell(row=new_row, column=1, value=ref)
                sayfa.cell(row=new_row, column=2, value=ct)
        
        wb.save(excel_path)
        return jsonify({'basarili': True, 'mesaj': f'{len(ref_map)} referans güncellendi.'})
    except Exception as e:
        return jsonify({'hata': str(e)}), 500


@app.route('/api/referanslar/<string:kod>', methods=['DELETE'])
def referans_sil(kod):
    """Referansı listeden sil."""
    conn = get_db()
    try:
        conn.execute('DELETE FROM referans_listesi WHERE referans_kodu = ?', (kod,))
        conn.commit()
    except Exception as e:
        conn.close()
        return jsonify({'hata': str(e)}), 400
    conn.close()
    return jsonify({'basarili': True}), 200


@app.route('/api/robotlar', methods=['GET'])
def robot_listesi():
    """Bölüm bazlı robot/hat/makine listesi.
    - kaynak: ABB1..ABB9 (sabit)
    - montaj: vardiyalardan distinct robot_no (HAT 1, HAT 2 ...) — operatörler girdikçe dinamik büyür
    - metal: 300T, 400T, 500T, Şerit Testere (sabit)
    """
    bolum = request.args.get('bolum', 'kaynak')
    if bolum == 'metal':
        robotlar = ['300T', '400T', '500T', 'Şerit Testere']
    elif bolum == 'montaj':
        conn = get_db()
        rows = conn.execute(
            "SELECT DISTINCT robot_no FROM vardiyalar WHERE COALESCE(bolum, 'kaynak') = 'montaj' AND robot_no IS NOT NULL AND robot_no != '' ORDER BY robot_no"
        ).fetchall()
        conn.close()
        robotlar = [r['robot_no'] for r in rows]
        # Hiç kayıt yoksa varsayılan olarak HAT 1'i göster — operatör formundan ilk girişte yenileri otomatik eklenir
        if not robotlar:
            robotlar = ['HAT 1']
    else:
        robotlar = [f'ABB{i}' for i in range(1, 10)]
    return jsonify(robotlar)


@app.route('/api/operatorler', methods=['GET'])
def operator_listesi():
    """Operatör listesini döndür. ?bolum= ile filtrelenebilir."""
    bolum = request.args.get('bolum', '')
    conn = get_db()
    if bolum:
        rows = conn.execute('SELECT id, ad, bolum FROM operatorler WHERE bolum = ? ORDER BY ad', (bolum,)).fetchall()
    else:
        rows = conn.execute('SELECT id, ad, bolum FROM operatorler ORDER BY ad').fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/api/import_excel', methods=['POST'])
def excel_import():
    """Excel dosyalarından verileri çeker.
    ?bolum=kaynak|montaj|metal → sadece o bölümün dosyası;
    parametre yoksa: hepsi.
    """
    bolum = request.args.get('bolum')
    if bolum and bolum not in ('kaynak', 'montaj', 'metal'):
        return jsonify({'hata': f"Geçersiz bölüm: {bolum}", 'basarili': False}), 400
    try:
        sonuc = import_data(bolum=bolum)
        return jsonify(sonuc), 200
    except Exception as e:
        return jsonify({'hata': str(e), 'basarili': False}), 500


@app.route('/api/export_arsiv', methods=['POST'])
def export_arsiv_endpoint():
    """Verileri masaüstündeki UretimTakipArsiv.xlsx dosyasina yaz."""
    data = request.get_json() or {}
    tarih_bas = data.get('tarih_bas')  # opsiyonel, orn: '2026-02-01'
    tarih_bit = data.get('tarih_bit')  # opsiyonel, orn: '2026-03-10'
    try:
        sonuc = export_arsiv(tarih_bas=tarih_bas, tarih_bit=tarih_bit)
        return jsonify(sonuc), 200
    except Exception as e:
        return jsonify({'basarili': False, 'hata': str(e)}), 500


# ─────────────────────────────────────────────────────────────
# ÖZET / DASHBOARD VERİSİ
# ─────────────────────────────────────────────────────────────

@app.route('/api/ozet', methods=['GET'])
def ozet():
    """Dashboard için kapsamlı özet verisi."""
    conn = get_db()
    c = conn.cursor()

    bugun = date.today().isoformat()
    tarih_bas = request.args.get('tarih_bas', bugun)
    tarih_bit = request.args.get('tarih_bit', bugun)
    robot = request.args.get('robot')
    bolum = request.args.get('bolum')

    param_vardiya = [tarih_bas, tarih_bit]
    sart_vardiya = 'tarih BETWEEN ? AND ?'
    if robot:
        sart_vardiya += ' AND robot_no = ?'
        param_vardiya.append(robot)
    if bolum:
        sart_vardiya += " AND COALESCE(v.bolum, 'kaynak') = ?"
        param_vardiya.append(bolum)

    # Vardiya sayıları
    vardiya_sayisi = c.execute(
        f'SELECT COUNT(*) as cnt FROM vardiyalar v WHERE {sart_vardiya}',
        param_vardiya
    ).fetchone()['cnt']

    # Robot bazlı üretim özeti
    robot_uretim = c.execute(f'''
        SELECT v.robot_no,
               SUM(u.ok_adet)  AS toplam_ok,
               SUM(u.nok_adet) AS toplam_nok,
               SUM(u.ok_adet + u.nok_adet) AS toplam_uretim
        FROM vardiyalar v
        LEFT JOIN uretim_kayitlari u ON u.vardiya_id = v.id
        WHERE {sart_vardiya}
        GROUP BY v.robot_no
        ORDER BY v.robot_no
    ''', param_vardiya).fetchall()

    # Referans bazlı üretim
    referans_uretim = c.execute(f'''
        SELECT u.referans_kodu,
               SUM(u.ok_adet)  as toplam_ok,
               SUM(u.nok_adet) as toplam_nok,
               SUM(u.ok_adet + u.nok_adet) as toplam_uretim,
               ROUND(AVG(u.cycle_time_sn), 0) as ort_cycle_time_sn
        FROM uretim_kayitlari u
        JOIN vardiyalar v ON v.id = u.vardiya_id
        WHERE {sart_vardiya}
        GROUP BY u.referans_kodu
        ORDER BY toplam_uretim DESC
    ''', param_vardiya).fetchall()

    # Duruş sebep bazlı dağılım
    durus_dagilim = c.execute(f'''
        SELECT d.durus_sebebi,
               d.durus_tipi,
               SUM(d.sure_dk) as toplam_sure,
               COUNT(*) as adet
        FROM duruslar d
        JOIN vardiyalar v ON v.id = d.vardiya_id
        WHERE {sart_vardiya}
        GROUP BY d.durus_sebebi
        ORDER BY toplam_sure DESC
    ''', param_vardiya).fetchall()

    # Planlı / Plansız duruş tip özeti
    durus_tipi_ozet = c.execute(f'''
        SELECT d.durus_tipi,
               SUM(d.sure_dk) as toplam_sure,
               COUNT(*) as adet
        FROM duruslar d
        JOIN vardiyalar v ON v.id = d.vardiya_id
        WHERE {sart_vardiya}
        GROUP BY d.durus_tipi
    ''', param_vardiya).fetchall()

    # Üretim Kayıtları (Detaylı - Düzenlenebilir)
    uretim_detay_listesi = c.execute(f'''
        SELECT u.id as uretim_id,
               v.tarih,
               v.robot_no,
               v.operator_adi,
               u.referans_kodu,
               u.ok_adet,
               u.nok_adet,
               u.hedef_adet,
               u.cycle_time_sn
        FROM uretim_kayitlari u
        JOIN vardiyalar v ON v.id = u.vardiya_id
        WHERE {sart_vardiya}
        ORDER BY v.tarih DESC, v.robot_no ASC
    ''', param_vardiya).fetchall()

    # Robot + Referans bazlı üretim kırılımı
    robot_referans_uretim = c.execute(f'''
        SELECT v.robot_no,
               u.referans_kodu,
               SUM(u.ok_adet)  as toplam_ok,
               SUM(u.nok_adet) as toplam_nok,
               SUM(u.ok_adet + u.nok_adet) as toplam_uretim,
               ROUND(AVG(u.cycle_time_sn), 0) as ort_cycle_time_sn
        FROM uretim_kayitlari u
        JOIN vardiyalar v ON v.id = u.vardiya_id
        WHERE {sart_vardiya}
        GROUP BY v.robot_no, u.referans_kodu
        ORDER BY v.robot_no, toplam_uretim DESC
    ''', param_vardiya).fetchall()

    # Robot + Duruş bazlı kırılım (Güncellendi: Tekil kayıt - Düzenlenebilir)
    robot_durus_kirilim = c.execute(f'''
        SELECT d.id as durus_id,
               v.robot_no,
               d.durus_sebebi,
               d.durus_tipi,
               d.sure_dk as toplam_sure,
               v.operator_adi,
               1 as adet
        FROM duruslar d
        JOIN vardiyalar v ON v.id = d.vardiya_id
        WHERE {sart_vardiya}
        ORDER BY v.robot_no, d.sure_dk DESC
    ''', param_vardiya).fetchall()

    # Son 50 vardiya kayıtları (OEE trendi için)
    son_vardiyalar = c.execute(
        f'SELECT v.id FROM vardiyalar v WHERE {sart_vardiya} ORDER BY v.tarih DESC, v.id DESC LIMIT 50',
        param_vardiya
    ).fetchall()

    oee_listesi = []
    for row in son_vardiyalar:
        oee_data = hesapla_oee(row['id'])
        if oee_data:
            oee_listesi.append({
                'vardiya_id': oee_data['vardiya_id'],
                'tarih': oee_data['tarih'],
                'vardiya': oee_data['vardiya_turu'],
                'robot': oee_data['robot_no'],
                'operator': oee_data['operator'],
                'baslangic_saati': oee_data.get('baslangic_saati', ''),
                'bitis_saati': oee_data.get('bitis_saati', ''),
                'oee': oee_data['oee'],
                'availability': oee_data['availability'],
                'performance': oee_data['performance'],
                'quality': oee_data['quality'],
                'ok': oee_data['toplam_ok'],
                'nok': oee_data['toplam_nok'],
                'hedef': oee_data['toplam_hedef'],
                'planli_durus_dk': oee_data.get('planli_durus_dk', 0),
                'plansiz_durus_dk': oee_data.get('plansiz_durus_dk', 0),
                'durus_dk': oee_data.get('planli_durus_dk', 0) + oee_data.get('plansiz_durus_dk', 0),
            })


    conn.close()

    ort_oee = round(sum(o['oee'] for o in oee_listesi) / len(oee_listesi), 1) if oee_listesi else 0

    return jsonify({
        'tarih_aralik': {'bas': tarih_bas, 'bit': tarih_bit},
        'vardiya_sayisi': vardiya_sayisi,
        'ort_oee': ort_oee,
        'robot_uretim': [dict(r) for r in robot_uretim],
        'referans_uretim': [dict(r) for r in referans_uretim],
        'uretim_detay_listesi': [dict(r) for r in uretim_detay_listesi],
        'robot_referans_uretim': [dict(r) for r in robot_referans_uretim],
        'durus_dagilim': [dict(r) for r in durus_dagilim],
        'durus_tipi_ozet': [dict(r) for r in durus_tipi_ozet],
        'robot_durus_kirilim': [dict(r) for r in robot_durus_kirilim],
        'vardiyalar': oee_listesi
    })


# ─────────────────────────────────────────────────────────────
# RAPOR API
# ─────────────────────────────────────────────────────────────

@app.route('/api/rapor', methods=['GET'])
def rapor_api():
    """Rapor verisi: operator, referans veya teep."""
    rapor_tipi = request.args.get('rapor_tipi', 'operator')
    tarih_bas = request.args.get('tarih_bas', date.today().isoformat())
    tarih_bit = request.args.get('tarih_bit', date.today().isoformat())
    robot = request.args.get('robot', '')
    operator_filtre = request.args.get('operator', '')
    bolum = request.args.get('bolum', '')

    conn = get_db()
    c = conn.cursor()

    # Temel vardiya filtresi
    sart = 'v.tarih BETWEEN ? AND ?'
    params = [tarih_bas, tarih_bit]
    if robot:
        sart += ' AND v.robot_no = ?'
        params.append(robot)
    if operator_filtre:
        sart += ' AND v.operator_adi = ?'
        params.append(operator_filtre)
    if bolum:
        sart += " AND COALESCE(v.bolum, 'kaynak') = ?"
        params.append(bolum)

    if rapor_tipi == 'operator':
        # Operator bazli rapor
        rows = c.execute(f'''
            SELECT v.operator_adi,
                   COUNT(DISTINCT v.id) as vardiya_sayisi,
                   SUM(v.toplam_sure_dk) as toplam_calisma_dk,
                   COALESCE(SUM(u.ok_adet), 0) as toplam_ok,
                   COALESCE(SUM(u.nok_adet), 0) as toplam_nok
            FROM vardiyalar v
            LEFT JOIN uretim_kayitlari u ON u.vardiya_id = v.id
            WHERE {sart}
            GROUP BY v.operator_adi
            ORDER BY toplam_ok DESC
        ''', params).fetchall()

        sonuclar = []
        for r in rows:
            op = r['operator_adi']
            # Bu operatorun vardiya id'lerini al
            vp = list(params)
            v_sart = sart + " AND v.operator_adi = ?"
            vp.append(op)
            vids = [x['id'] for x in c.execute(
                f'SELECT v.id FROM vardiyalar v WHERE {v_sart}', vp
            ).fetchall()]

            # Durus toplami
            if vids:
                placeholders = ','.join('?' * len(vids))
                planli = c.execute(
                    f"SELECT COALESCE(SUM(sure_dk),0) as t FROM duruslar WHERE vardiya_id IN ({placeholders}) AND durus_tipi='planli'",
                    vids
                ).fetchone()['t']
                plansiz = c.execute(
                    f"SELECT COALESCE(SUM(sure_dk),0) as t FROM duruslar WHERE vardiya_id IN ({placeholders}) AND durus_tipi='plansiz'",
                    vids
                ).fetchone()['t']
            else:
                planli = plansiz = 0

            # OEE ortalamasi
            oee_list = []
            avail_list = []
            perf_list = []
            qual_list = []
            for vid in vids:
                oee_d = hesapla_oee(vid)
                if oee_d:
                    oee_list.append(oee_d['oee'])
                    avail_list.append(oee_d['availability'])
                    perf_list.append(oee_d['performance'])
                    qual_list.append(oee_d['quality'])

            ort_oee = round(sum(oee_list) / len(oee_list), 1) if oee_list else 0
            ort_avail = round(sum(avail_list) / len(avail_list), 1) if avail_list else 0
            ort_perf = round(sum(perf_list) / len(perf_list), 1) if perf_list else 0
            ort_qual = round(sum(qual_list) / len(qual_list), 1) if qual_list else 0

            toplam_uretim = (r['toplam_ok'] or 0) + (r['toplam_nok'] or 0)
            sonuclar.append({
                'operator': op,
                'vardiya_sayisi': r['vardiya_sayisi'],
                'toplam_calisma_dk': r['toplam_calisma_dk'] or 0,
                'toplam_ok': r['toplam_ok'] or 0,
                'toplam_nok': r['toplam_nok'] or 0,
                'toplam_uretim': toplam_uretim,
                'planli_durus_dk': planli,
                'plansiz_durus_dk': plansiz,
                'ort_oee': ort_oee,
                'ort_availability': ort_avail,
                'ort_performance': ort_perf,
                'ort_quality': ort_qual,
            })
        conn.close()
        return jsonify({'rapor_tipi': 'operator', 'tarih': {'bas': tarih_bas, 'bit': tarih_bit}, 'veriler': sonuclar})

    elif rapor_tipi == 'referans':
        # Referans bazli rapor
        rows = c.execute(f'''
            SELECT u.referans_kodu,
                   SUM(u.ok_adet) as toplam_ok,
                   SUM(u.nok_adet) as toplam_nok,
                   SUM(u.ok_adet + u.nok_adet) as toplam_uretim,
                   ROUND(AVG(u.cycle_time_sn), 1) as ort_ct,
                   GROUP_CONCAT(DISTINCT v.robot_no) as robotlar,
                   GROUP_CONCAT(DISTINCT v.operator_adi) as operatorler,
                   COUNT(DISTINCT v.id) as vardiya_sayisi
            FROM uretim_kayitlari u
            JOIN vardiyalar v ON v.id = u.vardiya_id
            WHERE {sart}
            GROUP BY u.referans_kodu
            ORDER BY toplam_uretim DESC
        ''', params).fetchall()

        sonuclar = []
        for r in rows:
            toplam = (r['toplam_ok'] or 0) + (r['toplam_nok'] or 0)
            kalite = round((r['toplam_ok'] or 0) / toplam * 100, 1) if toplam > 0 else 0
            # Hedef CT'yi referans listesinden al
            ref_row = c.execute(
                'SELECT hedef_cycle_time_sn FROM referans_listesi WHERE referans_kodu = ?',
                (r['referans_kodu'],)
            ).fetchone()
            hedef_ct = ref_row['hedef_cycle_time_sn'] if ref_row else 0

            sonuclar.append({
                'referans_kodu': r['referans_kodu'],
                'toplam_ok': r['toplam_ok'] or 0,
                'toplam_nok': r['toplam_nok'] or 0,
                'toplam_uretim': toplam,
                'kalite_pct': kalite,
                'ort_ct': r['ort_ct'] or 0,
                'hedef_ct': hedef_ct or 0,
                'robotlar': r['robotlar'] or '',
                'operatorler': r['operatorler'] or '',
                'vardiya_sayisi': r['vardiya_sayisi'],
            })
        conn.close()
        return jsonify({'rapor_tipi': 'referans', 'tarih': {'bas': tarih_bas, 'bit': tarih_bit}, 'veriler': sonuclar})

    elif rapor_tipi == 'teep':
        # TEEP Selalesi
        # 1) Takvim suresi
        from datetime import timedelta
        d_bas = datetime.strptime(tarih_bas, '%Y-%m-%d').date()
        d_bit = datetime.strptime(tarih_bit, '%Y-%m-%d').date()
        gun_sayisi = max((d_bit - d_bas).days + 1, 1)
        takvim_dk = gun_sayisi * 24 * 60  # 24h x gun

        # 2) Toplam vardiya suresi
        row = c.execute(f'''
            SELECT COALESCE(SUM(v.toplam_sure_dk), 0) as toplam
            FROM vardiyalar v WHERE {sart}
        ''', params).fetchone()
        toplam_vardiya_dk = row['toplam']

        # 3) Planli duruslar
        row = c.execute(f'''
            SELECT COALESCE(SUM(d.sure_dk), 0) as toplam
            FROM duruslar d JOIN vardiyalar v ON v.id = d.vardiya_id
            WHERE {sart} AND d.durus_tipi = 'planli'
        ''', params).fetchone()
        planli_durus_dk = row['toplam']

        # 4) Plansiz duruslar
        row = c.execute(f'''
            SELECT COALESCE(SUM(d.sure_dk), 0) as toplam
            FROM duruslar d JOIN vardiyalar v ON v.id = d.vardiya_id
            WHERE {sart} AND d.durus_tipi = 'plansiz'
        ''', params).fetchone()
        plansiz_durus_dk = row['toplam']

        # 5) Net plan suresi
        net_plan_dk = max(0, toplam_vardiya_dk - planli_durus_dk)
        # 6) Fiili calisma
        fiili_calisma_dk = max(0, net_plan_dk - plansiz_durus_dk)

        # 7) Uretim verileri
        uretim_rows = c.execute(f'''
            SELECT u.ok_adet, u.nok_adet, u.cycle_time_sn,
                   r.hedef_cycle_time_sn as guncel_ct
            FROM uretim_kayitlari u
            JOIN vardiyalar v ON v.id = u.vardiya_id
            LEFT JOIN referans_listesi r ON u.referans_kodu = r.referans_kodu
            WHERE {sart}
        ''', params).fetchall()

        toplam_ok = sum(r['ok_adet'] for r in uretim_rows)
        toplam_nok = sum(r['nok_adet'] for r in uretim_rows)
        toplam_uretim = toplam_ok + toplam_nok

        # Ideal uretim suresi (sn -> dk)
        ideal_uretim_sn = 0
        for r in uretim_rows:
            ct = r['guncel_ct'] if (r['guncel_ct'] and r['guncel_ct'] > 0) else (r['cycle_time_sn'] or 0)
            adet = r['ok_adet'] + r['nok_adet']
            ideal_uretim_sn += adet * ct
        ideal_uretim_dk = ideal_uretim_sn / 60

        # Performans kaybi
        performans_kaybi_dk = max(0, fiili_calisma_dk - ideal_uretim_dk)

        # Kalite kaybi
        kalite_kaybi_dk = (toplam_nok / toplam_uretim * ideal_uretim_dk) if toplam_uretim > 0 else 0

        # Net degerli uretim suresi
        net_uretim_dk = max(0, ideal_uretim_dk - kalite_kaybi_dk)

        # TEEP
        teep_pct = round((net_uretim_dk / takvim_dk) * 100, 1) if takvim_dk > 0 else 0
        # OEE (referans)
        oee_pct = round((net_uretim_dk / net_plan_dk) * 100, 1) if net_plan_dk > 0 else 0
        # Loading (kullanim orani)
        loading_pct = round((toplam_vardiya_dk / takvim_dk) * 100, 1) if takvim_dk > 0 else 0
        # Kullanilabilirlik icin dururm sayilari
        plansiz_durus_kalemler = c.execute(f'''
            SELECT d.durus_sebebi, SUM(d.sure_dk) as toplam, COUNT(*) as adet
            FROM duruslar d JOIN vardiyalar v ON v.id = d.vardiya_id
            WHERE {sart} AND d.durus_tipi = 'plansiz'
            GROUP BY d.durus_sebebi ORDER BY toplam DESC
        ''', params).fetchall()
        planli_durus_kalemler = c.execute(f'''
            SELECT d.durus_sebebi, SUM(d.sure_dk) as toplam, COUNT(*) as adet
            FROM duruslar d JOIN vardiyalar v ON v.id = d.vardiya_id
            WHERE {sart} AND d.durus_tipi = 'planli'
            GROUP BY d.durus_sebebi ORDER BY toplam DESC
        ''', params).fetchall()

        conn.close()

        # Uretilmeyen sure (takvimde olup vardiyaya alinmamis)
        uretilmeyen_dk = max(0, takvim_dk - toplam_vardiya_dk)

        return jsonify({
            'rapor_tipi': 'teep',
            'tarih': {'bas': tarih_bas, 'bit': tarih_bit},
            'gun_sayisi': gun_sayisi,
            'selale': {
                'takvim_dk': takvim_dk,
                'uretilmeyen_dk': round(uretilmeyen_dk, 1),
                'toplam_vardiya_dk': round(toplam_vardiya_dk, 1),
                'planli_durus_dk': round(planli_durus_dk, 1),
                'net_plan_dk': round(net_plan_dk, 1),
                'plansiz_durus_dk': round(plansiz_durus_dk, 1),
                'fiili_calisma_dk': round(fiili_calisma_dk, 1),
                'performans_kaybi_dk': round(performans_kaybi_dk, 1),
                'kalite_kaybi_dk': round(kalite_kaybi_dk, 1),
                'net_uretim_dk': round(net_uretim_dk, 1),
            },
            'yuzde': {
                'teep': teep_pct,
                'oee': oee_pct,
                'loading': loading_pct,
            },
            'uretim': {
                'toplam_ok': toplam_ok,
                'toplam_nok': toplam_nok,
                'toplam_uretim': toplam_uretim,
            },
            'plansiz_kalemler': [dict(r) for r in plansiz_durus_kalemler],
            'planli_kalemler': [dict(r) for r in planli_durus_kalemler],
        })

    conn.close()
    return jsonify({'hata': 'Gecersiz rapor_tipi'}), 400


# ─────────────────────────────────────────────────────────────
# FİKSTÜR ADRESLERİ API
# ─────────────────────────────────────────────────────────────

@app.route('/api/fikstur', methods=['GET'])
def get_fikstur_listesi():
    """Tüm fikstür adreslerini getir."""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM fikstur_adresleri ORDER BY referans_kodu ASC")
    kayitlar = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify(kayitlar)

@app.route('/api/fikstur', methods=['POST'])
def kaydet_fikstur():
    """Yeni fikstür adresi ekle veya mevcut olanı güncelle, ODS'yi de güncelle."""
    data = request.json
    ref = data.get('referans_kodu', '').strip()
    raf = data.get('raf_adresi', '').strip()
    notlar = data.get('notlar', '').strip()

    if not ref:
        return jsonify({'error': 'Referans kodu zorunludur.'}), 400

    conn = get_db()
    c = conn.cursor()
    
    # Eger referans varsa Guncelle, yoksa Ekle
    c.execute("SELECT id FROM fikstur_adresleri WHERE referans_kodu=?", (ref,))
    row = c.fetchone()
    if row:
        c.execute("UPDATE fikstur_adresleri SET raf_adresi=?, notlar=?, created_at=datetime('now', 'localtime') WHERE id=?", 
                  (raf, notlar, row['id']))
    else:
        c.execute("INSERT INTO fikstur_adresleri (referans_kodu, raf_adresi, notlar) VALUES (?,?,?)", 
                  (ref, raf, notlar))
    
    conn.commit()

    # ODS'yi güncelle
    c.execute("SELECT referans_kodu, raf_adresi, notlar FROM fikstur_adresleri ORDER BY referans_kodu ASC")
    tum_kayitlar = [(r['referans_kodu'], r['raf_adresi'], r['notlar']) for r in c.fetchall()]
    conn.close()
    _ods_guncelle(tum_kayitlar)

    return jsonify({'message': 'Fikstür adresi kaydedildi.'})


@app.route('/api/fikstur/ods_import', methods=['POST'])
def fikstur_ods_import():
    """ODS dosyasındaki tüm kayıtları DB'ye aktarır (zaten varsa atlar)."""
    kayitlar = _ods_kayitlari_oku()
    if not kayitlar:
        return jsonify({'error': 'ODS dosyası okunamadı veya boş.'}), 400

    conn = get_db()
    c = conn.cursor()
    eklenen, atlanan = 0, 0
    for ref, raf in kayitlar.items():
        c.execute("SELECT id FROM fikstur_adresleri WHERE referans_kodu=?", (ref,))
        if c.fetchone():
            atlanan += 1
        else:
            c.execute("INSERT INTO fikstur_adresleri (referans_kodu, raf_adresi, notlar) VALUES (?,?,?)",
                      (ref, raf, ''))
            eklenen += 1
    conn.commit()
    conn.close()
    return jsonify({'eklenen': eklenen, 'atlanan': atlanan,
                    'mesaj': f'{eklenen} yeni kayıt eklendi, {atlanan} kayıt zaten mevcuttu.'})


@app.route('/api/fikstur/<int:id>', methods=['DELETE'])
def sil_fikstur(id):
    """Fikstür adresini sil ve ODS'yi güncelle."""
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM fikstur_adresleri WHERE id=?", (id,))
    conn.commit()
    # ODS güncelle
    c.execute("SELECT referans_kodu, raf_adresi, notlar FROM fikstur_adresleri ORDER BY referans_kodu ASC")
    tum_kayitlar = [(r['referans_kodu'], r['raf_adresi'], r['notlar']) for r in c.fetchall()]
    conn.close()
    _ods_guncelle(tum_kayitlar)
    return jsonify({'message': 'Fikstür silindi.'})


# ─────────────────────────────────────────────────────────────
# ROBOT İSTASYON PROGRAMLARI API
# ─────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────
# REFERANS ROBOT UYUMU API
# ─────────────────────────────────────────────────────────────

@app.route('/api/referans_uyum', methods=['GET'])
def get_referans_uyum():
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM referans_robot_uyumu ORDER BY referans_kodu ASC")
    kayitlar = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify(kayitlar)

@app.route('/api/referans_uyum', methods=['POST'])
def kaydet_referans_uyum():
    data = request.json
    ref = data.get('referans_kodu', '').strip()
    robot_no = data.get('robot_no', '').strip()
    istasyon = int(data.get('istasyon', 0))
    
    if not ref or not robot_no or not istasyon:
        return jsonify({'error': 'Eksik bilgi.'}), 400
        
    conn = get_db()
    c = conn.cursor()
    # Aynı eşleşme var mı kontrol et
    c.execute("SELECT id FROM referans_robot_uyumu WHERE referans_kodu=? AND robot_no=? AND istasyon=?", (ref, robot_no, istasyon))
    if not c.fetchone():
        c.execute("INSERT INTO referans_robot_uyumu (referans_kodu, robot_no, istasyon) VALUES (?,?,?)", (ref, robot_no, istasyon))
    conn.commit()
    conn.close()
    return jsonify({'message': 'Uyum kaydedildi.'})

@app.route('/api/referans_uyum/<int:id>', methods=['DELETE'])
def sil_referans_uyum(id):
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM referans_robot_uyumu WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return jsonify({'message': 'Silindi.'})

@app.route('/api/referans_uyum/ods_import', methods=['POST'])
def referans_uyum_ods_import():
    kayitlar = _ods_uyum_oku()
    if not kayitlar:
        return jsonify({'error': 'ODS okunamadı veya eşleşme bulunamadı.'}), 400
        
    conn = get_db()
    c = conn.cursor()
    eklenen, atlanan = 0, 0
    # Eski kayıtları temizleyelim (Opsiyonel: Veya mevcutları koruyup üstüne ekleyebiliriz)
    # Hızlı kurulum için temizlemek faydalı olabilir ancak manuel eklenenleri silmemek adına sadece yeni ekleyelim.
    for ref, r_no, ist in kayitlar:
        c.execute("SELECT id FROM referans_robot_uyumu WHERE referans_kodu=? AND robot_no=? AND istasyon=?", (ref, r_no, ist))
        if c.fetchone():
            atlanan += 1
        else:
            c.execute("INSERT INTO referans_robot_uyumu (referans_kodu, robot_no, istasyon) VALUES (?,?,?)", (ref, r_no, ist))
            eklenen += 1
    conn.commit()
    conn.close()
    return jsonify({'eklenen': eklenen, 'atlanan': atlanan, 'mesaj': f'{eklenen} yeni uyum eklendi, {atlanan} kayıt zaten mevcuttu.'})

@app.route('/api/robot_programlari', methods=['GET'])
def get_robot_programlari():
    """Robot istasyon programlarını getir (tüm kayıtlar, robot+istasyon sırası ile)."""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM robot_programlari ORDER BY robot_no ASC, istasyon ASC, guncelleme_tarihi DESC")
    kayitlar = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify(kayitlar)

@app.route('/api/robot_programlari', methods=['POST'])
def ekle_robot_programi():
    """Robot istasyon programı ekle (çoklu kayıt - INSERT)."""
    data = request.json
    robot_no = data.get('robot_no')
    istasyon = data.get('istasyon')
    ref = data.get('referans_kodu', '').strip()
    operator = data.get('guncelleyen', '').strip()

    if not robot_no or istasyon is None:
        return jsonify({'error': 'Robot no ve istasyon zorunludur.'}), 400
    if not ref:
        return jsonify({'error': 'Referans kodu zorunludur.'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("""
        INSERT INTO robot_programlari (robot_no, istasyon, referans_kodu, guncelleyen)
        VALUES (?, ?, ?, ?)
    """, (robot_no, istasyon, ref, operator))
    new_id = c.lastrowid
    conn.commit()
    conn.close()
    
    return jsonify({'message': 'Robot programı eklendi.', 'id': new_id})

@app.route('/api/robot_programlari/<int:id>', methods=['DELETE'])
def sil_robot_programi(id):
    """Robot programı kaydını sil."""
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM robot_programlari WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return jsonify({'message': 'Kayıt silindi.'})



# ─────────────────────────────────────────────────────────────
# ROBOT İŞ ATAMA API
# ─────────────────────────────────────────────────────────────

@app.route('/api/robot_is_atamalari', methods=['GET'])
def is_atamalari_listesi():
    """Tüm aktif iş atamalarını döndür."""
    conn = get_db()
    rows = conn.execute(
        'SELECT * FROM robot_is_atamalari WHERE durum="bekliyor" ORDER BY atama_tarihi DESC'
    ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/robot_is_atamalari', methods=['POST'])
def is_atamasi_ekle():
    """Yeni iş ataması ekle."""
    data = request.get_json() or {}
    robot_no = data.get('robot_no', '').strip()
    referans_kodu = data.get('referans_kodu', '').strip()
    if not robot_no or not referans_kodu:
        return jsonify({'error': 'robot_no ve referans_kodu zorunludur'}), 400
    conn = get_db()
    c = conn.cursor()
    c.execute('''
        INSERT INTO robot_is_atamalari (robot_no, istasyon, referans_kodu, aciklama, atayan)
        VALUES (?, ?, ?, ?, ?)
    ''', (
        robot_no,
        int(data.get('istasyon', 0)),
        referans_kodu,
        data.get('aciklama', ''),
        data.get('atayan', '')
    ))
    conn.commit()
    conn.close()
    return jsonify({'basarili': True}), 201

@app.route('/api/robot_is_atamalari/<int:id>', methods=['DELETE'])
def is_atamasi_sil(id):
    """İş atamasını tamamlandı olarak işaretle (sil)."""
    conn = get_db()
    conn.execute('DELETE FROM robot_is_atamalari WHERE id=?', (id,))
    conn.commit()
    conn.close()
    return jsonify({'basarili': True})



# ─────────────────────────────────────────────────────────────
# ANDON ROBOT AYARLARI API
# ─────────────────────────────────────────────────────────────

@app.route('/api/andon_robot_ayarlari', methods=['GET'])
def andon_robot_ayarlari_listesi():
    conn = get_db()
    rows = conn.execute('SELECT * FROM andon_robot_ayarlari ORDER BY sira').fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/andon_robot_ayarlari', methods=['PATCH'])
def andon_robot_ayarlari_guncelle():
    """Toplu güncelleme: [{robot_no, goster, sira}, ...]"""
    data = request.get_json() or []
    conn = get_db()
    for item in data:
        conn.execute(
            'UPDATE andon_robot_ayarlari SET goster=?, sira=? WHERE robot_no=?',
            (int(item.get('goster', 1)), int(item.get('sira', 0)), item['robot_no'])
        )
    conn.commit()
    conn.close()
    return jsonify({'basarili': True})
    
@app.route('/api/ayarlar', methods=['GET'])
def get_ayarlar():
    conn = get_db()
    rows = conn.execute('SELECT anahtar, deger FROM genel_ayarlar').fetchall()
    conn.close()
    return jsonify({r['anahtar']: r['deger'] for r in rows})

@app.route('/api/ayarlar', methods=['POST'])
def save_ayarlar():
    data = request.get_json() or {}
    conn = get_db()
    for k, v in data.items():
        conn.execute('INSERT OR REPLACE INTO genel_ayarlar (anahtar, deger) VALUES (?, ?)', (k, str(v)))
    conn.commit()
    conn.close()
    return jsonify({'basarili': True})



# ─────────────────────────────────────────────────────────────
# REFERANS TAKİP (LAUNCH / İŞ EMRİ) API
# ─────────────────────────────────────────────────────────────

DURUM_SIRASI = ['launch_alinacak', 'launch_alindi', 'launch_hazir', 'uretimde', 'uretim_tamamlandi']

@app.route('/api/referans_takip', methods=['GET'])
def referans_takip_listesi():
    bolum = request.args.get('bolum')
    conn = get_db()
    # Montaj için öncelik sırasına göre listele (NULL olanlar en sonda),
    # diğer bölümlerde oluşturma tarihine göre.
    if bolum == 'montaj':
        rows = conn.execute('''
            SELECT rt.*, rl.hedef_cycle_time_sn
            FROM referans_takip rt
            LEFT JOIN referans_listesi rl ON REPLACE(rt.referans_kodu, ' ', '') = REPLACE(rl.referans_kodu, ' ', '')
            WHERE COALESCE(rt.bolum, 'kaynak') = 'montaj'
            ORDER BY (rt.oncelik IS NULL), rt.oncelik ASC, rt.olusturma_tarihi DESC
        ''').fetchall()
    elif bolum:
        rows = conn.execute('''
            SELECT rt.*, rl.hedef_cycle_time_sn
            FROM referans_takip rt
            LEFT JOIN referans_listesi rl ON REPLACE(rt.referans_kodu, ' ', '') = REPLACE(rl.referans_kodu, ' ', '')
            WHERE COALESCE(rt.bolum, 'kaynak') = ?
            ORDER BY rt.olusturma_tarihi DESC
        ''', (bolum,)).fetchall()
    else:
        rows = conn.execute('''
            SELECT rt.*, rl.hedef_cycle_time_sn
            FROM referans_takip rt
            LEFT JOIN referans_listesi rl ON REPLACE(rt.referans_kodu, ' ', '') = REPLACE(rl.referans_kodu, ' ', '')
            ORDER BY rt.olusturma_tarihi DESC
        ''').fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

def _oncelik_yer_ac(c, bolum, yeni_oncelik, exclude_id=None):
    """Montaj'da bir kayıt belirli bir önceliğe yerleştirilirken, mevcut
    bu önceliğe sahip ve daha aşağıdaki kayıtların önceliğini +1 kaydırır.
    exclude_id verilirse o kaydı kaydırmadan dışarıda bırakır (PATCH için).
    """
    if bolum != 'montaj' or yeni_oncelik is None:
        return
    if exclude_id is not None:
        c.execute('''
            UPDATE referans_takip
            SET oncelik = oncelik + 1
            WHERE COALESCE(bolum, 'kaynak') = 'montaj'
              AND oncelik IS NOT NULL AND oncelik >= ?
              AND id != ?
        ''', (yeni_oncelik, exclude_id))
    else:
        c.execute('''
            UPDATE referans_takip
            SET oncelik = oncelik + 1
            WHERE COALESCE(bolum, 'kaynak') = 'montaj'
              AND oncelik IS NOT NULL AND oncelik >= ?
        ''', (yeni_oncelik,))


@app.route('/api/referans_takip', methods=['POST'])
def referans_takip_ekle():
    try:
        data = request.get_json() or {}
        ref = data.get('referans_kodu', '').strip()
        if not ref:
            return jsonify({'error': 'referans_kodu zorunludur'}), 400

        bolum = (data.get('bolum') or 'kaynak').strip()
        if bolum not in ('kaynak', 'montaj', 'metal'):
            bolum = 'kaynak'

        # Öncelik (yalnızca montaj için anlamlı). Boş/None: belirtilmemiş.
        oncelik_raw = data.get('oncelik')
        oncelik = None
        if oncelik_raw not in (None, '', 0, '0'):
            try:
                oncelik = int(oncelik_raw)
                if oncelik < 1: oncelik = None
            except (TypeError, ValueError):
                oncelik = None

        conn = get_db()
        c = conn.cursor()
        # Montaj + öncelik verilmişse: mevcut >=N olanları kaydır
        _oncelik_yer_ac(c, bolum, oncelik)
        c.execute('''
            INSERT INTO referans_takip (referans_kodu, hedef_adet, aciklama, durum, olusturan, robot_no, istasyon, bolum, oncelik)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (ref, int(data.get('hedef_adet', 0)), data.get('aciklama', ''),
              data.get('durum', 'launch_alinacak'), data.get('olusturan', ''),
              data.get('robot_no', ''), int(data.get('istasyon', 0)), bolum, oncelik))
        conn.commit()
        return jsonify({'basarili': True}), 201
    except Exception as e:
        print(f"HATA (referans_takip_ekle): {traceback.format_exc()}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/referans_takip/<int:id>', methods=['PATCH'])
def referans_takip_guncelle(id):
    data = request.get_json() or {}
    conn = get_db()
    c = conn.cursor()

    # Öncelik değişiyorsa shift mantığı çalıştır (montaj için)
    if 'oncelik' in data:
        mevcut = c.execute("SELECT bolum, oncelik FROM referans_takip WHERE id=?", (id,)).fetchone()
        if mevcut:
            bolum_kayit = (mevcut['bolum'] or 'kaynak')
            yeni_raw = data.get('oncelik')
            yeni_onc = None
            if yeni_raw not in (None, '', 0, '0'):
                try:
                    yeni_onc = int(yeni_raw)
                    if yeni_onc < 1: yeni_onc = None
                except (TypeError, ValueError):
                    yeni_onc = None
            _oncelik_yer_ac(c, bolum_kayit, yeni_onc, exclude_id=id)
            data['oncelik'] = yeni_onc  # normalize edildi

    fields, vals = [], []
    if 'durum' in data:
        fields.append('durum=?'); vals.append(data['durum'])
    if 'aciklama' in data:
        fields.append('aciklama=?'); vals.append(data['aciklama'])
    if 'hedef_adet' in data:
        fields.append('hedef_adet=?'); vals.append(int(data['hedef_adet']))
    if 'referans_kodu' in data:
        fields.append('referans_kodu=?'); vals.append(data['referans_kodu'])
    if 'robot_no' in data:
        fields.append('robot_no=?'); vals.append(data['robot_no'])
    if 'istasyon' in data:
        fields.append('istasyon=?'); vals.append(int(data['istasyon']))
    if 'oncelik' in data:
        fields.append('oncelik=?'); vals.append(data['oncelik'])  # None olabilir
    if not fields:
        conn.close()
        return jsonify({'error': 'Güncellenecek alan yok'}), 400
    fields.append("guncelleme_tarihi=datetime('now','localtime')")
    vals.append(id)
    c.execute(f"UPDATE referans_takip SET {','.join(fields)} WHERE id=?", vals)
    conn.commit()
    conn.close()
    return jsonify({'basarili': True})

@app.route('/api/robot/<robot_no>/atamalar', methods=['GET'])
def robot_atama_listesi(robot_no):
    """Belirli bir robota atanmış aktif işleri getir."""
    conn = get_db()
    rows = conn.execute('''
        SELECT * FROM referans_takip 
        WHERE robot_no = ? AND durum != 'uretim_tamamlandi'
        ORDER BY guncelleme_tarihi DESC
    ''', (robot_no,)).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/api/referans_takip/<int:id>', methods=['DELETE'])
def referans_takip_sil(id):
    conn = get_db()
    conn.execute('DELETE FROM referans_takip WHERE id=?', (id,))
    conn.commit()
    conn.close()
    return jsonify({'basarili': True})


# ─────────────────────────────────────────────────────────────

# Bellekte tutulan kayan duyuru mesajı
_andon_mesaj = {'metin': '', 'yazar': ''}

@app.route('/api/andon', methods=['GET'])
def andon_veri():
    """Andon TV için bugünün tüm üretim verilerini tek sorguda döndürür. ?bolum= ile filtrelenebilir."""
    bugun = date.today().isoformat()
    bolum = request.args.get('bolum', '')
    conn = get_db()
    c = conn.cursor()

    # Bölüm filtresi
    bolum_sart = ''
    bolum_params = [bugun]
    if bolum:
        bolum_sart = ' AND v.bolum = ?'
        bolum_params = [bugun, bolum]

    # Bugünkü tüm vardiyalar
    vardiyalar = c.execute(
        f'SELECT * FROM vardiyalar v WHERE v.tarih = ?{bolum_sart} ORDER BY v.baslangic_saati',
        bolum_params
    ).fetchall()

    # Vardiya id'leri
    vardiya_ids = [v['id'] for v in vardiyalar]
    if not vardiya_ids:
        vardiya_ids_placeholder = '(-1)'
    else:
        vardiya_ids_placeholder = '(' + ','.join(str(v) for v in vardiya_ids) + ')'

    # Bugünkü tüm üretim kayıtları
    uretim_rows = c.execute(f'''
        SELECT u.* FROM uretim_kayitlari u
        WHERE u.vardiya_id IN {vardiya_ids_placeholder}
    ''').fetchall()

    # Bugünkü duruşlar
    durus_rows = c.execute(f'''
        SELECT d.* FROM duruslar d
        WHERE d.vardiya_id IN {vardiya_ids_placeholder}
        ORDER BY d.sure_dk DESC
    ''').fetchall()

    # Toplamlar
    toplam_ok  = sum(r['ok_adet']  for r in uretim_rows)
    toplam_nok = sum(r['nok_adet'] for r in uretim_rows)
    toplam_hedef = sum(r['hedef_adet'] for r in uretim_rows)
    toplam_uretim = toplam_ok + toplam_nok
    kalite_pct = round((toplam_ok / toplam_uretim * 100), 1) if toplam_uretim > 0 else 0

    # Toplam planlı süre ve duruş
    toplam_planli_dk = sum(v['toplam_sure_dk'] or 0 for v in vardiyalar)
    toplam_durus_dk  = sum(d['sure_dk'] or 0 for d in durus_rows)
    planli_durus_dk  = sum(d['sure_dk'] or 0 for d in durus_rows if d['durus_tipi'] == 'planli')
    plansiz_durus_dk = toplam_durus_dk - planli_durus_dk

    # Basit OEE (birleşik)
    if toplam_planli_dk > 0:
        kullanilabilirlik = max(0, (toplam_planli_dk - plansiz_durus_dk) / toplam_planli_dk * 100)
    else:
        kullanilabilirlik = 0

    # Performans (Adet bazlı basitleştirilmiş)
    performans = (toplam_ok / toplam_hedef * 100) if toplam_hedef > 0 else 0
    if performans > 100: performans = 100 # %100'ü geçmesin görsel amaçlı
    
    # OEE = K * P * Q
    oee = (kullanilabilirlik / 100) * (performans / 100) * (kalite_pct / 100) * 100
    
    # Plansız duruşlar (uyarı bandı için)
    plansiz_duruslar = [
        {'sebep': d['durus_sebebi'], 'sure_dk': d['sure_dk']}
        for d in durus_rows if d['durus_tipi'] == 'plansiz'
    ]

    # ── İş atamaları (aktif) — bölüme göre filtrelenir ──
    if bolum:
        atama_rows = c.execute('''
            SELECT id, istasyon, referans_kodu, aciklama, olusturan as atayan, robot_no
            FROM referans_takip
            WHERE robot_no != '' AND robot_no IS NOT NULL
              AND durum != 'uretim_tamamlandi'
              AND COALESCE(bolum, 'kaynak') = ?
            ORDER BY guncelleme_tarihi DESC
        ''', (bolum,)).fetchall()
    else:
        atama_rows = c.execute('''
            SELECT id, istasyon, referans_kodu, aciklama, olusturan as atayan, robot_no
            FROM referans_takip
            WHERE robot_no != '' AND robot_no IS NOT NULL
              AND durum != 'uretim_tamamlandi'
            ORDER BY guncelleme_tarihi DESC
        ''').fetchall()

    # Robot bazlı duruşları çek (bölüm filtresine uygun)
    robot_durus_rows = c.execute(f'''
        SELECT v.robot_no, COALESCE(SUM(d.sure_dk), 0) as toplam_durus_dk, COUNT(d.id) as durus_adet
        FROM vardiyalar v
        LEFT JOIN duruslar d ON d.vardiya_id = v.id
        WHERE v.id IN {vardiya_ids_placeholder}
        GROUP BY v.robot_no
    ''').fetchall()

    robot_durus_detay = c.execute(f'''
        SELECT v.robot_no, d.durus_sebebi, d.durus_tipi, d.sure_dk
        FROM duruslar d
        JOIN vardiyalar v ON d.vardiya_id = v.id
        WHERE v.id IN {vardiya_ids_placeholder}
        ORDER BY v.robot_no, d.sure_dk DESC
    ''').fetchall()

    robot_durus_map = {}
    for r in robot_durus_rows:
        robot_durus_map[r['robot_no']] = {'toplam_durus_dk': r['toplam_durus_dk'], 'durus_adet': r['durus_adet'], 'detay': []}
    for d in robot_durus_detay:
        rn = d['robot_no']
        if rn not in robot_durus_map: robot_durus_map[rn] = {'toplam_durus_dk': 0, 'durus_adet': 0, 'detay': []}
        robot_durus_map[rn]['detay'].append({'sebep': d['durus_sebebi'], 'sure_dk': d['sure_dk'], 'tip': d['durus_tipi']})

    # Vardiya bazlı duruş haritası (her vardiya için ayrı duruş istatistiği)
    # — montaj/metal'de aynı hat/makinede birden fazla vardiya açıkken duruşları
    # operatörler arasında karıştırmamak için.
    vardiya_durus_rows = c.execute(f'''
        SELECT v.id as vid, COALESCE(SUM(d.sure_dk), 0) as toplam_durus_dk, COUNT(d.id) as durus_adet
        FROM vardiyalar v
        LEFT JOIN duruslar d ON d.vardiya_id = v.id
        WHERE v.id IN {vardiya_ids_placeholder}
        GROUP BY v.id
    ''').fetchall()
    vardiya_durus_detay = c.execute(f'''
        SELECT d.vardiya_id as vid, d.durus_sebebi, d.durus_tipi, d.sure_dk
        FROM duruslar d
        WHERE d.vardiya_id IN {vardiya_ids_placeholder}
        ORDER BY d.vardiya_id, d.sure_dk DESC
    ''').fetchall()
    vardiya_durus_map = {}
    for r in vardiya_durus_rows:
        vardiya_durus_map[r['vid']] = {'toplam_durus_dk': r['toplam_durus_dk'], 'durus_adet': r['durus_adet'], 'detay': []}
    for d in vardiya_durus_detay:
        vid = d['vid']
        if vid not in vardiya_durus_map:
            vardiya_durus_map[vid] = {'toplam_durus_dk': 0, 'durus_adet': 0, 'detay': []}
        vardiya_durus_map[vid]['detay'].append({'sebep': d['durus_sebebi'], 'sure_dk': d['sure_dk'], 'tip': d['durus_tipi']})

    # ── aktif_vardiyalar: her açık vardiya = ayrı kart ──
    # Bu, andon ekranında aynı hat/makineye birden fazla operatör açtığında
    # her operatörün kendi kartını ve kendi üretim/duruş kayıtlarını görmesini sağlar.
    aktif_vardiyalar = []
    for v in vardiyalar:
        if v['durum'] != 'acik':
            continue
        durus_info = vardiya_durus_map.get(v['id'], {'toplam_durus_dk': 0, 'durus_adet': 0, 'detay': []})
        item = {
            'vardiya_id': v['id'],
            'robot_no': v['robot_no'],
            'operator': v['operator_adi'],
            'vardiya': v['vardiya_turu'],
            'baslangic': v['baslangic_saati'],
            'bitis': v['bitis_saati'],
            'istasyon_1': [], 'istasyon_2': [], 'diger': [], 'atamalar': [],
            'durus_dk': durus_info['toplam_durus_dk'],
            'durus_adet': durus_info['durus_adet'],
            'durus_detay': durus_info['detay'],
        }
        for u in uretim_rows:
            if u['vardiya_id'] == v['id']:
                row = {'ref': u['referans_kodu'], 'launch': u['launch_adet'] or 0, 'tamamlandi': 1 if u['tamamlandi'] else 0}
                ist = u['istasyon'] or 0
                if ist == 1:   item['istasyon_1'].append(row)
                elif ist == 2: item['istasyon_2'].append(row)
                else:          item['diger'].append(row)
        aktif_vardiyalar.append(item)

    # Atamaları her uygun shift kartına ekle (aynı robot_no'lu kartlara)
    # Montaj'da "atama" kavramı kullanılmıyor — operatöre sıradaki iş olarak
    # bireysel atama yansımaz (öncelik listesi yalnızca yönetici görünümünde anlamlı).
    if bolum != 'montaj':
        for a in atama_rows:
            rn = a['robot_no']
            atama_item = {'id': a['id'], 'istasyon': a['istasyon'], 'referans_kodu': a['referans_kodu'], 'aciklama': a['aciklama'], 'atayan': a['atayan']}
            for it in aktif_vardiyalar:
                if it['robot_no'] == rn:
                    it['atamalar'].append(atama_item)

    # Eski 'robotlar' alanı (legacy andon.html için backward compat — robot bazında ilk vardiya)
    robotlar = {}
    for v in vardiyalar:
        if v['durum'] != 'acik':
            continue
        r = v['robot_no']
        durus_info = robot_durus_map.get(r, {'toplam_durus_dk': 0, 'durus_adet': 0, 'detay': []})
        if r not in robotlar:
            robotlar[r] = {
                'robot_no': r, 'operator': v['operator_adi'], 'vardiya': v['vardiya_turu'],
                'baslangic': v['baslangic_saati'], 'bitis': v['bitis_saati'],
                'istasyon_1': [], 'istasyon_2': [], 'diger': [], 'atamalar': [],
                'durus_dk': durus_info['toplam_durus_dk'], 'durus_adet': durus_info['durus_adet'],
                'durus_detay': durus_info.get('detay', [])
            }
        for u in uretim_rows:
            if u['vardiya_id'] == v['id']:
                row = {'ref': u['referans_kodu'], 'launch': u['launch_adet'] or 0, 'tamamlandi': 1 if u['tamamlandi'] else 0}
                ist = u['istasyon'] or 0
                if ist == 1: robotlar[r]['istasyon_1'].append(row)
                elif ist == 2: robotlar[r]['istasyon_2'].append(row)
                else: robotlar[r]['diger'].append(row)

    for a in atama_rows:
        rn = a['robot_no']
        if rn not in robotlar:
            robotlar[rn] = {'robot_no':rn, 'operator':'', 'vardiya':'', 'baslangic':'', 'bitis':'', 'istasyon_1':[], 'istasyon_2':[], 'diger':[], 'atamalar':[]}
        robotlar[rn]['atamalar'].append({'id': a['id'], 'istasyon': a['istasyon'], 'referans_kodu': a['referans_kodu'], 'aciklama': a['aciklama'], 'atayan': a['atayan']})

    # Referans takip listesi — bölüme göre filtrelenir
    if bolum:
        rt_rows = c.execute('''
            SELECT rt.*, rl.hedef_cycle_time_sn
            FROM referans_takip rt
            LEFT JOIN referans_listesi rl ON REPLACE(rt.referans_kodu, ' ', '') = REPLACE(rl.referans_kodu, ' ', '')
            WHERE COALESCE(rt.bolum, 'kaynak') = ?
            ORDER BY rt.olusturma_tarihi DESC
        ''', (bolum,)).fetchall()
    else:
        rt_rows = c.execute('''
            SELECT rt.*, rl.hedef_cycle_time_sn
            FROM referans_takip rt
            LEFT JOIN referans_listesi rl ON REPLACE(rt.referans_kodu, ' ', '') = REPLACE(rl.referans_kodu, ' ', '')
            ORDER BY rt.olusturma_tarihi DESC
        ''').fetchall()
    referans_takip_list = [dict(row) for row in rt_rows]

    # Tum ayarlari yukle
    ayarlar_rows = c.execute("SELECT anahtar, deger FROM genel_ayarlar").fetchall()
    
    response_data = {
        'tarih': bugun,
        'toplam_ok': toplam_ok,
        'toplam_nok': toplam_nok,
        'toplam_hedef': toplam_hedef,
        'kalite_pct': kalite_pct,
        'kullanilabilirlik_pct': round(kullanilabilirlik, 1),
        'performans_pct': round(performans, 1),
        'oee_pct': round(oee, 1),
        'toplam_durus_dk': toplam_durus_dk,
        'plansiz_durus_dk': plansiz_durus_dk,
        'vardiya_sayisi': len(vardiyalar),
        'robotlar': list(robotlar.values()),
        'aktif_vardiyalar': aktif_vardiyalar,
        'plansiz_duruslar': plansiz_duruslar,
        'referans_takip': referans_takip_list,
        'duyuru': _andon_mesaj
    }
    
    # Ayarlari response_data'ya entegre et
    for row in ayarlar_rows:
        response_data[row['anahtar']] = row['deger']
        
    return jsonify(response_data)




@app.route('/api/andon/mesaj', methods=['POST'])
def andon_mesaj_guncelle():
    """Andon TV kayan duyuru mesajını güncelle."""
    data = request.get_json() or {}
    _andon_mesaj['metin'] = data.get('metin', '')
    _andon_mesaj['yazar'] = data.get('yazar', '')
    return jsonify({'basarili': True})


# ─────────────────────────────────────────────────────────────
# GÜNLÜK ÜRETİM RAPORU (Excel Görünümü) API
# ─────────────────────────────────────────────────────────────

@app.route('/api/rapor/gunluk_detay', methods=['GET'])
def gunluk_rapor_detay():
    """Belirli bir tarih, vardiya ve bolum icin uretim detaylarini getirir.
    Sira (key) bolume gore degisir:
      - kaynak: 1..9, M
      - montaj: o gun calisan distinct hat'lar (HAT 1, HAT 2 ...)
      - metal:  300T, 400T, 500T, Şerit Testere
    """
    tarih = request.args.get('tarih')
    vardiya_turu = request.args.get('vardiya')
    bolum = (request.args.get('bolum') or 'kaynak').strip()
    if bolum not in ('kaynak', 'montaj', 'metal'):
        bolum = 'kaynak'

    if not tarih or not vardiya_turu:
        return jsonify({'hata': 'tarih ve vardiya parametreleri zorunludur'}), 400

    try:
        conn = get_db()

        rows = conn.execute("""
            SELECT
                v.robot_no,
                v.operator_adi,
                u.referans_kodu,
                SUM(u.ok_adet)    as ok_toplam,
                SUM(u.nok_adet)   as nok_toplam,
                SUM(u.tamir_adet) as tamir_toplam
            FROM vardiyalar v
            LEFT JOIN uretim_kayitlari u ON v.id = u.vardiya_id
            WHERE v.tarih = ?
              AND UPPER(REPLACE(REPLACE(v.vardiya_turu,'ü','u'),'Ü','U')) =
                  UPPER(REPLACE(REPLACE(?,'ü','u'),'Ü','U'))
              AND COALESCE(v.bolum, 'kaynak') = ?
            GROUP BY v.robot_no, v.operator_adi, u.referans_kodu
            ORDER BY v.robot_no, v.operator_adi
        """, (tarih, vardiya_turu, bolum)).fetchall()
        conn.close()

        # Bolume gore robot/hat/makine sirasi
        if bolum == 'kaynak':
            robot_listesi = ['1','2','3','4','5','6','7','8','9','M']
        elif bolum == 'metal':
            robot_listesi = ['300T', '400T', '500T', 'Şerit Testere']
        else:  # montaj
            # O gun calisan distinct hat'lari sirayla
            seen = []
            for r in rows:
                rn = (r['robot_no'] or '').strip()
                if rn and rn not in seen:
                    seen.append(rn)
            robot_listesi = seen

        rapor_data = {r: [] for r in robot_listesi}

        def normalize_robot_kaynak(r_no):
            r = str(r_no or '').strip().upper()
            for prefix in ('ABB-', 'ABB', 'ROBOT-', 'ROBOT'):
                if r.startswith(prefix):
                    tail = r[len(prefix):].strip()
                    if tail in robot_listesi:
                        return tail
            if r == 'M':
                return 'M'
            return r

        for row in rows:
            if bolum == 'kaynak':
                target_r = normalize_robot_kaynak(row['robot_no'])
            else:
                target_r = (row['robot_no'] or '').strip()
            if not target_r:
                continue

            if not row['referans_kodu'] and (row['ok_toplam'] or 0) == 0:
                continue

            entry = {
                'operator':   row['operator_adi'] or '',
                'parca_kodu': row['referans_kodu'] or '',
                'adet':       row['ok_toplam']    or 0,
                'tamir':      row['tamir_toplam'] or 0,
                'hurda':      row['nok_toplam']   or 0,
            }

            if target_r not in rapor_data:
                rapor_data[target_r] = []
            rapor_data[target_r].append(entry)

        return jsonify({
            'tarih':   tarih,
            'vardiya': vardiya_turu,
            'bolum':   bolum,
            'siralama': robot_listesi,
            'data':    rapor_data,
        })
    except Exception as e:
        if 'conn' in locals():
            conn.close()
        return jsonify({'hata': f'Rapor hatasi: {str(e)}'}), 500


@app.route('/api/rapor/vardiya_listesi', methods=['GET'])
def rapor_vardiya_listesi():
    """Belirtilen tarihteki ve bolumdeki mevcut vardiya turlerini dondurur."""
    tarih = request.args.get('tarih')
    bolum = request.args.get('bolum')
    if not tarih:
        return jsonify({'hata': 'tarih zorunludur'}), 400
    try:
        conn = get_db()
        if bolum:
            rows = conn.execute(
                "SELECT DISTINCT vardiya_turu FROM vardiyalar WHERE tarih = ? AND COALESCE(bolum, 'kaynak') = ? ORDER BY vardiya_turu",
                (tarih, bolum)
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT DISTINCT vardiya_turu FROM vardiyalar WHERE tarih = ? ORDER BY vardiya_turu",
                (tarih,)
            ).fetchall()
        conn.close()
        return jsonify([r['vardiya_turu'] for r in rows])
    except Exception as e:
        return jsonify({'hata': str(e)}), 500


# ─────────────────────────────────────────────────────────────

# BAŞLATMA
# ─────────────────────────────────────────────────────────────

if __name__ == '__main__':
    # Flask debug modundayken (reloader) çift çalışmayı önlemek için kontrol
    if os.environ.get('WERKZEUG_RUN_MAIN') != 'true':
        init_db()
    
    print("\n" + "="*55)
    print("  COFLE MANAGE - URETIM TAKIP SISTEMI CALISIYOR")
    print("="*55)
    print("  Operator Formu : https://coflemanage.online")
    print("  Yonetici Panel : https://coflemanage.online/dashboard")
    print("  Andon Ekrani   : https://coflemanage.online/andon")
    print("="*55)
    print("  Sistem artik bu adresten yayinlanmaktadir.")
    print("="*55 + "\n")
    app.run(host='0.0.0.0', port=5000, debug=True)
