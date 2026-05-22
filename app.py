from flask import Flask, request, jsonify, render_template, send_file, g
from flask_cors import CORS
from datetime import datetime, date
from functools import wraps
import json
import os
import traceback # For debugging

from database import get_db as db_connect, init_db
from oee import hesapla_oee, hesapla_oee_ozet
from import_excel import import_data, durus_sebepleri_yukle, import_tum, export_referans_cycle_times
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
# OPERATÖR YETKİLENDİRME (PIN tabanlı)
# Operatör mobile'dan gelen istekler X-Operator + X-Operator-Pin
# header'larini icermeli. Dashboard/yonetici cagrilarinda bu
# header'lar yoktur ve serbest erisim verilir (geri uyumluluk).
# ─────────────────────────────────────────────────────────────
def operator_required(f):
    """Header'da X-Operator + X-Operator-Pin varsa PIN doğrular.
    Yoksa (dashboard/yönetici) izin verir ama g.operator_adi None olur.

    X-Operator header'i frontend'de encodeURIComponent ile encode edilir
    (HTTP header'lar ISO-8859-1 olmali, Turkce karakterler bozulur).
    """
    from urllib.parse import unquote
    @wraps(f)
    def wrapper(*args, **kwargs):
        op_raw = (request.headers.get('X-Operator', '') or '').strip()
        op = unquote(op_raw) if op_raw else ''
        pin = (request.headers.get('X-Operator-Pin', '') or '').strip()
        g.operator_adi = None
        if op and pin:
            conn = get_db()
            row = conn.execute('SELECT pin FROM operatorler WHERE ad=?', (op,)).fetchone()
            if not row:
                return jsonify({'hata': f'Operatör bulunamadı: {op}'}), 403
            if (row['pin'] or '0000') != pin:
                return jsonify({'hata': 'PIN yanlış'}), 403
            g.operator_adi = op
        elif op or pin:
            return jsonify({'hata': 'Operatör adı ve PIN birlikte gönderilmeli'}), 403
        return f(*args, **kwargs)
    return wrapper


def _vardiya_sahibi_kontrol(vardiya_id):
    """g.operator_adi varsa vardiya o operatöre ait mi kontrol eder.
    g.operator_adi None ise (dashboard) izin verir.
    Returns: (ok: bool, hata_mesaji: str|None)
    """
    if not getattr(g, 'operator_adi', None):
        return True, None  # dashboard modu — operatör yetkisi gerekmez
    conn = get_db()
    row = conn.execute('SELECT operator_adi FROM vardiyalar WHERE id=?', (vardiya_id,)).fetchone()
    if not row:
        return False, f'Vardiya bulunamadı (id={vardiya_id})'
    if row['operator_adi'] != g.operator_adi:
        return False, f'Bu vardiya başka operatöre ait: {row["operator_adi"]}'
    return True, None


def _uretim_vardiya_bul(uretim_id):
    """uretim_kayit id'den vardiya_id getirir (yoksa None)."""
    conn = get_db()
    row = conn.execute('SELECT vardiya_id FROM uretim_kayitlari WHERE id=?', (uretim_id,)).fetchone()
    return row['vardiya_id'] if row else None


def _durus_vardiya_bul(durus_id):
    """durus id'den vardiya_id getirir (yoksa None)."""
    conn = get_db()
    row = conn.execute('SELECT vardiya_id FROM duruslar WHERE id=?', (durus_id,)).fetchone()
    return row['vardiya_id'] if row else None


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
@operator_required
def vardiya_ekle():
    """Yeni vardiya kaydı oluştur.
    Operatör mobile'dan gelirse X-Operator header'ı body'deki operator_adi ile eşleşmeli.
    """
    data = request.get_json()

    zorunlu = ['tarih', 'vardiya_turu', 'robot_no', 'operator_adi', 'baslangic_saati', 'bitis_saati']
    for alan in zorunlu:
        if not data.get(alan):
            return jsonify({'hata': f'"{alan}" alanı zorunludur'}), 400

    # Operatör mobile'dan gelen istekte header'daki operator = body'deki operator olmalı
    # (başkasının adına vardiya açma engeli). Dashboard modunda g.operator_adi None.
    if getattr(g, 'operator_adi', None) and g.operator_adi != data['operator_adi']:
        return jsonify({'hata': f'Sadece kendi adınıza vardiya açabilirsiniz ({g.operator_adi})'}), 403

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
@operator_required
def vardiya_sil(vid):
    """Vardiyayi sil (uretim ve duruslarla birlikte).
    Operatör sadece kendi vardiyasını silebilir.
    """
    ok, hata = _vardiya_sahibi_kontrol(vid)
    if not ok:
        return jsonify({'hata': hata}), 403
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
@operator_required
def vardiya_kapat(vid):
    """Vardiyayi kapat ve bitis saatini guncelle.
    Operatör sadece kendi vardiyasını kapatabilir.
    """
    ok, hata = _vardiya_sahibi_kontrol(vid)
    if not ok:
        return jsonify({'hata': hata}), 403
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


@app.route('/api/vardiya/<int:vid>/robotla_calisiyor', methods=['PATCH'])
@operator_required
def vardiya_robotla_calisiyor(vid):
    """Vardiya 'robotla çalışıyor / tam otomasyon' bayrağını aç/kapat.
    Metal enjeksiyonda makine+robot otomatik çalışırken kullanılır.
    Body: { robotla_calisiyor: 0|1 }
    """
    ok, hata = _vardiya_sahibi_kontrol(vid)
    if not ok:
        return jsonify({'hata': hata}), 403
    data = request.get_json() or {}
    yeni = 1 if data.get('robotla_calisiyor') else 0
    conn = get_db()
    c = conn.cursor()
    v = c.execute('SELECT id FROM vardiyalar WHERE id=?', (vid,)).fetchone()
    if not v:
        conn.close()
        return jsonify({'hata': 'Vardiya bulunamadı'}), 404
    c.execute('UPDATE vardiyalar SET robotla_calisiyor=? WHERE id=?', (yeni, vid))
    conn.commit()
    conn.close()
    return jsonify({'basarili': True, 'robotla_calisiyor': yeni})


@app.route('/api/vardiya/<int:vid>/ac', methods=['PATCH'])
@operator_required
def vardiya_ac(vid):
    """Kapanmış vardiyayı yeniden aç."""
    ok, hata = _vardiya_sahibi_kontrol(vid)
    if not ok:
        return jsonify({'hata': hata}), 403
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
@operator_required
def vardiya_guncelle(vid):
    """Vardiya saatlerini ve detaylarını güncelle."""
    ok, hata = _vardiya_sahibi_kontrol(vid)
    if not ok:
        return jsonify({'hata': hata}), 403
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
@operator_required
def uretim_ekle():
    """Üretim kaydı ekle."""
    data = request.get_json() or {}

    if not data.get('vardiya_id') or not data.get('referans_kodu'):
        return jsonify({'hata': 'vardiya_id ve referans_kodu zorunludur'}), 400

    # Operatör yetki kontrolü: header'da operator+pin geldiyse vardiya sahibi mi?
    ok, hata = _vardiya_sahibi_kontrol(int(data['vardiya_id']))
    if not ok:
        return jsonify({'hata': hata}), 403

    conn = None
    try:
        conn = get_db()
        c = conn.cursor()

        # Vardiyanın bölümünü öğren — yeni eklenen referansların doğru bölüme yazılması için
        v_row = c.execute(
            "SELECT COALESCE(bolum, 'kaynak') as bolum FROM vardiyalar WHERE id = ?",
            (data['vardiya_id'],)
        ).fetchone()
        vardiya_bolum = v_row['bolum'] if v_row else 'kaynak'

        # Birden fazla kayıt gelebilir
        satirlar = data.get('satirlar', [data])

        eklenen = 0
        for satir in satirlar:
            ref = (satir.get('referans_kodu') or data.get('referans_kodu') or '').strip()
            ct_in = float(satir.get('cycle_time_sn', 0) or 0)
            # cycle_time gönderilmediyse ya da 0 ise referanslardan otomatik çek
            if ct_in <= 0 and ref:
                ref_row = c.execute(
                    "SELECT hedef_cycle_time_sn FROM referans_listesi WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ',''))",
                    (ref,)
                ).fetchone()
                if ref_row:
                    ct_in = float(ref_row['hedef_cycle_time_sn'] or 0)

            c.execute('''
                INSERT INTO uretim_kayitlari (vardiya_id, referans_kodu, ok_adet, nok_adet, tamir_adet, hedef_adet, cycle_time_sn, istasyon, launch_adet, aciklama)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                data['vardiya_id'],
                ref,
                int(satir.get('ok_adet', 0) or 0),
                int(satir.get('nok_adet', 0) or 0),
                int(satir.get('tamir_adet', 0) or 0),
                int(satir.get('hedef_adet', 0) or 0),
                ct_in,
                int(satir.get('istasyon', data.get('istasyon', 0)) or 0),
                int(satir.get('launch_adet', data.get('launch_adet', 0)) or 0),
                (satir.get('aciklama') or data.get('aciklama') or '').strip()
            ))
            # Referansı listeye otomatik ekle — VARDIYANIN bölümüyle etiketle
            # (montaj operatörü tanımsız bir kod girince montaj'a düşsün, kaynak'a değil)
            c.execute(
                'INSERT OR IGNORE INTO referans_listesi (referans_kodu, bolum) VALUES (?, ?)',
                (ref, vardiya_bolum)
            )
            eklenen += 1

        conn.commit()
        return jsonify({'basarili': True, 'eklenen': eklenen}), 201
    except Exception as e:
        print(f"HATA (uretim_ekle): {traceback.format_exc()}")
        return jsonify({'hata': f'Üretim eklenemedi: {str(e)}'}), 500
    finally:
        if conn:
            conn.close()


@app.route('/api/uretim/<int:uid>', methods=['DELETE'])
@operator_required
def uretim_sil(uid):
    vid = _uretim_vardiya_bul(uid)
    if vid:
        ok, hata = _vardiya_sahibi_kontrol(vid)
        if not ok:
            return jsonify({'hata': hata}), 403
    conn = get_db()
    conn.execute('DELETE FROM uretim_kayitlari WHERE id = ?', (uid,))
    conn.commit()
    conn.close()
    return jsonify({'basarili': True})


@app.route('/api/uretim/<int:uid>', methods=['PUT'])
@operator_required
def uretim_guncelle(uid):
    """Uretim kaydini guncelle."""
    vid = _uretim_vardiya_bul(uid)
    if vid:
        ok, hata = _vardiya_sahibi_kontrol(vid)
        if not ok:
            return jsonify({'hata': hata}), 403
    data = request.get_json() or {}
    conn = None
    try:
        conn = get_db()
        c = conn.cursor()
        mevcut = c.execute('SELECT id FROM uretim_kayitlari WHERE id = ?', (uid,)).fetchone()
        if not mevcut:
            return jsonify({'hata': 'Kayit bulunamadi'}), 404

        ref = (data.get('referans_kodu') or '').strip()
        ct = float(data.get('cycle_time_sn', 0) or 0)
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
            SET referans_kodu=?, ok_adet=?, nok_adet=?, tamir_adet=?, hedef_adet=?, cycle_time_sn=?, istasyon=?, launch_adet=?, aciklama=?
            WHERE id=?
        ''', (
            ref,
            int(data.get('ok_adet', 0) or 0),
            int(data.get('nok_adet', 0) or 0),
            int(data.get('tamir_adet', 0) or 0),
            int(data.get('hedef_adet', 0) or 0),
            ct,
            int(data.get('istasyon', 0) or 0),
            int(data.get('launch_adet', 0) or 0),
            (data.get('aciklama') or '').strip(),
            uid
        ))
        conn.commit()
        return jsonify({'basarili': True})
    except Exception as e:
        print(f"HATA (uretim_guncelle): {traceback.format_exc()}")
        return jsonify({'hata': f'Üretim güncellenemedi: {str(e)}'}), 500
    finally:
        if conn:
            conn.close()


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
@operator_required
def uretim_tamamlandi_guncelle(uid):
    """Üretim kaydını tamamlandı / tamamlanmadı olarak işaretle."""
    vid = _uretim_vardiya_bul(uid)
    if vid:
        ok, hata = _vardiya_sahibi_kontrol(vid)
        if not ok:
            return jsonify({'hata': hata}), 403
    data = request.get_json() or {}
    deger = 1 if data.get('tamamlandi') else 0
    conn = get_db()
    conn.execute('UPDATE uretim_kayitlari SET tamamlandi=? WHERE id=?', (deger, uid))
    conn.commit()
    conn.close()
    return jsonify({'basarili': True})

@app.route('/api/durus', methods=['POST'])
@operator_required
def durus_ekle():
    """Duruş kaydı ekle. Aynı vardiya + aynı durus_sebebi için zaten kayıt
    varsa: yenisini eklemek yerine süreyi mevcut kayda ekler (toplama yapar).
    Açıklama varsa eski açıklamaya ' | ' ile ayırarak sona eklenir."""
    data = request.get_json() or {}

    if not data.get('vardiya_id') or not data.get('durus_sebebi'):
        return jsonify({'hata': 'vardiya_id ve durus_sebebi zorunludur'}), 400

    # Operatör yetki kontrolü: vardiya sahibi mi?
    ok, hata = _vardiya_sahibi_kontrol(int(data['vardiya_id']))
    if not ok:
        return jsonify({'hata': hata}), 403

    # Birden fazla duruş gelebilir
    satirlar = data.get('satirlar', [data])

    conn = None
    try:
        conn = get_db()
        c = conn.cursor()
        eklenen = 0
        birlestirilen = 0
        for satir in satirlar:
            vid    = data['vardiya_id']
            sebep  = (satir.get('durus_sebebi') or data.get('durus_sebebi') or '').strip()
            if not sebep:
                continue
            sure   = int(satir.get('sure_dk', 0) or 0)
            saat   = satir.get('baslangic_saati', '') or ''
            acikl  = (satir.get('aciklama') or '').strip()
            tipi   = satir.get('durus_tipi', data.get('durus_tipi', 'plansiz'))

            # Aynı vardiya + aynı sebep var mı?
            mevcut = c.execute(
                "SELECT id, sure_dk, aciklama FROM duruslar WHERE vardiya_id=? AND durus_sebebi=? LIMIT 1",
                (vid, sebep)
            ).fetchone()

            if mevcut:
                yeni_sure = (mevcut['sure_dk'] or 0) + sure
                # Açıklamaları birleştir (boş olanı atla)
                eski_acikl = (mevcut['aciklama'] or '').strip()
                if eski_acikl and acikl:
                    yeni_acikl = eski_acikl + ' | ' + acikl
                else:
                    yeni_acikl = eski_acikl or acikl
                c.execute(
                    "UPDATE duruslar SET sure_dk=?, aciklama=?, durus_tipi=? WHERE id=?",
                    (yeni_sure, yeni_acikl, tipi, mevcut['id'])
                )
                birlestirilen += 1
            else:
                c.execute('''
                    INSERT INTO duruslar (vardiya_id, durus_sebebi, aciklama, sure_dk, baslangic_saati, durus_tipi)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (vid, sebep, acikl, sure, saat, tipi))
                eklenen += 1

        conn.commit()
        return jsonify({'basarili': True, 'eklenen': eklenen, 'birlestirilen': birlestirilen}), 201
    except Exception as e:
        print(f"HATA (durus_ekle): {traceback.format_exc()}")
        return jsonify({'hata': f'Duruş eklenemedi: {str(e)}'}), 500
    finally:
        if conn:
            conn.close()


@app.route('/api/durus/<int:did>', methods=['DELETE'])
@operator_required
def durus_sil(did):
    vid = _durus_vardiya_bul(did)
    if vid:
        ok, hata = _vardiya_sahibi_kontrol(vid)
        if not ok:
            return jsonify({'hata': hata}), 403
    conn = get_db()
    conn.execute('DELETE FROM duruslar WHERE id = ?', (did,))
    conn.commit()
    conn.close()
    return jsonify({'basarili': True})


@app.route('/api/durus/<int:did>', methods=['PUT'])
@operator_required
def durus_guncelle(did):
    """Durus kaydini guncelle."""
    vid = _durus_vardiya_bul(did)
    if vid:
        ok, hata = _vardiya_sahibi_kontrol(vid)
        if not ok:
            return jsonify({'hata': hata}), 403
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
    """Referans autocomplete listesi. ?bolum= ile filtrelenebilir.
    Kaynak için kaynak_suresi_sn ve soktak_suresi_sn alanlarını da döner."""
    q = request.args.get('q', '')
    bolum = request.args.get('bolum', '')
    conn = get_db()
    base = "SELECT referans_kodu, aciklama, hedef_cycle_time_sn, kaynak_suresi_sn, soktak_suresi_sn, sure_teyit, sure_teyit_tarihi FROM referans_listesi"
    if bolum:
        rows = conn.execute(
            base + " WHERE REPLACE(referans_kodu, ' ', '') LIKE REPLACE(?, ' ', '') AND bolum = ? ORDER BY referans_kodu",
            (f'%{q}%', bolum)
        ).fetchall()
    else:
        rows = conn.execute(
            base + " WHERE REPLACE(referans_kodu, ' ', '') LIKE REPLACE(?, ' ', '') ORDER BY referans_kodu",
            (f'%{q}%',)
        ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/api/referanslar', methods=['POST'])
def referans_ekle():
    """Yeni referans tanımla veya güncelle (bolum opsiyonel — varsayılan 'kaynak').
    Cycle time tanımlandığında Excel'in ilgili bölüm sayfası otomatik
    güncellenir (data/uretim_verileri.xlsx).
    """
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

    # Excel'i otomatik güncelle (sadece bu bölümün cycle time'larını yazar,
    # diğer sayfalar dokunulmaz). Hata olsa bile referans kaydı başarılı sayılır.
    try:
        export_referans_cycle_times(bolum=bolum)
    except Exception as e:
        print(f'[referans_ekle] Excel auto-sync hatası ({bolum}): {e}')

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
    """Veritabanındaki referans listesini data/uretim_verileri.xlsx'in
    aktif bölüm sayfasına yazar.

    Body: { bolum: 'kaynak' | 'montaj' | 'metal' } — yoksa kaynak.
    """
    BOLUM_SAYFA = {
        'kaynak': 'Kaynak Referans',
        'montaj': 'Montaj Referans',
        'metal':  'Metal Referans',
    }
    body = request.get_json(silent=True) or {}
    bolum = body.get('bolum', 'kaynak')
    if bolum not in BOLUM_SAYFA:
        return jsonify({'hata': f"Geçersiz bölüm: {bolum}"}), 400
    sayfa_adi = BOLUM_SAYFA[bolum]

    excel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'uretim_verileri.xlsx')
    if not os.path.exists(excel_path):
        return jsonify({'hata': 'data/uretim_verileri.xlsx bulunamadı.'}), 404

    try:
        import openpyxl

        conn = get_db()
        rows = conn.execute(
            "SELECT referans_kodu, hedef_cycle_time_sn FROM referans_listesi "
            "WHERE COALESCE(bolum, 'kaynak') = ? ORDER BY referans_kodu",
            (bolum,)
        ).fetchall()
        conn.close()

        wb = openpyxl.load_workbook(excel_path)
        if sayfa_adi not in wb.sheetnames:
            return jsonify({'hata': f"'{sayfa_adi}' sayfası bulunamadı"}), 404
        sayfa = wb[sayfa_adi]
        
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
    - metal: 300T, 400T, 550T, Şerit Testere (sabit)
    """
    bolum = request.args.get('bolum', 'kaynak')
    if bolum == 'metal':
        robotlar = ['300T', '400T', '550T', 'Şerit Testere']
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
    """Operatör listesini döndür. ?bolum= ile filtrelenebilir.
    PIN bilgisi bu endpoint'te VERİLMEZ (operatör mobile bunu çağırır).
    Yönetim için /api/operatorler/yonetim kullanın.
    """
    bolum = request.args.get('bolum', '')
    conn = get_db()
    if bolum:
        rows = conn.execute('SELECT id, ad, bolum FROM operatorler WHERE bolum = ? ORDER BY ad', (bolum,)).fetchall()
    else:
        rows = conn.execute('SELECT id, ad, bolum FROM operatorler ORDER BY ad').fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/api/operatorler/yonetim', methods=['GET'])
def operator_listesi_yonetim():
    """Yönetici için operatör listesi + PIN bilgisi.
    Dashboard'da PIN tanımlama panelinde kullanılır.
    """
    conn = get_db()
    rows = conn.execute('SELECT id, ad, bolum, pin FROM operatorler ORDER BY bolum, ad').fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/api/operator/oturum_ac', methods=['POST'])
def operator_oturum_ac():
    """Operatör mobile'da PIN ile oturum açma.
    Body: {operator_adi, pin}
    Return 200 ok / 403 hata
    """
    data = request.get_json() or {}
    op = (data.get('operator_adi') or '').strip()
    pin = (data.get('pin') or '').strip()
    if not op or not pin:
        return jsonify({'hata': 'operator_adi ve pin zorunlu'}), 400
    conn = get_db()
    row = conn.execute('SELECT pin, bolum FROM operatorler WHERE ad=?', (op,)).fetchone()
    conn.close()
    if not row:
        return jsonify({'hata': f'Operatör bulunamadı: {op}'}), 403
    if (row['pin'] or '0000') != pin:
        return jsonify({'hata': 'PIN yanlış'}), 403
    return jsonify({
        'basarili':     True,
        'operator_adi': op,
        'bolum':        row['bolum'] or '',
    }), 200


@app.route('/api/operator/<int:oid>/pin', methods=['PATCH'])
def operator_pin_guncelle(oid):
    """Yönetici operatörün PIN'ini günceller. Body: {pin}
    PIN 4 haneli rakam olmalıdır.
    """
    data = request.get_json() or {}
    yeni_pin = (data.get('pin') or '').strip()
    if not yeni_pin or len(yeni_pin) != 4 or not yeni_pin.isdigit():
        return jsonify({'hata': "PIN 4 haneli rakam olmalı (örn: '1234')"}), 400
    conn = get_db()
    row = conn.execute('SELECT ad FROM operatorler WHERE id=?', (oid,)).fetchone()
    if not row:
        conn.close()
        return jsonify({'hata': 'Operatör bulunamadı'}), 404
    conn.execute('UPDATE operatorler SET pin=? WHERE id=?', (yeni_pin, oid))
    conn.commit()
    conn.close()
    return jsonify({'basarili': True, 'operator_adi': row['ad']}), 200


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


@app.route('/api/veri/import_tum', methods=['POST'])
def veri_import_tum():
    """Tüm Excel sayfalarını okuyup DB'yi günceller (tek tıkla import).
    Kapsam: referans + operatör (3 bölüm), robot program, fikstür raf.
    Duruş sebepleri her API isteğinde Excel'den okunur, import gerekmez.
    """
    try:
        sonuc = import_tum()
        return jsonify(sonuc), 200
    except Exception as e:
        return jsonify({'hata': str(e), 'basarili': False}), 500


@app.route('/api/referans/teyit', methods=['PATCH'])
def referans_teyit_guncelle():
    """Bir referansın süre teyidi bayrağını aç/kapat.
    Body: {referans_kodu, sure_teyit: 0|1}
    """
    data = request.get_json() or {}
    kod = (data.get('referans_kodu') or '').strip()
    teyit = 1 if data.get('sure_teyit') else 0
    if not kod:
        return jsonify({'hata': 'referans_kodu zorunlu'}), 400
    conn = get_db()
    try:
        if teyit == 1:
            conn.execute(
                "UPDATE referans_listesi SET sure_teyit=1, sure_teyit_tarihi=datetime('now','localtime') "
                "WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ',''))",
                (kod,)
            )
        else:
            conn.execute(
                "UPDATE referans_listesi SET sure_teyit=0, sure_teyit_tarihi=NULL "
                "WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ',''))",
                (kod,)
            )
        conn.commit()
    finally:
        conn.close()
    return jsonify({'basarili': True, 'referans_kodu': kod, 'sure_teyit': teyit}), 200


@app.route('/api/saha_cihazlari', methods=['GET'])
def saha_cihazlari():
    """Beklenen tüm saha cihazlarının durumu (kaynak/montaj/metal).

    Beklenen cihaz listesi sabit:
      - 9 robot:    ABB1-IO ... ABB9-IO       (bolum: kaynak)
      - 12 montaj:  MONTAJ-M1 ... MONTAJ-M12  (bolum: montaj)
      - 3 metal:    300T-IO, 400T-IO, 550T-IO (bolum: metal)

    Her cihaz için pilot.db'deki cihaz_kayitlari ile eşleştirilir.
    Eşleşme yoksa 'beklemede' (hiç bağlanmamış).
    """
    import os
    pilot_db = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pilot', 'pilot.db')

    BEKLENEN = {
        'kaynak': [{'cihaz_id': f'ABB{i}-IO', 'robot_no': f'ABB{i}'} for i in range(1, 10)],
        'montaj': [{'cihaz_id': f'MONTAJ-M{i}', 'robot_no': f'M{i}'} for i in range(1, 13)],
        'metal':  [
            {'cihaz_id': '300T-IO', 'robot_no': '300T'},
            {'cihaz_id': '400T-IO', 'robot_no': '400T'},
            {'cihaz_id': '550T-IO', 'robot_no': '550T'},
        ],
    }

    # Pilot DB'den mevcut cihaz_kayitlari + bugünün ist1/ist2 sayıları
    mevcut = {}
    ist_sayimlari = {}   # {cihaz_id: {1: adet, 2: adet}}
    reset_map = {}       # {(cihaz_id, istasyon): reset_ts}
    if os.path.exists(pilot_db):
        import sqlite3
        c = sqlite3.connect(pilot_db)
        c.row_factory = sqlite3.Row
        try:
            rows = c.execute('''
                SELECT *,
                       CAST((julianday('now','localtime') - julianday(son_heartbeat)) * 1440 AS INTEGER) as son_heartbeat_dk
                FROM cihaz_kayitlari
            ''').fetchall()
            for r in rows:
                d = dict(r)
                d['durum'] = 'offline' if (d.get('son_heartbeat_dk') or 0) > 2 else 'online'
                mevcut[d['cihaz_id']] = d

            # Bugünün her cihaz + istasyon başına pulse adetleri (kaynak için ist1/ist2 ayrı göstermek için)
            for r in c.execute('''
                SELECT cihaz_id, istasyon, COUNT(*) as adet
                FROM sayac_olaylari
                WHERE date(ts) = date('now','localtime')
                GROUP BY cihaz_id, istasyon
            ''').fetchall():
                cid = r['cihaz_id']
                if cid not in ist_sayimlari:
                    ist_sayimlari[cid] = {}
                ist_sayimlari[cid][r['istasyon']] = r['adet']

            # Manuel reset noktalari — dashboard'dan "sayaci sifirla" tıklayınca yazıldı
            # Sayım sırasında filtre olarak kullanılır (ts > reset_ts olanları say)
            try:
                for r in c.execute('SELECT cihaz_id, istasyon, reset_ts FROM sayac_reset_noktalari').fetchall():
                    reset_map[(r['cihaz_id'], int(r['istasyon'] or 0))] = r['reset_ts']
            except Exception:
                pass  # Tablo henüz yoksa (pilot_app.py yeniden başlatılmamışsa) — atla
        except Exception as e:
            print(f'[saha_cihazlari] DB hata: {e}')
        finally:
            c.close()

    # ─── Aktif referans bilgisi (durum='uretimde') ──────────────────
    # Her robot+istasyon icin: hangi referans uretimde + ne zamandir + simdiye
    # kadar kac pulse var. Operator yeni referans secince uretime_baslama_ts
    # yenilenir, sayim sifirdan baslamis gibi gorunur.
    aktif_ref_map = {}  # {(robot_no, istasyon): {referans_kodu, basla_ts}}
    # Aktif vardiyalar — sayac sifirlama icin ek baslangic referansi
    aktif_vardiya_map = {}  # {(bolum, robot_no): vardiya_basla_ts ISO}
    try:
        conn_main = get_db()
        for r in conn_main.execute('''
            SELECT robot_no, istasyon, referans_kodu,
                   COALESCE(uretime_baslama_ts, olusturma_tarihi) as basla_ts
            FROM referans_takip
            WHERE durum='uretimde' AND robot_no <> ''
        ''').fetchall():
            key = (r['robot_no'], int(r['istasyon'] or 0))
            # Eger ayni robot+istasyon icin birden fazla varsa en sonuncuyu al
            mevcut_ref = aktif_ref_map.get(key)
            if (not mevcut_ref) or ((r['basla_ts'] or '') > (mevcut_ref['basla_ts'] or '')):
                aktif_ref_map[key] = {'referans_kodu': r['referans_kodu'], 'basla_ts': r['basla_ts']}

        # Bugun acik vardiyalari cek — operator vardiya acince sayim o andan
        # itibaren olur (yeni referans secilmemis bile olsa sifirdan baslar)
        bugun_str = datetime.now().strftime('%Y-%m-%d')
        for r in conn_main.execute('''
            SELECT robot_no, COALESCE(bolum,'kaynak') as bolum, baslangic_saati, tarih
            FROM vardiyalar
            WHERE durum='aktif' AND tarih=?
        ''', (bugun_str,)).fetchall():
            bs = (r['baslangic_saati'] or '').strip()
            if bs and r['robot_no']:
                # baslangic_saati 'HH:MM' formatinda — tarih ile birlestir ISO yap
                try:
                    ts_iso = f"{r['tarih']} {bs}:00" if len(bs) == 5 else f"{r['tarih']} {bs}"
                    # Ayni (bolum, robot_no) icin birden fazla aktif vardiya varsa
                    # en sonra acilmis olani al (en buyuk ts)
                    key = (r['bolum'], r['robot_no'])
                    if (key not in aktif_vardiya_map) or (ts_iso > aktif_vardiya_map[key]):
                        aktif_vardiya_map[key] = ts_iso
                except Exception:
                    pass
        conn_main.close()
    except Exception as e:
        print(f'[saha_cihazlari] referans_takip/vardiyalar hata: {e}')

    # ─── Pulse sayımı (pilot DB'den, baslangic_ts sonrası) ───
    def _aktif_pulse_say(cihaz_id, istasyon, basla_ts):
        """Verilen cihaz/istasyon icin basla_ts'den sonraki pulse sayisi."""
        if not basla_ts or not os.path.exists(pilot_db):
            return 0
        try:
            import sqlite3
            pc = sqlite3.connect(pilot_db)
            try:
                q = '''SELECT COUNT(*) FROM sayac_olaylari
                       WHERE cihaz_id=? AND ts >= ?'''
                params = [cihaz_id, basla_ts]
                if istasyon and istasyon > 0:
                    q += ' AND istasyon=?'
                    params.append(istasyon)
                row = pc.execute(q, params).fetchone()
                return row[0] if row else 0
            finally:
                pc.close()
        except Exception:
            return 0

    def _en_son_ts(*candidates):
        """Verilen ISO timestamp string'lerinden en buyugunu (en son) dondur.
        None/bos olanlar atlanir. Hicbiri yoksa None."""
        gecerli = [t for t in candidates if t]
        return max(gecerli) if gecerli else None

    sonuc = {}
    for bolum, cihazlar in BEKLENEN.items():
        liste = []
        for beklenen_cihaz in cihazlar:
            cid = beklenen_cihaz['cihaz_id']
            rno = beklenen_cihaz['robot_no']
            kayit = mevcut.get(cid)

            # Aktif vardiya baslangic ts (varsa) — sayac sifirlama referansi
            vardiya_ts = aktif_vardiya_map.get((bolum, rno))
            # Manuel reset ts'leri (varsa) — dashboard'dan butonla sifirlanmis
            reset_ist1 = reset_map.get((cid, 1))
            reset_ist2 = reset_map.get((cid, 2))
            reset_all  = reset_map.get((cid, 0))  # tum istasyonlari kapsar

            # Aktif referans tespiti — kaynakta her istasyon ayri olabilir,
            # montaj/metal'de tek istasyon (genelde 0 veya 1).
            # Sayim icin filtre = max(vardiya_ts, referans_ts, manuel_reset_ts):
            # hangisi daha sonraysa o andan itibaren sayim baslar.
            aktif_referanslar = []
            if bolum == 'kaynak':
                for ist in (1, 2):
                    a = aktif_ref_map.get((rno, ist))
                    if a:
                        reset_ist = reset_ist1 if ist == 1 else reset_ist2
                        filt_ts = _en_son_ts(vardiya_ts, a['basla_ts'], reset_ist, reset_all)
                        aktif_referanslar.append({
                            'istasyon':      ist,
                            'referans_kodu': a['referans_kodu'],
                            'basla_ts':      filt_ts,
                            'pulse_sayisi':  _aktif_pulse_say(cid, ist, filt_ts),
                        })
            else:
                # Montaj/metal: istasyon belirtilmemis (0) veya 1 — her ikisini de yakala
                a = None
                for ist in (0, 1):
                    a = aktif_ref_map.get((rno, ist))
                    if a:
                        break
                # Montaj fallback: operator hat secmez, robot_no='MONTAJ' default (eski kayitlar)
                if not a and bolum == 'montaj':
                    for fallback_rno in ('MONTAJ', ''):
                        for ist in (0, 1):
                            a = aktif_ref_map.get((fallback_rno, ist))
                            if a:
                                break
                        if a:
                            break
                if a:
                    filt_ts = _en_son_ts(vardiya_ts, a['basla_ts'], reset_all, reset_ist1)
                    aktif_referanslar.append({
                        'istasyon':      None,
                        'referans_kodu': a['referans_kodu'],
                        'basla_ts':      filt_ts,
                        'pulse_sayisi':  _aktif_pulse_say(cid, 0, filt_ts),
                    })

            # Sayac degerlerini hesapla:
            # - Aktif vardiya varsa veya manuel reset yapilmissa: o andan sonraki pulse'lar
            # - Hicbiri yoksa: bugunun tamamindaki pulse'lar (eski davranis)
            if kayit:
                ist_map = ist_sayimlari.get(cid, {})
                # Istasyon basina filtre ts: vardiya, kendi reset, all reset
                ist1_filt = _en_son_ts(vardiya_ts, reset_ist1, reset_all)
                ist2_filt = _en_son_ts(vardiya_ts, reset_ist2, reset_all)
                if ist1_filt or ist2_filt:
                    ist1_v = _aktif_pulse_say(cid, 1, ist1_filt) if ist1_filt else ist_map.get(1, 0)
                    ist2_v = _aktif_pulse_say(cid, 2, ist2_filt) if ist2_filt else ist_map.get(2, 0)
                    # ist=0 (montaj/metal generic) icin de all-reset uygula
                    toplam_ts = _en_son_ts(vardiya_ts, reset_all)
                    toplam_0 = _aktif_pulse_say(cid, 0, toplam_ts) if toplam_ts else 0
                    toplam_v = ist1_v + ist2_v + toplam_0
                else:
                    ist1_v = ist_map.get(1, 0)
                    ist2_v = ist_map.get(2, 0)
                    toplam_v = sum(ist_map.values())

                liste.append({
                    'cihaz_id':           cid,
                    'robot_no':           rno,
                    'bolum':              bolum,
                    'durum':              kayit.get('durum', 'offline'),
                    'ip_adresi':          kayit.get('ip_adresi', ''),
                    'wifi_rssi':          kayit.get('wifi_rssi', 0),
                    'son_heartbeat':      kayit.get('son_heartbeat', ''),
                    'son_heartbeat_dk':   kayit.get('son_heartbeat_dk', 0),
                    'firmware_ver':       kayit.get('firmware_ver', ''),
                    'toplam_sinyal':      kayit.get('toplam_sinyal', 0),
                    'buffer_kuyruk':      kayit.get('buffer_kuyruk', 0),
                    'uptime_sn':          kayit.get('uptime_sn', 0),
                    'free_heap':          kayit.get('free_heap', 0),
                    'bugun_ist1':         ist1_v,
                    'bugun_ist2':         ist2_v,
                    'bugun_toplam':       toplam_v,
                    'robot_calisiyor':    1 if kayit.get('robot_calisiyor') else 0,
                    'aktif_vardiya_ts':   vardiya_ts,
                    'aktif_referanslar':  aktif_referanslar,
                    'kayitli':            True,
                })
            else:
                # Hiç bağlanmamış
                liste.append({
                    'cihaz_id':           cid,
                    'robot_no':           rno,
                    'bolum':              bolum,
                    'durum':              'beklemede',
                    'aktif_vardiya_ts':   vardiya_ts,
                    'aktif_referanslar':  aktif_referanslar,
                    'kayitli':            False,
                })
        sonuc[bolum] = liste

    return jsonify(sonuc), 200


@app.route('/api/saha_cihazlari/sayac_reset', methods=['POST'])
def saha_cihazlari_sayac_reset():
    """Manuel sayaç sıfırlama — dashboard kartından çağrılır.

    Body:
      - cihaz_id: 'ABB2-IO', 'MONTAJ-M3', '400T-IO' vb. (zorunlu)
      - istasyon: 0 (tüm), 1 (sadece ist.1), 2 (sadece ist.2) — default 0
      - yapan:    opsiyonel, logging için

    Pilot.db.sayac_reset_noktalari'na INSERT OR REPLACE yapılır.
    /api/saha_cihazlari endpoint'i sayım sırasında bu ts'i de en sona dahil eder.
    """
    import os, sqlite3
    data = request.get_json() or {}
    cihaz_id = (data.get('cihaz_id') or '').strip()
    istasyon = int(data.get('istasyon') or 0)
    yapan    = (data.get('yapan') or '').strip()

    if not cihaz_id:
        return jsonify({'hata': 'cihaz_id zorunlu'}), 400
    if istasyon not in (0, 1, 2):
        return jsonify({'hata': 'istasyon 0/1/2 olmalı'}), 400

    pilot_db = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pilot', 'pilot.db')
    if not os.path.exists(pilot_db):
        return jsonify({'hata': 'Pilot DB bulunamadı (pilot_app.py çalışıyor mu?)'}), 503

    try:
        pc = sqlite3.connect(pilot_db)
        try:
            # Tablo henüz yoksa (pilot_app.py yeniden başlatılmamışsa) oluştur
            pc.execute('''
                CREATE TABLE IF NOT EXISTS sayac_reset_noktalari (
                    cihaz_id   TEXT NOT NULL,
                    istasyon   INTEGER NOT NULL DEFAULT 0,
                    reset_ts   TEXT NOT NULL,
                    yapan      TEXT DEFAULT '',
                    PRIMARY KEY (cihaz_id, istasyon)
                )
            ''')
            cur = pc.execute('''
                INSERT OR REPLACE INTO sayac_reset_noktalari
                (cihaz_id, istasyon, reset_ts, yapan)
                VALUES (?, ?, datetime('now','localtime'), ?)
            ''', (cihaz_id, istasyon, yapan))
            pc.commit()
            # Yazılan ts'yi geri oku
            row = pc.execute(
                'SELECT reset_ts FROM sayac_reset_noktalari WHERE cihaz_id=? AND istasyon=?',
                (cihaz_id, istasyon)
            ).fetchone()
            reset_ts = row[0] if row else None
        finally:
            pc.close()
    except Exception as e:
        return jsonify({'hata': f'Reset yazılamadı: {str(e)}'}), 500

    ist_label = 'tüm istasyonlar' if istasyon == 0 else f'İstasyon {istasyon}'
    print(f"[sayac_reset] {cihaz_id} / {ist_label} → reset_ts={reset_ts} (yapan: {yapan or 'belirtilmemiş'})")
    return jsonify({
        'basarili': True,
        'cihaz_id': cihaz_id,
        'istasyon': istasyon,
        'reset_ts': reset_ts,
    }), 200


@app.route('/api/pilot/sinyal_analiz', methods=['GET'])
def pilot_sinyal_analiz():
    """Pilot sayaç pulse'larının zaman analizi.

    Query parametreleri:
      - bolum:    'kaynak' | 'montaj' | 'metal' (zorunlu)
      - robot_no: 'ABB2', 'M3', '400T' vb. (zorunlu)
      - istasyon: '1' | '2' | '' (boş = tüm istasyonlar)
      - tarih:    'YYYY-MM-DD' (default bugün)

    Dönen:
      - olaylar:        her pulse {ts, istasyon, gap_sn}
      - ozet:           toplam_pulse, ort_gap_sn, en_uzun_gap, en_kisa_gap
      - saatlik:        [{hour, pulse_sayisi, ort_gap_sn}]
      - duruslar:       [{basla, biti, sebep, tip, sure_dk}]
      - referans_donemler: o gün üretilen referansların created_at timeline'ı
      - aktif_referanslar: şu an referans_takip'te uretimde olanlar
    """
    import os, sqlite3
    from datetime import datetime, timedelta

    bolum    = request.args.get('bolum', 'kaynak')
    robot_no = request.args.get('robot_no', '').strip()
    istasyon = request.args.get('istasyon', '').strip()
    tarih    = request.args.get('tarih', datetime.now().strftime('%Y-%m-%d'))

    if not robot_no:
        return jsonify({'hata': 'robot_no parametresi zorunlu'}), 400

    # ─── Pilot DB'den pulse'lar ───────────────────────────────────
    pilot_db = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pilot', 'pilot.db')
    olaylar = []
    if os.path.exists(pilot_db):
        pc = sqlite3.connect(pilot_db)
        pc.row_factory = sqlite3.Row
        try:
            q = '''SELECT ts, istasyon FROM sayac_olaylari
                   WHERE bolum=? AND robot_no=? AND date(ts)=?'''
            params = [bolum, robot_no, tarih]
            if istasyon:
                q += ' AND istasyon=?'
                params.append(int(istasyon))
            q += ' ORDER BY ts'
            rows = pc.execute(q, params).fetchall()

            # Her istasyon için ayrı gap takibi (ist1 ile ist2 birbirine karışmasın)
            son_ts_per_ist = {}
            for r in rows:
                try:
                    ts_dt = datetime.fromisoformat(r['ts'])
                except Exception:
                    continue
                ist = r['istasyon'] or 0
                prev = son_ts_per_ist.get(ist)
                gap_sn = round((ts_dt - prev).total_seconds(), 1) if prev else None
                olaylar.append({
                    'ts':       r['ts'],
                    'istasyon': ist,
                    'gap_sn':   gap_sn,
                })
                son_ts_per_ist[ist] = ts_dt
        except Exception as e:
            print(f'[sinyal_analiz] Pilot DB hata: {e}')
        finally:
            pc.close()

    # ─── Özet metrikleri ──────────────────────────────────────────
    gaps = [o['gap_sn'] for o in olaylar if o['gap_sn'] is not None]
    # Medyan
    _medyan = 0
    if gaps:
        _sorted = sorted(gaps)
        _mid = len(_sorted) // 2
        _medyan = round(_sorted[_mid] if len(_sorted) % 2 else (_sorted[_mid-1] + _sorted[_mid]) / 2, 1)
    # Mod (1 sn yuvarlanmış değerlerin en sık görüleni)
    _mod = 0
    if gaps:
        _counts = {}
        for v in gaps:
            r = round(v)
            _counts[r] = _counts.get(r, 0) + 1
        _mod = max(_counts, key=_counts.get)
    ozet = {
        'toplam_pulse':    len(olaylar),
        'ortalama_gap_sn': round(sum(gaps) / len(gaps), 1) if gaps else 0,
        'medyan_gap_sn':   _medyan,
        'mod_gap_sn':      _mod,
        'en_uzun_gap_sn':  round(max(gaps), 1) if gaps else 0,
        'en_kisa_gap_sn':  round(min(gaps), 1) if gaps else 0,
    }

    # ─── Saatlik dağılım ─────────────────────────────────────────
    saatlik_map = {}  # {hour: [gap1, gap2, ...]}
    saatlik_sayim = {}  # {hour: pulse_count}
    for o in olaylar:
        try:
            h = datetime.fromisoformat(o['ts']).hour
        except Exception:
            continue
        saatlik_sayim[h] = saatlik_sayim.get(h, 0) + 1
        if o['gap_sn'] is not None:
            saatlik_map.setdefault(h, []).append(o['gap_sn'])
    saatlik = []
    for h in range(24):
        gaps_h = saatlik_map.get(h, [])
        if h in saatlik_sayim:
            saatlik.append({
                'hour':         h,
                'pulse_sayisi': saatlik_sayim[h],
                'ort_gap_sn':   round(sum(gaps_h) / len(gaps_h), 1) if gaps_h else 0,
            })

    # ─── Duruşlar (ana DB) ──────────────────────────────────────
    conn = get_db()
    durus_rows = conn.execute('''
        SELECT d.baslangic_saati, d.sure_dk, d.durus_sebebi, d.durus_tipi, d.aciklama
        FROM duruslar d
        JOIN vardiyalar v ON d.vardiya_id = v.id
        WHERE v.robot_no=? AND v.tarih=?
        ORDER BY d.baslangic_saati
    ''', (robot_no, tarih)).fetchall()
    duruslar = []
    for r in durus_rows:
        bs = (r['baslangic_saati'] or '').strip()
        sure_dk = r['sure_dk'] or 0
        if not bs or sure_dk <= 0:
            continue
        # baslangic_saati "HH:MM" formatında — tarih ile birleştir
        try:
            if 'T' in bs or len(bs) > 8:
                basla_iso = bs
            else:
                basla_iso = f'{tarih}T{bs}:00' if len(bs) == 5 else f'{tarih}T{bs}'
            basla_dt = datetime.fromisoformat(basla_iso)
            biti_dt = basla_dt + timedelta(minutes=sure_dk)
            duruslar.append({
                'basla':    basla_dt.isoformat(),
                'biti':     biti_dt.isoformat(),
                'sure_dk':  sure_dk,
                'sebep':    r['durus_sebebi'] or '',
                'tip':      r['durus_tipi'] or 'plansiz',
                'aciklama': r['aciklama'] or '',
            })
        except Exception:
            pass

    # ─── Referans dönemler (üretim kayıtları timeline'ı) ────────
    uretim_rows = conn.execute('''
        SELECT u.referans_kodu, u.ok_adet, u.nok_adet, u.cycle_time_sn, u.created_at,
               r.hedef_cycle_time_sn
        FROM uretim_kayitlari u
        JOIN vardiyalar v ON u.vardiya_id = v.id
        LEFT JOIN referans_listesi r ON r.referans_kodu = u.referans_kodu
                                     AND COALESCE(r.bolum,'kaynak') = COALESCE(v.bolum,'kaynak')
        WHERE v.robot_no=? AND v.tarih=?
        ORDER BY u.created_at
    ''', (robot_no, tarih)).fetchall() if _kolon_var(conn, 'vardiyalar', 'bolum') else conn.execute('''
        SELECT u.referans_kodu, u.ok_adet, u.nok_adet, u.cycle_time_sn, u.created_at,
               r.hedef_cycle_time_sn
        FROM uretim_kayitlari u
        JOIN vardiyalar v ON u.vardiya_id = v.id
        LEFT JOIN referans_listesi r ON r.referans_kodu = u.referans_kodu
        WHERE v.robot_no=? AND v.tarih=?
        ORDER BY u.created_at
    ''', (robot_no, tarih)).fetchall()
    referans_donemler = []
    for r in uretim_rows:
        referans_donemler.append({
            'ts':              r['created_at'],
            'referans_kodu':   r['referans_kodu'],
            'ok_adet':         r['ok_adet'] or 0,
            'nok_adet':        r['nok_adet'] or 0,
            'cycle_time_sn':   r['cycle_time_sn'] or 0,
            'hedef_cycle_sn':  r['hedef_cycle_time_sn'] or 0,
        })

    # ─── Aktif referanslar (şu an üretimde olanlar) ─────────────
    aktif_ref_rows = conn.execute('''
        SELECT referans_kodu, istasyon, hedef_adet
        FROM referans_takip
        WHERE durum='uretimde' AND robot_no=?
    ''', (robot_no,)).fetchall()
    aktif_referanslar = [dict(r) for r in aktif_ref_rows]

    # ─── Referans değişim noktaları (o gün üretime giren referansların başlangıç anları) ───
    # Sinyal Analizi grafiginde dikey cizgi olarak gosterilir
    # Montaj icin robot_no='MONTAJ' veya bos olabilir — fallback ile yakala
    robot_filtre = "(robot_no=? OR robot_no='MONTAJ' OR robot_no='')" if bolum == 'montaj' else "robot_no=?"
    degisim_rows = conn.execute(f'''
        SELECT referans_kodu, uretime_baslama_ts, istasyon
        FROM referans_takip
        WHERE uretime_baslama_ts IS NOT NULL
          AND date(uretime_baslama_ts)=?
          AND {robot_filtre}
        ORDER BY uretime_baslama_ts
    ''', (tarih, robot_no)).fetchall()
    referans_degisim_noktalari = []
    for r in degisim_rows:
        ist = int(r['istasyon'] or 0)
        # Istasyon filtresi varsa eslesmeyenleri at
        if istasyon and ist > 0 and ist != int(istasyon):
            continue
        referans_degisim_noktalari.append({
            'ts':            r['uretime_baslama_ts'],
            'referans_kodu': r['referans_kodu'],
            'istasyon':      ist if ist > 0 else None,
        })

    conn.close()

    return jsonify({
        'cihaz_id':                   f'{robot_no}-IO' if bolum != 'montaj' else f'MONTAJ-{robot_no}',
        'robot_no':                   robot_no,
        'bolum':                      bolum,
        'istasyon':                   int(istasyon) if istasyon else None,
        'tarih':                      tarih,
        'ozet':                       ozet,
        'olaylar':                    olaylar,
        'saatlik':                    saatlik,
        'duruslar':                   duruslar,
        'referans_donemler':          referans_donemler,
        'referans_degisim_noktalari': referans_degisim_noktalari,
        'aktif_referanslar':          aktif_referanslar,
    }), 200


def _kolon_var(conn, tablo, kolon):
    """Migration güvenliği — kolon var mı kontrol."""
    try:
        cols = [r[1] for r in conn.execute(f"PRAGMA table_info({tablo})").fetchall()]
        return kolon in cols
    except Exception:
        return False


@app.route('/api/referans/teyit_ozet', methods=['GET'])
def referans_teyit_ozet():
    """Bölüm bazında teyit özeti: toplam / teyitli / teyitsiz / süresiz sayıları."""
    bolum = request.args.get('bolum', 'kaynak')
    conn = get_db()
    rows = conn.execute('''
        SELECT
          COUNT(*) as toplam,
          SUM(CASE WHEN COALESCE(sure_teyit,0)=1 THEN 1 ELSE 0 END) as teyitli,
          SUM(CASE WHEN COALESCE(sure_teyit,0)=0 AND COALESCE(hedef_cycle_time_sn,0)>0 THEN 1 ELSE 0 END) as teyitsiz,
          SUM(CASE WHEN COALESCE(hedef_cycle_time_sn,0)=0 THEN 1 ELSE 0 END) as suresiz
        FROM referans_listesi
        WHERE COALESCE(bolum,'kaynak')=?
    ''', (bolum,)).fetchone()
    conn.close()
    d = dict(rows) if rows else {}
    return jsonify({
        'bolum': bolum,
        'toplam': d.get('toplam', 0) or 0,
        'teyitli': d.get('teyitli', 0) or 0,
        'teyitsiz': d.get('teyitsiz', 0) or 0,
        'suresiz': d.get('suresiz', 0) or 0,
    }), 200


@app.route('/api/pair_cycle', methods=['GET'])
def api_pair_cycle():
    """Robot kaynak — iki istasyona atanan referansların gerçek paralel çevrim
    süresini hesaplar.

    Query: ?ist1=10.130.X&ist2=10.130.Y (her ikisi de opsiyonel)

    Mantık: max(K1, S2) + max(K2, S1)
      K1=İst1 kaynak süresi, S1=İst1 söktak süresi (referans bazlı)
    """
    from oee import pair_cycle_hesapla
    ist1 = (request.args.get('ist1') or '').strip()
    ist2 = (request.args.get('ist2') or '').strip()
    sonuc = pair_cycle_hesapla(ist1 or None, ist2 or None)
    return jsonify(sonuc), 200


@app.route('/api/referans/sureler', methods=['PATCH'])
def referans_sureler_guncelle():
    """Bir referansın kaynak ve söktak sürelerini günceller.
    Body: {referans_kodu, kaynak_suresi_sn, soktak_suresi_sn}
    Excel'in 'Kaynak Referans' sayfasını da otomatik günceller.
    """
    data = request.get_json() or {}
    kod = (data.get('referans_kodu') or '').strip()
    if not kod:
        return jsonify({'hata': 'referans_kodu zorunlu'}), 400
    try:
        kaynak = float(data.get('kaynak_suresi_sn', 0) or 0)
        soktak = float(data.get('soktak_suresi_sn', 0) or 0)
    except (ValueError, TypeError):
        return jsonify({'hata': 'kaynak/soktak sayısal olmalı'}), 400
    toplam = round(kaynak + soktak, 2)

    conn = get_db()
    try:
        c = conn.cursor()
        c.execute(
            "UPDATE referans_listesi SET kaynak_suresi_sn=?, soktak_suresi_sn=?, hedef_cycle_time_sn=? "
            "WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ',''))",
            (kaynak, soktak, toplam, kod)
        )
        # Geriye dönük: bu kodla mevcut üretim kayıtlarındaki cycle_time'ı güncelle
        c.execute(
            "UPDATE uretim_kayitlari SET cycle_time_sn=? "
            "WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ',''))",
            (toplam, kod)
        )
        conn.commit()
    finally:
        conn.close()

    # Excel'e yaz (Kaynak Referans sayfası — kolon B=kaynak, C=söktak)
    try:
        from import_excel import EXCEL_YOL, BOLUM_SAYFA
        import openpyxl, os
        if os.path.exists(EXCEL_YOL):
            wb = openpyxl.load_workbook(EXCEL_YOL)
            sayfa_adi = BOLUM_SAYFA['kaynak']['ref']
            if sayfa_adi in wb.sheetnames:
                ws = wb[sayfa_adi]
                norm_target = kod.upper().replace(' ', '')
                for r in range(2, ws.max_row + 1):
                    cell = ws.cell(row=r, column=1).value
                    if cell and str(cell).upper().replace(' ', '') == norm_target:
                        ws.cell(row=r, column=2, value=kaynak)
                        ws.cell(row=r, column=3, value=soktak)
                        ws.cell(row=r, column=4, value=f'=B{r}+C{r}')
                        break
                wb.save(EXCEL_YOL)
    except Exception as e:
        print(f'[referans/sureler] Excel sync hatası: {e}')

    return jsonify({'basarili': True, 'kaynak': kaynak, 'soktak': soktak, 'toplam': toplam}), 200


@app.route('/api/veri/arsivle_simdi', methods=['POST'])
def veri_arsivle_simdi():
    """Manuel tetik — 18:00'lık otomatik arşivin elle çalıştırılması."""
    try:
        from scheduler import gunluk_arsiv_calistir
        gunluk_arsiv_calistir()
        return jsonify({'basarili': True}), 200
    except Exception as e:
        return jsonify({'hata': str(e), 'basarili': False}), 500


@app.route('/api/durus_sebepleri', methods=['GET'])
def durus_sebepleri_api():
    """Bölüme göre duruş sebeplerini Excel'den okuyup döner.

    ?bolum=kaynak → [{'sebep': 'Robot Prog. Çalışması', 'tip': 'planli'}, ...]

    Excel sayfaları: data/uretim_verileri.xlsx
      - Robotik Kaynak Duruş Listesi
      - Montaj Duruş Listesi
      - Metal Enjeksiyon Duruş Listesi

    Cache-Control: no-store → Excel'de yapılan değişiklikler anında yansır.
    """
    bolum = (request.args.get('bolum') or 'kaynak').strip()
    if bolum not in ('kaynak', 'montaj', 'metal'):
        return jsonify({'hata': f"Geçersiz bölüm: {bolum}"}), 400
    sebepler = durus_sebepleri_yukle(bolum)
    resp = jsonify({'bolum': bolum, 'sebepler': sebepler})
    resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    resp.headers['Pragma'] = 'no-cache'
    return resp, 200


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
    # Tüm bölümlerde öncelik ASC, NULL olanlar en sonda → oluşturma tarihine göre
    # Kaynak/Söktak süreleri de JOIN ile eklendi (pair_cycle hesabı için)
    if bolum:
        rows = conn.execute('''
            SELECT rt.*,
                   rl.hedef_cycle_time_sn,
                   rl.kaynak_suresi_sn,
                   rl.soktak_suresi_sn
            FROM referans_takip rt
            LEFT JOIN referans_listesi rl ON REPLACE(rt.referans_kodu, ' ', '') = REPLACE(rl.referans_kodu, ' ', '')
            WHERE COALESCE(rt.bolum, 'kaynak') = ?
            ORDER BY (rt.oncelik IS NULL), rt.oncelik ASC, rt.olusturma_tarihi DESC
        ''', (bolum,)).fetchall()
    else:
        rows = conn.execute('''
            SELECT rt.*,
                   rl.hedef_cycle_time_sn,
                   rl.kaynak_suresi_sn,
                   rl.soktak_suresi_sn
            FROM referans_takip rt
            LEFT JOIN referans_listesi rl ON REPLACE(rt.referans_kodu, ' ', '') = REPLACE(rl.referans_kodu, ' ', '')
            ORDER BY (rt.oncelik IS NULL), rt.oncelik ASC, rt.olusturma_tarihi DESC
        ''').fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

def _oncelik_clamp(c, bolum, yeni_oncelik, eski_oncelik=None, exclude_id=None):
    """Yeni öncelik değerini, o bölümdeki mevcut öncelikli kayıt sayısına göre clamp eder.

    INSERT (eski_oncelik=None): max = mevcut_sayi + 1 (yeni eklenecek)
    UPDATE (eski var, yeni var): max = mevcut_sayi (toplam değişmiyor)
    UPDATE (eski None, yeni var): max = mevcut_sayi + 1 (öncelik atıyor, +1 olur)

    None döndürürse clamp gerekmedi (zaten geçerli) veya yeni_oncelik None.
    """
    if yeni_oncelik is None:
        return None
    where_excl = ' AND id != ?' if exclude_id is not None else ''
    excl_args = (exclude_id,) if exclude_id is not None else ()
    row = c.execute(f'''
        SELECT COUNT(*) as cnt FROM referans_takip
        WHERE COALESCE(bolum, 'kaynak') = ? AND oncelik IS NOT NULL
        {where_excl}
    ''', (bolum,) + excl_args).fetchone()
    mevcut_sayi = row['cnt'] if row else 0
    # Bu kayıt da öncelikli olacak: max = mevcut_sayi + 1 (kendisi de dahil)
    max_izinli = mevcut_sayi + 1
    if yeni_oncelik > max_izinli:
        return max_izinli
    if yeni_oncelik < 1:
        return 1
    return yeni_oncelik


def _oncelik_kaydir(c, bolum, eski_oncelik, yeni_oncelik, exclude_id=None):
    """Bir bölüm kaydının öncelik geçişinde diğer satırları doğru yönde kaydırır.
    Tüm bölümler için çalışır (kaynak/montaj/metal).

    eski_oncelik: kaydın önceki değeri (None = öncesinde öncelik yoktu / yeni kayıt)
    yeni_oncelik: kaydın yeni değeri (None = öncelik kaldırılıyor)
    exclude_id:   bu id'li satıra dokunma (PATCH için kendisini hariç tut)
    """
    # Değişiklik yok
    if eski_oncelik == yeni_oncelik:
        return

    where_excl = ' AND id != ?' if exclude_id is not None else ''
    excl_args = (exclude_id,) if exclude_id is not None else ()

    # 1) Yeni öncelik atanıyor (önceden yoktu): >= yeni olanları +1 kaydır
    if eski_oncelik is None and yeni_oncelik is not None:
        c.execute(f'''
            UPDATE referans_takip
            SET oncelik = oncelik + 1
            WHERE COALESCE(bolum, 'kaynak') = ?
              AND oncelik IS NOT NULL AND oncelik >= ?
              {where_excl}
        ''', (bolum, yeni_oncelik) + excl_args)

    # 2) Öncelik kaldırılıyor: > eski olanları -1 yukarı çek (boşluğu kapat)
    elif eski_oncelik is not None and yeni_oncelik is None:
        c.execute(f'''
            UPDATE referans_takip
            SET oncelik = oncelik - 1
            WHERE COALESCE(bolum, 'kaynak') = ?
              AND oncelik IS NOT NULL AND oncelik > ?
              {where_excl}
        ''', (bolum, eski_oncelik) + excl_args)

    # 3) Yukarı taşıma (önem artıyor: eski > yeni, ör. 4 -> 2)
    elif yeni_oncelik < eski_oncelik:
        c.execute(f'''
            UPDATE referans_takip
            SET oncelik = oncelik + 1
            WHERE COALESCE(bolum, 'kaynak') = ?
              AND oncelik IS NOT NULL AND oncelik >= ? AND oncelik < ?
              {where_excl}
        ''', (bolum, yeni_oncelik, eski_oncelik) + excl_args)

    # 4) Aşağı taşıma (önem azalıyor: eski < yeni, ör. 1 -> 3)
    else:  # yeni_oncelik > eski_oncelik
        c.execute(f'''
            UPDATE referans_takip
            SET oncelik = oncelik - 1
            WHERE COALESCE(bolum, 'kaynak') = ?
              AND oncelik IS NOT NULL AND oncelik > ? AND oncelik <= ?
              {where_excl}
        ''', (bolum, eski_oncelik, yeni_oncelik) + excl_args)


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
        # Öncelik clamp: kullanıcı çok yüksek bir sayı girdiyse mevcut listedeki
        # ref sayısına göre maksimum izinliye düşür (örn. 3 ref var, 5 girilmişse → 4)
        oncelik = _oncelik_clamp(c, bolum, oncelik)
        # Yeni kayıt: eski_oncelik=None, yeni_oncelik girildiyse mevcut >=N olanları +1
        _oncelik_kaydir(c, bolum, None, oncelik)
        durum_init = data.get('durum', 'launch_alinacak')
        # Eger referans dogrudan 'uretimde' durumda ekleniyorsa, pilot sayac sifirlama
        # zamanini su an olarak isaretle (yeni referans secimi = sayac sifir)
        uretime_baslama_init = "datetime('now','localtime')" if durum_init == 'uretimde' else None
        if uretime_baslama_init:
            c.execute(f'''
                INSERT INTO referans_takip (referans_kodu, hedef_adet, aciklama, durum, olusturan, robot_no, istasyon, bolum, oncelik, uretime_baslama_ts)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, {uretime_baslama_init})
            ''', (ref, int(data.get('hedef_adet', 0)), data.get('aciklama', ''),
                  durum_init, data.get('olusturan', ''),
                  data.get('robot_no', ''), int(data.get('istasyon', 0)), bolum, oncelik))
        else:
            c.execute('''
                INSERT INTO referans_takip (referans_kodu, hedef_adet, aciklama, durum, olusturan, robot_no, istasyon, bolum, oncelik)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (ref, int(data.get('hedef_adet', 0)), data.get('aciklama', ''),
                  durum_init, data.get('olusturan', ''),
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

    # Öncelik değişiyorsa shift mantığı çalıştır (tüm bölümler için)
    if 'oncelik' in data:
        mevcut = c.execute("SELECT bolum, oncelik FROM referans_takip WHERE id=?", (id,)).fetchone()
        if mevcut:
            bolum_kayit = (mevcut['bolum'] or 'kaynak')
            eski_onc = mevcut['oncelik']  # None olabilir
            yeni_raw = data.get('oncelik')
            yeni_onc = None
            if yeni_raw not in (None, '', 0, '0'):
                try:
                    yeni_onc = int(yeni_raw)
                    if yeni_onc < 1: yeni_onc = None
                except (TypeError, ValueError):
                    yeni_onc = None
            # Clamp: girilen sayı listedeki ref sayısını aşıyorsa düşür
            yeni_onc = _oncelik_clamp(c, bolum_kayit, yeni_onc, eski_oncelik=eski_onc, exclude_id=id)
            _oncelik_kaydir(c, bolum_kayit, eski_onc, yeni_onc, exclude_id=id)
            data['oncelik'] = yeni_onc  # normalize edildi

    fields, vals = [], []
    if 'durum' in data:
        fields.append('durum=?'); vals.append(data['durum'])
        # Durum 'uretimde'ye geciliyorsa pilot sayac sifirlama referansi olarak
        # uretime_baslama_ts'yi simdiye set et. (Duraklatip tekrar baslatirsa
        # sayac yeniden sifirlanir — su anki uretim hizini gosterir.)
        if data['durum'] == 'uretimde':
            fields.append("uretime_baslama_ts=datetime('now','localtime')")
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
        # robotla_calisiyor — sütun yoksa eski DB; güvenli erişim
        try:
            rc = 1 if v['robotla_calisiyor'] else 0
        except (KeyError, IndexError):
            rc = 0
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
            'robotla_calisiyor': rc,
            'pair_cycle': None,  # kaynak için iki istasyondaki aktif referansların gerçek çevrimi
        }
        for u in uretim_rows:
            if u['vardiya_id'] == v['id']:
                row = {
                    'ref': u['referans_kodu'],
                    'launch': u['launch_adet'] or 0,
                    'tamamlandi': 1 if u['tamamlandi'] else 0,
                    'teyit': 0,        # aşağıda ref_durum_map'ten doldurulacak
                    'suresiz': 0,
                }
                ist = u['istasyon'] or 0
                if ist == 1:   item['istasyon_1'].append(row)
                elif ist == 2: item['istasyon_2'].append(row)
                else:          item['diger'].append(row)
        aktif_vardiyalar.append(item)

    # Referans durumu (teyit / süresiz) — andonda işaretlemek için
    # Tek SQL ile tüm bölüm için map oluştur, sonra her satıra ekle
    ref_durum_map = {}
    if bolum == 'kaynak':
        for r in c.execute(
            "SELECT referans_kodu, COALESCE(sure_teyit,0) as teyit, COALESCE(hedef_cycle_time_sn,0) as ct "
            "FROM referans_listesi WHERE COALESCE(bolum,'kaynak')='kaynak'"
        ).fetchall():
            norm = str(r['referans_kodu'] or '').upper().replace(' ', '')
            ref_durum_map[norm] = {
                'teyit': int(r['teyit'] or 0),
                'suresiz': 1 if (r['ct'] or 0) <= 0 else 0,
            }
        # Her aktif_vardiya'nın her ref satırına teyit/süresiz bayraklarını yaz
        for it in aktif_vardiyalar:
            for koleksiyon in ('istasyon_1', 'istasyon_2', 'diger'):
                for row in it[koleksiyon]:
                    norm = str(row.get('ref') or '').upper().replace(' ', '')
                    durum = ref_durum_map.get(norm)
                    if durum:
                        row['teyit'] = durum['teyit']
                        row['suresiz'] = durum['suresiz']

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

    # Pair cycle hesabı (sadece kaynak bölümü için).
    # Her vardiyada İst.1 ve İst.2'deki en son üretim referansını al, çevrim hesapla.
    if bolum == 'kaynak':
        from oee import pair_cycle_hesapla
        for it in aktif_vardiyalar:
            ist1_kod = it['istasyon_1'][-1]['ref'] if it['istasyon_1'] else None
            ist2_kod = it['istasyon_2'][-1]['ref'] if it['istasyon_2'] else None
            if ist1_kod or ist2_kod:
                try:
                    it['pair_cycle'] = pair_cycle_hesapla(ist1_kod, ist2_kod)
                except Exception as e:
                    print(f'[andon pair_cycle] hata: {e}')
                    it['pair_cycle'] = None

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
      - metal:  300T, 400T, 550T, Şerit Testere
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
            robot_listesi = ['300T', '400T', '550T', 'Şerit Testere']
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

# Modul yüklenirken DB migration'larını çalıştır (debug auto-reload sonrası da
# eksik kolonların eklendiğinden emin olmak için).
init_db()


if __name__ == '__main__':
    # Flask debug modundayken (reloader) çift çalışmayı önlemek için kontrol
    if os.environ.get('WERKZEUG_RUN_MAIN') != 'true':
        init_db()

    # Scheduler — sadece bir kez başlat (reloader child process'inde başlat,
    # debug modda da main process'te çift olmasın)
    if not os.environ.get('FLASK_DEBUG') or os.environ.get('WERKZEUG_RUN_MAIN') == 'true':
        try:
            from scheduler import start_scheduler
            start_scheduler()
        except Exception as _e:
            print(f'[SCHED] başlatılamadı: {_e}')

    print("\n" + "="*55)
    print("  COFLE MANAGE - URETIM TAKIP SISTEMI CALISIYOR")
    print("="*55)
    print("  Operator Formu : https://coflemanage.online")
    print("  Yonetici Panel : https://coflemanage.online/dashboard")
    print("  Andon Ekrani   : https://coflemanage.online/andon")
    print("="*55)
    print("  Sistem artik bu adresten yayinlanmaktadir.")
    print("  Otomatik arsiv : Her gun 18:00 (data/arsiv/)")
    print("="*55 + "\n")
    app.run(host='0.0.0.0', port=5000, debug=True)
