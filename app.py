from flask import Flask, request, jsonify, render_template, send_file, g, session, redirect
from flask_cors import CORS
from datetime import datetime, date, timedelta
from functools import wraps
from werkzeug.security import generate_password_hash, check_password_hash
import json
import os
import re
import secrets
import traceback # For debugging

from database import get_db as db_connect, init_db
from oee import hesapla_oee, hesapla_oee_ozet
from import_excel import import_data, durus_sebepleri_yukle, import_tum, export_referans_cycle_times, import_tk1
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

# ── Oturum (session) gizli anahtarı ──
# Panel girişi imzalı cookie ile tutulur. Anahtar RESTART'lar arasında SABİT
# olmalı (yoksa her restart tüm oturumları düşürür). Öncelik: COFLE_SECRET_KEY
# env; yoksa proje kökünde tek seferlik üretilip saklanan .flask_secret dosyası
# (gitignore'da). Dosya yazılamazsa (salt-okunur FS) süreç-ömürlü rastgele anahtar.
def _secret_key_yukle():
    env_key = os.environ.get('COFLE_SECRET_KEY')
    if env_key:
        return env_key
    yol = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.flask_secret')
    try:
        if os.path.exists(yol):
            with open(yol, 'r', encoding='utf-8') as f:
                k = f.read().strip()
                if k:
                    return k
        k = secrets.token_hex(32)
        with open(yol, 'w', encoding='utf-8') as f:
            f.write(k)
        return k
    except Exception:
        return secrets.token_hex(32)

app.secret_key = _secret_key_yukle()
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    PERMANENT_SESSION_LIFETIME=timedelta(hours=12),
)

# ── Panel yetki modeli ──
# İzin verilebilen sayfa anahtarları (dashboard_v2 sidebar data-sayfa değerleri).
# 'kullanicilar' burada YOK — o admin'e özeldir (rol ile korunur, izinle değil).
PANEL_SAYFALAR = [
    'ozet', 'bolum', 'kayitlar', 'is-yonetimi', 'fikstur', 'referanslar',
    'operatorler', 'saha-cihazlari', 'sinyal-analizi', 'andon-ayarlari', 'raporlar',
    'as400-teyit',
]

def panel_kullanici():
    """Aktif oturumdaki panel kullanıcısını DB'den taze çeker (izin/aktiflik anlık).
    Oturum yoksa veya kullanıcı pasif/silinmişse None."""
    kid = session.get('panel_uid')
    if not kid:
        return None
    try:
        conn = get_db()
        row = conn.execute(
            "SELECT id, kullanici_adi, ad_soyad, rol, izinler, aktif, sifre_gecici "
            "FROM panel_kullanicilari WHERE id=?", (kid,)
        ).fetchone()
    except Exception:
        return None
    if not row or not row['aktif']:
        return None
    try:
        izinler = json.loads(row['izinler'] or '[]')
        if not isinstance(izinler, list):
            izinler = []
    except Exception:
        izinler = []
    admin = (row['rol'] == 'admin')
    return {
        'id': row['id'],
        'kullanici_adi': row['kullanici_adi'],
        'ad_soyad': row['ad_soyad'] or '',
        'rol': row['rol'],
        'admin': admin,
        # admin tüm sayfaları görür; kullanıcı yalnız izin listesi
        'izinler': PANEL_SAYFALAR[:] if admin else [s for s in izinler if s in PANEL_SAYFALAR],
        'sifre_gecici': bool(row['sifre_gecici']),
    }

def panel_gerekli(izin=None, admin=False):
    """Endpoint koruma decorator'ı.
      panel_gerekli()                → sadece giriş yapılmış olmalı
      panel_gerekli(izin='referanslar') → o sayfa izni (admin daima geçer)
      panel_gerekli(admin=True)      → yalnız admin
    401 = giriş yok, 403 = yetki yok. Operatör mobil/andon uç noktalarına
    UYGULANMAZ (onlar oturumsuz çalışır)."""
    def dekorator(fn):
        @wraps(fn)
        def sarmal(*args, **kwargs):
            ku = panel_kullanici()
            if not ku:
                return jsonify({'hata': 'Oturum gerekli', 'giris_gerekli': True}), 401
            if admin and not ku['admin']:
                return jsonify({'hata': 'Bu işlem için yönetici yetkisi gerekli'}), 403
            if izin and not ku['admin'] and izin not in ku['izinler']:
                return jsonify({'hata': f'Bu sayfa için yetkiniz yok ({izin})'}), 403
            g.panel_ku = ku
            return fn(*args, **kwargs)
        return sarmal
    return dekorator

# Şablon hot-reload — debug'dan BAĞIMSIZ. debug=False (güvenlik) iken Flask normalde
# şablonları önbelleğe alır ve .html değişiklikleri ancak restart'ta yansır. Bunu açık
# tutarak HTML değişikliklerini RESTART OLMADAN canlıya alıyoruz (ufak mtime kontrolü,
# ihmal edilebilir maliyet) — güvenlik debug'ı kapalı kalır.
app.config['TEMPLATES_AUTO_RELOAD'] = True

@app.after_request
def _html_no_cache(resp):
    """HTML sayfaları her açılışta TAZE gelsin — PWA/ana ekran kısayolu bayat
    sürüm göstermesin (tarayıcı her seferinde sunucuyla doğrulasın)."""
    try:
        if resp.mimetype == 'text/html':
            resp.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
            resp.headers['Pragma'] = 'no-cache'
    except Exception:
        pass
    return resp

# GUVENLIK: wildcard CORS YOK. Tum erisim same-origin oldugu icin varsayilan: cross-origin kapali.
# Frontend ayri domaine bolunurse: COFLE_CORS_ORIGINS=https://alan.adi (virgulle coklu) ile ac.
_cors_origins = [o.strip() for o in os.environ.get('COFLE_CORS_ORIGINS', '').split(',') if o.strip()]
if _cors_origins:
    CORS(app, origins=_cors_origins)

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


@app.errorhandler(Exception)
def api_hata_json(e):
    """/api/* yollarında yakalanmamış istisna HTML 500 yerine JSON döner.
    Mobil/dashboard r.json() bekler — HTML gelince operatör anlamsız 'Bağlantı hatası'
    görür ve gerçek sebep gizlenir (2026-07-06 'closed database' arızası dersi).
    Sayfa route'ları (HTML) etkilenmez. HTTPException'lar (404/400 vb.) aynen geçer."""
    from werkzeug.exceptions import HTTPException
    if isinstance(e, HTTPException):
        # /api/* yollarında 400/404/405 vb. de JSON — mobil r.json() her zaman parse edebilsin
        if request.path.startswith('/api/'):
            return jsonify({'hata': e.description or e.name, 'kod': e.code}), e.code
        return e
    if request.path.startswith('/api/'):
        print(f"HATA ({request.method} {request.path}): {traceback.format_exc()}")
        return jsonify({'hata': f'Sunucu hatası: {type(e).__name__}: {e}'}), 500
    raise e

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
        g.operator_lokasyon = None
        if op and pin:
            # Lokasyon (varsa, fetch wrapper query'e ekler) ile daralt — composite UNIQUE
            # sonrası aynı ad iki tesiste olabilir. PIN'i tutan SATIRIN lokasyonu g'ye yazılır
            # (sahiplik kontrolü tesis karşılaştırması yapabilsin).
            lok = (request.args.get('lokasyon') or '').strip()
            # DIKKAT: get_db() flask.g'de REQUEST-PAYLASIMLI baglanti doner — BURADA KAPATMA!
            # (Decorator istegin BASINDA calisir; kapatirsak handler ayni kapali baglantiyi
            # alir → "Cannot operate on a closed database" 500. Kapatma teardown_db'de.)
            conn = get_db()
            if lok:
                rows = conn.execute("SELECT pin, COALESCE(lokasyon,'TK2') as lokasyon FROM operatorler WHERE ad=? AND COALESCE(lokasyon,'TK2')=?", (op, lok)).fetchall()
            else:
                rows = conn.execute("SELECT pin, COALESCE(lokasyon,'TK2') as lokasyon FROM operatorler WHERE ad=?", (op,)).fetchall()
            if not rows:
                return jsonify({'hata': f'Operatör bulunamadı: {op}'}), 403
            eslesen = next((r for r in rows if (r['pin'] or '0000') == pin), None)
            if eslesen is None:
                return jsonify({'hata': 'PIN yanlış'}), 403
            g.operator_adi = op
            g.operator_lokasyon = eslesen['lokasyon'] or 'TK2'
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
    if g.operator_adi == ADMIN_ADI:
        return True, None  # yönetici — tüm vardiyalara erişir
    conn = get_db()
    row = conn.execute("SELECT operator_adi, COALESCE(lokasyon,'TK2') as lokasyon FROM vardiyalar WHERE id=?", (vardiya_id,)).fetchone()
    if not row:
        return False, f'Vardiya bulunamadı (id={vardiya_id})'
    if row['operator_adi'] != g.operator_adi:
        return False, f'Bu vardiya başka operatöre ait: {row["operator_adi"]}'
    # Tesis karşılaştırması: composite UNIQUE sonrası iki tesiste aynı ad olabilir —
    # TK1'deki 'X', TK2'deki aynı adlı 'X'in vardiyasına dokunamasın.
    op_lok = getattr(g, 'operator_lokasyon', None)
    if op_lok and row['lokasyon'] != op_lok:
        return False, f'Bu vardiya başka tesise ait ({row["lokasyon"]})'
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
# SAYAÇ ENTEGRASYONU — pilot sensör pulse'larını üretim kaydına bağlar
# ─────────────────────────────────────────────────────────────
# Sadece sahada kurulu + doğrulanmış cihazlarda OK adet sensörden otomatik beslenir.
# Yeni cihaz devreye alınıp test edildikçe robot_no buraya eklenir.
# Kaynak ABB1..ABB9: tüm robotlar devreye alındı (2026-06). NOT: 2 robotun fiziksel sayaç
#   montajı 2026-06-26'da tamamlanacak — o robotlarda cihaz sayana kadar operatör ADET'i ELLE
#   girmeli (auto mod sensörü 0 okur → boş bırakılırsa üretim 0 görünür).
# Montaj M1..M12: tüm hatlara sayaç modülü kuruldu (2026-06) — hepsi sensör destekli.
# Metal: 300T + 400T + 550T. 300T [PULSE] testinden geçti, sahada sayıyor (2026-07-16).
# NOT: andon "makine çalışıyor" + sayaç rozeti bu allowlist'ten BAĞIMSIZ (saha_cihazlari
# robot_calisiyor'u sinyalden türetir); allowlist SADECE operatör mobil otomatik sayacını açar.
SAYAC_AUTO_CIHAZLAR = (
    {f'ABB{i}' for i in range(1, 10)} | {'300T', '400T', '550T'} | {f'M{i}' for i in range(1, 13)}
    | {'YF1', 'LF-LFP'}
    # 2026-07-23 yeni sayaç modülleri: abkant (röle) + eksantrik pres (iki-el AND) →
    # bölüm 'pres'; plastik enjeksiyon (320T/407T, metal mantığı) → bölüm 'plastik' (TK1).
    # 2026-07-27 saha rollout: yapıştırma (kısa röle pulse, plastik hattının makinesi)
    # → bölüm 'plastik' (TK1). robot_no ASCII 'Yapistirma' (firmware ile birebir).
    | {'Abkant 1', 'Abkant 2', 'Abkant 3'} | {f'Pres {i}' for i in range(1, 6)}
    | {'320T', '407T', 'Yapistirma'}
)
# YF1: TK1 (yan tesis) montaj deneme modülü — MONTAJ-YF1 cihazı robot_no='YF1' ile flash'lı
# (sahada zaten yüklü, yeniden flash YOK). pilot.db'de bolum=montaj/robot_no=YF1; saha_cihazlari
# BEKLENEN'inde lokasyon=TK1 olarak izlenir. robot_no global benzersiz → firmware'de lokasyon gerekmez.

# TK1 (yan tesis) cihazlarının robot_no'ları — pilot.db'de lokasyon kolonu olmadığı için
# saha_cihazlari/sinyal_kalitesi'ni lokasyona göre filtrelerken robot_no ile eşleriz.
# YF1 = sahadaki deneme modülü; diğerleri TK1 hatları (ileride cihaz takılırsa kapsanır).
TK1_ROBOT_NOLARI = {'YF1', 'Pull', 'Push-Pull', 'Iveco', 'LF-LFP',
                    '320T', '407T', 'Yapistirma'}   # 2026-07-27 plastik enj + yapıştırma (TK1)

# Hat → sayaç cihazının robot_no eşlemesi: operatör hattı seçer (vardiya.robot_no='LF-LFP')
# ama saha modülü pulse'ları robot_no='YF1' ile pilot.db'ye düşer. Sayım yapılırken hat
# adı cihazın robot_no'suna çevrilir. (Yeniden flash gerektirmeden eşleştirme.)
HAT_SAYAC_CIHAZI = {'LF-LFP': 'YF1'}

# ── PRES (Pres Abkant) MAKİNE MODELİ (kullanıcı 2026-07-28) ──
# Operatör gün içinde makine değiştirerek çalışıyor → makine VARDİYADA DEĞİL, her
# ÜRETİM KAYDINDA seçilir (lazer modelinin aynısı). Vardiya hattı tek ve sabittir:
# PRES_SABIT_HAT; makine, uretim_kayitlari.istasyon kolonunda 1..8 olarak durur.
# Sıra ÖNEMLİ: istasyon no = bu listedeki sıra (1=Abkant 1 ... 8=Pres 5); makine
# adları pilot.db cihaz robot_no'larıyla (firmware) BİREBİR aynı olmalı.
PRES_SABIT_HAT  = 'Pres Abkant'
PRES_MAKINELERI = ['Abkant 1', 'Abkant 2', 'Abkant 3',
                   'Pres 1', 'Pres 2', 'Pres 3', 'Pres 4', 'Pres 5']


def pres_makine_ad(istasyon):
    """istasyon no (1..8) → makine adı. Geçersiz/0 ise '' (makine seçilmemiş)."""
    try:
        i = int(istasyon or 0)
    except (TypeError, ValueError):
        return ''
    return PRES_MAKINELERI[i - 1] if 1 <= i <= len(PRES_MAKINELERI) else ''


def pres_ist_no(makine_ad):
    """Makine adı → istasyon no (1..8). Listede yoksa 0."""
    try:
        return PRES_MAKINELERI.index(str(makine_ad or '').strip()) + 1
    except ValueError:
        return 0


def _sayac_cihazi(bolum, robot_no, istasyon=0):
    """Vardiya hattı + üretim kaydı istasyonundan pilot.db CIHAZ robot_no'sunu bulur.
    - pres: hat sabit, MAKİNE istasyondan gelir (1=Abkant 1 ... 8=Pres 5).
      istasyon 0 ise (eski kayıtlar: makine vardiyada seçiliyordu) hat adı aynen
      kullanılır → geriye dönük uyum korunur.
    - diğer bölümler: hat→cihaz eşlemesi (LF-LFP→YF1), yoksa hattın kendisi."""
    if bolum == 'pres':
        ad = pres_makine_ad(istasyon)
        if ad:
            return ad
    return HAT_SAYAC_CIHAZI.get(robot_no, robot_no)


def _sayac_destekli(bolum, robot_no, istasyon=0):
    """KAYIT düzeyi: bu üretim kaydı için otomatik sensör sayımı var mı (allowlist)?
    KATI — pres'te makine (istasyon) çözülemiyorsa False döner. Aksi halde makinesiz
    kayıt 'auto' işaretlenir, sayım hep 0 okur ve gerçek ok_adet'i EZERDİ."""
    return _sayac_cihazi(bolum, robot_no, istasyon) in SAYAC_AUTO_CIHAZLAR


def _sayac_destekli_bolum(bolum, robot_no):
    """VARDİYA düzeyi UI ipucu: bu vardiyada sensör sayımı mümkün mü? pres'te makine
    üretim kaydında seçildiğinden bölümdeki herhangi bir sayaçlı makine yeterlidir
    (mobil sensör rozeti/sıfırla butonu vardiya açılışında da görünsün)."""
    if bolum == 'pres' and robot_no == PRES_SABIT_HAT:
        return any(m in SAYAC_AUTO_CIHAZLAR for m in PRES_MAKINELERI)
    return _sayac_destekli(bolum, robot_no, 0)

# Yönetici kullanıcı — operatorler tablosunda 'Admin' (PIN varsayılan 9999, panelden
# değiştirilebilir). Bu adla PIN girişi yapan TÜM vardiyalara erişir (sahiplik kontrolü
# bypass, mobilde kart kilidi yok) ve başka operatör adına vardiya açabilir.
ADMIN_ADI = 'Admin'

# Geçerli bölüm anahtarları — TEK kaynak. isleme/lazer/pres (2026-07): TK2'ye eklenen
# yeni bölümler; sayaç/andon/iş takibi YOK, sadece vardiya + üretim girişi + rapor.
# Yeni bölüm eklerken: import_excel.BOLUM_SAYFA + mobil/dashboard/rapor template
# toggle'ları da güncellenmeli (anahtarlar ASCII slug — URL/CSS class/JS map'lerde kullanılır).
GECERLI_BOLUMLER = ('kaynak', 'montaj', 'metal', 'isleme', 'lazer', 'pres', 'plastik')
BOLUM_AD = {
    'kaynak': 'Robot Kaynak',
    'montaj': 'Montaj',
    'metal':  'Metal Enjeksiyon',
    'isleme': 'İşleme',
    'lazer':  'Lazer Kesim',
    'pres':   'Pres Abkant',
    'plastik': 'Plastik Enjeksiyon',
}

# Push bildirim alıcıları — bölüme yeni referans/launch eklendiğinde kimin
# telefonuna bildirim gitsin. İsimler operatorler tablosundaki adlarla
# (harf/Türkçe karakter duyarsız) eşleştirilir. Yeni alıcı: buraya ekle.
BILDIRIM_ALICILARI = {
    'kaynak': ['Cihan Zilci', 'Mevlüt Bora'],
    'montaj': ['Mustafa Kuş'],
}

# Metal "makine calisiyor" durumu — her sinyal aslinda makine-calisiyor rolesi
# darbesidir (400T firmware o ~22sn aktif roleyi izler). Yesil SADECE role aktifken
# yansin diye esik role-aktif penceresine yakin tutulur: sinyal gelince yesil, ~20sn
# sonra (role birakinca / sinyal kesilince) soner. Boylece andon makinenin nabzini
# gosterir. NOT: andon bu durumu 3sn'de bir hizli poll eder (ana 30sn refresh'i
# beklemeden) — yoksa 20sn'lik pencere 30sn poll arasinda kacardi.
METAL_CALISIYOR_ESIK_SN = 20


def _pilot_db_yolu():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pilot', 'pilot.db')


def _pilot_pulse_say(bolum, robot_no, istasyon, basla_ts, biti_ts=None):
    """Pilot DB'de (bolum, robot_no) için basla_ts ile biti_ts arasındaki pulse sayısı.
    istasyon > 0 ise o istasyona filtrele; 0/None ise tüm istasyonlar (montaj/metal
    tek istasyon, firmware istasyon=1 gönderir → filtre yok hepsini sayar)."""
    # Hat → cihaz robot_no eşlemesi (LF-LFP→YF1; pres'te istasyon→makine)
    cihaz = _sayac_cihazi(bolum, robot_no, istasyon)
    # PRES: istasyon MAKİNE seçimidir, cihaz-içi istasyon değil — makineye çevrildiyse
    # istasyon filtresi KALKAR (abkant/pres firmware'i montaj gibi istasyon=1 gönderir;
    # filtre bırakılırsa istasyon=3 aranır ve hiç pulse bulunamazdı).
    if bolum == 'pres' and pres_makine_ad(istasyon):
        istasyon = 0
    robot_no = cihaz
    pilot_db = _pilot_db_yolu()
    if not basla_ts:
        return 0
    if not os.path.exists(pilot_db):
        return None   # pilot.db yok = OKUNAMADI (hata); gercek 0 ile karistirma
    import sqlite3
    try:
        pc = sqlite3.connect(pilot_db, timeout=5.0)
        try:
            pc.execute('PRAGMA busy_timeout=5000')   # kilitliyse 5sn bekle, hemen hata atma
            q = 'SELECT COUNT(*) FROM sayac_olaylari WHERE bolum=? AND robot_no=? AND ts >= ?'
            params = [bolum, robot_no, basla_ts]
            if biti_ts:
                q += ' AND ts < ?'
                params.append(biti_ts)
            if istasyon and istasyon > 0:
                q += ' AND istasyon=?'
                params.append(istasyon)
            row = pc.execute(q, params).fetchone()
            return row[0] if row else 0
        finally:
            pc.close()
    except Exception:
        return None   # KILIT/HATA: 0 DONDURME — caller ok_adet'i sifirlamasin


def _pilot_son_pulse_dk(bolum, robot_no, gun=None, istasyon=0):
    """Bu cihazdan en son pulse'tan bu yana geçen dakika. Hiç pulse yoksa None.
    10 dk duruş uyarısı için kullanılır. pres'te istasyon = operatörün o an
    çalıştığı makine (üretim kaydından gelir)."""
    robot_no = _sayac_cihazi(bolum, robot_no, istasyon)   # LF-LFP→YF1; pres: istasyon→makine
    pilot_db = _pilot_db_yolu()
    if not os.path.exists(pilot_db):
        return None
    import sqlite3
    from datetime import datetime as _dt
    gun = gun or date.today().isoformat()
    try:
        pc = sqlite3.connect(pilot_db)
        try:
            row = pc.execute(
                "SELECT MAX(ts) FROM sayac_olaylari WHERE bolum=? AND robot_no=? AND date(ts)=?",
                (bolum, robot_no, gun)
            ).fetchone()
            if not row or not row[0]:
                return None
            try:
                son = _dt.fromisoformat(row[0])
            except Exception:
                return None
            return max(0, int((_dt.now() - son).total_seconds() // 60))
        finally:
            pc.close()
    except Exception:
        return None


def _test_basari_say(conn, cihaz_id, referans_kodu, basla_ts, biti_ts=None):
    """Cofle test cihazından (cofle_test poller'ı doldurur) basla_ts ile biti_ts
    arasında o referansa ait BAŞARILI test sayısı. ESP32 pulse sayımının test-cihazı
    karşılığı: 1 başarılı test = 1 pulse. Kaynak: test_sonuclari (ana DB).
    basla_ts/biti_ts yerel '%Y-%m-%d %H:%M:%S' string; epoch'a çevrilir. Hata → None."""
    if not cihaz_id or not basla_ts:
        return None
    try:
        bas_epoch = int(datetime.strptime(basla_ts, '%Y-%m-%d %H:%M:%S').timestamp())
    except Exception:
        return None
    norm = (referans_kodu or '').strip().upper().replace(' ', '')
    q = ("SELECT COUNT(*) FROM test_sonuclari "
         "WHERE cihaz_id=? AND basarili=1 AND ts_epoch>=? AND (?='' OR referans_norm=?)")
    params = [cihaz_id, bas_epoch, norm, norm]
    if biti_ts:
        try:
            params.append(int(datetime.strptime(biti_ts, '%Y-%m-%d %H:%M:%S').timestamp()))
            q += " AND ts_epoch < ?"
        except Exception:
            pass
    try:
        row = conn.execute(q, params).fetchone()
        return row[0] if row else 0
    except Exception:
        return None   # tablo yok/hata: ok_adet'i SIFIRLAMA


def _test_son_aktivite_dk(conn, vardiya_id):
    """Vardiyanın AKTİF test-cihazı kayıtları (sayac_otomatik=1 + test_cihaz_id) için
    son BAŞARILI testten beri geçen dakika. Pilot pulse'ın test-cihazı karşılığı:
    operatör test cihazıyla çalışırken ESP32'ye pulse düşmez — '10dk sinyal yok'
    duruş uyarısı test aktivitesine bakmazsa YANLIŞ duruş sayar.
    Henüz hiç test yoksa referans başlangıcından beri sayar (o da gerçek duruştur).
    Test-cihazı kaydı yoksa None (pilot mantığı tek başına geçerli)."""
    try:
        rows = conn.execute(
            "SELECT test_cihaz_id, referans_kodu, sayac_baslangic_ts FROM uretim_kayitlari "
            "WHERE vardiya_id=? AND sayac_otomatik=1 AND test_cihaz_id IS NOT NULL "
            "AND sayac_baslangic_ts IS NOT NULL",
            (vardiya_id,)
        ).fetchall()
        if not rows:
            return None
        en_yeni_epoch = None
        for r in rows:
            try:
                bas_epoch = int(datetime.strptime(r['sayac_baslangic_ts'], '%Y-%m-%d %H:%M:%S').timestamp())
            except Exception:
                continue
            norm = (r['referans_kodu'] or '').strip().upper().replace(' ', '')
            row = conn.execute(
                "SELECT MAX(ts_epoch) FROM test_sonuclari "
                "WHERE cihaz_id=? AND basarili=1 AND ts_epoch>=? AND (?='' OR referans_norm=?)",
                (r['test_cihaz_id'], bas_epoch, norm, norm)
            ).fetchone()
            aktivite = (row[0] if row and row[0] else bas_epoch)  # hiç test yok → başlangıç anı
            if en_yeni_epoch is None or aktivite > en_yeni_epoch:
                en_yeni_epoch = aktivite
        if en_yeni_epoch is None:
            return None
        return max(0, int((datetime.now().timestamp() - en_yeni_epoch) // 60))
    except Exception:
        return None


def _bukum_bol(pulse, bukum_op):
    """Ham pulse sayısını ÜRETİLEN PARÇAYA çevirir (Pres Abkant büküm modeli).

    Bükümde bir parçaya sırayla tüm büküm operasyonları uygulanır, sonra sıradaki
    parçaya geçilir. 3 operasyonlu bir referansta abkant 3 kez sinyal üretir ama
    ortaya 1 parça çıkar → üretim teyidi 3. sinyaldir.

    TABAN BÖLME bilinçli: 7 pulse / 3 op = 2 parça (7. pulse üçüncü parçanın ilk
    bükümü, parça henüz BİTMEDİ). Kalan, kayıt açık kaldığı sürece bir sonraki
    senkronda tamamlanır — ok_adet her seferinde ham pulse'tan yeniden hesaplandığı
    için kayıp olmaz.

    pulse None (pilot.db okunamadı) ise None döner: çağıran ok_adet'i EZMEZ.
    bukum_op ≤ 1 ise dokunmaz — tüm eski kayıtlar ve diğer bölümler etkilenmez."""
    if pulse is None:
        return None
    try:
        n = int(bukum_op or 1)
    except (TypeError, ValueError):
        n = 1
    if n <= 1:
        return pulse
    return pulse // n


def _bukum_excel_yaz(ref, adet):
    """Büküm operasyon sayısını 'Pres Abkant Referans' sayfasına yazar (D kolonu).
    Sayfa 3 kolonluydu (Kod | Açıklama | Süre) → D yeni. Başlık yoksa açılır.
    EN İYİ ÇABA: Excel biri tarafından açıksa PermissionError gelir, üretim kaydı
    bundan ETKİLENMEZ — hata yalnız loglanır (bkz. referans/sureler aynı kalıp)."""
    try:
        from import_excel import EXCEL_YOL, BOLUM_SAYFA
        import openpyxl, os
        if not os.path.exists(EXCEL_YOL):
            return
        wb = openpyxl.load_workbook(EXCEL_YOL)
        sayfa_adi = BOLUM_SAYFA['pres']['ref']
        if sayfa_adi not in wb.sheetnames:
            return
        ws = wb[sayfa_adi]
        if not ws.cell(row=1, column=4).value:
            ws.cell(row=1, column=4, value='Büküm Op.')
        hedef = str(ref).upper().replace(' ', '')
        for r in range(2, ws.max_row + 1):
            hucre = ws.cell(row=r, column=1).value
            if hucre and str(hucre).upper().replace(' ', '') == hedef:
                ws.cell(row=r, column=4, value=int(adet))
                break
        wb.save(EXCEL_YOL)
    except Exception as e:
        print(f'[bukum] Excel sync hatasi ({ref}): {e}')


def _bukum_op_coz(c, gelen, ref, bolum, lokasyon):
    """Üretim kaydına yazılacak büküm operasyon sayısını belirler.

    - pres DIŞINDAKİ bölümlerde her zaman 1 → özellik kapalı, hiçbir şey değişmez.
    - İstemci değer gönderdiyse: referansa da YAZILIR (bir sonraki üretimde hazır
      gelsin) ve Excel'e senkronlanır — ama YALNIZ değer gerçekten değiştiyse,
      yoksa her üretim kaydında 700 satırlık Excel açılıp kaydedilirdi.
    - Göndermediyse: referansın hatırlanan değeri kullanılır (yoksa 1)."""
    if (bolum or '') != 'pres':
        return 1
    kayitli = 1
    try:
        row = c.execute(
            "SELECT COALESCE(bukum_operasyon,1) AS b FROM referans_listesi "
            "WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ','')) "
            "AND COALESCE(bolum,'kaynak')=? AND COALESCE(lokasyon,'TK2')=? LIMIT 1",
            (ref, bolum, lokasyon)
        ).fetchone()
        if row:
            kayitli = int(row['b'] or 1)
    except Exception:
        kayitli = 1

    if gelen in (None, '', 0, '0'):
        return max(1, kayitli)
    try:
        yeni = int(gelen)
    except (TypeError, ValueError):
        return max(1, kayitli)
    yeni = max(1, min(99, yeni))   # 1..99 — sıfır/negatif bölme çöktürür, 99 üstü hatalı giriş

    if yeni != kayitli:
        try:
            c.execute(
                "UPDATE referans_listesi SET bukum_operasyon=? "
                "WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ','')) "
                "AND COALESCE(bolum,'kaynak')=? AND COALESCE(lokasyon,'TK2')=?",
                (yeni, ref, bolum, lokasyon)
            )
        except Exception as e:
            print(f'[bukum] referans guncellenemedi ({ref}): {e}')
        _bukum_excel_yaz(ref, yeni)
    return yeni


def _uretim_sayac_senkron(conn, vardiya_id):
    """Auto modlu (sayac_otomatik=1) üretim kayıtlarının ok_adet'ini sensör sayımıyla
    günceller. GET /api/vardiya ve OEE okumadan önce çağrılır, böylece liste/rapor/OEE
    hep güncel sensör sayısını gösterir. İki sayaç kaynağı:
      • test_cihaz_id dolu → Cofle test cihazı (test_sonuclari, başarılı test)
      • aksi + robot allowlist'te → ESP32 pulse (pilot.db)"""
    try:
        v = conn.execute(
            "SELECT COALESCE(bolum,'kaynak') as bolum, robot_no FROM vardiyalar WHERE id=?",
            (vardiya_id,)
        ).fetchone()
        if not v:
            return
        rows = conn.execute(
            "SELECT id, istasyon, sayac_baslangic_ts, test_cihaz_id, referans_kodu, "
            "       COALESCE(bukum_operasyon, 1) AS bukum_operasyon "
            "FROM uretim_kayitlari "
            "WHERE vardiya_id=? AND sayac_otomatik=1 AND sayac_baslangic_ts IS NOT NULL",
            (vardiya_id,)
        ).fetchall()
        for r in rows:
            # Allowlist kontrolü SATIR bazlı: pres'te makine üretim kaydındadır
            # (istasyon), dolayısıyla desteklenirlik satırdan satıra değişebilir.
            robot_allow = bool(v['robot_no']) and _sayac_destekli(
                v['bolum'], v['robot_no'], r['istasyon'] or 0)
            if r['test_cihaz_id']:
                cnt = _test_basari_say(conn, r['test_cihaz_id'], r['referans_kodu'],
                                       r['sayac_baslangic_ts'])
            elif robot_allow:
                cnt = _pilot_pulse_say(v['bolum'], v['robot_no'], r['istasyon'] or 0,
                                       r['sayac_baslangic_ts'])
                cnt = _bukum_bol(cnt, r['bukum_operasyon'])
            else:
                cnt = None
            if cnt is not None:   # kaynak okunamadıysa ok_adet'i SIFIRLAMA (gerçek üretimi ezme)
                conn.execute("UPDATE uretim_kayitlari SET ok_adet=? WHERE id=?", (cnt, r['id']))
        conn.commit()
    except Exception as e:
        print(f"[_uretim_sayac_senkron] hata: {e}")


def _acik_auto_vardiyalari_senkronla(conn, bolum=None, lokasyon=None):
    """Bugünün AÇIK vardiyalarındaki TÜM auto sayaç kayıtlarını (ESP32 + test cihazı)
    sensör sayımıyla tazeler. Andon ve dashboard-özet okumaları ÖNCESİ çağrılır —
    yoksa canlı makine adetleri (pilot rozeti) akarken performans/OEE bayat ok_adet'le
    hesaplanır (operatör mobili açık değilse saatlerce güncellenmez). Yalnız AÇIK
    vardiyalar: kapalı vardiyanın sayacı kapanış değerinde donuk kalmalı."""
    try:
        sql = ("SELECT DISTINCT v.id FROM vardiyalar v "
               "JOIN uretim_kayitlari u ON u.vardiya_id = v.id "
               "WHERE v.tarih = date('now','localtime') "
               "AND COALESCE(v.durum,'acik') != 'kapali' AND u.sayac_otomatik = 1")
        params = []
        if bolum:
            sql += " AND COALESCE(v.bolum,'kaynak') = ?"
            params.append(bolum)
        if lokasyon:
            sql += " AND COALESCE(v.lokasyon,'TK2') = ?"
            params.append(lokasyon)
        for r in conn.execute(sql, params).fetchall():
            _uretim_sayac_senkron(conn, r['id'])
    except Exception as e:
        print(f"[auto-senkron] hata: {e}")


# ─────────────────────────────────────────────────────────────
# OTOMATIK MOLALAR — sabit saatlerde planlı duruş (çay + yemek)
# ─────────────────────────────────────────────────────────────
# Şirket politikası: 09:00/13:00/15:00 çay (10dk), 11:00 yemek (30dk). Vardiya o
# saatten ÖNCE açıldıysa ve saat GELDİYSE duruş otomatik işlenir — operatör elle
# girmesin. AYNI TÜRDEN molalar tek duruş kaydında BİRİKİR (süre üzerine eklenir);
# yeni satır sadece o türden kayıt yoksa açılır. cay_molasi_bayrak bitmask'i her
# molayı bir kez işler (operatör kaydı silse de geri gelmez). Sadece BUGÜN.
# Çay tüm bölümlerde aynı; YEMEK bölüme göre değişir:
#   metal → 'Maden Tkv. ve Yemek Molası' 40dk (döküm alanı: maden takviyesi + yemek)
#   diğer → 'Yemek Molası' 30dk
def _bolum_molalari(bolum):
    m = [
        (1, '09:00', 10, 'Çay Molası'),
        (2, '13:00', 10, 'Çay Molası'),
        (4, '15:00', 10, 'Çay Molası'),
    ]
    if (bolum or 'kaynak') == 'metal':
        m.append((8, '11:00', 40, 'Maden Tkv. ve Yemek Molası'))
    else:
        m.append((8, '11:00', 30, 'Yemek Molası'))
    return sorted(m, key=lambda x: x[1])

def _saat_to_time(s):
    """'HH:MM' / 'HH:MM:SS' metnini datetime.time'a çevir; bozuksa None."""
    s = (s or '').strip()
    try:
        return datetime.strptime(s[:5], '%H:%M').time()
    except ValueError:
        return None

_TR_FOLD = str.maketrans('çÇıİöÖşŞüÜğĞ', 'cciioossuugg')
def _mola_norm(s):
    """Sebep adını eşleştirme için normalize et: Türkçe harf katla + küçült.
    'Çay Molası' / 'Cay Molasi' / 'ÇAY MOLASI' → 'cay molasi'."""
    return (s or '').translate(_TR_FOLD).lower().strip()

def _otomatik_mola_uygula(conn, vardiya):
    """Saati gelmiş otomatik molaları bu vardiyaya işler. Aynı türden planlı kayıt
    varsa süresini ARTIRIR, yoksa yeni satır açar. Idempotent (bitmask). Caller
    commit eder. Sadece bugünkü vardiya; eksik kolon/bozuk veri sessizce geçilir."""
    try:
        if 'cay_molasi_bayrak' not in vardiya.keys():
            return                      # migration henüz çalışmadı
    except Exception:
        return
    if (vardiya['tarih'] or '') != datetime.now().strftime('%Y-%m-%d'):
        return                          # geçmiş ayrı (tek seferlik düzenleme yapıldı)
    bas_t = _saat_to_time(vardiya['baslangic_saati'])
    if bas_t is None:
        return
    # Etkin bitiş: vardiya açıksa ŞİMDİ, kapalıysa bitiş saati
    if (vardiya['durum'] or '') == 'kapali':
        etkin_bitis = _saat_to_time(vardiya['bitis_saati']) or datetime.now().time()
    else:
        etkin_bitis = datetime.now().time()
    try:
        vardiya_bolum = (vardiya['bolum'] if 'bolum' in vardiya.keys() else 'kaynak') or 'kaynak'
    except Exception:
        vardiya_bolum = 'kaynak'
    bayrak = vardiya['cay_molasi_bayrak'] or 0
    yeni = bayrak
    for bit, mola_saati, mola_dk, sebep in _bolum_molalari(vardiya_bolum):
        if bayrak & bit:
            continue                    # bu mola zaten islendi
        mola_t = _saat_to_time(mola_saati)
        # Vardiya moladan ÖNCE açılmış + mola saati gelmiş olmalı
        if not (mola_t and bas_t < mola_t <= etkin_bitis):
            continue
        # Aynı türden mevcut planlı kayıt var mı? (yazım varyantlarına dayanıklı)
        hedef_id = None
        for r in conn.execute(
                "SELECT id, durus_sebebi FROM duruslar WHERE vardiya_id=? AND durus_tipi='planli' ORDER BY id",
                (vardiya['id'],)).fetchall():
            if _mola_norm(r['durus_sebebi']) == _mola_norm(sebep):
                hedef_id = r['id']
                break
        if hedef_id is not None:
            # BİRİKTİR: süreyi üzerine ekle, açıklamaya saati not düş
            conn.execute(
                "UPDATE duruslar SET sure_dk = sure_dk + ?, "
                "aciklama = CASE WHEN COALESCE(aciklama,'')='' THEN ? ELSE aciklama || ' + ' || ? END "
                "WHERE id=?",
                (mola_dk, f'Otomatik: {mola_saati} ({mola_dk}dk)', f'{mola_saati} ({mola_dk}dk)', hedef_id)
            )
        else:
            conn.execute(
                "INSERT INTO duruslar (vardiya_id, durus_sebebi, aciklama, sure_dk, baslangic_saati, durus_tipi) "
                "VALUES (?, ?, ?, ?, ?, 'planli')",
                (vardiya['id'], sebep, f'Otomatik: {mola_saati} ({mola_dk}dk)', mola_dk, mola_saati)
            )
        yeni |= bit
    if yeni != bayrak:
        conn.execute("UPDATE vardiyalar SET cay_molasi_bayrak=? WHERE id=?", (yeni, vardiya['id']))


# ─────────────────────────────────────────────────────────────
# WEB PUSH BİLDİRİM — operatör telefonuna (VAPID)
# ─────────────────────────────────────────────────────────────
# Operatör mobil sayfada PIN ile girince tarayıcı push aboneliği kaydedilir
# (/api/push/abone). Launch eklenince BILDIRIM_ALICILARI'ndaki operatörlerin
# tüm cihazlarına bildirim gider. Anahtarlar ilk kullanımda üretilip
# genel_ayarlar'da saklanır. pywebpush yoksa özellik sessizce devre dışı.
_vapid_cache = {}

def _vapid_keys():
    """VAPID anahtar çifti (yoksa üret, genel_ayarlar'a yaz). {'priv','pub'} | None."""
    if _vapid_cache:
        return _vapid_cache
    try:
        import base64
        from py_vapid import Vapid
        from cryptography.hazmat.primitives import serialization
        conn = db_connect()
        try:
            r_priv = conn.execute("SELECT deger FROM genel_ayarlar WHERE anahtar='push_vapid_priv'").fetchone()
            r_pub  = conn.execute("SELECT deger FROM genel_ayarlar WHERE anahtar='push_vapid_pub'").fetchone()
            if r_priv and r_pub:
                try:
                    Vapid.from_string(r_priv['deger'])   # format dogrulamasi (self-heal)
                    _vapid_cache.update(priv=r_priv['deger'], pub=r_pub['deger'])
                    return _vapid_cache
                except Exception:
                    print('[PUSH] Kayitli VAPID anahtari gecersiz formatta — yeniden uretiliyor')
            v = Vapid()
            v.generate_keys()
            # pywebpush, vapid_private_key olarak 32-byte HAM degerin base64url'unu bekler
            # (Vapid.from_string -> from_raw). PEM DEGIL.
            priv_raw = v.private_key.private_numbers().private_value.to_bytes(32, 'big')
            priv_b64 = base64.urlsafe_b64encode(priv_raw).rstrip(b'=').decode()
            pub_raw = v.public_key.public_bytes(
                serialization.Encoding.X962, serialization.PublicFormat.UncompressedPoint)
            pub_b64 = base64.urlsafe_b64encode(pub_raw).rstrip(b'=').decode()
            conn.execute("INSERT OR REPLACE INTO genel_ayarlar (anahtar, deger) VALUES ('push_vapid_priv', ?)", (priv_b64,))
            conn.execute("INSERT OR REPLACE INTO genel_ayarlar (anahtar, deger) VALUES ('push_vapid_pub', ?)", (pub_b64,))
            conn.commit()
            _vapid_cache.update(priv=priv_b64, pub=pub_b64)
            return _vapid_cache
        finally:
            conn.close()
    except Exception as e:
        print(f'[PUSH] VAPID anahtar hatası: {e}')
        return None

def _push_gonder(alici_adlari, baslik, govde, url='/'):
    """Adı geçen operatörlerin TÜM cihazlarına push gönder (senkron).
    Ölü abonelikler (404/410) otomatik silinir. Hata bildirimi düşürmez."""
    try:
        from pywebpush import webpush, WebPushException
    except ImportError:
        print('[PUSH] pywebpush kurulu değil — bildirim atlandı')
        return
    keys = _vapid_keys()
    if not keys:
        return
    hedef = {_mola_norm(a) for a in alici_adlari}
    payload = json.dumps({'title': baslik, 'body': govde, 'url': url})
    try:
        conn = db_connect()
    except Exception as e:
        print(f'[PUSH] DB: {e}')
        return
    try:
        subs = [s for s in conn.execute('SELECT * FROM push_abonelikler').fetchall()
                if _mola_norm(s['operator_adi']) in hedef]
        for s in subs:
            try:
                webpush(
                    subscription_info={'endpoint': s['endpoint'],
                                       'keys': {'p256dh': s['p256dh'], 'auth': s['auth']}},
                    data=payload,
                    vapid_private_key=keys['priv'],
                    vapid_claims={'sub': 'mailto:bildirim@coflemanage.online'},
                    timeout=10,
                )
            except WebPushException as e:
                kod = getattr(getattr(e, 'response', None), 'status_code', None)
                if kod in (404, 410):   # abonelik ölmüş (uygulama silinmiş vb.)
                    conn.execute('DELETE FROM push_abonelikler WHERE id=?', (s['id'],))
                    conn.commit()
                else:
                    print(f'[PUSH] gönderim hatası ({s["operator_adi"]}): {e}')
            except Exception as e:
                print(f'[PUSH] {e}')
    finally:
        conn.close()

def _push_gonder_async(alici_adlari, baslik, govde, url='/'):
    """Push'u arka plan thread'inde gönder — HTTP yanıtını bekletmez."""
    import threading
    threading.Thread(target=_push_gonder, args=(list(alici_adlari), baslik, govde, url),
                     daemon=True).start()


@app.route('/sw.js')
def service_worker_js():
    """Service worker — kök kapsam (push almak için / scope şart)."""
    resp = send_file(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static', 'sw.js'),
                     mimetype='application/javascript')
    resp.headers['Cache-Control'] = 'no-cache'
    return resp


@app.route('/api/push/vapid', methods=['GET'])
def push_vapid_public():
    keys = _vapid_keys()
    if not keys:
        return jsonify({'hata': 'push devre dışı'}), 503
    return jsonify({'public_key': keys['pub']})


@app.route('/api/push/abone', methods=['POST'])
@operator_required
def push_abone():
    """Operatör mobil cihazının push aboneliğini kaydet (PIN doğrulamalı)."""
    if not getattr(g, 'operator_adi', None):
        return jsonify({'hata': 'Operatör girişi gerekli'}), 400
    sub = (request.get_json() or {}).get('subscription') or {}
    endpoint = (sub.get('endpoint') or '').strip()
    anahtarlar = sub.get('keys') or {}
    p256dh, auth = anahtarlar.get('p256dh'), anahtarlar.get('auth')
    if not (endpoint and p256dh and auth):
        return jsonify({'hata': 'Geçersiz abonelik'}), 400
    conn = get_db()
    # endpoint UNIQUE: aynı cihaz tekrar abone olursa operatör adı güncellenir
    conn.execute('''
        INSERT INTO push_abonelikler (operator_adi, endpoint, p256dh, auth)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(endpoint) DO UPDATE SET
            operator_adi = excluded.operator_adi,
            p256dh = excluded.p256dh,
            auth = excluded.auth
    ''', (g.operator_adi, endpoint, p256dh, auth))
    conn.commit()
    return jsonify({'basarili': True})


def _vardiya_efektif_dk(vardiya):
    """Vardiyanın EFEKTİF süresi (dk) — canlı performans/kullanılabilirlik paydası.
    Kapalı: toplam_sure_dk (gerçek açılış→kapanış; kapatırken yazılır).
    Açık (bugün): başlangıçtan ŞU ANA kadar geçen dakika — süre ilerledikçe artar,
    performans buna göre canlı değişir. Açık ama eski günde unutulmuş: kayıtlı süre."""
    try:
        durum = (vardiya['durum'] or 'kapali')
        if durum == 'kapali':
            return vardiya['toplam_sure_dk'] or 0
        if (vardiya['tarih'] or '') == datetime.now().strftime('%Y-%m-%d'):
            bas = _saat_to_time(vardiya['baslangic_saati'])
            if bas is None:
                return vardiya['toplam_sure_dk'] or 0
            simdi = datetime.now().time()
            gecen = (simdi.hour * 60 + simdi.minute) - (bas.hour * 60 + bas.minute)
            if gecen < 0:
                gecen += 1440   # gece yarısını geçen vardiya
            return max(0, gecen)
        return vardiya['toplam_sure_dk'] or 0
    except Exception:
        try:
            return vardiya['toplam_sure_dk'] or 0
        except Exception:
            return 0


# ─────────────────────────────────────────────────────────────
# SAYFA ROUTES
# ─────────────────────────────────────────────────────────────

@app.route('/')
def operator_sayfasi():
    """Operatör mobil giriş sayfası (v2). Varsayılan lokasyon: TK2 (mevcut fabrika)."""
    return render_template('mobile_v2.html', lokasyon='TK2')


@app.route('/mobile_v2')
def operator_v2_preview():
    """Operatör v2 önizlemesi (canlı ile aynı şablon — direkt link için saklanıyor)."""
    return render_template('mobile_v2.html', lokasyon='TK2')


@app.route('/tk1')
def operator_tk1_sayfasi():
    """TK1 (yan tesis) operatör mobil veri giriş sayfası. Lokasyon URL'den SABİT (TK1);
    operatör ekranda TK1/TK2 seçmez — bu adres TK1 operatörlerinin telefonu içindir."""
    return render_template('mobile_v2.html', lokasyon='TK1')


@app.route('/onizleme')
def operator_onizleme_sayfasi():
    """UI v3 "ENDÜSTRİYEL" ÖNİZLEMESİ (2026-07-29, design_handoff_ui_v3) —
    operatör ekranı adayı; canlı / ve /tk1'e DOKUNMAZ. OK/OEE sayaç halkası +
    ink/amber/kırmızı-sınır CTA reçetesi. Beğenilirse mobile_onizleme.html
    içeriği mobile_v2.html'e taşınır."""
    return render_template('mobile_onizleme.html', lokasyon='TK2')


@app.route('/mobile_legacy')
def operator_legacy_sayfasi():
    """Eski operatör tasarımı (geri dönüş için saklanıyor)."""
    return render_template('mobile.html')


@app.route('/dashboard')
def dashboard_sayfasi():
    """Yönetici dashboard — YENİ tasarım (v2: önce genel bakış, gruplu menü, i18n).
    Oturum yoksa giriş ekranı sunulur. Kullanıcının rol/izin bilgisi şablona
    enjekte edilir (ekran titremeden yetkisiz sayfalar gizlenir)."""
    ku = panel_kullanici()
    if not ku:
        return render_template('panel_giris.html')
    return render_template('dashboard_v2.html', panel_ku=json.dumps({
        'kullanici_adi': ku['kullanici_adi'],
        'ad_soyad': ku['ad_soyad'],
        'rol': ku['rol'],
        'admin': ku['admin'],
        'izinler': ku['izinler'],
        'sifre_gecici': ku['sifre_gecici'],
    }, ensure_ascii=False))


@app.route('/dashboard/onizleme')
def dashboard_onizleme_sayfasi():
    """UI v3 "ENDÜSTRİYEL" ÖNİZLEMESİ (2026-07-29, design_handoff_ui_v3) —
    canlı /dashboard'a DOKUNMAZ. Nötr gri + indigo + ink chrome; düz yüzeyler,
    gölgesiz, mono rakam. Beğenilirse dashboard_onizleme.html içeriği
    dashboard_v2.html'e taşınır. Auth/panel_ku canlı rotayla birebir aynı."""
    ku = panel_kullanici()
    if not ku:
        return render_template('panel_giris.html')
    return render_template('dashboard_onizleme.html', panel_ku=json.dumps({
        'kullanici_adi': ku['kullanici_adi'],
        'ad_soyad': ku['ad_soyad'],
        'rol': ku['rol'],
        'admin': ku['admin'],
        'izinler': ku['izinler'],
        'sifre_gecici': ku['sifre_gecici'],
    }, ensure_ascii=False))


@app.route('/dashboard/tezgah')
def dashboard_tezgah_sayfasi():
    """TASARIM ADAYI 2 — "Tezgah / Sıcak Nötr" (2026-07-28). Marka turuncusu
    (logo #EC6736) birincil kimlik, ılık nötr zemin, yoğun tablo, süs katmanı yok.
    /dashboard/onizleme (aday 1) ve /dashboard (canlı) ile yan yana karşılaştırılır.
    Canlıya DOKUNMAZ."""
    ku = panel_kullanici()
    if not ku:
        return render_template('panel_giris.html')
    return render_template('dashboard_tezgah.html', panel_ku=json.dumps({
        'kullanici_adi': ku['kullanici_adi'],
        'ad_soyad': ku['ad_soyad'],
        'rol': ku['rol'],
        'admin': ku['admin'],
        'izinler': ku['izinler'],
        'sifre_gecici': ku['sifre_gecici'],
    }, ensure_ascii=False))


@app.route('/dashboard_eski')
def dashboard_eski_sayfasi():
    """Eski yönetim paneli (v2 canlıya alınınca yedek olarak saklanıyor).
    Bu panel sayfa-izni filtresi TAŞIMAZ (her şeyi gösterir) — bu yüzden yalnız
    admin erişebilir; kısıtlı kullanıcı buradan kısıtlamayı aşamaz."""
    ku = panel_kullanici()
    if not ku:
        return render_template('panel_giris.html')
    if not ku['admin']:
        return redirect('/dashboard')
    return render_template('dashboard.html')


@app.route('/dashboard_legacy')
def dashboard_legacy_sayfasi():
    """Eski dashboard tasarımı (geri dönüş için saklanıyor). Admin-only (yukarıdaki
    gerekçeyle — izin filtresi yok)."""
    ku = panel_kullanici()
    if not ku:
        return render_template('panel_giris.html')
    if not ku['admin']:
        return redirect('/dashboard')
    return render_template('dashboard_legacy.html')


# ═══════════════════════ PANEL GİRİŞ / OTURUM ═══════════════════════
@app.route('/api/panel/giris', methods=['POST'])
def panel_giris():
    """Panel (dashboard) giriş. Body: {kullanici_adi, sifre}.
    Başarılı → imzalı oturum cookie'si kurulur. 401 = hatalı/pasif."""
    data = request.get_json(silent=True) or {}
    ka = (data.get('kullanici_adi') or '').strip().lower()
    sifre = (data.get('sifre') or '')
    if not ka or not sifre:
        return jsonify({'hata': 'Kullanıcı adı ve şifre zorunlu'}), 400
    conn = get_db()
    row = conn.execute(
        "SELECT id, sifre_hash, aktif FROM panel_kullanicilari WHERE lower(kullanici_adi)=?",
        (ka,)
    ).fetchone()
    if not row or not check_password_hash(row['sifre_hash'], sifre):
        return jsonify({'hata': 'Kullanıcı adı veya şifre hatalı'}), 401
    if not row['aktif']:
        return jsonify({'hata': 'Bu hesap pasif durumda. Yöneticinize başvurun.'}), 403
    conn.execute("UPDATE panel_kullanicilari SET son_giris=datetime('now','localtime') WHERE id=?", (row['id'],))
    conn.commit()
    session.permanent = True
    session['panel_uid'] = row['id']
    ku = panel_kullanici()
    return jsonify({'basarili': True, 'kullanici': {
        'kullanici_adi': ku['kullanici_adi'],
        'ad_soyad': ku['ad_soyad'],
        'rol': ku['rol'],
        'admin': ku['admin'],
        'izinler': ku['izinler'],
        'sifre_gecici': ku['sifre_gecici'],
    }}), 200


@app.route('/api/panel/cikis', methods=['POST'])
def panel_cikis():
    """Oturumu kapatır."""
    session.pop('panel_uid', None)
    session.clear()
    return jsonify({'basarili': True}), 200


@app.route('/api/panel/oturum', methods=['GET'])
def panel_oturum():
    """Aktif oturum bilgisi (frontend yetki senkronu için)."""
    ku = panel_kullanici()
    if not ku:
        return jsonify({'oturum': False}), 200
    return jsonify({'oturum': True, 'kullanici': {
        'kullanici_adi': ku['kullanici_adi'],
        'ad_soyad': ku['ad_soyad'],
        'rol': ku['rol'],
        'admin': ku['admin'],
        'izinler': ku['izinler'],
        'sifre_gecici': ku['sifre_gecici'],
    }}), 200


@app.route('/api/panel/sifre_degistir', methods=['POST'])
def panel_sifre_degistir():
    """Giriş yapmış kullanıcı kendi şifresini değiştirir.
    Body: {eski_sifre, yeni_sifre}. Geçici şifre bayrağı temizlenir."""
    ku = panel_kullanici()
    if not ku:
        return jsonify({'hata': 'Oturum gerekli', 'giris_gerekli': True}), 401
    data = request.get_json(silent=True) or {}
    eski = (data.get('eski_sifre') or '')
    yeni = (data.get('yeni_sifre') or '')
    if len(yeni) < 6:
        return jsonify({'hata': 'Yeni şifre en az 6 karakter olmalı'}), 400
    conn = get_db()
    row = conn.execute("SELECT sifre_hash FROM panel_kullanicilari WHERE id=?", (ku['id'],)).fetchone()
    if not row or not check_password_hash(row['sifre_hash'], eski):
        return jsonify({'hata': 'Mevcut şifre hatalı'}), 403
    conn.execute(
        "UPDATE panel_kullanicilari SET sifre_hash=?, sifre_gecici=0 WHERE id=?",
        (generate_password_hash(yeni), ku['id'])
    )
    conn.commit()
    return jsonify({'basarili': True}), 200


# ═══════════════════════ PANEL KULLANICI YÖNETİMİ (admin) ═══════════════════════
def _panel_ku_json(row):
    try:
        izinler = json.loads(row['izinler'] or '[]')
        if not isinstance(izinler, list):
            izinler = []
    except Exception:
        izinler = []
    admin = (row['rol'] == 'admin')
    return {
        'id': row['id'],
        'kullanici_adi': row['kullanici_adi'],
        'ad_soyad': row['ad_soyad'] or '',
        'rol': row['rol'],
        'izinler': PANEL_SAYFALAR[:] if admin else [s for s in izinler if s in PANEL_SAYFALAR],
        'aktif': bool(row['aktif']),
        'sifre_gecici': bool(row['sifre_gecici']),
        'son_giris': row['son_giris'] if 'son_giris' in row.keys() else None,
    }


def _temiz_izinler(deger):
    """Gelen izin listesini geçerli sayfa anahtarlarına süzer."""
    if not isinstance(deger, list):
        return []
    return [s for s in deger if s in PANEL_SAYFALAR]


@app.route('/api/panel/kullanicilar', methods=['GET'])
@panel_gerekli(admin=True)
def panel_kullanicilar_liste():
    """Tüm panel kullanıcılarını listeler (admin)."""
    conn = get_db()
    rows = conn.execute(
        "SELECT id, kullanici_adi, ad_soyad, rol, izinler, aktif, sifre_gecici, son_giris "
        "FROM panel_kullanicilari ORDER BY (rol='admin') DESC, kullanici_adi"
    ).fetchall()
    return jsonify({'sayfalar': PANEL_SAYFALAR, 'kullanicilar': [_panel_ku_json(r) for r in rows]}), 200


@app.route('/api/panel/kullanicilar', methods=['POST'])
@panel_gerekli(admin=True)
def panel_kullanici_ekle():
    """Yeni panel kullanıcısı (admin). Body:
    {kullanici_adi, ad_soyad, sifre, rol, izinler[]}. sifre_gecici=1 verilir."""
    data = request.get_json(silent=True) or {}
    ka = (data.get('kullanici_adi') or '').strip().lower()
    ad_soyad = (data.get('ad_soyad') or '').strip()
    sifre = (data.get('sifre') or '')
    rol = 'admin' if (data.get('rol') == 'admin') else 'kullanici'
    izinler = _temiz_izinler(data.get('izinler'))
    if not ka:
        return jsonify({'hata': 'Kullanıcı adı zorunlu'}), 400
    if len(sifre) < 6:
        return jsonify({'hata': 'Şifre en az 6 karakter olmalı'}), 400
    conn = get_db()
    var = conn.execute("SELECT 1 FROM panel_kullanicilari WHERE lower(kullanici_adi)=?", (ka,)).fetchone()
    if var:
        return jsonify({'hata': f'Bu kullanıcı adı zaten var: {ka}'}), 409
    conn.execute(
        "INSERT INTO panel_kullanicilari (kullanici_adi, ad_soyad, sifre_hash, rol, izinler, aktif, sifre_gecici) "
        "VALUES (?, ?, ?, ?, ?, 1, 1)",
        (ka, ad_soyad, generate_password_hash(sifre), rol, json.dumps(izinler))
    )
    conn.commit()
    return jsonify({'basarili': True, 'kullanici_adi': ka}), 201


@app.route('/api/panel/kullanicilar/<int:uid>', methods=['PATCH'])
@panel_gerekli(admin=True)
def panel_kullanici_guncelle(uid):
    """Kullanıcı ad_soyad / rol / izinler / aktif günceller (admin).
    Kilitlenme koruması: admin kendi rolünü düşüremez/kendini pasifleştiremez;
    son aktif admin'in rolü düşürülemez."""
    data = request.get_json(silent=True) or {}
    conn = get_db()
    row = conn.execute("SELECT id, rol, aktif FROM panel_kullanicilari WHERE id=?", (uid,)).fetchone()
    if not row:
        return jsonify({'hata': 'Kullanıcı bulunamadı'}), 404
    ben = g.panel_ku
    yeni_rol = row['rol']
    if 'rol' in data:
        yeni_rol = 'admin' if (data.get('rol') == 'admin') else 'kullanici'
    yeni_aktif = row['aktif']
    if 'aktif' in data:
        yeni_aktif = 1 if data.get('aktif') else 0

    # Kendini kilitleme koruması
    if uid == ben['id']:
        if yeni_rol != 'admin':
            return jsonify({'hata': 'Kendi yönetici rolünüzü kaldıramazsınız'}), 400
        if not yeni_aktif:
            return jsonify({'hata': 'Kendi hesabınızı pasifleştiremezsiniz'}), 400
    # Son admin koruması
    if row['rol'] == 'admin' and (yeni_rol != 'admin' or not yeni_aktif):
        aktif_admin = conn.execute(
            "SELECT COUNT(*) c FROM panel_kullanicilari WHERE rol='admin' AND aktif=1"
        ).fetchone()['c']
        if aktif_admin <= 1:
            return jsonify({'hata': 'Sistemde en az bir aktif yönetici kalmalı'}), 400

    ad_soyad = data.get('ad_soyad')
    if ad_soyad is None:
        ad_soyad = conn.execute("SELECT ad_soyad FROM panel_kullanicilari WHERE id=?", (uid,)).fetchone()['ad_soyad']
    else:
        ad_soyad = ad_soyad.strip()
    izinler = _temiz_izinler(data['izinler']) if 'izinler' in data else None
    if izinler is None:
        izinler_json = conn.execute("SELECT izinler FROM panel_kullanicilari WHERE id=?", (uid,)).fetchone()['izinler']
    else:
        izinler_json = json.dumps(izinler)
    conn.execute(
        "UPDATE panel_kullanicilari SET ad_soyad=?, rol=?, izinler=?, aktif=? WHERE id=?",
        (ad_soyad, yeni_rol, izinler_json, yeni_aktif, uid)
    )
    conn.commit()
    return jsonify({'basarili': True}), 200


@app.route('/api/panel/kullanicilar/<int:uid>/sifre', methods=['POST'])
@panel_gerekli(admin=True)
def panel_kullanici_sifre_sifirla(uid):
    """Admin bir kullanıcının şifresini sıfırlar. Body: {sifre}.
    sifre_gecici=1 → kullanıcı ilk girişte değiştirmeye zorlanır."""
    data = request.get_json(silent=True) or {}
    sifre = (data.get('sifre') or '')
    if len(sifre) < 6:
        return jsonify({'hata': 'Şifre en az 6 karakter olmalı'}), 400
    conn = get_db()
    row = conn.execute("SELECT id FROM panel_kullanicilari WHERE id=?", (uid,)).fetchone()
    if not row:
        return jsonify({'hata': 'Kullanıcı bulunamadı'}), 404
    conn.execute(
        "UPDATE panel_kullanicilari SET sifre_hash=?, sifre_gecici=1 WHERE id=?",
        (generate_password_hash(sifre), uid)
    )
    conn.commit()
    return jsonify({'basarili': True}), 200


@app.route('/api/panel/kullanicilar/<int:uid>', methods=['DELETE'])
@panel_gerekli(admin=True)
def panel_kullanici_sil(uid):
    """Kullanıcı siler (admin). Kendini veya son aktif admin'i silemez."""
    ben = g.panel_ku
    if uid == ben['id']:
        return jsonify({'hata': 'Kendi hesabınızı silemezsiniz'}), 400
    conn = get_db()
    row = conn.execute("SELECT rol, aktif FROM panel_kullanicilari WHERE id=?", (uid,)).fetchone()
    if not row:
        return jsonify({'hata': 'Kullanıcı bulunamadı'}), 404
    if row['rol'] == 'admin' and row['aktif']:
        aktif_admin = conn.execute(
            "SELECT COUNT(*) c FROM panel_kullanicilari WHERE rol='admin' AND aktif=1"
        ).fetchone()['c']
        if aktif_admin <= 1:
            return jsonify({'hata': 'Sistemde en az bir aktif yönetici kalmalı'}), 400
    conn.execute("DELETE FROM panel_kullanicilari WHERE id=?", (uid,))
    conn.commit()
    return jsonify({'basarili': True}), 200


@app.route('/logo')
def logo_serve():
    """Marka logosu (COFLE FORGE — 2026-07-29 kimlik yenilemesi).

    ?tema=koyu → static/logo_koyu.png   (wordmark BEYAZ; koyu zeminli ekranlar:
                                         andon panoları, giriş ekranı)
    ?tip=mark  → static/logo_mark.png   (SEMBOL tek başına, 512×512 kare)
                 Sekme ikonu/PWA için: 1600×480 lockup 16×16'ya sıkışınca
                 okunamayan bir çizgiye dönüşüyordu.
    varsayılan → static/logo.png        (açık tema lockup)

    Tercih sırası: istenen varyant → static/logo.png → masaüstü override →
    static/logo.svg. static/logo.png BAŞTA: masaüstündeki eski 'LOGO_COFLE ONLY.png'
    varsa bile repodaki resmi logo kazanır (handoff README'si tersini söylüyor,
    o bilgi static/logo.svg'deki bayat yoruma dayanıyor — kod hep böyleydi)."""
    proje_dir = os.path.dirname(os.path.abspath(__file__))
    if (request.args.get('tip') or '').strip() == 'mark':
        mark = os.path.join(proje_dir, 'static', 'logo_mark.png')
        if os.path.exists(mark):
            return send_file(mark, mimetype='image/png')
        mark_svg = os.path.join(proje_dir, 'static', 'logo_mark.svg')
        if os.path.exists(mark_svg):
            return send_file(mark_svg, mimetype='image/svg+xml')
    if (request.args.get('tema') or '').strip() == 'koyu':
        koyu = os.path.join(proje_dir, 'static', 'logo_koyu.png')
        if os.path.exists(koyu):
            return send_file(koyu, mimetype='image/png')
    logo_yollar = [
        (os.path.join(proje_dir, 'static', 'logo.png'), 'image/png'),
        (os.path.join(os.path.expanduser('~'), 'OneDrive', 'Masaüstü', 'LOGO_COFLE ONLY.png'), 'image/png'),
        (os.path.join(os.path.expanduser('~'), 'Desktop', 'LOGO_COFLE ONLY.png'), 'image/png'),
        (os.path.join(proje_dir, 'static', 'logo.svg'), 'image/svg+xml'),
    ]
    for yol, mime in logo_yollar:
        if os.path.exists(yol):
            return send_file(yol, mimetype=mime)
    return '', 404


# ───── Favicon / iOS / Android home-screen ikonları ─────
# Telefonda "Ana ekrana ekle" denildiğinde Cofle logosu görünür.
@app.route('/favicon.ico')
@app.route('/favicon.png')
@app.route('/apple-touch-icon.png')
@app.route('/apple-touch-icon-precomposed.png')
def favicon_serve():
    """Tüm tarayıcı/iOS/Android ikon istekleri SEMBOLE yönlendirilir.

    Eskiden lockup (1600×480) dönülüyordu: 16×16 sekme ikonuna küçülünce
    okunamayan bir çizgi oluyordu. logo_mark.png kare (512×512) ve tek başına
    okunur. Kaynak: handoff README, madde 5."""
    proje_dir = os.path.dirname(os.path.abspath(__file__))
    mark = os.path.join(proje_dir, 'static', 'logo_mark.png')
    if os.path.exists(mark):
        return send_file(mark, mimetype='image/png')
    return logo_serve()


@app.route('/manifest.json')
def web_manifest():
    """Android/iOS PWA manifesti — ana ekrana eklendiğinde Cofle logosu + neon tema.

    ?lokasyon=TK1 → start_url '/tk1' (TK1 operatör telefonu ana ekran kısayolu TK1'den
    açılır; yoksa PWA '/' açıp localStorage lokasyonu TK2'ye ezerdi ve TK2 operatörleri
    listelenirdi). id alanı lokasyona göre farklı → iki tesis kısayolu aynı telefonda
    ayrı uygulama olarak yaşayabilir.
    """
    lokasyon = (request.args.get('lokasyon') or 'TK2').upper()
    tk1 = (lokasyon == 'TK1')
    return jsonify({
        "id": "/tk1" if tk1 else "/",
        "name": "Cofle Forge TK1" if tk1 else "Cofle Forge",
        "short_name": "Cofle TK1" if tk1 else "Cofle",
        "description": "Cofle Forge Üretim Takip Sistemi" + (" — TK1" if tk1 else ""),
        "start_url": "/tk1" if tk1 else "/",
        "scope": "/",
        "display": "standalone",
        "orientation": "any",
        # Açılış (splash) renkleri UYGULAMANIN GERÇEK zeminiyle eşleşmeli: eskiden
        # koyu mor (#060414/#1c1f3a) idi, operatör ekranı tezgah paletine geçince
        # koyu bir splash'tan açık bir uygulamaya düşüyordu (göz alıyordu).
        "background_color": "#ECECF1",
        "theme_color": "#ECECF1",
        "lang": "tr",
        # PWA ikonu ayrı asset: /logo lockup'ı (1600×480) kare ikon yuvasına
        # sıkıştırılıyordu. pwa_icon.png 512×512 kare (handoff README, madde 6).
        "icons": [
            {"src": "/static/pwa_icon.png", "sizes": "192x192", "type": "image/png", "purpose": "any"},
            {"src": "/static/pwa_icon.png", "sizes": "512x512", "type": "image/png", "purpose": "any"},
            {"src": "/static/pwa_icon.png", "sizes": "512x512", "type": "image/png", "purpose": "maskable"}
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


@app.route('/andon_tk1')
def andon_tk1_sayfasi():
    """TK1 (yan tesis) Andon ekranı — montaj mantığı, lokasyon=TK1 (v5 tasarım)."""
    return render_template('andon_v5.html',
                           bolum='montaj',
                           lokasyon='TK1',
                           bolum_ad='TK1 Montaj',
                           bolum_ikon='🏢',
                           panel_baslik='TK1 Hat Durum Paneli')


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
    lokasyon = request.args.get('lokasyon')
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
    if lokasyon:
        query += " AND COALESCE(lokasyon, 'TK2') = ?"
        params.append(lokasyon)

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

    zorunlu = ['tarih', 'vardiya_turu', 'robot_no', 'operator_adi']
    for alan in zorunlu:
        if not data.get(alan):
            return jsonify({'hata': f'"{alan}" alanı zorunludur'}), 400

    # Operatör mobile'dan gelen istekte header'daki operator = body'deki operator olmalı
    # (başkasının adına vardiya açma engeli). Dashboard modunda g.operator_adi None.
    # Admin bu kısıttan muaf — başka operatör adına vardiya açabilir.
    if (getattr(g, 'operator_adi', None) and g.operator_adi != data['operator_adi']
            and g.operator_adi != ADMIN_ADI):
        return jsonify({'hata': f'Sadece kendi adınıza vardiya açabilirsiniz ({g.operator_adi})'}), 403

    # Başlangıç: verilmediyse ŞU AN (operatör vardiyayı açtığı saat — sabit 07:00 yok).
    # Bitiş: vardiya KAPATILINCA belli olur; açılışta boş kalabilir (toplam_sure_dk=0,
    # açık vardiyada performans zaten "o ana kadar geçen süre" üzerinden hesaplanır).
    bas_str = (data.get('baslangic_saati') or '').strip() or datetime.now().strftime('%H:%M')
    bit_str = (data.get('bitis_saati') or '').strip()
    try:
        bas = datetime.strptime(bas_str, '%H:%M')
    except ValueError:
        return jsonify({'hata': 'Başlangıç saat formatı HH:MM olmalıdır'}), 400
    fark_dk = 0
    if bit_str:
        try:
            bit = datetime.strptime(bit_str, '%H:%M')
        except ValueError:
            return jsonify({'hata': 'Bitiş saat formatı HH:MM olmalıdır'}), 400
        fark_dk = int((bit - bas).total_seconds() / 60)
        if fark_dk < 0:
            fark_dk += 1440  # Gece yarısı geçen vardiyalar için

    conn = get_db()
    c = conn.cursor()
    c.execute('''
        INSERT INTO vardiyalar (tarih, vardiya_turu, robot_no, operator_adi, baslangic_saati, bitis_saati, toplam_sure_dk, notlar, bolum, lokasyon)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        data['tarih'],
        data['vardiya_turu'],
        data['robot_no'],
        data['operator_adi'],
        bas_str,
        bit_str,
        fark_dk,
        data.get('notlar', ''),
        data.get('bolum', 'kaynak'),
        (data.get('lokasyon') or request.args.get('lokasyon') or 'TK2')
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

    # Otomatik molalar (çay/yemek) — saati gelmişse bu vardiyaya işle (idempotent)
    _otomatik_mola_uygula(c, vardiya)
    conn.commit()

    # Auto sayaç kayıtlarının ok_adet'ini sensör pulse sayımıyla tazele
    _uretim_sayac_senkron(conn, vid)

    uretim = c.execute('SELECT * FROM uretim_kayitlari WHERE vardiya_id = ?', (vid,)).fetchall()
    duruslar = c.execute('SELECT * FROM duruslar WHERE vardiya_id = ?', (vid,)).fetchall()

    # 10dk duruş uyarısı: son AKTİVİTE'den beri geçen dakika. İki kaynak:
    # ESP32 pulse (pilot.db) VE test cihazı başarılı testleri (test_sonuclari) —
    # test cihazıyla çalışan hatta pulse düşmez, yalnız pilot'a bakmak YANLIŞ duruş sayar.
    # En yakın (en küçük dk) aktivite geçerli.
    vbolum = (vardiya['bolum'] if 'bolum' in vardiya.keys() else None) or 'kaynak'
    # pres: makine üretim kaydında → o an çalışılan makineyi son auto kayıttan al
    # (duruş uyarısı doğru cihazın pulse'ına bakmalı; makine yoksa 0 = hat adı).
    _pres_ist = 0
    if vbolum == 'pres':
        _pr = c.execute(
            "SELECT istasyon FROM uretim_kayitlari WHERE vardiya_id=? "
            "AND sayac_baslangic_ts IS NOT NULL ORDER BY id DESC LIMIT 1", (vid,)).fetchone()
        _pres_ist = int((_pr['istasyon'] if _pr else 0) or 0)
    son_pulse_dk = _pilot_son_pulse_dk(vbolum, vardiya['robot_no'], istasyon=_pres_ist)
    test_dk = _test_son_aktivite_dk(conn, vid)
    if test_dk is not None:
        son_pulse_dk = test_dk if son_pulse_dk is None else min(son_pulse_dk, test_dk)
    conn.close()

    return jsonify({
        'vardiya': dict(vardiya),
        'uretim': [dict(r) for r in uretim],
        'duruslar': [dict(r) for r in duruslar],
        'son_pulse_dk': son_pulse_dk,
        # Bu cihazda sensör otomatik sayım destekli mi (allowlist) — mobil UI buna göre
        # sensör rozeti/ipucu/sıfırla butonu gösterir
        'sayac_destekli': _sayac_destekli_bolum(vbolum, vardiya['robot_no']),
    })


@app.route('/api/vardiya/<int:vid>/sayac_log', methods=['GET'])
def vardiya_sayac_log(vid):
    """Operatör mobil 'Sinyal Geçmişi' — vardiyanın sayaç olay logu (en yeni üstte, son 100).
    İki kaynak birleşik: ESP32 pulse (pilot.db sayac_olaylari, vardiya günü + cihaz)
    ve test cihazı kayıtları (test_sonuclari, referans başlangıcından beri)."""
    conn = get_db()
    v = conn.execute(
        "SELECT robot_no, COALESCE(bolum,'kaynak') as bolum, tarih FROM vardiyalar WHERE id=?",
        (vid,)
    ).fetchone()
    if not v:
        return jsonify({'hata': 'Vardiya bulunamadı'}), 404
    loglar = []

    # ── ESP32 pulse'ları (hat → cihaz eşlemesi: LF-LFP→YF1) ──
    robot = HAT_SAYAC_CIHAZI.get(v['robot_no'], v['robot_no'])
    pilot_db = _pilot_db_yolu()
    if os.path.exists(pilot_db):
        try:
            import sqlite3 as _sq
            pc = _sq.connect(pilot_db, timeout=5.0)
            pc.row_factory = _sq.Row
            try:
                pc.execute('PRAGMA busy_timeout=5000')
                for r in pc.execute(
                    "SELECT ts, istasyon FROM sayac_olaylari "
                    "WHERE bolum=? AND robot_no=? AND date(ts)=? ORDER BY ts DESC LIMIT 100",
                    (v['bolum'], robot, v['tarih'])
                ).fetchall():
                    loglar.append({
                        'ts': r['ts'], 'kaynak': 'sensor',
                        'detay': (f"İst.{r['istasyon']}" if (r['istasyon'] or 0) > 0 else 'Sayım'),
                    })
            finally:
                pc.close()
        except Exception as e:
            print(f"[sayac_log] pilot okuma hatası: {e}")

    # ── Test cihazı kayıtları (başarısızlar da listelenir — operatör tam geçmişi görsün) ──
    try:
        for u in conn.execute(
            "SELECT test_cihaz_id, referans_kodu, sayac_baslangic_ts FROM uretim_kayitlari "
            "WHERE vardiya_id=? AND test_cihaz_id IS NOT NULL", (vid,)
        ).fetchall():
            norm = (u['referans_kodu'] or '').strip().upper().replace(' ', '')
            bas_epoch = 0
            if u['sayac_baslangic_ts']:
                try:
                    bas_epoch = int(datetime.strptime(u['sayac_baslangic_ts'], '%Y-%m-%d %H:%M:%S').timestamp())
                except Exception:
                    pass
            for t in conn.execute(
                "SELECT ts_epoch, sonuc, basarili, cihaz_adi FROM test_sonuclari "
                "WHERE cihaz_id=? AND ts_epoch>=? AND (?='' OR referans_norm=?) "
                "ORDER BY ts_epoch DESC LIMIT 100",
                (u['test_cihaz_id'], bas_epoch, norm, norm)
            ).fetchall():
                loglar.append({
                    'ts': datetime.fromtimestamp(t['ts_epoch']).strftime('%Y-%m-%d %H:%M:%S'),
                    'kaynak': 'test',
                    'basarili': int(t['basarili'] or 0),
                    'detay': (t['sonuc'] or ('BAŞARILI' if t['basarili'] else 'BAŞARISIZ')),
                })
    except Exception as e:
        print(f"[sayac_log] test okuma hatası: {e}")

    loglar.sort(key=lambda x: x['ts'], reverse=True)
    return jsonify({'robot_no': v['robot_no'], 'loglar': loglar[:100]})


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
    """Bugune ait tum vardiylari getirir (operatorn secim ekrani icin). ?bolum= ve ?lokasyon= ile filtrelenebilir."""
    from datetime import date as _date
    bugun = _date.today().isoformat()
    bolum = request.args.get('bolum')
    lokasyon = request.args.get('lokasyon')
    q = 'SELECT * FROM vardiyalar WHERE tarih = ?'
    params = [bugun]
    if bolum:
        q += " AND COALESCE(bolum, 'kaynak') = ?"
        params.append(bolum)
    if lokasyon:
        q += " AND COALESCE(lokasyon, 'TK2') = ?"
        params.append(lokasyon)
    q += ' ORDER BY baslangic_saati DESC'
    conn = get_db()
    rows = conn.execute(q, params).fetchall()
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
    # Kapanmadan önce: vardiya boyunca geçen otomatik molaları kesinleştir (rapor tam olsun)
    _otomatik_mola_uygula(c, vardiya)
    # Saat parse KORUMALI: DB'de bozuk baslangic_saati varsa (eski/elle düzenlenmiş kayıt)
    # veya body'de bozuk bitis gelirse vardiya SONSUZA dek kapatılamaz hale gelmesin —
    # 'HH:MM:SS' toleransı + hata durumunda 400 JSON.
    try:
        bas = datetime.strptime(str(vardiya['baslangic_saati'])[:5], '%H:%M')
        bit = datetime.strptime(str(bitis)[:5], '%H:%M')
    except (ValueError, TypeError):
        conn.close()
        return jsonify({'hata': f"Saat formatı geçersiz (başlangıç='{vardiya['baslangic_saati']}', bitiş='{bitis}') — HH:MM olmalı"}), 400
    fark_dk = int((bit - bas).total_seconds() / 60)
    if fark_dk < 0:
        fark_dk += 1440
    # Auto sayaç kayıtlarını KAPANIŞ DEĞERİNDE DONDUR: son sensör/test sayımını yaz,
    # otomatik modu kapat. Yoksa kayıt sayac_otomatik=1 kaldığı için sonraki okumalar
    # (mobil GET/OEE) vardiya KAPANDIKTAN sonraki pulse'ları da bu kayda yazar —
    # yeni vardiyayla ÇİFT SAYIM olur.
    _uretim_sayac_senkron(conn, vid)
    c.execute("UPDATE uretim_kayitlari SET sayac_otomatik=0 WHERE vardiya_id=? AND sayac_otomatik=1", (vid,))
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


@app.route('/api/vardiya/<int:vid>/hat', methods=['PATCH'])
@operator_required
def vardiya_hat_degistir(vid):
    """Açık vardiyanın hattını (robot_no) değiştir — operatör masa değiştirince
    vardiyayı kapatmadan geçtiği masanın sayacına bağlanır.

    Eski masadaki AKTİF sayaç kayıtları o ana kadarki değerde DONDURULUR (yeni
    masanın pulse'ları eski referansa karışmasın); operatör yeni masada '+ Ekle'
    ile sayımı sıfırdan başlatır. Bölüm değişmez (montaj içinde masa değişimi).
    Body: { robot_no }
    """
    ok, hata = _vardiya_sahibi_kontrol(vid)
    if not ok:
        return jsonify({'hata': hata}), 403
    data = request.get_json() or {}
    yeni_robot = (data.get('robot_no') or '').strip()
    if not yeni_robot:
        return jsonify({'hata': 'Hat (robot_no) zorunludur'}), 400

    conn = get_db()
    try:
        c = conn.cursor()
        v = c.execute(
            "SELECT COALESCE(bolum,'kaynak') as bolum, robot_no, durum FROM vardiyalar WHERE id=?",
            (vid,)
        ).fetchone()
        if not v:
            return jsonify({'hata': 'Vardiya bulunamadı'}), 404
        if v['durum'] == 'kapali':
            return jsonify({'hata': 'Kapalı vardiyanın hattı değiştirilemez'}), 400

        eski_robot = v['robot_no']
        if yeni_robot == eski_robot:
            return jsonify({'basarili': True, 'degisti': False, 'robot_no': yeni_robot,
                            'sayac_destekli': (yeni_robot in SAYAC_AUTO_CIHAZLAR)})

        # Eski masadaki aktif sayaç kayıtlarını O ANA kadar dondur (eski robot_no ile say).
        # uretim_ekle'deki referans-değişimi dondurma mantığıyla aynı.
        simdi = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        aktif = c.execute(
            "SELECT id, istasyon, sayac_baslangic_ts, COALESCE(bukum_operasyon,1) AS bukum_operasyon "
            "FROM uretim_kayitlari "
            "WHERE vardiya_id=? AND sayac_otomatik=1 AND sayac_baslangic_ts IS NOT NULL",
            (vid,)
        ).fetchall()
        for o in aktif:
            fcnt = _pilot_pulse_say(v['bolum'], eski_robot, o['istasyon'] or 0,
                                    o['sayac_baslangic_ts'], biti_ts=simdi)
            # Dondurma da bölünmeli — yoksa kayıt kapanınca adet ham pulse'a fırlar
            fcnt = _bukum_bol(fcnt, o['bukum_operasyon'])
            if fcnt is not None:
                c.execute("UPDATE uretim_kayitlari SET ok_adet=?, sayac_otomatik=0 WHERE id=?",
                          (fcnt, o['id']))
            else:
                # pilot.db okunamadı — ok_adet'i SIFIRLAMA; mevcut değerle manuel'e al
                c.execute("UPDATE uretim_kayitlari SET sayac_otomatik=0 WHERE id=?", (o['id'],))

        c.execute("UPDATE vardiyalar SET robot_no=? WHERE id=?", (yeni_robot, vid))
        conn.commit()
        return jsonify({'basarili': True, 'degisti': True, 'robot_no': yeni_robot,
                        'eski_robot': eski_robot, 'donduruldu': len(aktif),
                        'sayac_destekli': (yeni_robot in SAYAC_AUTO_CIHAZLAR)})
    finally:
        conn.close()


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

    # Saat FORMAT DOĞRULAMA: bozuk değer DB'ye yazılırsa vardiya_kapat her denemede
    # patlar ve vardiya kapatılamaz hale gelir — bozuk formatı 400 ile reddet.
    # (Boş bitis_saati meşru: açık vardiya. 'HH:MM:SS' ilk 5 karaktere kırpılır.)
    def _saat_ok(s):
        if s is None or str(s).strip() == '':
            return ''
        try:
            datetime.strptime(str(s).strip()[:5], '%H:%M')
            return str(s).strip()[:5]
        except ValueError:
            return None
    bas = _saat_ok(bas)
    bit = _saat_ok(bit)
    if bas is None or bit is None or bas == '':
        conn.close()
        return jsonify({'hata': 'Saat formatı HH:MM olmalıdır (örn: 07:30)'}), 400

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
    try:
        vardiya_id_int = int(data['vardiya_id'])
    except (ValueError, TypeError):
        return jsonify({'hata': f"vardiya_id sayısal olmalı: {data.get('vardiya_id')!r}"}), 400

    # Operatör yetki kontrolü: header'da operator+pin geldiyse vardiya sahibi mi?
    ok, hata = _vardiya_sahibi_kontrol(vardiya_id_int)
    if not ok:
        return jsonify({'hata': hata}), 403

    conn = None
    try:
        conn = get_db()
        c = conn.cursor()

        # Vardiyanın bölümünü + robot_no'sunu + lokasyonunu öğren
        v_row = c.execute(
            "SELECT COALESCE(bolum, 'kaynak') as bolum, robot_no, COALESCE(lokasyon, 'TK2') as lokasyon FROM vardiyalar WHERE id = ?",
            (data['vardiya_id'],)
        ).fetchone()
        vardiya_bolum = v_row['bolum'] if v_row else 'kaynak'
        vardiya_robot = v_row['robot_no'] if v_row else ''
        vardiya_lokasyon = v_row['lokasyon'] if v_row else 'TK2'

        # Birden fazla kayıt gelebilir
        satirlar = data.get('satirlar', [data])
        tek_satir = (len(satirlar) == 1)
        simdi = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        eklenen = 0
        for satir in satirlar:
            ref = (satir.get('referans_kodu') or data.get('referans_kodu') or '').strip()
            ct_in = float(satir.get('cycle_time_sn', 0) or 0)
            # cycle_time gönderilmediyse ya da 0 ise referanslardan otomatik çek.
            # LOKASYON filtresi ŞART: composite UNIQUE sonrası aynı kod TK1+TK2'de
            # iki satır olabilir — vardiyanın tesisinin cycle'ı çekilmeli.
            # BÖLÜM TERCİHİ: aynı kod birden fazla bölümde farklı cycle ile olabilir —
            # önce vardiyanın bölümündeki satır; yoksa (eski davranış) herhangi biri.
            if ct_in <= 0 and ref:
                ref_row = c.execute(
                    "SELECT hedef_cycle_time_sn FROM referans_listesi "
                    "WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ','')) AND COALESCE(lokasyon,'TK2')=? "
                    "ORDER BY (COALESCE(bolum,'kaynak') = ?) DESC LIMIT 1",
                    (ref, vardiya_lokasyon, vardiya_bolum)
                ).fetchone()
                if ref_row:
                    ct_in = float(ref_row['hedef_cycle_time_sn'] or 0)

            ist_val = int(satir.get('istasyon', data.get('istasyon', 0)) or 0)
            ok_val  = int(satir.get('ok_adet', 0) or 0)

            # Test cihazı (Cofle/wicow) seçildiyse sayaç kaynağı o cihaz olur (ESP32 yerine)
            tc_raw = satir.get('test_cihaz_id', data.get('test_cihaz_id'))
            try:
                tc_id = int(tc_raw) if tc_raw not in (None, '', 0, '0') else None
            except (ValueError, TypeError):
                tc_id = None

            # Sayaç otomatik mod: tek satır + OK girilmemiş (0) + (cihaz allowlist'te VEYA
            # test cihazı seçilmiş). Operatör referansa YENİ başlıyor → sayaç o andan sıfırlanır.
            # (Toplu/geçmiş giriş, OK elle yazılmış veya desteklenmeyen kaynak → manuel mod.)
            # pres: allowlist kontrolü SATIRIN makinesine (istasyon) göre — makine
            # seçilmemişse otomatik sayaç YOK (manuel giriş; 0-ezme tuzağını önler).
            auto = 1 if (tek_satir and ok_val == 0
                         and (_sayac_destekli(vardiya_bolum, vardiya_robot, ist_val) or tc_id)) else 0

            if auto and tc_id:
                # Aynı test cihazında önceki auto kayıt(lar)ı dondur (referans değişti)
                onceki = c.execute(
                    "SELECT id, sayac_baslangic_ts, referans_kodu FROM uretim_kayitlari "
                    "WHERE vardiya_id=? AND test_cihaz_id=? AND sayac_otomatik=1",
                    (data['vardiya_id'], tc_id)
                ).fetchall()
                for o in onceki:
                    fcnt = _test_basari_say(conn, tc_id, o['referans_kodu'],
                                            o['sayac_baslangic_ts'], biti_ts=simdi)
                    if fcnt is not None:
                        c.execute("UPDATE uretim_kayitlari SET ok_adet=?, sayac_otomatik=0 WHERE id=?",
                                  (fcnt, o['id']))
                    else:
                        c.execute("UPDATE uretim_kayitlari SET sayac_otomatik=0 WHERE id=?", (o['id'],))
            elif auto:
                # Aynı vardiya+istasyon önceki ESP32 auto kayıt(lar)ı dondur (test-cihazı hariç)
                onceki = c.execute(
                    "SELECT id, istasyon, sayac_baslangic_ts, COALESCE(bukum_operasyon,1) AS bukum_operasyon "
                    "FROM uretim_kayitlari "
                    "WHERE vardiya_id=? AND istasyon=? AND sayac_otomatik=1 AND test_cihaz_id IS NULL",
                    (data['vardiya_id'], ist_val)
                ).fetchall()
                for o in onceki:
                    fcnt = _pilot_pulse_say(vardiya_bolum, vardiya_robot, ist_val,
                                            o['sayac_baslangic_ts'], biti_ts=simdi)
                    # Dondurma da bölünmeli (bkz. _bukum_bol)
                    fcnt = _bukum_bol(fcnt, o['bukum_operasyon'])
                    if fcnt is not None:
                        c.execute("UPDATE uretim_kayitlari SET ok_adet=?, sayac_otomatik=0 WHERE id=?",
                                  (fcnt, o['id']))
                    else:
                        # pilot.db okunamadi — ok_adet'i SIFIRLAMA; mevcut degerle dondur (manuel'e al)
                        c.execute("UPDATE uretim_kayitlari SET sayac_otomatik=0 WHERE id=?", (o['id'],))

            # Referansı listeye otomatik ekle — VARDIYANIN bölümü VE lokasyonuyla etiketle
            # (montaj operatörü tanımsız bir kod girince montaj'a düşsün, kaynak'a değil;
            #  TK1 operatörünün girdiği kod TK1'e düşsün, TK2'ye karışmasın)
            # NORMALIZE kontrol: UNIQUE ham kod üzerinde — operatör '94. LTK. 341/10' gibi
            # boşluklu/farklı yazarsa INSERT OR IGNORE yeni VARYANT satır açıyordu; bu
            # duplike satırlar JOIN'lerde iş emirlerini çoğaltıyordu (fan-out). Normalize
            # eş (aynı lokasyonda) varsa yeni satır AÇMA.
            # SIRA: bu blok üretim INSERT'inden ÖNCE olmalı — _bukum_op_coz büküm
            # sayısını referans satırına yazıyor, satır henüz yoksa değer kaybolurdu.
            ref_var = c.execute(
                "SELECT 1 FROM referans_listesi "
                "WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ','')) "
                "AND COALESCE(lokasyon,'TK2')=? LIMIT 1",
                (ref, vardiya_lokasyon)
            ).fetchone()
            if not ref_var:
                c.execute(
                    'INSERT OR IGNORE INTO referans_listesi (referans_kodu, bolum, lokasyon) VALUES (?, ?, ?)',
                    (ref, vardiya_bolum, vardiya_lokasyon)
                )

            # BÜKÜM OPERASYON SAYISI — yalnız pres (abkant) için anlamlı.
            # İstemci göndermediyse referansın hatırlanan değerine düşer; o da yoksa 1
            # (bölme yok = eski davranış). Kayda KOPYALANIR: referansın varsayılanı
            # sonradan değişse bile bu kaydın adedi retroaktif kaymasın.
            bop = _bukum_op_coz(c, satir.get('bukum_operasyon', data.get('bukum_operasyon')),
                                ref, vardiya_bolum, vardiya_lokasyon)

            c.execute('''
                INSERT INTO uretim_kayitlari (vardiya_id, referans_kodu, ok_adet, nok_adet, tamir_adet, hedef_adet, cycle_time_sn, istasyon, launch_adet, aciklama, sayac_baslangic_ts, sayac_otomatik, test_cihaz_id, bukum_operasyon)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                data['vardiya_id'],
                ref,
                ok_val,
                int(satir.get('nok_adet', 0) or 0),
                int(satir.get('tamir_adet', 0) or 0),
                int(satir.get('hedef_adet', 0) or 0),
                ct_in,
                ist_val,
                int(satir.get('launch_adet', data.get('launch_adet', 0)) or 0),
                (satir.get('aciklama') or data.get('aciklama') or '').strip(),
                simdi if auto else None,
                auto,
                tc_id,
                bop
            ))
            eklenen += 1

        conn.commit()
        return jsonify({'basarili': True, 'eklenen': eklenen}), 201
    except Exception as e:
        print(f"HATA (uretim_ekle): {traceback.format_exc()}")
        return jsonify({'hata': f'Üretim eklenemedi: {str(e)}'}), 500
    finally:
        if conn:
            conn.close()


@app.route('/api/uretim/<int:uid>/sayac_sifirla', methods=['POST'])
@operator_required
def uretim_sayac_sifirla(uid):
    """Operatör emniyet reseti — bu üretim kaydının sayacını ŞİMDİDEN sıfırlar.
    Deneme/test sinyalleri karışıklık yaratırsa operatör sayacı 0'lar.
    baseline=now → bu andan sonraki pulse'lar sayılır; otomatik mod yeniden aktif."""
    vid = _uretim_vardiya_bul(uid)
    if vid:
        ok, hata = _vardiya_sahibi_kontrol(vid)
        if not ok:
            return jsonify({'hata': hata}), 403
    conn = get_db()
    try:
        simdi = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        cur = conn.execute(
            "UPDATE uretim_kayitlari SET sayac_baslangic_ts=?, ok_adet=0, sayac_otomatik=1 WHERE id=?",
            (simdi, uid)
        )
        conn.commit()
        if cur.rowcount == 0:
            return jsonify({'hata': 'Kayit bulunamadi'}), 404
        return jsonify({'basarili': True, 'sayac_baslangic_ts': simdi})
    except Exception as e:
        # DB kilidi/hata: mobil JSON bekler — HTML 500 yerine anlaşılır hata
        return jsonify({'hata': f'Sayaç sıfırlanamadı: {e}'}), 500
    finally:
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
        mevcut = c.execute(
            'SELECT id, vardiya_id, istasyon, sayac_otomatik, sayac_baslangic_ts '
            'FROM uretim_kayitlari WHERE id = ?', (uid,)
        ).fetchone()
        if not mevcut:
            return jsonify({'hata': 'Kayit bulunamadi'}), 404

        ref = (data.get('referans_kodu') or '').strip()
        ct = float(data.get('cycle_time_sn', 0) or 0)
        # Gelen CT 0 veya gönderilmediyse referanslardan çek — vardiyanın LOKASYONUYLA
        # (composite UNIQUE sonrası aynı kod iki tesiste olabilir; yanlış tesisin ct'si çekilmesin)
        if ct <= 0 and ref:
            v_lok_row = c.execute(
                "SELECT COALESCE(lokasyon,'TK2') as lok FROM vardiyalar WHERE id = ?",
                (mevcut['vardiya_id'],)
            ).fetchone()
            v_lok = v_lok_row['lok'] if v_lok_row else 'TK2'
            ref_row = c.execute(
                "SELECT hedef_cycle_time_sn FROM referans_listesi "
                "WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ','')) AND COALESCE(lokasyon,'TK2')=?",
                (ref, v_lok)
            ).fetchone()
            if ref_row:
                ct = float(ref_row['hedef_cycle_time_sn'] or 0)

        submitted_ok = int(data.get('ok_adet', 0) or 0)
        # Sayaç otomatik modu: mobil 'ok_elle_degisti' bayrağı gönderirse (operatör
        # OK alanına dokundu) otomatik kapatılır (değer dondurulur). Sadece NOK/açıklama
        # değişip OK'e dokunulmadıysa bayrak gelmez → auto devam eder.
        yeni_otomatik = mevcut['sayac_otomatik'] or 0
        if data.get('ok_elle_degisti'):
            yeni_otomatik = 0

        # BÜKÜM OPERASYON SAYISI — operatör kaydı düzenlerken de değiştirebilir.
        # Vardiyanın bölüm/lokasyonuyla çözülür; pres dışında her zaman 1 döner.
        v_bl = c.execute(
            "SELECT COALESCE(bolum,'kaynak') AS bolum, COALESCE(lokasyon,'TK2') AS lokasyon "
            "FROM vardiyalar WHERE id=?", (mevcut['vardiya_id'],)
        ).fetchone()
        bop = _bukum_op_coz(c, data.get('bukum_operasyon'), ref,
                            v_bl['bolum'] if v_bl else 'kaynak',
                            v_bl['lokasyon'] if v_bl else 'TK2')

        c.execute('''
            UPDATE uretim_kayitlari
            SET referans_kodu=?, ok_adet=?, nok_adet=?, tamir_adet=?, hedef_adet=?, cycle_time_sn=?, istasyon=?, launch_adet=?, aciklama=?, sayac_otomatik=?, bukum_operasyon=?
            WHERE id=?
        ''', (
            ref,
            submitted_ok,
            int(data.get('nok_adet', 0) or 0),
            int(data.get('tamir_adet', 0) or 0),
            int(data.get('hedef_adet', 0) or 0),
            ct,
            int(data.get('istasyon', 0) or 0),
            int(data.get('launch_adet', 0) or 0),
            (data.get('aciklama') or '').strip(),
            yeni_otomatik,
            bop,
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
    # Referans eşleşmesi kaydın VARDİYASININ lokasyonuna kapalı — composite UNIQUE
    # sonrası aynı kod TK1+TK2'de olabilir; yanlış tesisin cycle'ı yazılmasın.
    cur = conn.execute("""
        UPDATE uretim_kayitlari
        SET cycle_time_sn = (
            SELECT r.hedef_cycle_time_sn
            FROM referans_listesi r
            WHERE UPPER(REPLACE(r.referans_kodu,' ','')) = UPPER(REPLACE(uretim_kayitlari.referans_kodu,' ',''))
              AND COALESCE(r.lokasyon,'TK2') = (
                    SELECT COALESCE(v.lokasyon,'TK2') FROM vardiyalar v WHERE v.id = uretim_kayitlari.vardiya_id)
              AND r.hedef_cycle_time_sn > 0
            LIMIT 1
        )
        WHERE (cycle_time_sn IS NULL OR cycle_time_sn = 0)
          AND EXISTS (
            SELECT 1 FROM referans_listesi r
            WHERE UPPER(REPLACE(r.referans_kodu,' ','')) = UPPER(REPLACE(uretim_kayitlari.referans_kodu,' ',''))
              AND COALESCE(r.lokasyon,'TK2') = (
                    SELECT COALESCE(v.lokasyon,'TK2') FROM vardiyalar v WHERE v.id = uretim_kayitlari.vardiya_id)
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
    try:
        durus_vid_int = int(data['vardiya_id'])
    except (ValueError, TypeError):
        return jsonify({'hata': f"vardiya_id sayısal olmalı: {data.get('vardiya_id')!r}"}), 400

    # Operatör yetki kontrolü: vardiya sahibi mi?
    ok, hata = _vardiya_sahibi_kontrol(durus_vid_int)
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
    data = request.get_json(silent=True) or {}
    # sure_dk güvenli parse (boş input '' → ValueError → HTML 500 olmasın; durus_ekle kalıbı)
    try:
        sure_dk_val = int(data.get('sure_dk', 0) or 0)
    except (ValueError, TypeError):
        return jsonify({'hata': f"Süre sayısal olmalı: {data.get('sure_dk')!r}"}), 400
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
        sure_dk_val,
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
    # OEE'nin güncel sensör sayımını kullanması için auto kayıtları tazele.
    # NOT: get_db() g-paylaşımlı — BURADA KAPATMA (teardown kapatır); erken close
    # sonrası get_db kullanan kod eklenirse 'closed database' 500'e döner.
    try:
        _uretim_sayac_senkron(get_db(), vid)
    except Exception:
        pass
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
    lokasyon = request.args.get('lokasyon')   # TK1/TK2 montaj OEE'si karışmasın

    # Varsayılan: bugün
    if not tarih_bas and not tarih_bit:
        bugun = date.today().isoformat()
        tarih_bas = bugun
        tarih_bit = bugun

    # Aralık bugünü kapsıyorsa auto sayaçları tazele — OEE özeti canlı adetlerle hesaplansın
    _bugun_o = date.today().isoformat()
    if (not tarih_bas or tarih_bas <= _bugun_o) and (not tarih_bit or tarih_bit >= _bugun_o):
        _acik_auto_vardiyalari_senkronla(get_db(), bolum or None, lokasyon or None)

    sonuc = hesapla_oee_ozet(tarih_bas, tarih_bit, robot, bolum, lokasyon)
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
    lokasyon = request.args.get('lokasyon', '')
    conn = get_db()
    base = ("SELECT referans_kodu, aciklama, hedef_cycle_time_sn, kaynak_suresi_sn, soktak_suresi_sn, "
            "sure_teyit, sure_teyit_tarihi, COALESCE(bukum_operasyon,1) AS bukum_operasyon "
            "FROM referans_listesi")
    sql = base + " WHERE REPLACE(referans_kodu, ' ', '') LIKE REPLACE(?, ' ', '')"
    params = [f'%{q}%']
    if bolum:
        sql += " AND bolum = ?"
        params.append(bolum)
    if lokasyon:
        sql += " AND COALESCE(lokasyon, 'TK2') = ?"
        params.append(lokasyon)
    sql += " ORDER BY referans_kodu"
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/api/referanslar', methods=['POST'])
@panel_gerekli(izin='referanslar')
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
    if bolum not in GECERLI_BOLUMLER:
        bolum = 'kaynak'
    lokasyon = (data.get('lokasyon') or request.args.get('lokasyon') or 'TK2').strip() or 'TK2'

    conn = get_db()
    try:
        # BÖLÜM+LOKASYONA KAPALI UPSERT: aynı (referans_kodu, bolum, lokasyon) varsa GÜNCELLE,
        # yoksa EKLE. (INSERT OR REPLACE kullanılmıyor → sure_teyit/kaynak/söktak gibi diğer
        # kolonlar korunur; diğer fabrikanın/bölümün aynı kodlu satırına ASLA dokunulmaz —
        # aynı kod birden fazla bölümde farklı cycle ile yaşayabilir.)
        mevcut = conn.execute(
            "SELECT id FROM referans_listesi "
            "WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ','')) "
            "AND COALESCE(bolum,'kaynak')=? AND COALESCE(lokasyon,'TK2')=?",
            (ref, bolum, lokasyon)
        ).fetchone()
        if mevcut:
            conn.execute(
                "UPDATE referans_listesi SET aciklama=?, hedef_cycle_time_sn=? WHERE id=?",
                (desc, ct, mevcut['id'])
            )
        else:
            conn.execute(
                "INSERT INTO referans_listesi (referans_kodu, aciklama, hedef_cycle_time_sn, bolum, lokasyon) VALUES (?, ?, ?, ?, ?)",
                (ref, desc, ct, bolum, lokasyon)
            )

        # Geriye dönük: aynı referansın geçmiş üretim kayıtlarının cycle'ını güncelle —
        # AMA sadece bu lokasyonun + bu bölümün vardiyalarına ait olanları
        # (diğer fabrikayı/bölümü ezme — aynı kodun bölüm başına cycle'ı farklı olabilir).
        if ct > 0:
            conn.execute(
                "UPDATE uretim_kayitlari SET cycle_time_sn = ? "
                "WHERE UPPER(REPLACE(referans_kodu, ' ', '')) = UPPER(REPLACE(?, ' ', '')) "
                "AND vardiya_id IN (SELECT id FROM vardiyalar WHERE COALESCE(lokasyon,'TK2')=? AND COALESCE(bolum,'kaynak')=?)",
                (ct, ref, lokasyon, bolum)
            )

        conn.commit()
    except Exception as e:
        conn.close()
        return jsonify({'hata': str(e)}), 400
    conn.close()

    # Excel auto-sync: SADECE TK2 (uretim_verileri.xlsx). TK1 referansları cycle time
    # kullanmaz ve ayrı Excel'dedir → TK1'de export atlanır (TK1 verisi TK2 dosyasına sızmaz).
    try:
        export_referans_cycle_times(bolum=bolum, lokasyon=lokasyon)
    except Exception as e:
        print(f'[referans_ekle] Excel auto-sync hatası ({bolum}/{lokasyon}): {e}')

    return jsonify({'basarili': True}), 201


# ── Referans KODU değiştirme (rename) — TÜM ilgili tablolarda ──
# Referans kodu 11 tablo.kolonda serbest string (FK yok); rename her yeri günceller.
# as400_teyit_log (audit) DOKUNULMAZ.
_REFERANS_KOD_TABLOLARI = [
    ('referans_listesi', 'referans_kodu'), ('uretim_kayitlari', 'referans_kodu'),
    ('referans_takip', 'referans_kodu'), ('referans_robot_uyumu', 'referans_kodu'),
    ('robot_programlari', 'referans_kodu'), ('robot_programlari_eski', 'referans_kodu'),
    ('robot_is_atamalari', 'referans_kodu'), ('fikstur_adresleri', 'referans_kodu'),
    ('fikstur_raf', 'referans_kodu'), ('test_sonuclari', 'referans_kodu'),
    ('test_sonuclari', 'referans_norm'),
]


def _guvenli_kod_replace(s, eski, yeni):
    """s içindeki 'eski' kodunu 'yeni' ile değiştirir AMA 'eski' daha büyük bir
    kodun parçasıysa (öncesi/sonrası alfasayısal veya - / ) o eşleşmeyi KORUR.
    Örn: eski='10.300.4081' → '10.300.4081W' DOKUNULMAZ (W devam), ama
    '10.300.4081 1. op' değişir (boşluk sınırı). Op etiketi hep korunur."""
    if not s or eski not in s:
        return s
    out = []
    i = 0
    n = len(eski)
    while True:
        j = s.find(eski, i)
        if j < 0:
            out.append(s[i:])
            break
        son = j + n
        onceki = s[j - 1] if j > 0 else ''
        sonraki = s[son] if son < len(s) else ''
        # DIKKAT: `sonraki in '-/'` KULLANMA — bos string ('' in '-/') True doner,
        # string SONUNDA biten tam eslesmeler asla degistirilemezdi (2026-07-22 bug).
        if onceki.isalnum() or sonraki.isalnum() or sonraki in ('-', '/'):
            out.append(s[i:son])          # daha büyük kodun parçası → dokunma
        else:
            out.append(s[i:j])
            out.append(yeni)
        i = son
    return ''.join(out)


def _referans_kodu_rename(conn, eski, yeni, uygula):
    """Tüm ilgili tablolarda kodu değiştirir. uygula=False → sadece önizleme (say + örnek).
    MASTER BİRLEŞTİRME (2026-07-20, 4081 dersi): referans_listesi'nde yeni kod aynı
    bölüm+lokasyonda ZATEN varsa (yazım hatası → doğru koda düzeltme senaryosu),
    hatalı master satırı SİLİNİR, hedef kayıt korunur (süreleri elle taşınabilir).
    Döner: (ozet {tablo.kolon: adet}, ornekler [{tablo, eski, yeni}])."""
    eski = (eski or '').strip()
    yeni = (yeni or '').strip()
    ozet = {}
    ornekler = []
    for t, k in _REFERANS_KOD_TABLOLARI:
        try:
            rows = conn.execute(f"SELECT rowid, {k} FROM {t} WHERE {k} LIKE ?", (f'%{eski}%',)).fetchall()
        except Exception:
            continue   # tablo yoksa atla
        say = 0
        for r in rows:
            rid, deger = r[0], r[1]
            yd = _guvenli_kod_replace(deger, eski, yeni)
            if yd != deger:
                if t == 'referans_listesi':
                    cakisan = conn.execute(
                        "SELECT rowid FROM referans_listesi WHERE rowid != ? "
                        "AND UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ','')) "
                        "AND COALESCE(bolum,'kaynak')=(SELECT COALESCE(bolum,'kaynak') FROM referans_listesi WHERE rowid=?) "
                        "AND COALESCE(lokasyon,'TK2')=(SELECT COALESCE(lokasyon,'TK2') FROM referans_listesi WHERE rowid=?)",
                        (rid, yd, rid, rid)).fetchone()
                    if cakisan:
                        say += 1
                        if len(ornekler) < 15:
                            ornekler.append({'tablo': t, 'eski': deger, 'yeni': yd + ' (BİRLEŞTİRME: hedef kayıt zaten var, hatalı tanım silinir)'})
                        if uygula:
                            conn.execute("DELETE FROM referans_listesi WHERE rowid=?", (rid,))
                        continue
                say += 1
                if len(ornekler) < 15:
                    ornekler.append({'tablo': t, 'eski': deger, 'yeni': yd})
                if uygula:
                    conn.execute(f"UPDATE {t} SET {k}=? WHERE rowid=?", (yd, rid))
        if say:
            ozet[f'{t}.{k}'] = say
    return ozet, ornekler


@app.route('/api/referans/kod_degistir', methods=['POST'])
@panel_gerekli(izin='referanslar')
def referans_kod_degistir():
    """Bir referans kodunu TÜM ilgili tablolarda değiştirir (rename).
    Body: {eski, yeni, onizleme?:bool}. onizleme=true → uygulamadan say+örnek döndürür."""
    data = request.get_json(silent=True) or {}
    eski = (data.get('eski') or '').strip()
    yeni = (data.get('yeni') or '').strip()
    onizleme = bool(data.get('onizleme'))
    if not eski or not yeni:
        return jsonify({'hata': 'eski ve yeni kod zorunlu'}), 400
    if eski == yeni:
        return jsonify({'hata': 'eski ve yeni kod aynı'}), 400
    conn = get_db()
    try:
        ozet, ornekler = _referans_kodu_rename(conn, eski, yeni, uygula=not onizleme)
        if onizleme:
            return jsonify({'onizleme': True, 'eski': eski, 'yeni': yeni,
                            'toplam': sum(ozet.values()), 'ozet': ozet, 'ornekler': ornekler})
        conn.commit()
        return jsonify({'basarili': True, 'eski': eski, 'yeni': yeni,
                        'toplam': sum(ozet.values()), 'ozet': ozet})
    except Exception as e:
        conn.rollback()
        msg = str(e)
        if 'UNIQUE' in msg.upper():
            return jsonify({'hata': f'Hedef kod aynı bölüm/lokasyonda zaten var: {msg}'}), 409
        return jsonify({'hata': msg}), 400


@app.route('/api/referanslar/eksik', methods=['GET'])
def referans_eksik_listesi():
    """Süresi tanımlanmamış (0 veya NULL) referansları getirir. ?bolum= ve ?lokasyon= ile filtrelenebilir."""
    bolum = request.args.get('bolum')
    lokasyon = request.args.get('lokasyon')
    conn = get_db()
    c = conn.cursor()
    sql = ("SELECT referans_kodu FROM referans_listesi "
           "WHERE (hedef_cycle_time_sn IS NULL OR hedef_cycle_time_sn = 0)")
    params = []
    if bolum:
        sql += " AND COALESCE(bolum, 'kaynak') = ?"
        params.append(bolum)
    if lokasyon:
        sql += " AND COALESCE(lokasyon, 'TK2') = ?"
        params.append(lokasyon)
    sql += " ORDER BY referans_kodu"
    rows = c.execute(sql, params).fetchall()
    conn.close()
    return jsonify([r['referans_kodu'] for r in rows])


@app.route('/api/referanslar/export_excel', methods=['POST'])
@panel_gerekli(izin='referanslar')
def referans_excel_export():
    """Veritabanındaki referans listesini data/uretim_verileri.xlsx'in
    aktif bölüm sayfasına yazar.

    Body: { bolum: 'kaynak' | 'montaj' | 'metal' } — yoksa kaynak.
    """
    # Sayfa adları import_excel.BOLUM_SAYFA'dan (tek kaynak) — yeni bölümler dahil.
    from import_excel import BOLUM_SAYFA as _IE_SAYFA
    BOLUM_SAYFA = {b: v['ref'] for b, v in _IE_SAYFA.items()}
    body = request.get_json(silent=True) or {}
    bolum = body.get('bolum', 'kaynak')
    lokasyon = (body.get('lokasyon') or request.args.get('lokasyon') or 'TK2').strip() or 'TK2'
    if bolum not in BOLUM_SAYFA:
        return jsonify({'hata': f"Geçersiz bölüm: {bolum}"}), 400
    # TK1 referansları cycle time kullanmaz ve ayrı dosyadadır (tek kolon) → TK2 Excel'ine yazma
    if lokasyon == 'TK1':
        return jsonify({'basarili': True, 'mesaj': 'TK1 referansları için cycle time Excel aktarımı yok.'}), 200
    # TEK implementasyon: import_excel.export_referans_cycle_times — kaynak sayfasında
    # 4+1 kolon (Kaynak/Söktak/Toplam/Süre Teyit), montaj/metal'de 2 kolon yazar.
    # (Eski yerel kopya kaynak'ta B kolonuna toplam cycle yazıp Kaynak Süresi'ni eziyordu.)
    try:
        sonuc = export_referans_cycle_times(bolum=bolum, lokasyon=lokasyon)
        if not sonuc.get('basarili'):
            return jsonify({'hata': sonuc.get('hata', 'Excel yazılamadı')}), 500
        return jsonify({'basarili': True, 'mesaj': f"{sonuc.get('yazilan', 0)} referans Excel'e yazıldı."})
    except Exception as e:
        return jsonify({'hata': str(e)}), 500


@app.route('/api/referanslar/<string:kod>', methods=['DELETE'])
@panel_gerekli(izin='referanslar')
def referans_sil(kod):
    """Referansı listeden sil — SADECE aktif lokasyonun satırı (diğer fabrikanın aynı kodlu
    referansına dokunma). lokasyon query/body'den; verilmezse TK2.
    ?bolum= verilirse yalnız o bölümün satırı silinir (aynı kod birden fazla bölümde
    yaşayabilir); verilmezse eski davranış (koddaki tüm TK2 satırları)."""
    lokasyon = (request.args.get('lokasyon') or 'TK2').strip() or 'TK2'
    bolum = (request.args.get('bolum') or '').strip()
    conn = get_db()
    try:
        if bolum:
            conn.execute(
                "DELETE FROM referans_listesi WHERE referans_kodu = ? AND COALESCE(bolum,'kaynak') = ? AND COALESCE(lokasyon,'TK2') = ?",
                (kod, bolum, lokasyon)
            )
        else:
            conn.execute(
                "DELETE FROM referans_listesi WHERE referans_kodu = ? AND COALESCE(lokasyon,'TK2') = ?",
                (kod, lokasyon)
            )
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
    - metal: 300T, 400T, 550T, Şerit Testere (sabit)
    - lazer: tek sabit hat 'Lazer' — makine (Lazer 1..6) vardiyada DEĞİL üretim kaydında
      seçilir (istasyon kolonu 1..6; operatör aynı anda birden fazla makine çalıştırır)
    - pres: tek sabit hat 'Pres Abkant' (2026-07-28) — makine (Abkant 1-3 / Pres 1-5)
      vardiyada DEĞİL üretim kaydında seçilir (istasyon 1..8); operatör gün içinde
      makine değiştirebildiği için lazer modeline geçildi. Eski vardiyaların makine
      adları (Abkant 2, Pres...) listede KALIR → dashboard geçmiş filtreleri çalışsın.
    - montaj/isleme: vardiyalardan distinct robot_no — operatörler girdikçe dinamik büyür
    """
    bolum = request.args.get('bolum', 'kaynak')
    lokasyon = request.args.get('lokasyon', 'TK2')
    # Plastik enjeksiyon (TK1, 2026-07-23): sabit 2 makine + operatörün eklediği ekstralar.
    # TK1 montaj-hat kısayolundan ÖNCE — aksi halde plastik montaj hatlarıyla ezilirdi.
    if bolum == 'plastik':
        conn = get_db()
        rows = conn.execute(
            "SELECT DISTINCT robot_no FROM vardiyalar WHERE COALESCE(bolum,'')='plastik' "
            "AND robot_no IS NOT NULL AND robot_no != '' ORDER BY robot_no").fetchall()
        conn.close()
        # 320T/407T = enjeksiyon; Yapistirma = yapıştırma makinesi (aynı bölüm, TK1).
        taban = ['320T', '407T', 'Yapistirma']
        return jsonify(taban + [r['robot_no'] for r in rows if r['robot_no'] not in taban])
    if lokasyon == 'TK1':
        # TK1 (yan tesis) sabit hatları — montaj mantığı
        return jsonify(['Pull', 'Push-Pull', 'Iveco', 'LF-LFP'])
    # Sabit makine tabanı olan bölümler (liste her zaman tam görünür; DISTINCT ekstraları eklenir)
    # pres BURADA DEĞİL: hattı sabit (aşağıda), makineleri üretim kaydında seçilir.
    SABIT_TABAN = {'plastik': ['320T', '407T', 'Yapistirma']}
    # Dinamik bölümlerde hiç vardiya yokken gösterilecek varsayılan makine/hat adı
    DINAMIK_VARSAYILAN = {'montaj': 'HAT 1', 'isleme': 'Tezgah 1'}
    if bolum == 'metal':
        robotlar = ['300T', '400T', '550T', 'Şerit Testere']
    elif bolum == 'lazer':
        robotlar = ['Lazer']
    elif bolum == 'pres':
        # Sabit hat + GEÇMİŞ vardiyaların makine adları (eski model: makine vardiyadaydı).
        # Mobil bu listeyi kullanmaz (hat gizli/sabit); dashboard geçmiş filtreleri kullanır.
        conn = get_db()
        rows = conn.execute(
            "SELECT DISTINCT robot_no FROM vardiyalar WHERE COALESCE(bolum,'')='pres' "
            "AND robot_no IS NOT NULL AND robot_no != '' ORDER BY robot_no").fetchall()
        conn.close()
        robotlar = [PRES_SABIT_HAT] + [r['robot_no'] for r in rows if r['robot_no'] != PRES_SABIT_HAT]
    elif bolum in SABIT_TABAN or bolum in DINAMIK_VARSAYILAN:
        conn = get_db()
        rows = conn.execute(
            "SELECT DISTINCT robot_no FROM vardiyalar WHERE COALESCE(bolum, 'kaynak') = ? "
            "AND COALESCE(lokasyon, 'TK2') = ? AND robot_no IS NOT NULL AND robot_no != '' ORDER BY robot_no",
            (bolum, lokasyon or 'TK2')
        ).fetchall()
        conn.close()
        kullanilan = [r['robot_no'] for r in rows]
        taban = SABIT_TABAN.get(bolum, [])
        robotlar = taban + [r for r in kullanilan if r not in taban]
        # Hiç kayıt yoksa varsayılanı göster — operatör formundan ilk girişte yenileri otomatik eklenir
        if not robotlar:
            robotlar = [DINAMIK_VARSAYILAN[bolum]]
    else:
        robotlar = [f'ABB{i}' for i in range(1, 10)]
    return jsonify(robotlar)


@app.route('/api/test_cihazlari', methods=['GET'])
def test_cihazlari():
    """Cofle test cihazı listesi (operatör mobil dropdown'u için). Token sunucuda
    kalır — mobil API'ye doğrudan erişmez. Entegrasyon kapalıysa boş liste döner."""
    try:
        import cofle_test
        if not cofle_test.etkin():
            return jsonify({'etkin': False, 'cihazlar': []})
        return jsonify({'etkin': True, 'cihazlar': cofle_test.cihaz_listesi()})
    except Exception as e:
        print(f"[test_cihazlari] hata: {e}")
        return jsonify({'etkin': False, 'cihazlar': []})


@app.route('/api/operatorler', methods=['GET'])
def operator_listesi():
    """Operatör listesini döndür. ?bolum= ile filtrelenebilir.
    PIN bilgisi bu endpoint'te VERİLMEZ (operatör mobile bunu çağırır).
    Yönetim için /api/operatorler/yonetim kullanın.
    """
    bolum = request.args.get('bolum', '')
    lokasyon = request.args.get('lokasyon', '')
    conn = get_db()
    sql = "SELECT id, ad, bolum, COALESCE(lokasyon,'TK2') as lokasyon FROM operatorler WHERE 1=1"
    params = []
    if bolum:
        # Admin her bölümün listesinde görünür (PIN modal + vardiya formu dropdown'u)
        sql += " AND (bolum = ? OR ad = ?)"
        params.extend([bolum, ADMIN_ADI])
    if lokasyon:
        sql += " AND COALESCE(lokasyon,'TK2') = ?"
        params.append(lokasyon)
    sql += " ORDER BY ad"
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/api/operatorler/yonetim', methods=['GET'])
@panel_gerekli(izin='operatorler')
def operator_listesi_yonetim():
    """Yönetici için operatör listesi + PIN bilgisi. ?lokasyon= ile tesise göre filtrelenir.
    Dashboard'da PIN tanımlama panelinde kullanılır (aktif lokasyona göre).
    """
    lokasyon = request.args.get('lokasyon')
    conn = get_db()
    sql = "SELECT id, ad, bolum, pin, COALESCE(lokasyon,'TK2') as lokasyon FROM operatorler WHERE 1=1"
    params = []
    if lokasyon:
        sql += " AND COALESCE(lokasyon,'TK2') = ?"
        params.append(lokasyon)
    sql += " ORDER BY COALESCE(lokasyon,'TK2'), bolum, ad"
    rows = conn.execute(sql, params).fetchall()
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
    lokasyon = (data.get('lokasyon') or request.args.get('lokasyon') or '').strip()
    if not op or not pin:
        return jsonify({'hata': 'operator_adi ve pin zorunlu'}), 400
    conn = get_db()
    # lokasyon verildiyse o lokasyondaki operatörü seç (ad artık lokasyonlar arası tek değil);
    # verilmezse ilk eşleşen (geriye dönük uyum).
    # ÇOKLU BÖLÜM: UNIQUE(ad, bolum, lokasyon) sonrası aynı kişinin bölüm başına satırı
    # olabilir — herhangi bir satırın PIN'i tutarsa giriş başarılı (satırlar tek PIN paylaşır,
    # import + PIN güncelleme bunu korur).
    if lokasyon:
        rows = conn.execute("SELECT pin, bolum, COALESCE(lokasyon,'TK2') as lokasyon FROM operatorler WHERE ad=? AND COALESCE(lokasyon,'TK2')=?", (op, lokasyon)).fetchall()
    else:
        rows = conn.execute("SELECT pin, bolum, COALESCE(lokasyon,'TK2') as lokasyon FROM operatorler WHERE ad=?", (op,)).fetchall()
    conn.close()
    if not rows:
        return jsonify({'hata': f'Operatör bulunamadı: {op}'}), 403
    eslesen = next((r for r in rows if (r['pin'] or '0000') == pin), None)
    if eslesen is None:
        return jsonify({'hata': 'PIN yanlış'}), 403
    return jsonify({
        'basarili':     True,
        'operator_adi': op,
        'bolum':        eslesen['bolum'] or '',
        'lokasyon':     eslesen['lokasyon'] or 'TK2',
    }), 200


@app.route('/api/operator/<int:oid>/pin', methods=['PATCH'])
@panel_gerekli(izin='operatorler')
def operator_pin_guncelle(oid):
    """Yönetici operatörün PIN'ini günceller. Body: {pin}
    PIN 4 haneli rakam olmalıdır.
    """
    data = request.get_json() or {}
    yeni_pin = (data.get('pin') or '').strip()
    if not yeni_pin or len(yeni_pin) != 4 or not yeni_pin.isdigit():
        return jsonify({'hata': "PIN 4 haneli rakam olmalı (örn: '1234')"}), 400
    conn = get_db()
    row = conn.execute("SELECT ad, COALESCE(lokasyon,'TK2') as lokasyon FROM operatorler WHERE id=?", (oid,)).fetchone()
    if not row:
        conn.close()
        return jsonify({'hata': 'Operatör bulunamadı'}), 404
    # Bir kişi = tek PIN: aynı kişinin tüm bölüm satırlarına yaz (aynı tesiste).
    # (Çoklu bölüm modeli: UNIQUE(ad, bolum, lokasyon) — kişi başına bölüm satırları var.)
    conn.execute(
        "UPDATE operatorler SET pin=? WHERE ad=? AND COALESCE(lokasyon,'TK2')=?",
        (yeni_pin, row['ad'], row['lokasyon'])
    )
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
    lokasyon = request.args.get('lokasyon')
    # TK1: ayrı Excel + import_tk1 (mirror-sync YOK, TK2 verisini silmez)
    if lokasyon == 'TK1':
        try:
            return jsonify(import_tk1()), 200
        except Exception as e:
            return jsonify({'hata': str(e), 'basarili': False}), 500
    if bolum and bolum not in GECERLI_BOLUMLER:
        return jsonify({'hata': f"Geçersiz bölüm: {bolum}", 'basarili': False}), 400
    try:
        sonuc = import_data(bolum=bolum)
        # Kaynak bölümü aktarımı ek sayfaları da kapsar: Robot Program Listesi + Fikstür
        # Raf (eski 'Toplu Veri Yönetimi' butonundan taşındı — tek buton, tek akış).
        if sonuc.get('basarili') and (bolum == 'kaynak' or not bolum):
            from import_excel import kaynak_ek_import
            sonuc.update(kaynak_ek_import())
        return jsonify(sonuc), 200
    except Exception as e:
        return jsonify({'hata': str(e), 'basarili': False}), 500


@app.route('/api/veri/import_tum', methods=['POST'])
def veri_import_tum():
    """Tüm Excel sayfalarını okuyup DB'yi günceller (tek tıkla import).
    Kapsam: referans + operatör (3 bölüm), robot program, fikstür raf.
    Duruş sebepleri her API isteğinde Excel'den okunur, import gerekmez.
    """
    # TK1: ayrı Excel + import_tk1 (TK2 dosyasını okumaz, TK2 verisini silmez)
    lokasyon = request.args.get('lokasyon')
    try:
        if lokasyon == 'TK1':
            return jsonify(import_tk1()), 200
        sonuc = import_tum()
        return jsonify(sonuc), 200
    except Exception as e:
        return jsonify({'hata': str(e), 'basarili': False}), 500


@app.route('/api/referans/teyit', methods=['PATCH'])
@panel_gerekli(izin='referanslar')
def referans_teyit_guncelle():
    """Bir referansın süre teyidi bayrağını aç/kapat.
    Body: {referans_kodu, sure_teyit: 0|1, lokasyon?} — lokasyon ile o fabrikanın
    satırı hedeflenir (diğer fabrikanın aynı kodlu referansının teyidi değişmez)."""
    data = request.get_json() or {}
    kod = (data.get('referans_kodu') or '').strip()
    teyit = 1 if data.get('sure_teyit') else 0
    lokasyon = (data.get('lokasyon') or request.args.get('lokasyon') or 'TK2').strip() or 'TK2'
    # Süre teyidi kaynak bölümüne özgü bir özellik — aynı kod başka bölümde de (lazer/pres)
    # yaşayabildiğinden güncelleme bölümle kapatılır (body'de bolum gelirse o, yoksa kaynak).
    bolum = (data.get('bolum') or 'kaynak').strip()
    if bolum not in GECERLI_BOLUMLER:
        bolum = 'kaynak'
    if not kod:
        return jsonify({'hata': 'referans_kodu zorunlu'}), 400
    conn = get_db()
    try:
        if teyit == 1:
            conn.execute(
                "UPDATE referans_listesi SET sure_teyit=1, sure_teyit_tarihi=datetime('now','localtime') "
                "WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ','')) AND COALESCE(bolum,'kaynak')=? AND COALESCE(lokasyon,'TK2')=?",
                (kod, bolum, lokasyon)
            )
        else:
            conn.execute(
                "UPDATE referans_listesi SET sure_teyit=0, sure_teyit_tarihi=NULL "
                "WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ','')) AND COALESCE(bolum,'kaynak')=? AND COALESCE(lokasyon,'TK2')=?",
                (kod, bolum, lokasyon)
            )
        conn.commit()
        ref_row = {'bolum': bolum}
    finally:
        conn.close()
    # Excel auto-sync: teyit bilgisi artık Excel'de 'Süre Teyit' (E) kolonunda tutulur.
    # (Eskiden teyit yalnız DB'ye yazılıyor, kullanıcı Excel'de değişiklik göremiyordu.)
    excel_uyari = None
    if lokasyon != 'TK1' and ref_row:
        try:
            sonuc = export_referans_cycle_times(bolum=ref_row['bolum'], lokasyon=lokasyon)
            if not sonuc.get('basarili'):
                excel_uyari = sonuc.get('hata')
        except Exception as e:
            excel_uyari = str(e)
        if excel_uyari:
            print(f'[referans/teyit] Excel sync hatası: {excel_uyari}')
    return jsonify({'basarili': True, 'referans_kodu': kod, 'sure_teyit': teyit,
                    'excel_uyari': excel_uyari}), 200


@app.route('/api/saha_cihazlari', methods=['GET'])
def saha_cihazlari():
    """Beklenen tüm saha cihazlarının durumu (kaynak/montaj/metal).

    Beklenen cihaz listesi sabit:
      - 9 robot:    ABB1-IO ... ABB9-IO       (bolum: kaynak, TK2)
      - 12 montaj:  MONTAJ-M1 ... MONTAJ-M12  (bolum: montaj, TK2) + MONTAJ-YF1 (TK1)
      - 3 metal:    300T-IO, 400T-IO, 550T-IO (bolum: metal, TK2)
      - 3 abkant:   ABKANT-A1/A2/A3            (bolum: pres, TK2)
      - 3 plastik:  PLASTIK-320T/407T/YAPISTIRMA (bolum: plastik, TK1)

    Her cihaz için pilot.db'deki cihaz_kayitlari ile eşleştirilir.
    Eşleşme yoksa 'beklemede' (hiç bağlanmamış). ?lokasyon= ile tesise göre süzülür.
    """
    import os
    pilot_db = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pilot', 'pilot.db')

    # ?lokasyon= verilirse sadece o tesisin cihazları döner (saha cihazları + sinyal
    # analizi TK1/TK2'ye özel). Entry'de lokasyon yoksa 'TK2' varsayılır.
    lokasyon = request.args.get('lokasyon')
    BEKLENEN = {
        'kaynak': [{'cihaz_id': f'ABB{i}-IO', 'robot_no': f'ABB{i}'} for i in range(1, 10)],
        'montaj': [{'cihaz_id': f'MONTAJ-M{i}', 'robot_no': f'M{i}'} for i in range(1, 13)]
                  + [{'cihaz_id': 'MONTAJ-YF1', 'robot_no': 'YF1', 'lokasyon': 'TK1'}],
        'metal':  [
            {'cihaz_id': '300T-IO', 'robot_no': '300T'},
            {'cihaz_id': '400T-IO', 'robot_no': '400T'},
            {'cihaz_id': '550T-IO', 'robot_no': '550T'},
        ],
        # 2026-07-27 saha rollout — abkant (pres bölümü, TK2) + plastik enjeksiyon &
        # yapıştırma (plastik bölümü, TK1). cihaz_id + robot_no firmware ile birebir
        # (generate.py DEVICES). Eksantrik Pres (Pres 1-5) modülleri hazır olunca eklenir.
        'pres':   [
            {'cihaz_id': 'ABKANT-A1', 'robot_no': 'Abkant 1'},
            {'cihaz_id': 'ABKANT-A2', 'robot_no': 'Abkant 2'},
            {'cihaz_id': 'ABKANT-A3', 'robot_no': 'Abkant 3'},
        ],
        'plastik': [
            {'cihaz_id': 'PLASTIK-320T',       'robot_no': '320T',       'lokasyon': 'TK1'},
            {'cihaz_id': 'PLASTIK-407T',       'robot_no': '407T',       'lokasyon': 'TK1'},
            {'cihaz_id': 'PLASTIK-YAPISTIRMA', 'robot_no': 'Yapistirma', 'lokasyon': 'TK1'},
        ],
    }

    # Pilot DB'den mevcut cihaz_kayitlari + bugünün ist1/ist2 sayıları
    mevcut = {}
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
    # Operator uretim baseline — yeni referans kaydi acinca sayac sifirlanir
    op_baseline_map = {}  # {(bolum, robot_no, istasyon): sayac_baslangic_ts}
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
            WHERE durum != 'kapali' AND tarih=?
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

        # Operatör üretim baseline — operatör yeni referans kaydı açınca sayac_baslangic_ts
        # set edilir; andon o andan sonrasını sayar (operatör aksiyonuyla sayaç sıfırlanır).
        # (bolum, robot_no, istasyon) -> en son baseline ts
        for r in conn_main.execute('''
            SELECT COALESCE(v.bolum,'kaynak') as bolum, v.robot_no, u.istasyon,
                   MAX(u.sayac_baslangic_ts) as bts
            FROM uretim_kayitlari u
            JOIN vardiyalar v ON u.vardiya_id = v.id
            WHERE v.durum != 'kapali' AND v.tarih=? AND u.sayac_baslangic_ts IS NOT NULL
            GROUP BY v.bolum, v.robot_no, u.istasyon
        ''', (bugun_str,)).fetchall():
            if r['robot_no'] and r['bts']:
                op_baseline_map[(r['bolum'], r['robot_no'], int(r['istasyon'] or 0))] = r['bts']
        # NOT: conn_main = get_db() g-paylaşımlı — istek ortasında KAPATMA (teardown kapatır).
        # Erken close, sonrasına get_db kullanan kod eklenirse 'closed database' 500 üretir.
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
            dev_lok = beklenen_cihaz.get('lokasyon', 'TK2')
            if lokasyon and dev_lok != lokasyon:
                continue   # bu tesisin cihazı değil — atla
            kayit = mevcut.get(cid)

            # PRES (2026-07-28): makine artik VARDIYADA degil URETIM KAYDINDA secilir.
            # Cihaz 'Abkant 2' → vardiya hatti PRES_SABIT_HAT, baseline istasyonu = 2.
            # ESKI kayitlar (vardiya.robot_no='Abkant 2', istasyon=0) icin fallback var.
            v_rno, pres_ist = rno, 0
            if bolum == 'pres':
                _pi = pres_ist_no(rno)
                if _pi:
                    v_rno, pres_ist = PRES_SABIT_HAT, _pi

            # Aktif vardiya baslangic ts (varsa) — sayac sifirlama referansi
            # (yeni model once; bulunamazsa eski makine-adli vardiyaya dus)
            vardiya_ts = aktif_vardiya_map.get((bolum, v_rno)) or aktif_vardiya_map.get((bolum, rno))
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
                        op_bl = op_baseline_map.get((bolum, rno, ist))
                        filt_ts = _en_son_ts(vardiya_ts, a['basla_ts'], reset_ist, reset_all, op_bl)
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
                    op_bl = op_baseline_map.get((bolum, rno, 0)) or op_baseline_map.get((bolum, rno, 1))
                    filt_ts = _en_son_ts(vardiya_ts, a['basla_ts'], reset_all, reset_ist1, op_bl)
                    aktif_referanslar.append({
                        'istasyon':      None,
                        'referans_kodu': a['referans_kodu'],
                        'basla_ts':      filt_ts,
                        'pulse_sayisi':  _aktif_pulse_say(cid, 0, filt_ts),
                    })

            # Sayac degerlerini hesapla — uretim kapisi = AKTIF VARDIYA (bkz. vardiya_acik):
            # - Aktif vardiya varsa: vardiya basindan (veya daha gec reset/baseline'dan) say
            # - Aktif vardiya yoksa: 0 (vardiya acilmadan gelen sinyal = hat paraziti, uretim degil)
            if kayit:
                # Istasyon basina baslangic referanslari (operator uretim baseline'i)
                op_bl_1 = op_baseline_map.get((bolum, rno, 1))
                op_bl_2 = op_baseline_map.get((bolum, rno, 2))
                op_bl_0 = op_baseline_map.get((bolum, rno, 0))
                # PRES: baseline anahtarinda istasyon = MAKINE no (sabit hat + 1..8).
                # Cihaz tek makinedir → tek sayac (op_bl_0 gibi davranir); eski
                # makine-adli kayitlarin baseline'i (rno, 0) fallback olarak kalir.
                if pres_ist:
                    op_bl_0 = _en_son_ts(op_baseline_map.get((bolum, PRES_SABIT_HAT, pres_ist)),
                                         op_bl_0)
                    op_bl_1 = op_bl_2 = None
                # ── URETIM KAPISI: AKTIF VARDIYA ──────────────────────────────
                # Sayim ancak ACIK bir vardiya (veya o vardiyada acilmis operator
                # uretim baseline'i) varsa yapilir. Aksi halde uretim baglami yoktur:
                # vardiya acilmadan gelen sinyaller (gece/hafta sonu hat paraziti)
                # uretim olarak SAYILMAZ.
                # NOT: gunler oncesinden kalan BAYAT manuel reset noktasi TEK BASINA
                # sayim baslatmaz; reset ancak acik vardiya icinde "su andan say" demektir.
                vardiya_acik = bool(vardiya_ts or op_bl_0 or op_bl_1 or op_bl_2)
                if not vardiya_acik:
                    ist1_v = 0
                    ist2_v = 0
                    toplam_v = 0
                else:
                    # En gec baslangic noktasindan say: vardiya / manuel reset / baseline
                    ist1_filt = _en_son_ts(vardiya_ts, reset_ist1, reset_all, op_bl_1)
                    ist2_filt = _en_son_ts(vardiya_ts, reset_ist2, reset_all, op_bl_2)
                    ist1_v = _aktif_pulse_say(cid, 1, ist1_filt) if ist1_filt else 0
                    ist2_v = _aktif_pulse_say(cid, 2, ist2_filt) if ist2_filt else 0
                    # toplam = cihazin TUM pulse'lari (istasyon farketmez).
                    # _aktif_pulse_say(cid, 0, ts) -> istasyon=0 falsy oldugu icin filtre
                    # UYGULAMAZ, tum pulse'lari sayar. ESKI BUG: ist1_v + ist2_v + toplam_0
                    # idi; montaj istasyon=1 gonderdigi icin o pulse hem ist1_v'de hem
                    # toplam_0'da sayiliyordu = 2x. Duzeltme: dogrudan toplam say.
                    toplam_ts = _en_son_ts(vardiya_ts, reset_all, op_bl_0)
                    toplam_v = _aktif_pulse_say(cid, 0, toplam_ts) if toplam_ts else (ist1_v + ist2_v)

                # ── "Makine calisiyor" durumu ──────────────────────────────────
                # Metal: sinyal akisindan TURETILIR. Her sinyal = makine-calisiyor rolesi
                # darbesi (400T o roleyi izler). Son sinyal ESIK sn icindeyse calisiyor;
                # sinyal kesilince durmus. Metal firmware robot_calisiyor=0 gonderir
                # (kullanmaz), o yuzden bu alani metal'de biz uretiyoruz. Vardiya sart:
                # vardiya yokken gelen sinyal hat parazitidir, "calisiyor" demek degildir.
                # Kaynak/montaj: cihazin gercek robot_calisiyor IO sinyali aynen gecer.
                if bolum == 'metal':
                    robot_calisiyor_deger = 0
                    son_sin = kayit.get('son_sinyal')
                    if vardiya_acik and son_sin:
                        try:
                            gecen_sn = (datetime.now() - datetime.strptime(son_sin, '%Y-%m-%d %H:%M:%S')).total_seconds()
                            if 0 <= gecen_sn < METAL_CALISIYOR_ESIK_SN:
                                robot_calisiyor_deger = 1
                        except (ValueError, TypeError):
                            pass
                else:
                    robot_calisiyor_deger = 1 if kayit.get('robot_calisiyor') else 0

                liste.append({
                    'cihaz_id':           cid,
                    'robot_no':           rno,
                    'bolum':              bolum,
                    'lokasyon':           dev_lok,
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
                    'robot_calisiyor':    robot_calisiyor_deger,
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
                    'lokasyon':           dev_lok,
                    'durum':              'beklemede',
                    'aktif_vardiya_ts':   vardiya_ts,
                    'aktif_referanslar':  aktif_referanslar,
                    'kayitli':            False,
                })
        sonuc[bolum] = liste

    return jsonify(sonuc), 200


@app.route('/api/saha_cihazlari/sayac_reset', methods=['POST'])
@panel_gerekli(izin='saha-cihazlari')
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


@app.route('/api/pilot/durus_tanimla', methods=['POST'])
@panel_gerekli(izin='andon-ayarlari')
def pilot_durus_tanimla():
    """Sinyal Analizi'nde 'Kaydedilmemiş' bir boşluğa sebep atar.
    Sinyalden gelen GERÇEK başlangıç saati + süre ile bir duruş kaydı oluşturur
    (baslangic_saati dolu → bir dahaki analizde tam eşleşir). Rapor SUM ile çalışır,
    bu yüzden saatli/ayrı kayıt Andon/OEE toplamlarını bozmaz."""
    data = request.get_json() or {}
    robot_no = (data.get('robot_no') or '').strip()
    tarih    = (data.get('tarih') or '').strip()
    bolum    = (data.get('bolum') or '').strip()
    basla    = (data.get('basla') or '').strip()   # ISO datetime (boşluğun başlangıcı)
    sebep    = (data.get('durus_sebebi') or '').strip()
    tip      = (data.get('durus_tipi') or 'plansiz').strip()
    try:
        sure_dk = int(round(float(data.get('sure_dk') or 0)))
    except (ValueError, TypeError):
        sure_dk = 0
    if not robot_no or not tarih or not sebep or sure_dk <= 0:
        return jsonify({'hata': 'robot_no, tarih, durus_sebebi, sure_dk zorunlu'}), 400
    # Başlangıç saatini HH:MM'e çevir (timed kayıt — gelecekte kesin eşleşme)
    hhmm = ''
    try:
        hhmm = datetime.fromisoformat(basla).strftime('%H:%M') if basla else ''
    except Exception:
        hhmm = ''
    conn = get_db()
    try:
        v = None
        if bolum and _kolon_var(conn, 'vardiyalar', 'bolum'):
            v = conn.execute("SELECT id FROM vardiyalar WHERE robot_no=? AND tarih=? AND bolum=? ORDER BY id DESC LIMIT 1",
                             (robot_no, tarih, bolum)).fetchone()
        if not v:
            v = conn.execute("SELECT id FROM vardiyalar WHERE robot_no=? AND tarih=? ORDER BY id DESC LIMIT 1",
                             (robot_no, tarih)).fetchone()
        if not v:
            return jsonify({'hata': 'Bu cihaz/tarih için vardiya bulunamadı'}), 404
        conn.execute('''INSERT INTO duruslar (vardiya_id, durus_sebebi, aciklama, sure_dk, baslangic_saati, durus_tipi)
                        VALUES (?, ?, ?, ?, ?, ?)''',
                     (v['id'], sebep, (data.get('aciklama') or '').strip(), sure_dk, hhmm, tip))
        conn.commit()
    finally:
        conn.close()
    return jsonify({'basarili': True, 'baslangic_saati': hhmm, 'sure_dk': sure_dk}), 201


@app.route('/api/pilot/sinyal_kalitesi', methods=['GET'])
@panel_gerekli(izin='sinyal-analizi')
def pilot_sinyal_kalitesi():
    """Sinyal Kalitesi / Tani paneli — cihaz başına sayım kararları + kayıp mutabakatı.
    Pilot.db'den okunur: cihaz_kayitlari (tani sayaçları + pulse_ist baseline) + sayac_olaylari.
      - sayildi  : gerçek sayım (firmware kararı)
      - parazit  : elenen gürültü (ham HIGH→LOW kenar) — sistemin ayıkladığı
      - erken    : MIN_PULSE_GAP altında reddedilen
      - kayip    : cihaz NVS sayımı − server'a ulaşan (transit/WiFi kaybı)
      - son_kararlar: son tani olayları (SAYILDI/ERKEN) zaman akışı
    """
    import os, sqlite3
    lokasyon = request.args.get('lokasyon')   # TK1/TK2'ye göre cihaz filtresi (robot_no eşlemesi)
    pilot_db = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pilot', 'pilot.db')
    cihazlar, son_kararlar = [], []
    if not os.path.exists(pilot_db):
        return jsonify({'cihazlar': [], 'son_kararlar': [], 'hata': 'Pilot DB yok'}), 200
    try:
        pc = sqlite3.connect(pilot_db)
        pc.row_factory = sqlite3.Row
        try:
            rows = pc.execute('''
                SELECT cihaz_id, bolum, robot_no, wifi_rssi,
                       tani_sayildi, tani_parazit, tani_erken,
                       pulse_ist1, pulse_ist2, base_ist1, base_ist2,
                       CAST((julianday('now','localtime')-julianday(son_heartbeat))*1440 AS INTEGER) hb_dk
                FROM cihaz_kayitlari
            ''').fetchall()
            for r in rows:
                cid = r['cihaz_id']
                sc1 = pc.execute("SELECT COUNT(*) FROM sayac_olaylari WHERE cihaz_id=? AND istasyon=1", (cid,)).fetchone()[0]
                sc2 = pc.execute("SELECT COUNT(*) FROM sayac_olaylari WHERE cihaz_id=? AND istasyon=2", (cid,)).fetchone()[0]
                b1 = r['base_ist1'] if r['base_ist1'] is not None else 0
                b2 = r['base_ist2'] if r['base_ist2'] is not None else 0
                kayip = max(0, (r['pulse_ist1'] or 0) + b1 - sc1) + max(0, (r['pulse_ist2'] or 0) + b2 - sc2)
                sayildi = r['tani_sayildi'] or 0
                parazit = r['tani_parazit'] or 0   # TÜM ham HIGH→LOW kenar (gerçek pulse kenarı dahil)
                erken   = r['tani_erken'] or 0
                # Elenen gürültü = ham kenar − sayılan − erken (gerçek pulse'ların kenarı düşülür)
                # Temiz sinyalde ~0 (her kenar sayıldı); 400T gibi gürültülü hatta yüksek.
                elenen_gurultu = max(0, parazit - sayildi - erken)
                cihazlar.append({
                    'cihaz_id': cid, 'bolum': r['bolum'], 'robot_no': r['robot_no'],
                    'wifi_rssi': r['wifi_rssi'], 'online': (r['hb_dk'] if r['hb_dk'] is not None else 999) <= 2,
                    'sayildi': sayildi, 'elenen_gurultu': elenen_gurultu, 'erken': erken,
                    'ham_kenar': parazit, 'kayip': kayip,
                })
            for r in pc.execute('''
                SELECT cihaz_id, robot_no, tip, gap_ms, low_ms, ts
                FROM tani_olaylari ORDER BY id DESC LIMIT 30
            ''').fetchall():
                son_kararlar.append(dict(r))
        finally:
            pc.close()
    except Exception as e:
        return jsonify({'cihazlar': [], 'son_kararlar': [], 'hata': str(e)}), 200
    # Lokasyon filtresi (TK1 = TK1_ROBOT_NOLARI; TK2 = geri kalan)
    if lokasyon == 'TK1':
        cihazlar = [c for c in cihazlar if c.get('robot_no') in TK1_ROBOT_NOLARI]
        son_kararlar = [k for k in son_kararlar if k.get('robot_no') in TK1_ROBOT_NOLARI]
    elif lokasyon == 'TK2':
        cihazlar = [c for c in cihazlar if c.get('robot_no') not in TK1_ROBOT_NOLARI]
        son_kararlar = [k for k in son_kararlar if k.get('robot_no') not in TK1_ROBOT_NOLARI]
    # En çok sinyal işleyen (tani verisi olan) cihazlar üstte
    cihazlar.sort(key=lambda x: (x['sayildi'] + x['ham_kenar']), reverse=True)
    return jsonify({'cihazlar': cihazlar, 'son_kararlar': son_kararlar})


@app.route('/api/pilot/sinyal_analiz', methods=['GET'])
@panel_gerekli(izin='sinyal-analizi')
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
    # Duruş eşiği: bu süreyi aşan gap = duruş (gerçek cycle değil). ?durus_esik_sn ile ayarlanır.
    DURUS_ESIK_SN = max(60, int(request.args.get('durus_esik_sn', 300) or 300))
    # Kayıtlı duruşlar (SEBEP eşleştirmesi için). NOT: duruslar tablosu aynı vardiya+sebep'i
    # TOPLAR (Andon toplamı için doğru) → kayıt zamanı timeline'da yanıltıcı. O yüzden
    # duruşları SİNYALDEN (uzun gap) gerçek zamanında türetip sebebi buradan eşleştiriyoruz.
    conn = get_db()
    kayitli_duruslar = []
    try:
        for _r in conn.execute('''
            SELECT d.baslangic_saati, d.sure_dk, d.durus_sebebi, d.durus_tipi, d.aciklama
            FROM duruslar d JOIN vardiyalar v ON d.vardiya_id = v.id
            WHERE v.robot_no=? AND v.tarih=? ORDER BY d.baslangic_saati
        ''', (robot_no, tarih)).fetchall():
            # NOT: kayıtların ~%100'ünde baslangic_saati BOŞ — bu yüzden eşleştirme
            # öncelikle SÜRE ile yapılır (40dk gap ↔ 40dk kayıt). Saat varsa zaman da kullanılır.
            _bs = (_r['baslangic_saati'] or '').strip()
            _sure_sn = (_r['sure_dk'] or 0) * 60
            _bdt = None
            if _bs:
                try:
                    _iso = _bs if ('T' in _bs or len(_bs) > 8) else (f'{tarih}T{_bs}:00' if len(_bs) == 5 else f'{tarih}T{_bs}')
                    _bdt = datetime.fromisoformat(_iso)
                except Exception:
                    _bdt = None
            kayitli_duruslar.append({
                'basla_dt': _bdt,
                'biti_dt':  (_bdt + timedelta(seconds=_sure_sn)) if _bdt else None,
                'sure_sn':  _sure_sn,
                'sebep':    _r['durus_sebebi'] or '',
                'tip':      _r['durus_tipi'] or 'plansiz',
                'aciklama': _r['aciklama'] or '',
                'used':     False,
            })
    except Exception as _e:
        print(f'[sinyal_analiz] kayitli durus hata: {_e}')

    def _durus_gap_mi(prev_dt, cur_dt, gap_sn):
        """Gap bir duruşa mı denk geliyor? (eşik üstü VEYA kayıtlı duruşla örtüşen)"""
        if gap_sn is not None and gap_sn > DURUS_ESIK_SN:
            return True
        if prev_dt is None or cur_dt is None:
            return False
        for kd in kayitli_duruslar:
            if kd['basla_dt'] is not None and prev_dt < kd['biti_dt'] and cur_dt > kd['basla_dt']:
                return True
        return False

    # Cycle istatistikleri SADECE gerçek cycle gap'lerinden (duruş gap'leri hariç) hesaplanır
    cycle_gaps = []
    haric_durus = 0
    for o in olaylar:
        g = o['gap_sn']
        if g is None:
            o['durus'] = False
            continue
        try:
            _cur = datetime.fromisoformat(o['ts'])
            _prev = _cur - timedelta(seconds=g)
        except Exception:
            _cur = _prev = None
        _isdurus = _durus_gap_mi(_prev, _cur, g)
        o['durus'] = _isdurus            # grafik scatter bu bayrakla duruş noktasını atlar
        if _isdurus:
            haric_durus += 1
            continue
        cycle_gaps.append(g)
    gaps = cycle_gaps
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
        'haric_durus_gap': haric_durus,   # cycle dışı bırakılan (duruş) gap sayısı
    }

    # ─── Saatlik dağılım ─────────────────────────────────────────
    saatlik_map = {}  # {hour: [gap1, gap2, ...]}
    saatlik_sayim = {}  # {hour: pulse_count}
    for o in olaylar:
        try:
            _cur = datetime.fromisoformat(o['ts'])
        except Exception:
            continue
        h = _cur.hour
        saatlik_sayim[h] = saatlik_sayim.get(h, 0) + 1
        g = o['gap_sn']
        if g is not None and not _durus_gap_mi(_cur - timedelta(seconds=g), _cur, g):
            saatlik_map.setdefault(h, []).append(g)
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
    # SİNYALDEN türet: birleşik pulse akışında eşik üstü boşluk = bir duruş (GERÇEK zaman).
    # Kayıttaki toplanmış süre/zaman KULLANILMAZ; sebep, kayıtlı duruştan eşleştirilir.
    def _match_durus(prev_dt, cur_dt, gap_sn):
        # 1) Saati DOLU kayıtla zaman örtüşmesi (nadir — kayıtların çoğu saatsiz)
        for kd in kayitli_duruslar:
            if (not kd['used'] and kd['basla_dt'] is not None
                    and prev_dt < kd['biti_dt'] and cur_dt > kd['basla_dt']):
                kd['used'] = True
                return kd
        # 2) SÜRE eşleşmesi (saatsiz kayıtlar): süresi gap'e en yakın, tolerans içinde kayıt
        aday, en_fark = None, None
        tol = max(120, gap_sn * 0.25)   # ±%25 veya en az 2dk
        for kd in kayitli_duruslar:
            if kd['used']:
                continue
            fark = abs(kd['sure_sn'] - gap_sn)
            if fark <= tol and (en_fark is None or fark < en_fark):
                en_fark, aday = fark, kd
        if aday is not None:
            aday['used'] = True
        return aday
    tum_ts = []
    for o in olaylar:
        try:
            tum_ts.append(datetime.fromisoformat(o['ts']))
        except Exception:
            pass
    tum_ts.sort()
    # Kaydedilmemiş (sebebi olmayan) duruş ancak >=10dk ise timeline'a gelsin —
    # kısa boşluklar "Kaydedilmemiş duruş" olarak kalabalık yapmasın. ?kayitsiz_esik_sn ile ayarlanır.
    KAYITSIZ_DURUS_ESIK_SN = max(0, int(request.args.get('kayitsiz_esik_sn', 600) or 600))
    duruslar = []
    for _i in range(1, len(tum_ts)):
        prev_dt, cur_dt = tum_ts[_i-1], tum_ts[_i]
        gap_sn = (cur_dt - prev_dt).total_seconds()
        if gap_sn <= DURUS_ESIK_SN:
            continue
        kd = _match_durus(prev_dt, cur_dt, gap_sn)
        if kd is None and gap_sn < KAYITSIZ_DURUS_ESIK_SN:
            continue  # kısa + kaydedilmemiş → timeline'a ekleme
        duruslar.append({
            'basla':    prev_dt.isoformat(),
            'biti':     cur_dt.isoformat(),
            'sure_dk':  round(gap_sn / 60, 1),
            'sebep':    (kd['sebep'] if kd else '') or 'Kaydedilmemiş',
            'tip':      kd['tip'] if kd else 'plansiz',
            'aciklama': kd['aciklama'] if kd else '',
        })

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
    """Bölüm bazında teyit özeti: toplam / teyitli / teyitsiz / süresiz sayıları.
    ?lokasyon= ile tesise göre filtrelenir (montaj çağrısında TK1/TK2 karışmasın)."""
    bolum = request.args.get('bolum', 'kaynak')
    lokasyon = request.args.get('lokasyon')
    sql = '''
        SELECT
          COUNT(*) as toplam,
          SUM(CASE WHEN COALESCE(sure_teyit,0)=1 THEN 1 ELSE 0 END) as teyitli,
          SUM(CASE WHEN COALESCE(sure_teyit,0)=0 AND COALESCE(hedef_cycle_time_sn,0)>0 THEN 1 ELSE 0 END) as teyitsiz,
          SUM(CASE WHEN COALESCE(hedef_cycle_time_sn,0)=0 THEN 1 ELSE 0 END) as suresiz
        FROM referans_listesi
        WHERE COALESCE(bolum,'kaynak')=?
    '''
    params = [bolum]
    if lokasyon:
        sql += " AND COALESCE(lokasyon,'TK2')=?"
        params.append(lokasyon)
    conn = get_db()
    rows = conn.execute(sql, params).fetchone()
    conn.close()
    d = dict(rows) if rows else {}
    return jsonify({
        'bolum': bolum,
        'lokasyon': lokasyon or '',
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
        # Bu endpoint kaynak (TK2) içindir → kaynak bölümü + TK2'ye kapat
        # (aynı kodlu TK1 satırına ya da başka bölümün — lazer/pres — satırına dokunma)
        c = conn.cursor()
        c.execute(
            "UPDATE referans_listesi SET kaynak_suresi_sn=?, soktak_suresi_sn=?, hedef_cycle_time_sn=? "
            "WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ','')) "
            "AND COALESCE(bolum,'kaynak')='kaynak' AND COALESCE(lokasyon,'TK2')='TK2'",
            (kaynak, soktak, toplam, kod)
        )
        # Geriye dönük: bu kodla mevcut TK2 KAYNAK üretim kayıtlarındaki cycle_time'ı güncelle
        c.execute(
            "UPDATE uretim_kayitlari SET cycle_time_sn=? "
            "WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ','')) "
            "AND vardiya_id IN (SELECT id FROM vardiyalar WHERE COALESCE(lokasyon,'TK2')='TK2' AND COALESCE(bolum,'kaynak')='kaynak')",
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
    lokasyon = (request.args.get('lokasyon') or 'TK2').strip() or 'TK2'
    if bolum not in GECERLI_BOLUMLER:
        return jsonify({'hata': f"Geçersiz bölüm: {bolum}"}), 400
    sebepler = durus_sebepleri_yukle(bolum, lokasyon)
    resp = jsonify({'bolum': bolum, 'lokasyon': lokasyon, 'sebepler': sebepler})
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
    lokasyon = request.args.get('lokasyon')

    # Sorgu bugünü kapsıyorsa: açık vardiyaların auto sayaçlarını ÖNCE tazele —
    # dashboard performans/OEE'si canlı makine adetleriyle hesaplansın
    # (andon rozeti canlıyken performansın bayat kalması sorunu).
    if tarih_bas <= bugun <= tarih_bit:
        _acik_auto_vardiyalari_senkronla(conn, bolum=bolum or None, lokasyon=lokasyon or None)

    param_vardiya = [tarih_bas, tarih_bit]
    sart_vardiya = 'tarih BETWEEN ? AND ?'
    if robot:
        sart_vardiya += ' AND robot_no = ?'
        param_vardiya.append(robot)
    if bolum:
        sart_vardiya += " AND COALESCE(v.bolum, 'kaynak') = ?"
        param_vardiya.append(bolum)
    if lokasyon:
        sart_vardiya += " AND COALESCE(v.lokasyon, 'TK2') = ?"
        param_vardiya.append(lokasyon)

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
    # istasyon: lazer bölümünde makine (1..6) — frontend 'Lazer N' kırılımı bundan çıkar
    # (lazer'de v.robot_no hep 'Lazer'; makine ayrımı yalnız u.istasyon'da).
    uretim_detay_listesi = c.execute(f'''
        SELECT u.id as uretim_id,
               v.tarih,
               v.robot_no,
               v.operator_adi,
               u.referans_kodu,
               u.ok_adet,
               u.nok_adet,
               u.hedef_adet,
               u.cycle_time_sn,
               COALESCE(u.istasyon, 0) as istasyon
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
    lokasyon = request.args.get('lokasyon', '')

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
    if lokasyon:
        # TK1/TK2 montaj raporları karışmasın (operatör/referans/TEEP tüm tipler bu sartı paylaşır)
        sart += " AND COALESCE(v.lokasyon, 'TK2') = ?"
        params.append(lokasyon)

    # Rapor bugünü kapsıyorsa: açık vardiyaların auto sayaçlarını ÖNCE tazele —
    # operatör/referans/TEEP raporları canlı makine adetleriyle hesaplansın
    _bugun_r = date.today().isoformat()
    if tarih_bas <= _bugun_r <= tarih_bit:
        _acik_auto_vardiyalari_senkronla(conn, bolum or None, lokasyon or None)

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
            v_rows = c.execute(
                f'SELECT v.* FROM vardiyalar v WHERE {v_sart}', vp
            ).fetchall()
            vids = [x['id'] for x in v_rows]
            # Açık vardiyalarda "o ana kadar geçen süre" (canlı), kapalılarda gerçek süre
            efektif_calisma_dk = sum(_vardiya_efektif_dk(x) for x in v_rows)

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
                'toplam_calisma_dk': efektif_calisma_dk,
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
            # Hedef CT'yi referans listesinden al — rapor lokasyon filtreliyse o tesisin satırı
            # (composite UNIQUE sonrası aynı kod TK1+TK2'de olabilir)
            if lokasyon:
                ref_row = c.execute(
                    "SELECT hedef_cycle_time_sn FROM referans_listesi WHERE referans_kodu = ? AND COALESCE(lokasyon,'TK2') = ?",
                    (r['referans_kodu'], lokasyon)
                ).fetchone()
            else:
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

        # 2) Toplam vardiya suresi — açık vardiyalarda "o ana kadar geçen süre" (canlı)
        v_rows = c.execute(f'SELECT v.* FROM vardiyalar v WHERE {sart}', params).fetchall()
        toplam_vardiya_dk = sum(_vardiya_efektif_dk(v) for v in v_rows)

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
    """?bolum= ve ?lokasyon= ile filtrele; parametresizse hepsi döner.
    NOT: TK1 hatları (Pull/...) ve TK2 montaj (M1-M12) aynı bolum='montaj' altında —
    lokasyon filtresi ikisini ayırır."""
    bolum = request.args.get('bolum', '').strip()
    lokasyon = request.args.get('lokasyon', '').strip()
    sql = "SELECT * FROM andon_robot_ayarlari WHERE 1=1"
    params = []
    if bolum:
        sql += " AND bolum=?"
        params.append(bolum)
    if lokasyon:
        sql += " AND COALESCE(lokasyon,'TK2')=?"
        params.append(lokasyon)
    sql += " ORDER BY bolum, sira"
    conn = get_db()
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/andon_robot_ayarlari', methods=['PATCH'])
@panel_gerekli(izin='andon-ayarlari')
def andon_robot_ayarlari_guncelle():
    """Toplu güncelleme: [{robot_no, bolum, goster, sira}, ...]
    Eski payload'larda bolum yoksa 'kaynak' varsayılır (geriye dönük uyumluluk)."""
    data = request.get_json() or []
    arg_lok = request.args.get('lokasyon')
    conn = get_db()
    for item in data:
        bolum = item.get('bolum') or 'kaynak'
        lokasyon = item.get('lokasyon') or arg_lok or 'TK2'
        conn.execute(
            "UPDATE andon_robot_ayarlari SET goster=?, sira=? WHERE robot_no=? AND bolum=? AND COALESCE(lokasyon,'TK2')=?",
            (int(item.get('goster', 1)), int(item.get('sira', 0)), item['robot_no'], bolum, lokasyon)
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
# GÜNLÜK ÜRETİM RAPORU E-POSTASI — alıcı yönetimi + manuel gönderim
# SMTP bilgileri mail_config.json'da (gitignore); alıcılar mail_alicilari tablosunda.
# ─────────────────────────────────────────────────────────────
import re as _re
_EMAIL_RE = _re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')


@app.route('/api/mail/durum', methods=['GET'])
@panel_gerekli(izin='raporlar')
def mail_durum():
    """Dashboard göstergesi: SMTP config durumu (şifre içermez) + aktif alıcı sayısı."""
    try:
        import mail_raporu
        d = mail_raporu.durum_ozeti()
        conn = get_db()
        d['alici_sayisi'] = conn.execute("SELECT COUNT(*) FROM mail_alicilari WHERE COALESCE(aktif,1)=1").fetchone()[0]
        conn.close()
        return jsonify(d)
    except Exception as e:
        return jsonify({'config_var': False, 'etkin': False, 'hata': str(e)})


@app.route('/api/mail/alicilar', methods=['GET'])
@panel_gerekli(izin='raporlar')
def mail_alicilar_listesi():
    conn = get_db()
    rows = conn.execute(
        "SELECT id, email, ad, COALESCE(aktif,1) as aktif, created_at FROM mail_alicilari ORDER BY email"
    ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/api/mail/alicilar', methods=['POST'])
@panel_gerekli(izin='raporlar')
def mail_alici_ekle():
    data = request.get_json() or {}
    email = (data.get('email') or '').strip().lower()
    ad = (data.get('ad') or '').strip()
    if not email or not _EMAIL_RE.match(email):
        return jsonify({'hata': 'Geçerli bir e-posta adresi giriniz.'}), 400
    conn = get_db()
    try:
        conn.execute("INSERT INTO mail_alicilari (email, ad, aktif) VALUES (?, ?, 1)", (email, ad))
        conn.commit()
    except Exception:
        conn.close()
        return jsonify({'hata': 'Bu e-posta zaten ekli.'}), 400
    conn.close()
    return jsonify({'basarili': True}), 201


@app.route('/api/mail/alicilar/<int:aid>', methods=['DELETE'])
@panel_gerekli(izin='raporlar')
def mail_alici_sil(aid):
    conn = get_db()
    conn.execute("DELETE FROM mail_alicilari WHERE id=?", (aid,))
    conn.commit()
    conn.close()
    return jsonify({'basarili': True})


@app.route('/api/mail/alicilar/<int:aid>', methods=['PATCH'])
@panel_gerekli(izin='raporlar')
def mail_alici_guncelle(aid):
    """aktif bayrağını aç/kapat (listeden silmeden geçici durdurma)."""
    data = request.get_json() or {}
    aktif = 1 if data.get('aktif') else 0
    conn = get_db()
    conn.execute("UPDATE mail_alicilari SET aktif=? WHERE id=?", (aktif, aid))
    conn.commit()
    conn.close()
    return jsonify({'basarili': True})


@app.route('/api/mail/gonder_simdi', methods=['POST'])
@panel_gerekli(izin='raporlar')
def mail_gonder_simdi():
    """Günlük raporu ŞİMDİ üretip aktif alıcılara gönderir (17:00'ı beklemeden test/manuel)."""
    data = request.get_json(silent=True) or {}
    tarih = data.get('tarih')  # opsiyonel YYYY-MM-DD; yoksa bugün
    try:
        import mail_raporu
        sonuc = mail_raporu.gunluk_mail_gonder(tarih=tarih)
        return jsonify(sonuc), (200 if sonuc.get('basarili') else 400)
    except Exception as e:
        return jsonify({'basarili': False, 'mesaj': str(e)}), 500


@app.route('/api/mail/test', methods=['POST'])
@panel_gerekli(izin='raporlar')
def mail_test_gonder():
    """Tek adrese test maili — SMTP kurulumunu doğrulamak için."""
    data = request.get_json() or {}
    email = (data.get('email') or '').strip().lower()
    if not email or not _EMAIL_RE.match(email):
        return jsonify({'basarili': False, 'mesaj': 'Geçerli bir e-posta adresi giriniz.'}), 400
    try:
        import mail_raporu
        sonuc = mail_raporu.test_gonder(email)
        return jsonify(sonuc), (200 if sonuc.get('basarili') else 400)
    except Exception as e:
        return jsonify({'basarili': False, 'mesaj': str(e)}), 500



# ─────────────────────────────────────────────────────────────
# REFERANS TAKİP (LAUNCH / İŞ EMRİ) API
# ─────────────────────────────────────────────────────────────

DURUM_SIRASI = ['launch_alinacak', 'launch_alindi', 'launch_hazir', 'uretimde', 'uretim_tamamlandi']

@app.route('/api/referans_takip', methods=['GET'])
def referans_takip_listesi():
    bolum = request.args.get('bolum')
    lokasyon = request.args.get('lokasyon')
    conn = get_db()
    # Tüm bölümlerde öncelik ASC, NULL olanlar en sonda → oluşturma tarihine göre
    # Kaynak/Söktak süreleri de JOIN ile eklendi (pair_cycle hesabı için)
    # GROUP BY rt.id — FAN-OUT KORUMASI: referans_listesi'nde aynı normalize koda sahip
    # birden çok satır olabilir (örn. '94.LTK.341/10' vs '94. LTK. 341/10' — operatörün
    # boşluklu yazımı auto-create ile ayrı satır açmıştı). JOIN ikisiyle de eşleşince
    # TEK iş emri listede İKİ kez görünüyordu. MAX(): dolu cycle değeri tercih edilir.
    sql = '''
        SELECT rt.*,
               MAX(rl.hedef_cycle_time_sn) as hedef_cycle_time_sn,
               MAX(rl.kaynak_suresi_sn)    as kaynak_suresi_sn,
               MAX(rl.soktak_suresi_sn)    as soktak_suresi_sn
        FROM referans_takip rt
        LEFT JOIN referans_listesi rl ON REPLACE(rt.referans_kodu, ' ', '') = REPLACE(rl.referans_kodu, ' ', '')
            AND COALESCE(rl.lokasyon,'TK2') = COALESCE(rt.lokasyon,'TK2')
        WHERE 1=1'''
    params = []
    if bolum:
        sql += " AND COALESCE(rt.bolum, 'kaynak') = ?"
        params.append(bolum)
    if lokasyon:
        sql += " AND COALESCE(rt.lokasyon, 'TK2') = ?"
        params.append(lokasyon)
    sql += " GROUP BY rt.id ORDER BY (rt.oncelik IS NULL), rt.oncelik ASC, rt.olusturma_tarihi DESC"
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

def _oncelik_clamp(c, bolum, yeni_oncelik, eski_oncelik=None, exclude_id=None, lokasyon=None):
    """Yeni öncelik değerini, o bölüm(+lokasyon)daki mevcut öncelikli kayıt sayısına göre clamp eder.

    INSERT (eski_oncelik=None): max = mevcut_sayi + 1 (yeni eklenecek)
    UPDATE (eski var, yeni var): max = mevcut_sayi (toplam değişmiyor)
    UPDATE (eski None, yeni var): max = mevcut_sayi + 1 (öncelik atıyor, +1 olur)

    lokasyon verilirse sayım o tesise kapanır (TK1/TK2 launch sıraları karışmasın).
    None döndürürse clamp gerekmedi (zaten geçerli) veya yeni_oncelik None.
    """
    if yeni_oncelik is None:
        return None
    where_excl = ' AND id != ?' if exclude_id is not None else ''
    excl_args = (exclude_id,) if exclude_id is not None else ()
    lok_sql = " AND COALESCE(lokasyon,'TK2') = ?" if lokasyon else ''
    lok_args = (lokasyon,) if lokasyon else ()
    row = c.execute(f'''
        SELECT COUNT(*) as cnt FROM referans_takip
        WHERE COALESCE(bolum, 'kaynak') = ? AND oncelik IS NOT NULL
        {lok_sql}{where_excl}
    ''', (bolum,) + lok_args + excl_args).fetchone()
    mevcut_sayi = row['cnt'] if row else 0
    # Bu kayıt da öncelikli olacak: max = mevcut_sayi + 1 (kendisi de dahil)
    max_izinli = mevcut_sayi + 1
    if yeni_oncelik > max_izinli:
        return max_izinli
    if yeni_oncelik < 1:
        return 1
    return yeni_oncelik


def _oncelik_kaydir(c, bolum, eski_oncelik, yeni_oncelik, exclude_id=None, lokasyon=None):
    """Bir bölüm kaydının öncelik geçişinde diğer satırları doğru yönde kaydırır.
    Tüm bölümler için çalışır (kaynak/montaj/metal).

    eski_oncelik: kaydın önceki değeri (None = öncesinde öncelik yoktu / yeni kayıt)
    yeni_oncelik: kaydın yeni değeri (None = öncelik kaldırılıyor)
    exclude_id:   bu id'li satıra dokunma (PATCH için kendisini hariç tut)
    lokasyon:     verilirse kaydırma o tesise kapanır (TK1/TK2 launch sıraları karışmasın)
    """
    # Değişiklik yok
    if eski_oncelik == yeni_oncelik:
        return

    where_excl = ' AND id != ?' if exclude_id is not None else ''
    excl_args = (exclude_id,) if exclude_id is not None else ()
    lok_sql = " AND COALESCE(lokasyon,'TK2') = ?" if lokasyon else ''
    lok_args = (lokasyon,) if lokasyon else ()

    # 1) Yeni öncelik atanıyor (önceden yoktu): >= yeni olanları +1 kaydır
    if eski_oncelik is None and yeni_oncelik is not None:
        c.execute(f'''
            UPDATE referans_takip
            SET oncelik = oncelik + 1
            WHERE COALESCE(bolum, 'kaynak') = ?
              AND oncelik IS NOT NULL AND oncelik >= ?
              {lok_sql}{where_excl}
        ''', (bolum, yeni_oncelik) + lok_args + excl_args)

    # 2) Öncelik kaldırılıyor: > eski olanları -1 yukarı çek (boşluğu kapat)
    elif eski_oncelik is not None and yeni_oncelik is None:
        c.execute(f'''
            UPDATE referans_takip
            SET oncelik = oncelik - 1
            WHERE COALESCE(bolum, 'kaynak') = ?
              AND oncelik IS NOT NULL AND oncelik > ?
              {lok_sql}{where_excl}
        ''', (bolum, eski_oncelik) + lok_args + excl_args)

    # 3) Yukarı taşıma (önem artıyor: eski > yeni, ör. 4 -> 2)
    elif yeni_oncelik < eski_oncelik:
        c.execute(f'''
            UPDATE referans_takip
            SET oncelik = oncelik + 1
            WHERE COALESCE(bolum, 'kaynak') = ?
              AND oncelik IS NOT NULL AND oncelik >= ? AND oncelik < ?
              {lok_sql}{where_excl}
        ''', (bolum, yeni_oncelik, eski_oncelik) + lok_args + excl_args)

    # 4) Aşağı taşıma (önem azalıyor: eski < yeni, ör. 1 -> 3)
    else:  # yeni_oncelik > eski_oncelik
        c.execute(f'''
            UPDATE referans_takip
            SET oncelik = oncelik - 1
            WHERE COALESCE(bolum, 'kaynak') = ?
              AND oncelik IS NOT NULL AND oncelik > ? AND oncelik <= ?
              {lok_sql}{where_excl}
        ''', (bolum, eski_oncelik, yeni_oncelik) + lok_args + excl_args)


@app.route('/api/referans_takip', methods=['POST'])
@panel_gerekli(izin='is-yonetimi')
def referans_takip_ekle():
    try:
        data = request.get_json() or {}
        ref = data.get('referans_kodu', '').strip()
        if not ref:
            return jsonify({'error': 'referans_kodu zorunludur'}), 400

        bolum = (data.get('bolum') or 'kaynak').strip()
        if bolum not in GECERLI_BOLUMLER:
            bolum = 'kaynak'
        lokasyon = (data.get('lokasyon') or request.args.get('lokasyon') or 'TK2').strip() or 'TK2'

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
        oncelik = _oncelik_clamp(c, bolum, oncelik, lokasyon=lokasyon)
        # Yeni kayıt: eski_oncelik=None, yeni_oncelik girildiyse mevcut >=N olanları +1
        _oncelik_kaydir(c, bolum, None, oncelik, lokasyon=lokasyon)
        durum_init = data.get('durum', 'launch_alinacak')
        # Eger referans dogrudan 'uretimde' durumda ekleniyorsa, pilot sayac sifirlama
        # zamanini su an olarak isaretle (yeni referans secimi = sayac sifir)
        uretime_baslama_init = "datetime('now','localtime')" if durum_init == 'uretimde' else None
        if uretime_baslama_init:
            c.execute(f'''
                INSERT INTO referans_takip (referans_kodu, hedef_adet, aciklama, durum, olusturan, robot_no, istasyon, bolum, oncelik, lokasyon, uretime_baslama_ts)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, {uretime_baslama_init})
            ''', (ref, int(data.get('hedef_adet', 0)), data.get('aciklama', ''),
                  durum_init, data.get('olusturan', ''),
                  data.get('robot_no', ''), int(data.get('istasyon', 0)), bolum, oncelik, lokasyon))
        else:
            c.execute('''
                INSERT INTO referans_takip (referans_kodu, hedef_adet, aciklama, durum, olusturan, robot_no, istasyon, bolum, oncelik, lokasyon)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (ref, int(data.get('hedef_adet', 0)), data.get('aciklama', ''),
                  durum_init, data.get('olusturan', ''),
                  data.get('robot_no', ''), int(data.get('istasyon', 0)), bolum, oncelik, lokasyon))
        conn.commit()

        # PUSH bildirim — bölümün sorumlu operatörlerinin telefonuna (arka planda).
        # LOKASYON: BILDIRIM_ALICILARI TK2 (ana fabrika) sorumluları — TK1 launch'ı
        # TK2 montaj sorumlusuna GITMESIN (ayrı fabrika). TK1 için alıcı tanımlanınca
        # burada (bolum, lokasyon) anahtarlı yapıya geçilir.
        alicilar = (BILDIRIM_ALICILARI.get(bolum) or []) if lokasyon == 'TK2' else []
        if alicilar:
            parcalar = [ref]
            if data.get('robot_no'):
                parcalar.append(str(data.get('robot_no')))
            if data.get('istasyon'):
                parcalar.append(f"İst.{data.get('istasyon')}")
            if data.get('hedef_adet'):
                parcalar.append(f"{data.get('hedef_adet')} adet")
            if data.get('olusturan'):
                parcalar.append(f"Ekleyen: {data.get('olusturan')}")
            bolum_ad = BOLUM_AD.get(bolum, bolum)
            _push_gonder_async(alicilar, f'🚀 Yeni Launch — {bolum_ad}', ' · '.join(parcalar))

        return jsonify({'basarili': True}), 201
    except Exception as e:
        print(f"HATA (referans_takip_ekle): {traceback.format_exc()}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/referans_takip/<int:id>', methods=['PATCH'])
@panel_gerekli(izin='is-yonetimi')
def referans_takip_guncelle(id):
    data = request.get_json() or {}
    conn = get_db()
    c = conn.cursor()

    # Öncelik değişiyorsa shift mantığı çalıştır (tüm bölümler için)
    if 'oncelik' in data:
        mevcut = c.execute("SELECT bolum, oncelik, COALESCE(lokasyon,'TK2') as lokasyon FROM referans_takip WHERE id=?", (id,)).fetchone()
        if mevcut:
            bolum_kayit = (mevcut['bolum'] or 'kaynak')
            lok_kayit = mevcut['lokasyon'] or 'TK2'
            eski_onc = mevcut['oncelik']  # None olabilir
            yeni_raw = data.get('oncelik')
            yeni_onc = None
            if yeni_raw not in (None, '', 0, '0'):
                try:
                    yeni_onc = int(yeni_raw)
                    if yeni_onc < 1: yeni_onc = None
                except (TypeError, ValueError):
                    yeni_onc = None
            # Clamp: girilen sayı listedeki ref sayısını aşıyorsa düşür (o tesise kapalı)
            yeni_onc = _oncelik_clamp(c, bolum_kayit, yeni_onc, eski_oncelik=eski_onc, exclude_id=id, lokasyon=lok_kayit)
            _oncelik_kaydir(c, bolum_kayit, eski_onc, yeni_onc, exclude_id=id, lokasyon=lok_kayit)
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
    if 'depo_notu' in data:
        fields.append('depo_notu=?'); vals.append((data.get('depo_notu') or '').strip())
    if 'hedef_adet' in data:
        try:
            fields.append('hedef_adet=?'); vals.append(int(data.get('hedef_adet') or 0))
        except (ValueError, TypeError):
            conn.close()
            return jsonify({'error': f"hedef_adet sayısal olmalı: {data.get('hedef_adet')!r}"}), 400
    if 'referans_kodu' in data:
        fields.append('referans_kodu=?'); vals.append(data['referans_kodu'])
    if 'robot_no' in data:
        fields.append('robot_no=?'); vals.append(data['robot_no'])
    if 'istasyon' in data:
        try:
            fields.append('istasyon=?'); vals.append(int(data.get('istasyon') or 0))
        except (ValueError, TypeError):
            conn.close()
            return jsonify({'error': f"istasyon sayısal olmalı: {data.get('istasyon')!r}"}), 400
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
@panel_gerekli(izin='is-yonetimi')
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
    lokasyon = request.args.get('lokasyon', '')
    conn = get_db()
    c = conn.cursor()

    # Bölüm + lokasyon filtresi
    bolum_sart = ''
    bolum_params = [bugun]
    if bolum:
        bolum_sart += ' AND v.bolum = ?'
        bolum_params.append(bolum)
    if lokasyon:
        bolum_sart += " AND COALESCE(v.lokasyon,'TK2') = ?"
        bolum_params.append(lokasyon)

    # Bugünkü tüm vardiyalar
    vardiyalar = c.execute(
        f'SELECT * FROM vardiyalar v WHERE v.tarih = ?{bolum_sart} ORDER BY v.baslangic_saati',
        bolum_params
    ).fetchall()

    # Otomatik molalar (çay/yemek) — saati gelenleri bu vardiyalara işle (idempotent, canlı)
    for _v in vardiyalar:
        _otomatik_mola_uygula(c, _v)
    conn.commit()

    # SAYAÇ SENKRONU — açık vardiyaların TÜM auto kayıtları (ESP32 + test cihazı).
    # Andon 30sn'de bir çağrılır: canlı makine adetleri üretim/performans hesabına
    # da anlık yansır (eskiden yalnız operatör mobili açıkken tazeleniyordu →
    # pilot rozeti canlı ama adet/performans bayat görünüyordu).
    _acik_auto_vardiyalari_senkronla(conn, bolum=bolum or None, lokasyon=lokasyon or None)

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
    # Açık vardiyalarda "o ana kadar geçen süre" (canlı), kapalılarda gerçek süre
    toplam_planli_dk = sum(_vardiya_efektif_dk(v) for v in vardiyalar)
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
        if v['durum'] == 'kapali':   # 'acik' VEYA 'aktif' (yeniden acilan) goster; sadece 'kapali' gizle
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
                # test_cihaz_id: eski DB'de kolon olmayabilir — güvenli erişim
                try:
                    _tc = u['test_cihaz_id']
                except (KeyError, IndexError):
                    _tc = None
                row = {
                    'ref': u['referans_kodu'],
                    'launch': u['launch_adet'] or 0,
                    'tamamlandi': 1 if u['tamamlandi'] else 0,
                    'teyit': 0,        # aşağıda ref_durum_map'ten doldurulacak
                    'suresiz': 0,
                    'ct': u['cycle_time_sn'] or 0,   # metal andon CT gösterimi (referans_listesi hedefi ile override edilir)
                    'ok': u['ok_adet'] or 0,
                    # Test cihazı sayaç kaynağı — andon ref satırında canlı adet rozeti göstermek için
                    'test_cihaz': 1 if _tc else 0,
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
    elif bolum == 'metal':
        # Metal andon: her referans satırına güncel hedef cycle time'ı yaz.
        # referans_listesi hedefi varsa o önceliklidir (panelden güncellenince andon yansır);
        # yoksa row['ct'] zaten kayıttaki cycle_time_sn (fallback) olarak kalır.
        ref_ct_map = {}
        for r in c.execute(
            "SELECT referans_kodu, COALESCE(hedef_cycle_time_sn,0) as ct "
            "FROM referans_listesi WHERE COALESCE(bolum,'kaynak')='metal'"
        ).fetchall():
            norm = str(r['referans_kodu'] or '').upper().replace(' ', '')
            ref_ct_map[norm] = r['ct']
        for it in aktif_vardiyalar:
            for koleksiyon in ('istasyon_1', 'istasyon_2', 'diger'):
                for row in it[koleksiyon]:
                    norm = str(row.get('ref') or '').upper().replace(' ', '')
                    hedef = ref_ct_map.get(norm)
                    if hedef and hedef > 0:
                        row['ct'] = hedef

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
        if v['durum'] == 'kapali':   # 'acik' VEYA 'aktif' (yeniden acilan) goster; sadece 'kapali' gizle
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

    # Referans takip listesi — bölüm VE lokasyona göre filtrelenir
    # (TK1 andonunda TK2 montaj iş emirleri görünmesin / tersi)
    # GROUP BY rt.id — fan-out koruması (normalize-duplike referans satırları JOIN'de
    # tek iş emrini çoğaltmasın; bkz. /api/referans_takip'teki aynı koruma)
    rt_sql = '''
        SELECT rt.*, MAX(rl.hedef_cycle_time_sn) as hedef_cycle_time_sn
        FROM referans_takip rt
        LEFT JOIN referans_listesi rl ON REPLACE(rt.referans_kodu, ' ', '') = REPLACE(rl.referans_kodu, ' ', '')
            AND COALESCE(rl.lokasyon,'TK2') = COALESCE(rt.lokasyon,'TK2')
        WHERE 1=1
    '''
    rt_params = []
    if bolum:
        rt_sql += " AND COALESCE(rt.bolum, 'kaynak') = ?"
        rt_params.append(bolum)
    if lokasyon:
        rt_sql += " AND COALESCE(rt.lokasyon, 'TK2') = ?"
        rt_params.append(lokasyon)
    rt_sql += " GROUP BY rt.id ORDER BY rt.olusturma_tarihi DESC"
    rt_rows = c.execute(rt_sql, rt_params).fetchall()
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
    if bolum not in GECERLI_BOLUMLER:
        bolum = 'kaynak'
    # Lokasyon verilmezse TK2 varsay: TK1 vardiyaları da bolum='montaj' olduğundan
    # filtresiz sorgu TK1 operatörlerini TK2 montaj raporuna karıştırıyordu.
    # (TK1 raporu isteyen istemciler — mobil/rapor sayfası — lokasyon=TK1 gönderir.)
    lokasyon = (request.args.get('lokasyon') or 'TK2').strip() or 'TK2'

    if not tarih or not vardiya_turu:
        return jsonify({'hata': 'tarih ve vardiya parametreleri zorunludur'}), 400

    try:
        conn = get_db()

        # Bugünün detayı isteniyorsa auto sayaçları tazele (canlı adetlerle raporla)
        if tarih == date.today().isoformat():
            _acik_auto_vardiyalari_senkronla(conn, bolum or None, lokasyon or None)

        # LAZER: makine vardiyada değil ÜRETİM KAYDINDA seçilir (istasyon 1..6) —
        # rapor makine bazında kırılır, gruplamaya istasyon eklenir.
        lazer_modu = (bolum == 'lazer')
        ist_kolon = ", COALESCE(u.istasyon, 0) as istasyon" if lazer_modu else ""
        sql = f"""
            SELECT
                v.robot_no,
                v.operator_adi,
                u.referans_kodu,
                SUM(u.ok_adet)    as ok_toplam,
                SUM(u.nok_adet)   as nok_toplam,
                SUM(u.tamir_adet) as tamir_toplam{ist_kolon}
            FROM vardiyalar v
            LEFT JOIN uretim_kayitlari u ON v.id = u.vardiya_id
            WHERE v.tarih = ?
              AND UPPER(REPLACE(REPLACE(v.vardiya_turu,'ü','u'),'Ü','U')) =
                  UPPER(REPLACE(REPLACE(?,'ü','u'),'Ü','U'))
              AND COALESCE(v.bolum, 'kaynak') = ?
        """
        params = [tarih, vardiya_turu, bolum]
        if lokasyon:
            sql += " AND COALESCE(v.lokasyon, 'TK2') = ?"
            params.append(lokasyon)
        if lazer_modu:
            sql += " GROUP BY v.robot_no, v.operator_adi, u.referans_kodu, COALESCE(u.istasyon, 0) ORDER BY istasyon, v.operator_adi"
        else:
            sql += " GROUP BY v.robot_no, v.operator_adi, u.referans_kodu ORDER BY v.robot_no, v.operator_adi"
        rows = conn.execute(sql, params).fetchall()
        conn.close()

        # Bolume gore robot/hat/makine sirasi
        if bolum == 'kaynak':
            robot_listesi = ['1','2','3','4','5','6','7','8','9','M']
        elif bolum == 'metal':
            robot_listesi = ['300T', '400T', '550T', 'Şerit Testere']
        elif lazer_modu:
            robot_listesi = [f'Lazer {i}' for i in range(1, 7)]
        else:  # montaj / isleme / pres
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
            elif lazer_modu:
                # Kayıt hangi makinedeyse o gruba (istasyon 1..6 → 'Lazer N');
                # makinesiz eski/boş kayıtlar genel 'Lazer' grubuna düşer.
                ist = int(row['istasyon'] or 0)
                target_r = f'Lazer {ist}' if ist > 0 else 'Lazer'
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
    # Lokasyon verilmezse TK2 varsay (gunluk_rapor_detay ile aynı gerekçe:
    # TK1 vardiya türleri TK2 rapor dropdown'una karışmasın).
    lokasyon = (request.args.get('lokasyon') or 'TK2').strip() or 'TK2'
    if not tarih:
        return jsonify({'hata': 'tarih zorunludur'}), 400
    try:
        conn = get_db()
        sql = "SELECT DISTINCT vardiya_turu FROM vardiyalar WHERE tarih = ?"
        params = [tarih]
        if bolum:
            sql += " AND COALESCE(bolum, 'kaynak') = ?"
            params.append(bolum)
        if lokasyon:
            sql += " AND COALESCE(lokasyon, 'TK2') = ?"
            params.append(lokasyon)
        sql += " ORDER BY vardiya_turu"
        rows = conn.execute(sql, params).fetchall()
        conn.close()
        return jsonify([r['vardiya_turu'] for r in rows])
    except Exception as e:
        return jsonify({'hata': str(e)}), 500


# ─────────────────────────────────────────────────────────────
# AS400 (ERP) TEYIT — günün üretimi ↔ açık launch eşlemesi
# Kaynak: as400/launch_esle.py → canlı TKC0301F.XPRO90 (pyodbc; şifre Windows
# Kimlik Bilgileri Yöneticisi'nde). Salt-okunur; ERP'ye hiçbir şey yazmaz.
# ─────────────────────────────────────────────────────────────
@app.route('/api/as400/teyit_listesi', methods=['GET'])
@panel_gerekli(izin='as400-teyit')
def as400_teyit_listesi():
    """Sabah teyit KUYRUĞU. İki mod:
      ?gunler=3[&bugun=1]  → son 3 tamamlanmış günün (dün, önceki, ...) listesi;
                             bugun=1 ise bugünü de ekler (varsayılan mod)
      ?tarih=YYYY-MM-DD    → tek gün (geçmiş inceleme)
    AS400 okumaları tek sefer yapılır (esle_coklu). Yanıt her gün için
    kategorili liste + son robot gönderimleri (as400_teyit_log) içerir."""
    tarih = (request.args.get('tarih') or '').strip()
    if tarih:
        tarihler = [tarih]
    else:
        try:
            gunler = max(1, min(7, int(request.args.get('gunler') or 3)))
        except ValueError:
            gunler = 3
        tarihler = []
        if (request.args.get('bugun') or '') in ('1', 'true'):
            tarihler.append(date.today().isoformat())
        tarihler += [(date.today() - timedelta(days=i)).isoformat() for i in range(1, gunler + 1)]
    try:
        import sys as _sys
        _d = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'as400')
        if _d not in _sys.path:
            _sys.path.insert(0, _d)
        import launch_esle as _le
        coklu = _le.esle_coklu(tarihler)
    except Exception as e:
        # AS400 kapalı/şifre yok/ağ hatası → paneli kırmadan anlaşılır mesaj
        return jsonify({'hata': f'AS400 sorgusu başarısız: {e}'}), 502
    conn = get_db()
    son_gonderimler = [dict(r) for r in conn.execute(
        "SELECT id, created_at, uretim_tarihi, yil, launch_no, referans, adet, sonuc, mesaj, olusturan "
        "FROM as400_teyit_log ORDER BY id DESC LIMIT 15").fetchall()]

    # COP verilmiş hurdalar (2026-07-24): "♻ Hurda → COP" bölümü zaten COP'lanmış
    # satırları TEKRAR göstermesin (sayfa yenilenince mükerrer görünüyordu). Anahtar =
    # "uretim_tarihi|article" (article = COP log'daki launch_no). Dedup ile aynı mantık.
    cop_verildi = {}
    try:
        for _cr in conn.execute(
            "SELECT uretim_tarihi, launch_no, SUM(adet) tp FROM as400_teyit_log "
            "WHERE yil='CO' AND sonuc='ok' GROUP BY uretim_tarihi, launch_no").fetchall():
            cop_verildi[f"{_cr['uretim_tarihi']}|{_cr['launch_no']}"] = _cr['tp']
    except Exception as _e:
        print(f'[teyit_listesi] cop_verildi atlandı: {_e}')

    # İŞARETLER (2026-07-20): günlük 'kontrol edildi' + kalıcı 'gerek_yok'.
    # Backend kanonik eşleştirir → her satıra 'isaret'/'kalici_haric' enjekte edilir
    # (frontend'in normalize etmesine gerek kalmaz). Tablo yoksa (eski DB) sessiz geç.
    try:
        gun_isaret = {}
        for t in tarihler:
            gun_isaret[t] = {r['referans']: {'durum': r['durum'], 'aciklama': r['aciklama']}
                             for r in conn.execute(
                                 "SELECT referans, durum, aciklama FROM as400_teyit_isaret "
                                 "WHERE kapsam='gun' AND uretim_tarihi=?", (t,)).fetchall()}
        kalici_isaret = {r['referans']: r['aciklama'] for r in conn.execute(
            "SELECT referans, aciklama FROM as400_teyit_isaret WHERE kapsam='kalici' AND durum='gerek_yok'").fetchall()}
        for t in tarihler:
            for satirlar in coklu[t].values():
                for r in satirlar:
                    kn = _le.kanonik(r.get('referans', ''))
                    if kn in gun_isaret[t]:
                        r['isaret'] = gun_isaret[t][kn]
                    if kn in kalici_isaret:
                        r['kalici_haric'] = True
    except Exception as e:
        print(f'[teyit_listesi] işaret enjeksiyonu atlandı: {e}')

    return jsonify({
        'gunler': [{'tarih': t,
                    'ozet': {k: len(v) for k, v in coklu[t].items()},
                    'liste': coklu[t]} for t in tarihler],
        'son_gonderimler': son_gonderimler,
        'cop_verildi': cop_verildi,
    })


@app.route('/api/as400/teyit_isaret', methods=['POST'])
@panel_gerekli(izin='as400-teyit')
def as400_teyit_isaret():
    """Teyit ekranında bir satırı işaretle / işareti kaldır.
    Body: {kapsam:'gun'|'kalici', durum?:'kontrol'|'gerek_yok', referans, bolum?,
           uretim_tarihi?, aciklama?, kaldir?}
    kapsam='kalici' → referans HER GÜN teyit dışı (HARIC). kapsam='gun' → o güne özel.
    durum='teyit_ver' (2026-07-27, 4536B vakası): satır Dikkat/varyant/ara-op'a düşse
    bile 'kod DOĞRU, teyit ver' demektir → CFI kuyruğuna zorlanır (robot + oto koşu).
    Referans launch_esle.kanonik ile normalize saklanır (eşleşme tutarlı)."""
    data = request.get_json(silent=True) or {}
    kapsam = (data.get('kapsam') or 'gun').strip()
    durum = (data.get('durum') or '').strip()
    referans = (data.get('referans') or '').strip()
    bolum = (data.get('bolum') or '').strip()
    u_tarih = (data.get('uretim_tarihi') or '').strip()
    aciklama = (data.get('aciklama') or '').strip()
    kaldir = bool(data.get('kaldir'))
    if kapsam not in ('gun', 'kalici') or not referans:
        return jsonify({'hata': 'kapsam ve referans zorunlu'}), 400
    if kapsam == 'kalici':
        u_tarih = ''
        durum = 'gerek_yok'
    else:
        if not u_tarih:
            return jsonify({'hata': 'günlük işaret için uretim_tarihi zorunlu'}), 400
        if durum not in ('kontrol', 'gerek_yok', 'teyit_ver'):
            durum = 'kontrol'
    try:
        import sys as _sys
        _d = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'as400')
        if _d not in _sys.path:
            _sys.path.insert(0, _d)
        import launch_esle as _le
        ref_norm = _le.kanonik(referans)
    except Exception:
        ref_norm = referans.upper().replace(' ', '')
    conn = get_db()
    kullanici = g.panel_ku['kullanici_adi']
    try:
        if kaldir:
            conn.execute(
                "DELETE FROM as400_teyit_isaret WHERE kapsam=? AND uretim_tarihi=? AND bolum=? AND referans=?",
                (kapsam, u_tarih, bolum, ref_norm))
        else:
            conn.execute(
                "INSERT INTO as400_teyit_isaret (kapsam, uretim_tarihi, bolum, referans, durum, aciklama, olusturan) "
                "VALUES (?,?,?,?,?,?,?) "
                "ON CONFLICT(kapsam, uretim_tarihi, bolum, referans) "
                "DO UPDATE SET durum=excluded.durum, aciklama=excluded.aciklama, olusturan=excluded.olusturan",
                (kapsam, u_tarih, bolum, ref_norm, durum, aciklama, kullanici))
        conn.commit()
        return jsonify({'basarili': True, 'kapsam': kapsam, 'durum': durum, 'referans': ref_norm, 'kaldirildi': kaldir})
    except Exception as e:
        conn.rollback()
        return jsonify({'hata': str(e)}), 400


# ── AS400 teyit GÖNDERİMİ (yazma) ──
# Session B'yi tek seferde tek istek sürebilir (fiziksel tek ekran) → global kilit.
import threading as _threading
_AS400_KILIT = _threading.Lock()
# Nazik durdurma bayrağı: kullanıcı "Durdur"a basınca set edilir. Döngü HER satır
# başında kontrol eder → set ise kalan satırlar 'atlandi'. Çalışan cscript ASLA
# kill edilmez (record-lock riski); mevcut launch biter, SONRAKİLERE geçilmez.
_AS400_DURDUR = _threading.Event()

# ── TEYIT-AGENT KÖPRÜSÜ (2026-07-24) ──
# Server'da app NSSM SERVİSİ olarak Session 0'da koşar; PCOMM (pcsws.exe) ise RDP
# kullanıcı oturumundadır. Session 0'dan başlatılan cscript PCOMM oturumlarını GÖREMEZ
# (Windows oturum izolasyonu) → robot, RDP oturumunda koşan as400/teyit_agent.py'ye
# HTTP ile devredilir (127.0.0.1:5010). Agent yoksa (laptop/interaktif) yerel subprocess
# — config dosyası YOK, tespit otomatik (localhost probe; agent kapalıysa anında reddedilir).
_TEYIT_AGENT_URL = 'http://127.0.0.1:5010'


def _teyit_agent_var():
    try:
        import requests
        r = requests.get(_TEYIT_AGENT_URL + '/durum', timeout=1.5)
        return r.status_code == 200 and (r.json() or {}).get('agent') == 'cofle-teyit'
    except Exception:
        return False


def _as400_robot_calistir(script_dosya, args, timeout_sn):
    """Robot cscript'ini çalıştır (agent varsa onda, yoksa yerel).
    Döner: (cikti:str, hata:str|None) — hata doluysa satır 'hata' olarak loglanmalı.
    KRİTİK: agent'a istek GİTTİKTEN sonra bağlantı koparsa yerel TEKRAR DENENMEZ
    (robot agent'ta hâlâ koşuyor olabilir → çift giriş riski)."""
    args = [str(a) for a in args]
    if _teyit_agent_var():
        try:
            import requests
            r = requests.post(_TEYIT_AGENT_URL + '/calistir',
                              json={'script': script_dosya, 'args': args, 'timeout': timeout_sn},
                              timeout=(3, timeout_sn + 15))
            if r.status_code == 200:
                return (r.json() or {}).get('cikti', ''), None
            if r.status_code == 504:
                return '', f'Robot zaman aşımı ({timeout_sn} sn) — sıradakine geçildi'
            return '', f'Teyit-agent hatası (HTTP {r.status_code}): {(r.json() or {}).get("hata", "")}'
        except Exception as e:
            return '', f'Teyit-agent bağlantısı koptu ({e}) — robot durumu belirsiz, Session B\'yi kontrol edin'
    import subprocess
    script = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'as400', script_dosya)
    try:
        pr = subprocess.run([r'C:\Windows\SysWOW64\cscript.exe', '//nologo', script] + args,
                            capture_output=True, timeout=timeout_sn)
        return (pr.stdout or b'').decode('cp1254', errors='replace'), None
    except subprocess.TimeoutExpired:
        # Timeout → cscript KILL edilir → yarım transaction + record-lock riski;
        # sonraki satıra devam edilir (sonraki robot G0'da record-lock'u kurtarır).
        return '', f'Robot zaman aşımı ({timeout_sn} sn) — sıradakine geçildi'


def _as400_launch_durum(yil, no):
    """BPROF0'dan launch'ın (teyitli Q0QTRI, durum Q0AVAN) — bağımsız doğrulama.
    Döner: (teyitli:float|None, durum:int|None)."""
    import sys as _sys
    _d = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'as400')
    if _d not in _sys.path:
        _sys.path.insert(0, _d)
    import keyring, pyodbc
    import as400_config as _cfg
    pw = _cfg.sifre_al()
    cn = pyodbc.connect(_cfg.baglanti_dizesi(pw), timeout=15, autocommit=True)
    try:
        r = cn.cursor().execute(
            "SELECT Q0QTRI, Q0AVAN FROM tkc0301F.BPROF0 WHERE Q0RED2=? AND Q0RENU=?",
            (int(yil), int(no))).fetchone()
        if not r:
            return None, None
        try:
            durum = int(str(r[1]).strip())
        except (TypeError, ValueError):
            durum = None
        return float(r[0]), durum
    finally:
        cn.close()


def _article_gecersiz_mi(article):
    """ERP article master kontrolü — robot ekrana yazmadan ÖNCE (2026-07-30).
    Döner: None (kod geçerli / doğrulanamadı) | reddetme mesajı (str).

    Master OKUNAMAZSA None döner: bilinmiyorken meşru teyidi bloklamak yanlış
    yön. Bu bir "bozuk kodu erken yakala" katmanı, tek fren değil."""
    try:
        import sys as _sys
        _d = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'as400')
        if _d not in _sys.path:
            _sys.path.insert(0, _d)
        import launch_esle as _la
        var = _la.article_tanimli([article])
        if var is None:          # sorgu/bağlantı hatası — karar verme
            return None
        if _la.kanonik(article) in var:
            return None
        return (f'ERP\'de böyle bir article YOK: "{article}" — operatör kodu yanlış '
                f'yazmış olabilir (örn. kod alanına operasyon bilgisi karışmış). '
                f'Robota gönderilmedi; üretim kaydındaki referans kodunu düzeltin. '
                f'NOT: bu satır gönderilseydi ekran kilitlenip AYNI KOŞUDAKİ diğer '
                f'referansların teyidi de verilemezdi.')
    except Exception as e:
        print(f'[article] kontrol yapılamadı ({article}): {e}')
    return None


def _kapasite_reddi(conn, referans, uretim_tarihi, adet):
    """SON SAVUNMA (2026-07-30) — robot çalıştırmadan hemen önce, DB'den BAĞIMSIZ
    kapasite kontrolü. Döner: None (sorun yok) | reddetme mesajı (str).

    NEDEN AYRI HESAP: frontend listesi bayat olabilir, kuyruk elle POST edilebilir,
    /api/as400/teyit_gonder doğrudan çağrılabilir. Hayali stok girişini ENGELLEYEN
    asıl yer burasıdır; UI'daki Dikkat kovası yalnızca erken uyarıdır.

    Cycle time tanımsız (0/NULL) → None döner, kontrol YOK (kullanıcı kararı).
    'teyit_ver' işareti kullanıcının açık onayı → kontrol atlanır.
    Hesap hata verirse None döner: belirsizken meşru teyidi bloklamak yanlış yön
    (asıl fren zaten operatör kaydının kendisi, bu ek bir emniyet katmanı)."""
    try:
        import sys as _sys
        _d = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'as400')
        if _d not in _sys.path:
            _sys.path.insert(0, _d)
        import launch_esle as _lk
        kanon = _lk.kanonik(referans)
        # Kullanıcı bu satırı açıkça onayladıysa geç
        onay = conn.execute(
            "SELECT 1 FROM as400_teyit_isaret WHERE kapsam='gun' AND uretim_tarihi=? "
            "AND referans=? AND durum='teyit_ver' LIMIT 1", (uretim_tarihi, kanon)).fetchone()
        if onay:
            return None
        # O gün o referansı üreten vardiyalar + cycle time (bölüm/lokasyon eşleşmeli)
        # ct SKALAR alt-sorgu (LEFT JOIN DEĞİL): normalize aynı koda düşen birden
        # fazla referans satırı varsa JOIN satır çoğaltır (bkz. launch_esle
        # gun_uretimi'ndeki 94.LTK.340 vakası). Burada adet toplanmadığı için
        # zararı sınırlıydı ama aynı tuzağı bırakmıyoruz.
        rows = conn.execute("""
            SELECT (SELECT MAX(rl.hedef_cycle_time_sn) FROM referans_listesi rl
                     WHERE UPPER(REPLACE(rl.referans_kodu,' ','')) = UPPER(REPLACE(u.referans_kodu,' ',''))
                       AND COALESCE(rl.bolum,'kaynak') = COALESCE(v.bolum,'kaynak')
                       AND COALESCE(rl.lokasyon,'TK2') = COALESCE(v.lokasyon,'TK2')) ct,
                   v.id vid,
                   COALESCE(v.toplam_sure_dk,0) sure_dk,
                   COALESCE(v.baslangic_saati,'') b, COALESCE(v.bitis_saati,'') e
            FROM uretim_kayitlari u JOIN vardiyalar v ON v.id = u.vardiya_id
            WHERE v.tarih = ?
              AND UPPER(REPLACE(u.referans_kodu,' ','')) = UPPER(REPLACE(?,' ',''))
        """, (uretim_tarihi, referans)).fetchall()
        if not rows:
            return None
        ct = 0.0
        sureler = {}
        for r in rows:
            try:
                c = float(r['ct'] or 0)
            except (TypeError, ValueError):
                c = 0.0
            if c > 0 and ct <= 0:
                ct = c
            sureler[r['vid']] = _lk._vardiya_suresi_dk(r['sure_dk'], r['b'], r['e'])
        if ct <= 0:
            return None                       # referans tanımsız → kontrol yok
        sure_dk = sum(v for v in sureler.values() if v and v > 0) or _lk.KAPASITE_VARSAYILAN_DK
        teorik = (sure_dk * 60.0) / ct
        azami = teorik * _lk.KAPASITE_TOLERANS
        if adet > azami:
            gerek = (adet * ct) / 3600.0
            return (f'KAPASİTE AŞIMI — {adet} adet × {ct:g} sn = {gerek:.1f} saat gerekir, '
                    f'vardiya {sure_dk / 60.0:.1f} saat (makul üst sınır ~{int(azami)} adet). '
                    f'Operatör kaydı hatalı olabilir; doğruysa panelden "➕ teyit ver" ile onaylayın.')
    except Exception as e:
        print(f'[kapasite] kontrol yapılamadı ({referans} {uretim_tarihi}): {e}')
    return None


def _teyit_gonder_calistir(conn, satirlar, kullanici, varsayilan_tarih='', zorla=False):
    """Launch teyit satırlarını robotla işler — endpoint VE 17:10 oto koşusu ortak
    çekirdeği. ÇAĞIRAN _AS400_KILIT'i tutuyor olmalı. Döner: sonuclar listesi.
    Her satır BAĞIMSIZ denenir — bir launch hata/iptal verirse SONRAKİ launch'lara
    DEVAM edilir (kullanıcı 2026-07-20: hata alınca ilk sayfaya çıkıp sıradaki
    launch'tan devam et). Robot her çağrıda G0'da B'yi temizler + record-lock kurtarır."""
    sonuclar = []
    for idx, s in enumerate(satirlar):
        # NAZIK DURDURMA: kullanıcı "Durdur"a bastıysa kalan satırlara GEÇME (çalışan
        # cscript öldürülMEZ; mevcut satır zaten bittiğinde bu kontrole geliriz).
        if _AS400_DURDUR.is_set():
            for kalan in satirlar[idx:]:
                sonuclar.append({
                    'yil': str(kalan.get('yil') or '').strip(),
                    'no': str(kalan.get('no') or '').strip(),
                    'article': str(kalan.get('article') or '').strip(),
                    'referans': str(kalan.get('referans') or '').strip(),
                    'adet': kalan.get('adet'), 'sonuc': 'atlandi',
                    'mesaj': 'Kullanıcı durdurdu'})
            break
        yil = str(s.get('yil') or '').strip()
        no = str(s.get('no') or '').strip()
        article = str(s.get('article') or '').strip()
        referans = str(s.get('referans') or '').strip()
        # Üretim tarihi SATIR bazlı (kuyruk görünümü birden çok günü birlikte gönderir)
        u_tarih = (str(s.get('uretim_tarihi') or '').strip() or varsayilan_tarih)
        try:
            adet = int(s.get('adet') or 0)
        except (TypeError, ValueError):
            adet = 0
        kayit = {'yil': yil, 'no': no, 'article': article, 'referans': referans,
                 'adet': adet, 'uretim_tarihi': u_tarih}

        if not (yil.isdigit() and len(yil) == 2 and no.isdigit() and 0 < adet <= 99999 and article and u_tarih):
            sonuclar.append({**kayit, 'sonuc': 'hata', 'mesaj': 'Geçersiz satır parametresi'})
            continue
        # Mükerrer koruması 1: kendi gönderim log'umuz
        var = conn.execute(
            "SELECT id FROM as400_teyit_log WHERE uretim_tarihi=? AND yil=? AND launch_no=? AND sonuc='ok'",
            (u_tarih, yil, no)).fetchone()
        if var and not zorla:
            sonuclar.append({**kayit, 'sonuc': 'atlandi',
                             'mesaj': f'Bu üretim günü için bu launch\'a zaten gönderilmiş (log #{var["id"]})'})
            continue
        # Mükerrer koruması 1b — REFERANS+GÜN (2026-07-30, 94.LTK.487/10 olayı):
        # yukarıdaki kontrol launch numarasına bağlı; ilk launch kapanıp YENİ launch
        # açılınca (26-150830 → 26-153549) anahtar değişiyor ve aynı üretim günü
        # İKİNCİ KEZ teyit alabiliyordu. Aynı gün + aynı referans için başarılı bir
        # launch teyidi varsa dur. Meşru bölünmüş teyit için 'zorla' ile aşılabilir.
        if not zorla:
            var2 = conn.execute(
                "SELECT id, launch_no, adet FROM as400_teyit_log "
                "WHERE uretim_tarihi=? AND sonuc='ok' AND yil NOT IN ('CF','CO') "
                "AND UPPER(REPLACE(referans,' ',''))=UPPER(REPLACE(?,' ','')) LIMIT 1",
                (u_tarih, referans)).fetchone()
            if var2:
                sonuclar.append({**kayit, 'sonuc': 'atlandi',
                                 'mesaj': f'Bu üretim günü ({u_tarih}) için bu referansa zaten teyit '
                                          f'verilmiş: launch {var2["launch_no"]}, {var2["adet"]} adet '
                                          f'(log #{var2["id"]}). Launch numarası farklı olsa bile aynı '
                                          f'üretim iki kez teyit edilmemeli — gerçekten bölünmüş teyit '
                                          f'gerekiyorsa "zorla" ile gönderin.'})
                continue
        # Mükerrer koruması 2 (ASIL): ERP'nin GERÇEK teyit hareketleri —
        # operatörün ELLE girdiği teyitler de burada görünür. Frontend
        # atlansa/eski liste gönderilse bile burada yakalanır.
        if not zorla:
            try:
                import launch_esle as _le2
                hrk = _le2.teyit_hareketleri([article]).get(_le2.kanonik(article), [])
                zt, ilgili = _le2._zaten_teyitli(hrk, u_tarih, adet,
                                                 _le2.ref_uretim_gecmisi(referans, u_tarih))
            except Exception:
                zt, ilgili = None, []
            if zt == 'kesin':
                det = ', '.join(f"{h['tarih']}: {h['adet']:g}" for h in ilgili)
                sonuclar.append({**kayit, 'sonuc': 'atlandi',
                                 'mesaj': f'ERP\'de bu üretim için zaten teyit var ({det}) — mükerrer olurdu'})
                continue
            # Mükerrer koruması 3: KAPASİTE (2026-07-30) — hayali stok girişini engeller
            kap = _kapasite_reddi(conn, referans, u_tarih, adet)
            if kap:
                sonuclar.append({**kayit, 'sonuc': 'hata', 'mesaj': kap})
                continue

        bayrak = 'S' if str(s.get('bayrak') or 'A').upper() == 'S' else 'A'
        try:
            once, once_durum = _as400_launch_durum(yil, no)
        except Exception as e:
            sonuclar.append({**kayit, 'bayrak': bayrak, 'sonuc': 'hata', 'mesaj': f'Ön okuma hatası: {e}'})
            continue

        cikti, robot_hata = _as400_robot_calistir('teyit_gir.js', [yil, no, article, adet, bayrak], 150)
        if robot_hata:
            sonuclar.append({**kayit, 'bayrak': bayrak, 'sonuc': 'hata', 'mesaj': robot_hata, 'teyitli_once': once})
            continue

        robot_ok = 'SONUC=OK' in cikti
        try:
            sonra, sonra_durum = _as400_launch_durum(yil, no)
        except Exception:
            sonra, sonra_durum = None, None
        # Doğrulama: teyitli tam +adet artmalı; S ise launch KAPANMIŞ (durum 70) olmalı.
        teyit_ok = (once is not None and sonra is not None and abs(sonra - once - adet) < 0.001)
        kapanma_ok = (bayrak == 'A') or (sonra_durum == 70)
        dogrulandi = robot_ok and teyit_ok and kapanma_ok
        son_satirlar = ' | '.join(l for l in cikti.splitlines() if l.strip())[-500:]
        sonuc = 'ok' if dogrulandi else 'hata'
        if dogrulandi:
            mesaj = f'Doğrulandı: teyitli {once:g} → {sonra:g}' + (
                f' · launch KAPANDI (durum {sonra_durum})' if bayrak == 'S' else '')
            mesaj += _is_emri_dus_router(conn, referans, adet)
            # COP (hurda) BURADA GİRİLMEZ (kullanıcı 2026-07-23: robot Rientro↔07>01
            # mekik dokuyordu) — hurda ayrı /cop_gonder fazında EN SONA girilir.
        elif robot_ok and teyit_ok and not kapanma_ok:
            mesaj = f'Teyit işlendi ({once:g}→{sonra:g}) ama S ile KAPANMADI (durum {sonra_durum}) — kontrol edin'
        elif robot_ok:
            mesaj = f'Robot OK ama doğrulama tutmadı (önce {once} sonra {sonra}, durum {sonra_durum})'
        else:
            mesaj = f'Robot iptal: {son_satirlar}'
        conn.execute(
            "INSERT INTO as400_teyit_log (uretim_tarihi, yil, launch_no, referans, article, adet, bayrak, sonuc, mesaj, teyitli_once, teyitli_sonra, olusturan) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
            (u_tarih, yil, no, referans, article, adet, bayrak, sonuc, mesaj, once, sonra, kullanici))
        conn.commit()
        sonuclar.append({**kayit, 'bayrak': bayrak, 'sonuc': sonuc, 'mesaj': mesaj,
                         'teyitli_once': once, 'teyitli_sonra': sonra})
        # Hata olsa bile SONRAKİ satıra devam (durdurma YOK).

    return sonuclar


@app.route('/api/as400/teyit_gonder', methods=['POST'])
@panel_gerekli(izin='as400-teyit')
def as400_teyit_gonder():
    """Seçili satırları AS400'e teyit olarak girer (Session B ekran robotu).
    Body: {uretim_tarihi, satirlar: [{yil, no, article, referans, adet, bayrak}], zorla}
    Çekirdek _teyit_gonder_calistir'da (17:10 oto koşusuyla ORTAK). Aynı üretim
    günü + launch için ikinci gönderim reddedilir (zorla=true hariç). Her satır
    sonrası BPROF0.Q0QTRI ile bağımsız doğrulama."""
    data = request.get_json(silent=True) or {}
    varsayilan_tarih = (data.get('uretim_tarihi') or '').strip()
    satirlar = data.get('satirlar') or []
    zorla = bool(data.get('zorla'))
    if not satirlar:
        return jsonify({'hata': 'satirlar zorunlu'}), 400
    if len(satirlar) > 60:
        return jsonify({'hata': 'Tek seferde en fazla 60 satır'}), 400
    if not _AS400_KILIT.acquire(blocking=False):
        return jsonify({'hata': 'Devam eden bir AS400 gönderimi var — bitmesini bekleyin'}), 409
    try:
        _AS400_DURDUR.clear()   # yeni gönderim → önceki "durdur" bayrağını sıfırla
        sonuclar = _teyit_gonder_calistir(get_db(), satirlar, g.panel_ku['kullanici_adi'],
                                          varsayilan_tarih, zorla)
        ozet = {k: sum(1 for r in sonuclar if r['sonuc'] == k) for k in ('ok', 'hata', 'atlandi')}
        return jsonify({'ozet': ozet, 'sonuclar': sonuclar, 'durduruldu': _AS400_DURDUR.is_set()})
    finally:
        _AS400_DURDUR.clear()
        _AS400_KILIT.release()


@app.route('/api/as400/teyit_durum', methods=['GET'])
@panel_gerekli(izin='as400-teyit')
def as400_teyit_durum():
    """Şu an bir AS400 gönderimi çalışıyor mu? (Gönder butonundan sonra Durdur'u
    doğru göstermek için sayfa açılışında sorulur.)"""
    return jsonify({'calisiyor': _AS400_KILIT.locked(), 'durduruluyor': _AS400_DURDUR.is_set()})


@app.route('/api/as400/teyit_durdur', methods=['POST'])
@panel_gerekli(izin='as400-teyit')
def as400_teyit_durdur():
    """Devam eden gönderimi NAZİKÇE durdurur: bayrağı set eder; döngü işlenmekte
    olan launch'ı bitirince kalan satırlara GEÇMEZ. Çalışan cscript ASLA öldürülmez
    (record-lock riski). Gönderim yoksa 'bos' döner."""
    if not _AS400_KILIT.locked():
        return jsonify({'durum': 'bos', 'mesaj': 'Şu an devam eden gönderim yok'})
    _AS400_DURDUR.set()
    return jsonify({'durum': 'durduruluyor',
                    'mesaj': 'Durdurma istendi — işlenmekte olan launch bitince duracak'})


def _is_emri_dus(conn, referans, adet):
    """Başarılı teyit sonrası İŞ YÖNETİMİ kit düşümü (kullanıcı 2026-07-22):
    kaynak/montaj iş emrinden (referans_takip) teyit edilen adet düşülür;
    kalan <= 3 ise (tam / fazla / 3 eksik üretim) iş emri LİSTEDEN SİLİNİR —
    'üretildi ama listeden silinmesi unutuldu' problemi biter.
    Eşleşme gevşek (boşluk+nokta+büyük/küçük duyarsız). Döner: ek mesaj | ''."""
    try:
        adet = int(adet or 0)
    except (TypeError, ValueError):
        return ''
    if not referans or adet <= 0:
        return ''
    try:
        satir = conn.execute(
            "SELECT id, referans_kodu, hedef_adet, bolum FROM referans_takip "
            "WHERE COALESCE(bolum,'') IN ('kaynak','montaj') "
            "  AND COALESCE(lokasyon,'TK2')='TK2' "
            "  AND COALESCE(hedef_adet,0) > 0 "
            "  AND UPPER(REPLACE(REPLACE(referans_kodu,' ',''),'.','')) = "
            "      UPPER(REPLACE(REPLACE(?,' ',''),'.','')) "
            "ORDER BY id LIMIT 1", (referans,)).fetchone()
        if not satir:
            return ''
        kalan = int(satir['hedef_adet']) - adet
        if kalan <= 3:
            conn.execute("DELETE FROM referans_takip WHERE id=?", (satir['id'],))
            conn.commit()
            return f" · iş emri listeden SİLİNDİ ({satir['bolum']}, hedef {satir['hedef_adet']}, kalan {kalan})"
        conn.execute(
            "UPDATE referans_takip SET hedef_adet=?, guncelleme_tarihi=datetime('now','localtime') WHERE id=?",
            (kalan, satir['id']))
        conn.commit()
        return f" · iş emrinden düşüldü ({satir['bolum']}: kalan {kalan})"
    except Exception as e:
        return f" · iş emri düşümü yapılamadı: {e}"


# ── LAPTOP→SERVER İŞ EMRİ DÜŞÜM KÖPRÜSÜ (2026-07-24) ──
# Teyit LAPTOPTA koşuyor (PCOMM) ama iş takip verisi SERVER'da (canlı DB). Bu yüzden
# laptop'un yerel düşümü kullanıcının gördüğü listeye ulaşmıyordu. Köprü: LAPTOP'ta
# server_sync.json ({"url":"http://<server>:5000"}) VARSA teyit sonrası düşüm SERVER'a
# HTTP ile yollanır; server yoksa/ulaşılamazsa yerel düşüme düşülür. SERVER tarafında EK
# CONFIG/TOKEN GEREKMEZ — endpoint yalnız İÇ AĞ'dan (LAN) kabul eder, dış (cloudflared) red.
def _server_sync_url():
    """Laptop client modu: server_sync.json içindeki url (yoksa None = yerel mod).
    Bu dosya makineye özel (gitignore) — SADECE laptopta bulunur, server'da YOK."""
    try:
        p = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'server_sync.json')
        if not os.path.exists(p):
            return None
        with open(p, encoding='utf-8') as f:
            return (json.load(f).get('url') or None)
    except Exception as e:
        print(f'[server_sync] config okunamadı: {e}')
        return None


def _is_ic_ag_istegi():
    """dus_sync makine-arası İÇ uç noktası: yalnız doğrudan LAN'dan (192.168.x) kabul;
    Cloudflare tüneli (CF-Ray/CF-Connecting-IP header'lı DIŞ istek) reddedilir.
    (App ProxyFix kullanmıyor → remote_addr gerçek TCP peer, spoof edilemez.)"""
    if request.headers.get('CF-Ray') or request.headers.get('CF-Connecting-IP'):
        return False
    return (request.remote_addr or '').startswith('192.168.')


def _is_emri_sunucuya_bildir(url, referans, adet):
    """Laptop: başarılı teyit sonrası server'ın iş takip düşümünü HTTP ile tetikler.
    Best-effort. Döner: (ok:bool, mesaj:str)."""
    try:
        import requests
        r = requests.post(url.rstrip('/') + '/api/is_emri/dus_sync',
                          json={'referans': referans, 'adet': adet}, timeout=8)
        if r.status_code == 200:
            return True, ((r.json() or {}).get('mesaj', '') or '')
        return False, f" · ⚠ iş emri server düşümü başarısız (HTTP {r.status_code})"
    except Exception as e:
        return False, f" · ⚠ iş emri server'a ULAŞILAMADI ({e})"


def _is_emri_dus_router(conn, referans, adet):
    """server_sync url varsa (laptop) düşümü SERVER'da yap (authoritative); server
    ulaşılamazsa YEREL düşüme düş + uyarı (interim kayıp olmasın). url yoksa yerel."""
    url = _server_sync_url()
    if not url:
        return _is_emri_dus(conn, referans, adet)          # server/standalone
    ok, mesaj = _is_emri_sunucuya_bildir(url, referans, adet)
    if ok:
        return mesaj                                       # server authoritative (boş = eşleşme yok)
    return _is_emri_dus(conn, referans, adet) + mesaj + ' — elle kontrol'


@app.route('/api/is_emri/dus_sync', methods=['POST'])
def api_is_emri_dus_sync():
    """Makine-arası iş emri düşümü (laptop→server). Yalnız İÇ AĞ'dan (LAN) kabul edilir
    (dış/cloudflared reddedilir); panel oturumu/token GEREKMEZ. Server KENDİ
    referans_takip'inde _is_emri_dus çalıştırır."""
    if not _is_ic_ag_istegi():
        return jsonify({'hata': 'yalnız iç ağ (LAN) erişimi'}), 403
    data = request.get_json(silent=True) or {}
    referans = (data.get('referans') or '').strip()
    try:
        adet = int(data.get('adet') or 0)
    except (TypeError, ValueError):
        adet = 0
    conn = get_db()
    mesaj = _is_emri_dus(conn, referans, adet)
    return jsonify({'mesaj': mesaj, 'dustu': bool(mesaj)})


def _as400_cfi_bugun(article, causal='CFI'):
    """BUGÜN girilen CFI/COP hareket adetleri (bağımsız doğrulama kanalı).
    Döner: [adet, ...] — BMMAF0 MGCACD=causal, hareket tarihi = bugün."""
    import sys as _sys
    _d = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'as400')
    if _d not in _sys.path:
        _sys.path.insert(0, _d)
    import keyring, pyodbc
    import as400_config as _cfg
    bugun = date.today()
    pw = _cfg.sifre_al()
    cn = pyodbc.connect(_cfg.baglanti_dizesi(pw), timeout=20, autocommit=True)
    try:
        rows = cn.cursor().execute(
            "SELECT MGQTA FROM tkc0301F.BMMAF0 WHERE MGCACD=? AND MGARCD=? "
            "AND MGDSSO=? AND MGDAAO=? AND MGDMMO=? AND MGDGGO=?",
            (causal, article, bugun.year // 100, bugun.year % 100, bugun.month, bugun.day)).fetchall()
        return [float(r[0] or 0) for r in rows]
    finally:
        cn.close()


def _cfi_gonder_calistir(conn, satirlar, kullanici, zorla=False):
    """CFI depo girişi satırlarını robotla işler — endpoint VE 17:10 oto koşusu
    ortak çekirdeği. ÇAĞIRAN _AS400_KILIT'i tutuyor olmalı. Döner: sonuclar.
    Kullanıcı akışı (2026-07-20): Warehouse=01D, Causal=CFI, Counterpart=01D,
    kod + adet + Enter. Mükerrer koruması: (1) as400_teyit_log (yil='CF'),
    (2) BMMAF0 RPR+CFI hareketleri (_zaten_teyitli). Doğrulama: bugünkü CFI
    hareketlerinde adet birebir aranır."""
    sonuclar = []
    import sys as _sys
    _d = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'as400')
    if _d not in _sys.path:
        _sys.path.insert(0, _d)
    import launch_esle as _le
    for idx, s in enumerate(satirlar):
        if _AS400_DURDUR.is_set():
            for kalan in satirlar[idx:]:
                sonuclar.append({'referans': str(kalan.get('referans') or '').strip(),
                                 'article': str(kalan.get('article') or '').strip(),
                                 'adet': kalan.get('adet'), 'sonuc': 'atlandi',
                                 'mesaj': 'Kullanıcı durdurdu'})
            break
        referans = str(s.get('referans') or '').strip()
        article = (str(s.get('article') or '').strip() or referans)
        u_tarih = str(s.get('uretim_tarihi') or '').strip()
        try:
            adet = int(s.get('adet') or 0)
        except (TypeError, ValueError):
            adet = 0
        kayit = {'referans': referans, 'article': article, 'adet': adet, 'uretim_tarihi': u_tarih}
        if not (article and 0 < adet <= 99999 and u_tarih):
            sonuclar.append({**kayit, 'sonuc': 'hata', 'mesaj': 'Geçersiz satır parametresi'})
            continue
        # Mükerrer 1: kendi log'umuz (CFI kayıtları yil='CF', launch_no=article)
        var = conn.execute(
            "SELECT id FROM as400_teyit_log WHERE uretim_tarihi=? AND yil='CF' AND launch_no=? AND sonuc='ok'",
            (u_tarih, article)).fetchone()
        if var and not zorla:
            sonuclar.append({**kayit, 'sonuc': 'atlandi', 'mesaj': 'Bu üretim günü için CFI zaten gönderilmiş (log)'})
            continue
        # Mükerrer 2: ERP hareketleri (RPR+CFI) — HEM article HEM üretim referansı
        # (2026-07-21 4082W dersi: launch teyidi gerçek article'a, CFI adayı çıplak
        # koda bakıyordu → aynı üretim iki kez gitti). Aynı-gün + geçmiş-açıklama kurallı.
        try:
            hrk_map = _le.teyit_hareketleri([article, referans])
            hrk = list(hrk_map.get(_le.kanonik(article), []))
            if _le.kanonik(referans) != _le.kanonik(article):
                hrk += hrk_map.get(_le.kanonik(referans), [])
            zt, ilgili = _le._zaten_teyitli(hrk, u_tarih, adet,
                                            _le.ref_uretim_gecmisi(referans, u_tarih))
        except Exception:
            zt, ilgili = None, []
        if zt == 'kesin' and not zorla:
            h0 = ilgili[0] if ilgili else {}
            sonuclar.append({**kayit, 'sonuc': 'atlandi',
                             'mesaj': f"ERP'de karşılığı var: {h0.get('tarih','')} {h0.get('adet',0):g} adet ({h0.get('tur','')})"})
            continue
        # Mükerrer 3: KAPASİTE (2026-07-30) — CFI de depoya stok yazar, aynı fren
        if not zorla:
            kap = _kapasite_reddi(conn, referans, u_tarih, adet)
            if kap:
                sonuclar.append({**kayit, 'sonuc': 'hata', 'mesaj': kap})
                continue
        # KOD DOĞRULAMA (2026-07-30): article ERP master'ında yoksa robota GÖNDERME.
        # '10.300.1992A.OP.1' vakası: operatör kod alanına operasyon bilgisi yazmış;
        # robot ekrana yazamayınca alan taştı ve AYNI KOŞUDAKİ DİĞER satırlar da
        # gönderilemedi. Tek bozuk kod tüm kuyruğu durduruyordu.
        gecersiz = _article_gecersiz_mi(article)
        if gecersiz:
            sonuclar.append({**kayit, 'sonuc': 'hata', 'mesaj': gecersiz})
            continue
        cikti, robot_hata = _as400_robot_calistir('cfi_gir.js', [article, adet], 120)
        if robot_hata:
            sonuclar.append({**kayit, 'sonuc': 'hata', 'mesaj': robot_hata})
            continue
        robot_ok = 'SONUC=OK' in cikti
        # Bağımsız doğrulama: bugünkü CFI hareketlerinde adet birebir var mı
        try:
            dogru = any(abs(q - adet) < 0.001 for q in _as400_cfi_bugun(article))
        except Exception:
            dogru = None   # okunamadı — yalnız robot çıktısına güven
        dogrulandi = robot_ok and (dogru is not False)
        son_satirlar = ' | '.join(l for l in cikti.splitlines() if l.strip())[-400:]
        sonuc = 'ok' if dogrulandi else 'hata'
        if dogrulandi:
            mesaj = f'CFI girildi: {article} → {adet} adet' + ('' if dogru else ' (hareket sorgusu doğrulanamadı, robot OK)')
            mesaj += _is_emri_dus_router(conn, referans, adet)
            # COP (hurda) ayrı fazda EN SONA girilir (bkz. /api/as400/cop_gonder)
        elif robot_ok:
            mesaj = f'Robot OK ama bugünkü CFI hareketlerinde {adet} bulunamadı — elle kontrol edin'
        else:
            mesaj = f'Robot iptal: {son_satirlar}'
        conn.execute(
            "INSERT INTO as400_teyit_log (uretim_tarihi, yil, launch_no, referans, article, adet, bayrak, sonuc, mesaj, olusturan) "
            "VALUES (?,?,?,?,?,?,?,?,?,?)",
            (u_tarih, 'CF', article, referans, article, adet, 'CFI', sonuc, mesaj, kullanici))
        conn.commit()
        sonuclar.append({**kayit, 'sonuc': sonuc, 'mesaj': mesaj})
    return sonuclar


@app.route('/api/as400/cfi_gonder', methods=['POST'])
@panel_gerekli(izin='as400-teyit')
def as400_cfi_gonder():
    """Launch'sız üretim için CFI depo girişi (Session B robotu, 07>01>F1 ekranı).
    Body: {satirlar: [{referans, article?, adet, uretim_tarihi}], zorla}
    Çekirdek _cfi_gonder_calistir'da (17:10 oto koşusuyla ORTAK)."""
    data = request.get_json(silent=True) or {}
    satirlar = data.get('satirlar') or []
    zorla = bool(data.get('zorla'))
    if not satirlar:
        return jsonify({'hata': 'satirlar zorunlu'}), 400
    if len(satirlar) > 60:
        return jsonify({'hata': 'Tek seferde en fazla 60 satır'}), 400
    if not _AS400_KILIT.acquire(blocking=False):
        return jsonify({'hata': 'Devam eden bir AS400 gönderimi var — bitmesini bekleyin'}), 409
    try:
        _AS400_DURDUR.clear()
        sonuclar = _cfi_gonder_calistir(get_db(), satirlar, g.panel_ku['kullanici_adi'], zorla)
        ozet = {k: sum(1 for r in sonuclar if r['sonuc'] == k) for k in ('ok', 'hata', 'atlandi')}
        return jsonify({'ozet': ozet, 'sonuclar': sonuclar, 'durduruldu': _AS400_DURDUR.is_set()})
    finally:
        _AS400_DURDUR.clear()
        _AS400_KILIT.release()


# ── AS400 PLANLAMA (kaynak MRP yardımcısı, kullanıcı 2026-07-23 akışı öğretti) ──
# Manuel akış: 07>10>01 stok → sipariş yoksa üst kodlar (implosion) → üst siparişleri
# → üretilecekse alt parça stok kontrolü → launch. Dosyalar (keşif 2026-07-23,
# çapalarla doğrulandı): STOK=TKC0300F.BSMAF0 (partita nedeniyle DAİMA GROUP BY);
# SİPARİŞ=TKC0301F.BORDF1 (açık: OCDOSO<>'C' AND OCQTOR>OCQTCO; PR/BFABF0 KULLANMA —
# operasyon süreli MRP önerileri); BOM=TKC0301F.BSPEF2 (Y2ARTI üst, Y2RICD alt,
# Y2IVQT birim miktar — Y2QTA hep 0, kullanma). CHAR(21) alanlar boşluk dolgulu →
# eşitlik parametreyle (DB2 blank-pad), TRIM yalnız SELECT'te.

def _plan_stoklar(cn, articles):
    """{artikel: {'depolar': {depo: {...}}, 'kullanilabilir': 01D+CF2+MK2+MT2 stok toplamı}}"""
    if not articles:
        return {}
    yer = ','.join('?' * len(articles))
    sonuc = {}
    for r in cn.cursor().execute(
            f"SELECT TRIM(S0ARTI), TRIM(S0MGCD), SUM(S0GIAD), SUM(S0IMPP)+SUM(S0IMPC), "
            f"SUM(S0ORDF)+SUM(S0ORDP) FROM TKC0300F.BSMAF0 "
            f"WHERE S0SOCI='01' AND S0ARTI IN ({yer}) GROUP BY TRIM(S0ARTI), TRIM(S0MGCD)",
            list(articles)):
        art, depo = r[0], r[1]
        stok, engaged, ordered = float(r[2] or 0), float(r[3] or 0), float(r[4] or 0)
        if not (stok or engaged or ordered):
            continue
        a = sonuc.setdefault(art, {'depolar': {}, 'kullanilabilir': 0.0, 'toplam_stok': 0.0})
        a['depolar'][depo] = {'stok': stok, 'engaged': engaged, 'ordered': ordered,
                              'avail': stok - engaged + ordered}
        a['toplam_stok'] += stok
        if depo in ('01D', 'CF2', 'MK2', 'MT2'):
            a['kullanilabilir'] += stok
    return sonuc


def _plan_siparisler(cn, articles):
    """{artikel: {'acik_toplam': N, 'satirlar': [{siparis, musteri, adet, acik, teslim}]}}"""
    if not articles:
        return {}
    yer = ','.join('?' * len(articles))
    sonuc = {}
    for r in cn.cursor().execute(
            f"SELECT TRIM(d.OCARCD), d.OCRED2, d.OCRENU, TRIM(a.ABRGSC), "
            f"d.OCQTOR, d.OCQTCO, d.OCCCD2, d.OCCCD3, d.OCCCD4 "
            f"FROM TKC0301F.BORDF1 d LEFT JOIN TKC0301F.BANAF0 a ON a.ABCFCD=d.OCCFCD "
            f"WHERE d.OCARCD IN ({yer}) AND d.OCDOSO<>'C' AND d.OCQTOR>d.OCQTCO "
            f"ORDER BY d.OCCCD2, d.OCCCD3, d.OCCCD4",
            list(articles)):
        art = r[0]
        acik = float(r[4] or 0) - float(r[5] or 0)
        s = sonuc.setdefault(art, {'acik_toplam': 0.0, 'satirlar': []})
        s['acik_toplam'] += acik
        if len(s['satirlar']) < 10:
            # EBCDIC(1026)→ODBC çevirisinde Türkçe harfler bozuk gelir: "=Ü £=Ö $=İ §=Ş
            musteri = str(r[3] or '').strip().translate(str.maketrans('"£$§', 'ÜÖİŞ'))
            s['satirlar'].append({
                'siparis': f'{int(r[1])}-{int(r[2])}',
                'musteri': musteri,
                'acik': acik,
                'teslim': f'20{int(r[6] or 0):02d}-{int(r[7] or 0):02d}-{int(r[8] or 0):02d}',
            })
    return sonuc


def _cop_gonder_calistir(conn, satirlar, kullanici, zorla=False):
    """HURDA (COP) girişi satırlarını robotla işler — endpoint VE 17:10 oto koşusu
    ortak çekirdeği. ÇAĞIRAN _AS400_KILIT'i tutuyor olmalı. Döner: sonuclar.
    COP'lar EN SON faz (kullanıcı 2026-07-23: launch/CFI fazlarına ARAYA GİRMESİN;
    robot Rientro↔07>01 mekik dokuyordu). Tüm hurdalar ardışık girilir: B,
    07>01>F1 ekranında kalır (alanlar kalıcı — hızlı)."""
    sonuclar = []
    for idx, s in enumerate(satirlar):
        if _AS400_DURDUR.is_set():
            for kalan in satirlar[idx:]:
                sonuclar.append({'referans': str(kalan.get('referans') or '').strip(),
                                 'article': str(kalan.get('article') or '').strip(),
                                 'adet': kalan.get('adet'), 'bayrak': 'COP',
                                 'sonuc': 'atlandi', 'mesaj': 'Kullanıcı durdurdu'})
            break
        referans = str(s.get('referans') or '').strip()
        article = (str(s.get('article') or '').strip() or referans)
        u_tarih = str(s.get('uretim_tarihi') or '').strip()
        try:
            adet = int(s.get('adet') or 0)
        except (TypeError, ValueError):
            adet = 0
        kayit = {'referans': referans, 'article': article, 'adet': adet,
                 'uretim_tarihi': u_tarih, 'bayrak': 'COP'}
        if not (article and 0 < adet <= 99999 and u_tarih):
            sonuclar.append({**kayit, 'sonuc': 'hata', 'mesaj': 'Geçersiz satır parametresi'})
            continue
        # Mükerrer: aynı üretim günü + article için COP zaten girildiyse atla
        var = conn.execute(
            "SELECT id FROM as400_teyit_log WHERE uretim_tarihi=? AND yil='CO' AND launch_no=? AND sonuc='ok'",
            (u_tarih, article)).fetchone()
        if var:
            sonuclar.append({**kayit, 'sonuc': 'atlandi', 'mesaj': 'Bu üretim günü için hurda COP zaten girilmiş'})
            continue
        # Mükerrer koruması 2 — ERP'NİN KENDİSİ (2026-07-30, 10.300.4180A olayı):
        # yukarıdaki kontrol yalnızca KENDİ log'umuza bakar. Robot COP'u ERP'ye
        # yazdığı hâlde 'SONUC=OK' dönmezse (ekran akışı beklenmedik bitti,
        # bağlantı koptu) log'a 'hata' düşer; satır COP bekleyenlerde kalır ve
        # tekrar gönderilirse ERP'ye İKİNCİ kez COP girilir. 4180A'da ERP'de COP
        # 18 vardı ama panel hâlâ bekliyor gösteriyordu. ERP'de bugün bu adet
        # zaten varsa gönderme; log'a 'ok' yaz ki panel de düzelsin.
        if not zorla:
            try:
                _mevcut = _as400_cfi_bugun(article, causal='COP')
            except Exception:
                _mevcut = None          # ERP okunamadı → engelleme (güvenli yön)
            if _mevcut and any(abs(q - adet) < 0.001 for q in _mevcut):
                _msg = (f'ERP\'de bugün {article} için {adet:g} adetlik COP hareketi ZATEN var '
                        f'(robot daha önce yazmış ama log\'a "ok" düşmemiş). Mükerrer COP '
                        f'girilmedi; kayıt eşitlendi.')
                conn.execute(
                    "INSERT INTO as400_teyit_log (uretim_tarihi, yil, launch_no, referans, article, "
                    "adet, bayrak, sonuc, mesaj, olusturan) VALUES (?,?,?,?,?,?,?,?,?,?)",
                    (u_tarih, 'CO', article, referans, article, adet, 'COP', 'ok', _msg, kullanici))
                conn.commit()
                sonuclar.append({**kayit, 'sonuc': 'atlandi', 'mesaj': _msg})
                continue
        # COP da aynı ekranı (07>01>F1) kullanır — bozuk kod burada da kuyruğu kilitler
        gecersiz = _article_gecersiz_mi(article)
        if gecersiz:
            sonuclar.append({**kayit, 'sonuc': 'hata', 'mesaj': gecersiz})
            continue
        cikti, robot_hata = _as400_robot_calistir('cfi_gir.js', [article, adet, 'COP'], 120)
        if robot_hata:
            sonuclar.append({**kayit, 'sonuc': 'hata', 'mesaj': robot_hata})
            continue
        robot_ok = 'SONUC=OK' in cikti
        try:
            dogru = any(abs(q - adet) < 0.001 for q in _as400_cfi_bugun(article, causal='COP'))
        except Exception:
            dogru = None
        dogrulandi = robot_ok and (dogru is not False)
        sonuc = 'ok' if dogrulandi else 'hata'
        if dogrulandi:
            mesaj = f'♻ Hurda COP girildi: {article} → {adet} adet'
        elif robot_ok:
            mesaj = f'Robot OK ama bugünkü COP hareketlerinde {adet} bulunamadı — elle kontrol edin'
        else:
            son_satirlar = ' | '.join(l for l in cikti.splitlines() if l.strip())[-300:]
            mesaj = f'Robot iptal: {son_satirlar}'
        conn.execute(
            "INSERT INTO as400_teyit_log (uretim_tarihi, yil, launch_no, referans, article, adet, bayrak, sonuc, mesaj, olusturan) "
            "VALUES (?,?,?,?,?,?,?,?,?,?)",
            (u_tarih, 'CO', article, referans, article, adet, 'COP', sonuc, mesaj, kullanici))
        conn.commit()
        sonuclar.append({**kayit, 'sonuc': sonuc, 'mesaj': mesaj})
    return sonuclar


@app.route('/api/as400/cop_gonder', methods=['POST'])
@panel_gerekli(izin='as400-teyit')
def as400_cop_gonder():
    """HURDA (COP) girişleri — EN SON faz. Çekirdek _cop_gonder_calistir'da
    (17:10 oto koşusuyla ORTAK). Body: {satirlar: [{referans, article, adet(hurda),
    uretim_tarihi}]}"""
    data = request.get_json(silent=True) or {}
    satirlar = data.get('satirlar') or []
    if not satirlar:
        return jsonify({'hata': 'satirlar zorunlu'}), 400
    if len(satirlar) > 60:
        return jsonify({'hata': 'Tek seferde en fazla 60 satır'}), 400
    if not _AS400_KILIT.acquire(blocking=False):
        return jsonify({'hata': 'Devam eden bir AS400 gönderimi var — bitmesini bekleyin'}), 409
    try:
        _AS400_DURDUR.clear()
        sonuclar = _cop_gonder_calistir(get_db(), satirlar, g.panel_ku['kullanici_adi'],
                                        zorla=bool(data.get('zorla')))
        ozet = {k: sum(1 for r in sonuclar if r['sonuc'] == k) for k in ('ok', 'hata', 'atlandi')}
        return jsonify({'ozet': ozet, 'sonuclar': sonuclar, 'durduruldu': _AS400_DURDUR.is_set()})
    finally:
        _AS400_DURDUR.clear()
        _AS400_KILIT.release()


@app.route('/api/as400/teyit_gecmisi', methods=['GET'])
@panel_gerekli(izin='as400-teyit')
def as400_teyit_gecmisi():
    """Geçmişe dönük teyit kaydı (2026-07-30, kullanıcı isteği: "hangi gün ne
    teyit verildiğini görebileceğimiz bir filtre bölümü olsun").

    Sayfadaki listeler yalnızca SON birkaç günü ve robotun SON 15 gönderimini
    gösteriyordu; "geçen hafta şu koda teyit verdik mi" sorusunun cevabı yoktu.
    Kaynak as400_teyit_log — yani BİZİM gönderdiklerimiz (ERP'nin tamamı değil).

    Parametreler: bas, bit (YYYY-MM-DD) · alan=uretim|gonderim (hangi tarihe göre
    filtrelenecek) · tip=hepsi|launch|cfi|cop · sonuc=hepsi|ok|hata|atlandi ·
    ref (kod parçası, opsiyonel)."""
    bas = (request.args.get('bas') or '').strip()
    bit = (request.args.get('bit') or '').strip()
    if not (bas and bit):
        bit = date.today().isoformat()
        bas = (date.today() - timedelta(days=7)).isoformat()
    alan = 'gonderim' if (request.args.get('alan') == 'gonderim') else 'uretim'
    tip = (request.args.get('tip') or 'hepsi').lower()
    sonuc_f = (request.args.get('sonuc') or 'hepsi').lower()
    ref = (request.args.get('ref') or '').strip()

    # gonderim tarihi created_at'in gun kismi; uretim tarihi zaten YYYY-MM-DD
    tarih_ifade = "DATE(created_at)" if alan == 'gonderim' else "uretim_tarihi"
    kosul = [f"{tarih_ifade} BETWEEN ? AND ?"]
    par = [bas, bit]
    if tip == 'launch':
        kosul.append("yil NOT IN ('CF','CO')")
    elif tip == 'cfi':
        kosul.append("yil='CF'")
    elif tip == 'cop':
        kosul.append("yil='CO'")
    if sonuc_f in ('ok', 'hata', 'atlandi'):
        kosul.append("sonuc=?"); par.append(sonuc_f)
    if ref:
        kosul.append("UPPER(REPLACE(referans,' ','')) LIKE ?")
        par.append('%' + ref.upper().replace(' ', '') + '%')

    conn = get_db()
    satirlar = [dict(r) for r in conn.execute(
        "SELECT id, created_at, uretim_tarihi, yil, launch_no, referans, article, adet, "
        "bayrak, sonuc, mesaj, olusturan FROM as400_teyit_log "
        f"WHERE {' AND '.join(kosul)} ORDER BY id DESC LIMIT 500", par).fetchall()]
    for s in satirlar:
        s['tip'] = 'COP' if s['yil'] == 'CO' else ('CFI' if s['yil'] == 'CF' else 'LAUNCH')

    ozet = {'satir': len(satirlar), 'ok': 0, 'hata': 0, 'atlandi': 0,
            'LAUNCH': 0, 'CFI': 0, 'COP': 0, 'adet_ok': 0.0}
    gunler = {}
    for s in satirlar:
        ozet[s['sonuc']] = ozet.get(s['sonuc'], 0) + 1
        ozet[s['tip']] += 1
        g = (s['created_at'] or '')[:10] if alan == 'gonderim' else s['uretim_tarihi']
        d = gunler.setdefault(g, {'tarih': g, 'ok': 0, 'hata': 0, 'atlandi': 0, 'adet': 0.0})
        d[s['sonuc']] = d.get(s['sonuc'], 0) + 1
        if s['sonuc'] == 'ok':
            ozet['adet_ok'] += float(s['adet'] or 0)
            d['adet'] += float(s['adet'] or 0)
    return jsonify({'satirlar': satirlar, 'ozet': ozet,
                    'gunler': sorted(gunler.values(), key=lambda x: x['tarih'] or '', reverse=True),
                    'filtre': {'bas': bas, 'bit': bit, 'alan': alan, 'tip': tip,
                               'sonuc': sonuc_f, 'ref': ref},
                    'limit_asildi': len(satirlar) >= 500})


@app.route('/api/as400/planlama', methods=['GET'])
@panel_gerekli(izin='as400-teyit')
def as400_planlama():
    """Tek referans için planlama görünümü: depo stokları + doğrudan açık
    siparişler + üst kodlar (3 seviye implosion; her biri stok+sipariş ile) +
    alt parçalar (1 seviye explosion; 01D/CF2/MK2/MT2 stoklarıyla)."""
    ref = (request.args.get('ref') or '').strip()
    if not ref or len(ref) > 21:
        return jsonify({'hata': 'ref zorunlu (en fazla 21 karakter)'}), 400
    import sys as _sys
    _d = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'as400')
    if _d not in _sys.path:
        _sys.path.insert(0, _d)
    import keyring, pyodbc
    import launch_esle as _le
    import as400_config as _cfg
    ref = _le.bosluk_nokta(ref)
    # ASCII dışı karakter (Türkçe İ/Ü/Ö...) AS400 sorgusunu CWBNL0107 ile patlatır
    import re as _re2
    if not _re2.match(r'^[A-Za-z0-9./\-]{3,21}$', ref):
        return jsonify({'hata': f"Geçersiz referans: '{ref}' — yalnız harf/rakam/nokta/tire (Türkçe karakter olmadan)"}), 400
    try:
        pw = _cfg.sifre_al()
        cn = pyodbc.connect(_cfg.baglanti_dizesi(pw), timeout=30, autocommit=True)
    except Exception as e:
        return jsonify({'hata': f'AS400 bağlantısı kurulamadı: {e}'}), 502
    try:
        # ÜST KODLAR — 3 seviyeye kadar implosion (iteratif; döngü/patlama korumalı)
        ustler = []          # [{kod, seviye, birim}]
        gorulen = {ref}
        seviye_kume = [ref]
        for seviye in (1, 2, 3):
            if not seviye_kume or len(gorulen) > 40:
                break
            yer = ','.join('?' * len(seviye_kume))
            yeni = []
            for r in cn.cursor().execute(
                    f"SELECT DISTINCT TRIM(Y2ARTI), TRIM(Y2RICD), Y2IVQT "
                    f"FROM TKC0301F.BSPEF2 WHERE Y2RICD IN ({yer})", seviye_kume):
                ust = r[0]
                if ust and ust not in gorulen:
                    gorulen.add(ust)
                    ustler.append({'kod': ust, 'seviye': seviye,
                                   'alt': r[1], 'birim': float(r[2] or 0)})
                    yeni.append(ust)
            seviye_kume = yeni
        # ALT PARÇA AĞACI — ÇOK SEVİYELİ explosion (kullanıcı 2026-07-23 kademeli
        # kontrol: bu kademenin stoğu yetmiyorsa bir alt kademenin parçalarıyla
        # üretilebilir mi? — derinlik 4, toplam ≤ 80 düğüm, döngü korumalı).
        agac_kodlari = set()

        def _patlat(kod, derinlik, yol):
            if derinlik <= 0 or len(agac_kodlari) > 80 or kod in yol:
                return []
            cocuklar = []
            for r in cn.cursor().execute(
                    "SELECT TRIM(Y2RICD), Y2IVQT, Y2IVUM FROM TKC0301F.BSPEF2 "
                    "WHERE Y2ARTI = ? ORDER BY Y2RIPR", (kod,)):
                alt_kod = str(r[0] or '').strip()
                if not alt_kod:
                    continue
                agac_kodlari.add(alt_kod)
                cocuklar.append({
                    'kod': alt_kod, 'birim': float(r[1] or 0), 'um': str(r[2] or '').strip(),
                    'cocuklar': _patlat(alt_kod, derinlik - 1, yol | {kod}),
                })
            return cocuklar

        agac = _patlat(ref, 4, set())
        # STOK + SİPARİŞLER (ref + üstler + ağaçtaki tüm kodlar tek seferde)
        tum_kodlar = [ref] + [u['kod'] for u in ustler] + sorted(agac_kodlari)
        stoklar = _plan_stoklar(cn, tum_kodlar)
        siparisler = _plan_siparisler(cn, [ref] + [u['kod'] for u in ustler])
        for u in ustler:
            u['stok'] = stoklar.get(u['kod'], {'depolar': {}, 'kullanilabilir': 0, 'toplam_stok': 0})
            u['siparis'] = siparisler.get(u['kod'], {'acik_toplam': 0, 'satirlar': []})

        def _stok_bagla(dugumler):
            for d in dugumler:
                s = stoklar.get(d['kod'], {'depolar': {}, 'kullanilabilir': 0, 'toplam_stok': 0})
                d['stok'] = s
                # MDT = FASONDA (kullanıcı 2026-07-23): üretilmiş, dış işlemde; dönüşte
                # ÜST koda girer → arz sayılır ama ayrı gösterilir.
                d['fasonda'] = float((s['depolar'].get('MDT') or {}).get('stok', 0) or 0)
                _stok_bagla(d['cocuklar'])
        _stok_bagla(agac)
        kok_stok = stoklar.get(ref, {'depolar': {}, 'kullanilabilir': 0, 'toplam_stok': 0})
        return jsonify({
            'ref': ref,
            'stok': kok_stok,
            'fasonda': float((kok_stok['depolar'].get('MDT') or {}).get('stok', 0) or 0),
            'dogrudan_siparis': siparisler.get(ref, {'acik_toplam': 0, 'satirlar': []}),
            'ustler': ustler,
            'agac': agac,
        })
    except Exception as e:
        return jsonify({'hata': f'AS400 planlama sorgusu başarısız: {e}'}), 502
    finally:
        cn.close()


def _acik_transferler_sorgula(gunler=60):
    """BMMAF0'daki iptal edilmemiş TA 02→01 transferlerini döndürür (liste).
    GET endpoint'i VE 16:45 oto iptal koşusu ortak sorgusu. Hata → raise."""
    bas = date.today() - timedelta(days=gunler)
    esik = bas.year * 10000 + bas.month * 100 + bas.day   # YYYYMMDD karşılaştırma
    import sys as _sys
    _d = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'as400')
    if _d not in _sys.path:
        _sys.path.insert(0, _d)
    import keyring, pyodbc
    import as400_config as _cfg
    pw = _cfg.sifre_al()
    cn = pyodbc.connect(_cfg.baglanti_dizesi(pw), timeout=30, autocommit=True)
    try:
        rows = cn.cursor().execute(
            "SELECT MGDSSO, MGDAAO, MGDMMO, MGDGGO, MGARCD, MGQTA, MGUTVA, MGANRE, MGNURE, MGPRRE "
            "FROM tkc0301F.BMMAF0 "
            "WHERE MGCACD='TA' AND MGMGCD='02' AND MGMCCD='01' "
            "  AND (MGDSSO*1000000 + MGDAAO*10000 + MGDMMO*100 + MGDGGO) >= ? "
            "ORDER BY (MGDSSO*1000000 + MGDAAO*10000 + MGDMMO*100 + MGDGGO) DESC, MGNURE DESC, MGPRRE",
            (esik,)).fetchall()
        return [{
            'tarih': f'{int(r[0])}{int(r[1]):02d}-{int(r[2]):02d}-{int(r[3]):02d}',
            'article': str(r[4]).strip(),
            'adet': float(r[5] or 0),
            'kullanici': str(r[6]).strip(),
            'kayit': f'{int(r[7])}-{int(r[8])}',
            'satir': int(r[9]),
        } for r in rows]
    finally:
        cn.close()


@app.route('/api/as400/acik_transferler', methods=['GET'])
@panel_gerekli(izin='as400-teyit')
def as400_acik_transferler():
    """TK1 'hayali stok' transferleri (kullanıcı 2026-07-22 öğretti): alt parça
    yolda iken launch açabilmek için 07>01>F1'den TA causal ile 02→01 stok
    taşınır (02 eksiye düşer); parça gelince Shift+F4 ile İPTAL edilmeli —
    iptal kaydı fiziksel SİLER. Dolayısıyla BMMAF0'da hâlâ duran her TA 02→01
    hareketi = iptal edilmemiş transfer. Tarih bazlı liste döner."""
    try:
        gunler = max(7, min(365, int(request.args.get('gunler') or 60)))
    except (TypeError, ValueError):
        gunler = 60
    try:
        transferler = _acik_transferler_sorgula(gunler)
        return jsonify({'gunler': gunler, 'transferler': transferler})
    except Exception as e:
        return jsonify({'hata': f'AS400 sorgusu başarısız: {e}'}), 502


def _transfer_kayit_duruyor_mu(kayit, satir):
    """İptal DOĞRULAMASI: BMMAF0'da (MGANRE-MGNURE, MGPRRE) hâlâ var mı?
    İptal (Shift+F4) kaydı FİZİKSEL siler → yoksa iptal kesinleşmiştir.
    Döner: True (duruyor) | False (silinmiş) | None (okunamadı)."""
    try:
        yil, no = str(kayit).split('-', 1)
        import sys as _sys
        _d = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'as400')
        if _d not in _sys.path:
            _sys.path.insert(0, _d)
        import keyring, pyodbc
        import as400_config as _cfg
        pw = _cfg.sifre_al()
        cn = pyodbc.connect(_cfg.baglanti_dizesi(pw), timeout=15, autocommit=True)
        try:
            r = cn.cursor().execute(
                "SELECT COUNT(*) FROM tkc0301F.BMMAF0 "
                "WHERE MGCACD='TA' AND MGANRE=? AND MGNURE=? AND MGPRRE=?",
                (int(yil), int(no), int(satir))).fetchone()
            return bool(r and int(r[0]) > 0)
        finally:
            cn.close()
    except Exception:
        return None


def _transfer_iptal_calistir(conn, satirlar, kullanici, dryrun=False):
    """Transfer iptal satırlarını robotla işler (transfer_iptal.js) — manuel
    endpoint VE 16:45 oto koşusu ortak çekirdeği. ÇAĞIRAN _AS400_KILIT'i tutmalı.
    satirlar: [{kayit:'26-245458', satir:200, article, adet, tarih?}]
    Doğrulama: robot SONUC=OK dese bile BMMAF0'da kayıt aranır — silinmişse 'ok'.
    Her sonuç as400_transfer_log'a yazılır. Döner: sonuclar."""
    sonuclar = []
    for idx, s in enumerate(satirlar):
        if _AS400_DURDUR.is_set():
            for kalan in satirlar[idx:]:
                sonuclar.append({'kayit': str(kalan.get('kayit') or ''), 'satir': kalan.get('satir'),
                                 'article': str(kalan.get('article') or ''), 'adet': kalan.get('adet'),
                                 'sonuc': 'atlandi', 'mesaj': 'Kullanıcı durdurdu'})
            break
        kayit = str(s.get('kayit') or '').strip()
        article = str(s.get('article') or '').strip()
        tarih = str(s.get('tarih') or '').strip()
        try:
            satir = int(s.get('satir') or 0)
        except (TypeError, ValueError):
            satir = 0
        try:
            adet = float(s.get('adet') or 0)
        except (TypeError, ValueError):
            adet = 0
        kaydet = {'kayit': kayit, 'satir': satir, 'article': article, 'adet': adet, 'tarih': tarih}
        m = re.match(r'^(\d{2})-(\d{1,7})$', kayit)
        if not (m and satir > 0 and article and adet > 0):
            sonuclar.append({**kaydet, 'sonuc': 'hata', 'mesaj': 'Geçersiz satır parametresi'})
            continue
        # Mükerrer/bayat koruması: kayıt AS400'de hâlâ duruyor mu? (Başka biri
        # elle iptal etmiş olabilir; robot boşuna ekrana girmesin.)
        duruyor = _transfer_kayit_duruyor_mu(kayit, satir)
        if duruyor is False:
            sonuclar.append({**kaydet, 'sonuc': 'atlandi', 'mesaj': 'Kayıt AS400\'de zaten yok (iptal edilmiş)'})
            _transfer_logla(conn, kaydet, 'atlandi', 'Zaten iptal edilmiş', kullanici)
            continue
        adet_arg = ('%g' % adet)
        robot_args = [m.group(1), m.group(2), satir, article, adet_arg] + (['DRYRUN'] if dryrun else [])
        cikti, robot_hata = _as400_robot_calistir('transfer_iptal.js', robot_args, 90)
        if robot_hata:
            sonuclar.append({**kaydet, 'sonuc': 'hata', 'mesaj': robot_hata})
            _transfer_logla(conn, kaydet, 'hata', robot_hata, kullanici)
            continue
        robot_ok = ('SONUC=OK' in cikti) or (dryrun and 'SONUC=DRYRUN-OK' in cikti)
        son_satirlar = ' | '.join(l for l in cikti.splitlines() if l.strip())[-400:]
        if dryrun:
            sonuc = 'ok' if robot_ok else 'hata'
            mesaj = ('DRYRUN doğrulandı (Shift+F4 basılmadı)' if robot_ok
                     else f'DRYRUN iptal: {son_satirlar}')
        else:
            # ASIL doğrulama: kayıt BMMAF0'dan silindi mi?
            silindi = _transfer_kayit_duruyor_mu(kayit, satir)
            if robot_ok and silindi is False:
                sonuc, mesaj = 'ok', f'İptal doğrulandı: {kayit} satır {satir} ({article}, {adet:g}) BMMAF0\'dan silindi'
            elif robot_ok and silindi is None:
                sonuc, mesaj = 'ok', 'Robot OK; BMMAF0 doğrulaması okunamadı — sabah kontrol edin'
            elif robot_ok:
                sonuc, mesaj = 'hata', 'Robot OK dedi ama kayıt BMMAF0\'da HÂLÂ DURUYOR — elle kontrol edin'
            else:
                sonuc, mesaj = 'hata', f'Robot iptal: {son_satirlar}'
        _transfer_logla(conn, kaydet, sonuc, mesaj, kullanici)
        sonuclar.append({**kaydet, 'sonuc': sonuc, 'mesaj': mesaj})
    return sonuclar


def _transfer_logla(conn, k, sonuc, mesaj, kullanici):
    try:
        conn.execute(
            "INSERT INTO as400_transfer_log (kayit, satir, article, adet, hareket_tarihi, sonuc, mesaj, olusturan) "
            "VALUES (?,?,?,?,?,?,?,?)",
            (k['kayit'], k['satir'], k['article'], k['adet'], k.get('tarih') or '', sonuc, mesaj, kullanici))
        conn.commit()
    except Exception as e:
        print(f'[transfer_iptal] log yazılamadı: {e}')


@app.route('/api/as400/transfer_iptal', methods=['POST'])
@panel_gerekli(izin='as400-teyit')
def as400_transfer_iptal():
    """Seçili açık transferleri ROBOTLA iptal eder (07>01, satır başına 2 + Enter,
    Shift+F4). Body: {satirlar: [{kayit, satir, article, adet, tarih?}], dryrun?}
    dryrun=true → robot ekrana girer, satırı doğrular ama Shift+F4'e BASMAZ."""
    data = request.get_json(silent=True) or {}
    satirlar = data.get('satirlar') or []
    dryrun = bool(data.get('dryrun'))
    if not satirlar:
        return jsonify({'hata': 'satirlar zorunlu'}), 400
    if len(satirlar) > 80:
        return jsonify({'hata': 'Tek seferde en fazla 80 satır'}), 400
    if not _AS400_KILIT.acquire(blocking=False):
        return jsonify({'hata': 'Devam eden bir AS400 gönderimi var — bitmesini bekleyin'}), 409
    try:
        _AS400_DURDUR.clear()
        sonuclar = _transfer_iptal_calistir(get_db(), satirlar, g.panel_ku['kullanici_adi'], dryrun)
        ozet = {k: sum(1 for r in sonuclar if r['sonuc'] == k) for k in ('ok', 'hata', 'atlandi')}
        return jsonify({'ozet': ozet, 'sonuclar': sonuclar, 'durduruldu': _AS400_DURDUR.is_set()})
    finally:
        _AS400_DURDUR.clear()
        _AS400_KILIT.release()


# ─────────────────────────────────────────────────────────────
# OTOMATİK AKŞAM KOŞULARI (kullanıcı 2026-07-27):
#   16:45 — açık TA 02→01 transferlerinin TÜMÜ robotla iptal edilir (07>01,
#           satır başına 2 + Enter, Shift+F4). Kullanıcı kararı: hepsi.
#   17:10 — günün TÜM teyit kuyruğu: launch teyitleri (A / tamamlayana S) +
#           CFI girişleri + hurda COP. Bugün dahil (gün 17:00'da bitmiş sayılır).
# Emin olunamayan satırlar GÖNDERİLMEZ → as400_oto_kosu 'sabah_kontrol'
# listesine yazılır; dashboard teyit sayfasındaki "Sabah Kontrol" paneli gösterir.
# MAKİNE GUARD'I: yalnız teyit-agent'lı makinede koşar (server). Laptop'ta app
# açık kalsa bile çifte koşu OLMAZ (agent 127.0.0.1:5010 yalnız server'da).
# ─────────────────────────────────────────────────────────────
_OTO_CONFIG_YOL = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'as400', 'oto_config.json')
_OTO_VARSAYILAN = {
    'transfer_iptal': {'etkin': True, 'saat': '16:45', 'gunler': 60, 'dryrun': False},
    'oto_teyit':      {'etkin': True, 'saat': '17:10', 'gunler': 3},
}


def _oto_config():
    """oto_config.json + varsayılanlar. utf-8-sig: BOM toleransı (mail_config dersi)."""
    cfg = json.loads(json.dumps(_OTO_VARSAYILAN))   # derin kopya
    try:
        with open(_OTO_CONFIG_YOL, encoding='utf-8-sig') as f:
            dosya = json.load(f) or {}
        for k, v in dosya.items():
            if isinstance(v, dict) and k in cfg:
                cfg[k].update(v)
    except FileNotFoundError:
        pass
    except Exception as e:
        print(f'[OTO] oto_config.json okunamadı ({e}) — varsayılanlar kullanılıyor')
    return cfg


def _oto_config_yaz(cfg):
    with open(_OTO_CONFIG_YOL, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def _oto_telafi_gerek_mi(kosu_turu):
    """scheduler'a verilecek 'bugün kaçırıldı mı?' yordamını üretir (closure).

    True döner ⇔ as400_oto_kosu'da BUGÜNE ait o türden HİÇ kayıt yok. Yani iş ne
    çalıştı ne de atlandı → uygulama planlı saatte ayakta değildi.
    Atlama da kayıt bıraktığı için (_oto_atlandi_kaydet), "ayar kapalıydı" durumu
    telafi TETİKLEMEZ — kapalı iş zaten tekrar atlanacaktı, boşuna koşmasın."""
    def _kontrol():
        try:
            _c = db_connect()
            try:
                r = _c.execute(
                    "SELECT 1 FROM as400_oto_kosu WHERE tur=? "
                    "AND date(created_at) = date('now','localtime') LIMIT 1",
                    (kosu_turu,)
                ).fetchone()
                return r is None
            finally:
                _c.close()
        except Exception as e:
            # Okuyamadıysak telafi ETME — belirsizken robot koşturmak yanlış yön
            print(f'[SCHED] telafi kontrolü okunamadı ({kosu_turu}): {e}')
            return False
    return _kontrol


def _oto_atlandi_kaydet(tur, ozet_metin, neden):
    """Erken ÇIKIŞ (atlama) durumlarını da as400_oto_kosu'ya yazar — böylece Sabah
    Kontrol paneli 'hiç çalışmadı mı yoksa atladı mı' ayırt edilebilir (2026-07-28
    dersi: agent kapalı/bölüm kapalı olunca job sessizce dönüyordu → panelde iz yok,
    'dün akşam neden verilmedi' cevaplanamıyordu). Kendi bağlantısını açar/kapatır."""
    try:
        _c = db_connect()
        try:
            _oto_kosu_kaydet(_c, tur, [], [{'tip': 'atlandi', 'neden': neden}], ozet_metin)
        finally:
            _c.close()
    except Exception as e:
        print(f'[OTO-{tur.upper()}] atlanma kaydı yazılamadı: {e}')


def _oto_kosu_kaydet(conn, tur, sonuclar, sabah_kontrol, ozet_metin):
    """Koşu raporunu as400_oto_kosu'ya yazar (Sabah Kontrol paneli buradan okur)."""
    try:
        say = {k: sum(1 for r in sonuclar if r.get('sonuc') == k) for k in ('ok', 'hata', 'atlandi')}
        conn.execute(
            "INSERT INTO as400_oto_kosu (tur, ozet, detay, ok, hata, atlandi, kontrol) VALUES (?,?,?,?,?,?,?)",
            (tur, ozet_metin, json.dumps({'sonuclar': sonuclar, 'sabah_kontrol': sabah_kontrol},
                                         ensure_ascii=False, default=str),
             say['ok'], say['hata'], say['atlandi'], len(sabah_kontrol)))
        # Eski koşuları buda (son 120 kalsın)
        conn.execute("DELETE FROM as400_oto_kosu WHERE id NOT IN "
                     "(SELECT id FROM as400_oto_kosu ORDER BY id DESC LIMIT 120)")
        conn.commit()
        print(f'[OTO-{tur.upper()}] {ozet_metin}')
    except Exception as e:
        print(f'[OTO-{tur.upper()}] koşu raporu yazılamadı: {e}')


def _kilit_bekleyerek_al(maks_sn, etiket):
    """_AS400_KILIT'i sabırla alır (16:45 koşusu uzarsa 17:10 bekleyebilsin).
    Döner: True (alındı) | False (süre doldu)."""
    import time as _t
    son = _t.time() + maks_sn
    while _t.time() < son:
        if _AS400_KILIT.acquire(blocking=False):
            return True
        print(f'[{etiket}] AS400 kilidi meşgul — 15 sn sonra tekrar denenecek')
        _t.sleep(15)
    return False


def oto_transfer_iptal_job():
    """16:45 — açık TA 02→01 transferlerinin TÜMÜNÜ robotla iptal eder."""
    cfg = _oto_config()['transfer_iptal']
    if not cfg.get('etkin'):
        print('[OTO-TRANSFER] kapalı (oto_config) — atlandı')
        _oto_atlandi_kaydet('transfer', 'ATLANDI: oto transfer iptali KAPALI (oto_config)',
                            'Ayar kapalı — Sabah Kontrol panelinden aç.')
        return
    if not _teyit_agent_var():
        print('[OTO-TRANSFER] teyit-agent yok — bu makine robot koşusu için yapılandırılmamış, atlandı')
        _oto_atlandi_kaydet('transfer', 'ATLANDI: teyit-agent kapalı (PCOMM/Session B yok)',
                            'teyit-agent 127.0.0.1:5010 yanıt vermedi — RDP oturumunda PCOMM+agent açık olmalı.')
        return
    conn = db_connect()
    try:
        try:
            transferler = _acik_transferler_sorgula(int(cfg.get('gunler') or 60))
        except Exception as e:
            _oto_kosu_kaydet(conn, 'transfer', [], [{'neden': f'AS400 sorgusu başarısız: {e}'}],
                             f'HATA: açık transferler okunamadı ({e})')
            return
        if not transferler:
            _oto_kosu_kaydet(conn, 'transfer', [], [], 'Açık transfer yok — iptal edilecek hareket bulunmadı')
            return
        # Aynı kayıt numarasının satırları ardışık işlensin (robot her çağrıda
        # kayıt no'yu yeniden girer; sıralı gidince ekran akışı kısa kalır).
        satirlar = sorted(transferler, key=lambda t: (t['kayit'], t['satir']))
        if not _kilit_bekleyerek_al(20 * 60, 'OTO-TRANSFER'):
            _oto_kosu_kaydet(conn, 'transfer', [], [{'neden': 'AS400 kilidi 20 dk boşalmadı — koşu yapılamadı'}],
                             'HATA: kilit alınamadı')
            return
        try:
            _AS400_DURDUR.clear()
            dryrun = bool(cfg.get('dryrun'))
            sonuclar = _transfer_iptal_calistir(conn, satirlar, 'oto-16:45', dryrun=dryrun)
        finally:
            _AS400_DURDUR.clear()
            _AS400_KILIT.release()
        sabah = [{'tip': 'transfer', 'neden': s.get('mesaj', ''), **{k: s.get(k) for k in ('kayit', 'satir', 'article', 'adet', 'tarih')}}
                 for s in sonuclar if s.get('sonuc') != 'ok']
        ok_n = sum(1 for s in sonuclar if s.get('sonuc') == 'ok')
        _oto_kosu_kaydet(conn, 'transfer', sonuclar, sabah,
                         f'{len(satirlar)} açık transfer: {ok_n} iptal edildi'
                         + (f', {len(sabah)} sabah kontrole ayrıldı' if sabah else '')
                         + (' [DRYRUN]' if dryrun else ''))
    finally:
        conn.close()


def _oto_kuyruk_olustur(conn, tarihler):
    """17:10 koşusu için kuyrukları kurar — dashboard as4Render bucket mantığının
    Python karşılığı. Döner: (launchQ, cfiQ, copQ, sabah_kontrol).
    KURALLAR (UI ile aynı): zaten_teyitli 'kesin' → dokunma; 'olasi'/zombi/varyant →
    sabah kontrol; işaret 'kontrol' → elle ilgilenildi, dokunma; işaret 'teyit_ver' →
    kod DOĞRU, CFI kuyruğuna zorla (4536B vakası). Bugün DAHİL (17:00'da gün bitti)."""
    import sys as _sys
    _d = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'as400')
    if _d not in _sys.path:
        _sys.path.insert(0, _d)
    import launch_esle as _le
    coklu = _le.esle_coklu(tarihler)
    gun_isaret = {}
    for t in tarihler:
        gun_isaret[t] = {r['referans']: r['durum'] for r in conn.execute(
            "SELECT referans, durum FROM as400_teyit_isaret WHERE kapsam='gun' AND uretim_tarihi=?",
            (t,)).fetchall()}
    cop_verildi = {f"{r['uretim_tarihi']}|{r['launch_no']}" for r in conn.execute(
        "SELECT uretim_tarihi, launch_no FROM as400_teyit_log WHERE yil='CO' AND sonuc='ok'").fetchall()}

    launchQ, cfiQ, copQ, sabah = [], [], [], []

    def isr(t, r):
        return gun_isaret.get(t, {}).get(_le.kanonik(r.get('referans', '')), '')

    def sabaha(tip, t, r, neden, launch=''):
        sabah.append({'tip': tip, 'tarih': t, 'bolum': r.get('bolum', ''),
                      'referans': r.get('referans', ''), 'adet': r.get('adet', 0),
                      'launch': launch, 'neden': neden})

    def cfiye(t, r, article=''):
        cfiQ.append({'referans': r['referans'], 'article': (article or r['referans']),
                     'adet': int(round(float(r.get('adet') or 0))), 'uretim_tarihi': t,
                     'bolum': r.get('bolum', '')})

    def kapasite_asti(t, r):
        """Kapasite aşımı → oto koşuda HİÇBİR kuyruğa girmez, sabah kontrole düşer.
        UI ile aynı kural (2026-07-30): cycle time tanımlı + vardiya süresinde
        fiziken üretilemeyecek adet. '➕ teyit ver' işareti kullanıcı onayıdır."""
        k = r.get('kapasite_asim')
        if not k or isr(t, r) == 'teyit_ver':
            return False
        sabaha('kapasite', t, r,
               f"ADET ŞÜPHELİ: {k.get('adet')} adet × {k.get('ct')} sn = "
               f"{k.get('gerekli_saat')} saat gerekir, vardiya "
               f"{round((k.get('sure_dk') or 0) / 60.0, 1)} saat — makul üst sınır "
               f"~{k.get('azami')} adet. Operatör kaydını düzeltin ya da ➕ teyit ver ile onaylayın.",
               (( r.get('launchlar') or [{}])[0] or {}).get('launch', ''))
        return True

    for t in tarihler:
        L = coklu.get(t) or {}
        for r in L.get('ACIK', []):
            ls = r.get('launchlar') or []
            if any(l.get('zaten_teyitli') == 'kesin' for l in ls):
                continue
            # SATIR düzeyi teyit: teyit KENDİ koduna CFI ile verilmiş olabilir —
            # launch'ın article'ında görünmez (2026-07-30, 6175B vakası).
            if r.get('zaten_teyitli') == 'kesin':
                continue
            if kapasite_asti(t, r):
                continue
            im = isr(t, r)
            if im == 'kontrol':
                continue
            temiz = [l for l in ls if not l.get('zombi')]
            if not temiz:
                sabaha('zombi', t, r, 'Launch kalanı ≤ 0 (S unutulmuş olabilir) — elle S ile kapatın',
                       (ls[0] or {}).get('launch', '') if ls else '')
                continue
            # 'olasi' şüphesi kullanıcı ➕ teyit ver işaretiyle EZİLEBİLİR (karar insanın)
            # Launch article'ı VEYA satırın kendi kodu 'olasi' ise elle kontrol.
            if ((any(l.get('zaten_teyitli') == 'olasi' for l in ls)
                 or r.get('zaten_teyitli') == 'olasi') and im != 'teyit_ver'):
                sabaha('olasi', t, r, 'Kısmi/uyuşmayan teyit görünüyor — elle kontrol edin',
                       temiz[0].get('launch', ''))
                continue
            l0 = temiz[0]
            adet = int(round(float(r.get('adet') or 0)))
            if adet <= 0:
                continue
            bayrak = 'S' if adet >= float(l0.get('kalan') or 0) - 1 else 'A'
            launchQ.append({'yil': l0['red2'], 'no': l0['renu'], 'article': l0['article'],
                            'referans': r['referans'], 'adet': adet, 'uretim_tarihi': t,
                            'bayrak': bayrak})
        for r in L.get('SUPHELI', []):
            ls = r.get('launchlar') or []
            if any(l.get('zaten_teyitli') == 'kesin' for l in ls):
                continue
            # SATIR düzeyi teyit: varyantta teyit KENDİ koduna CFI ile verilmiş
            # olabilir (2026-07-30, 6175B) — launch article'ında görünmez.
            if r.get('zaten_teyitli') == 'kesin':
                continue
            if kapasite_asti(t, r):
                continue
            im = isr(t, r)
            if im == 'kontrol':
                continue
            if r.get('zaten_teyitli') == 'olasi' and im != 'teyit_ver':
                sabaha('olasi', t, r, 'Kısmi/uyuşmayan teyit görünüyor — elle kontrol edin',
                       (ls[0] or {}).get('launch', '') if ls else '')
                continue
            if im == 'teyit_ver':
                cfiye(t, r)   # kullanıcı 'kod doğru' dedi → kendi koduna CFI
                continue
            sabaha('varyant', t, r, 'Yazım/varyant farkı — panelden 🔁 düzelt ya da ➕ teyit ver',
                   (ls[0] or {}).get('launch', '') if ls else '')
        for kat in ('OPR10', 'YOK'):
            for r in L.get(kat, []):
                if r.get('zaten_teyitli') == 'kesin':
                    continue
                if kapasite_asti(t, r):
                    continue
                im = isr(t, r)
                if im == 'kontrol':
                    continue
                if r.get('zaten_teyitli') == 'olasi' and im != 'teyit_ver':
                    sabaha('olasi', t, r, 'Kısmi/uyuşmayan teyit görünüyor — elle kontrol edin')
                    continue
                art = ''
                if kat == 'OPR10':
                    ls = r.get('launchlar') or []
                    art = (ls[0] or {}).get('article', '') if ls else ''
                cfiye(t, r, art)
        for r in L.get('KAPALI', []):
            im = isr(t, r)
            if r.get('zaten_teyitli') == 'kesin' or im == 'kontrol':
                continue
            if kapasite_asti(t, r):
                continue
            if im == 'teyit_ver':
                cfiye(t, r)
            # değilse: TK1 akışı (durum 45/50) — bilgi, sabah listesini şişirme
        for r in L.get('ARAOP', []):
            if isr(t, r) == 'teyit_ver':
                cfiye(t, r)
        # ── Hurda → COP (HARIC ve ARAOP hariç tüm kovalar) ──
        for kat, rows in L.items():
            if kat in ('HARIC', 'ARAOP'):
                continue
            for r in rows:
                try:
                    h = int(round(float(r.get('hurda') or 0)))
                except (TypeError, ValueError):
                    h = 0
                if h <= 0:
                    continue
                ls = r.get('launchlar') or []
                if kat == 'SUPHELI':
                    sabaha('cop_supheli', t, r, f'Hurda {h} adet: varyant eşleşme — COP kodunu doğrulayıp elle gönderin')
                    continue
                art = ((ls[0] or {}).get('article') or r['referans']) if ls else r['referans']
                if f'{t}|{art}' in cop_verildi:
                    continue
                copQ.append({'referans': r['referans'], 'article': art, 'adet': h,
                             'uretim_tarihi': t, 'bolum': r.get('bolum', '')})
    return launchQ, cfiQ, copQ, sabah


def oto_teyit_job():
    """17:10 — günün tüm teyit kuyruğu: launch (A/S) → CFI → COP."""
    cfg = _oto_config()['oto_teyit']
    if not cfg.get('etkin'):
        print('[OTO-TEYIT] kapalı (oto_config) — atlandı')
        _oto_atlandi_kaydet('teyit', 'ATLANDI: oto teyit KAPALI (oto_config)',
                            'Ayar kapalı — Sabah Kontrol panelinden aç.')
        return
    if not _teyit_agent_var():
        print('[OTO-TEYIT] teyit-agent yok — bu makine robot koşusu için yapılandırılmamış, atlandı')
        _oto_atlandi_kaydet('teyit', 'ATLANDI: teyit-agent kapalı (PCOMM/Session B yok)',
                            'teyit-agent 127.0.0.1:5010 yanıt vermedi — RDP oturumunda PCOMM+agent açık olmalı.')
        return
    gunler = max(1, min(7, int(cfg.get('gunler') or 3)))
    tarihler = [date.today().isoformat()] + [(date.today() - timedelta(days=i)).isoformat()
                                             for i in range(1, gunler + 1)]
    conn = db_connect()
    try:
        if not _kilit_bekleyerek_al(45 * 60, 'OTO-TEYIT'):
            _oto_kosu_kaydet(conn, 'teyit', [], [{'neden': 'AS400 kilidi 45 dk boşalmadı — koşu yapılamadı'}],
                             'HATA: kilit alınamadı')
            return
        try:
            _AS400_DURDUR.clear()
            try:
                launchQ, cfiQ, copQ, sabah = _oto_kuyruk_olustur(conn, tarihler)
            except Exception as e:
                _oto_kosu_kaydet(conn, 'teyit', [], [{'neden': f'Kuyruk oluşturulamadı: {e}'}],
                                 f'HATA: AS400 eşleme başarısız ({e})')
                return
            sonuclar = []
            for faz, kuyruk in (('launch', launchQ), ('cfi', cfiQ), ('cop', copQ)):
                if _AS400_DURDUR.is_set():
                    break
                if not kuyruk:
                    continue
                if faz == 'launch':
                    s = _teyit_gonder_calistir(conn, kuyruk, 'oto-17:10')
                elif faz == 'cfi':
                    s = _cfi_gonder_calistir(conn, kuyruk, 'oto-17:10')
                else:
                    s = _cop_gonder_calistir(conn, kuyruk, 'oto-17:10')
                sonuclar += [{**r, 'faz': faz} for r in s]
        finally:
            _AS400_DURDUR.clear()
            _AS400_KILIT.release()
        # Hatalı gönderimler de sabah kontrolüne düşsün
        for r in sonuclar:
            if r.get('sonuc') == 'hata':
                sabah.append({'tip': 'gonderim_hata', 'tarih': r.get('uretim_tarihi', ''),
                              'bolum': r.get('bolum', ''), 'referans': r.get('referans', ''),
                              'adet': r.get('adet', 0), 'launch': r.get('no', '') or r.get('article', ''),
                              'neden': f"{r.get('faz', '')}: {r.get('mesaj', '')}"})
        ok_n = sum(1 for r in sonuclar if r.get('sonuc') == 'ok')
        atl_n = sum(1 for r in sonuclar if r.get('sonuc') == 'atlandi')
        _oto_kosu_kaydet(conn, 'teyit', sonuclar, sabah,
                         f'{len(launchQ)} launch + {len(cfiQ)} CFI + {len(copQ)} COP kuyruğa alındı: '
                         f'{ok_n} ok, {atl_n} atlandı'
                         + (f', {len(sabah)} satır sabah kontrole ayrıldı' if sabah else ''))
    finally:
        conn.close()


@app.route('/api/as400/oto_kosu', methods=['GET'])
@panel_gerekli(izin='as400-teyit')
def as400_oto_kosu():
    """Sabah Kontrol paneli: son otomatik koşular (16:45 transfer + 17:10 teyit)."""
    try:
        limit = max(1, min(30, int(request.args.get('limit') or 6)))
    except (TypeError, ValueError):
        limit = 6
    conn = get_db()
    kosular = []
    try:
        for r in conn.execute(
                "SELECT id, tur, ozet, detay, ok, hata, atlandi, kontrol, created_at "
                "FROM as400_oto_kosu ORDER BY id DESC LIMIT ?", (limit,)).fetchall():
            d = dict(r)
            try:
                d['detay'] = json.loads(d.get('detay') or '{}')
            except Exception:
                d['detay'] = {}
            kosular.append(d)
    except Exception as e:
        return jsonify({'hata': str(e)}), 500

    # ── ÇÖZÜLMÜŞ SATIRLARI İŞARETLE (2026-07-30, kullanıcı: "kontrol kısmında
    # düzeltilen şeyleri göstermeyelim") ─────────────────────────────────────
    # Sabah kontrol listesi koşu ANINDA donuyor: o gün emin olunamayan satırlar
    # JSON'a yazılıyor ve bir daha güncellenmiyordu. Sonradan teyit verilse veya
    # kullanıcı "gerek yok"/"kontrol edildi" işaretlese bile satır listede
    # kalıyor, panel çözülmüş işleri bekliyormuş gibi gösteriyordu.
    # Burada her satır CANLI durumla karşılaştırılır; çözülmüşler cozuldu=True
    # alır ve panelde gizlenir (sayısı ayrıca bildirilir — sessiz kesme yok).
    try:
        import sys as _sys
        _d = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'as400')
        if _d not in _sys.path:
            _sys.path.insert(0, _d)
        import launch_esle as _le

        # (üretim günü, kanonik kod) → o güne o koda BAŞARILI gönderim yapılmış
        basarili = set()
        for r in conn.execute(
                "SELECT uretim_tarihi, referans, article FROM as400_teyit_log WHERE sonuc='ok'").fetchall():
            for kod in (r['referans'], r['article']):
                if kod:
                    basarili.add((r['uretim_tarihi'] or '', _le.kanonik(kod)))
        # kullanıcı işaretleri: kalıcı 'gerek_yok' + o güne özel her işaret
        kalici = _le.kalici_haric_set()
        gunluk = set()
        for r in conn.execute(
                "SELECT uretim_tarihi, referans FROM as400_teyit_isaret WHERE kapsam='gun'").fetchall():
            gunluk.add((r['uretim_tarihi'] or '', _le.kanonik(r['referans'] or '')))

        # ERP KANALI (2026-07-30, 94.ltk.704 olayı): robot hareketi ERP'ye YAZDIĞI
        # hâlde ekran doğrulaması tutmazsa log'a 'hata' düşer ve satır sonsuza
        # kadar "kontrol bekliyor" görünür. 704'te ERP'de 07-29 CFI 24 vardı ama
        # panel hata gösteriyordu. Kendi log'umuz tek kanal olamaz — ERP'ye de bak.
        # Sorgu yalnız listede GERÇEKTEN açık kalan kodlar için yapılır (pahalı).
        acik_adaylar = set()
        for k in kosular:
            for s in ((k.get('detay') or {}).get('sabah_kontrol') or []):
                kod = (s.get('referans') or s.get('article') or '').strip()
                kn = _le.kanonik(kod)
                if kod and kn and kn not in kalici and (s.get('tarih') or '', kn) not in basarili \
                        and (s.get('tarih') or '', kn) not in gunluk:
                    # ERP sorgusu TAM eşleşme; operatör kodu küçük harfle yazmış
                    # olabilir (94.ltk.704 ⇄ ERP'de 94.LTK.704). Her iki yazımı da
                    # sor — sonuç zaten kanonik anahtarla toplanıyor.
                    acik_adaylar.add(kod)
                    acik_adaylar.add(kod.upper())
        erp_hrk = {}
        if acik_adaylar:
            try:
                erp_hrk = _le.teyit_hareketleri(sorted(acik_adaylar)) or {}
            except Exception as _e:
                print(f'[oto_kosu] ERP kanali atlandi: {_e}')

        for k in kosular:
            sabah = (k.get('detay') or {}).get('sabah_kontrol') or []
            acik = 0
            for s in sabah:
                kn = _le.kanonik(s.get('referans') or s.get('article') or '')
                t = s.get('tarih') or ''
                if not kn:
                    acik += 1
                    continue
                if kn in kalici:
                    s['cozuldu'], s['cozum'] = True, 'kalıcı olarak "teyit gerekmez" işaretli'
                elif (t, kn) in basarili:
                    s['cozuldu'], s['cozum'] = True, 'bu üretim günü için teyit sonradan verilmiş'
                elif (t, kn) in gunluk:
                    s['cozuldu'], s['cozum'] = True, 'panelden elle işaretlenmiş'
                else:
                    # ERP'de üretim gününde VEYA sonrasında bu adetle hareket var mı?
                    _h = [x for x in (erp_hrk.get(kn) or [])
                          if (not t or x.get('tarih', '') >= t)
                          and abs(float(x.get('adet') or 0) - float(s.get('adet') or 0)) < 0.001]
                    if _h:
                        s['cozuldu'] = True
                        s['cozum'] = (f"ERP'de kayıtlı: {_h[0]['tarih']} · {_h[0].get('tur', '')} "
                                      f"{_h[0]['adet']:g} adet (robot yazdı, ekran doğrulaması tutmamış)")
                    else:
                        acik += 1
            k['acik_kontrol'] = acik
            k['cozulen_kontrol'] = len(sabah) - acik
    except Exception as e:
        # Tespit edilemezse hiçbir satır gizlenmez (güvenli yön: eksik gösterme)
        print(f'[oto_kosu] cozuldu tespiti atlandi: {e}')

    return jsonify({'kosular': kosular, 'config': _oto_config(), 'agent': _teyit_agent_var()})


@app.route('/api/as400/oto_config', methods=['POST'])
@panel_gerekli(izin='as400-teyit')
def as400_oto_config_degistir():
    """Oto koşu aç/kapat (+ transfer dryrun). Body: {tur, etkin?, dryrun?}.
    Saat değişikliği DOSYADAN yapılır ve restart ister (thread saati startta okur)."""
    data = request.get_json(silent=True) or {}
    tur = (data.get('tur') or '').strip()
    if tur not in _OTO_VARSAYILAN:
        return jsonify({'hata': f'geçersiz tur: {tur}'}), 400
    cfg = _oto_config()
    if 'etkin' in data:
        cfg[tur]['etkin'] = bool(data.get('etkin'))
    if 'dryrun' in data and tur == 'transfer_iptal':
        cfg[tur]['dryrun'] = bool(data.get('dryrun'))
    try:
        _oto_config_yaz(cfg)
    except Exception as e:
        return jsonify({'hata': f'config yazılamadı: {e}'}), 500
    return jsonify({'basarili': True, 'config': cfg})


# ─────────────────────────────────────────────────────────────

# BAŞLATMA
# ─────────────────────────────────────────────────────────────

# Modul yüklenirken DB migration'larını çalıştır (debug auto-reload sonrası da
# eksik kolonların eklendiğinden emin olmak için).
init_db()


if __name__ == '__main__':
    # Windows konsolu (cp1254) Unicode sembolleri (→ • ✓ emoji) encode EDEMEZ →
    # print() UnicodeEncodeError fırlatır; bu bir yanıt yolunda olursa (örn. mail
    # başarı logu) BAŞARI, HATA gibi görünür. errors='replace' → print ASLA çökmesin.
    try:
        import sys as _sysio
        _sysio.stdout.reconfigure(errors='replace')
        _sysio.stderr.reconfigure(errors='replace')
    except Exception:
        pass

    # Flask debug modundayken (reloader) çift çalışmayı önlemek için kontrol
    if os.environ.get('WERKZEUG_RUN_MAIN') != 'true':
        init_db()

    # Scheduler — sadece bir kez başlat (reloader child process'inde başlat,
    # debug modda da main process'te çift olmasın)
    if not os.environ.get('FLASK_DEBUG') or os.environ.get('WERKZEUG_RUN_MAIN') == 'true':
        try:
            from scheduler import start_scheduler
            # AS400 oto koşuları (2026-07-27): saatler oto_config.json'dan.
            # Jobs kendi içinde etkin/agent kontrolü yapar (kapatmak restart istemez;
            # saat DEĞİŞİKLİĞİ restart ister — thread saati burada okur).
            _ek = []
            _ocfg = _oto_config()
            # kosu_turu: as400_oto_kosu.tur — telafi kontrolü bu kayda bakar
            for _tur, _fn, _et, _vs, _kt in (
                    ('transfer_iptal', oto_transfer_iptal_job, 'AS400 Transfer İptali (02→01)', (16, 45), 'transfer'),
                    ('oto_teyit', oto_teyit_job, 'AS400 Oto Teyit Kuyruğu', (17, 10), 'teyit')):
                try:
                    _h, _m = (int(x) for x in str(_ocfg[_tur].get('saat') or '').split(':'))
                except (TypeError, ValueError):
                    _h, _m = _vs
                # TELAFİ: uygulama planlı saatte ayakta değilse (restart/çökme) o gün
                # sessizce kaybolurdu. Açılışta "bugün hiç koşmadıysa bir kez çalıştır".
                # Güvenli: her iki iş de kendi içinde mükerrer korumalı (as400_teyit_log
                # 'ok' kaydı olan launch tekrar gönderilmez), ayrıca kapalıysa zaten atlar.
                _ek.append((_h, _m, _fn, _et, _oto_telafi_gerek_mi(_kt)))
            start_scheduler(ek_gorevler=_ek)
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
    # GUVENLIK: internete acik sistemde debug=True = uzaktan kod calistirma (RCE).
    # Varsayilan KAPALI; yerel gelistirmede ortam degiskeni ile ac: COFLE_DEBUG=1
    # threaded=True ŞART: teyit gönderimi bir robot çağrısında 15-150sn bloklar;
    # tek-thread'de bu sürede TÜM app donar (Durdur yoklaması, operatör mobil, andon
    # cevap alamaz — "sistem durdu" belirtisi). Çok-thread'de web responsive kalır,
    # robot yine _AS400_KILIT ile tek tek çalışır. (get_db flask.g request-scoped →
    # her thread kendi bağlantısını açar, SQLite thread-güvenli; yazımlar busy-timeout.)
    app.run(host='0.0.0.0', port=5000, threaded=True, debug=(os.environ.get('COFLE_DEBUG') == '1'))
