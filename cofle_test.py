# -*- coding: utf-8 -*-
"""
cofle_test.py — Cofle test cihazı (wicow) entegrasyonu.

Harici test sistemi (beta.coflesvp.com → API: apibeta.wicow.io/cofle) operatörlerin
ürün test cihazlarını ve yaptıkları testleri tutar. Bu modül o API'den "başarılı test"
kayıtlarını çekip yerel `test_sonuclari` tablosuna yazar; böylece bir test cihazı,
ESP32 pulse sensörü gibi bir SAYAÇ KAYNAĞI olur (başarılı test = 1 pulse).

Mimari (pilot.db/sayac_olaylari ile birebir paralel):
  • Arka plan poller (scheduler.py) aktif test-cihazı oturumlarının cihazlarını
    periyodik çekip test_sonuclari'na `test_id` (API'nin benzersiz id'si) ile UPSERT eder.
  • app.py `_test_basari_say()` referans başlangıcından beri PASSED kayıtları sayar.
  • Operatör mobilde "Test cihazı ile çalışıyorum" seçerse üretim kaydı test_cihaz_id
    taşır ve sayaç bu kaynaktan beslenir.

API erişimi cofle_test_config.json'dan (gitignore'lu) okunur: api_base + Bearer token.
Tüm ağ çağrıları hata-güvenli: başarısızsa None döner, ASLA exception fırlatmaz
(sayaç değeri ezilmesin / uygulama çökmesin).
"""
import os
import json
import time
import sqlite3
import urllib.request
import urllib.error

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(PROJECT_DIR, 'cofle_test_config.json')
DB_PATH     = os.path.join(PROJECT_DIR, 'uretim.db')

_cfg_cache = {'mtime': None, 'data': None}


def _config():
    """cofle_test_config.json'u oku (mtime cache). Dosya yoksa/bozuksa None."""
    try:
        st = os.stat(CONFIG_PATH)
    except OSError:
        return None
    if _cfg_cache['mtime'] != st.st_mtime:
        try:
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                _cfg_cache['data'] = json.load(f)
            _cfg_cache['mtime'] = st.st_mtime
        except Exception as e:
            print(f'[cofle_test] config okunamadı: {e}')
            return None
    return _cfg_cache['data']


def etkin():
    """Entegrasyon açık mı (config var + etkin=true + token dolu)."""
    c = _config()
    return bool(c and c.get('etkin') and c.get('token') and c.get('api_base'))


def poll_aralik_sn():
    c = _config() or {}
    try:
        return max(5, int(c.get('poll_aralik_sn', 20)))
    except (ValueError, TypeError):
        return 20


def _api_get(path, timeout=15):
    """GET <api_base><path> → parse JSON. Hata/timeout/auth sorununda None (sessiz)."""
    c = _config()
    if not c or not c.get('token') or not c.get('api_base'):
        return None
    url = c['api_base'].rstrip('/') + path
    req = urllib.request.Request(url, headers={
        'authorization': 'Bearer ' + c['token'],
        'accept': 'application/json',
        'origin': 'https://beta.coflesvp.com',
    })
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        # 401/403 = token süresi dolmuş olabilir → operatöre değil log'a düşür
        print(f'[cofle_test] API HTTP {e.code} ({path}) — token süresi dolmuş olabilir')
        return None
    except Exception as e:
        print(f'[cofle_test] API hatası ({path}): {e}')
        return None


def _norm(s):
    return (s or '').strip().upper().replace(' ', '')


def _basarili_mi(result):
    """API result alanı 'PASSED_164.20kg' / 'FAILED_...' biçiminde — PASSED ile başlıyorsa başarılı."""
    return 1 if str(result or '').strip().upper().startswith('PASSED') else 0


# ── Dropdown: cihaz listesi (overview) ──────────────────────────────
_liste_cache = {'ts': 0, 'data': None}


def cihaz_listesi(cache_sn=60):
    """Test cihazlarını [{'id', 'ad'}] olarak döner (operatör dropdown'u için).
    60sn cache — her form açılışında API'yi dövmemek için. Hata olursa son
    başarılı listeyi (veya boş) döner."""
    now = time.time()
    if _liste_cache['data'] is not None and (now - _liste_cache['ts']) < cache_sn:
        return _liste_cache['data']
    data = _api_get('/test-logs')
    if data is None:
        return _liste_cache['data'] or []
    liste = []
    for d in data:
        if not isinstance(d, dict) or d.get('device_id') is None:
            continue
        liste.append({'id': d['device_id'], 'ad': (d.get('device_name') or str(d['device_id'])).strip()})
    liste.sort(key=lambda x: x['ad'].lower())
    _liste_cache['data'] = liste
    _liste_cache['ts'] = now
    return liste


def cihaz_adi_haritasi():
    """device_id → device_name haritası (poll'da kayıtlara isim yazmak için)."""
    return {d['id']: d['ad'] for d in cihaz_listesi(cache_sn=300)}


def cihaz_detay(cihaz_id):
    """Bir cihazın test kayıtları (ham). Hata olursa None."""
    return _api_get(f'/test-logs/{int(cihaz_id)}/details')


# ── Poller: aktif cihazların testlerini yerel tabloya çek ───────────
def aktif_cihaz_idleri(db_path=None):
    """Açık (bugün) bir vardiyada, test_cihaz_id ile çalışan üretim kayıtlarının
    benzersiz cihaz id'leri. Sadece bunları poll ederiz (API'yi gereksiz yormayız)."""
    db_path = db_path or DB_PATH
    try:
        conn = sqlite3.connect(db_path, timeout=5.0)
        try:
            rows = conn.execute(
                "SELECT DISTINCT u.test_cihaz_id "
                "FROM uretim_kayitlari u JOIN vardiyalar v ON v.id = u.vardiya_id "
                "WHERE u.test_cihaz_id IS NOT NULL "
                "  AND u.sayac_otomatik = 1 "
                "  AND COALESCE(v.durum,'aktif') IN ('aktif','acik') "
                "  AND v.tarih = date('now','localtime')"
            ).fetchall()
            return [r[0] for r in rows if r[0] is not None]
        finally:
            conn.close()
    except Exception as e:
        print(f'[cofle_test] aktif cihaz sorgusu hatası: {e}')
        return []


def upsert_kayitlar(conn, cihaz_id, cihaz_adi, kayitlar):
    """Bir cihazın test kayıtlarını test_sonuclari'na yaz (test_id ile dedup)."""
    eklenen = 0
    for r in kayitlar or []:
        tid = r.get('id')
        if tid is None:
            continue
        ref = r.get('test_conf_name') or ''
        op = ((r.get('operator_name') or '').strip() + ' ' + (r.get('operator_surname') or '').strip()).strip()
        cur = conn.execute(
            "INSERT OR IGNORE INTO test_sonuclari "
            "(test_id, cihaz_id, cihaz_adi, referans_kodu, referans_norm, sonuc, basarili, ts_epoch, operator) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (tid, cihaz_id, cihaz_adi, ref, _norm(ref), r.get('result') or '',
             _basarili_mi(r.get('result')), r.get('timestamp') or 0, op)
        )
        eklenen += cur.rowcount
    return eklenen


def poll(db_path=None):
    """Aktif test cihazlarının yeni testlerini çekip test_sonuclari'na yazar.
    scheduler.py periyodik çağırır. Hata-güvenli."""
    if not etkin():
        return {'cihaz': 0, 'eklenen': 0, 'pasif': True}
    db_path = db_path or DB_PATH
    ids = aktif_cihaz_idleri(db_path)
    if not ids:
        return {'cihaz': 0, 'eklenen': 0}
    ad_harita = cihaz_adi_haritasi()
    toplam = 0
    try:
        conn = sqlite3.connect(db_path, timeout=10.0)
        conn.execute('PRAGMA busy_timeout=5000')
        try:
            for cid in ids:
                kayitlar = cihaz_detay(cid)
                if kayitlar is None:
                    continue  # API hatası — bu cihazı atla, sonra tekrar denenir
                toplam += upsert_kayitlar(conn, cid, ad_harita.get(cid, str(cid)), kayitlar)
            conn.commit()
        finally:
            conn.close()
    except Exception as e:
        print(f'[cofle_test] poll DB hatası: {e}')
    if toplam:
        print(f'[cofle_test] poll: {len(ids)} cihaz, {toplam} yeni test kaydı')
    return {'cihaz': len(ids), 'eklenen': toplam}
