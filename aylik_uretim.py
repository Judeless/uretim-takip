# -*- coding: utf-8 -*-
"""
aylik_uretim.py — AYLIK ÜRETİM RAPORU: teyidi verilmiş üretim, Forge bölümlerine göre
(kullanıcı 2026-10-09: "aylık üretilen, üretim teyidi verilen referansları bölümlere göre
ayırıp Excel hâlinde raporlayacağız … Forge bölümleri olsun, performans da eklensin,
çalışma saatini görelim; yönetici ile paylaşalım").

PLANLAMANIN ESKİ YÖNTEMİ
  · Q:\\UretimPlanlama\\EMRE\\Yeni klasör\\Üretim\\Yıllık Üretim\\2025 yıllık Üretim.xlsx —
    AS400 10-05-03-15 raporu (aylık), RPR ve CFI hareketleri; depo 01 = TK1, 01D = TK2.
  · Q:\\UretimPlanlama\\Aylık Kapasite Sunum\\Kapasite Kullanım Oranı 2026.xlsx — ayın kod
    × adet listesi 'Database' (kod → üretim hattı + saatlik adet) ile eşlenir; Summary:
    adet · teorik süre (adet ÷ saatlik adet) · çalışma saati ('Çalışma Saati' sayfası) ·
    performans (hedef %90). Database'de olmayan kod özete GİRMİYORDU (Ağustos 2026:
    677 kodun 226'sı, adedin %45'i; pres/abkant hiç yoktu).

FORGE
  · Adet: AS400 BMMAF0 — MGCACD 'RPR' (launch teyidi) + 'CFI' (launch'sız üretim girişi),
    HAREKET tarihine göre ay (10-05-03-15 ile aynı). Tesis hareketin DEPOSUNDAN
    (01D → TK2, 01/02 → TK1); bilinmeyen depoda kodun Forge tesisi.
  · Bölüm: FORGE (referans_listesi; tel'de adım ekli satırlar kökten eşlenir). Kod bir
    tesiste birden çok bölümde tanımlıysa her bölümün işine sayılır (o bölüm o koda iş
    yaptı); tesis toplamı kodu BİR kez sayar. Forge'da yoksa planlamanın Anaveri'si
    (TK-1/2 · Hat), o da yoksa 'Tanımsız' listesi.
  · Teorik süre (kapasite modülüyle AYNI öncelik): kapasite Excel süresi (kapasite_sure)
    → Forge çevrim süresi (tel'de adımların toplamı) → ERP rota süresi → yok.
  · Çalışma saati: kapasite Excel'inin 'Çalışma Saati' sayfası (NçS + %90/%95/%100 mesai,
    bölüm adıyla). PP / Pull saatleri TK1 TEL bölümüne yazılır (PP telleri Forge'da tel).
  · Performans = teorik süre ÷ toplam çalışma saati (NçS + mesai). Planlamanın eski
    tablosu yalnız NçS'ye bölüyordu — o değer de ayrı sütunda.
  · 20.* / 21.* hammadde (metre, kg) adet toplamına girmez — ayrı listede.
  · COP (hurda, kullanıcı 2026-10-12: "bölümlere göre cop adetlerini de getirelim … aylık
    üretime göre fire oranını da görelim"): ayın COP hareketleri kodun ÜRETİMİYLE aynı
    bölüme yazılır; adet (üretim) toplamına GİRMEZ. Fire % = COP ÷ üretilen adet (RPR+CFI).
    Ay boyunca hiç üretimi olmayan kodun COP'u yine bölümüne yazılır, kod sayısını artırmaz.

AS400'e YALNIZ SELECT gider. Ayın okuması saklanır (aylik_uretim_cekim); sınıflama ve
süreler rapor anında güncel tanımlarla uygulanır.
"""
import json
import os
import re
import threading
import time
from datetime import date, datetime, timedelta

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
HAREKET_TABLO = 'tkc0301F.BMMAF0'
CAUSALLAR = ('RPR', 'CFI', 'COP')   # COP = hurda; üretim adedine girmez
DEPO_TESIS = {'01D': 'TK2', '01': 'TK1', '02': 'TK1'}
HAMMADDE_ONEK = ('20.', '21.')
HEDEF_PERFORMANS = 0.90
# Excel süresi KULLANILMAYAN bölümler (kapasite modülüyle aynı kural, kullanıcı 2026-09-17): metalde
# Excel'in saatlik adedi ÇEVRİM başına, Forge'unki parça başına (kalıp gözü) — Forge süresi önce gelir.
SURE_EXCEL_DISI = ('metal',)
VARSAYILAN_KAPASITE_KLASORU = r'Q:\UretimPlanlama\Aylık Kapasite Sunum'
# Kapasite Excel'i 'Çalışma Saati' BÖLÜM adı → Forge (tesis, bölüm)
CALISMA_BOLUM = {
    'WELDING': ('TK2', 'kaynak'), 'LASER CUTTING': ('TK2', 'lazer'), 'METAL INJECTION': ('TK2', 'metal'),
    'ASSEMBLY': ('TK2', 'montaj'), 'ABKANT': ('TK2', 'pres'),
    'PP': ('TK1', 'tel'), 'PULL': ('TK1', 'tel'),
    'IVECO': ('TK1', 'montaj'), 'JKP,JKW': ('TK1', 'montaj'), 'LF,LFP,LAG (LTK),LSB,PBL': ('TK1', 'montaj'),
    'PLASTIC INJECTION': ('TK1', 'plastik'),
}
# Anaveri (TK-1/2 · Hat) → Forge bölümü — yalnız kod Forge'da HİÇ yoksa
ANAVERI_BOLUM = {
    ('TK-2', 'MONTAJ'): ('TK2', 'montaj'), ('TK-2', 'TRAKTÖR'): ('TK2', 'montaj'), ('TK-2', 'PEDAL'): ('TK2', 'montaj'),
    ('TK-2', 'PROJE'): ('TK2', 'montaj'), ('TK-2', 'KAYNAK'): ('TK2', 'kaynak'), ('TK-2', 'LAZER KESIM'): ('TK2', 'lazer'),
    ('TK-2', 'LAZER KESİM'): ('TK2', 'lazer'), ('TK-2', 'METAL ENJEKSIYON'): ('TK2', 'metal'),
    ('TK-2', 'METAL ENJEKSİYON'): ('TK2', 'metal'), ('TK-2', 'SAC BÜKÜM'): ('TK2', 'pres'),
    ('TK-2', 'İŞLEME'): ('TK2', 'isleme'), ('TK-2', 'ISLEME'): ('TK2', 'isleme'),
    ('TK-1', 'PP'): ('TK1', 'tel'), ('TK-1', 'KESIM'): ('TK1', 'tel'), ('TK-1', 'KESİM'): ('TK1', 'tel'),
    ('TK-1', 'LF'): ('TK1', 'montaj'), ('TK-1', 'JOYSTICK'): ('TK1', 'montaj'), ('TK-1', 'IVECO'): ('TK1', 'montaj'),
    ('TK-1', 'PLASTIK ENJEKSIYON'): ('TK1', 'plastik'), ('TK-1', 'PLASTIK ENJEKSİYON'): ('TK1', 'plastik'),
    # Tabo hortumları (92.*) Forge'da TK1 Montaj
    ('TABO', 'B. HOSE'): ('TK1', 'montaj'),
}
# FASON / TEDARİKÇİ (kullanıcı 2026-10-09: "Pull telleri şu anda fasonda üretiliyor"): Anaveri'de
# Pandora / Pull tellerinin RPR-CFI'ı tedarikçinin işidir — TK1 Tel'e sayılsaydı teorik süresi TK1
# çalışanlarının saatine bölünür, performansı şişirirdi. Ayrı satır, performansa girmez.
FASON = 'fason'
# RAPORA DAHİL EDİLMEYEN bölümler (kullanıcı 2026-11-03: "işleme ve fason kısımlarını bu rapora dahil
# etmeyelim"): yalnız bunlarda tanımlı kodlar özete, tesis toplamına, performansa ve Forge
# karşılaştırmasına girmez; kaç kod / adet olduğu notta yazar.
RAPOR_DISI = ('isleme', FASON)


def _w_teli(kod):
    """Kodu W ile biten tel (93.*) fason telidir (kullanıcı 2026-10-13) — Forge tanımını da ezer."""
    k = str(kod or '').strip().upper()
    return k.startswith('93.') and k.endswith('W')


def _fason_mi(a, kesin=True):
    """kesin=True: Pandora / Pull (Forge tanımını da ezer). False: Anaveri 'Fason' (yalnız
    Forge'da tanım yoksa — Forge'da tanımlıysa içeride üretiyoruz demektir)."""
    if not a:
        return False
    tk, hat = (a.get('tk') or '').upper(), (a.get('hat') or '').upper()
    if kesin:
        return tk == 'PANDORA' or hat == 'PULL'
    return tk == 'FASON' or hat == 'FASON'
BOLUM_AD = {'kaynak': 'Kaynak', 'montaj': 'Montaj', 'metal': 'Metal Enjeksiyon', 'lazer': 'Lazer Kesim',
            'pres': 'Pres / Abkant', 'isleme': 'İşleme', 'plastik': 'Plastik Enjeksiyon', 'tel': 'Tel Üretimi',
            'fason': 'Fason (tedarikçi)'}
BOLUM_SIRA = {'TK2': ('kaynak', 'lazer', 'pres', 'metal', 'isleme', 'montaj', 'fason'),
              'TK1': ('montaj', 'tel', 'plastik', 'fason')}
AY_AD = ('', 'Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz', 'Ağustos', 'Eylül', 'Ekim',
         'Kasım', 'Aralık')
_GUVENLI_KOD = re.compile(r'^[A-Za-z0-9./\- ]{3,40}$')
# AS400 açıklamalarında kontrol karakteri olabiliyor ('VW EL FREN TEL' + görünmez bayt): openpyxl
# bunları reddediyor ve TÜM Excel çıktısı düşüyordu (IllegalCharacterError). Küme openpyxl'inkiyle aynı.
_YASAK_KARAKTER = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')


_TR_ACIKLAMA = str.maketrans('"£$§°', 'ÜÖİŞÇ')     # EBCDIC(1026)→ODBC bozulması: '$§LEME GR' = 'İŞLEME GR'


def _temiz(v):
    return _YASAK_KARAKTER.sub('', str(v or '')).strip().translate(_TR_ACIKLAMA)


def tablolari_kur(conn):
    conn.execute("""CREATE TABLE IF NOT EXISTS aylik_uretim_cekim (
        id INTEGER PRIMARY KEY AUTOINCREMENT, yil INTEGER NOT NULL, ay INTEGER NOT NULL,
        ts TEXT NOT NULL, kullanici TEXT DEFAULT '', satir INTEGER DEFAULT 0, sure_sn REAL,
        veri TEXT NOT NULL)""")
    conn.execute("CREATE INDEX IF NOT EXISTS ix_auc_ay ON aylik_uretim_cekim(yil, ay)")
    conn.execute("""CREATE TABLE IF NOT EXISTS aylik_calisma (
        yil INTEGER NOT NULL, ay INTEGER NOT NULL, excel_bolum TEXT NOT NULL,
        nos REAL DEFAULT 0, mesai REAL DEFAULT 0, kisi INTEGER DEFAULT 0, isimsiz REAL DEFAULT 0,
        dosya TEXT DEFAULT '', kullanici TEXT DEFAULT '', guncellendi TEXT DEFAULT '',
        PRIMARY KEY (yil, ay, excel_bolum))""")
    # Kapasite Excel'inin 'Database' sayfası: kod → saatlik adet + üretim hattı (ayın değil, son yüklenen)
    conn.execute("""CREATE TABLE IF NOT EXISTS aylik_kapasite_db (
        kod TEXT PRIMARY KEY, saatlik REAL, hat TEXT DEFAULT '', makine TEXT DEFAULT '',
        dosya TEXT DEFAULT '', guncellendi TEXT DEFAULT '')""")
    # Rapordan hariç tutma kuralları (kullanıcı 2026-10-09: "92'li kodları dahil etmeyelim … kaldırabileyim").
    # Tüm aylara uygulanır; panelden eklenir / kaldırılır.
    yeni = not conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='aylik_uretim_haric'").fetchone()
    conn.execute("""CREATE TABLE IF NOT EXISTS aylik_uretim_haric (
        id INTEGER PRIMARY KEY AUTOINCREMENT, desen TEXT NOT NULL UNIQUE, tur TEXT NOT NULL DEFAULT 'kod',
        aciklama TEXT DEFAULT '', kullanici TEXT DEFAULT '', olusturma TEXT DEFAULT '')""")
    if yeni:        # tohum YALNIZ ilk kurulumda — kullanıcı kaldırırsa geri gelmez
        conn.execute("INSERT OR IGNORE INTO aylik_uretim_haric (desen, tur, aciklama, kullanici, olusturma) "
                     "VALUES ('92.', 'onek', ?, 'sistem', ?)",
                     ("92'li kodlar (Tabo el fren telleri)", datetime.now().strftime('%Y-%m-%d %H:%M')))
    conn.commit()


def _norm(k):
    return re.sub(r'\s+', '', str(k or '')).upper()


def _kok(k):
    return _norm(str(k or '').strip().split(' ')[0])


# ── RAPORDAN HARİÇ TUTMA ─────────────────────────────────────────────────────
HARIC_TURLER = ('onek', 'kod')
_HARIC_DESEN = re.compile(r'^[A-Z0-9./\-]{2,40}$')


def haric_kurallari(conn):
    """Önce tek kod kuralları (daha özel), sonra önekler."""
    tablolari_kur(conn)
    return [{'id': r[0], 'desen': r[1], 'tur': r[2], 'aciklama': r[3] or '', 'kullanici': r[4] or '',
             'olusturma': r[5] or ''} for r in conn.execute(
        "SELECT id, desen, tur, aciklama, kullanici, olusturma FROM aylik_uretim_haric ORDER BY tur, desen")]


def haric_ekle(conn, desen, tur, aciklama, kullanici):
    """desen: önek ('92.') ya da tam kod — raporla aynı normalize edilir (boşluksuz, büyük harf)."""
    n = _norm(desen)
    if tur not in HARIC_TURLER:
        raise ValueError("Tür 'onek' ya da 'kod' olmalı")
    if not _HARIC_DESEN.match(n):
        raise ValueError('Geçersiz kod / önek (2-40 karakter: harf, rakam, nokta, / ve -)')
    tablolari_kur(conn)
    if conn.execute("SELECT 1 FROM aylik_uretim_haric WHERE desen=?", (n,)).fetchone():
        raise ValueError(f'{n} zaten hariç tutuluyor')
    cur = conn.execute("INSERT INTO aylik_uretim_haric (desen, tur, aciklama, kullanici, olusturma) VALUES (?,?,?,?,?)",
                       (n, tur, str(aciklama or '').strip()[:200], kullanici or '',
                        datetime.now().strftime('%Y-%m-%d %H:%M')))
    conn.commit()
    return cur.lastrowid


def haric_sil(conn, kid):
    tablolari_kur(conn)
    n = conn.execute("DELETE FROM aylik_uretim_haric WHERE id=?", (kid,)).rowcount
    conn.commit()
    return n


def _haric_bul(n, kurallar):
    for k in kurallar:
        if (k['tur'] == 'kod' and n == k['desen']) or (k['tur'] == 'onek' and n.startswith(k['desen'])):
            return k
    return None


# ── AS400 ────────────────────────────────────────────────────────────────────
def as400_oku(cn, yil, ay):
    """Ayın RPR + CFI (+ COP hurda) hareketleri, kod × depo × causal toplamı + ürün açıklaması."""
    cu = cn.cursor()
    cu.execute(f"SELECT TRIM(MGARCD), TRIM(MGMGCD), TRIM(MGCACD), SUM(MGQTA), COUNT(*) FROM {HAREKET_TABLO} "
               f"WHERE MGCACD IN ({','.join(repr(c) for c in CAUSALLAR)}) AND MGDSSO=? AND MGDAAO=? AND MGDMMO=? "
               f"GROUP BY TRIM(MGARCD), TRIM(MGMGCD), TRIM(MGCACD)", (yil // 100, yil % 100, ay))
    hareket = [{'kod': r[0], 'depo': r[1] or '', 'causal': r[2], 'adet': float(r[3] or 0), 'satir': int(r[4] or 0)}
               for r in cu.fetchall() if r[0]]
    kodlar = sorted({h['kod'] for h in hareket if _GUVENLI_KOD.match(h['kod'])})
    aciklama = {}
    for i in range(0, len(kodlar), 50):
        grup = kodlar[i:i + 50]
        cu.execute("SELECT TRIM(A0ARTI), TRIM(A0ARDS) FROM TKC0301F.BARTF0 "
                   f"WHERE A0ARTI IN ({','.join('?' * len(grup))})", grup)
        aciklama.update({r[0]: r[1] or '' for r in cu.fetchall()})
    return {'yil': yil, 'ay': ay, 'ts': datetime.now().strftime('%Y-%m-%d %H:%M'),
            'hareket': hareket, 'aciklama': aciklama,
            'causallar': list(CAUSALLAR)}      # eski okumalarda yok → COP okunmamış demektir


def kaydet(conn, veri, kullanici, sure_sn=None):
    tablolari_kur(conn)
    cur = conn.execute("INSERT INTO aylik_uretim_cekim (yil, ay, ts, kullanici, satir, sure_sn, veri) "
                       "VALUES (?,?,?,?,?,?,?)",
                       (veri['yil'], veri['ay'], veri['ts'], kullanici, len(veri['hareket']), sure_sn,
                        json.dumps(veri, ensure_ascii=False)))
    # Ayın yalnız son 3 okuması tutulur
    conn.execute("DELETE FROM aylik_uretim_cekim WHERE yil=? AND ay=? AND id NOT IN (SELECT id FROM "
                 "aylik_uretim_cekim WHERE yil=? AND ay=? ORDER BY id DESC LIMIT 3)",
                 (veri['yil'], veri['ay'], veri['yil'], veri['ay']))
    conn.commit()
    return cur.lastrowid


def son_cekim(conn, yil, ay):
    tablolari_kur(conn)
    r = conn.execute("SELECT id, ts, kullanici, sure_sn, veri FROM aylik_uretim_cekim WHERE yil=? AND ay=? "
                     "ORDER BY id DESC LIMIT 1", (yil, ay)).fetchone()
    if not r:
        return None
    v = json.loads(r[4])
    v['_id'], v['_kullanici'], v['_sure_sn'] = r[0], r[2], r[3]
    return v


def cekilen_aylar(conn):
    tablolari_kur(conn)
    return [{'yil': r[0], 'ay': r[1], 'ts': r[2]} for r in conn.execute(
        "SELECT yil, ay, MAX(ts) FROM aylik_uretim_cekim GROUP BY yil, ay ORDER BY yil DESC, ay DESC")]


_DENEME = {'calisiyor': False, 'basladi': None, 'bitti': None, 'kim': '', 'hata': '', 'sonuc': None,
           'yil': None, 'ay': None}
_KILIT = threading.Lock()


def deneme_durumu():
    return dict(_DENEME)


def hazirla(conn_ac, erp_ac, yil, ay, kullanici):
    """Senkron: AS400'den oku + sakla."""
    if not _KILIT.acquire(blocking=False):
        raise RuntimeError('Aylık rapor zaten hazırlanıyor')
    t0 = time.time()
    _DENEME.update(calisiyor=True, basladi=datetime.now().strftime('%Y-%m-%d %H:%M:%S'), bitti=None,
                   kim=kullanici, hata='', sonuc=None, yil=yil, ay=ay)
    conn = conn_ac()
    try:
        cn = erp_ac()
        try:
            veri = as400_oku(cn, yil, ay)
        finally:
            try:
                cn.close()
            except Exception:
                pass
        cid = kaydet(conn, veri, kullanici, round(time.time() - t0, 1))
        _DENEME['sonuc'] = {'id': cid, 'satir': len(veri['hareket']),
                            'kod': len({h['kod'] for h in veri['hareket']})}
        return _DENEME['sonuc']
    except Exception as e:
        _DENEME['hata'] = str(e)
        raise
    finally:
        _DENEME.update(calisiyor=False, bitti=datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        try:
            conn.close()
        except Exception:
            pass
        _KILIT.release()


def hazirla_arka_planda(*args):
    if _DENEME['calisiyor'] or _KILIT.locked():
        return False
    _DENEME.update(calisiyor=True, basladi=datetime.now().strftime('%Y-%m-%d %H:%M:%S'), bitti=None, hata='',
                   sonuc=None)

    def _kos():
        try:
            hazirla(*args)
        except Exception as e:
            print(f'[AYLIK ÜRETİM] hazırlama hatası: {e}')
    threading.Thread(target=_kos, daemon=True, name='aylik-uretim').start()
    return True


# ── ÇALIŞMA SAATİ (kapasite Excel'i → 'Çalışma Saati' sayfası) ─────────────
def _bolum_adi(v):
    return ' '.join(str(v or '').strip().upper().split())


def calisma_oku(yol):
    """→ (ay, {excel_bolum: {'nos', 'mesai', 'kisi', 'isimsiz'}}) — İsim · Ay · NçS · %90 · %95 · %100 · BÖLÜM.
    Ayı yazılmamış (isimsiz ek) satırlar dosyadaki ASIL aya sayılır (kapasite modülüyle aynı kural)."""
    from openpyxl import load_workbook
    wb = load_workbook(yol, read_only=True, data_only=True)
    try:
        ad = next((s for s in wb.sheetnames if 'saat' in s.strip().lower()), None)
        if not ad:
            raise ValueError("Dosyada 'Çalışma Saati' sayfası yok — kapasite Excel'ini (Kapasite Kullanım Oranı) yükleyin")
        aylar, out = {}, {}
        for i, r in enumerate(wb[ad].iter_rows(values_only=True), start=1):
            if i == 1 or not r or len(r) < 7:
                continue
            bolum = _bolum_adi(r[6])
            if not bolum:
                continue
            f = lambda v: float(v) if isinstance(v, (int, float)) else 0.0   # noqa: E731
            nos, mesai = f(r[2]), f(r[3]) + f(r[4]) + f(r[5])
            if nos + mesai <= 0:
                continue
            try:
                a = int(float(r[1] or 0))
            except (TypeError, ValueError):
                a = 0
            if a:
                aylar[a] = aylar.get(a, 0) + 1
            d = out.setdefault(bolum, {'nos': 0.0, 'mesai': 0.0, 'kisi': 0, 'isimsiz': 0.0})
            d['nos'] += nos
            d['mesai'] += mesai
            if str(r[0] or '').strip():
                d['kisi'] += 1
            else:
                d['isimsiz'] += nos + mesai
        if not aylar:
            raise ValueError("'Çalışma Saati' sayfasında ay bilgisi bulunamadı")
        return max(aylar, key=aylar.get), out
    finally:
        wb.close()


def database_oku(yol):
    """Kapasite Excel'i 'Database' sayfası → [{kod, saatlik, hat, makine}] (yoksa [])."""
    from openpyxl import load_workbook
    wb = load_workbook(yol, read_only=True, data_only=True)
    try:
        ad = next((s for s in wb.sheetnames if s.strip().lower().startswith('database')), None)
        if not ad:
            return []
        out, gorulen = [], set()
        for i, r in enumerate(wb[ad].iter_rows(values_only=True), start=1):
            if i == 1 or not r or not r[0]:
                continue
            kod = _norm(r[0])
            if kod in gorulen:
                continue
            gorulen.add(kod)
            try:
                saatlik = float(r[2] or 0)
            except (TypeError, ValueError):
                saatlik = 0.0
            out.append({'kod': kod, 'saatlik': saatlik if saatlik > 0 else None,
                        'hat': _bolum_adi(r[3] if len(r) > 3 else ''), 'makine': str(r[1] or '').strip()})
        return out
    finally:
        wb.close()


def database_yaz(conn, satirlar, dosya):
    tablolari_kur(conn)
    if not satirlar:
        return 0
    simdi = datetime.now().strftime('%Y-%m-%d %H:%M')
    conn.execute("DELETE FROM aylik_kapasite_db")
    conn.executemany("INSERT OR REPLACE INTO aylik_kapasite_db (kod, saatlik, hat, makine, dosya, guncellendi) "
                     "VALUES (?,?,?,?,?,?)", [(x['kod'], x['saatlik'], x['hat'], x['makine'],
                                               os.path.basename(str(dosya)), simdi) for x in satirlar])
    conn.commit()
    return len(satirlar)


def calisma_yaz(conn, yil, ay, saatler, dosya, kullanici):
    tablolari_kur(conn)
    conn.execute("DELETE FROM aylik_calisma WHERE yil=? AND ay=?", (yil, ay))
    simdi = datetime.now().strftime('%Y-%m-%d %H:%M')
    conn.executemany("INSERT INTO aylik_calisma (yil, ay, excel_bolum, nos, mesai, kisi, isimsiz, dosya, kullanici, "
                     "guncellendi) VALUES (?,?,?,?,?,?,?,?,?,?)",
                     [(yil, ay, b, round(d['nos'], 1), round(d['mesai'], 1), d['kisi'], round(d['isimsiz'], 1),
                       os.path.basename(str(dosya)), kullanici, simdi) for b, d in saatler.items()])
    conn.commit()
    return len(saatler)


def en_yeni_kapasite(klasor=VARSAYILAN_KAPASITE_KLASORU):
    if not os.path.isdir(klasor):
        raise FileNotFoundError(f'Klasöre erişilemiyor: {klasor}')
    ad = [os.path.join(klasor, f) for f in os.listdir(klasor)
          if f.lower().startswith('kapasite') and f.lower().endswith(('.xlsx', '.xlsm')) and not f.startswith('~$')]
    if not ad:
        raise FileNotFoundError(f"'{klasor}' içinde kapasite Excel'i yok")
    return max(ad, key=os.path.getmtime)


def calisma_aylari(conn):
    tablolari_kur(conn)
    return [{'yil': r[0], 'ay': r[1], 'dosya': r[2]} for r in conn.execute(
        "SELECT yil, ay, MAX(dosya) FROM aylik_calisma GROUP BY yil, ay ORDER BY yil DESC, ay DESC")]


# ── ANA VERİ › 'Aylık NçS' (kullanıcı 2026-10-13) ──────────────────────────────
# "Aylık üretim raporu için de bizim Ana Veri Excel'imizden devam edelim; ek bir sayfaya
# aylık gerekli olan NçS'leri gireceğim, yeni ayda o kısmı değiştiririm." Kapasite
# Excel'inin kişi bazlı 'Çalışma Saati' yerine bölüm bazlı tablo: Ay hücresi + Tesis ·
# Bölüm · NçS · Mesai · Kişi. 'Ana Veri Yükle' onaylanınca o AYIN saatleri aylik_calisma'ya
# yazılır (ayın eski satırlarının — kapasite Excel'inden gelmiş olsalar da — yerine geçer);
# dosyada ay değişince önceki ayların saatleri veritabanında kalır.
NCS_SAYFA = 'Aylık NçS'
NCS_KAYNAK = 'Ana Veri › Aylık NçS'
NCS_BOLUM_AD = {('TK2', 'kaynak'): 'Kaynak', ('TK2', 'lazer'): 'Lazer Kesim',
                ('TK2', 'pres'): 'Büküm (Pres / Abkant)', ('TK2', 'metal'): 'Metal Enjeksiyon',
                ('TK2', 'isleme'): 'İşleme', ('TK2', 'montaj'): 'Montaj',
                ('TK1', 'montaj'): 'Montaj', ('TK1', 'tel'): 'Tel Üretimi', ('TK1', 'plastik'): 'Plastik Enjeksiyon'}
_AY_ADLARI = {'ocak': 1, 'şubat': 2, 'subat': 2, 'mart': 3, 'nisan': 4, 'mayıs': 5, 'mayis': 5, 'haziran': 6,
              'temmuz': 7, 'ağustos': 8, 'agustos': 8, 'eylül': 9, 'eylul': 9, 'ekim': 10, 'kasım': 11,
              'kasim': 11, 'aralık': 12, 'aralik': 12}


def _kucuk(v):
    return str(v if v is not None else '').replace('I', 'ı').replace('İ', 'i').lower().strip()


def _ncs_bolum(tesis, ad):
    """('TK2', 'Büküm (Pres / Abkant)') → ('TK2', 'pres') | None"""
    t = {'TK1': 'TK1', 'TK2': 'TK2'}.get(re.sub(r'[\s-]', '', str(tesis or '')).upper())
    a = _kucuk(ad)
    if not t or not a:
        return None
    for anahtar, b in (('plastik', 'plastik'), ('metal', 'metal'), ('lazer', 'lazer'), ('kaynak', 'kaynak'),
                       ('büküm', 'pres'), ('bukum', 'pres'), ('pres', 'pres'), ('abkant', 'pres'),
                       ('işleme', 'isleme'), ('isleme', 'isleme'), ('tel', 'tel'), ('montaj', 'montaj')):
        if anahtar in a:
            return (t, b) if (t, b) in NCS_BOLUM_AD else None
    return None


def _ncs_ay(v):
    """Ay hücresi → (yıl, ay) | None. Tarih · '2026-10' · '10.2026' · 'Ekim 2026'."""
    if isinstance(v, (datetime, date)):
        return v.year, v.month
    s_ = _kucuk(v)
    m = re.match(r'^(\d{4})\D+(\d{1,2})', s_)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.match(r'^(\d{1,2})\D+(\d{4})', s_)
    if m:
        return int(m.group(2)), int(m.group(1))
    m = re.match(r'^([a-zçğıöşü]+)\D*(\d{4})', s_)
    if m and m.group(1) in _AY_ADLARI:
        return int(m.group(2)), _AY_ADLARI[m.group(1)]
    return None


def ncs_sayfadan(ws):
    """Aylık NçS sayfası → {'yil', 'ay', 'saatler': {(tesis, bölüm): {nos, mesai, kisi}}, 'hatalar': [...]}.
    NçS'si boş satır 'girilmedi' sayılır (o bölüm yazılmaz)."""
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    out = {'yil': None, 'ay': None, 'saatler': {}, 'hatalar': []}
    for r in rows[:15]:
        for i, v in enumerate(r[:-1]):
            if _kucuk(v).rstrip(':') == 'ay':
                ya = _ncs_ay(r[i + 1])
                if ya and 1 <= ya[1] <= 12:
                    out['yil'], out['ay'] = ya
                else:
                    out['hatalar'].append(f"Ay hücresi okunamadı: {r[i + 1]!r} (ör. 2026-10)")
                break
        if out['yil'] or out['hatalar']:
            break
    if not out['yil'] and not out['hatalar']:
        out['hatalar'].append("'Ay' hücresi bulunamadı — A sütununa 'Ay', yanına 2026-10 yazın")
    bas, kol = None, {}
    for i, r in enumerate(rows[:30]):
        adlar = [_kucuk(v) for v in r]
        if 'tesis' in adlar and any(a.startswith(('bölüm', 'bolum')) for a in adlar):
            bas = i
            for j, a in enumerate(adlar):
                if a == 'tesis':
                    kol['tesis'] = j
                elif a.startswith(('bölüm', 'bolum')):
                    kol.setdefault('bolum', j)
                elif a.startswith(('nçs', 'ncs')):
                    kol.setdefault('nos', j)
                elif a.startswith('mesai'):
                    kol.setdefault('mesai', j)
                elif a.startswith(('kişi', 'kisi')):
                    kol.setdefault('kisi', j)
            break
    if bas is None or 'nos' not in kol:
        out['hatalar'].append("Tablo başlığı bulunamadı — 'Tesis', 'Bölüm', 'NçS' sütunları olmalı")
        return out

    def sayi(r, a):
        j = kol.get(a)
        v = r[j] if j is not None and j < len(r) else None
        if v in (None, ''):
            return None
        try:
            return float(str(v).replace(',', '.'))
        except ValueError:
            raise ValueError(f'{a} sayı değil: {v!r}')
    for no, r in enumerate(rows[bas + 1:], start=bas + 2):
        t = r[kol['tesis']] if kol['tesis'] < len(r) else None
        b = r[kol['bolum']] if kol['bolum'] < len(r) else None
        if not t and not b:
            continue
        hedef = _ncs_bolum(t, b)
        if not hedef:
            out['hatalar'].append(f'satır {no}: bölüm tanınmadı ({t} / {b})')
            continue
        try:
            nos, mesai, kisi = sayi(r, 'nos'), sayi(r, 'mesai'), sayi(r, 'kisi')
        except ValueError as e:
            out['hatalar'].append(f'satır {no}: {e}')
            continue
        if nos is None:
            continue
        if nos < 0 or (mesai or 0) < 0:
            out['hatalar'].append(f'satır {no}: negatif saat')
            continue
        dd = out['saatler'].setdefault(hedef, {'nos': 0.0, 'mesai': 0.0, 'kisi': 0})
        dd['nos'] += nos
        dd['mesai'] += mesai or 0.0
        dd['kisi'] += int(kisi or 0)
    return out


def _ncs_anahtar(excel_bolum):
    """aylik_calisma.excel_bolum → (tesis, bölüm) YALNIZ Ana Veri NçS anahtarıysa ('TK2/kaynak');
    kapasite Excel'inin '#N/A' gibi adları None döner."""
    p = str(excel_bolum or '').split('/', 1)
    return tuple(p) if len(p) == 2 and tuple(p) in NCS_BOLUM_AD else None


def ncs_farki(conn, n):
    """Sayfadaki ay ↔ veritabanındaki o ayın saatleri (önizleme için)."""
    if not n.get('yil'):
        return {'yil': None, 'ay': None, 'satir': 0, 'satirlar': [], 'degisen': 0, 'hatalar': n.get('hatalar') or []}
    eski, kaynak = {}, set()
    for r in conn.execute("SELECT excel_bolum, nos, mesai, dosya FROM aylik_calisma WHERE yil=? AND ay=?",
                          (n['yil'], n['ay'])):
        kaynak.add(r[3] or '')
        h = CALISMA_BOLUM.get(r[0]) or _ncs_anahtar(r[0])
        if h:
            e = eski.setdefault(h, [0.0, 0.0])
            e[0] += r[1] or 0
            e[1] += r[2] or 0
    sira = list(NCS_BOLUM_AD)
    satirlar, degisen = [], 0
    for h, v in sorted(n['saatler'].items(), key=lambda x: sira.index(x[0])):
        e = eski.get(h)
        fark = e is None or abs(e[0] - v['nos']) > 0.05 or abs(e[1] - v['mesai']) > 0.05
        degisen += 1 if fark else 0
        satirlar.append({'tesis': h[0], 'bolum': NCS_BOLUM_AD[h], 'nos': round(v['nos'], 1),
                         'mesai': round(v['mesai'], 1), 'kisi': v['kisi'],
                         'eski_nos': round(e[0], 1) if e else None, 'degisti': fark})
    return {'yil': n['yil'], 'ay': n['ay'], 'ay_ad': f"{AY_AD[n['ay']]} {n['yil']}", 'satir': len(satirlar),
            'toplam_nos': round(sum(v['nos'] for v in n['saatler'].values()), 1), 'satirlar': satirlar,
            'degisen': degisen + (1 if len(eski) != len(n['saatler']) else 0),
            'mevcut_kaynak': sorted(k for k in kaynak if k), 'hatalar': n.get('hatalar') or []}


def ncs_yaz(conn, yil, ay, saatler, kullanici, commit=True, dosya=NCS_KAYNAK):
    """O ayın aylik_calisma satırlarını sayfadakilerle DEĞİŞTİRİR (anahtar 'TK2/kaynak').
    commit=False: çağıranın işlemi içinde (tablolar önceden kurulmuş olmalı).
    dosya: panelde '⏱ Çalışma saati' yanında görünen kaynak adı (Ana Veri NçS ya da IK Excel'i)."""
    if commit:
        tablolari_kur(conn)
    conn.execute("DELETE FROM aylik_calisma WHERE yil=? AND ay=?", (yil, ay))
    simdi = datetime.now().strftime('%Y-%m-%d %H:%M')
    conn.executemany("INSERT INTO aylik_calisma (yil, ay, excel_bolum, nos, mesai, kisi, isimsiz, dosya, kullanici, "
                     "guncellendi) VALUES (?,?,?,?,?,?,0,?,?,?)",
                     [(yil, ay, f'{t}/{b}', round(v['nos'], 1), round(v['mesai'], 1), int(v['kisi'] or 0),
                       dosya, kullanici or '', simdi) for (t, b), v in saatler.items()])
    if commit:
        conn.commit()
    return len(saatler)


# ── IK ÇALIŞMA SAATİ EXCEL'İ (kullanıcı 2026-11-03) ───────────────────────────────
# "Aylık toplam çalışma saatini IK'dan bu Excel ile alıyoruz; dashboard'a yükleyeyim ve hangi
#  sütunları toplayacağını yazabileyim — total NçS için I, mesaileri de katmak için R,S,T."
# Dosya: her ay ayrı sayfa ('2026-09'); kişi başı satır: Sicil No · İsim · Ay · … · NçS (I) ·
# … · %90 (R) · %95 (S) · %100 (T) …. BÖLÜM SÜTUNU YOK → kişi → (tesis, bölüm) eşlemesi SİCİL
# ile kalıcı saklanır (ik_personel). İlk yüklemede Forge operatör adıyla tek eşleşen kişi
# önerilir; kalanlar panelde atanır, 'disi' = üretim dışı (saate girmez). Toplanacak sütunlar
# panelden yazılır (aylik_ayar). Yazım Aylık NçS ile aynı biçimde (ncs_yaz, 'TK2/kaynak').
IK_VARSAYILAN = {'nos': 'I', 'mesai': 'R,S,T'}
IK_DISI = 'disi'
IK_KLASOR = os.path.join(PROJECT_DIR, 'data', 'ik_saat')
_TR_ASCII = str.maketrans('çğıöşüÇĞİÖŞÜ', 'cgiosuCGIOSU')


def ik_tablolari_kur(conn):
    conn.execute("""CREATE TABLE IF NOT EXISTS ik_personel (
        sicil TEXT PRIMARY KEY, isim TEXT DEFAULT '', hedef TEXT DEFAULT '', kaynak TEXT DEFAULT '',
        kullanici TEXT DEFAULT '', guncellendi TEXT DEFAULT '')""")
    conn.execute("CREATE TABLE IF NOT EXISTS aylik_ayar (anahtar TEXT PRIMARY KEY, deger TEXT)")


def _harf_indeks(h):
    n = 0
    for ch in h:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _indeks_harf(i):
    s = ''
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def sutun_listesi(metin):
    """'I' · 'R,S,T' · 'R+S+T' · 'R:T' / 'R-T' → [indeks] (A=0). Boş → []. Hatalıysa ValueError."""
    out = []
    for parca in re.split(r'[,;+\s]+', str(metin or '').upper().strip()):
        if not parca:
            continue
        m = re.fullmatch(r'([A-Z]{1,3})(?:[:\-]([A-Z]{1,3}))?', parca)
        if not m:
            raise ValueError(f'Sütun harfi anlaşılmadı: {parca!r} (ör. I · R,S,T · R:T)')
        a = _harf_indeks(m.group(1))
        b = _harf_indeks(m.group(2)) if m.group(2) else a
        a, b = min(a, b), max(a, b)
        out += [i for i in range(a, b + 1) if i not in out]
    return out


def ik_ayar(conn):
    ik_tablolari_kur(conn)
    r = conn.execute("SELECT deger FROM aylik_ayar WHERE anahtar='ik_sutunlar'").fetchone()
    try:
        d = json.loads(r[0]) if r else {}
    except ValueError:
        d = {}
    return {k: (d.get(k) or v) for k, v in IK_VARSAYILAN.items()}


def ik_ayar_yaz(conn, nos, mesai):
    if not sutun_listesi(nos):
        raise ValueError('NçS için en az bir sütun yazın (ör. I)')
    sutun_listesi(mesai)
    ik_tablolari_kur(conn)
    d = {'nos': str(nos).strip().upper(), 'mesai': str(mesai or '').strip().upper()}
    conn.execute("INSERT OR REPLACE INTO aylik_ayar (anahtar, deger) VALUES ('ik_sutunlar', ?)", (json.dumps(d),))
    conn.commit()
    return d


def _ik_sayfa_ay(ad):
    """'2026-09' → (2026, 9) | None."""
    m = re.match(r'^\s*(\d{4})\D+(\d{1,2})\s*$', str(ad or ''))
    return (int(m.group(1)), int(m.group(2))) if m and 1 <= int(m.group(2)) <= 12 else None


def _sayi_hucre(v):
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip().replace(',', '.')) if v not in (None, '') else 0.0
    except ValueError:
        return 0.0


def ik_oku(yol, sayfa=None):
    """IK Excel'i → {'sayfalar': [{ad, yil, ay}], 'sayfa', 'satirlar', 'basliklar', 's_i', 'i_i'}.
    sayfa verilmezse yıl-ay adlı EN SON sayfa."""
    from openpyxl import load_workbook
    wb = load_workbook(yol, read_only=True, data_only=True)
    try:
        sayfalar = [{'ad': s, 'yil': (_ik_sayfa_ay(s) or (None, None))[0], 'ay': (_ik_sayfa_ay(s) or (None, None))[1]}
                    for s in wb.sheetnames]
        if not sayfa:
            aylik = [s for s in sayfalar if s['yil']]
            sayfa = max(aylik, key=lambda s: (s['yil'], s['ay']))['ad'] if aylik else wb.sheetnames[0]
        if sayfa not in wb.sheetnames:
            raise ValueError(f"Dosyada '{sayfa}' sayfası yok")
        satirlar = [list(r) for r in wb[sayfa].iter_rows(values_only=True)]
    finally:
        wb.close()
    bas = next((i for i, r in enumerate(satirlar[:20]) if any('sicil' in _kucuk(v) for v in r)), None)
    if bas is None:
        raise ValueError(f"'{sayfa}' sayfasında başlık satırı bulunamadı — 'Sicil No' sütunu olmalı")
    h = [str(v if v is not None else '').strip() for v in satirlar[bas]]
    s_i = next(i for i, v in enumerate(h) if 'sicil' in _kucuk(v))
    i_i = next((i for i, v in enumerate(h) if _kucuk(v) in ('isim', 'ad soyad', 'adı soyadı', 'adi soyadi', 'ad')), s_i + 1)
    return {'sayfalar': sayfalar, 'sayfa': sayfa, 'satirlar': satirlar[bas + 1:], 'basliklar': h,
            's_i': s_i, 'i_i': i_i}


def ik_kisiler(okuma, nos_kol, mesai_kol):
    """Satırlar → [{sicil, isim, nos, mesai}] — seçili sütunların toplamı."""
    out = []
    for r in okuma['satirlar']:
        s = r[okuma['s_i']] if okuma['s_i'] < len(r) else None
        if s in (None, ''):
            continue
        sicil = str(int(s)) if isinstance(s, float) and s.is_integer() else str(s).strip()
        if not re.search(r'\d', sicil):
            continue                          # alt toplam / açıklama satırı
        topla = lambda kol: sum(_sayi_hucre(r[j]) for j in kol if j < len(r))   # noqa: E731
        out.append({'sicil': sicil, 'isim': str(r[okuma['i_i']] or '').strip() if okuma['i_i'] < len(r) else '',
                    'nos': round(topla(nos_kol), 2), 'mesai': round(topla(mesai_kol), 2)})
    return out


def _ad_anahtar(ad):
    return re.sub(r'[^A-Z]', '', str(ad or '').translate(_TR_ASCII).upper())


def ik_eslestir(conn, kisiler):
    """Her kişiye 'hedef' ('TK2/kaynak' | 'disi' | '') + 'kaynak' (kayıtlı / operatör / '') ekler."""
    ik_tablolari_kur(conn)
    kayitli = {r[0]: (r[1], r[2]) for r in conn.execute("SELECT sicil, hedef, kaynak FROM ik_personel")}
    ops = {}
    for ad, bolum, lok in conn.execute("SELECT ad, COALESCE(bolum,'kaynak'), COALESCE(lokasyon,'TK2') FROM operatorler"):
        if (lok, bolum) in NCS_BOLUM_AD:
            ops.setdefault(_ad_anahtar(ad), set()).add(f'{lok}/{bolum}')
    for k in kisiler:
        h = kayitli.get(k['sicil'])
        if h and h[0]:
            k['hedef'], k['kaynak'] = h[0], 'kayıtlı'
        else:
            aday = ops.get(_ad_anahtar(k['isim'])) or set()
            k['hedef'], k['kaynak'] = (next(iter(aday)), 'operatör') if len(aday) == 1 else ('', '')
    return kisiler


def ik_toplam(kisiler):
    """→ (saatler {(tesis, bölüm): {nos, mesai, kisi}}, atanmamış {kisi, saat}, üretim dışı {kisi, saat})."""
    saatler, atanmamis, disi = {}, {'kisi': 0, 'saat': 0.0}, {'kisi': 0, 'saat': 0.0}
    for k in kisiler:
        top = k['nos'] + k['mesai']
        if top <= 0:
            continue
        if not k.get('hedef') or k['hedef'] == IK_DISI:
            d = disi if k.get('hedef') == IK_DISI else atanmamis
            d['kisi'] += 1
            d['saat'] = round(d['saat'] + top, 1)
            continue
        t, b = k['hedef'].split('/', 1)
        d = saatler.setdefault((t, b), {'nos': 0.0, 'mesai': 0.0, 'kisi': 0})
        d['nos'] += k['nos']
        d['mesai'] += k['mesai']
        d['kisi'] += 1
    return saatler, atanmamis, disi


def ik_hedefler():
    return [{'deger': f'{t}/{b}', 'ad': f'{t} · {ad}'} for (t, b), ad in NCS_BOLUM_AD.items()] + \
           [{'deger': IK_DISI, 'ad': 'Üretim dışı (saate girmez)'}]


def ik_onizle(conn, yol, sayfa, nos, mesai, atamalar=None):
    """Önizleme: bölüm toplamları + kişi listesi (atamalar: panelde değiştirilen {sicil: hedef})."""
    nos_kol, mesai_kol = sutun_listesi(nos), sutun_listesi(mesai)
    if not nos_kol:
        raise ValueError('NçS için en az bir sütun yazın (ör. I)')
    ok = ik_oku(yol, sayfa)
    kisiler = ik_eslestir(conn, ik_kisiler(ok, nos_kol, mesai_kol))
    gecerli = {h['deger'] for h in ik_hedefler()}
    for k in kisiler:
        if atamalar and k['sicil'] in atamalar:
            v = atamalar[k['sicil']]
            k['hedef'], k['kaynak'] = (v, 'elle') if v in gecerli else ('', '')
    saatler, atanmamis, disi = ik_toplam(kisiler)
    ya = _ik_sayfa_ay(ok['sayfa'])
    bas = ok['basliklar']
    etiket = lambda kol: [{'harf': _indeks_harf(j), 'baslik': bas[j] if j < len(bas) else ''} for j in kol]  # noqa: E731
    sira = list(NCS_BOLUM_AD)
    mevcut = []
    if ya:
        tablolari_kur(conn)
        mevcut = sorted({r[0] for r in conn.execute("SELECT DISTINCT dosya FROM aylik_calisma WHERE yil=? AND ay=?", ya) if r[0]})
    return {'sayfalar': ok['sayfalar'], 'sayfa': ok['sayfa'], 'yil': ya[0] if ya else None, 'ay': ya[1] if ya else None,
            'ay_ad': f'{AY_AD[ya[1]]} {ya[0]}' if ya else '', 'nos_sutun': etiket(nos_kol), 'mesai_sutun': etiket(mesai_kol),
            'bolumler': [{'hedef': f'{t}/{b}', 'tesis': t, 'bolum': NCS_BOLUM_AD[(t, b)], 'nos': round(v['nos'], 1),
                          'mesai': round(v['mesai'], 1), 'kisi': v['kisi']}
                         for (t, b), v in sorted(saatler.items(), key=lambda x: sira.index(x[0]))],
            'toplam': {'nos': round(sum(v['nos'] for v in saatler.values()), 1),
                       'mesai': round(sum(v['mesai'] for v in saatler.values()), 1),
                       'kisi': sum(v['kisi'] for v in saatler.values())},
            'atanmamis': atanmamis, 'disi': disi, 'kisiler': kisiler, 'hedefler': ik_hedefler(), 'mevcut': mevcut}


def ik_kaydet(conn, yol, dosya_adi, sayfa, nos, mesai, atamalar, kullanici, yil=None, ay=None):
    """Eşlemeyi (ik_personel) kalıcı yazar + ayın çalışma saatlerini DEĞİŞTİRİR. Ay sayfa adından
    ('2026-09'); ad aya benzemiyorsa çağıranın yil/ay'ı."""
    o = ik_onizle(conn, yol, sayfa, nos, mesai, atamalar)
    yil, ay = (o['yil'], o['ay']) if o['yil'] else (yil, ay)
    if not (yil and ay):
        raise ValueError(f"'{sayfa}' sayfa adından ay anlaşılmadı (ör. 2026-09)")
    simdi = datetime.now().strftime('%Y-%m-%d %H:%M')
    conn.executemany(
        "INSERT INTO ik_personel (sicil, isim, hedef, kaynak, kullanici, guncellendi) VALUES (?,?,?,?,?,?) "
        "ON CONFLICT(sicil) DO UPDATE SET isim=excluded.isim, hedef=excluded.hedef, kaynak=excluded.kaynak, "
        "kullanici=excluded.kullanici, guncellendi=excluded.guncellendi",
        [(k['sicil'], k['isim'], k['hedef'], 'elle' if k['kaynak'] == 'elle' else (k['kaynak'] or ''), kullanici, simdi)
         for k in o['kisiler'] if k['hedef'] and k['kaynak'] != 'kayıtlı'])
    saatler, atanmamis, disi = ik_toplam(o['kisiler'])
    ncs_yaz(conn, yil, ay, saatler, kullanici, dosya=f'IK › {dosya_adi} · {sayfa}')
    return {'yil': yil, 'ay': ay, 'ay_ad': f'{AY_AD[ay]} {yil}', 'bolum': len(saatler), 'toplam': o['toplam'],
            'atanmamis': atanmamis, 'disi': disi}


def calisma_haritasi(conn, yil, ay):
    """{(tesis, bölüm): {'nos','mesai','kisi'}} + atanmamış (üretim dışı) satırlar."""
    tablolari_kur(conn)
    atanan, disarida, bilgi = {}, [], None
    for r in conn.execute("SELECT excel_bolum, nos, mesai, kisi, dosya, guncellendi FROM aylik_calisma "
                          "WHERE yil=? AND ay=? ORDER BY excel_bolum", (yil, ay)):
        bilgi = {'dosya': r[4], 'guncellendi': r[5]}
        nk = _ncs_anahtar(r[0])                      # Ana Veri › Aylık NçS satırı (bkz. ncs_yaz)
        hedef = CALISMA_BOLUM.get(r[0]) or nk
        if hedef:
            d = atanan.setdefault(hedef, {'nos': 0.0, 'mesai': 0.0, 'kisi': 0, 'excel': []})
            d['nos'] += r[1] or 0
            d['mesai'] += r[2] or 0
            d['kisi'] += r[3] or 0
            d['excel'].append(NCS_BOLUM_AD[nk] if nk else r[0])
        else:
            disarida.append({'excel_bolum': r[0], 'nos': r[1] or 0, 'mesai': r[2] or 0, 'kisi': r[3] or 0})
    return atanan, disarida, bilgi


# ── SINIFLAMA + SÜRE ─────────────────────────────────────────────────────────
def _forge(conn):
    """{tesis: (tam, kök)} — {kod: {bölüm: toplam çevrim sn}}."""
    out = {}
    for lok in ('TK1', 'TK2'):
        tam, kok = {}, {}
        for kod, bolum, ct in conn.execute("SELECT referans_kodu, COALESCE(bolum,'kaynak'), COALESCE(hedef_cycle_time_sn,0) "
                                           "FROM referans_listesi WHERE COALESCE(lokasyon,'TK2')=?", (lok,)):
            for h, k in ((tam, _norm(kod)), (kok, _kok(kod))):
                h.setdefault(k, {}).setdefault(bolum, 0.0)
                h[k][bolum] += float(ct or 0)
        out[lok] = (tam, kok)
    return out


# ÇOK ADIMLI KAYNAK (kullanıcı 2026-11-03: "130.5422GRW'yi üretmek için 1. op kaynağını da yapıyoruz,
# bu gibi ön işlem kaynakları da göz önünde bulundurmalıyız"). Bölüm Kapasitesi'nin kuralı
# (bolum_kapasite._adimlari_topla, kullanıcı 2026-10-26): ara adım satırları ('KOD 1. op',
# 'KOD 1.ist', 'KOD(hazırlık)') ERP teyidi ALMAZ; süreleri teyidi alan ANA kodun süresine EKLENİR
# (ana kodun kendi süresi yoksa süre adımların toplamıdır). Önceden ana kodun kendi süresi doluysa
# adımlar atlanıyordu (2757W-10: 116 sn sayılıyor, hazırlık 233 sn yoktu; 4415W: 120, op'lar 178 yok)
# ve paranteze yazılmış adım ('…(hazırlık)') hiç bulunmuyordu.
ADIM_BOLUMLERI = (('TK2', 'kaynak'),)


def _adimlar(conn):
    """{(tesis, bölüm): {ANA_KOD: [(adım adı, sn)]}} — referans_listesi'ndeki ara adım satırları."""
    import bolum_kapasite as BK
    out = {}
    for kod, bolum, lok, ct in conn.execute(
            "SELECT referans_kodu, COALESCE(bolum,'kaynak'), COALESCE(lokasyon,'TK2'), COALESCE(hedef_cycle_time_sn,0) "
            "FROM referans_listesi"):
        if (lok, bolum) not in ADIM_BOLUMLERI or not ct:
            continue
        m = BK.ADIM_DESENI.match(str(kod or '').strip())
        if not m:
            continue
        ana, ad = m.group(1).upper(), m.group(2).strip()
        if ana == _norm(kod) or any(x in ad.upper() for x in BK.HARIC_ADIM):
            continue
        out.setdefault((lok, bolum), {}).setdefault(ana, []).append((ad, float(ct)))
    return out


# ── FORGE KARŞILAŞTIRMASI (kullanıcı 2026-11-03) ──────────────────────────────────
# "RPR / CFI'ı AS400'den, NçS'yi IK'dan alıyoruz. Karşılaştırma yapılabilmesi ve gerçek veriyi
#  bilmek için aynı verileri Forge'a göre de hesaplayıp raporda görelim."
#  · Adet: operatörlerin üretim kayıtları, BİTMİŞ ÜRÜN (haftalik_uretim kuralı — ara op, tamir /
#    punta / rework, TK1 test cihazı sayılmaz; son op ana koda yazılır). TEL: her adım ayrı kayıt
#    ('93.TK.464 KAPAMA') → kablo başına adım toplamlarının EN BÜYÜĞÜ (adım sayısı kadar katlanmasın).
#  · Teorik süre: ERP tarafıyla AYNI birim süre (aynı ürüne aynı çarpan; fark yalnız adetten gelir).
#  · Çalışma saati: vardiya giriş–çıkış; kişi-gün başına aralıkların BİRLEŞİMİ (aynı kişi iki robota
#    paralel vardiya açabiliyor — toplanırsa saat ikiye katlanır), çay / yemek molası ve namaz
#    düşülür (setup, fikstür, robot programı gibi planlı işler çalışma sayılır). Kişi = farklı operatör.
FAALIYET = ('depo', 'kalite')


def _ay_araligi(yil, ay):
    bas = date(yil, ay, 1)
    return bas, date(yil + (ay == 12), ay % 12 + 1, 1) - timedelta(days=1)


# İŞLEM EKİ: operatör kodun yanına işlemi yazıyor — '10.130.5158 delme', '10.300.6342A(somun)',
# '10.300.1845. Op.1', '10.300.2315.op.1', '93.TK.464 KAPAMA'. Aynı parça her işlemde AYRI kayıt;
# toplanırsa işlem sayısı kadar katlanır (Eylül işleme: ERP 4.965 ↔ ham Forge 52.173).
_OP_SONEK = re.compile(r'(?i)[\s.\-_]*(op[r]?t?\.?\s*\d+|\d+\s*\.?\s*(op[r]?t?|ist)\.?)\s*$')


def _taban(kod):
    """İşlem ekini atar: '10.300.1845. Op.1' → '10.300.1845', '10.300.6342A(somun)' → '10.300.6342A'."""
    k = re.split(r'[\s(]', str(kod or '').strip(), maxsplit=1)[0]
    return _OP_SONEK.sub('', k).rstrip('.-_ ') or str(kod or '').strip()


def forge_uretim(conn, yil, ay):
    """{(tesis, bölüm): {ANAHTAR: {'kod', 'adet', 'tur'}}} — Forge'un bitmiş ürün adedi. ANAHTAR nokta /
    boşluk farkını yutar (haftalik_uretim.ref_anahtar: '10 300 0609' = '10.300.0609' = '10.3000609').
    Kodun YALIN kaydı varsa adet odur (işlem ekli kayıtlar ara işlemdir, süresi birim süreye girer);
    yalnız işlem ekli kayıt varsa her işlem aynı parçaları işlediği için EN BÜYÜĞÜ; tel'de kablonun
    girdiği SON adım (kesim parça sayar — en büyük adım kesim olurdu)."""
    import haftalik_uretim as HU
    import tel_proses as TP
    bas, bit = _ay_araligi(yil, ay)
    gruplar = {}
    for lok in ('TK2', 'TK1'):
        for r in HU._kayitlar(conn, lok, bas, bit, FAALIYET)[0]:
            ref, ok = r['referans'], r['ok'] or 0
            taban = TP.tel_ek_ayikla(ref) if r['bolum'] == 'tel' else _taban(ref)
            a = HU.ref_anahtar(taban)
            g = gruplar.setdefault((lok, r['bolum'], a), {'kod': taban, 'yalin': None, 'islem': {}})
            adim = TP.tel_koddan_adim(ref) if r['bolum'] == 'tel' else None
            # Yalın kayıt: ekin hiçbiri yok ('10 300 0609' boşluklu yazımı nokta sayılır; parantez /
            # boşluk sonrası ad / op eki = işlem kaydı — ref_anahtar parantezi yuttuğu için ayrıca bakılır)
            temiz = HU._le.bosluk_nokta(ref) if HU._le else ref
            if adim or re.search(r'[\s(]', temiz) or _OP_SONEK.search(temiz) or HU.ref_anahtar(ref) != a:
                i = adim or ref.strip().upper()
                g['islem'][i] = g['islem'].get(i, 0.0) + ok
                continue
            g['yalin'] = (g['yalin'] or 0.0) + ok
            if temiz.count('.') > g['kod'].count('.'):    # noktalı yazım görünsün
                g['kod'] = temiz
    out = {}
    for (lok, b, a), g in gruplar.items():
        if g['yalin'] is not None:
            adet, tur = g['yalin'], ''
        elif b == 'tel':
            son = max(g['islem'], key=lambda x: TP.TEL_ADIM_SIRA.get(x, 0))
            adet, tur = g['islem'][son], f'son adım: {son}'
        else:
            adet, tur = max(g['islem'].values()), 'işlem kayıtlarının en büyüğü'
        out.setdefault((lok, b), {})[a] = {'kod': g['kod'], 'adet': adet, 'tur': tur}
    return out


def _dakika(s):
    m = re.match(r'^\s*(\d{1,2}):(\d{2})', str(s or ''))
    return int(m.group(1)) * 60 + int(m.group(2)) if m else None


def forge_calisma(conn, yil, ay):
    """{(tesis, bölüm): {'saat', 'brut', 'mola', 'kisi', 'kisi_gun', 'durus', 'durus_planli', 'durus_plansiz',
    'durus_sebep'}} — vardiyalardan. 'saat' = vardiya − mola; DURUŞ (kullanıcı 2026-11-03: "performansı
    doğru hesaplamak için duruşları da düşmek gerekir") = mola dışındaki planlı + plansız duruşlar
    (setup, fikstür, arıza, malzeme bekleme…); kişi-gün başına tekil (paralel vardiyada aynı sebep +
    başlangıç bir kez) ve o günün çalışma süresini aşamaz."""
    bas, bit = _ay_araligi(yil, ay)
    bugun, simdi = date.today().isoformat(), datetime.now()
    araliklar, vk = {}, {}
    for vid, t, lok, b, op, bs, bt, dk in conn.execute(
            "SELECT id, tarih, COALESCE(lokasyon,'TK2'), COALESCE(bolum,'kaynak'), COALESCE(operator_adi,''), "
            "baslangic_saati, bitis_saati, COALESCE(toplam_sure_dk,0) FROM vardiyalar WHERE tarih BETWEEN ? AND ?",
            (bas.isoformat(), bit.isoformat())):
        if b in FAALIYET:
            continue
        s, e = _dakika(bs), _dakika(bt)
        if s is None:
            if not dk:
                continue
            s, e = 0, dk                       # saat yok: süre tek başına (birleşime giremez)
        elif e is None:
            e = (simdi.hour * 60 + simdi.minute) if t == bugun else s + (dk or 0)   # açık vardiya
        elif e < s:
            e += 1440                          # gece yarısını geçen vardiya
        k = (lok, b, op.strip().upper(), t)
        araliklar.setdefault(k, []).append((s, e))
        vk[vid] = k
    molalar, duruslar = {}, {}
    for vid, sebep, tip, ds, sdk in conn.execute(
            "SELECT d.vardiya_id, d.durus_sebebi, COALESCE(d.durus_tipi,''), d.baslangic_saati, COALESCE(d.sure_dk,0) "
            "FROM duruslar d JOIN vardiyalar v ON v.id = d.vardiya_id WHERE v.tarih BETWEEN ? AND ?",
            (bas.isoformat(), bit.isoformat())):
        s = _kucuk(sebep)
        k = vk.get(vid)
        if not k or not sdk:
            continue
        # Paralel vardiyaların otomatik molaları / duruşları aynı saatte iki kez yazılır → (sebep, başlangıç) tekil
        anahtar = (s, ds or f'v{vid}')
        if tip == 'planli' and ('mola' in s or 'namaz' in s):
            m = molalar.setdefault(k, {})
            m[anahtar] = max(m.get(anahtar, 0.0), float(sdk))
        else:
            m = duruslar.setdefault(k, {})
            onceki = m.get(anahtar)
            if not onceki or float(sdk) > onceki[0]:
                m[anahtar] = (float(sdk), str(sebep or '?').strip() or '?', 'plansiz' if tip == 'plansiz' else 'planli')
    out = {}
    for k, ar in araliklar.items():
        ar.sort()
        sure, son_s, son_e = 0.0, None, None
        for s, e in ar:
            if son_e is None or s > son_e:
                if son_e is not None:
                    sure += son_e - son_s
                son_s, son_e = s, e
            else:
                son_e = max(son_e, e)
        sure += (son_e - son_s) if son_e is not None else 0
        mola = min(sum(molalar.get(k, {}).values()), sure)
        dl = list(duruslar.get(k, {}).values())
        dtop = sum(x[0] for x in dl)
        oran = min(1.0, (sure - mola) / dtop) if dtop else 0.0      # duruş o günün çalışmasını aşamaz
        o = out.setdefault((k[0], k[1]), {'brut': 0.0, 'mola': 0.0, 'kisi': set(), 'kisi_gun': 0,
                                          'durus': {'planli': 0.0, 'plansiz': 0.0}, 'sebep': {}})
        o['brut'] += sure
        o['mola'] += mola
        o['kisi'].add(k[2])
        o['kisi_gun'] += 1
        for dk_, sebep, tip in dl:
            o['durus'][tip] += dk_ * oran
            o['sebep'][(sebep, tip)] = o['sebep'].get((sebep, tip), 0.0) + dk_ * oran
    return {k: {'saat': round((v['brut'] - v['mola']) / 60, 1), 'brut': round(v['brut'] / 60, 1),
                'mola': round(v['mola'] / 60, 1), 'kisi': len(v['kisi']), 'kisi_gun': v['kisi_gun'],
                'durus': round((v['durus']['planli'] + v['durus']['plansiz']) / 60, 1),
                'durus_planli': round(v['durus']['planli'] / 60, 1), 'durus_plansiz': round(v['durus']['plansiz'] / 60, 1),
                'durus_sebep': [{'sebep': s, 'tip': t, 'saat': round(dk_ / 60, 1)}
                                for (s, t), dk_ in sorted(v['sebep'].items(), key=lambda x: -x[1])][:12]}
            for k, v in out.items()}


def _sure_kaynaklari(conn):
    kap, rota, db = {}, {}, {}
    try:
        for kod, saatlik, hat in conn.execute("SELECT kod, saatlik, hat FROM aylik_kapasite_db"):
            db[kod] = {'saatlik': saatlik, 'hat': hat or ''}
    except Exception:
        pass
    try:
        for kod, lok, b, sn in conn.execute("SELECT kod, COALESCE(lokasyon,'TK2'), bolum, sure_sn FROM kapasite_sure "
                                            "WHERE sure_sn > 0"):
            kap[(_norm(kod), lok, b)] = float(sn)
    except Exception:
        pass
    try:
        for kod, b, sn in conn.execute("SELECT kod, bolum, sure_sn FROM kapasite_urun_bolum WHERE sure_sn > 0"):
            rota[(_norm(kod), b)] = float(sn)
    except Exception:
        pass
    return kap, rota, db


def rapor(conn, yil, ay):
    """Ayın raporu (son AS400 okumasından). None → o ay hiç okunmamış."""
    import anaveri as AV
    veri = son_cekim(conn, yil, ay)
    if not veri:
        return None
    forge = _forge(conn)
    adimlar = _adimlar(conn)
    kap, rota, kdb = _sure_kaynaklari(conn)
    ana = AV.haritasi(conn)
    # Teller (93.*): önce Ana Veri'nin tel satırları (yüklendiyse), yoksa planlamanın Anaveri'si
    tel = AV.tel_haritasi(conn) if AV.tel_kaynagi_ana_veri(conn) else {}
    kurallar = haric_kurallari(conn)
    acik = veri.get('aciklama') or {}
    cop_okundu = 'COP' in (veri.get('causallar') or [])

    def forge_sn_bul(n, kod, tesis, b):
        """Ana Veri / Forge birim süresi (çok adımlı kaynakta ara adımlar eklenmiş) + adım notu."""
        tam, kok = forge[tesis]
        t = (tam.get(n) or {}).get(b, 0.0)
        # Tel: ERP kodu köktür, Forge'da adım ekli satırlar (KESIM, KAPAMA…) — süre kökte TOPLANIR;
        # eksiz satır süresizse kökün toplamı kullanılır.
        fsn = t if t > 0 else (kok.get(_kok(kod)) or {}).get(b, 0.0)
        ad = adimlar.get((tesis, b), {}).get(n)
        if not ad:
            return fsn, None
        return t + sum(x for _, x in ad), ' + '.join(([f'ana kod {t:g}'] if t else []) + [f'{a} {x:g}' for a, x in ad])

    def sure_sec(n, kod, tesis, b, forge_sn):
        """Birim süre → (sn, kaynak). ÖNCE ANA VERİ (kullanıcı 2026-10-13: "önce Ana Veri, yoksa
        planlama süresi"): TK2'de Forge süresi = Ana Veri'nin çevrim/kaynak süresi → varsa o; yoksa
        planlamanın kapasite süresi, o da yoksa ERP rotası. TK1 Ana Veri'de DEĞİL: orada planlama
        süresi önce kalır (tel süreleri Forge'da adım bazlı, kök toplamı eksik olur)."""
        if b == FASON:
            return 0.0, ''
        k = kap.get((n, tesis, b)) or kap.get((_kok(kod), tesis, b))
        ks = (kdb.get(n) or {}).get('saatlik')
        if (tesis == 'TK2' or b in SURE_EXCEL_DISI) and forge_sn > 0:
            k, ks = None, None
        if k:
            return k, 'kapasite Excel'
        if ks:
            return 3600.0 / ks, 'kapasite Excel'
        if forge_sn > 0:
            return forge_sn, 'Ana Veri' if tesis == 'TK2' else 'Forge'
        if rota.get((n, b)):
            return rota[(n, b)], 'ERP rota'
        return 0.0, ''

    # FORGE KARŞILAŞTIRMASI (2026-11-03): aynı ay, Forge'un kendi kayıtlarından (bkz. forge_uretim)
    import haftalik_uretim as HU
    f_uretim = forge_uretim(conn, yil, ay)
    erp_gorulen = set()
    # Kod × tesis birleştir (causal ayrı tutulur)
    kodlar = {}
    depo_ozet = {}
    for h in veri['hareket']:
        kod = h['kod'].strip()
        n = _norm(kod)
        tesis = DEPO_TESIS.get(h['depo'].upper())
        if not tesis:
            iki = [lok for lok in ('TK2', 'TK1') if forge[lok][0].get(n) or forge[lok][1].get(_kok(kod))]
            tesis = iki[0] if len(iki) == 1 else ''
        d = kodlar.setdefault((n, tesis), {'kod': kod, 'tesis': tesis, 'RPR': 0.0, 'CFI': 0.0, 'COP': 0.0,
                                           'depolar': {}})
        d[h['causal']] = d.get(h['causal'], 0.0) + h['adet']
        if h['causal'] != 'COP':        # depo sütunu üretimin deposunu gösterir
            d['depolar'][h['depo']] = d['depolar'].get(h['depo'], 0.0) + h['adet']
        o = depo_ozet.setdefault(h['depo'] or '?', {'depo': h['depo'] or '?', 'tesis': DEPO_TESIS.get(h['depo'].upper(), ''),
                                                    'RPR': 0.0, 'CFI': 0.0, 'COP': 0.0, 'kod': set()})
        o[h['causal']] = o.get(h['causal'], 0.0) + h['adet']
        o['kod'].add(n)
    # COP ÜRETİMİN YANINA: hurda depodan düşülür ve TK1 plastik hurdası 08.10.2026'ya kadar 01D'den
    # (TK2) düşülüyordu. Üretimi YALNIZ bir başka tesiste olan kodun yalnız-COP kaydı oraya eklenir.
    for (n, tesis), d in list(kodlar.items()):
        if d['RPR'] + d['CFI'] > 0 or d['COP'] <= 0:
            continue
        uretim = [k for k, x in kodlar.items() if k[0] == n and k[1] != tesis and x['RPR'] + x['CFI'] > 0]
        if len(uretim) == 1:
            kodlar[uretim[0]]['COP'] += d['COP']
            del kodlar[(n, tesis)]
    bolumler, tanimsiz, hammadde, haric, tesis_top, rapor_disi = {}, [], [], [], {}, {}
    for (n, tesis), d in kodlar.items():
        adet = d['RPR'] + d['CFI']
        cop = d['COP']
        sayilir = 1 if adet > 0 else 0      # ay içinde üretimi olmayan (yalnız COP) kod sayılmaz
        fa = HU.ref_anahtar(d['kod'])       # Forge eşleşmesi: nokta / boşluk farkı yutulur
        satir = {'kod': d['kod'], 'aciklama': _temiz(acik.get(d['kod'], '')), 'tesis': tesis, 'rpr': round(d['RPR'], 2),
                 'cfi': round(d['CFI'], 2), 'adet': round(adet, 2), 'cop': round(cop, 2),
                 'depolar': ' '.join(f'{k or "?"}:{v:g}' for k, v in d['depolar'].items())}
        hk = _haric_bul(n, kurallar)
        if hk:                      # kullanıcının hariç tuttuğu kod: hiçbir toplama / performansa girmez
            haric.append(dict(satir, kural=hk['desen'] + ('…' if hk['tur'] == 'onek' else ''), kural_id=hk['id']))
            continue
        if n.startswith(HAMMADDE_ONEK):
            hammadde.append(satir)
            continue
        # Bölüm: Pull/Pandora fason → Forge (bu tesiste) → kapasite Database → Anaveri → tanımsız
        hedef, kaynak, adim_notu = {}, '', {}
        a_kod = ((tel.get(n) or (tel.get(n[:-1]) if n.endswith('W') else None)) if n.startswith('93.') else None) \
            or ana.get(n) or (ana.get(n[:-1]) if n.endswith('W') else None)
        if _fason_mi(a_kod, True) or _w_teli(n):
            tesis = tesis or 'TK1'
            satir['tesis'] = tesis
            hedef, kaynak = {FASON: 0.0}, 'anaveri'
        if tesis and not hedef:
            tam, kok = forge[tesis]
            t_h, k_h = tam.get(n), kok.get(_kok(d['kod'])) or {}
            # Tel: ERP kodu köktür, Forge'da adım ekli satırlar (KESIM, KAPAMA…) — süre kökte
            # TOPLANIR; eksiz satır süresizse kökün toplamı kullanılır.
            hedef = dict.fromkeys(t_h or k_h)
            kaynak = 'forge' if hedef else ''
            for b in hedef:          # çok adımlı kaynakta ana kod + ara adımlar (bkz. _adimlar)
                hedef[b], adim_notu[b] = forge_sn_bul(n, d['kod'], tesis, b)
        if not hedef:
            # Kapasite Excel'i Database hattı (planlamanın sınıflaması) — tesis uyuşuyorsa
            m = CALISMA_BOLUM.get((kdb.get(n) or {}).get('hat', ''))
            if m and (not tesis or m[0] == tesis):
                tesis = tesis or m[0]
                satir['tesis'] = tesis
                hedef, kaynak = {m[1]: 0.0}, 'kapasite Excel'
        if not hedef and tesis and _fason_mi(a_kod, False):
            hedef, kaynak = {FASON: 0.0}, 'anaveri'
        if not hedef:
            a = a_kod
            m = ANAVERI_BOLUM.get(((a or {}).get('tk', '').upper(), (a or {}).get('hat', '').upper())) if a else None
            if m and (not tesis or m[0] == tesis):
                tesis = tesis or m[0]
                satir['tesis'] = tesis
                hedef, kaynak = {m[1]: 0.0}, 'anaveri'
        if not hedef or not tesis:
            a = ana.get(n)
            satir['anaveri'] = f"{a['tk']} / {a['hat']}" if a else ''
            tanimsiz.append(satir)
            continue
        # İşleme + fason rapora dahil DEĞİL (kullanıcı 2026-11-03): kod başka bir bölümde de tanımlıysa
        # orada sayılmaya devam eder; yalnız bunlardaysa özete / tesis toplamına / performansa girmez.
        disari = [b for b in hedef if b in RAPOR_DISI]
        if disari:
            hedef = {b: v for b, v in hedef.items() if b not in RAPOR_DISI}
            if not hedef:               # yalnız işleme / fason kodu: tamamen dışarıda (notta sayılır)
                x = rapor_disi.setdefault(disari[0], {'bolum': BOLUM_AD.get(disari[0], disari[0]), 'kod': 0, 'adet': 0.0})
                x['kod'] += sayilir
                x['adet'] += adet
                continue
        t = tesis_top.setdefault(tesis, {'adet': 0.0, 'kod': 0, 'rpr': 0.0, 'cfi': 0.0, 'cop': 0.0})
        t['adet'] += adet
        t['kod'] += sayilir
        t['rpr'] += d['RPR']
        t['cfi'] += d['CFI']
        t['cop'] += cop
        for b, forge_sn in hedef.items():
            sn, sk = sure_sec(n, d['kod'], tesis, b, forge_sn or 0.0)
            erp_gorulen.add((tesis, b, fa))
            bl = bolumler.setdefault((tesis, b), {'tesis': tesis, 'bolum': b, 'ad': BOLUM_AD.get(b, b), 'adet': 0.0,
                                                  'kod': 0, 'rpr': 0.0, 'cfi': 0.0, 'cop': 0.0, 'teorik_sn': 0.0,
                                                  'sureli_adet': 0.0, 'satirlar': []})
            bl['adet'] += adet
            bl['kod'] += sayilir
            bl['rpr'] += d['RPR']
            bl['cfi'] += d['CFI']
            bl['cop'] += cop
            if sn > 0 and adet > 0:
                bl['teorik_sn'] += adet * sn
                bl['sureli_adet'] += adet
            bl['satirlar'].append(dict(satir, sure_sn=round(sn, 2) if sn else None, sure_kaynak=sk,
                                       sure_adim=adim_notu.get(b) if sk in ('Ana Veri', 'Forge') else None,
                                       teorik_saat=round(adet * sn / 3600, 2) if sn and adet > 0 else None,
                                       sinif=kaynak, cok_bolum=len(hedef) > 1,
                                       forge_adet=round(((f_uretim.get((tesis, b)) or {}).get(fa) or {}).get('adet') or 0, 2)))
    calisma, disarida, calisma_bilgi = calisma_haritasi(conn, yil, ay)
    # Forge tarafı: aynı hariç tutma + aynı birim süre; ERP'de bu bölümde olmayan kodlar ayrıca listelenir
    f_saat = forge_calisma(conn, yil, ay)
    f_ozet = {}
    for (tesis, b), kf in f_uretim.items():
        o = f_ozet.setdefault((tesis, b), {'adet': 0.0, 'kod': 0, 'teorik_sn': 0.0, 'sureli_adet': 0.0, 'yalniz': []})
        for fa, x in kf.items():
            n = _norm(x['kod'])
            if x['adet'] <= 0 or _haric_bul(n, kurallar) or n.startswith(HAMMADDE_ONEK):
                continue
            sn, sk = sure_sec(n, x['kod'], tesis, b, forge_sn_bul(n, x['kod'], tesis, b)[0])
            o['adet'] += x['adet']
            o['kod'] += 1
            if sn > 0:
                o['teorik_sn'] += x['adet'] * sn
                o['sureli_adet'] += x['adet']
            if (tesis, b, fa) not in erp_gorulen:
                o['yalniz'].append({'kod': x['kod'], 'adet': round(x['adet'], 2), 'sure_sn': round(sn, 2) if sn else None,
                                    'sure_kaynak': sk, 'teorik_saat': round(x['adet'] * sn / 3600, 2) if sn else None,
                                    'tur': x['tur']})
    f_tesis = {}
    ozet = []
    for tesis in ('TK2', 'TK1'):
        sira = BOLUM_SIRA.get(tesis, ())
        anahtarlar = sorted(({b for (t, b) in bolumler if t == tesis} | {b for (t, b) in calisma if t == tesis} |
                             {b for (t, b) in set(f_ozet) | set(f_saat) if t == tesis and b in sira}) - set(RAPOR_DISI),
                            key=lambda b: (sira.index(b) if b in sira else 99, b))
        for b in anahtarlar:
            bl = bolumler.get((tesis, b)) or {'tesis': tesis, 'bolum': b, 'ad': BOLUM_AD.get(b, b), 'adet': 0.0,
                                              'kod': 0, 'rpr': 0.0, 'cfi': 0.0, 'cop': 0.0, 'teorik_sn': 0.0,
                                              'sureli_adet': 0.0, 'satirlar': []}
            c = calisma.get((tesis, b)) or {}
            nos, mesai = c.get('nos', 0.0), c.get('mesai', 0.0)
            top = nos + mesai
            teorik = bl['teorik_sn'] / 3600
            fsn = b == FASON
            fo, fs = f_ozet.get((tesis, b)) or {}, f_saat.get((tesis, b)) or {}
            f_teorik = (fo.get('teorik_sn') or 0.0) / 3600
            forge_o = None if fsn else {
                'adet': round(fo.get('adet', 0.0), 2), 'kod': fo.get('kod', 0), 'teorik_saat': round(f_teorik, 1),
                'sure_kapsami': round(fo['sureli_adet'] / fo['adet'], 3) if fo.get('adet') else None,
                'calisma': fs.get('saat', 0.0), 'brut': fs.get('brut', 0.0), 'mola': fs.get('mola', 0.0),
                'kisi': fs.get('kisi', 0), 'kisi_gun': fs.get('kisi_gun', 0),
                'performans': round(f_teorik / fs['saat'], 3) if fs.get('saat') else None,
                'yalniz_sayi': len(fo.get('yalniz', [])), 'yalniz_adet': round(sum(x['adet'] for x in fo.get('yalniz', [])), 2)}
            # DURUŞLAR DÜŞÜLMÜŞ (kullanıcı 2026-11-03): Forge'un mola dışı duruşları hem IK saatinden hem
            # Forge vardiya saatinden düşülür; performans = teorik ÷ net çalışma
            durus_o = None
            if forge_o:
                d_s = fs.get('durus', 0.0)
                ik_net = round(top - d_s, 1) if top else 0.0
                f_net = round(fs.get('saat', 0.0) - d_s, 1)
                durus_o = {'saat': d_s, 'planli': fs.get('durus_planli', 0.0), 'plansiz': fs.get('durus_plansiz', 0.0),
                           'sebepler': fs.get('durus_sebep', []), 'ik_net': ik_net, 'forge_net': f_net,
                           'performans_erp': round(teorik / ik_net, 3) if ik_net > 0 else None,
                           'performans_forge': round(f_teorik / f_net, 3) if f_net > 0 else None}
                forge_o['durus'] = d_s
                ft = f_tesis.setdefault(tesis, {'adet': 0.0, 'teorik_saat': 0.0, 'calisma': 0.0, 'kisi': 0, 'durus': 0.0,
                                                'ik_net': 0.0, 'teorik_erp': 0.0})
                ft['adet'] += forge_o['adet']
                ft['teorik_saat'] += forge_o['teorik_saat']
                ft['calisma'] += forge_o['calisma']
                ft['kisi'] += forge_o['kisi']
                ft['durus'] += d_s
                if top:
                    ft['ik_net'] += ik_net
                    ft['teorik_erp'] += teorik
            # KOD FARKLARI (kullanıcı 2026-11-03: "aradaki farkların hangi kodlardan kaynaklandığını görmeliyim")
            farklar = [{'kod': s_['kod'], 'aciklama': s_['aciklama'], 'erp': s_['adet'], 'forge': s_.get('forge_adet') or 0.0,
                        'sure_sn': s_['sure_sn']} for s_ in bl['satirlar']] + \
                      [{'kod': x['kod'], 'aciklama': "yalnız Forge'da" + (f" ({x['tur']})" if x.get('tur') else ''), 'erp': 0.0,
                        'forge': x['adet'], 'sure_sn': x['sure_sn']} for x in fo.get('yalniz', [])]
            for x in farklar:
                x['fark'] = round(x['forge'] - x['erp'], 2)
                x['fark_saat'] = round(x['fark'] * x['sure_sn'] / 3600, 2) if x['sure_sn'] else None
            farklar = sorted((x for x in farklar if abs(x['fark']) >= 0.5), key=lambda x: (-abs(x['fark']), x['kod']))
            ozet.append({'fason': fsn, 'tesis': tesis, 'bolum': b, 'ad': bl['ad'], 'adet': round(bl['adet'], 2), 'kod': bl['kod'],
                         'rpr': round(bl['rpr'], 2), 'cfi': round(bl['cfi'], 2),
                         'cop': round(bl['cop'], 2) if cop_okundu else None,
                         'fire': round(bl['cop'] / bl['adet'], 4) if cop_okundu and bl['adet'] else None,
                         'teorik_saat': None if fsn else round(teorik, 1),
                         'sure_kapsami': round(bl['sureli_adet'] / bl['adet'], 3) if bl['adet'] and not fsn else None,
                         'nos': round(nos, 1), 'mesai': round(mesai, 1), 'calisma': round(top, 1),
                         'kisi': c.get('kisi', 0), 'calisma_excel': c.get('excel', []),
                         'performans': round(teorik / top, 3) if top else None,
                         'performans_nos': round(teorik / nos, 3) if nos else None,
                         'forge': forge_o,
                         'forge_yalniz': sorted(fo.get('yalniz', []), key=lambda x: -x['adet'])[:300],
                         'durus': durus_o,
                         'kod_farklari': farklar[:400], 'kod_fark_sayisi': len(farklar),
                         'satirlar': sorted(bl['satirlar'], key=lambda s: -s['adet'])})
    for ft in f_tesis.values():
        ft['performans'] = round(ft['teorik_saat'] / ft['calisma'], 3) if ft['calisma'] else None
        ft['forge_net'] = round(ft['calisma'] - ft['durus'], 1)
        ft['performans_forge_net'] = round(ft['teorik_saat'] / ft['forge_net'], 3) if ft['forge_net'] > 0 else None
        ft['performans_erp_net'] = round(ft['teorik_erp'] / ft['ik_net'], 3) if ft['ik_net'] > 0 else None
        for k in ('durus', 'ik_net', 'teorik_erp'):
            ft[k] = round(ft[k], 1)
        for k in ('adet', 'teorik_saat', 'calisma'):
            ft[k] = round(ft[k], 1)
    for t in tesis_top.values():
        t['fire'] = round(t['cop'] / t['adet'], 4) if cop_okundu and t['adet'] else None
        for k in ('adet', 'rpr', 'cfi', 'cop'):
            t[k] = round(t[k], 2)
        if not cop_okundu:
            t['cop'] = None
    tanimsiz.sort(key=lambda s: -s['adet'])
    hammadde.sort(key=lambda s: -s['adet'])
    haric.sort(key=lambda s: -s['adet'])
    hsay = {}
    for x in haric:
        v = hsay.setdefault(x['kural_id'], [0, 0.0])
        v[0] += 1
        v[1] += x['adet']
    haric_kurallar = [dict(k, kod=hsay.get(k['id'], [0, 0.0])[0], adet=round(hsay.get(k['id'], [0, 0.0])[1], 2))
                      for k in kurallar]
    return {'yil': yil, 'ay': ay, 'ay_ad': f'{AY_AD[ay]} {yil}', 'olcum': veri['ts'], 'olcan': veri.get('_kullanici', ''),
            'devam_eden': (yil, ay) == (date.today().year, date.today().month),
            'ozet': ozet, 'tesis': tesis_top, 'forge_tesis': f_tesis, 'tanimsiz': tanimsiz, 'hammadde': hammadde, 'cop_okundu': cop_okundu,
            'depolar': sorted(({**o, 'kod': len(o['kod']), 'RPR': round(o['RPR'], 2), 'CFI': round(o['CFI'], 2),
                                'COP': round(o['COP'], 2)}
                               for o in depo_ozet.values()), key=lambda o: -(o['RPR'] + o['CFI'])),
            'calisma_disarida': disarida, 'calisma_bilgi': calisma_bilgi,
            'tanimsiz_adet': round(sum(s['adet'] for s in tanimsiz), 2),
            'hammadde_adet': round(sum(s['adet'] for s in hammadde), 2),
            'haric': haric, 'haric_adet': round(sum(s['adet'] for s in haric), 2), 'haric_kurallar': haric_kurallar,
            'rapor_disi': [dict(x, adet=round(x['adet'], 2)) for x in rapor_disi.values()]}


# ── EXCEL (yöneticiyle paylaşılacak) ─────────────────────────────────────────
def excel(r):
    import io
    from openpyxl import Workbook
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    ws = wb.active
    ws.title = 'Özet'
    koyu = PatternFill('solid', fgColor='1F3864')
    acik = PatternFill('solid', fgColor='E9E4F7')
    ince = Side(style='thin', color='D0D0D0')
    kenar = Border(top=ince, bottom=ince, left=ince, right=ince)
    ws.append([f"AYLIK ÜRETİM RAPORU — {r['ay_ad']}" + (' (ay devam ediyor)' if r['devam_eden'] else '')])
    ws['A1'].font = Font(bold=True, size=14, color='1F3864')
    ws.append([f"Teyidi verilmiş üretim: AS400 RPR (launch teyidi) + CFI (launch'sız üretim girişi), hareket tarihine "
               f"göre · COP = hurda (üretime girmez), fire % = COP ÷ üretilen · okuma {r['olcum']} · bölümler Forge "
               f"tanımlarına göre"])
    ws['A2'].font = Font(italic=True, color='666666')
    ws.append([])
    bas = ['Tesis', 'Bölüm', 'Üretilen adet', 'Kod sayısı', 'RPR adet', 'CFI adet', 'COP (hurda) adet', 'Fire %',
           'Teorik süre (saat)', 'Süre kapsamı', 'Çalışma saati (NçS)', 'Mesai (saat)', 'Toplam çalışma (saat)',
           'Kişi', 'Performans', 'Hedef', 'Performans (yalnız NçS)']
    ws.append(bas)
    hb = ws.max_row
    for i in range(1, len(bas) + 1):
        c = ws.cell(row=hb, column=i)
        c.font, c.fill, c.border = Font(bold=True, color='FFFFFF'), koyu, kenar
        c.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
    ws.row_dimensions[hb].height = 32
    for tesis in ('TK2', 'TK1'):
        satirlar = [o for o in r['ozet'] if o['tesis'] == tesis]
        if not satirlar:
            continue
        ilk = ws.max_row + 1
        for o in satirlar:
            if o.get('fason'):
                continue            # fason satırı tesis toplamının altında ayrı yazılır
            ws.append([tesis, o['ad'], o['adet'], o['kod'], o['rpr'], o['cfi'], o.get('cop'), o.get('fire'),
                       o['teorik_saat'], o['sure_kapsami'],
                       o['nos'] or None, o['mesai'] or None, o['calisma'] or None, o['kisi'] or None,
                       o['performans'], HEDEF_PERFORMANS if o['calisma'] else None, o['performans_nos']])
        son = ws.max_row
        t = r['tesis'].get(tesis, {})
        L = get_column_letter
        tr_ = son + 1
        ws.append([f'{tesis} TOPLAM', '(kod tekil)', t.get('adet'), t.get('kod'), t.get('rpr'), t.get('cfi'),
                   t.get('cop'), (f'=IF(C{tr_}>0,G{tr_}/C{tr_},"")' if r.get('cop_okundu') else None),
                   f'=SUM(I{ilk}:I{son})', None, f'=SUM(K{ilk}:K{son})', f'=SUM(L{ilk}:L{son})',
                   f'=SUM(M{ilk}:M{son})', f'=SUM(N{ilk}:N{son})', f'=IF(M{tr_}>0,I{tr_}/M{tr_},"")',
                   HEDEF_PERFORMANS, f'=IF(K{tr_}>0,I{tr_}/K{tr_},"")'])
        for i in range(1, len(bas) + 1):
            c = ws.cell(row=ws.max_row, column=i)
            c.font, c.fill = Font(bold=True), acik
        for o in [x for x in satirlar if x.get('fason')]:
            ws.append([tesis, o['ad'], o['adet'], o['kod'], o['rpr'], o['cfi'], o.get('cop'), o.get('fire'),
                       None, None, None, None, None, None,
                       None, None, 'tedarikçinin işi — tesis toplamına ve performansa girmez'])
            ws.cell(row=ws.max_row, column=2).font = Font(italic=True, color='666666')
        for rr in range(ilk, ws.max_row + 1):
            for i in range(1, len(bas) + 1):
                ws.cell(row=rr, column=i).border = kenar
            for col in ('C', 'D', 'E', 'F', 'G'):
                ws[f'{col}{rr}'].number_format = '#,##0'
            ws[f'H{rr}'].number_format = '0.0%'
            for col in ('I', 'K', 'L', 'M'):
                ws[f'{col}{rr}'].number_format = '#,##0.0'
            for col in ('J', 'O', 'P', 'Q'):
                ws[f'{col}{rr}'].number_format = '0%'
        ws.append([])
    alan = f'O5:O{ws.max_row}'
    ws.conditional_formatting.add(alan, CellIsRule(operator='lessThan', formula=[str(HEDEF_PERFORMANS)],
                                                   fill=PatternFill('solid', fgColor='F8CBAD')))
    ws.conditional_formatting.add(alan, CellIsRule(operator='greaterThanOrEqual', formula=[str(HEDEF_PERFORMANS)],
                                                   fill=PatternFill('solid', fgColor='C6EFCE')))
    notlar = [
        'NOTLAR',
        "• Bir kod tesisinde birden çok bölümde tanımlıysa (ör. kaynak + montaj) her bölümün işine sayılır; tesis toplamı kodu bir kez sayar.",
        '• Teorik süre = adet × birim süre. Birim süre önceliği: kapasite Excel\'i (Database saatlik adet) → Forge çevrim süresi (tel: adımların toplamı) → ERP rota süresi. '
        '"Süre kapsamı" = birim süresi bilinen adetin payı; süresiz kodlar teorik süreye girmez.',
        "• Çok adımlı kaynak: ara operasyonlar ('KOD 1. op', 'KOD(hazırlık)') teyit almaz; süreleri teyidi alan ana kodun "
        "birim süresine eklenir (Bölüm Kapasitesi ile aynı kural) — 'Süre kaynağı'nda adım dökümü.",
        '• Performans = teorik süre ÷ toplam çalışma saati (NçS + mesai). Son sütun, eski planlama tablosu gibi yalnız NçS\'ye böler.',
        f"• Tanımsız (Forge'da ve Anaveri'de bölümü olmayan) kodlar: {len(r['tanimsiz'])} kod, {r['tanimsiz_adet']:,.0f} adet — "
        "'Tanımsız kodlar' sayfası; özete girmez.",
        f"• Hammadde (20.* / 21.*, metre-kg): {len(r['hammadde'])} kod — adet toplamına girmez.",
        ("• COP (hurda): ayın COP hareketleri, kodun üretimiyle aynı bölüme yazılır; üretilen adede girmez. "
         "Fire % = COP ÷ üretilen adet (RPR + CFI)." if r.get('cop_okundu') else
         "• COP (hurda): bu okuma COP hareketlerini içermiyor — panelden 'AS400'den hazırla' ile yeniden okuyun."),
        "• İşleme ve fason rapora dahil değil: " + (', '.join(f"{x['bolum']} {x['kod']} kod · {x['adet']:,.0f} adet"
                                                              for x in r.get('rapor_disi') or []) or 'bu ay kod yok') +
        " (yalnız bu bölümlerde tanımlı kodlar; başka bölümde de tanımlı kod orada sayılır).",
        "• Fason (tedarikçi): Pandora / Pull telleri (Makine 'Pandora'), kodu W ile biten teller ve Anaveri'de 'Fason' işaretli kodlar "
        "(Forge'da tanımı yoksa) — rapora dahil edilmez.",
        f"• Rapordan hariç tutulan ({', '.join(k['desen'] + ('…' if k['tur'] == 'onek' else '') for k in r.get('haric_kurallar') or []) or 'kural yok'}): "
        f"{len(r.get('haric') or [])} kod, {r.get('haric_adet') or 0:,.0f} adet — 'Hariç tutulan' sayfası; hiçbir toplama ve performansa girmez.",
    ]
    if r.get('calisma_bilgi'):
        notlar.append(f"• Çalışma saati kaynağı: {r['calisma_bilgi']['dosya']} ({r['calisma_bilgi']['guncellendi']}).")
    else:
        notlar.append('• Bu ay için çalışma saati yüklenmedi — performans hesaplanmadı.')
    for n in notlar:
        ws.append([n])
        ws.cell(row=ws.max_row, column=1).font = Font(bold=(n == 'NOTLAR'), color='444444')
    for i, w in enumerate([12, 22, 14, 10, 12, 12, 12, 9, 14, 11, 14, 12, 14, 7, 12, 8, 14], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'C5'
    # Bölüm sayfaları
    for o in r['ozet']:
        if not o['satirlar']:
            continue
        ad = f"{o['tesis']} {o['ad']}"[:31].replace('/', '-')
        w = wb.create_sheet(ad)
        w.append([f"{o['tesis']} {o['ad']} — {r['ay_ad']} · {o['kod']} kod · {o['adet']:,.0f} adet · " +
                  ('tedarikçinin işi — tesis toplamına ve performansa girmez' if o.get('fason')
                   else f"teorik {o['teorik_saat']:,.1f} saat")])
        w['A1'].font = Font(bold=True, size=12)
        b2 = ['Kod', 'Açıklama', 'RPR', 'CFI', 'Toplam adet', 'COP (hurda)', 'Birim süre (sn)', 'Süre kaynağı',
              'Teorik süre (saat)', 'Sınıflama', 'Depo', 'Not', 'Forge adet']
        w.append(b2)
        for i in range(1, len(b2) + 1):
            c = w.cell(row=2, column=i)
            c.font, c.fill = Font(bold=True, color='FFFFFF'), koyu
        for s in o['satirlar']:
            w.append([s['kod'], s['aciklama'], s['rpr'] or None, s['cfi'] or None, s['adet'], s.get('cop') or None,
                      s['sure_sn'], s['sure_kaynak'] + (f" ({s['sure_adim']})" if s.get('sure_adim') else ''),
                      s['teorik_saat'], {'forge': 'Forge', 'anaveri': 'Anaveri'}.get(s['sinif'], s['sinif']),
                      s['depolar'], 'birden çok bölümde tanımlı' if s['cok_bolum'] else '', s.get('forge_adet') or None])
        for x in o.get('forge_yalniz') or []:
            w.append([x['kod'], '', None, None, None, None, x['sure_sn'], x['sure_kaynak'], None, '', '',
                      "yalnız Forge'da (ERP'de bu bölümde teyit yok)", x['adet']])
        for i, wd in enumerate([18, 34, 10, 10, 12, 11, 12, 14, 13, 10, 18, 22, 11], start=1):
            w.column_dimensions[get_column_letter(i)].width = wd
        w.freeze_panes = 'B3'
        w.auto_filter.ref = f'A2:M{w.max_row}'
    # FORGE KARŞILAŞTIRMASI (2026-11-03)
    w = wb.create_sheet('Forge karşılaştırma')
    w.append([f"ERP / IK ↔ Forge — {r['ay_ad']}: adet (teyitli ↔ Forge bitmiş ürün), çalışma saati (IK ↔ vardiya), performans"])
    w['A1'].font = Font(bold=True, size=12)
    bas = ['Tesis', 'Bölüm', 'Adet ERP', 'Adet Forge', 'Fark (Forge − ERP)', 'Teorik saat ERP', 'Teorik saat Forge',
           'Çalışma IK (NçS+mesai)', 'Çalışma Forge (vardiya − mola)', 'Forge mola', 'Kişi IK', 'Kişi Forge',
           'Performans ERP ÷ IK', 'Performans Forge ÷ Forge', "Yalnız Forge'da kod", "Yalnız Forge'da adet"]
    w.append(bas)
    for i in range(1, len(bas) + 1):
        c = w.cell(row=2, column=i)
        c.font, c.fill = Font(bold=True, color='FFFFFF'), koyu
        c.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
    for o in r['ozet']:
        f = o.get('forge')
        if not f:
            continue
        w.append([o['tesis'], o['ad'], o['adet'], f['adet'], round(f['adet'] - o['adet'], 2), o['teorik_saat'], f['teorik_saat'],
                  o['calisma'] or None, f['calisma'] or None, f['mola'] or None, o['kisi'] or None, f['kisi'] or None,
                  o['performans'], f['performans'], f['yalniz_sayi'] or None, f['yalniz_adet'] or None])
        for col in (13, 14):
            w.cell(row=w.max_row, column=col).number_format = '0%'
    w.append([])
    for x in ['Forge adet: operatörlerin üretim kayıtları, bitmiş ürün — ara operasyon, tamir / punta / rework, TK1 test cihazı '
              've işlem eki taşıyan (delme, zımpara, somun…) kayıtlar ayrıca sayılmaz; yalnız işlem kaydı olan parçada en büyük '
              'işlem, tel\'de kablonun son adımı. İşleme ve fason rapora dahil değil.',
              'Teorik saat: iki tarafta da aynı birim süre — fark yalnız adetten gelir.',
              'Forge çalışma: vardiya giriş–çıkış, kişi-gün başına birleşim (paralel vardiya çift sayılmaz), çay / yemek molası '
              've namaz düşülmüş; setup / fikstür gibi planlı işler çalışma sayılır.']:
        w.append([x])
    for i, wd in enumerate([7, 22, 11, 11, 13, 12, 12, 14, 16, 10, 8, 9, 13, 14, 12, 12], start=1):
        w.column_dimensions[get_column_letter(i)].width = wd
    w.row_dimensions[2].height = 44
    w.freeze_panes = 'C3'
    # DURUŞLAR DÜŞÜLMÜŞ (2026-11-03)
    w = wb.create_sheet('Duruş düşülmüş')
    w.append([f"Duruşlar düşülmüş performans — {r['ay_ad']}: mola dışı planlı + plansız duruşlar (Forge) "
              "IK saatinden ve Forge vardiya saatinden düşülür"])
    w['A1'].font = Font(bold=True, size=12)
    bas = ['Tesis', 'Bölüm', 'Duruş (saat)', 'Planlı', 'Plansız', 'En çok duruş', 'Teorik ERP', 'Çalışma IK', 'IK − duruş',
           'Performans ERP ÷ IK', 'Performans ERP ÷ (IK − duruş)', 'Teorik Forge', 'Çalışma Forge', 'Forge − duruş',
           'Performans Forge', 'Performans Forge (duruş düşülmüş)']
    w.append(bas)
    for i in range(1, len(bas) + 1):
        c = w.cell(row=2, column=i)
        c.font, c.fill = Font(bold=True, color='FFFFFF'), koyu
        c.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
    for o in r['ozet']:
        f, du = o.get('forge'), o.get('durus')
        if not f or not du:
            continue
        w.append([o['tesis'], o['ad'], du['saat'] or None, du['planli'] or None, du['plansiz'] or None,
                  ' · '.join(f"{x['sebep']} {x['saat']:g}" for x in du['sebepler'][:4]),
                  o['teorik_saat'], o['calisma'] or None, du['ik_net'] or None, o['performans'], du['performans_erp'],
                  f['teorik_saat'], f['calisma'] or None, du['forge_net'] or None, f['performans'], du['performans_forge']])
        for col in (10, 11, 15, 16):
            w.cell(row=w.max_row, column=col).number_format = '0%'
    w.append([])
    w.append(['Duruş: Forge duruş kayıtları — çay / yemek molası ve namaz hariç (onlar çalışma saatinden zaten düşülüyor); '
              'aynı kişinin paralel vardiyalarında aynı sebep + başlangıç bir kez; bir günün duruşu o günün çalışmasını aşamaz. '
              'IK saatinden de AYNI Forge duruşu düşülür (IK duruş tutmaz).'])
    for i, wd in enumerate([7, 22, 10, 9, 9, 46, 11, 11, 11, 13, 15, 11, 12, 12, 12, 15], start=1):
        w.column_dimensions[get_column_letter(i)].width = wd
    w.row_dimensions[2].height = 44
    w.freeze_panes = 'C3'
    # Tanımsız + hammadde + depo + çalışma saati
    for baslik, liste, kol, kol_bas in (('Tanımsız kodlar', r['tanimsiz'], 'anaveri', 'Anaveri (TK-1/2 / Hat)'),
                                        ('Hariç tutulan', r.get('haric') or [], 'kural', 'Kural'),
                                        ('Hammadde (adet dışı)', r['hammadde'], None, None)):
        w = wb.create_sheet(baslik)
        w.append(['Kod', 'Açıklama', 'Tesis (depo)', 'RPR', 'CFI', 'Toplam', 'COP (hurda)', 'Depo']
                 + ([kol_bas] if kol else []))
        for c in w[1]:
            c.font, c.fill = Font(bold=True, color='FFFFFF'), koyu
        for s in liste:
            w.append([s['kod'], s['aciklama'], s['tesis'], s['rpr'] or None, s['cfi'] or None, s['adet'],
                      s.get('cop') or None, s['depolar']] + ([s.get(kol, '')] if kol else []))
        for i, wd in enumerate([18, 34, 10, 10, 10, 12, 11, 18, 24], start=1):
            w.column_dimensions[get_column_letter(i)].width = wd
    w = wb.create_sheet('Depo ve çalışma saati')
    w.append(['Depo', 'Tesis', 'RPR', 'CFI', 'COP (hurda)', 'Kod sayısı'])
    for c in w[1]:
        c.font, c.fill = Font(bold=True, color='FFFFFF'), koyu
    for d in r['depolar']:
        w.append([d['depo'], d['tesis'], d['RPR'], d['CFI'], d.get('COP'), d['kod']])
    w.append([])
    w.append(['Çalışma saati (kapasite Excel\'i)', 'Forge bölümü', 'NçS', 'Mesai', 'Kişi'])
    for c in w[w.max_row]:
        c.font, c.fill = Font(bold=True, color='FFFFFF'), koyu
    for o in r['ozet']:
        if o['calisma']:
            w.append([', '.join(o['calisma_excel']), f"{o['tesis']} {o['ad']}", o['nos'], o['mesai'], o['kisi']])
    for d in r['calisma_disarida']:
        w.append([d['excel_bolum'], '— üretim dışı / atanmamış', d['nos'], d['mesai'], d['kisi']])
    for i, wd in enumerate([34, 26, 12, 12, 10], start=1):
        w.column_dimensions[get_column_letter(i)].width = wd
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()
