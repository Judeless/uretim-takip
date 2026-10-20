# -*- coding: utf-8 -*-
"""MARKALI KURULUM HAZIRLAMA — sıfırdan yeni firma kurulumu (2026-10-02).

Aynı kodun başka bir firmaya (ör. tedarikçi) kurulumu. Bu betik YENİ bir klasörde
(bulut sunucuda git clone'dan sonra) BİR KEZ çalıştırılır:

    python kurulum_hazirla.py --profil ornek_kurulum/tedarikci.json \
        --excel "ÜRETİM SÜRELERİ DATA.xlsb" --yonetici admin --yonetici-ad "Ad Soyad"

Yaptıkları:
  1. Profili data/kurulum.json olarak kopyalar (zaten varsa ona dokunmaz)
  2. Boş veritabanını kurar — Cofle'nin bilinen şifreli yöneticisi AÇILMAZ
  3. Yönetici hesabını RASTGELE şifreyle açar; şifre yalnız ekrana bir kez yazılır
     (ilk girişte değiştirilmesi istenir)
  4. Excel verilmişse referansları içe alır (metal: göz + parça başı süre; CNC
     işleme: delme/diş/taşlama; montaj: bileşenler; plastik: 'PLASTİK ENJEKSİYON
     KODLAR' bloğu + tonaj sayfasındaki plastik hammaddeli satırlar)
  5. Bölüm başına varsayılan duruş sebeplerini ekler (panelden değiştirilir)

Kurulumdan SONRA Excel'e yeni referans eklenince (var olanlara dokunmaz):

    python kurulum_hazirla.py --sadece-referans --excel "ÜRETİM SÜRELERİ DATA.xlsb"

GÜVENLİK: Cofle'nin kendi kurulumunu DÖNÜŞTÜRMEZ — klasörde üretim verisi olan bir
veritabanı ya da Cofle'nin varsayılan yöneticisi varsa hiçbir şey yazmadan durur.
"""
import argparse
import json
import os
import re
import secrets
import shutil
import sqlite3
import sys

KOK = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, KOK)
for _a in (sys.stdout, sys.stderr):
    try:
        _a.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

HEDEF_PROFIL = os.path.join(KOK, 'data', 'kurulum.json')
DB = os.path.join(KOK, 'uretim.db')

DURUS_ORTAK = [('Mola / Yemek', 'planli'), ('Planlı Bakım', 'planli'), ('Toplantı / Eğitim', 'planli'),
               ('Makine Arızası', 'plansiz'), ('Malzeme Bekleme', 'plansiz'),
               ('Kalite Problemi', 'plansiz'), ('Operatör Yok', 'plansiz'),
               ('Elektrik / Hava Kesintisi', 'plansiz'), ('Ayar', 'plansiz')]
DURUS_BOLUM = {
    'metal': [('Kalıp Değişimi', 'planli'), ('Kalıp Arızası', 'plansiz'), ('Kalıp Isıtma', 'planli')],
    'isleme': [('Takım Değişimi', 'planli'), ('Program / Ayar', 'planli'), ('Takım Kırılması', 'plansiz')],
    'montaj': [('Model Değişimi', 'planli'), ('Eksik Parça', 'plansiz')],
    'plastik': [('Kalıp Değişimi', 'planli'), ('Kalıp Arızası', 'plansiz'), ('Malzeme Kurutma', 'planli')],
}


def dur(mesaj):
    print('\nDURDU: ' + mesaj)
    sys.exit(1)


# ── güvenlik ──────────────────────────────────────────────────────────────
def cofle_verisi_var_mi():
    if not os.path.exists(DB):
        return None
    try:
        cn = sqlite3.connect(DB)
        try:
            tablolar = {r[0] for r in cn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            if 'vardiyalar' in tablolar and cn.execute('SELECT COUNT(*) FROM vardiyalar').fetchone()[0]:
                return 'veritabanında vardiya kaydı var'
            if 'panel_kullanicilari' in tablolar and cn.execute(
                    "SELECT 1 FROM panel_kullanicilari WHERE kullanici_adi='emre.dogutekin'").fetchone():
                return "Cofle'nin varsayılan yöneticisi var"
        finally:
            cn.close()
    except sqlite3.Error as e:
        return f'veritabanı okunamadı: {e}'
    return None


# ── Excel okuma ───────────────────────────────────────────────────────────
def excel_sayfalari(yol):
    """{sayfa adı: satırlar} — sayfa sırası korunur. .xlsb → pyxlsb, .xlsx → openpyxl."""
    if yol.lower().endswith('.xlsb'):
        from pyxlsb import open_workbook
        out = {}
        with open_workbook(yol) as wb:
            for ad in wb.sheets:
                with wb.get_sheet(ad) as s:
                    out[ad] = [[c.v for c in r] for r in s.rows()]
        return out
    import openpyxl
    wb = openpyxl.load_workbook(yol, data_only=True, read_only=True)
    return {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}


def excel_satirlari(yol):
    """İlk sayfanın satırları (liste listesi)."""
    return next(iter(excel_sayfalari(yol).values()))


def _sayi(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _kod(v):
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _gecersiz_kod(kod):
    """'6' / '4' gibi kısa salt sayılar kod değil, hücreye kayan adet/göz değeri."""
    return kod.isdigit() and len(kod) < 5


def _kod_ayir(v):
    """Plastik tablosundaki kod hücresi → [(kod, not)]:
    '10.300.2607 ( İÇ PARÇA)' → not olarak ayrılır;
    '10.300.0499-0500' / '10.300.2612 - 0500' (aile kalıbı, tek baskıda iki parça) → iki kod."""
    s = _kod(v)
    m = re.match(r'^(\d{2}\.\d{3}\.\d{4}[A-Z]?)\s*\(\s*(.*?)\s*\)\s*$', s)
    if m:
        return [(m.group(1), m.group(2))]
    m = re.match(r'^(\d{2}\.\d{3}\.)(\d{4})\s*-\s*(\d{4})$', s)
    if m:
        return [(m.group(1) + m.group(2), f'aile kalıbı {s}'), (m.group(1) + m.group(3), f'aile kalıbı {s}')]
    return [(s, '')]


def _tekille(liste):
    """Aynı kod birden çok satırda: süresi DOLU olan kalır (yoksa ilki)."""
    out = {}
    for r in liste:
        if r['kod'] not in out or (not out[r['kod']]['ct'] and r['ct']):
            out[r['kod']] = r
    return list(out.values())


# Tonaj sayfasında hammaddesi bu desene uyan satır METAL, diğer dolu hammaddeler
# (PA66.GF30, PP, ABS, ASA, TPE...) PLASTİK enjeksiyon ürünüdür
_METAL_HAMMADDE = re.compile(r'zamak|zamac|alümin|alumin|pirin|brass', re.I)


def _makine(v):
    """'50 Ton' / '51 Ton' / '100 TON' / '50/100 Ton' → '50 Ton' / '100 Ton' / '50 / 100 Ton'.
    51/53/101 gibi değerler veri girişi farkı — en yakın sınıfa yuvarlanır."""
    sayilar = [int(x) for x in re.findall(r'\d+', str(v or ''))]
    if not sayilar:
        return str(v or '').strip()
    sinif = sorted({50 if s < 75 else 100 for s in sayilar})
    return ' / '.join(str(s) for s in sinif) + ' Ton'


def referanslari_cikar(satirlar):
    """Tedarikçinin 'general table' biçimi: başlık satırı 'Group/Ext. Code' içerir."""
    bas = next((i for i, r in enumerate(satirlar[:15])
                if any(str(x or '').strip().lower() == 'group/ext. code' for x in r)), None)
    if bas is None:
        dur("Excel'de 'Group/Ext. Code' başlıklı satır bulunamadı — beklenen biçim değil")
    h = [str(x or '').strip().lower() for x in satirlar[bas]]

    def kol(*adlar):
        for ad in adlar:
            if ad in h:
                return h.index(ad)
        return None
    k = {'kod': kol('group/ext. code'), 'makine': kol('machine'), 'goz': kol('# of cavities'),
         'ct': kol('cycle time (sec)'), 'isl': kol('machining description'), 'del': kol('drilling'),
         'dis': kol('threading'), 'tas': kol('grinding'), 'mon': kol('assembled parts')}
    eksik = [a for a in ('kod', 'goz', 'ct') if k[a] is None]
    if eksik:
        dur(f'Excel başlığında eksik kolon: {eksik}')
    metal, isleme, montaj, gorulen = [], [], [], set()
    for r in satirlar[bas + 1:]:
        if not r or k['kod'] >= len(r):
            continue
        kod = _kod(r[k['kod']])
        if not kod or kod in gorulen or len(kod) > 40:
            continue
        gorulen.add(kod)
        goz = max(1, int(_sayi(r[k['goz']]) or 1))
        baski = _sayi(r[k['ct']])
        mk = _makine(r[k['makine']]) if k['makine'] is not None else ''
        metal.append({'kod': kod, 'goz': goz, 'ct': round(baski / goz, 2) if baski else 0,
                      'aciklama': f'{mk} · {goz} göz · {baski:g} sn/baskı'.strip(' ·')})
        if k['del'] is not None:
            d, t, g = (int(_sayi(r[k[x]])) if k[x] is not None and k[x] < len(r) else 0
                       for x in ('del', 'dis', 'tas'))
            if d or t or g:
                ops = ' · '.join(f'{ad} {n}' for ad, n in (('Delme', d), ('Diş', t), ('Taşlama', g)) if n)
                isleme.append({'kod': kod, 'aciklama': ops})
        if k['mon'] is not None:
            parcalar = [str(x).strip() for x in r[k['mon']:k['mon'] + 8]
                        if x not in (None, '') and re.match(r'^\d{2}\.', str(x).strip())
                        and '+' not in str(x)]          # 'a+b+c' özet hücresi tekrar etmesin
            if parcalar:
                montaj.append({'kod': kod, 'aciklama': 'Bileşenler: ' + ', '.join(dict.fromkeys(parcalar))})
    return metal, isleme, montaj, plastik_bloku(satirlar, bas, k['kod'])


def plastik_bloku(satirlar, bas, ilk_kod_kol):
    """Aynı sayfada metal tablosunun SAĞINDA ikinci bir 'Group/Ext. Code' bloğu,
    üstündeki başlıkta 'PLASTİK ENJEKSİYON KODLAR' (Aspanda Excel'i, BJ sütunu): kod ·
    göz · baskı süresi. Makine yok — plastikte operatör makineyi (107/350 Ton) vardiyada seçer."""
    h = [str(x or '').strip().lower() for x in satirlar[bas]]
    out = []
    for j, x in enumerate(h):
        if j == ilk_kod_kol or x != 'group/ext. code':
            continue
        baslik = ' '.join(str(satirlar[i][c] or '') for i in range(bas)
                          for c in range(max(0, j - 1), j + 3) if c < len(satirlar[i]))
        if 'PLAST' not in baslik.upper():
            continue
        yakin = h[j:j + 10]
        k_goz = j + yakin.index('# of cavities') if '# of cavities' in yakin else None
        k_ct = j + yakin.index('cycle time (sec)') if 'cycle time (sec)' in yakin else None
        for r in satirlar[bas + 1:]:
            if not r or j >= len(r):
                continue
            goz = max(1, int(_sayi(r[k_goz]) or 1)) if k_goz is not None and k_goz < len(r) else 1
            baski = _sayi(r[k_ct]) if k_ct is not None and k_ct < len(r) else 0
            for kod, notu in _kod_ayir(r[j]):
                if not kod or len(kod) > 40 or _gecersiz_kod(kod):
                    continue
                aciklama = f'{goz} göz · {baski:g} sn/baskı' if baski else f'{goz} göz'
                out.append({'kod': kod, 'goz': goz, 'ct': round(baski / goz, 2) if baski else 0,
                            'aciklama': aciklama + (f' · {notu}' if notu else '')})
    return _tekille(out)


def _tonaj_bloku(ad, satirlar):
    """'200T' sayfasının makinesi + başlık satırı + kolonları; biçim uymuyorsa None."""
    m = re.match(r'^\s*(\d+)\s*T(?:on)?\s*$', str(ad), re.I)
    if not m:
        return None
    bas = next((i for i, r in enumerate(satirlar[:15])
                if any(str(x or '').strip().lower() == 'enjeksiyon prosesi' for x in r)), None)
    if bas is None:
        return None
    h = [str(x or '').strip().lower() for x in satirlar[bas]]
    kol = {a: (h.index(b) if b in h else None) for a, b in
           (('kod', 'enjeksiyon prosesi'), ('goz', 'göz adedi'), ('ct', 'çevrim süresi'), ('ham', 'hammadde cinsi'))}
    return f'{int(m.group(1))} Ton', bas, kol


def tonaj_sayfasi_refleri(ad, satirlar):
    """'200T' gibi tek makinelik maliyet sayfası (Aspanda Excel'i, 2026-10-19): başlık
    satırında 'Enjeksiyon Prosesi' (kod) · 'Göz Adedi' · 'Çevrim Süresi' (baskı süresi).
    Makine sayfa adından gelir: '200T' → '200 Ton'. Biçim uymuyorsa boş liste.
    Sayfa zamak ile plastiği KARIŞIK tutar (2026-10-20): 'Hammadde Cinsi' ZAMAK → metal
    (makine = sayfa), PA66/PP/ABS/ASA/TPE... → plastik (her kayıtta 'bolum' anahtarı)."""
    blok = _tonaj_bloku(ad, satirlar)
    if not blok:
        return []
    mk, bas, k = blok

    def hucre(r, a):
        return r[k[a]] if k[a] is not None and k[a] < len(r) else None
    out, gorulen = [], set()
    for r in satirlar[bas + 1:]:
        if not r:
            continue
        kod = _kod(hucre(r, 'kod'))
        if not kod or kod in gorulen or len(kod) > 40 or _gecersiz_kod(kod):
            continue
        gorulen.add(kod)
        goz = max(1, int(_sayi(hucre(r, 'goz')) or 1))
        baski = _sayi(hucre(r, 'ct'))
        ham = str(hucre(r, 'ham') or '').strip()
        ct = round(baski / goz, 2) if baski else 0
        if ham and not _METAL_HAMMADDE.search(ham):
            out.append({'bolum': 'plastik', 'kod': kod, 'goz': goz, 'ct': ct,
                        'aciklama': f'{ham} · {goz} göz · {baski:g} sn/baskı'})
        else:
            out.append({'bolum': 'metal', 'kod': kod, 'goz': goz, 'ct': ct,
                        'aciklama': f'{mk} · {goz} göz · {baski:g} sn/baskı'})
    return out


def _tonaj_eski_kodlari(ad, satirlar):
    """Eski sürümün (2026-10-19) bu sayfadan METAL diye yazdığı kodların tamamı —
    düzeltme modunda yanlış bölüme düşenleri bulmak için."""
    blok = _tonaj_bloku(ad, satirlar)
    if not blok:
        return None
    mk, bas, k = blok
    return {kod for r in satirlar[bas + 1:] if r and k['kod'] < len(r)
            for kod in [_kod(r[k['kod']])] if kod and len(kod) <= 40}, mk


def excel_referanslari(yol):
    """Excel → {bölüm: [referans]} (metal · isleme · montaj · plastik)."""
    sayfalar = excel_sayfalari(yol)
    metal, isleme, montaj, plastik = referanslari_cikar(next(iter(sayfalar.values())))
    # Tonaj sayfaları ('200T'): o makinede basılan zamak ürünler metal, plastik hammaddeliler plastik
    for _ad, _sat in sayfalar.items():
        for r in tonaj_sayfasi_refleri(_ad, _sat):
            (plastik if r['bolum'] == 'plastik' else metal).append(r)
    return {'metal': metal, 'isleme': isleme, 'montaj': montaj, 'plastik': plastik}, sayfalar


def referans_ekle(cn, refs, bolumler, lok):
    """INSERT OR IGNORE — var olan referansa (süresi/açıklaması elle düzeltilmiş olabilir) DOKUNMAZ."""
    sayac = {b: 0 for b in refs}
    for bolum, liste in refs.items():
        if bolum not in bolumler:
            continue
        for r in liste:
            sayac[bolum] += cn.execute(
                "INSERT OR IGNORE INTO referans_listesi (referans_kodu, aciklama, hedef_cycle_time_sn, "
                "bolum, lokasyon, kalip_goz, paket_adedi, sure_teyit) VALUES (?,?,?,?,?,?,?,0)",
                (r['kod'], r['aciklama'], r.get('ct', 0), bolum, lok, r.get('goz', 1),
                 r.get('goz', 1) if bolum in ('metal', 'plastik') else 1)).rowcount
    return sayac


def referans_guncelle(excel):
    """--sadece-referans: MEVCUT markalı kurulumda Excel'deki EKSİK referansları ekler.
    Profil/yönetici/duruş sebeplerine dokunmaz; var olan referansı değiştirmez. Eski sürümün
    tonaj sayfasından METAL diye yazdığı plastik ürünleri (ve '6'/'4' gibi bozuk kodları)
    siler — yalnız üretim kaydında hiç kullanılmamışsa."""
    if not excel:
        dur('--sadece-referans için --excel gerekli')
    if not os.path.exists(HEDEF_PROFIL) or not os.path.exists(DB):
        dur('Bu klasörde kurulum yok (data/kurulum.json ya da uretim.db eksik)')
    import kurulum
    p = kurulum.yukle(tazele=True)
    cn = sqlite3.connect(DB, timeout=30)
    try:
        if kurulum.cofle_mi() or cn.execute(
                "SELECT 1 FROM panel_kullanicilari WHERE kullanici_adi='emre.dogutekin'").fetchone():
            dur('Bu Cofle kurulumu — referans güncelleme yalnız markalı kurulumda çalışır')
        bolumler = p.get('bolumler') or []
        lok = ((p.get('lokasyonlar') or [{}])[0] or {}).get('kod') or 'TK2'
        refs, sayfalar = excel_referanslari(excel)
        yeni_metal = {r['kod'] for r in refs['metal']}
        silinen, korunan = [], []
        for _ad, _sat in sayfalar.items():
            eski = _tonaj_eski_kodlari(_ad, _sat)
            if not eski:
                continue
            kodlar, mk = eski
            for kod in sorted(kodlar - yeni_metal):
                if not cn.execute("SELECT 1 FROM referans_listesi WHERE referans_kodu=? AND bolum='metal' "
                                  "AND lokasyon=? AND aciklama LIKE ?", (kod, lok, f'{mk} ·%')).fetchone():
                    continue
                if cn.execute('SELECT 1 FROM uretim_kayitlari WHERE referans_kodu=? LIMIT 1', (kod,)).fetchone():
                    korunan.append(kod)
                    continue
                cn.execute("DELETE FROM referans_listesi WHERE referans_kodu=? AND bolum='metal' AND lokasyon=? "
                           "AND aciklama LIKE ?", (kod, lok, f'{mk} ·%'))
                silinen.append(kod)
        sayac = referans_ekle(cn, refs, bolumler, lok)
        cn.commit()
        toplam = dict(cn.execute('SELECT bolum, COUNT(*) FROM referans_listesi WHERE lokasyon=? GROUP BY bolum',
                                 (lok,)).fetchall())
    finally:
        cn.close()
    print(f"Metalden çıkarılan (plastik ya da bozuk kod): {len(silinen)}"
          + (f"  — {', '.join(silinen[:8])}{' …' if len(silinen) > 8 else ''}" if silinen else ''))
    if korunan:
        print(f"Üretim kaydı olduğu için metalde BIRAKILAN: {', '.join(korunan)}")
    print('Eklenen: ' + ' · '.join(f'{b} {n}' for b, n in sayac.items() if b in bolumler))
    print('Toplam : ' + ' · '.join(f'{b} {toplam.get(b, 0)}' for b in bolumler))
    print('\nTAMAM — yeniden başlatma gerekmez, panelde Referanslar sayfasını yenileyin.')


def main():
    ap = argparse.ArgumentParser(description='Markalı kurulumu sıfırdan hazırlar.')
    ap.add_argument('--profil', help='kurulum profili (ör. ornek_kurulum/tedarikci.json)')
    ap.add_argument('--excel', help='referans/süre verisi (.xlsb / .xlsx)')
    ap.add_argument('--yonetici', default='admin', help='yönetici kullanıcı adı')
    ap.add_argument('--yonetici-ad', default='Yönetici', help='yöneticinin adı soyadı')
    ap.add_argument('--sadece-referans', action='store_true',
                    help="MEVCUT markalı kurulumda yalnız Excel'deki eksik referansları ekler")
    a = ap.parse_args()

    if a.sadece_referans:
        referans_guncelle(a.excel)
        return

    neden = cofle_verisi_var_mi()
    if neden:
        dur(f'Bu klasör mevcut bir kurulum gibi görünüyor ({neden}). Betik yalnız BOŞ, '
            f'yeni bir klasörde çalışır — Cofle kurulumunu dönüştürmez.')

    if not os.path.exists(HEDEF_PROFIL):
        if not a.profil:
            dur('data/kurulum.json yok ve --profil verilmedi.')
        with open(a.profil, encoding='utf-8-sig') as f:
            json.load(f)                                  # bozuk profil kopyalanmasın
        os.makedirs(os.path.dirname(HEDEF_PROFIL), exist_ok=True)
        shutil.copyfile(a.profil, HEDEF_PROFIL)
        print(f'1. Profil kopyalandı → {HEDEF_PROFIL}')
    else:
        print(f'1. Profil zaten var → {HEDEF_PROFIL} (dokunulmadı)')

    import kurulum
    p = kurulum.yukle(tazele=True)
    if kurulum.cofle_mi() or p.get('varsayilan_yonetici', True):
        dur("Profil Cofle kimliğinde ya da 'varsayilan_yonetici' açık — markalı kurulumda "
            "kimlik 'cofle' olamaz ve 'varsayilan_yonetici': false olmalı.")

    import database
    database.init_db()
    print(f'2. Veritabanı kuruldu → {DB}')

    from werkzeug.security import generate_password_hash
    cn = sqlite3.connect(DB)
    try:
        sifre = None
        if not cn.execute('SELECT 1 FROM panel_kullanicilari WHERE kullanici_adi=?', (a.yonetici,)).fetchone():
            sifre = secrets.token_urlsafe(9)
            cn.execute("INSERT INTO panel_kullanicilari (kullanici_adi, ad_soyad, sifre_hash, rol, izinler, "
                       "aktif, sifre_gecici) VALUES (?,?,?,'admin','[]',1,1)",
                       (a.yonetici, a.yonetici_ad, generate_password_hash(sifre)))
        bolumler = p.get('bolumler') or []
        lok = ((p.get('lokasyonlar') or [{}])[0] or {}).get('kod') or 'TK2'
        # Mobilde TÜM vardiyalara erişen 'Admin' operatörü — rastgele 4 haneli PIN
        # (Cofle'deki 9999 internete açık sunucuda herkesin bildiği bir PIN olurdu)
        admin_pin = None
        if not cn.execute("SELECT 1 FROM operatorler WHERE ad='Admin' AND lokasyon=?", (lok,)).fetchone():
            admin_pin = ''.join(secrets.choice('0123456789') for _ in range(4))
            while admin_pin in ('0000', '1234', '9999') or len(set(admin_pin)) == 1:
                admin_pin = ''.join(secrets.choice('0123456789') for _ in range(4))
            cn.execute("INSERT INTO operatorler (ad, bolum, pin, lokasyon) VALUES ('Admin', ?, ?, ?)",
                       ((bolumler or ['montaj'])[0], admin_pin, lok))

        sayac = {'metal': 0, 'isleme': 0, 'montaj': 0, 'plastik': 0}
        if a.excel:
            sayac = referans_ekle(cn, excel_referanslari(a.excel)[0], bolumler, lok)

        durus = 0
        for b in bolumler:
            for i, (sebep, tip) in enumerate(DURUS_ORTAK + DURUS_BOLUM.get(b, []), start=1):
                durus += cn.execute(
                    "INSERT OR IGNORE INTO durus_sebebi (lokasyon, bolum, sebep, tip, aktif, sira, olusturan) "
                    "VALUES (?,?,?,?,1,?,'kurulum')", (lok, b, sebep, tip, i)).rowcount
        cn.commit()
    finally:
        cn.close()

    print(f'3. Yönetici: {a.yonetici}' + (f'  ·  GEÇİCİ ŞİFRE: {sifre}' if sifre else ' (zaten vardı, şifresine dokunulmadı)'))
    if sifre:
        print('   Bu şifre bir daha gösterilmez — ilk girişte değiştirmeniz istenecek.')
    if admin_pin:
        print(f"   Mobil 'Admin' operatörü PIN: {admin_pin}  (tüm vardiyaları düzenler — panelden değiştirilebilir)")
    if a.excel:
        print(f"4. Referanslar: metal {sayac['metal']} · plastik {sayac['plastik']} · "
              f"CNC işleme {sayac['isleme']} · montaj {sayac['montaj']}")
    else:
        print('4. Referans Excel\'i verilmedi — referanslar panelden eklenir')
    print(f'5. Varsayılan duruş sebepleri: {durus} (panel → Duruş Sebepleri)')
    print(f"\nHAZIR — {p['marka']['ad']}. Başlatmak: python app.py")


if __name__ == '__main__':
    main()
