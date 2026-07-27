# -*- coding: utf-8 -*-
"""Gunun uretimini AS400 acik launch'lariyla esler → sabah teyit is listesi.

Kaynaklar:
  - uretim.db (mail_raporu.gunluk_veri ile ayni sorgu mantigi: gun + referans + adet)
  - AS400 XB530W0 work-file (10>05>03>29 raporu; durum 40 = acik launch)

Cikti kategorileri:
  ACIK   : referansin durum-40 launch'i var → teyit verilebilir
  KAPALI : listede var ama durumu 40 degil (45/50 vb.) → kontrol gerek
  YOK    : listede hic yok → launch alinmamis (durum 10) VEYA teyidi coktan verilmis

Kullanim:  python as400/launch_esle.py [YYYY-MM-DD]   (varsayilan: dun)
"""
import sys, io, os, re, sqlite3, collections
from datetime import date, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import keyring, pyodbc
import as400_config as CFG

PROJE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
URETIM_DB = os.path.join(PROJE, 'uretim.db')


def kanonik(s):
    """Referans kodunu eslesme anahtarina cevirir.
    - '... 1.op' / '2.op' operasyon eklerini atar
    - '(hazırlık)' gibi parantezli ekleri atar
    - '1.ist'/'2.ist' istasyon eklerini atar
    - tum bosluklari siler, buyuk harfe cevirir
    """
    s = str(s or '')
    s = re.sub(r'\([^)]*\)', '', s)                       # (hazırlık) vb.
    s = re.sub(r'\s*\d+\.?\s*(op|ist)\.?\s*$', '', s, flags=re.I)
    return re.sub(r'\s+', '', s).upper()


def gevsek(s):
    """Ikincil anahtar: nokta/bosluk farklarini yutar ('10 300 0609' ≈ '10.300.0609').
    Yalniz alfasayisal karakterler kalir — eslesme 'SUPHELI' olarak isaretlenir."""
    return re.sub(r'[^A-Z0-9]', '', kanonik(s))


def kok(s):
    """Ucuncul anahtar: sondaki harf ekini (ve pesindeki -xx'i) atar.
    '10.130.5421W' ≈ '10.130.5421GRW', '10.300.6343A' ≈ '10.300.6343A-S'.
    Varyant karisma riski var → eslesme daima 'SUPHELI' (insan onayli)."""
    k = kanonik(s)
    return re.sub(r'[A-Z]+([\-/][A-Z0-9]+)?$', '', k) or k


# ── OP/IST kurali (kullanici 2026-07-17): 'X 1. op', 'X 2.op', 'X (2.Opr)',
# 'X 1.ist' gibi ekler ARA OPERASYON belirtir. Referans listesindeki (ve o gunun
# uretimindeki) EN SON op/ist numarasi hangisiyse yalniz O satir bitmis urundur:
# sondaki ek yokmus gibi (base kod) teyit edilir. Ara op'lar teyit EDILMEZ.
# Kanit: 0541W'de operator yalniz 2.ist adedini (198) teyit etmis.
# Dikkat: ek, kodun SONUNDA ve oncesinde bosluk/parantez olmali —
# '92.2.OP001R' (kod icinde OP) veya 'ABK1.op' (harfe bitisik) EK DEGILDIR.
_OP_EK = re.compile(r'[\s(](\d+)\s*\.?\s*(op[r]?|ist)\s*\.?\s*\)?\s*$', re.IGNORECASE)


def op_bilgi(referans):
    """(base_kod, op_no) dondurur; ek yoksa (referans, None).
    Parantezli aciklamalar ('(hazırlık)') once atilir ki '1.ist (hazırlık)' da yakalansin."""
    temiz = re.sub(r'(?i)\((?!\s*\d+\s*\.?\s*op)[^)]*\)', '', str(referans or '')).strip()
    m = _OP_EK.search(temiz)
    if not m:
        return str(referans or '').strip(), None
    base = temiz[:m.start()].strip().rstrip('(').strip()
    return base, int(m.group(1))


def _referans_listesi_son_op():
    """referans_listesi'nden base→en_buyuk_op haritasi (kanonik anahtarla)."""
    conn = sqlite3.connect(URETIM_DB)
    sonuc = {}
    try:
        for (rk,) in conn.execute("SELECT referans_kodu FROM referans_listesi").fetchall():
            base, op = op_bilgi(rk)
            if op is not None:
                kb = kanonik(base)
                sonuc[kb] = max(sonuc.get(kb, 0), op)
    finally:
        conn.close()
    return sonuc


# Bosluklu kod yazimi (kullanici 2026-07-22): '10 300 1393A' = '10.300.1393A'
# sayilir — nokta yerine bosluk birakilmis kodlar nokta bicimine cevrilerek
# teyit olusturulur. YALNIZ 'sayi sayi kalan' bicimi cevrilir; aciklamali
# kayitlar ('10.130.4031 zimpara', '... somun punta') DOKUNULMAZ.
_BOSLUK_KOD = re.compile(r'^(\d{1,4}) (\d{1,4}) (\S+)$')


def bosluk_nokta(ref):
    m = _BOSLUK_KOD.match(str(ref or '').strip())
    return f'{m.group(1)}.{m.group(2)}.{m.group(3)}' if m else str(ref or '').strip()


# PRES op eki (kullanici 2026-07-22): preste '1. op' / '2. op' yazan kayitlar
# TEYIT ALMAZ — teyit yalniz ayni referansin OP'SUZ yazimina verilir. Pres
# yazimlari dagitik oldugundan nokta/tire ile bitisik ekler de taninir
# ('...1393A.1.op', 'X-2.op', 'X 1. op').
_PRES_OP = re.compile(r'[\s.\-_](\d+)\s*\.?\s*(op[r]?|ist)\.?\s*$', re.IGNORECASE)


def op_kurali_uygula(rows):
    """Gun uretim satirlarina op/ist kuralini uygular.
    Donis: (islenmis_satirlar, ara_op_satirlari). Son op satirinin referansi
    base koda cevrilir (orijinal 'orijinal_referans' alaninda saklanir)."""
    liste_son = _referans_listesi_son_op()
    # O gunku uretimden de base→max op topla (liste eksikse gunun verisi tamamlar)
    gun_son = {}
    for r in rows:
        base, op = op_bilgi(r['referans'])
        if op is not None:
            kb = kanonik(base)
            gun_son[kb] = max(gun_son.get(kb, 0), op)
    islenmis, ara = [], []
    for r in rows:
        # HAZIRLIK on-operasyonu → teyit DISI (nihai teyit base koda verilir).
        # Kullanici 2026-07-20: '2757w-10(hazırlık)' ile base '2757w-10' AYNI launch'a
        # iki ayri teyit adayi oluyordu (mukerrer risk). Hazirligi 1.op gibi ARAOP'a al.
        rl = str(r['referans'] or '').lower()
        if 'hazırlık' in rl or 'hazirlik' in rl:
            base_h = re.sub(r'\s*\([^)]*\)\s*', '', str(r['referans'] or '')).strip()
            ara.append({**r, 'son_op': base_h or 'base kod'})
            continue
        # PRES op kurali (kullanici 2026-07-22): preste op'lu kayit TEYIT ALMAZ,
        # teyit yalniz op'suz yazima verilir (son-op/base cevirisi de YOK).
        if str(r.get('bolum') or '') == 'pres':
            m_p = _PRES_OP.search(str(r['referans'] or '').strip())
            if m_p:
                base_p = str(r['referans'])[:m_p.start()].strip().rstrip('.-_ ')
                ara.append({**r, 'son_op': (base_p + " op'suz yazimina verilir") if base_p else "op'suz yazim"})
                continue
        base, op = op_bilgi(r['referans'])
        if op is None:
            islenmis.append(r)
            continue
        kb = kanonik(base)
        son_op = max(liste_son.get(kb, 0), gun_son.get(kb, 0))
        if op >= son_op:
            yeni = dict(r)
            yeni['orijinal_referans'] = r['referans']
            yeni['referans'] = base          # son op → ek yokmus gibi teyit
            islenmis.append(yeni)
        else:
            ara.append({**r, 'son_op': son_op})
    return islenmis, ara


def as400_launchlar():
    """AS400'den acik launch haritasi: kanonik(article) -> [launch kaydi]"""
    pw = CFG.sifre_al()
    if not pw:
        raise RuntimeError('Sifre kasada yok: python as400/kaydet_sifre.py')
    cn = pyodbc.connect(CFG.baglanti_dizesi(pw), timeout=20, autocommit=True)
    tam = collections.defaultdict(list)
    gev = collections.defaultdict(list)
    kokm = collections.defaultdict(list)
    for r in cn.cursor().execute(CFG.SORGU):
        try:
            durum = int(str(r[4]).strip() or -1)
        except ValueError:
            durum = -1
        qty = float(r[5] or 0)
        teyitli = float(r[6] or 0)
        kayit = {
            'launch': f'{r[0]}-{r[1]}-{r[2]}',
            'red2': str(r[1]).strip(),    # yil (teyit gonderiminde Order No 1. alan)
            'renu': str(r[2]).strip(),    # launch no (Order No 2. alan)
            'article': str(r[3] or '').strip(),
            'durum': durum,
            'qty': qty,
            'teyitli': teyitli,
            'kalan': qty - teyitli,   # negatif = launch adedinden fazla teyit (A-modu)
        }
        tam[kanonik(r[3])].append(kayit)
        gev[gevsek(r[3])].append(kayit)
        kokm[kok(r[3])].append(kayit)
    cn.close()
    return tam, gev, kokm


# Teyit kapsamindaki bolumler. 2026-07-22 kullanici istegi: PRES de dahil
# (abkant = pres bolumunun makineleri: 'Abkant 1/2' + 'Pres' hatlari) — launch'i
# olan launch teyidi, olmayan CFI kuyruguna duser. ISLEME hala DISARIDA
# (delme/rayba/zimpara ara operasyonlari). Yalniz TK2 (TK1 = durum 45/50, sonra).
LAUNCH_BOLUMLERI = ('kaynak', 'metal', 'montaj', 'lazer', 'pres')
LOKASYON = 'TK2'


def teyit_hareketleri(articles):
    """Verilen ARTICLE'larin RPR (uretim teyidi) hareketlerini dondurur —
    LAUNCH'TAN BAGIMSIZ. Cunku operator ayni urunu baska bir launch'a girip
    S ile kapatmis olabilir; o launch durum 70 olur ve acik listede gorunmez.
    Donis: {kanonik(article): [{'tarih','adet','launch'}, ...]}
    Kaynak: BMMAF0 (ekrandaki 'Re-entry summary' satirlarinin ta kendisi)."""
    if not articles:
        return {}
    # ADAY FILTRESI — bozuk aday TUM sorguyu patlatir (hrk={} → tum teyit kaniti
    # kaybolur, teyitliler bekleyen gorunur!):
    #  - 21+ karakter → CWB0111 truncation (2026-07-21 'somun punta' vakasi)
    #  - ASCII disi (Turkce I/Ü/Ö...) → CWBNL0107 donusum hatasi (2026-07-23
    #    '150 MM-ALUMİNYUM' vakasi). Gercek MGARCD yalniz ASCII kod icerir →
    #    guvenli desene uymayanlar zaten eslesemez, sessizce ele.
    _guvenli = re.compile(r'^[A-Za-z0-9./\- ]{1,21}$')
    liste = sorted({str(a).strip() for a in articles
                    if str(a or '').strip() and _guvenli.match(str(a).strip())})
    if not liste:
        return {}
    pw = CFG.sifre_al()
    cn = pyodbc.connect(CFG.baglanti_dizesi(pw), timeout=60, autocommit=True)
    sonuc = collections.defaultdict(list)
    try:
        for i in range(0, len(liste), 60):
            grup = liste[i:i + 60]
            sql = CFG.HAREKET_SORGU.format(yer=','.join('?' * len(grup)))
            for r in cn.cursor().execute(sql, grup):
                try:
                    tarih = '%02d%02d-%02d-%02d' % (int(r[1]), int(r[2]), int(r[3]), int(r[4]))
                except (TypeError, ValueError):
                    continue
                tur = str(r[8]).strip() if len(r) > 8 else 'RPR'
                sonuc[kanonik(r[0])].append({
                    'tarih': tarih,
                    'adet': float(r[5] or 0),
                    'launch': f'{str(r[6]).strip()}-{str(r[7]).strip()}' if tur == 'RPR' else 'CFI',
                    'tur': tur,
                })
    finally:
        cn.close()
    for v in sonuc.values():
        v.sort(key=lambda x: x['tarih'])
    return sonuc


def _zaten_teyitli(hareketler, uretim_tarihi, adet, gecmis=None):
    """Bu launch'a, bu uretim gunu icin teyit verilmis gibi gorunuyor mu?
    Kural (2026-07-20): teyit normalde ERTESI gun girilir; hareket tarihi = GIRIS
    tarihi. Uretim gunuyle AYNI gunlu hareket ONCEKI gunun teyididir (1865W:
    07-17 uretimi 30, 07-17 hareketi 68 = 07-16'nin 68'lik uretiminin teyidi).
    - SONRAKI gunlerde adet birebir esleyen TEK hareket → 'kesin'
    - sonraki hareketlerin TOPLAMI birebir esliyorsa → 'kesin' (bolunmus teyit,
      orn 3680GAW 1+275=276)
    - AYNI-GUN istisnasi (kullanici 2026-07-20): ayni gunlu hareket bizim adetle
      BIREBIR esliyor VE onceki gunlerin uretim adetleriyle ACIKLANAMIYORSA
      (gecmis={tarih: {adetler}} bizim uretim kayitlarimizdan) → bizim teyidimiz,
      'kesin'. Aciklanabiliyorsa onceki gunundur, kanit sayilmaz.
    - GENEL ACIKLAMA KURALI: sonraki gunlerde girilen ama adedi BASKA bir uretim
      gununun adetiyle esleyen hareketler de o gunun teyididir → kanittan dusulur
      (orn pazartesi girilen 68 = cumanin 68'lik uretimi; bizim 30 hala teyitsiz).
    - kalan hareketlerin toplami >= bizim adet → 'olasi' (elle kontrol)
    Donis: (durum, ilgili_hareketler)  durum: None | 'kesin' | 'olasi'
    """
    gecmis = gecmis or {}

    def aciklanan(h):
        """Hareket BASKA bir gunun uretimiyle aciklaniyor mu? (kullanici 2026-07-20:
        'onceki gun uretimlerini kontrol ederek hangi uretimin teyidi verilmis
        anlayabiliriz'). Adet, bizim gun DISINDA bir uretim gununun adetiyle
        birebir esliyorsa (ve o gun hareketten once/ayni gunse) → o gunun teyidi."""
        for t, adl in gecmis.items():
            if t != uretim_tarihi and t <= h['tarih'] and any(abs(a - h['adet']) < 0.001 for a in adl):
                return True
        return False

    sonrakiler = [h for h in hareketler if h['tarih'] > uretim_tarihi]
    # 1) Sonraki gunlerde adet birebir esleyen hareket → bizim teyidimiz, kesin
    for h in sonrakiler:
        if abs(h['adet'] - adet) < 0.001:
            return 'kesin', [h]
    # 2) AYNI-GUN istisnasi: birebir adet + baska gunle aciklanamiyor → kesin
    for h in hareketler:
        if (h['tarih'] == uretim_tarihi and abs(h['adet'] - adet) < 0.001
                and not aciklanan(h)):
            return 'kesin', [h]
    # 3) Baska gunlerin uretimiyle aciklanan sonraki hareketler KANIT DEGIL
    #    (orn pazartesi girilen 68, cumanin 68'lik uretiminin teyidi olabilir).
    kalanlar = [h for h in sonrakiler if not aciklanan(h)]
    if not kalanlar:
        return None, []
    toplam = sum(h['adet'] for h in kalanlar)
    if abs(toplam - adet) < 0.001:
        return 'kesin', kalanlar        # bolunmus teyit (orn 1+275=276)
    if toplam >= adet - 0.001:
        return 'olasi', kalanlar
    return None, kalanlar


# REWORK: operator uretim kaydinin 'aciklama' alanina rework yazdiysa o KAYIT teyit
# DISI (kullanici 2026-07-27). Yazim degisken: rework/riwork/rowork/rivork... ->
# regex 'r' + 0-2 harf + work|vork (network/framework gibi r-oneksiz kelimeleri
# YAKALAMAZ). Kayit bazli: ayni referansin normal uretimi teyit edilir, rework
# kismi elenir. Bu yuzden SQL-SUM yerine kayitlari cekip Python'da gruplariz.
REWORK_RE = re.compile(r'r[a-z]{0,2}[wv]ork', re.I)


def rework_mi(*metinler):
    """Verilen metinlerden herhangi biri rework belirtiyor mu (aciklama VEYA referans)."""
    return any(m and REWORK_RE.search(m) for m in metinler)


def gun_uretimi(tarih):
    """O gunun referans+adet listesi — yalniz launch'li bolumler + TK2.
    Rework kayitlari (aciklama/referans 'rework' varyanti) haric tutulur."""
    conn = sqlite3.connect(URETIM_DB)
    conn.row_factory = sqlite3.Row
    rows = conn.execute(f"""
        SELECT COALESCE(v.lokasyon,'TK2') tesis, COALESCE(v.bolum,'kaynak') bolum,
               u.referans_kodu referans, COALESCE(u.ok_adet,0) ok,
               COALESCE(u.nok_adet,0) nok, COALESCE(u.aciklama,'') aciklama
        FROM uretim_kayitlari u JOIN vardiyalar v ON v.id = u.vardiya_id
        WHERE v.tarih = ? AND u.referans_kodu IS NOT NULL AND u.referans_kodu != ''
          AND COALESCE(v.lokasyon,'TK2') = ?
          AND COALESCE(v.bolum,'kaynak') IN ({','.join('?'*len(LAUNCH_BOLUMLERI))})
        ORDER BY tesis, bolum, u.referans_kodu""",
        (tarih, LOKASYON) + LAUNCH_BOLUMLERI).fetchall()
    conn.close()
    # Grupla — rework kayitlarini ATLA (kayit bazli; ayni referansin normal uretimi kalir)
    grup = {}
    for r in rows:
        if rework_mi(r['aciklama'], r['referans']):
            continue
        key = (r['tesis'], r['bolum'], r['referans'])
        g = grup.get(key)
        if g is None:
            g = grup[key] = {'tesis': r['tesis'], 'bolum': r['bolum'],
                             'referans': r['referans'], 'adet': 0, 'hurda': 0}
        g['adet']  += r['ok']
        g['hurda'] += r['nok']
    # adet>0 filtre + bosluklu kod yazimini nokta bicimine cevir ('10 300 1393A' →
    # '10.300.1393A'); orijinal yazim UI'da gosterilmek uzere saklanir (2026-07-22).
    sonuc = []
    for d in grup.values():
        if d['adet'] <= 0:
            continue
        duz = bosluk_nokta(d['referans'])
        if duz != d['referans']:
            d['orijinal_referans'] = d['referans']
            d['referans'] = duz
        sonuc.append(d)
    return sonuc


def uretim_gecmisi(bas_tarih, son_tarih):
    """Kendi uretim kayitlarimizdan tarih bazli adet gecmisi (2026-07-20).
    Ayni-gun hareketlerin 'onceki gunun teyidi mi' sorusunu cevaplamak icin:
    hareket adedi onceki bir gunun uretim adetiyle birebir esliyorsa o gunundur.
    Donis: {kanonik(referans): {tarih: {adetler}}} (raw referans_kodu bazinda).
    REWORK kayitlari HARIC (gun_uretimi ile tutarli — rework teyit edilmez, hareket
    aciklamasi olarak da sayilmamali)."""
    conn = sqlite3.connect(URETIM_DB)
    conn.row_factory = sqlite3.Row
    ara = collections.defaultdict(float)   # (tarih, ref_raw) -> toplam ok (rework haric)
    try:
        for r in conn.execute(f"""
            SELECT v.tarih tarih, u.referans_kodu ref, COALESCE(u.ok_adet,0) ok,
                   COALESCE(u.aciklama,'') aciklama
            FROM uretim_kayitlari u JOIN vardiyalar v ON v.id = u.vardiya_id
            WHERE v.tarih >= ? AND v.tarih <= ?
              AND u.referans_kodu IS NOT NULL AND u.referans_kodu != ''
              AND COALESCE(v.lokasyon,'TK2') = ?
              AND COALESCE(v.bolum,'kaynak') IN ({','.join('?'*len(LAUNCH_BOLUMLERI))})""",
                (bas_tarih, son_tarih, LOKASYON) + LAUNCH_BOLUMLERI).fetchall():
            if rework_mi(r['aciklama'], r['ref']):
                continue
            ara[(r['tarih'], r['ref'])] += r['ok']
    finally:
        conn.close()
    g = collections.defaultdict(lambda: collections.defaultdict(set))
    for (t, ref), toplam in ara.items():
        if toplam > 0:
            g[kanonik(bosluk_nokta(ref))][t].add(float(toplam))
    return g


def ref_uretim_gecmisi(referans, uretim_tarihi, gun=10):
    """Tek referansin uretim gecmisi (gonderim tarafi dedup icin) — {tarih: {adetler}}."""
    bas = (date.fromisoformat(uretim_tarihi) - timedelta(days=gun)).isoformat()
    return uretim_gecmisi(bas, uretim_tarihi).get(kanonik(referans), {})


def kalici_haric_set():
    """Kalici 'teyit gerekmez' isaretli referanslarin kanonik seti (2026-07-20).
    Kullanici teyit ekranindan bir referansi 'gerek_yok' isaretlerse (orn 6343a ara
    urun) o referans her gun teyit disi (HARIC) kalir. Tablo yoksa bos set."""
    conn = sqlite3.connect(URETIM_DB)
    s = set()
    try:
        for (rk,) in conn.execute(
                "SELECT referans FROM as400_teyit_isaret WHERE kapsam='kalici' AND durum='gerek_yok'").fetchall():
            s.add(kanonik(rk))
    except Exception:
        pass
    finally:
        conn.close()
    return s


def esle(tarih):
    """Tek gun icin teyit is listesi (geriye uyum)."""
    return esle_coklu([tarih])[tarih]


def esle_coklu(tarihler):
    """Birden cok uretim gunu icin is listesi — AS400 okumalari TEK SEFER yapilir
    (acik emirler + tum gunlerin article'lari icin hareketler), sonra gun gun
    kategorize edilir. Sabah kuyrugu gorunumunun motoru.
    Donis: {tarih: kategori_sozlugu}"""
    tam, gev, kokm = as400_launchlar()
    hazir = {}
    articles = set()
    for t in tarihler:
        satirlar, ara = op_kurali_uygula(gun_uretimi(t))
        hazir[t] = (satirlar, ara)
        for u in satirlar:
            ad = (tam.get(kanonik(u['referans'])) or gev.get(gevsek(u['referans']))
                  or kokm.get(kok(u['referans'])))
            for l in (ad or []):
                articles.add(l['article'])
            # Uretim referansinin KENDISI de DAIMA hareket adayi (2026-07-21 4082W
            # dersi: launch S ile kapaninca satir kok-eslesmeyle CIPLAK 10.300.4082'nin
            # OPR'sine baglandi; hareket sorgusu yalniz o article'a bakinca 4082W'ye
            # 07:28'de verilen teyidi GOREMEDI → ayni uretim CFI ile MUKERRER gitti.
            # DB2 CHAR bosluk-dolgusunu tolere eder; eslesmezse sorgu bos doner.)
            ref = str(u['referans'] or '').strip()
            if ref:
                articles.add(ref)
                articles.add(ref.upper())
    try:
        hrk = teyit_hareketleri(articles)
    except Exception as _e:
        # SESSIZ KALMA (2026-07-21 dersi: DataError tum kaniti yutup 4082W'nin
        # mukerrer CFI gonderimine yol acti) — en azindan konsola yaz.
        print(f'[launch_esle] hareket sorgusu HATASI (kanitsiz devam): {_e}')
        hrk = {}
    haric = kalici_haric_set()
    # Uretim gecmisi (ayni-gun hareket aciklamasi icin): en eski sorgu gununden
    # 10 gun oncesine kadar — TEK sqlite sorgusu, tum gunler icin ortak.
    try:
        bas = (date.fromisoformat(min(tarihler)) - timedelta(days=10)).isoformat()
        gecmis_map = uretim_gecmisi(bas, max(tarihler))
    except Exception:
        gecmis_map = {}
    return {t: _kategorize(t, hazir[t][0], hazir[t][1], tam, gev, kokm, hrk, haric, gecmis_map)
            for t in tarihler}


def _kategorize(tarih, satirlar, ara_oplar, tam, gev, kokm, hrk, haric_set=None, gecmis_map=None):
    """Bir gunun satirlarini kategorize eder (onceden yuklenmis AS400 verisiyle)."""
    # OPR10 = OPR acilmis ama launch ALINMAMIS (durum 10) — planlamaya sinyal
    # ARAOP = ara operasyon (son op degil) — teyit son op uretiminde verilir
    # HARIC = kullanicinin kalici 'teyit gerekmez' isaretledigi referanslar
    haric_set = haric_set or set()
    gecmis_map = gecmis_map or {}
    sonuc = {'ACIK': [], 'SUPHELI': [], 'OPR10': [], 'KAPALI': [], 'YOK': [], 'ARAOP': [], 'HARIC': []}
    sonuc['ARAOP'] = [{**r, 'launchlar': []} for r in ara_oplar]
    for u in satirlar:
        if kanonik(u['referans']) in haric_set:
            sonuc['HARIC'].append({**u, 'launchlar': []})
            continue
        adaylar = tam.get(kanonik(u['referans']))
        supheli = False
        if not adaylar:
            adaylar = gev.get(gevsek(u['referans'])) or kokm.get(kok(u['referans']))
            supheli = adaylar is not None
        if not adaylar:
            sonuc['YOK'].append({**u, 'launchlar': []})
            continue
        # KOPYA sart: ayni launch kaydi (as400_launchlar'daki dict) birden cok
        # uretim satirina eslesebilir (orn. 1847A ve 1847CS → ayni launch 17809).
        # Kopyalamazsak, sonraki satirin mukerrer-hesabi oncekinin uzerine yazar.
        acik = [dict(a) for a in adaylar if a['durum'] == 40]
        if acik:
            sonuc['SUPHELI' if supheli else 'ACIK'].append({**u, 'launchlar': acik})
            continue
        # SUPHELI (kok/gevsek) eslesmede ACIK launch yoksa aday guvenilir DEGIL:
        # farkli urunun OPR'si olabilir (2026-07-21: 4082W, kapali launch sonrasi
        # CIPLAK 10.300.4082'nin OPR'sine baglandi → CFI yanlis urune gitti).
        # Bu durumda YOK say → CFI hedefi uretim referansinin kendisi olur.
        if supheli:
            sonuc['YOK'].append({**u, 'launchlar': []})
            continue
        opr = [dict(a) for a in adaylar if a['durum'] == 10]
        if opr:
            sonuc['OPR10'].append({**u, 'launchlar': opr})
        else:
            sonuc['KAPALI'].append({**u, 'launchlar': [dict(a) for a in adaylar]})

    # ── MUKERRER KORUMASI (ARTICLE bazli) + ZOMBI launch isareti ──
    # Operatorun ELLE girdigi teyitler de gorunur; ayrica operator BASKA bir
    # launch'a girip kapatmis olabilir → article bazli bakiyoruz (hrk hazir gelir).
    for kat in ('ACIK', 'SUPHELI'):
        for r in sonuc[kat]:
            for l in r['launchlar']:
                # ZOMBI: durum 40 ama launch adedi asilmis (kalan <= 0) →
                # fiilen bitmis, S ile kapatilmamis. Teyit BURAYA verilmemeli.
                l['zombi'] = (l['kalan'] <= 0.001)
                hareketler = hrk.get(kanonik(l['article']), [])
                durum, ilgili = _zaten_teyitli(hareketler, tarih, r['adet'],
                                               gecmis_map.get(kanonik(r['referans'])))
                l['zaten_teyitli'] = durum          # None | 'kesin' | 'olasi'
                l['zaten_hareket'] = ilgili[:5]     # yalniz ilgili olanlar (JSON sismesin)

    # ── OPR10/KAPALI/YOK icin de SATIR duzeyinde hareket kontrolu (2026-07-20) ──
    # Sabah S ile kapatilan launch XPRO90'dan duser; ayni referansin YENI siparisi
    # durum 10 gorununce satir yanlislikla "launch alinmamis"e dusuyordu (94.CV.037,
    # 94.LTK.460/701/766 vakasi). Once RPR hareketine bak: uretim gunu+sonrasinda
    # adet birebir teyit varsa → satiri teyitli say (frontend Teyitli'ye tasir).
    for kat in ('OPR10', 'KAPALI', 'YOK'):
        for r in sonuc[kat]:
            anahtar = kanonik(r['referans'])
            hareketler = hrk.get(anahtar, [])
            if not hareketler and r['launchlar']:
                hareketler = hrk.get(kanonik(r['launchlar'][0]['article']), [])
            durum, ilgili = _zaten_teyitli(hareketler, tarih, r['adet'], gecmis_map.get(anahtar))
            r['zaten_teyitli'] = durum              # None | 'kesin' | 'olasi'
            r['zaten_hareket'] = ilgili[:5]
    return sonuc


if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    tarih = sys.argv[1] if len(sys.argv) > 1 else (date.today() - timedelta(days=1)).isoformat()
    print(f'Teyit is listesi — uretim gunu: {tarih}\n')
    s = esle(tarih)
    for kat, baslik in [('ACIK', 'ACIK LAUNCH VAR (durum 40) — teyit verilebilir'),
                        ('SUPHELI', 'ACIK ama yazim/varyant farki — kontrol et'),
                        ('OPR10', 'OPR VAR ama LAUNCH ALINMAMIS (durum 10) — planlamaya haber'),
                        ('KAPALI', 'yalniz durum 45/50 kaydi var (TK1 akisi?)'),
                        ('YOK', 'acik emirlerde HIC yok (kapali ya da OPR yok)'),
                        ('ARAOP', 'ARA OPERASYON — teyit son op uretiminde verilir')]:
        print(f'── {baslik}: {len(s[kat])}')
        for r in s[kat][:15]:
            ls = ', '.join(f"{l['launch']}(d{l['durum']},kln{l['kalan']:g})" for l in r['launchlar'][:3])
            print(f"   {r['tesis']:3} {r['bolum']:7} {r['referans']:26} adet={r['adet']:<6} {ls}")
        if len(s[kat]) > 15:
            print(f'   ... +{len(s[kat])-15} satir daha')
        print()
