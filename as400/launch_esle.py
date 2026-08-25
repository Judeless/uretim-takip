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

# ── TEYIT KAPSAMI: (tesis, bolumler) — kullanici 2026-08-25 ──────────────────
# TK1 PLASTIK EKLENDI. O bolumde launch ALINMIYOR, dolayisiyla satirlar dogal
# olarak 'YOK' kategorisine (= CFI kuyrugu) duser; kullanici da "cfi ile teyit
# veririz" dedi. Diger TK1 bolumleri (montaj/tel) HALA DISARIDA.
TEYIT_KAPSAMI = (
    ('TK2', LAUNCH_BOLUMLERI),
    ('TK1', ('plastik',)),
)

# app.PLASTIK_MAKINELERI ile BIREBIR olmali — sira = uretim_kayitlari.istasyon.
# app.py bu listeyi cagirmadan once dogrular (_le_kapsam_dogrula); tutmuyorsa
# konsola UYARI basar, cunku sira kaymasi yanlis makineyi teyit disi birakir.
PLASTIK_MAKINELERI = ['320T', '407T', 'Yapistirma', 'Sizdirmazlik Test']

# TEYIT VERILMEYEN MAKINELER (kullanici 2026-08-25):
# "yapistirma yapilan urun ile test makinesinde test edilen urunun kodu ayni
#  oldugu icin iki makinede de uretilene teyit vermeye gerek yok — test
#  cihazinda uretilen urunlere teyit vermeyelim."
# Satir kuyruktan SILINMEZ, 'HARIC' kategorisine duser ve gerekce yazilir:
# uretimin kayda gectigi gorunsun, mukerrer teyit olmasin.
TEYIT_DISI_MAKINE = {('TK1', 'plastik', 'Sizdirmazlik Test')}


def _kapsam_where(alias='v'):
    """TEYIT_KAPSAMI'ni SQL kosuluna cevirir: (tesis,bolum) ciftlerinin OR'u."""
    parcalar, parm = [], []
    for tesis, bolumler in TEYIT_KAPSAMI:
        yer = ','.join('?' * len(bolumler))
        parcalar.append(f"(COALESCE({alias}.lokasyon,'TK2') = ? "
                        f"AND COALESCE({alias}.bolum,'kaynak') IN ({yer}))")
        parm.append(tesis)
        parm.extend(bolumler)
    return '(' + ' OR '.join(parcalar) + ')', parm


def makine_adi(tesis, bolum, istasyon):
    """Uretim kaydinin makinesi (plastik/TK1). Bilinmiyorsa ''."""
    if bolum == 'plastik':
        try:
            i = int(istasyon or 0)
        except (TypeError, ValueError):
            return ''
        if 1 <= i <= len(PLASTIK_MAKINELERI):
            return PLASTIK_MAKINELERI[i - 1]
    return ''


def article_tanimli(kodlar):
    """ERP article master'inda (BARTF0.A0ARTI) TANIMLI olan kodlarin kumesi.

    NEDEN (2026-07-30): operator uretim kaydina '10.300.1992A.OP.1' yazdi —
    kod alanina OPERASYON bilgisi karistirmis ('1. OP' yerine 'OP.1'). Robot bu
    kodu 07>01>F1 ekranina yazmaya calisinca alan tasti, ekran 'article alani
    dolu' hatasinda kilitlendi ve AYNI KOSUDAKI DIGER REFERANSLARIN teyidi de
    verilemedi. Tek bozuk kod butun kuyrugu durduruyor.

    Uzunluk kriteri ISE YARAMAZ: ERP'de 17 karakterlik MESRU kodlar var
    ('92.10.4312      B' gibi bosluk dolgulu). Dogru kriter kodun master'da
    TANIMLI olup olmadigi. Donis: {kanonik(kod)} kumesi.
    Sorgu patlarsa BOS KUME degil None doner — cagiran 'bilinmiyor' ile
    'tanimsiz'i ayirt etsin (bilinmiyorken gonderimi bloklamak yanlis yon)."""
    if not kodlar:
        return set()
    # Beyaz liste DEGIL, gercek riskler (2026-07-30 duzeltmesi): sorguyu
    # patlatan tek sey 21+ karakter (CWB0111 truncation) ve ASCII disi karakter
    # (CWBNL0107 donusum). Karakter beyaz listesi ('.', '/', '-' disini ele)
    # 318 GERCEK article'i yanlislikla eliyordu — icinde 10.300.1992A_LK gibi
    # alt cizgili kodlar da vardi ve "ERP'de yok" sanilip teyitleri atlaniyordu.
    # Her ozel karakter (_ , & ( ) $ * " \x1a) gercek ERP sorgusuyla tek tek VE
    # toplu test edildi: parametreli sorgu oldugu icin hicbiri patlatmiyor.
    _guvenli = re.compile(r'^[\x00-\x7F]{1,21}$')
    # HARF DUYARLILIGI (2026-08-17 duzeltmesi): DB2 IN karsilastirmasi harf
    # DUYARLIDIR, ama donen kume kanonik(=UPPER) ile kuruluyor ve cagiran da
    # kanonik(article) ile bakiyor. Yani sorgu duyarli, karsilastirma duyarsiz —
    # tutarsizlik. Sonuc: operator kucuk harfle yazdiginda ('10.130.6202b')
    # ERP'de kod BUYUK harfle dururken sorgu HICBIR SEY bulmuyor ve mesru kod
    # "ERP'de boyle bir article YOK" diye reddediliyordu.
    # Cozum: hem yazildigi gibi hem BUYUK harfli bicimi IN listesine koy. Kolona
    # UPPER(A0ARTI) uygulamak indeksi devre disi birakirdi — bu yol esitlik
    # karsilastirmasini (ve indeksi) korur.
    ham = {str(k).strip() for k in kodlar
           if str(k or '').strip() and _guvenli.match(str(k).strip())}
    liste = sorted(ham | {k.upper() for k in ham})
    if not liste:
        return set()
    try:
        cn = pyodbc.connect(CFG.baglanti_dizesi(CFG.sifre_al()), timeout=60, autocommit=True)
    except Exception as e:
        print(f'[launch_esle] article master baglanti HATASI: {e}')
        return None
    var = set()
    try:
        cu = cn.cursor()
        for i in range(0, len(liste), 60):
            grup = liste[i:i + 60]
            sql = ("SELECT A0ARTI FROM tkc0301F.BARTF0 WHERE A0ARTI IN (%s)"
                   % ','.join('?' * len(grup)))
            for r in cu.execute(sql, grup):
                var.add(kanonik(r[0]))
    except Exception as e:
        print(f'[launch_esle] article master sorgu HATASI: {e}')
        return None
    finally:
        cn.close()
    return var


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
    # Beyaz liste DEGIL, gercek riskler (2026-07-30 duzeltmesi): sorguyu
    # patlatan tek sey 21+ karakter (CWB0111 truncation) ve ASCII disi karakter
    # (CWBNL0107 donusum). Karakter beyaz listesi ('.', '/', '-' disini ele)
    # 318 GERCEK article'i yanlislikla eliyordu — icinde 10.300.1992A_LK gibi
    # alt cizgili kodlar da vardi ve "ERP'de yok" sanilip teyitleri atlaniyordu.
    # Her ozel karakter (_ , & ( ) $ * " \x1a) gercek ERP sorgusuyla tek tek VE
    # toplu test edildi: parametreli sorgu oldugu icin hicbiri patlatmiyor.
    _guvenli = re.compile(r'^[\x00-\x7F]{1,21}$')
    # HARF DUYARLILIGI (2026-08-17) — BURADA EN TEHLIKELISI: sorgu MGARCD'yi HARF
    # DUYARLI karsilastirir, donen sozluk ise kanonik(=UPPER) anahtarli. Operator
    # kucuk harfle yazdiginda ('10.130.6206b') ERP'de kod BUYUK dururken hareket
    # sorgusu HICBIR SEY bulmuyor → o urunun MEVCUT TEYITLERI GORUNMEZ oluyor →
    # mukerrer freni calismiyor ve ayni uretim IKINCI KEZ teyit edilebiliyor.
    # Cozum: kodun hem yazildigi hem BUYUK harfli bicimini sorgula (esitlik ve
    # indeks korunur; kolona UPPER() uygulamak indeksi devre disi birakirdi).
    ham = {str(a).strip() for a in articles
           if str(a or '').strip() and _guvenli.match(str(a).strip())}
    liste = sorted(ham | {a.upper() for a in ham})
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
    # ADET YOKSA TEYIT KANITI DA YOK (2026-07-31): 'sadece hurda' satirlarinda
    # (ok_adet=0, nok_adet>0) adet 0'dir. Asagidaki 'toplam >= adet' kurali 0 ile
    # DAIMA saglanir ve satir sahte 'olasi' isaretlenip Dikkat kovasini sisirirdi.
    try:
        if float(adet or 0) <= 0:
            return None, []
    except (TypeError, ValueError):
        return None, []

    def aciklanan(h):
        """Hareket BASKA bir gunun uretimiyle aciklaniyor mu? (kullanici 2026-07-20:
        'onceki gun uretimlerini kontrol ederek hangi uretimin teyidi verilmis
        anlayabiliriz'). Adet, bizim gun DISINDA bir uretim gununun adetiyle
        birebir esliyorsa (ve o gun hareketten once/ayni gunse) → o gunun teyidi.

        BIR GUN BIR KEZ KARSILANIR (2026-07-30 duzeltmesi): o gecmis gunun adedini
        karsilayan DAHA ERKEN bir hareket zaten varsa o gun KAPANMISTIR; eldeki
        hareket ona ait olamaz, baska (muhtemelen BIZIM) uretime aittir.
        Gercek olay 94.LTK.487/10: 24.07 uretimi 30 adet -> 07-27'de RPR 30 ile
        teyit edilmis. 29.07 uretimi de 30 adet; 07-29'daki RPR 30 kanit olmasi
        gerekirken '24.07'nin teyidi' sayilip yok sayildi, satir teyitsiz gorundu
        ve 30.07'de IKINCI KEZ teyit gitti -> ERP'ye 30 adet FAZLA."""
        return id(h) in _aciklanan_kume

    # ── AÇIKLANAN HAREKETLER — TEK GECISTE, GUN TUKETEREK (2026-08-21) ──────
    # ESKI UYGULAMA: her hareket icin ayri ayri "bu adedi karsilayan DAHA ERKEN
    # TARIHLI bir hareket var mi" diye bakiliyordu. Ayni GUN girilen hareketler
    # icin bu kontrol HIC calismiyordu (x['tarih'] < h['tarih'] hicbir zaman
    # dogru olmuyor) -> ayni gunlu hareketlerin HEPSI ayni gecmis gune atfedilip
    # eleniyordu.
    # GERCEK OLAY 94.LTK.215 (2026-08-21): 19.08 ve 20.08'de 100'er adet uretim,
    # 20.08'de 3x100 hareket. Ucu de "19.08'in teyidi" sayildi -> 20.08 satiri
    # teyitsiz gorundu -> tekrar gonderildi -> ERP'ye 100 adet FAZLA girdi.
    # YENI: hareketler tarih sirasiyla gezilir, karsilanan gecmis gun TUKETILIR.
    # Bir gecmis gun yalnizca BIR harekete sahip olabilir; kalanlar kanit sayilir.
    _aciklanan_kume = set()
    _tuketilen = set()
    for h in sorted(hareketler, key=lambda x: x['tarih']):
        for t in sorted(gecmis):
            if t == uretim_tarihi or t in _tuketilen or t > h['tarih']:
                continue
            if any(abs(a - h['adet']) < 0.001 for a in gecmis[t]):
                _aciklanan_kume.add(id(h))
                _tuketilen.add(t)
                break

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
    # 3) BOLUNMUS TEYIT — AYNI GUN + SONRAKI GUNLER BIRLIKTE (2026-08-06 duzeltmesi).
    #    Eskiden yalniz 'sonrakiler' toplanirdi; ayni gun verilen PARCALI teyit
    #    hic sayilmiyordu. Gercek olay 94.LTK.890: 06.08 uretimi 80 adet, teyit AYNI
    #    GUN iki parca halinde verilmis (RPR 23 + RPR 57 = 80). Iki hareket de tek
    #    basina 80'e esit olmadigi icin kural 2'ye takilmiyor, kural 3 ise ayni gunu
    #    hic gormuyordu → satir "teyitsiz" kalip kuyrukta tekrar tekrar cikiyordu.
    #    Baska gunun uretimiyle aciklanan hareketler yine ELENIR (aciklanan): ayni
    #    gunlu bir hareket onceki gunun teyidi olabilir, o koruma bozulmuyor.
    kalanlar = [h for h in hareketler if h['tarih'] >= uretim_tarihi and not aciklanan(h)]
    if not kalanlar:
        return None, []
    toplam = sum(h['adet'] for h in kalanlar)
    if abs(toplam - adet) < 0.001:
        return 'kesin', kalanlar        # bolunmus teyit (orn 1+275=276)
    if toplam >= adet - 0.001:
        return 'olasi', kalanlar
    # 4) KISMI TEYIT (2026-08-21) — gunun bir PARCASI teyit edilmis.
    #    ESKIDEN None donuyordu, yani satir "hic teyit verilmemis" gorunuyor ve
    #    panel GUNUN TAMAMINI oneriyordu.
    #    GERCEK OLAY 94.LTK.09: 20.08'de iki kalem uretim (18 + 17 = 35), ertesi
    #    gun yalniz 18 teyit edilmis. Satir 35 adetle "teyitsiz" cikti; gonderilse
    #    ERP'ye 18 adet FAZLA girecekti.
    #    Cagiran KALAN adedi kanit hareketlerinin toplamindan cikarir:
    #        kalan = adet - sum(h['adet'] for h in ilgili)
    #    'kismi' OTOMATIK GONDERILMEZ — 'olasi' gibi elle onay ister; kanit
    #    hareketi bize ait olmayabilir, o zaman kalan yanlis hesaplanirdi.
    if toplam > 0.001:
        return 'kismi', kalanlar
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


# ─────────────────────────────────────────────────────────────────────────────
# KAPASITE MAKULLUK KONTROLU (2026-07-30)
# ─────────────────────────────────────────────────────────────────────────────
# OLAY: operator 10.130.3778 (metal, cycle 65 sn) icin 410 yerine 4410 yazdi.
# 4410 x 65 sn = 79.6 SAAT — tek vardiyada fiziken imkansiz. Teyit verilseydi
# ERP'ye ~4000 adetlik hayali stok girecekti (kullanici adedi elle duzeltti).
#
# KURAL: referansin cycle time'i TANIMLI ise, o gun o referansi ureten
# vardiyalarin toplam suresinden teorik azami adet hesaplanir; girilen adet
# bunun TOLERANS katini asiyorsa satir "kapasite asimi" olarak isaretlenir ve
# robot ona DOKUNMAZ (Dikkat kovasi + oto kosuda atlanir + gonderimde reddedilir).
# Cycle time TANIMSIZ ise (0/NULL) kontrol YAPILMAZ — kullanici boyle istedi.
#
# TOLERANS neden 2.0: hedef cycle muhafazakar konur, operator hedefin uzerine
# cikabilir (%50-100 fazla uretim gorulur). 2 kat pay gercek uretimi ASLA
# bloklamaz ama 10 kat parmak hatasini (410 -> 4410) kesin yakalar.
KAPASITE_TOLERANS = 2.0
# Vardiya suresi hic bilinemezse fiziksel gun siniri: tek makine gunde en fazla
# 24 saat doner. Hic kontrol etmemektense bu tavan kullanilir.
KAPASITE_VARSAYILAN_DK = 24 * 60


def _vardiya_suresi_dk(sure_dk, bas_saat, bit_saat):
    """Vardiyanin CALISMA suresi (dk). Once toplam_sure_dk; yoksa baslangic-bitis
    saatlerinden hesaplanir (acik vardiyada bitis bos olabilir -> 0 doner ve
    cagiran varsayilan tavana duser)."""
    try:
        s = float(sure_dk or 0)
    except (TypeError, ValueError):
        s = 0.0
    if s > 0:
        return s
    try:
        b = str(bas_saat or '').strip()[:5]
        e = str(bit_saat or '').strip()[:5]
        if len(b) == 5 and len(e) == 5:
            bh, bm = int(b[:2]), int(b[3:5])
            eh, em = int(e[:2]), int(e[3:5])
            dk = (eh * 60 + em) - (bh * 60 + bm)
            if dk < 0:          # gece vardiyasi (23:00 -> 07:00)
                dk += 24 * 60
            if dk > 0:
                return float(dk)
    except (TypeError, ValueError):
        pass
    return 0.0


def ayni_isi_birlestir(satirlar):
    """AYNI ISIN farkli satirlarini TEK teyit satirinda toplar (2026-07-30).

    NEDEN: gun_uretimi HAM YAZIMLA gruplar, op_kurali_uygula ise SONRA son-op
    satirinin referansini BASE koda cevirir. Bu sirada ayni base koda inen
    satirlar AYRI kaliyordu:
      '10.300.4415W 2. op' (30 adet) -> base '10.300.4415W'
      '10.300.4415W'       ( 5 adet) -> zaten base
    Sonuc: panelde ayni referanstan IKI satir; biri teyit edilip digeri
    unutuluyordu (kullanici 2026-07-30: "birden fazla operator ayni isi
    yaptiysa sadece 1 ini degil, adeti toplayip tek seferde teyit verelim").

    Anahtar kanonik base kod: buyuk/kucuk harf ve bosluk farklari da yutulur
    ('10.300.6175a' + '10.300.6175A' tek satir olur).
    adet ve hurda TOPLANIR; vardiya sureleri birlesir (kapasite hesabi
    birlesik adet + birlesik sure uzerinden yapilmali, yoksa toplanan adet
    tek vardiyanin suresine bolunup yanlis 'asim' verir)."""
    grup = {}
    for r in satirlar:
        # 'teyit_disi' ANAHTARA DAHIL (2026-08-25): op-kurali base koda cevirdikten
        # sonra teyit disi satir (test cihazi) teyit verilecek satirla ayni base
        # koda dusuyor ve BIRLESIYORDU -> hem HARIC ayrimi kayboluyor hem adet
        # toplanip ERP'ye iki kat stok giriyordu.
        anahtar = (r.get('tesis'), r.get('bolum'), kanonik(r.get('referans')),
                   bool(r.get('teyit_disi')))
        g = grup.get(anahtar)
        if g is None:
            grup[anahtar] = dict(r)
            continue
        g['adet'] += r.get('adet', 0)
        g['hurda'] = g.get('hurda', 0) + r.get('hurda', 0)
        for _m in (r.get('makineler') or []):
            if _m not in g.setdefault('makineler', []):
                g['makineler'].append(_m)
        # vardiya sureleri: id bazli tekille (ayni vardiya iki satirda gecebilir)
        gv = g.setdefault('_vardiyalar', {})
        gv.update(r.get('_vardiyalar') or {})
        # cycle time: ilk dolu deger
        if not g.get('ct') and r.get('ct'):
            g['ct'] = r['ct']
        # hangi yazimlarin birlestigi UI'da gorunsun (denetim izi)
        yazimlar = g.setdefault('birlesen_yazimlar',
                                [g.get('orijinal_referans') or g.get('referans')])
        yazimlar.append(r.get('orijinal_referans') or r.get('referans'))
    return list(grup.values())


def _kapasite_isle(d):
    """Grup sozlugune kapasite alanlarini yazar; asim varsa d['kapasite_asim'].
    _vardiyalar yardimci alani JSON'a sismesin diye temizlenir."""
    sureler = d.pop('_vardiyalar', {}) or {}
    sure_dk = sum(v for v in sureler.values() if v and v > 0)
    ct = d.get('ct') or 0
    d['kapasite_sure_dk'] = round(sure_dk, 1)
    if not ct or ct <= 0:
        d['ct'] = None            # tanimsiz -> kontrol yok (UI de boyle anlar)
        return
    # Sure hic bilinmiyorsa fiziksel gun tavani
    etkin_dk = sure_dk if sure_dk > 0 else KAPASITE_VARSAYILAN_DK
    teorik = (etkin_dk * 60.0) / float(ct)
    azami = teorik * KAPASITE_TOLERANS
    d['kapasite_teorik'] = int(teorik)
    d['kapasite_azami'] = int(azami)
    if d['adet'] > azami:
        gerek_sa = (d['adet'] * float(ct)) / 3600.0
        d['kapasite_asim'] = {
            'adet': d['adet'], 'ct': float(ct),
            'teorik': int(teorik), 'azami': int(azami),
            'sure_dk': round(etkin_dk, 1),
            'sure_varsayilan': sure_dk <= 0,
            'gerekli_saat': round(gerek_sa, 1),
            'kat': round(d['adet'] / teorik, 1) if teorik > 0 else None,
        }


def son_uretim_gunleri(adet=3, bugun_haric=True, azami_geri_gun=45):
    """Teyit kuyrugunun bakacagi son 'adet' URETIM GUNU (takvim gunu DEGIL).

    NEDEN (2026-08-17): kuyruk "bugun - 1..3" takvim gunune bakiyordu. Uretim
    olmayan gunler pencereyi yiyor:
      - Her PAZARTESI pencere Pzt/Paz/Cmt/Cum -> pratikte yalniz CUMA gorunuyor,
        persembenin teyit edilmemis isi sessizce pencereden dusuyordu.
      - 2026-08-07'de bir HAFTALIK ara verildi; donuste 07.08'in teyit edilmemis
        kodlari 'Yenile' ile HIC gorunmedi (10 gun geride, ustelik gunler
        parametresi 7 ile sinirli) — yalniz elle tarih filtresiyle bulunabildi.
    Uretim yapilan gunleri secince ara/hafta sonu/tatil pencereyi kaydirmaz.

    Kapsam gun_uretimi ile AYNI (TEYIT_KAPSAMI): kapsam disi bir tesiste/teyit
    disi bir bolumde uretim yapilan gun kuyruga girmez, aksi halde bos gun icin
    bosuna AS400 sorgusu atilirdi.

    azami_geri_gun: bu kadar gunden eski gunler alinmaz — sistem uzun sure bos
    kalirsa kuyruk aylar oncesini acmasin (o kayitlara elle tarih filtresiyle
    bakilir). Sorgu basarisiz olursa BOS liste doner; cagiran takvim gunune duser.
    """
    try:
        bugun = date.today()
        alt_sinir = (bugun - timedelta(days=azami_geri_gun)).isoformat()
        ust_sinir = (bugun - timedelta(days=1)).isoformat() if bugun_haric else bugun.isoformat()
        _kap_sql2, _kap_parm2 = _kapsam_where('v')
        conn = sqlite3.connect(URETIM_DB)
        try:
            rows = conn.execute(f"""
                SELECT DISTINCT v.tarih
                FROM uretim_kayitlari u JOIN vardiyalar v ON v.id = u.vardiya_id
                WHERE v.tarih BETWEEN ? AND ?
                  AND COALESCE(u.ok_adet,0) > 0
                  AND u.referans_kodu IS NOT NULL AND u.referans_kodu != ''
                  AND {_kap_sql2}
                ORDER BY v.tarih DESC LIMIT ?""",
                [alt_sinir, ust_sinir] + _kap_parm2 + [int(adet)]).fetchall()
        finally:
            conn.close()
        return [r[0] for r in rows]
    except Exception as e:
        print(f'[launch_esle] son_uretim_gunleri basarisiz, takvim gunune dusuluyor: {e}')
        return []


def gun_uretimi(tarih):
    """O gunun referans+adet listesi — TEYIT_KAPSAMI (TK2 launch bolumleri + TK1 plastik).
    Rework kayitlari (aciklama/referans 'rework' varyanti) haric tutulur."""
    conn = sqlite3.connect(URETIM_DB)
    conn.row_factory = sqlite3.Row
    # KAPASITE ALANLARI (2026-07-30): cycle time + vardiya suresi de cekilir ki
    # "8 saatte fiziken uretilemeyecek adet" teyit oncesi yakalanabilsin.
    # LEFT JOIN: referans tanimli DEGILSE ct NULL kalir -> kontrol YAPILMAZ.
    # ⚠ CYCLE TIME SKALAR ALT-SORGUYLA ALINIR, **LEFT JOIN İLE DEĞİL** (2026-07-30).
    # LEFT JOIN kullanıldığında (ilk yazım) referans_listesi'nde normalize olarak
    # AYNI koda düşen BİRDEN FAZLA satır varsa üretim kaydı O KADAR KEZ dönüyordu
    # ve g['adet'] += ok döngüsü adedi KATLIYORDU. Gerçek olay: '94.LTK.340' ve
    # '94.ltk.340' iki ayrı satır → 42 adetlik üretim 84 olarak teyit edildi ve
    # ERP'ye 42 adet FAZLA stok girdi. Montaj/TK2'de böyle 29 çift yazım var.
    # UNIQUE(referans_kodu,bolum,lokasyon) bu duplikeyi ENGELLEMEZ: kısıt HAM kodda,
    # eşleşme ise normalize (UPPER+REPLACE) yapılıyor.
    # Skalar alt-sorgu satır ÇOĞALTMAZ; tam-eş yazım öncelikli, sonra dolu ct.
    _kap_sql, _kap_parm = _kapsam_where('v')
    rows = conn.execute(f"""
        SELECT COALESCE(v.lokasyon,'TK2') tesis, COALESCE(v.bolum,'kaynak') bolum,
               u.referans_kodu referans, COALESCE(u.ok_adet,0) ok,
               COALESCE(u.nok_adet,0) nok, COALESCE(u.aciklama,'') aciklama,
               COALESCE(u.istasyon,0) istasyon,
               (SELECT MAX(rl.hedef_cycle_time_sn) FROM referans_listesi rl
                 WHERE UPPER(REPLACE(rl.referans_kodu,' ','')) = UPPER(REPLACE(u.referans_kodu,' ',''))
                   AND COALESCE(rl.bolum,'kaynak') = COALESCE(v.bolum,'kaynak')
                   AND COALESCE(rl.lokasyon,'TK2') = COALESCE(v.lokasyon,'TK2')) ct,
               -- AS400 DEPO KODLARI (TK1 plastik): her urunun deposu farkli.
               -- Skalar alt-sorgu (LEFT JOIN DEGIL): ayni koda normalize dusen
               -- birden fazla satir uretim kaydini COGALTIR, adet katlanir.
               (SELECT MAX(rl2.depo_kodu) FROM referans_listesi rl2
                 WHERE UPPER(REPLACE(rl2.referans_kodu,' ','')) = UPPER(REPLACE(u.referans_kodu,' ',''))
                   AND COALESCE(rl2.bolum,'kaynak') = COALESCE(v.bolum,'kaynak')
                   AND COALESCE(rl2.lokasyon,'TK2') = COALESCE(v.lokasyon,'TK2')) depo_kodu,
               (SELECT MAX(rl3.karsi_depo_kodu) FROM referans_listesi rl3
                 WHERE UPPER(REPLACE(rl3.referans_kodu,' ','')) = UPPER(REPLACE(u.referans_kodu,' ',''))
                   AND COALESCE(rl3.bolum,'kaynak') = COALESCE(v.bolum,'kaynak')
                   AND COALESCE(rl3.lokasyon,'TK2') = COALESCE(v.lokasyon,'TK2')) karsi_depo_kodu,
               v.id vardiya_id, COALESCE(v.toplam_sure_dk,0) sure_dk,
               COALESCE(v.baslangic_saati,'') bas_saat, COALESCE(v.bitis_saati,'') bit_saat
        FROM uretim_kayitlari u JOIN vardiyalar v ON v.id = u.vardiya_id
        WHERE v.tarih = ? AND u.referans_kodu IS NOT NULL AND u.referans_kodu != ''
          AND {_kap_sql}
        ORDER BY tesis, bolum, u.referans_kodu""",
        [tarih] + _kap_parm).fetchall()
    conn.close()
    # Grupla — rework kayitlarini ATLA (kayit bazli; ayni referansin normal uretimi kalir)
    grup = {}
    for r in rows:
        if rework_mi(r['aciklama'], r['referans']):
            continue
        # TEYIT DISI BAYRAGI ANAHTARA GIRER (kullanici 2026-08-25).
        # Yapistirma ile sizdirmazlik testinin kodu AYNI: ayni satirda birleserlerse
        # (a) test kaydi yapistirmanin icinde kaybolur, 'teyit gerekmez' ayrimi
        # yapilamaz, (b) adet TOPLANIR ve ERP'ye IKI KAT stok girer.
        # Anahtar MAKINE degil BAYRAK: 320T + 407T + Yapistirma ayni kodu
        # uretiyorsa TEK satirda toplanmalari DOGRU (hepsine teyit verilir);
        # ayrilmasi gereken yalniz teyit disi makinedir.
        _mak = makine_adi(r['tesis'], r['bolum'], r['istasyon'])
        _disi = (r['tesis'], r['bolum'], _mak) in TEYIT_DISI_MAKINE
        key = (r['tesis'], r['bolum'], r['referans'], _disi)
        g = grup.get(key)
        if g is None:
            g = grup[key] = {'tesis': r['tesis'], 'bolum': r['bolum'],
                             'referans': r['referans'], 'adet': 0, 'hurda': 0,
                             'ct': None, '_vardiyalar': {}, 'teyit_disi': _disi,
                             'makineler': [],
                             'depo_kodu': (r['depo_kodu'] or '').strip(),
                             'karsi_depo_kodu': (r['karsi_depo_kodu'] or '').strip()}
        if _mak and _mak not in g['makineler']:
            g['makineler'].append(_mak)
        g['adet']  += r['ok']
        g['hurda'] += r['nok']
        # Cycle time: ilk dolu deger (ayni referans+bolum tek satir olmali)
        try:
            _ct = float(r['ct']) if r['ct'] is not None else 0.0
        except (TypeError, ValueError):
            _ct = 0.0
        if _ct > 0 and not g['ct']:
            g['ct'] = _ct
        # Vardiya suresi: AYNI vardiya birden cok uretim kaydinda gecer ->
        # id'ye gore tekille, yoksa sure kat kat sisip kontrolu ise yaramaz kilar.
        g['_vardiyalar'][r['vardiya_id']] = _vardiya_suresi_dk(
            r['sure_dk'], r['bas_saat'], r['bit_saat'])
    # adet>0 filtre + bosluklu kod yazimini nokta bicimine cevir ('10 300 1393A' →
    # '10.300.1393A'); orijinal yazim UI'da gosterilmek uzere saklanir (2026-07-22).
    sonuc = []
    for d in grup.values():
        # SADECE HURDA (2026-07-31): ok_adet=0 ama nok_adet>0 olan gun/referans
        # satiri ESKIDEN buradan tamamen dusuyordu -> AS400 teyit sayfasinda HIC
        # gorunmuyor, hurdasi ♻ COP kuyruguna da girmiyordu. "Hepsi hurda cikti"
        # gunu sessizce kayboluyor ve COP asla girilemiyordu.
        # Satir artik KALIR; adet=0 oldugu icin teyit kuyruklarina giremez
        # (launch/CFI yollarinda adet>0 kapisi var), yalniz COP listesinde cikar.
        if d['adet'] <= 0 and (d.get('hurda') or 0) <= 0:
            continue
        duz = bosluk_nokta(d['referans'])
        if duz != d['referans']:
            d['orijinal_referans'] = d['referans']
            d['referans'] = duz
        # KAPASITE burada HESAPLANMAZ: op kurali sonrasi ayni base koda inen
        # satirlar birlestirilecek (ayni_isi_birlestir) ve adet/sure DEGISECEK.
        # Hesap birlesmeden once yapilirsa toplanan adet tek satirin suresine
        # bolunur ve yanlis 'kapasite asimi' uretir. Cagiran sirayi kurar.
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
    _kap_sql3, _kap_parm3 = _kapsam_where('v')
    try:
        for r in conn.execute(f"""
            SELECT v.tarih tarih, u.referans_kodu ref, COALESCE(u.ok_adet,0) ok,
                   COALESCE(u.aciklama,'') aciklama
            FROM uretim_kayitlari u JOIN vardiyalar v ON v.id = u.vardiya_id
            WHERE v.tarih >= ? AND v.tarih <= ?
              AND u.referans_kodu IS NOT NULL AND u.referans_kodu != ''
              AND {_kap_sql3}""",
                [bas_tarih, son_tarih] + _kap_parm3).fetchall():
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


def bizim_gonderilenler(tarihler):
    """(tarih, kanonik referans) -> BIZIM basariyla gonderdigimiz toplam adet.

    NEDEN (kullanici 2026-08-25, 10.300.2756W olayi): parcali uretimde gunun
    uretimi birden cok kayitta olusuyor (16+10+41+31 = 98) ve teyit de parca
    parca veriliyor. Panel gunun TAMAMINI oneriyordu; 41 zaten gonderilmisken
    98 daha gondermek ERP'ye 41 adet FAZLA stok yazardi. Artik kalan onerilir.

    COP (yil='CO') SAYILMAZ: hurda ayri bir depo hareketidir, uretim teyidi degil.
    Basarisiz denemeler de sayilmaz — ERP'ye yazilmadilar."""
    if not tarihler:
        return {}
    conn = sqlite3.connect(URETIM_DB)
    out = collections.defaultdict(float)
    try:
        yer = ','.join('?' * len(tarihler))
        for r in conn.execute(
                f"SELECT uretim_tarihi, referans, adet FROM as400_teyit_log "
                f"WHERE uretim_tarihi IN ({yer}) AND sonuc='ok' AND yil != 'CO'",
                list(tarihler)).fetchall():
            try:
                out[(r[0], kanonik(r[1]))] += float(r[2] or 0)
            except (TypeError, ValueError):
                continue
    except Exception as e:
        # Tablo yok / okunamadi -> bos don. Kalan hesabi yapilmaz, panel eski
        # davranisa (gunun tamami) duser; mukerrer freni yine yerinde.
        print(f'[launch_esle] bizim_gonderilenler okunamadi: {e}')
        return {}
    finally:
        conn.close()
    return dict(out)


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
        # Op kurali base koda cevirdi -> ayni ise dusen satirlari TEK satira topla,
        # sonra kapasiteyi BIRLESIK adet+sure uzerinden hesapla (sira kritik).
        satirlar = ayni_isi_birlestir(satirlar)
        for _s in satirlar:
            _kapasite_isle(_s)
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
    # BIZIM daha once gonderdigimiz adet (bolunmus teyit) — satira yazilir ki
    # panel gunun TAMAMINI degil KALANI onersin.
    try:
        gonderilmis = bizim_gonderilenler(tarihler)
    except Exception as _e:
        print(f'[launch_esle] gonderilmis hesabi atlandi: {_e}')
        gonderilmis = {}
    for t in tarihler:
        for u in hazir[t][0]:
            g = gonderilmis.get((t, kanonik(u.get('referans'))), 0.0)
            if g > 0:
                u['gonderilmis_adet'] = round(g, 3)
                u['kalan_gonderilecek'] = max(0, int(round((u.get('adet') or 0) - g)))
    return {t: _kategorize(t, hazir[t][0], hazir[t][1], tam, gev, kokm, hrk, haric, gecmis_map)
            for t in tarihler}


def _teyit_isle(hedef, durum, ilgili, adet):
    """_zaten_teyitli ciktisini satira/launch'a yazar.

    'kismi' (2026-08-21) icin KALAN adedi de hesaplar: panel gunun TAMAMINI
    degil kalani onermeli. 94.LTK.09'da gunun 35 adedinin 18'i teyitliydi;
    kalan hesaplanmadigi icin panel 35 oneriyordu ve gonderilse ERP'ye 18 adet
    FAZLA girecekti."""
    hedef['zaten_teyitli'] = durum          # None | 'kesin' | 'olasi' | 'kismi'
    hedef['zaten_hareket'] = ilgili[:5]     # yalniz ilgili olanlar (JSON sismesin)
    if durum == 'kismi':
        teyitli = sum(h['adet'] for h in ilgili)
        hedef['teyitli_adet'] = round(teyitli, 3)
        hedef['kalan_adet'] = round(max(0.0, (adet or 0) - teyitli), 3)
    return hedef


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
        # TEYIT DISI MAKINE (kullanici 2026-08-25): sizdirmazlik test cihazinda
        # uretilen urunun kodu yapistirmadakiyle AYNI — ikisine de teyit verilirse
        # ERP'ye iki kat stok girer. Satir listede KALIR (uretim gorunsun), fakat
        # HARIC olur ve gerekcesi yazilir.
        if u.get('teyit_disi'):
            sonuc['HARIC'].append({**u, 'launchlar': [],
                                   'haric_sebep': 'Test cihazı — teyit yapıştırmada verilir '
                                                  '(aynı kod, mükerrer olmasın)'})
            continue
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
                _teyit_isle(l, durum, ilgili, r['adet'])
            # SATIR duzeyi kontrol — KENDI referans kodunun hareketleri (2026-07-30).
            # ACIK/SUPHELI'de yalnizca launch'in ARTICLE'ina bakiliyordu; varyant
            # eslesmede (10.300.6175B -> launch article 10.300.6175W) teyit KENDI
            # koduna CFI olarak verilirse 6175W hareketlerinde GORUNMUYOR ve satir
            # sonsuza kadar "Dikkat gerektirir"de kaliyordu (kullanici 2026-07-30:
            # "6175B icin teyit vermisiz neden dikkatte cikmis").
            # Veri hazirdi: esle_coklu articles setine uretim referansini DA ekliyor
            # (2026-07-21 4082W dersi) — yalnizca burada kullanilmiyordu.
            _ref_hrk = hrk.get(kanonik(r['referans']), [])
            if _ref_hrk:
                _d, _i = _zaten_teyitli(_ref_hrk, tarih, r['adet'],
                                        gecmis_map.get(kanonik(r['referans'])))
                _teyit_isle(r, _d, _i, r['adet'])

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
            _teyit_isle(r, durum, ilgili, r['adet'])
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
