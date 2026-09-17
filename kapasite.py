# -*- coding: utf-8 -*-
"""KAPASİTE MODÜLÜ — ürün havuzu, bölüm sınıflandırması ve süreler (kullanıcı 2026-09-17).

Kullanıcı: "TK1 ve TK2 kapasite hesabı modülü oluşturalım. Önce TK1 ve TK2'deki bütün
referansları AS400'den çekip P olan kodları ayıralım. P olanlar üretim, A olanlar fason
kodları demek. Sonrasında P olan referansları bölümlere göre ayıracağız… yeni ve tanımlı
olmayan referanslar için atama yapabilelim… süre tanımlamaları da gerekecek."

KAYNAKLAR (AS400 / TKC0301F):
  BARTF0  Anagrafico articoli — A0ARTI kod, A0ARDS açıklama, A0PROV proveniens
          (P = üretim, A = satın alma/fason, C/F birkaç istisna), A0ARAN='A' iptal
  BSPEF1  Spec.prod. Dett.risorsa — ürünün ROTASI: hangi kaynakta (risorsa) kaç
          saniye (Y1IVUM='SS'). ERP'de 27.384 aktif P kodun rotası TANIMLI.
  BRISF0  Anagrafico risorse + B350F0 Tabella reparto — kaynak → ERP bölümü
          (WELDING→Welding, LASERCUT→Laser cutting, CNCBEND→CNC BENDING…)

TASARIM KARARLARI
  · Süre ERP'den GELİR ama BAĞLAYICI DEĞİLDİR: bizim ölçtüğümüz süre (referans_listesi)
    her zaman kazanır. ERP süresi "öneri" olarak gösterilir, tek tıkla uygulanır.
    (Örnek 10.130.4456W: biz 102.8 sn ölçtük, ERP 120 sn diyor.)
  · Atama referans_listesi'ne YAZILIR — OEE, analiz ve kapasite aynı tanımı kullansın.
    Otomatik yazılmaz: 25 binden fazla kod var, operatör listelerini boğardı.
  · AS400 tarafında TK1/TK2 ayrımı YOK (tek şirket). Lokasyon bizde belirlenir:
    kod referans_listesi'nde varsa oradan, yoksa atamada kullanıcı seçer.
"""
import sqlite3
from datetime import datetime

# Kaynak (ERP risorsa) → bizim bölüm. Sınıflandırmanın ÇEKİRDEĞİ; panelden
# değiştirilebilir (kapasite_kaynak tablosu), burası yalnız İLK kurulum değeridir.
# '' = bölüme sayılmaz (dış işlem / malzeme / test).
VARSAYILAN_KAYNAK_BOLUM = {
    'WELDING': 'kaynak',
    'LASERCUT': 'lazer',
    'CNCBEND': 'bukum',            # ⚠ 'bukum' bölümü sistemde YOK — bkz. BOLUM_DISI
    'OUTCASE': 'tel',
    'MEC.20T': 'pres', 'MEC.40T': 'pres', 'MEC.60T': 'pres', 'MEC.80T': 'pres',
    'MEC.200T': 'pres',
    'INJ.270T': 'plastik', 'INJ.320T': 'plastik',
    'INJ.50T': 'metal', 'INJ.100T': 'metal', 'INJ.200T': 'metal',
    'INJ.300T': 'metal', 'INJ.400T': 'metal',
    'DIECASTALL': 'metal', 'DIECASTZAM': 'metal',
    'BURRINGALL': 'metal', 'BURRINGZAM': 'metal', 'WRKDIECAST': 'metal',
    'BRO.4T': 'isleme', 'TURNING': 'isleme', 'DRILLING': 'isleme',
    'THREAD': 'isleme', 'THREAT': 'isleme', 'SCREWCUT': 'isleme',
    'SCREWCUTH': 'isleme', 'HOLE': 'isleme', 'HOLEHARD': 'isleme', 'CUTAA': 'isleme',
    # Genel işçilik kaynakları: ERP'de montaj/el işçiliği bunlarla yazılıyor.
    # NROPE 25 binden fazla kodda var ve KAYNAK/LAZER kodlarında da ikinci satır
    # olarak geçiyor → tek başına "montaj" demek değil. Bu yüzden türetmede
    # "yalnız işçilik varsa montaj" kuralı uygulanır (bkz. rota_bolumleri).
    'NROPE': 'montaj', 'NROPEAA': 'montaj', 'OPERAIO': 'montaj',
    'SPECOPEPP': 'montaj', 'T0006': 'montaj', 'WRKGENERIC': 'montaj',
    # Dış işlem / kaplama / test — kendi bölümümüz değil, kapasiteye girmez
    'T': '', 'TER': '', 'TEP': '', 'TEA': '', 'RAD': '', 'VACUUM': '',
    'PAINTING': '', 'T0001': '', 'T0002': '', 'T0003': '', 'T0004': '', 'T0005': '',
    'CATCOATING': '', 'EPOKSI': '', 'ZAMAPRINT': '', 'PLASTICPR': '',
    'ELECT TEST': '', 'EOL TEST': '', 'LEAK TEST': '', 'SMD': '', 'SOLDERING': '',
    'HOT PLT WD': '', 'T1 HEAT TR': '', 'MATERIALE': '', 'MPRE': '', 'MPRO': '',
}

# Yalnız işçilik sayılan kaynaklar: bir kodun rotasında BUNLARDAN BAŞKA bir şey
# yoksa iş montajdır; makine kaynağı da varsa işçilik o makinenin yanında sayılır.
ISCILIK_KAYNAKLARI = ('NROPE', 'NROPEAA', 'OPERAIO', 'SPECOPEPP', 'WRKGENERIC', 'T0006')

# Sistemde üretim bölümü olarak TANIMLI OLMAYAN kapasite bölümleri. Atama yapılamaz
# (referans_listesi bu bölümü kabul etmez); panelde "bölüm sistemde yok" uyarısı çıkar.
BOLUM_DISI = ('bukum',)

# Zaman birimleri (ERP Y1IVUM) → saniye çarpanı. AD/KG/NR malzeme satırıdır, süre değil.
UM_SANIYE = {'SS': 1.0, 'MN': 60.0, 'HH': 3600.0}

AS400_SEMA = 'TKC0301F'

SQL_URUN = f"""
    SELECT A0ARTI, A0ARDS, A0PROV, A0ARAN, A0TART, A0LNPR
    FROM {AS400_SEMA}.BARTF0
"""
SQL_ROTA = f"""
    SELECT r.Y1ARTI, r.Y1LIAG, r.Y1FAPR, r.Y1RIPR, r.Y1RICD,
           r.Y1IVUM, r.Y1IVQT, r.Y1IFUM, r.Y1IFQT, r.Y1RINM
    FROM {AS400_SEMA}.BSPEF1 r
    JOIN {AS400_SEMA}.BARTF0 a ON a.A0ARTI = r.Y1ARTI
    WHERE a.A0PROV = 'P' AND (a.A0ARAN IS NULL OR a.A0ARAN <> 'A')
"""
SQL_KAYNAK = f"""
    SELECT r.ARRICD, r.ARRIDS, r.ARRPCD, d.B35003, r.ARRIAN
    FROM {AS400_SEMA}.BRISF0 r
    LEFT JOIN {AS400_SEMA}.B350F0 d ON d.B35002 = r.ARRPCD
"""


def _t(v):
    return '' if v is None else str(v).strip()


def _f(v):
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def senkron(conn, kullanici='', baglan=None):
    """AS400'den ürün havuzunu + rotaları çeker, yerel tablolara yazar.
    Dönüş: {'urun', 'uretim', 'fason', 'rota', 'kaynak', 'yeni_kaynak': [...]}
    AS400 kapalıysa/yetki yoksa exception yükseltir — çağıran mesajı gösterir."""
    if baglan is None:
        import sys, os
        sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'as400'))
        import as400_config as CFG
        baglan = CFG.baglan
    cn = baglan(timeout=60)
    simdi = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    try:
        cur = cn.cursor()
        # 1) Ürünler — A ve P birlikte alınır: "bu kod fason mu" sorusu da cevaplansın
        cur.execute(SQL_URUN)
        urunler = [(_t(r[0]), _t(r[1]), _t(r[2]).upper(), 1 if _t(r[3]).upper() == 'A' else 0,
                    _t(r[4]), _t(r[5]), simdi) for r in cur.fetchall() if _t(r[0])]
        # 2) Rotalar (yalnız aktif P kodlar)
        cur.execute(SQL_ROTA)
        rotalar = []
        for r in cur.fetchall():
            um, mik = _t(r[5]).upper(), _f(r[6])
            h_um, h_mik = _t(r[7]).upper(), _f(r[8])
            rotalar.append((_t(r[0]), _t(r[1]), int(_f(r[2])), int(_f(r[3])), _t(r[4]), um, mik,
                            mik * UM_SANIYE.get(um, 0.0), h_mik * UM_SANIYE.get(h_um, 0.0),
                            max(1, int(_f(r[9])) or 1)))
        # 3) Kaynak (risorsa) listesi — bölüm haritasının satırları
        cur.execute(SQL_KAYNAK)
        kaynaklar = [(_t(r[0]), _t(r[1]), _t(r[2]), _t(r[3]), 1 if _t(r[4]).upper() == 'A' else 0)
                     for r in cur.fetchall() if _t(r[0])]
    finally:
        try:
            cn.close()
        except Exception:
            pass

    c = conn.cursor()
    # Ürün havuzu: UPSERT + bu senkronda görülmeyenleri 'guncel=0' yap (ERP'den
    # düşen kod sessizce kaybolmasın, listede "artık ERP'de yok" diye görünsün).
    c.execute("UPDATE kapasite_urun SET guncel = 0")
    c.executemany(
        "INSERT INTO kapasite_urun (kod, aciklama, prov, iptal, tip, urun_hatti, senk_at, guncel) "
        "VALUES (?,?,?,?,?,?,?,1) ON CONFLICT(kod) DO UPDATE SET aciklama=excluded.aciklama, "
        "prov=excluded.prov, iptal=excluded.iptal, tip=excluded.tip, "
        "urun_hatti=excluded.urun_hatti, senk_at=excluded.senk_at, guncel=1", urunler)
    # Rotalar tam yenilenir: ERP'de silinen adım bizde kalmasın.
    # INSERT OR IGNORE: ERP'de aynı satır iki kez yazılmış olabiliyor (2026-09-17'de
    # 34.485 satırda 1 tane: 93.00.1856 · NROPE 145,6 sn iki kez). Toplarsak o kodun
    # süresi iki katı çıkardı; mükerrer satır ELENİR ve sayısı senkron sonucunda döner.
    c.execute("DELETE FROM kapasite_rota")
    onceki = c.execute("SELECT COUNT(*) FROM kapasite_rota").fetchone()[0]
    c.executemany(
        "INSERT OR IGNORE INTO kapasite_rota (kod, surum, faz, sira, kaynak_kod, um, miktar, "
        "sure_sn, hazirlik_sn, kisi) VALUES (?,?,?,?,?,?,?,?,?,?)", rotalar)
    yazilan = c.execute("SELECT COUNT(*) FROM kapasite_rota").fetchone()[0] - onceki
    mukerrer = len(rotalar) - yazilan
    # Kaynak haritası: YENİ kaynak eklenirse varsayılan bölümüyle gelir, kullanıcının
    # elle verdiği bölüm EZİLMEZ (elle=1 satırlara dokunulmaz).
    yeni = []
    for kod, ad, rep, rep_ad, iptal in kaynaklar:
        var = c.execute("SELECT bolum, elle FROM kapasite_kaynak WHERE kaynak_kod=?", (kod,)).fetchone()
        if var is None:
            varsayilan = VARSAYILAN_KAYNAK_BOLUM.get(kod.upper(), '')
            c.execute("INSERT INTO kapasite_kaynak (kaynak_kod, aciklama, reparto, reparto_ad, "
                      "bolum, iptal, elle, updated_at) VALUES (?,?,?,?,?,?,0,?)",
                      (kod, ad, rep, rep_ad, varsayilan, iptal, simdi))
            # Yalnız HİÇ TANIMADIĞIMIZ kaynak bildirilir. Bilerek '' verilenler
            # (dış işlem/kaplama/test) "atanmamış" değildir — uyarı gürültüsü olurdu.
            if kod.upper() not in VARSAYILAN_KAYNAK_BOLUM:
                yeni.append(kod)          # bölümü bilinmeyen YENİ kaynak → kullanıcı atasın
        else:
            c.execute("UPDATE kapasite_kaynak SET aciklama=?, reparto=?, reparto_ad=?, iptal=? "
                      "WHERE kaynak_kod=?", (ad, rep, rep_ad, iptal, kod))
    uretim = sum(1 for u in urunler if u[2] == 'P' and not u[3])
    fason = sum(1 for u in urunler if u[2] == 'A' and not u[3])
    c.execute("INSERT INTO kapasite_senk (tarih, kullanici, urun, uretim, fason, rota, kaynak) "
              "VALUES (?,?,?,?,?,?,?)",
              (simdi, kullanici, len(urunler), uretim, fason, yazilan, len(kaynaklar)))
    conn.commit()
    return {'urun': len(urunler), 'uretim': uretim, 'fason': fason, 'rota': yazilan,
            'mukerrer_rota': mukerrer, 'kaynak': len(kaynaklar), 'yeni_kaynak': yeni,
            'tarih': simdi}


def kaynak_haritasi(conn):
    """{KAYNAK_KOD: (bolum, aciklama)} — panelden düzenlenen eşleme."""
    return {str(r['kaynak_kod']).strip().upper(): (r['bolum'] or '', r['aciklama'] or '')
            for r in conn.execute("SELECT kaynak_kod, bolum, aciklama FROM kapasite_kaynak")}


def rota_bolumleri(satirlar, harita):
    """Bir ürünün rota satırlarından {bolum: saniye} çıkarır.

    · Sürüm (LIAG): temel sürüm ('') varsa O kullanılır; yoksa tek/ilk sürüm
      (ERP'de 'PROD'/'TEL' sürümleri temelin kopyası — ikisini toplamak süreyi
      iki katına çıkarırdı).
    · İŞÇİLİK KURALI: rotada makine kaynağı varsa işçilik satırı (NROPE vb.) o
      koda ayrı bir 'montaj' işi AÇMAZ — süresi en uzun makine bölümüne eklenir.
      Yalnız işçilik varsa iş gerçekten manuel montajdır.
    """
    surumler = {str(s.get('surum') or '').strip() for s in satirlar}
    sec = '' if '' in surumler else (sorted(surumler)[0] if surumler else '')
    satirlar = [s for s in satirlar if str(s.get('surum') or '').strip() == sec]
    makine, iscilik = {}, 0.0
    for s in satirlar:
        kod = str(s.get('kaynak_kod') or '').strip().upper()
        sn = float(s.get('sure_sn') or 0)
        if sn <= 0:
            continue
        bolum = (harita.get(kod, ('', ''))[0] or '').strip()
        if not bolum:
            continue                                   # dış işlem / malzeme
        if kod in ISCILIK_KAYNAKLARI:
            iscilik += sn
        else:
            makine[bolum] = makine.get(bolum, 0.0) + sn
    if makine:
        if iscilik:
            enb = max(makine, key=lambda b: makine[b])
            makine[enb] += iscilik
        return {b: round(v, 2) for b, v in makine.items()}
    return {'montaj': round(iscilik, 2)} if iscilik else {}


def turet(conn):
    """Rota satırlarından (kod, bolum, sure_sn) türetip kapasite_urun_bolum'e yazar.

    Senkrondan SONRA ve kaynak→bölüm eşlemesi her değiştiğinde çalışır. Türetmeyi
    tabloya yazmanın sebebi listeleme: 29 binden fazla kodu her istekte Python'da
    süzmek yerine SQL süzer (sayfalama da ancak böyle doğru çalışır)."""
    harita = kaynak_haritasi(conn)
    ham = {}
    for r in conn.execute("SELECT kod, surum, kaynak_kod, sure_sn FROM kapasite_rota"):
        ham.setdefault(r['kod'], []).append(dict(r))
    satirlar = []
    for kod, rota in ham.items():
        for bolum, sn in rota_bolumleri(rota, harita).items():
            if sn > 0:
                satirlar.append((kod, bolum, sn))
    c = conn.cursor()
    c.execute("DELETE FROM kapasite_urun_bolum")
    c.executemany("INSERT INTO kapasite_urun_bolum (kod, bolum, sure_sn) VALUES (?,?,?)", satirlar)
    conn.commit()
    return len(satirlar)


def _norm(kod):
    """Kod eşlemesi: büyük harf + boşluksuz. NOKTA ANLAMLIDIR (10.300.1 ≠ 10.3001)."""
    return str(kod or '').strip().upper().replace(' ', '')


def _tanimli_sql(lokasyon, sureli=False):
    """Bizde TANIMLI kodların normalize listesi (alt sorgu metni).

    HIZ NOTU (2026-09-17): bu eşleme önce korele alt sorguydu
    (`EXISTS (… WHERE UPPER(REPLACE(r.referans_kodu,' ',''))=u.kod)`) ve SQLite her
    ürün için referans_listesi'ni baştan tarıyordu → 29 bin kodda **76 saniye**.
    IN (SELECT …) biçiminde liste BİR KEZ kurulup bloom filtresine giriyor → 0,03 sn.
    Biçimi korele hâle geri çevirme.
    """
    q = ("SELECT UPPER(REPLACE(referans_kodu,' ','')) FROM referans_listesi "
         "WHERE referans_kodu IS NOT NULL")
    if lokasyon:
        q += " AND COALESCE(lokasyon,'TK2') = ?"
    if sureli:
        q += " AND COALESCE(hedef_cycle_time_sn,0) > 0"
    return q


def ozet(conn, lokasyon=''):
    """Panel KPI'ları: havuzun ne kadarı sınıflanmış / süresi tanımlı."""
    par = [lokasyon] if lokasyon else []
    s = conn.execute(
        "SELECT COUNT(*) toplam, SUM(CASE WHEN prov='P' AND iptal=0 THEN 1 ELSE 0 END) uretim, "
        "SUM(CASE WHEN prov='A' AND iptal=0 THEN 1 ELSE 0 END) fason, "
        "SUM(CASE WHEN iptal=1 THEN 1 ELSE 0 END) iptal, "
        "SUM(CASE WHEN guncel=0 THEN 1 ELSE 0 END) dusmus FROM kapasite_urun").fetchone()
    rotali = conn.execute(
        "SELECT COUNT(*) n FROM kapasite_urun u WHERE u.prov='P' AND u.iptal=0 "
        "AND EXISTS (SELECT 1 FROM kapasite_urun_bolum b WHERE b.kod = u.kod)").fetchone()['n']
    tanimli = conn.execute(
        "SELECT COUNT(*) n FROM kapasite_urun u WHERE u.prov='P' AND u.iptal=0 "
        f"AND u.kod IN ({_tanimli_sql(lokasyon)})", par).fetchone()['n']
    sureli = conn.execute(
        "SELECT COUNT(*) n FROM kapasite_urun u WHERE u.prov='P' AND u.iptal=0 "
        f"AND u.kod IN ({_tanimli_sql(lokasyon, True)})", par).fetchone()['n']
    son = conn.execute("SELECT * FROM kapasite_senk ORDER BY id DESC LIMIT 1").fetchone()
    return {
        'toplam': s['toplam'] or 0, 'uretim': s['uretim'] or 0, 'fason': s['fason'] or 0,
        'iptal': s['iptal'] or 0, 'dusmus': s['dusmus'] or 0, 'rotali': rotali,
        'tanimli': tanimli, 'sureli': sureli,
        'atanmamis': max(0, (s['uretim'] or 0) - tanimli), 'suresiz': max(0, tanimli - sureli),
        'lokasyon': lokasyon,
        'son_senk': ({'tarih': son['tarih'], 'kullanici': son['kullanici'], 'urun': son['urun'],
                      'rota': son['rota'], 'uretim': son['uretim'], 'fason': son['fason']}
                     if son else None),
    }


def liste(conn, prov='P', durum='', bolum='', lokasyon='', ara='', sirala='kod',
          limit=100, offset=0):
    """Ürün havuzu — her satırda ERP rotası (bölüm + saniye) ve bizdeki tanım(lar).

    durum: '' hepsi · 'atanmamis' (bizde tanım yok) · 'suresiz' (tanım var, süre yok)
           · 'tanimli' (süresi var) · 'rotasiz' (ERP'de rota yok → süre önerisi çıkmaz)
    bolum: ERP rotasından TÜRETİLEN bölüm (bizim atadığımız bölüm değil).
    """
    kosul, par = ["u.iptal = 0"], []
    if prov:
        kosul.append("u.prov = ?")
        par.append(prov)
    if ara:
        a = '%' + str(ara).strip().upper() + '%'
        kosul.append("(u.kod LIKE ? OR UPPER(u.aciklama) LIKE ?)")
        par += [a, a]
    if bolum:
        kosul.append("EXISTS (SELECT 1 FROM kapasite_urun_bolum b WHERE b.kod=u.kod AND b.bolum=?)")
        par.append(bolum)
    var = f"u.kod IN ({_tanimli_sql(lokasyon)})"
    sureli = f"u.kod IN ({_tanimli_sql(lokasyon, True)})"
    if durum == 'atanmamis':
        kosul.append(f"u.kod NOT IN ({_tanimli_sql(lokasyon)})")
        par += ([lokasyon] if lokasyon else [])
    elif durum == 'suresiz':
        kosul.append(var + f" AND u.kod NOT IN ({_tanimli_sql(lokasyon, True)})")
        par += ([lokasyon, lokasyon] if lokasyon else [])
    elif durum == 'tanimli':
        kosul.append(sureli)
        par += ([lokasyon] if lokasyon else [])
    elif durum == 'rotasiz':
        kosul.append("NOT EXISTS (SELECT 1 FROM kapasite_urun_bolum b WHERE b.kod = u.kod)")
    nere = " WHERE " + " AND ".join(kosul)
    toplam = conn.execute("SELECT COUNT(*) n FROM kapasite_urun u" + nere, par).fetchone()['n']
    duzen = ("ORDER BY (SELECT COALESCE(SUM(sure_sn),0) FROM kapasite_urun_bolum b "
             "WHERE b.kod=u.kod) DESC, u.kod" if sirala == 'sure' else "ORDER BY u.kod")
    satirlar = [dict(r) for r in conn.execute(
        "SELECT u.kod, u.aciklama, u.prov, u.tip, u.urun_hatti, u.guncel FROM kapasite_urun u"
        + nere + ' ' + duzen + " LIMIT ? OFFSET ?", par + [int(limit), int(offset)])]
    kodlar = [s['kod'] for s in satirlar]
    if kodlar:
        isaret = ','.join('?' * len(kodlar))
        rota = {}
        for r in conn.execute(
                f"SELECT kod, bolum, sure_sn FROM kapasite_urun_bolum WHERE kod IN ({isaret})",
                kodlar):
            rota.setdefault(r['kod'], {})[r['bolum']] = round(float(r['sure_sn'] or 0), 1)
        tanim = {}
        for r in conn.execute(
                "SELECT referans_kodu, COALESCE(bolum,'kaynak') bolum, COALESCE(lokasyon,'TK2') lokasyon, "
                "COALESCE(hedef_cycle_time_sn,0) ct FROM referans_listesi WHERE "
                f"UPPER(REPLACE(referans_kodu,' ','')) IN ({isaret})", kodlar):
            tanim.setdefault(_norm(r['referans_kodu']), []).append(
                {'bolum': r['bolum'], 'lokasyon': r['lokasyon'], 'ct': round(float(r['ct'] or 0), 1)})
        for s in satirlar:
            s['rota'] = rota.get(s['kod'], {})
            hepsi = tanim.get(s['kod'], [])
            s['tanimlar'] = [t for t in hepsi if not lokasyon or t['lokasyon'] == lokasyon]
            s['diger_tesis'] = [t for t in hepsi if lokasyon and t['lokasyon'] != lokasyon]
            s['durum'] = ('tanimli' if any(t['ct'] > 0 for t in s['tanimlar'])
                          else ('suresiz' if s['tanimlar'] else 'atanmamis'))
    return {'toplam': toplam, 'satirlar': satirlar, 'limit': int(limit), 'offset': int(offset)}
