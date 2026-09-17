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
from datetime import date, datetime, timedelta

# Kaynak (ERP risorsa) → bizim bölüm. Sınıflandırmanın ÇEKİRDEĞİ; panelden
# değiştirilebilir (kapasite_kaynak tablosu), burası yalnız İLK kurulum değeridir.
# '' = bölüme sayılmaz (dış işlem / malzeme / test).
VARSAYILAN_KAYNAK_BOLUM = {
    'WELDING': 'kaynak',
    'LASERCUT': 'lazer',
    # Kullanıcı 2026-09-17: "CNCBEND bölümü abkant olarak tanımlı bizde. Mechanical
    # press ile CNC bend'i bizdeki pres/abkant bölümü olarak düşüneceğiz."
    'CNCBEND': 'pres',
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

# Sistemde üretim bölümü olarak TANIMLI OLMAYAN kapasite bölümleri (atama yapılamaz).
# 2026-09-17: 'bukum' buradan ÇIKTI — CNC büküm bizde Pres/Abkant bölümünün parçası.
BOLUM_DISI = ()

# Zaman birimleri (ERP Y1IVUM) → saniye çarpanı. AD/KG/NR malzeme satırıdır, süre değil.
UM_SANIYE = {'SS': 1.0, 'MN': 60.0, 'HH': 3600.0}

# TESİSİN BÖLÜMLERİ (panel LOKASYON_BOLUMLERI ile aynı olmalı). Kullanıcı 2026-09-17:
# "TK1 için kapasite hesabını montaj hattı, tel üretimi ve plastik enjeksiyon olarak
# sınıflandıracağız." ERP rotasından türeyen bölüm o tesiste YOKSA o iş bu görünüme
# girmez (TK1'de 'lazer' satırı çıkması yanlış olurdu).
LOKASYON_BOLUMLERI = {
    'TK2': ('kaynak', 'montaj', 'metal', 'isleme', 'lazer', 'pres'),
    'TK1': ('montaj', 'tel', 'plastik'),
}

# ── ÜRETİM MÜDÜRÜ KAPASİTE EXCEL'İ (kullanıcı 2026-09-17) ───────────────────
# Q:\UretimPlanlama\Aylık Kapasite Sunum\Kapasite Kullanım Oranı 2026.xlsx
#   Database        : Kod · Makine · SAATLİK ÜRETİM ADEDİ · ÜRETİM HATTI · 2025'te üretildi mi
#   Çalışma Saati   : İsim · Ay · NÖS · %90 · %95 · %100 · BÖLÜM  (kişi bazlı aylık saat)
#   Summary Tk1/Tk2 : Performans = yapılan işin teorik süresi ÷ çalışma saati (hedef %90)
# Süre = 3600 / saatlik adet. Hat adı hem TESİSİ hem BÖLÜMÜ söyler: PP/PULL/IVECO/
# JKP-JKW/LF-LFP TK1 montaj hatlarıdır, ASSEMBLY ise TK2 montajıdır.
EXCEL_HAT_BOLUM = {
    'WELDING': ('TK2', 'kaynak'),
    'LASER CUTTING': ('TK2', 'lazer'),
    'METAL INJECTION': ('TK2', 'metal'),
    'ASSEMBLY': ('TK2', 'montaj'),
    'ABKANT': ('TK2', 'pres'),
    'PP': ('TK1', 'montaj'),
    'PULL': ('TK1', 'montaj'),
    'IVECO': ('TK1', 'montaj'),
    'JKP,JKW': ('TK1', 'montaj'),
    'LF,LFP,LAG (LTK),LSB,PBL': ('TK1', 'montaj'),
    'PLASTIC INJECTION': ('TK1', 'plastik'),
    # Kapasite hattı DEĞİL (bilerek dışarıda): ELECTRONIC, DEPO, POLISING, PTO WITH CABLE
}
EXCEL_ATLANAN = ('ELECTRONIC', 'DEPO', 'POLISING', 'PTO WITH CABLE', '0', '#N/A')

# SÜRESİ EXCEL'DEN ALINMAYAN BÖLÜMLER (kullanıcı 2026-09-17: "bizdeki metal enjeksiyon
# süreleri doğru, metal enjeksiyon süreleri için Forge'yi alalım"). Excel'in saatlik
# adedi çevrim başına; bizim süremiz parça başına (kalıp gözü) — karıştırmak yanlış.
EXCEL_SURE_DISI = ('metal',)

# BÖLÜM/TESİS KARARI FORGE'UN: Excel'in üretim hattı eskiden beri kullanıldığı için
# yanlış olabilir (ör. PBL kodları Excel'de TK1 'LF,LFP,…' hattında ama üretim TK2'de;
# 34 lazer kodu bizde pres/abkant). Excel yalnız SÜRE getirir; kod Forge'de tanımlıysa
# süre ORADAKİ bölüme yazılır, Excel'in hattı yalnız Forge'de HİÇ tanım yoksa kullanılır.


def _excel_hat(ad):
    """Excel hat/bölüm adını (lokasyon, bolum) çiftine çevirir; tanınmıyorsa None."""
    a = ' '.join(str(ad or '').strip().upper().split())
    return EXCEL_HAT_BOLUM.get(a)


def excel_oku(kaynak):
    """Kapasite Excel'ini okur (dosya yolu ya da dosya benzeri nesne).
    Dönüş: {'sureler': [...], 'calisma': [...], 'hatlar': {...}, 'uyarilar': [...]}"""
    import openpyxl
    wb = openpyxl.load_workbook(kaynak, data_only=True, read_only=True)
    try:
        sayfa = {ad.strip().lower(): ad for ad in wb.sheetnames}
        d_ad = next((sayfa[k] for k in sayfa if k.startswith('database')), None)
        c_ad = next((sayfa[k] for k in sayfa if 'saat' in k), None)
        uyarilar = []
        sureler, hatlar = [], {}
        if not d_ad:
            uyarilar.append("'Database' sayfası bulunamadı — süre aktarımı yapılamaz.")
        else:
            ws = wb[d_ad]
            for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
                if i == 1 or not r or r[0] in (None, ''):
                    continue
                kod = str(r[0]).strip()
                if not kod:
                    continue
                try:
                    saatlik = float(r[2] or 0)
                except (TypeError, ValueError):
                    saatlik = 0.0
                hat = ' '.join(str(r[3] or '').strip().upper().split())
                hedef = _excel_hat(hat)
                h = hatlar.setdefault(hat or '(boş)', {'hat': hat or '(boş)', 'kod': 0, 'sureli': 0,
                                                       'lokasyon': hedef[0] if hedef else '',
                                                       'bolum': hedef[1] if hedef else ''})
                h['kod'] += 1
                if saatlik > 0:
                    h['sureli'] += 1
                if not hedef or saatlik <= 0:
                    continue          # tanınmayan hat ya da süresi girilmemiş kod
                sureler.append({'kod': kod, 'lokasyon': hedef[0], 'bolum': hedef[1],
                                'saatlik': saatlik, 'sure_sn': round(3600.0 / saatlik, 2),
                                'makine': str(r[1] or '').strip(), 'hat': hat})
        calisma, ay_gorulen = {}, {}
        if not c_ad:
            uyarilar.append("'Çalışma Saati' sayfası bulunamadı — kapasite saatleri alınamadı.")
        else:
            ws = wb[c_ad]
            for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
                if i == 1 or not r:
                    continue
                # İSİMSİZ SATIRLAR DA SAYILIR: Excel'in altında isim yazmadan bölüme
                # eklenmiş saatler var (ör. 'pp 260 saat' — geçici/ödünç işçilik).
                # Atlarsak kapasite eksik çıkar (Ağustos'ta 260 saat kaybolur).
                hedef = _excel_hat(r[6] if len(r) > 6 else '')
                if not hedef:
                    continue
                try:
                    ay = int(float(r[1] or 0))
                except (TypeError, ValueError):
                    ay = 0
                # NÖS + %90/%95/%100 mesai = kişinin o ay fiilen çalıştığı saat
                saat = 0.0
                for j in (2, 3, 4, 5):
                    try:
                        saat += float(r[j] or 0)
                    except (TypeError, ValueError):
                        pass
                if saat <= 0:
                    continue
                if ay:
                    ay_gorulen[ay] = ay_gorulen.get(ay, 0) + 1
                k = (hedef[0], hedef[1], ay)
                c = calisma.setdefault(k, {'lokasyon': hedef[0], 'bolum': hedef[1], 'ay': ay,
                                           'saat': 0.0, 'kisi': 0, 'isimsiz': 0.0})
                if not str(r[0] or '').strip():
                    c['isimsiz'] += saat
                c['saat'] += saat
                if str(r[0] or '').strip():
                    c['kisi'] += 1          # kişi sayısı yalnız isimli satırlardan
        # AYI YAZILMAMIŞ satırlar (isimsiz ek saatler) dosyadaki ASIL AYA yazılır;
        # ayrı bir 'ay 0' kovası kapasiteyi ikiye bölerdi.
        ana_ay = max(ay_gorulen, key=lambda a: ay_gorulen[a]) if ay_gorulen else 0
        if ana_ay:
            for (lok, bol, ay), c in list(calisma.items()):
                if ay or not c['saat']:
                    continue
                hedef_k = (lok, bol, ana_ay)
                ana = calisma.setdefault(hedef_k, {'lokasyon': lok, 'bolum': bol, 'ay': ana_ay,
                                                   'saat': 0.0, 'kisi': 0, 'isimsiz': 0.0})
                ana['saat'] += c['saat']
                ana['isimsiz'] += c['saat']
                calisma.pop((lok, bol, ay))
        for c in calisma.values():
            c['saat'] = round(c['saat'], 1)
            c['isimsiz'] = round(c.get('isimsiz', 0), 1)
            # Aylık saat → haftalık: ortalama ay 4,345 hafta (365/7/12)
            c['haftalik_saat'] = round(c['saat'] / 4.345, 1)
        taninmayan = [h for h in hatlar.values()
                      if not h['lokasyon'] and h['hat'].upper() not in EXCEL_ATLANAN]
        if taninmayan:
            uyarilar.append('Tanınmayan üretim hattı: ' +
                            ', '.join(f"{h['hat']} ({h['kod']} kod)" for h in taninmayan[:6]))
        return {'sureler': sureler, 'calisma': sorted(calisma.values(), key=lambda x: (x['lokasyon'], x['bolum'])),
                'hatlar': sorted(hatlar.values(), key=lambda x: -x['kod']), 'uyarilar': uyarilar,
                'aylar': sorted(ay_gorulen)}
    finally:
        wb.close()


def _excel_hedefler(kod, f_tam, f_kok):
    """Excel'deki bir kodun yazılacağı hedefler.
    Dönüş: (hedefler, kaynak) · hedefler = [(lokasyon, bolum, mevcut_ct)]
      kaynak 'forge'  → kod Forge'de tanımlı, süre O bölüme yazılır
      kaynak 'adim'   → Forge'de ADIM EKLİ satırlar var (tel): Excel tek süre verir,
                        adımlara bölemeyiz → YAZILMAZ
      kaynak 'excel'  → Forge'de tanım yok, Excel'in hattı kullanılır
    """
    tam = f_tam.get(_norm(kod))
    if tam:
        return [(l, b, ct) for (l, b), ct in tam.items()], 'forge'
    kokler = f_kok.get(_kok(kod))
    if kokler:
        return [(l, b, ct) for (l, b), ct in kokler.items()], 'adim'
    return [], 'excel'


def _excel_forge_haritasi(conn):
    """{norm_kod: {(lokasyon, bolum): ct}} — tam kod ve kök kod için ayrı ayrı."""
    tam, kok = {}, {}
    for r in conn.execute("SELECT referans_kodu, COALESCE(bolum,'kaynak') b, "
                          "COALESCE(lokasyon,'TK2') l, COALESCE(hedef_cycle_time_sn,0) ct "
                          "FROM referans_listesi"):
        ct = float(r['ct'] or 0)
        anahtar = (r['l'], r['b'])
        tam.setdefault(_norm(r['referans_kodu']), {})[anahtar] = ct
        if _kok(r['referans_kodu']) != _norm(r['referans_kodu']):
            kok.setdefault(_kok(r['referans_kodu']), {})[anahtar] = ct
    return tam, kok


def excel_onizle(conn, veri, lokasyon=''):
    """Excel'deki sürelerin bizdeki tanımlarla farkını çıkarır (YAZMADAN).
    'yeni' = bizde o bölümde tanım yok · 'degisen' = süre farklı · 'ayni' = aynı."""
    f_tam, f_kok = _excel_forge_haritasi(conn)
    talep = {}
    for r in conn.execute("SELECT kod, SUM(kalan) k FROM kapasite_talep WHERE kalan > 0 GROUP BY kod"):
        talep[r['kod']] = float(r['k'] or 0)
    yeni = degisen = ayni = bos = 0
    atlanan = {'metal': 0, 'adim': 0, 'hat_farki': 0}
    ozet, buyuk_fark = {}, []
    for x in veri['sureler']:
        hedefler, kaynak = _excel_hedefler(x['kod'], f_tam, f_kok)
        if kaynak == 'adim':
            atlanan['adim'] += 1       # tel: Excel tek süre veriyor, adımlara bölemeyiz
            continue
        if kaynak == 'excel':
            hedefler = [(x['lokasyon'], x['bolum'], None)]
        for lok, bolum, eski in hedefler:
            if lokasyon and lok != lokasyon:
                continue
            if bolum in EXCEL_SURE_DISI:
                atlanan['metal'] += 1
                continue
            if kaynak == 'forge' and (lok, bolum) != (x['lokasyon'], x['bolum']):
                atlanan['hat_farki'] += 1   # Excel başka hat diyor; Forge kazanır
            if eski is None:
                durum = 'yeni'
            elif eski <= 0:
                durum = 'bos'          # tanım var, süre girilmemiş → ezme riski YOK
            else:
                durum = 'ayni' if abs(eski - x['sure_sn']) < 0.05 else 'degisen'
            if durum == 'yeni':
                yeni += 1
            elif durum == 'bos':
                bos += 1
            elif durum == 'degisen':
                degisen += 1
            else:
                ayni += 1
            o = ozet.setdefault((lok, bolum), {'lokasyon': lok, 'bolum': bolum, 'yeni': 0,
                                               'degisen': 0, 'ayni': 0, 'bos': 0, 'opr': 0})
            o[durum] += 1
            if talep.get(x['kod']):
                o['opr'] += 1
            if durum == 'degisen' and eski > 0 and abs(x['sure_sn'] - eski) / eski > 0.5:
                adet = talep.get(x['kod'], 0)
                buyuk_fark.append({
                    'kod': x['kod'], 'lokasyon': lok, 'bolum': bolum,
                    'forge': round(eski, 1), 'excel': x['sure_sn'],
                    'oran': round(100 * (x['sure_sn'] - eski) / eski),
                    'adet': round(adet),
                    'saat_etki': round(adet * abs(x['sure_sn'] - eski) / 3600, 1)})
    buyuk_fark.sort(key=lambda z: (-z['saat_etki'], -abs(z['oran'])))
    return {'yeni': yeni, 'degisen': degisen, 'ayni': ayni, 'bos': bos,
            'toplam': yeni + degisen + ayni + bos, 'atlanan': atlanan,
            'bolumler': sorted(ozet.values(), key=lambda x: (x['lokasyon'], x['bolum'])),
            'buyuk_fark': buyuk_fark[:60], 'buyuk_fark_sayisi': len(buyuk_fark),
            'buyuk_fark_saat': round(sum(z['saat_etki'] for z in buyuk_fark), 1)}


def excel_sure_uygula(conn, veri, kullanici='', lokasyon='', sadece_opr=False,
                      sadece_bos=True):
    """Excel sürelerini referans_listesi'ne yazar (bölüm/tesis Excel'in hattından).

    sadece_bos=True (VARSAYILAN) → bizde süresi OLMAYAN kodlara yazar, mevcut süreyi
      EZMEZ. Kaynak/lazer gibi bölümlerde bizim süreler sahada ölçüldü; Excel'in
      saatlik adedi yuvarlak bir plan değeri olabilir (ör. 10.130.2412: bizde 255,3 sn,
      Excel'de 36 sn). Ezmek istenirse bu bayrak kapatılır.
    sadece_opr=True → yalnız açık ihtiyacı olan kodlara dokunur."""
    talep = {r['kod'] for r in conn.execute("SELECT DISTINCT kod FROM kapasite_talep WHERE kalan > 0")}
    f_tam, f_kok = _excel_forge_haritasi(conn)
    yeni = guncel = ayni = 0
    atlanan = {'metal': 0, 'adim': 0}
    for x in veri['sureler']:
        if sadece_opr and x['kod'] not in talep:
            continue
        hedefler, kaynak = _excel_hedefler(x['kod'], f_tam, f_kok)
        if kaynak == 'adim':
            atlanan['adim'] += 1
            continue
        if kaynak == 'excel':
            hedefler = [(x['lokasyon'], x['bolum'], None)]
        for hedef_lok, hedef_bolum, _eski in hedefler:
            if lokasyon and hedef_lok != lokasyon:
                continue
            if hedef_bolum in EXCEL_SURE_DISI:
                atlanan['metal'] += 1      # metal süresi Forge'den kalır
                continue
            r = conn.execute(
                "SELECT id, COALESCE(hedef_cycle_time_sn,0) ct FROM referans_listesi "
                "WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ','')) AND bolum=? "
                "AND COALESCE(lokasyon,'TK2')=?", (x['kod'], hedef_bolum, hedef_lok)).fetchone()
            if r is None:
                conn.execute("INSERT INTO referans_listesi (referans_kodu, hedef_cycle_time_sn, bolum, "
                             "lokasyon) VALUES (?,?,?,?)",
                             (x['kod'], x['sure_sn'], hedef_bolum, hedef_lok))
                yeni += 1
            elif sadece_bos and float(r['ct'] or 0) > 0:
                ayni += 1        # süresi var, dokunma (mevcut tanım korunur)
                continue
            elif abs(float(r['ct']) - x['sure_sn']) >= 0.05:
                conn.execute("UPDATE referans_listesi SET hedef_cycle_time_sn=? WHERE id=?",
                             (x['sure_sn'], r['id']))
                guncel += 1
            else:
                ayni += 1
                continue
            # Geçmiş üretim kayıtlarının cycle'ı — /api/referanslar ve sure_yukle ile aynı kural
            conn.execute(
                "UPDATE uretim_kayitlari SET cycle_time_sn=? "
                "WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ','')) "
                "AND vardiya_id IN (SELECT id FROM vardiyalar WHERE COALESCE(lokasyon,'TK2')=? "
                "                   AND COALESCE(bolum,'kaynak')=?)",
                (x['sure_sn'], x['kod'], hedef_lok, hedef_bolum))
    conn.commit()
    return {'yeni': yeni, 'guncel': guncel, 'ayni': ayni, 'atlanan': atlanan}


def excel_calisma_uygula(conn, veri, kullanici=''):
    """Çalışma saatlerini kaydeder ve bölümlerin HAFTALIK KAPASİTESİNİ bu saatle
    günceller (kapasite_parametre.haftalik_saat_elle). Aylık saat ÷ 4,345 = haftalık."""
    yazilan = 0
    for c in veri['calisma']:
        conn.execute(
            "INSERT INTO kapasite_calisma_saati (lokasyon, bolum, ay, saat, kisi, kaynak, updated_at) "
            "VALUES (?,?,?,?,?, 'excel', datetime('now','localtime')) "
            "ON CONFLICT(lokasyon, bolum, ay) DO UPDATE SET saat=excluded.saat, kisi=excluded.kisi, "
            "updated_at=datetime('now','localtime')",
            (c['lokasyon'], c['bolum'], c['ay'], c['saat'], c['kisi']))
        conn.execute(
            "INSERT INTO kapasite_parametre (lokasyon, bolum, haftalik_saat_elle, guncelleyen, updated_at) "
            "VALUES (?,?,?,?, datetime('now','localtime')) "
            "ON CONFLICT(lokasyon, bolum) DO UPDATE SET haftalik_saat_elle=excluded.haftalik_saat_elle, "
            "guncelleyen=excluded.guncelleyen, updated_at=datetime('now','localtime')",
            (c['lokasyon'], c['bolum'], c['haftalik_saat'], kullanici))
        yazilan += 1
    conn.commit()
    return {'bolum': yazilan}


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
# AÇIK ÜRETİM İHTİYACI (OPR). XPRO90 = yalnız açık emirler görünümü:
#   Q0AVAN 10 = OPR — MRP ihtiyaç yarattı, launch alınmadı (asıl kapasite yükü burada)
#          40 = launch açık · 45/50 = TK1 akışı
#   Q0QTOR sipariş adedi · Q0QTRI teyit edilen → KALAN = ordine - rientrata
#   Q0FPD* = "Dt fine produzione" (termin): haftalık kovalar buna göre kurulur.
SQL_TALEP = f"""
    SELECT Q0ARTI, Q0AVAN, Q0QTOR, Q0QTRI, Q0RED1, Q0RED2, Q0RENU,
           Q0FPD1, Q0FPD2, Q0FPD3, Q0FPD4, Q0IPD1, Q0IPD2, Q0IPD3, Q0IPD4
    FROM {AS400_SEMA}.XPRO90
"""
TALEP_DURUM_AD = {'10': 'OPR (ihtiyaç)', '40': 'Launch açık', '45': 'Launch (TK1)',
                  '50': 'Launch (TK1)'}

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
            # Kod içindeki VARSAYILAN değişmişse (ör. 2026-09-17 CNCBEND 'bukum' → 'pres')
            # elle dokunulmamış satır yeni varsayılana geçer; elle=1 satıra ASLA dokunma.
            if not var['elle']:
                vars_yeni = VARSAYILAN_KAYNAK_BOLUM.get(kod.upper(), var['bolum'] or '')
                if vars_yeni != (var['bolum'] or ''):
                    c.execute("UPDATE kapasite_kaynak SET bolum=? WHERE kaynak_kod=?", (vars_yeni, kod))
    uretim = sum(1 for u in urunler if u[2] == 'P' and not u[3])
    fason = sum(1 for u in urunler if u[2] == 'A' and not u[3])
    c.execute("INSERT INTO kapasite_senk (tarih, kullanici, urun, uretim, fason, rota, kaynak) "
              "VALUES (?,?,?,?,?,?,?)",
              (simdi, kullanici, len(urunler), uretim, fason, yazilan, len(kaynaklar)))
    conn.commit()
    return {'urun': len(urunler), 'uretim': uretim, 'fason': fason, 'rota': yazilan,
            'mukerrer_rota': mukerrer, 'kaynak': len(kaynaklar), 'yeni_kaynak': yeni,
            'tarih': simdi}


def _erp_tarih(ss, aa, mm, gg):
    """ERP tarihi (yüzyıl/yıl/ay/gün ayrı kolonlar) → 'YYYY-MM-DD'. Geçersizse ''.
    DİKKAT: SS = YÜZYIL (20), AA = yıl (26) → 2026. '1900+' varsayımı 3926 üretir."""
    try:
        ss, aa, mm, gg = int(_f(ss)), int(_f(aa)), int(_f(mm)), int(_f(gg))
        if not (mm and gg):
            return ''
        yil = ss * 100 + aa if ss >= 19 else 2000 + aa
        return date(yil, mm, gg).isoformat()
    except (ValueError, TypeError):
        return ''


def talep_senkron(conn, kullanici='', baglan=None):
    """Açık üretim ihtiyacını (OPR + açık launch) AS400'den çeker.

    Kullanıcı 2026-09-17: "biz sistemde şu anda OPR'si yani üretim ihtiyacı oluşmuş
    referanslara bakacağız çünkü o kodları üretmemiz gerekiyor." Havuzdaki 29 bin
    koddan yalnız bunlar kapasiteyi ilgilendirir (2026-09-17: 1.900 civarı kod)."""
    if baglan is None:
        import sys as _sys, os as _os
        _sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'as400'))
        import as400_config as CFG
        baglan = CFG.baglan
    cn = baglan(timeout=60)
    simdi = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    try:
        cur = cn.cursor()
        cur.execute(SQL_TALEP)
        ham = cur.fetchall()
    finally:
        try:
            cn.close()
        except Exception:
            pass
    satirlar = []
    for r in ham:
        kod = _t(r[0])
        if not kod:
            continue
        adet, teyit = _f(r[2]), _f(r[3])
        emir = f'{int(_f(r[4])):02d}-{int(_f(r[5])):02d}-{int(_f(r[6]))}'
        satirlar.append((kod, _t(r[1]), emir, adet, teyit, max(0.0, adet - teyit),
                         _erp_tarih(r[7], r[8], r[9], r[10]), _erp_tarih(r[11], r[12], r[13], r[14]),
                         simdi))
    c = conn.cursor()
    c.execute("DELETE FROM kapasite_talep")      # tam yenileme: kapanan emir bizde kalmasın
    c.executemany(
        "INSERT INTO kapasite_talep (kod, durum, emir_no, adet, teyit, kalan, bitis, baslangic, "
        "senk_at) VALUES (?,?,?,?,?,?,?,?,?)", satirlar)
    conn.commit()
    bugun = date.today().isoformat()
    acik = [x for x in satirlar if x[5] > 0]
    return {'emir': len(satirlar), 'acik': len(acik),
            'kod': len({x[0] for x in acik}),
            'adet': round(sum(x[5] for x in acik)),
            'gecikmis': sum(1 for x in acik if x[6] and x[6] < bugun),
            'gecikmis_adet': round(sum(x[5] for x in acik if x[6] and x[6] < bugun)),
            'tarihsiz': sum(1 for x in acik if not x[6]),
            'tarih': simdi}


def parametreler(conn, lokasyon='TK2', bolumler=()):
    """Bölümlerin kapasite parametreleri (tanımsızsa varsayılan satır üretilir).
    Haftalık kapasite saati = makine × vardiya × vardiya_saat × gün × verimlilik."""
    out = {}
    for r in conn.execute("SELECT * FROM kapasite_parametre WHERE COALESCE(lokasyon,'TK2')=?",
                          (lokasyon,)):
        out[r['bolum']] = {k: r[k] for k in r.keys()}
    for b in bolumler:
        out.setdefault(b, {'lokasyon': lokasyon, 'bolum': b, 'makine': 1, 'vardiya': 2,
                           'vardiya_saat': 7.5, 'gun': 5, 'verimlilik': 75, 'oee_kullan': 0,
                           'not_metni': '', 'guncelleyen': '', 'updated_at': None,
                           'varsayilan': True})
    for b, p in out.items():
        p.setdefault('haftalik_saat_elle', 0)
        elle = float(p.get('haftalik_saat_elle') or 0)
        # ELLE SAAT (üretim müdürü Excel'indeki gerçek çalışma saati) formülü EZER:
        # gerçekte çalışılan saat, makine × vardiya varsayımından daha doğrudur.
        p['haftalik_saat'] = elle if elle > 0 else round(
            float(p['makine'] or 0) * float(p['vardiya'] or 0) * float(p['vardiya_saat'] or 0) *
            float(p['gun'] or 0) * (float(p['verimlilik'] or 0) / 100.0), 1)
        p['saat_kaynagi'] = 'excel' if elle > 0 else 'formul'
    return out


def gerceklesen_oee(bolum, lokasyon, hafta=4):
    """Bölümün son <hafta> haftadaki gerçekleşen OEE'si (%) — yoksa None.
    Kapasite parametresinde 'OEE kullan' seçilirse verimlilik yerine BU geçer."""
    try:
        from oee import hesapla_oee_ozet
        bit = date.today()
        bas = bit - timedelta(days=7 * hafta)
        o = hesapla_oee_ozet(bas.isoformat(), bit.isoformat(), None, bolum, lokasyon)
        v = float(o.get('ort_oee') or 0)
        return round(v, 1) if v > 0 else None
    except Exception as e:
        print(f'[kapasite] gerçekleşen OEE alınamadı ({bolum}/{lokasyon}): {e}')
        return None


def _kok(kod):
    """Adım eki atılmış kod: '93.00.1347 KESIM' → '93.00.1347'.
    TEL üretiminde bir ERP kodu Forge'de birden çok satırdır (her proses adımı ayrı
    referans, kendi süresiyle) — eşleme kökten yapılır."""
    return _norm(str(kod or '').split(' ')[0])


def forge_haritasi(conn, lokasyon):
    """Forge'deki (referans_listesi) bölüm tanımları — SINIFLANDIRMANIN BİRİNCİL KAYNAĞI.

    Dönüş: {anahtar: {bolum: {'sn': toplam_saniye, 'satir': kaç referans satırı}}}
    Anahtar hem TAM kod hem KÖK kod olarak yazılır; arama önce tam kodla yapılır.
    Aynı bölümdeki birden çok satırın süresi TOPLANIR: tel'de ürün kesim + kapama +
    son montaj adımlarının HEPSİNDEN geçer, kapasite yükü adımların toplamıdır.
    """
    tam, kok = {}, {}
    for r in conn.execute(
            "SELECT referans_kodu, COALESCE(bolum,'kaynak') bolum, "
            "COALESCE(hedef_cycle_time_sn,0) ct FROM referans_listesi "
            "WHERE COALESCE(lokasyon,'TK2') = ?", (lokasyon,)):
        sn = float(r['ct'] or 0)
        for harita, anahtar in ((tam, _norm(r['referans_kodu'])), (kok, _kok(r['referans_kodu']))):
            b = harita.setdefault(anahtar, {}).setdefault(r['bolum'], {'sn': 0.0, 'satir': 0})
            b['sn'] += sn
            b['satir'] += 1
    return tam, kok


def _forge_bolumleri(kod, tam, kok):
    """Kodun Forge'deki bölüm dağılımı: önce tam kod, yoksa kök (adım ekli tel kodları)."""
    return tam.get(_norm(kod)) or kok.get(_kok(kod))


def talep_ozet(conn, lokasyon='TK2', haftalar=(2, 4, 6, 8)):
    """Bölüm bazlı haftalık ihtiyaç: kaç adet ve kaç SAAT iş var.

    SINIFLANDIRMA SIRASI (kullanıcı 2026-09-17: "kod dağılımları Forge'de bölümlere göre
    tanımlı olan referanslardan oluşacak gibi düşün"):
      1) Kod bu tesiste Forge'de tanımlıysa → bölüm(ler) ve süre(ler) ORADAN gelir.
         (Tel'de adım ekli satırların süreleri toplanır; ERP rotası 'tel' demeyi bilmez,
          yalnız OUTCASE kaynağını tanır — 2026-09-17'de tel'de 21 kod görünmesinin sebebi
          buydu.)
      2) Kod DİĞER tesiste tanımlıysa → o tesisin işidir, bu görünüme girmez ('diger_tesis').
      3) Hiçbir tesiste tanımı yoksa → ERP rotasından türetilen bölüm kullanılır ('erp_kod');
         rotası tamamen dış işlemse 'dis_islem', hiç bölüm çıkmıyorsa 'siniflanamayan'.

    · Kovalar KÜMÜLATİF: "4 hafta" = 28 gün içindeki iş + TERMİNİ GEÇMİŞ iş.
    · Kapasite: makine × vardiya × vardiya_saat × gün × verimlilik (kapasite_parametre).
    """
    bugun = date.today()
    sinir = {h: (bugun + timedelta(days=7 * h)).isoformat() for h in haftalar}
    diger_lok = 'TK1' if lokasyon == 'TK2' else 'TK2'
    gecerli = set(LOKASYON_BOLUMLERI.get(lokasyon, LOKASYON_BOLUMLERI['TK2']))
    f_tam, f_kok = forge_haritasi(conn, lokasyon)
    d_tam, d_kok = forge_haritasi(conn, diger_lok)
    rota = {}
    for r in conn.execute("SELECT kod, bolum, sure_sn FROM kapasite_urun_bolum"):
        rota.setdefault(r['kod'], {})[r['bolum']] = float(r['sure_sn'] or 0)
    bolumler = {}
    dis_islem = {'kod': set(), 'adet': 0.0}
    diger_tesis = {'kod': set(), 'adet': 0.0}
    siniflanamayan = {'kod': set(), 'adet': 0.0}
    tarihsiz = {'adet': 0.0, 'emir': 0}
    for t in conn.execute("SELECT kod, kalan, bitis FROM kapasite_talep WHERE kalan > 0"):
        kod, kalan, bitis = t['kod'], float(t['kalan'] or 0), (t['bitis'] or '')
        forge = _forge_bolumleri(kod, f_tam, f_kok)
        if forge:
            # SÜRE: Forge'de tanımlıysa o, değilse (0 ise) ERP rota süresi — bölüm
            # Forge'den gelse bile süresi girilmemiş olabilir; ERP süresi hiç yoktan
            # iyidir, ama hangisinin kullanıldığı sayılır ve panelde gösterilir.
            erp_sn = rota.get(kod, {})
            dagilim = {}
            for b, v in forge.items():
                if b not in gecerli:
                    continue
                if v['sn'] > 0:
                    dagilim[b] = (v['sn'], 'forge', 'forge')
                else:
                    dagilim[b] = (erp_sn.get(b, 0.0), 'forge', 'erp' if erp_sn.get(b) else 'yok')
            if not dagilim:
                diger_tesis['kod'].add(kod)
                diger_tesis['adet'] += kalan
                continue
        elif _forge_bolumleri(kod, d_tam, d_kok):
            diger_tesis['kod'].add(kod)
            diger_tesis['adet'] += kalan
            continue
        else:
            erp = {b: sn for b, sn in (rota.get(kod) or {}).items() if b in gecerli}
            if not erp:
                # Rotası var ama bu tesiste geçerli bölüm çıkmıyor: ya tamamen dış
                # işlem, ya da öteki tesisin işi. Rota satırı hiç yoksa sınıflanamaz.
                if rota.get(kod):
                    diger_tesis['kod'].add(kod)
                    diger_tesis['adet'] += kalan
                elif conn.execute("SELECT 1 FROM kapasite_rota WHERE kod=? LIMIT 1", (kod,)).fetchone():
                    dis_islem['kod'].add(kod)
                    dis_islem['adet'] += kalan
                else:
                    siniflanamayan['kod'].add(kod)
                    siniflanamayan['adet'] += kalan
                continue
            dagilim = {b: (sn, 'erp', 'erp' if sn else 'yok') for b, sn in erp.items()}
        if not bitis:
            tarihsiz['adet'] += kalan
            tarihsiz['emir'] += 1
        gecikmis = bool(bitis) and bitis < bugun.isoformat()
        for bolum, (sn, bolum_kaynak, sure_kaynak) in dagilim.items():
            b = bolumler.setdefault(bolum, {
                'bolum': bolum, 'kod': set(), 'suresiz_kod': set(), 'forge_kod': set(),
                'erp_kod': set(), 'sure_forge': set(), 'sure_erp': set(),
                'adet': 0.0, 'gecikmis_adet': 0.0, 'gecikmis_sn': 0.0,
                'hafta': {h: {'adet': 0.0, 'sn': 0.0} for h in haftalar}})
            (b['forge_kod'] if bolum_kaynak == 'forge' else b['erp_kod']).add(kod)
            if sure_kaynak == 'forge':
                b['sure_forge'].add(kod)
            elif sure_kaynak == 'erp':
                b['sure_erp'].add(kod)
            else:
                b['suresiz_kod'].add(kod)
            b['kod'].add(kod)
            b['adet'] += kalan
            if gecikmis:
                b['gecikmis_adet'] += kalan
                b['gecikmis_sn'] += kalan * (sn or 0)
            for h in haftalar:
                if gecikmis or (bitis and bitis <= sinir[h]):
                    b['hafta'][h]['adet'] += kalan
                    b['hafta'][h]['sn'] += kalan * (sn or 0)
    out = []
    for b in bolumler.values():
        out.append({
            'bolum': b['bolum'], 'kod_sayisi': len(b['kod']),
            'forge_kod': len(b['forge_kod']), 'erp_kod': len(b['erp_kod']),
            'sure_forge': len(b['sure_forge']), 'sure_erp': len(b['sure_erp']),
            'suresiz_kod': len(b['suresiz_kod']),
            'adet': round(b['adet']), 'gecikmis_adet': round(b['gecikmis_adet']),
            'gecikmis_saat': round(b['gecikmis_sn'] / 3600, 1),
            'hafta': {str(h): {'adet': round(v['adet']), 'saat': round(v['sn'] / 3600, 1)}
                      for h, v in b['hafta'].items()},
        })
    # KAPASİTE KARŞILAŞTIRMASI: kova kümülatif olduğu için kapasite de kümülatiftir
    # (4 haftalık kova ↔ 4 haftalık kapasite). Gecikmiş iş kovaların içinde olduğundan
    # doluluk >%100 çıkabilir — bu bir hata değil, "yetişmiyoruz" demektir.
    par = parametreler(conn, lokasyon, [b['bolum'] for b in out])
    for b in out:
        p = dict(par.get(b['bolum'], {}))
        if p.get('oee_kullan'):
            oee = gerceklesen_oee(b['bolum'], lokasyon)
            p['gerceklesen_oee'] = oee
            if oee and not float(p.get('haftalik_saat_elle') or 0):
                p['verimlilik'] = oee
                p['haftalik_saat'] = round(float(p['makine'] or 0) * float(p['vardiya'] or 0) *
                                           float(p['vardiya_saat'] or 0) * float(p['gun'] or 0) *
                                           (oee / 100.0), 1)
        b['kapasite'] = p
        for h in haftalar:
            kap = round(float(p.get('haftalik_saat') or 0) * h, 1)
            saat = b['hafta'][str(h)]['saat']
            b['hafta'][str(h)]['kapasite_saat'] = kap
            b['hafta'][str(h)]['doluluk'] = round(100 * saat / kap, 1) if kap > 0 else None
    out.sort(key=lambda x: -x['hafta'][str(max(haftalar))]['saat'])
    son = conn.execute("SELECT MAX(senk_at) s FROM kapasite_talep").fetchone()
    havuz = conn.execute("SELECT COUNT(*) n FROM kapasite_urun_bolum").fetchone()['n']
    talep_var = conn.execute("SELECT COUNT(*) n FROM kapasite_talep WHERE kalan > 0").fetchone()['n']
    # "Hiç bölüm çıkmadı" iki AYRI sebepten olur; karıştırmak yanlış teşhis yaratır:
    #   rota_yok  → ürün/rota hiç çekilmedi, sınıflandırma yapılamıyor (Ürün + Rota Çek)
    #   talep_yok → ihtiyaç listesi boş
    uyari = ''
    if not havuz:
        uyari = ('Ürün ve rota listesi çekilmemiş — Forge\'de tanımı olmayan kodlar '
                 'sınıflandırılamaz. "Ürün + Rota Çek" ile listeyi alın.')
    elif not talep_var:
        uyari = 'Açık üretim ihtiyacı (OPR) bulunamadı — "İhtiyacı Tazele" ile listeyi çekin.'
    bugun_iso = bugun.isoformat()
    say = conn.execute(
        "SELECT COUNT(*) emir, COALESCE(SUM(kalan),0) adet, "
        "SUM(CASE WHEN bitis <> '' AND bitis < ? THEN 1 ELSE 0 END) gecikmis_emir, "
        "COALESCE(SUM(CASE WHEN bitis <> '' AND bitis < ? THEN kalan ELSE 0 END),0) gecikmis_adet "
        "FROM kapasite_talep WHERE kalan > 0", (bugun_iso, bugun_iso)).fetchone()
    toplam = {'kod': sum(b['kod_sayisi'] for b in out), 'emir': say['emir'],
              'adet': round(say['adet']), 'gecikmis_emir': say['gecikmis_emir'],
              'gecikmis_adet': round(say['gecikmis_adet']),
              'suresiz_kod': sum(b['suresiz_kod'] for b in out),
              'erp_kod': sum(b['erp_kod'] for b in out),
              'forge_kod': sum(b['forge_kod'] for b in out),
              'sure_forge': sum(b['sure_forge'] for b in out),
              'sure_erp': sum(b['sure_erp'] for b in out),
              'saat': {str(h): round(sum(b['hafta'][str(h)]['saat'] for b in out), 1) for h in haftalar},
              'kapasite_saat': {str(h): round(sum(b['hafta'][str(h)].get('kapasite_saat') or 0
                                                  for b in out), 1) for h in haftalar}}
    return {'lokasyon': lokasyon, 'haftalar': list(haftalar), 'bolumler': out, 'toplam': toplam,
            'dis_islem': {'kod': len(dis_islem['kod']), 'adet': round(dis_islem['adet'])},
            'diger_tesis': {'kod': len(diger_tesis['kod']), 'adet': round(diger_tesis['adet']),
                            'lokasyon': diger_lok},
            'siniflanamayan': {'kod': len(siniflanamayan['kod']), 'adet': round(siniflanamayan['adet'])},
            'tarihsiz': {'emir': tarihsiz['emir'], 'adet': round(tarihsiz['adet'])},
            'havuz_satiri': havuz, 'acik_talep': talep_var, 'uyari': uyari,
            'senk_at': (son['s'] if son else None)}


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
    makine, iscilik, iscilik_var = {}, 0.0, False
    for s in satirlar:
        kod = str(s.get('kaynak_kod') or '').strip().upper()
        sn = float(s.get('sure_sn') or 0)
        bolum = (harita.get(kod, ('', ''))[0] or '').strip()
        if not bolum:
            continue                                   # dış işlem / malzeme / test
        if kod in ISCILIK_KAYNAKLARI:
            iscilik_var = True
            iscilik += max(0.0, sn)
        else:
            # SÜRESİ 0 OLAN ADIM DA BÖLÜME YAZILIR (2026-09-17): ERP'de rota var ama
            # süre girilmemişse kod "bölümü yok" diye kaybolmamalı — tam tersine
            # süre tanımlanacaklar listesinde görünmesi gerekiyor.
            makine[bolum] = makine.get(bolum, 0.0) + max(0.0, sn)
    if makine:
        if iscilik:
            enb = max(makine, key=lambda b: makine[b])
            makine[enb] += iscilik
        return {b: round(v, 2) for b, v in makine.items()}
    return {'montaj': round(iscilik, 2)} if iscilik_var else {}


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
            satirlar.append((kod, bolum, sn))     # sn=0 → 'süresi tanımsız' olarak görünür
    c = conn.cursor()
    c.execute("DELETE FROM kapasite_urun_bolum")
    c.executemany("INSERT INTO kapasite_urun_bolum (kod, bolum, sure_sn) VALUES (?,?,?)", satirlar)
    conn.commit()
    return len(satirlar)


def _norm(kod):
    """Kod eşlemesi: büyük harf + boşluksuz. NOKTA ANLAMLIDIR (10.300.1 ≠ 10.3001)."""
    return str(kod or '').strip().upper().replace(' ', '')


def _tanimli_sql(lokasyon, sureli=False, bolum=''):
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
    if bolum:
        q += " AND COALESCE(bolum,'kaynak') = ?"
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
          talep='1', limit=100, offset=0):
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
        # SINIFLANDIRMA SIRASI ÖZETLE AYNI: önce Forge tanımı, Forge'de hiç tanımı
        # olmayan kodda ERP rotasından türetilen bölüm. (Eskiden yalnız ERP rotasına
        # bakıyordu: bizde 'pres' olan lazer kodları 'lazer' süzgecinde çıkıyordu.)
        kosul.append(f"(u.kod IN ({_tanimli_sql(lokasyon, bolum=bolum)}) OR "
                     f"(u.kod NOT IN ({_tanimli_sql(lokasyon)}) AND EXISTS "
                     "(SELECT 1 FROM kapasite_urun_bolum b WHERE b.kod=u.kod AND b.bolum=?)))")
        if lokasyon:
            par.append(lokasyon)
        par.append(bolum)
        if lokasyon:
            par.append(lokasyon)
        par.append(bolum)
    if talep == '1':
        # Kullanıcı 2026-09-17: "hepsini süzmeye ve atamaya gerek yok, OPR'si oluşmuş
        # referanslara bakacağız" → varsayılan süzgeç açık üretim ihtiyacıdır.
        kosul.append("EXISTS (SELECT 1 FROM kapasite_talep t WHERE t.kod=u.kod AND t.kalan > 0)")
    elif talep == '0':
        kosul.append("NOT EXISTS (SELECT 1 FROM kapasite_talep t WHERE t.kod=u.kod AND t.kalan > 0)")
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
    if sirala == 'sure':
        duzen = ("ORDER BY (SELECT COALESCE(SUM(sure_sn),0) FROM kapasite_urun_bolum b "
                 "WHERE b.kod=u.kod) DESC, u.kod")
    elif sirala == 'talep':      # en çok iş bekleyen kod üstte — süre girişi oradan başlasın
        duzen = ("ORDER BY (SELECT COALESCE(SUM(kalan),0) FROM kapasite_talep t "
                 "WHERE t.kod=u.kod) DESC, u.kod")
    else:
        duzen = "ORDER BY u.kod"
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
        talepler = {}
        for r in conn.execute(
                "SELECT kod, SUM(kalan) kalan, MIN(CASE WHEN bitis <> '' THEN bitis END) ilk_termin, "
                "COUNT(*) emir FROM kapasite_talep WHERE kalan > 0 AND kod IN "
                f"({isaret}) GROUP BY kod", kodlar):
            talepler[r['kod']] = {'kalan': round(float(r['kalan'] or 0)),
                                  'ilk_termin': r['ilk_termin'] or '', 'emir': r['emir']}
        for s in satirlar:
            s['talep'] = talepler.get(s['kod'])
            s['rota'] = rota.get(s['kod'], {})
            hepsi = tanim.get(s['kod'], [])
            s['tanimlar'] = [t for t in hepsi if not lokasyon or t['lokasyon'] == lokasyon]
            s['diger_tesis'] = [t for t in hepsi if lokasyon and t['lokasyon'] != lokasyon]
            s['durum'] = ('tanimli' if any(t['ct'] > 0 for t in s['tanimlar'])
                          else ('suresiz' if s['tanimlar'] else 'atanmamis'))
    return {'toplam': toplam, 'satirlar': satirlar, 'limit': int(limit), 'offset': int(offset)}
