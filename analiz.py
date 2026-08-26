# -*- coding: utf-8 -*-
"""
analiz.py — Üretim / duruş / sinyal verisinden OTOMATİK BULGU çıkaran analiz motoru.

Neden ayrı modül: app.py'yi import EDEMEZ (Flask'ı ayağa kaldırır) — mail_raporu.py
gibi bu da kendi bağlantısını açar. app.py, mail_raporu.py ve dashboard aynı
bulguları buradan alır; kural tek yerde durur.

İKİ KATMAN
----------
1. YEREL MOTOR (bu dosyanın gövdesi) — istatistik + kural tabanlı. İnternet
   gerektirmez, ücretsiz, aynı veriye HEP aynı sonucu verir ve her bulgu
   "hangi sayıdan geldi" ile birlikte döner (denetlenebilir).
2. YORUM KATMANI (yorum_uret) — bulguları Claude API'ye verip akıcı bir yönetici
   özeti yazdırır. ai_config.json yoksa/kapalıysa SESSİZCE atlanır; yerel bulgular
   her hâlükârda üretilir. Yani AI arızası analizi durdurmaz.

BULGU SÖZLEŞMESİ
----------------
{'kod', 'siddet': kritik|uyari|bilgi, 'baslik', 'detay', 'oneri', 'sayisal': {...}}
Her bulgu TEK bir cümlede ne olduğunu söyler; 'sayisal' hesabın girdisidir
(panelde tabloya dönüşür, mailde metne). Eşikler modül başında toplanmıştır —
sahadan geri bildirim gelince tek yerden ayarlanır.
"""
import json
import os
import sqlite3
from collections import defaultdict
from datetime import datetime, timedelta

from oee import PLANLI_DURUS_KATEGORILER, durus_tipi_belirle   # noqa: F401  (tip çözümü ortak)

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH     = os.path.join(PROJECT_DIR, 'uretim.db')
PILOT_DB    = os.path.join(PROJECT_DIR, 'pilot', 'pilot.db')
AI_CONFIG   = os.path.join(PROJECT_DIR, 'ai_config.json')

# ── EŞİKLER (tek yerde) ──────────────────────────────────────────────────────
# Sahadan "bu uyarı gereksiz" geri bildirimi gelirse YALNIZ burası değişir.
ESIK = {
    'oee_kritik':        40.0,   # % — altındaki hat kritik
    'oee_uyari':         55.0,   # %
    'min_vardiya':       3,      # bir hat/operatör için anlamlı sayılacak en az vardiya
    'hurda_uyari':       3.0,    # % NOK oranı
    'hurda_kritik':      8.0,    # %
    'min_adet':          20,     # hurda/hedef analizinde anlamlı sayılacak en az adet
    'hedef_gercek':      70.0,   # % — hedefin bu kadarının altı
    'durus_pareto':      25.0,   # % — tek sebep plansız duruşun bu kadarını yiyorsa
    'durus_tekrar':      4,      # aynı sebep+hat bu kadar kez tekrarlıyorsa
    'vardiya_fark':      15.0,   # puan — gündüz/gece OEE farkı
    'operator_fark':     15.0,   # puan — operatörün kendi hattının ortalamasına farkı
    'sinyal_bosluk_dk':  45,     # dk — kayıtsız sinyal boşluğu (duruş girilmemiş olabilir)
    'uzun_vardiya_dk':   16 * 60,
}

SIDDET_SIRA = {'kritik': 0, 'uyari': 1, 'bilgi': 2}

BOLUM_AD = {
    'kaynak': 'Robot Kaynak', 'montaj': 'Montaj', 'metal': 'Metal Enjeksiyon',
    'isleme': 'İşleme', 'lazer': 'Lazer', 'pres': 'Pres Abkant',
    'plastik': 'Plastik Enjeksiyon', 'tel': 'Tel Üretimi',
}


# ═════════════════════════════════════════════════════════════════════════════
#  VERİ TOPLAMA
# ═════════════════════════════════════════════════════════════════════════════

def _baglan(yol=None):
    conn = sqlite3.connect(yol or DB_PATH, timeout=20.0)
    conn.row_factory = sqlite3.Row
    return conn


def _gecen_dk(vardiya):
    """Vardiyanın OEE paydası. AÇIK vardiyada (bugün) başlangıçtan ŞU ANA kadar —
    oee.hesapla_oee ile birebir aynı kural; iki yerde farklı hesaplanırsa panelde
    aynı vardiya iki farklı OEE gösterirdi."""
    dk = vardiya['toplam_sure_dk'] or 0
    if (vardiya['durum'] or 'kapali') != 'kapali' and (vardiya['tarih'] or '') == datetime.now().strftime('%Y-%m-%d'):
        try:
            bas = datetime.strptime((vardiya['baslangic_saati'] or '').strip()[:5], '%H:%M').time()
            simdi = datetime.now().time()
            gecen = (simdi.hour * 60 + simdi.minute) - (bas.hour * 60 + bas.minute)
            if gecen < 0:
                gecen += 1440
            return max(0, gecen)
        except ValueError:
            pass
    return dk


# ── OEE HESABINA GİRMEYEN BÖLÜMLER (kullanıcı 2026-08-25) ───────────────────
# "TK2'de genel oee hesaplama kısmına işleme bölümünü katmayalım çünkü süre
#  tanımlamaları yapılı değil" → aynısı analiz ve operatör performansı için de
#  istendi.
# Cycle süresi tanımsız üretim performansa SIFIR katkı verir ama vardiyanın
# çalışma süresi paydaya girer; sonuç fabrika/operatör OEE'sini olduğundan DÜŞÜK
# gösterir (ölçüm: Temmuz TK2 genel OEE işleme dahil %78,3 — hariç %87,6).
#
# KURAL: bölüm AÇIKÇA seçilmişse hariç tutma YAPILMAZ. Yani işleme'yi incelemek
# isteyen bölüm filtresinden seçer ve her şeyi olduğu gibi görür; hariç tutma
# yalnız BÖLÜMLER ARASI (fabrika geneli) görünümde çalışır.
#
# Süreler tanımlanınca bu listeden ÇIKARILMALI — aksi hâlde gerçek üretim
# sessizce analiz dışında kalır.
OEE_DISI_BOLUMLER = ('isleme',)


def vardiya_metrikleri(conn, bas, bit, lokasyon=None, bolum=None):
    """Dönemdeki her vardiyanın OEE bileşenleri — TOPLU sorgularla.

    oee.hesapla_oee vardiya BAŞINA ayrı bağlantı açıp ~6 sorgu atıyor; 30 günlük
    analizde bu binlerce sorgu demek. Buradaki üç toplu sorgu aynı formülleri
    uygular (availability/performance/quality tanımları oee.py ile BİREBİR).

    Cycle time çözümü SKALAR ALT-SORGU ile yapılır, LEFT JOIN ile DEĞİL:
    aynı referans kodu birden fazla satırda olabildiği için (TK1+TK2, farklı
    bölümler) JOIN üretim satırlarını çoğaltır ve adetleri KATLARDI.
    """
    sql = ("SELECT id, tarih, vardiya_turu, robot_no, operator_adi, baslangic_saati, "
           "       toplam_sure_dk, durum, COALESCE(bolum,'kaynak') AS bolum, "
           "       COALESCE(lokasyon,'TK2') AS lokasyon, "
           "       COALESCE(robotla_calisiyor,0) AS robotla_calisiyor "
           "FROM vardiyalar WHERE tarih >= ? AND tarih <= ?")
    par = [bas, bit]
    if lokasyon:
        sql += " AND COALESCE(lokasyon,'TK2') = ?"
        par.append(lokasyon)
    if bolum:
        # Bölüm AÇIKÇA seçildi → hariç tutma YOK (işleme de incelenebilsin)
        sql += " AND COALESCE(bolum,'kaynak') = ?"
        par.append(bolum)
    elif OEE_DISI_BOLUMLER:
        # Bölümler arası görünüm → cycle süresi tanımsız bölümler havuza girmez
        sql += (" AND COALESCE(bolum,'kaynak') NOT IN ("
                + ','.join('?' * len(OEE_DISI_BOLUMLER)) + ")")
        par.extend(OEE_DISI_BOLUMLER)
    vardiyalar = conn.execute(sql, par).fetchall()
    if not vardiyalar:
        return []
    vid_list = [v['id'] for v in vardiyalar]
    yer = ','.join('?' * len(vid_list))

    # Duruşlar — tip bazlı toplam + sebep dökümü
    durus_top = defaultdict(lambda: {'planli': 0, 'plansiz': 0})
    durus_ayrinti = defaultdict(list)
    for d in conn.execute(
            f"SELECT vardiya_id, durus_sebebi, COALESCE(sure_dk,0) AS sure_dk, durus_tipi, "
            f"       baslangic_saati, baslangic_ts, bitis_ts "
            f"FROM duruslar WHERE vardiya_id IN ({yer})", vid_list):
        tip = d['durus_tipi'] or durus_tipi_belirle(d['durus_sebebi'])
        durus_top[d['vardiya_id']]['planli' if tip == 'planli' else 'plansiz'] += d['sure_dk']
        durus_ayrinti[d['vardiya_id']].append(
            {'sebep': d['durus_sebebi'], 'dk': d['sure_dk'], 'tip': tip,
             'saat': d['baslangic_saati'],
             # Damgalar operatör zaman birleştirmesi için (bkz. _operator_zamani):
             # iki makine AYNI ANDA moladaysa operatör 30 dk kaybetti, 60 değil.
             'bas_ts': d['baslangic_ts'], 'bit_ts': d['bitis_ts']})

    # Üretim + cycle time çözümü.
    #
    # CYCLE TIME SATIR BAŞINA ALT-SORGU İLE ÇÖZÜLMEZ: her üretim satırı için
    # referans_listesi'ne iki korele alt-sorgu atmak 23 günlük analizde 4.9 saniye
    # sürüyordu. Onun yerine referans listesi BİR KEZ haritaya alınıp eşleme
    # Python'da yapılır (aynı öncelik: vardiyanın kendi bölüm+lokasyonu, yoksa
    # koda göre herhangi biri).
    #
    # LEFT JOIN de KULLANILMAZ: aynı referans kodu birden fazla satırda olabildiği
    # için (TK1+TK2, farklı bölümler) JOIN üretim satırlarını çoğaltır ve adetleri
    # KATLARDI — oee.py'de 2026-08-04'te düzeltilen tuzağın aynısı.
    ct_tam, ct_kod = {}, {}
    for r in conn.execute(
            "SELECT referans_kodu, hedef_cycle_time_sn AS ct, COALESCE(bolum,'kaynak') AS bolum, "
            "       COALESCE(lokasyon,'TK2') AS lokasyon FROM referans_listesi "
            "WHERE hedef_cycle_time_sn > 0"):
        anahtar = r['referans_kodu'].replace(' ', '').upper()
        ct_tam[(anahtar, r['bolum'], r['lokasyon'])] = max(
            ct_tam.get((anahtar, r['bolum'], r['lokasyon']), 0), r['ct'])
        ct_kod[anahtar] = max(ct_kod.get(anahtar, 0), r['ct'])

    v_bilgi = {v['id']: (v['bolum'], v['lokasyon']) for v in vardiyalar}
    uretim = defaultdict(list)
    for u in conn.execute(
            f"SELECT vardiya_id, referans_kodu, ok_adet, nok_adet, hedef_adet, "
            f"       cycle_time_sn, istasyon FROM uretim_kayitlari "
            f"WHERE vardiya_id IN ({yer})", vid_list):
        bol, lok = v_bilgi.get(u['vardiya_id'], ('kaynak', 'TK2'))
        anahtar = (u['referans_kodu'] or '').replace(' ', '').upper()
        guncel = ct_tam.get((anahtar, bol, lok)) or ct_kod.get(anahtar) or 0
        uretim[u['vardiya_id']].append({
            'referans_kodu': u['referans_kodu'], 'ok_adet': u['ok_adet'],
            'nok_adet': u['nok_adet'], 'hedef_adet': u['hedef_adet'],
            'cycle_time_sn': u['cycle_time_sn'], 'istasyon': u['istasyon'],
            'guncel_ct': guncel,
        })

    sonuc = []
    for v in vardiyalar:
        vid = v['id']
        toplam_dk = _gecen_dk(v)
        planli = durus_top[vid]['planli']
        plansiz = durus_top[vid]['plansiz']
        net_plan = max(0, toplam_dk - planli)
        calisma = max(0, net_plan - plansiz)
        rows = uretim.get(vid, [])
        ok = sum(r['ok_adet'] or 0 for r in rows)
        nok = sum(r['nok_adet'] or 0 for r in rows)
        hedef = sum(r['hedef_adet'] or 0 for r in rows)
        gercek_sn = 0.0
        ct_yok_adet = 0
        for r in rows:
            ct = r['guncel_ct'] if (r['guncel_ct'] and r['guncel_ct'] > 0) else (r['cycle_time_sn'] or 0)
            adet = (r['ok_adet'] or 0) + (r['nok_adet'] or 0)
            if ct > 0:
                gercek_sn += adet * ct
            else:
                ct_yok_adet += adet          # süresi tanımsız → performansa HİÇ katkı vermez
        a = (calisma / net_plan) if net_plan > 0 else 0
        p = (gercek_sn / (calisma * 60)) if calisma > 0 else 0
        q = (ok / (ok + nok)) if (ok + nok) > 0 else 0
        sonuc.append({
            'vardiya_id': vid, 'tarih': v['tarih'], 'vardiya_turu': v['vardiya_turu'],
            'operator': v['operator_adi'], 'hat': v['robot_no'],
            'bolum': v['bolum'], 'lokasyon': v['lokasyon'], 'durum': v['durum'] or 'kapali',
            'baslangic': v['baslangic_saati'],
            # 'Robotla Çalışıyor' (metal, tam otomasyon): operatör bu vardiyada
            # makineye sürekli bağlı DEĞİL — aynı anda başka makineye de bakıyor
            # olabilir. Operatör toplamında süre bir kez sayılır.
            'robotlu': int(v['robotla_calisiyor'] or 0) == 1,
            'toplam_dk': toplam_dk, 'planli_dk': planli, 'plansiz_dk': plansiz,
            'net_plan_dk': net_plan, 'calisma_dk': calisma,
            'ok': ok, 'nok': nok, 'toplam': ok + nok, 'hedef': hedef,
            'gercek_uretim_sn': gercek_sn, 'ct_yok_adet': ct_yok_adet,
            'availability': round(a * 100, 1), 'performance': round(p * 100, 1),
            'quality': round(q * 100, 1), 'oee': round(a * min(p, 1.0) * q * 100, 1),
            'duruslar': durus_ayrinti.get(vid, []),
            'kayitlar': rows,
        })
    return sonuc


def _topla(vardiyalar):
    """Vardiya kümesini TEK bir OEE'ye indirger.

    Yüzdelerin ortalaması ALINMAZ — bileşenler toplanıp formül yeniden uygulanır.
    Ortalama alınsaydı 10 dakikalık bir vardiya 8 saatlik vardiyayla aynı ağırlığı
    taşır ve tek bir kısa kayıt tüm tabloyu yanıltırdı."""
    net = sum(v['net_plan_dk'] for v in vardiyalar)
    cal = sum(v['calisma_dk'] for v in vardiyalar)
    # ── ÜRETİM SÜRESİ VARDİYA BAŞINA KIRPILIR (kullanıcı 2026-08-25) ────────
    # Bir vardiya kendi çalışma süresinden FAZLASINI üretmiş sayılamaz. Kırpma
    # eskiden yalnız EN SONDA (min(P,1)) yapılıyordu; havuzda bu ÇOK GEÇ:
    # cycle time'ı cömert tanımlanmış bir vardiya %199 performansla katkı verip,
    # cycle time'ı TANIMSIZ olduğu için 0 sn üreten vardiyaları SÜBVANSE ediyordu.
    # Ölçüm (2026-07 TK2, 358 vardiya): 123'ü %100 üstü (ort. %199), 97'si %0 →
    # havuz OEE 93.0 çıkıyordu; vardiya başına kırpınca 66.4. Kullanıcı "%98
    # çıkıyor, o kadar yüksek olamaz" derken tam bu sapmayı görüyordu.
    # TEK VARDİYADA fark yok (min(P,1) ile aynı) — sapma yalnız TOPLAMDA oluşur.
    sn  = sum(min(v['gercek_uretim_sn'], v['calisma_dk'] * 60) for v in vardiyalar)
    ok  = sum(v['ok'] for v in vardiyalar)
    nok = sum(v['nok'] for v in vardiyalar)
    a = (cal / net) if net > 0 else 0
    p = (sn / (cal * 60)) if cal > 0 else 0
    q = (ok / (ok + nok)) if (ok + nok) > 0 else 0
    ad_top = sum(v['toplam'] for v in vardiyalar)
    ad_yok = sum(v['ct_yok_adet'] for v in vardiyalar)
    return {
        'vardiya': len(vardiyalar),
        'toplam_dk': sum(v['toplam_dk'] for v in vardiyalar),
        'planli_dk': sum(v['planli_dk'] for v in vardiyalar),
        'plansiz_dk': sum(v['plansiz_dk'] for v in vardiyalar),
        'net_plan_dk': net, 'calisma_dk': cal,
        'gercek_uretim_sn': sn,
        'ok': ok, 'nok': nok, 'toplam': ok + nok,
        'hedef': sum(v['hedef'] for v in vardiyalar),
        'ct_yok_adet': sum(v['ct_yok_adet'] for v in vardiyalar),
        'availability': round(a * 100, 1), 'performance': round(p * 100, 1),
        'quality': round(q * 100, 1), 'oee': round(a * min(p, 1.0) * q * 100, 1),
        'hurda_oran': round((nok / (ok + nok) * 100), 2) if (ok + nok) > 0 else 0.0,
        # ÖLÇÜM KAPSAMI: üretimin yüzde kaçının cycle time'ı tanımlı. Düşükse
        # performans (dolayısıyla OEE) YANILTICIDIR — tanımsız üretim performansa
        # hiç katkı vermez, OEE olduğundan düşük çıkar. Paneller bunu OEE'nin
        # yanında gösterir ki rakama ne kadar güvenileceği görünsün.
        'olcum_kapsami': round((ad_top - ad_yok) / ad_top * 100, 1) if ad_top > 0 else 0.0,
    }


# ═════════════════════════════════════════════════════════════════════════════
#  OPERATÖR PERFORMANSI
# ═════════════════════════════════════════════════════════════════════════════

# ── 1 OPERATÖR ÷ 2 MAKİNE (kullanıcı 2026-08-21) ────────────────────────────
# "Metal enjeksiyonda 1 operatör iki makineye birden bakıyorsa (makineler robotla
#  parça alma yapıyorsa bu mümkün) OEE hesaplaması 2 operatör 2 makine gibi
#  hesaplanmamalı — OEE düşük görünüyor."
#
# SORUN: operatör iki makine için İKİ vardiya açıyor. Operatör satırında bu iki
# vardiyanın süresi TOPLANIYORDU: 07:00-16:00 çalışan kişi 18 saat görünüyor,
# dolayısıyla adet/saat, performans ve OEE YARIYA iniyordu. Üretim iki katına
# çıkmasına rağmen kişi tabloda en alta düşüyordu.
#
# ÇÖZÜM: operatör satırında SÜRE bir kez sayılır — vardiya pencereleri toplanmaz,
# BİRLEŞTİRİLİR (union). Union hem eşzamanlı hem ardışık durumu doğru çözer:
#   eşzamanlı 07-16 + 07-16 → 9 saat (bir kez)
#   ardışık   07-11 + 11-16 → 9 saat (toplamla aynı)
# Üretim (ok/nok/süre) TOPLANIR — gerçekten iki katı üretim çıkmıştır.
#
# KAPSAM: yalnız 'Robotla Çalışıyor' işaretli vardiyalar (kullanıcı kararı —
# sinyal mevcut anahtardan gelir). İşaretsiz vardiyalar eskisi gibi toplanır;
# yani hiçbir mevcut bölümün sayısı kendiliğinden değişmez.
#
# MAKİNE (hat) TARAFI DEĞİŞMEZ: her makinenin kendi OEE'si aynen durur — makine
# performansını ölçen tek şey odur (kullanıcı kararı).


def _birlesik_dk(araliklar):
    """Çakışan (baslangic_dk, bitis_dk) aralıklarının BİRLEŞİMİNİN uzunluğu."""
    temiz = [(b, s) for b, s in araliklar if b is not None and s is not None and s > b]
    if not temiz:
        return 0
    temiz.sort()
    toplam = 0
    cb, cs = temiz[0]
    for b, s in temiz[1:]:
        if b > cs:
            toplam += cs - cb
            cb, cs = b, s
        elif s > cs:
            cs = s
    return toplam + (cs - cb)


def _dk_no(ts):
    """'YYYY-MM-DD HH:MM:SS' → mutlak dakika (aralık karşılaştırması için)."""
    if not ts:
        return None
    try:
        d = datetime.strptime(str(ts)[:16].replace('T', ' '), '%Y-%m-%d %H:%M')
    except ValueError:
        return None
    return int(d.timestamp() // 60)


def _vardiya_penceresi(v):
    """Vardiyanın (baslangic_dk, bitis_dk) penceresi; saat bozuksa None."""
    bas = _dk_no(f"{v['tarih']} {str(v.get('baslangic') or '')[:5]}")
    if bas is None:
        return None
    return (bas, bas + (v['toplam_dk'] or 0))


def _cakisan_gruplar(gunun_vardiyalari):
    """Aynı günün vardiyalarını ZAMANDA ÇAKIŞANLAR birlikte olacak şekilde gruplar.

    Penceresi çözülemeyen (saati bozuk) vardiya kendi başına bir gruptur — eski
    davranışa, yani toplamaya düşer."""
    pencereli, tekil = [], []
    for v in gunun_vardiyalari:
        p = _vardiya_penceresi(v)
        (pencereli if p else tekil).append((v, p))
    pencereli.sort(key=lambda x: x[1][0])
    gruplar = [[v] for v, _ in tekil]
    aktif, son = [], None
    for v, (b, s2) in pencereli:
        if aktif and b < son:            # önceki grupla çakışıyor
            aktif.append(v)
            son = max(son, s2)
        else:
            if aktif:
                gruplar.append(aktif)
            aktif, son = [v], s2
    if aktif:
        gruplar.append(aktif)
    return gruplar


def _grup_zamani(grup):
    """Çakışan vardiya grubunun BİRLEŞİK (toplam, planlı, plansız) dakikası."""
    pencereler = [_vardiya_penceresi(v) for v in grup]
    toplam = _birlesik_dk([p for p in pencereler if p])
    toplam += sum(v['toplam_dk'] for v, p in zip(grup, pencereler) if not p)
    sonuc = {}
    for tip in ('planli', 'plansiz'):
        araliklar, damgasiz = [], 0
        for v in grup:
            for d in v.get('duruslar', []):
                if d.get('tip') != tip:
                    continue
                b, e = _dk_no(d.get('bas_ts')), _dk_no(d.get('bit_ts'))
                if b is None or e is None:
                    damgasiz += d.get('dk', 0)   # damgasız kayıt → eski davranış
                else:
                    araliklar.append((b, e))
        sonuc[tip] = _birlesik_dk(araliklar) + damgasiz
    return toplam, sonuc['planli'], sonuc['plansiz']


def _operator_zamani(vs):
    """Operatörün GERÇEK süresi: (toplam_dk, planli_dk, plansiz_dk).

    BİRLEŞTİRME KOŞULU (2026-08-21, saha verisine göre): aynı gün ZAMANDA
    ÇAKIŞAN en az iki vardiya + bunlardan EN AZ BİRİ 'Robotla Çalışıyor'.
    "Hepsi robotlu olsun" demek YANLIŞ olurdu: operatör anahtarı yalnız robotun
    parçayı aldığı makinede işaretliyor, fiziksel olarak başında durduğu ikinci
    makinede işaretlemiyor. Gerçek kayıt (Serdar Göcük, 2026-07-22): 300T
    07:03-16:23 robotlu=1 + 550T 08:12-16:23 robotlu=0 → toplasaydık 1051 dk
    (17,5 saat) çıkardı; birleşince 560 dk (gerçek gün).

    Çakışmayan vardiyalar (kişi sabah bir makinede, öğleden sonra başkasında)
    TOPLANIR — orada süre gerçekten arka arkayadır."""
    toplam = planli = plansiz = 0
    gunler = defaultdict(list)
    for v in vs:
        gunler[v.get('tarih')].append(v)
    for gunun in gunler.values():
        for grup in _cakisan_gruplar(gunun):
            if len(grup) >= 2 and any(v.get('robotlu') for v in grup):
                t, pl, ps = _grup_zamani(grup)
            else:
                t = sum(v['toplam_dk'] for v in grup)
                pl = sum(v['planli_dk'] for v in grup)
                ps = sum(v['plansiz_dk'] for v in grup)
            toplam += t
            planli += pl
            plansiz += ps
    return toplam, planli, plansiz


def _birlesen_vardiya_sayisi(vs):
    """Süresi birleştirilen vardiya sayısı (tabloda 'çoklu makine' rozeti için)."""
    n = 0
    gunler = defaultdict(list)
    for v in vs:
        gunler[v.get('tarih')].append(v)
    for gunun in gunler.values():
        for grup in _cakisan_gruplar(gunun):
            if len(grup) >= 2 and any(v.get('robotlu') for v in grup):
                n += len(grup)
    return n


def _topla_operator(vs):
    """_topla'nın operatör sürümü: ÜRETİM toplanır, SÜRE birleştirilir.

    _topla DEĞİŞTİRİLMEDİ — o hat/makine tarafında kullanılıyor ve orada süre
    gerçekten toplanmalı (iki makine iki ayrı kapasitedir)."""
    t = _topla(vs)
    toplam_dk, planli_dk, plansiz_dk = _operator_zamani(vs)
    net = max(0, toplam_dk - planli_dk)
    cal = max(0, net - plansiz_dk)
    sn = t['gercek_uretim_sn'] if 'gercek_uretim_sn' in t else sum(v['gercek_uretim_sn'] for v in vs)
    ok, nok = t['ok'], t['nok']
    a = (cal / net) if net > 0 else 0
    p = (sn / (cal * 60)) if cal > 0 else 0
    q = (ok / (ok + nok)) if (ok + nok) > 0 else 0
    t.update({
        'toplam_dk': toplam_dk, 'planli_dk': planli_dk, 'plansiz_dk': plansiz_dk,
        'net_plan_dk': net, 'calisma_dk': cal,
        'availability': round(a * 100, 1), 'performance': round(p * 100, 1),
        'quality': round(q * 100, 1), 'oee': round(a * min(p, 1.0) * q * 100, 1),
        # Kaç vardiyanın süresi birleştirildi — tabloda "2 makine" rozeti için
        'paralel_vardiya': _birlesen_vardiya_sayisi(vs),
    })
    return t


def operator_performans(vardiyalar):
    """Operatör bazlı performans tablosu.

    ADİL KARŞILAŞTIRMA NOTU: OEE hat/makineye çok bağlıdır (yavaş bir makinede
    çalışan operatör düşük görünür). Bu yüzden her operatör için 'hat_farki' de
    hesaplanır: kendi çalıştığı hatların o dönemdeki ortalamasına göre kaç puan
    yukarıda/aşağıda olduğu. Kişileri sıralarken bakılması gereken sütun budur."""
    hat_grup = defaultdict(list)
    for v in vardiyalar:
        hat_grup[(v['lokasyon'], v['bolum'], v['hat'])].append(v)
    hat_oee = {k: _topla(vs)['oee'] for k, vs in hat_grup.items()}

    op_grup = defaultdict(list)
    for v in vardiyalar:
        if v['operator']:
            op_grup[v['operator']].append(v)

    tablo = []
    for op, vs in op_grup.items():
        # OPERATÖR toplamı: robot destekli paralel vardiyalarda süre BİR KEZ
        # sayılır (bkz. _topla_operator). Hat tablosu _topla kullanmaya devam eder.
        t = _topla_operator(vs)
        # Operatörün hat karışımına göre BEKLENEN OEE (hat ortalamalarının,
        # o hatta geçirdiği net süreyle ağırlıklı ortalaması)
        agirlik, beklenen = 0.0, 0.0
        for v in vs:
            k = (v['lokasyon'], v['bolum'], v['hat'])
            w = v['net_plan_dk'] or 0
            agirlik += w
            beklenen += hat_oee.get(k, 0) * w
        beklenen = (beklenen / agirlik) if agirlik > 0 else 0
        hatlar = sorted({v['hat'] for v in vs if v['hat']})
        tablo.append({
            'operator': op,
            'bolumler': sorted({v['bolum'] for v in vs}),
            'lokasyonlar': sorted({v['lokasyon'] for v in vs}),
            'hatlar': hatlar,
            'hat_sayisi': len(hatlar),
            'gun': len({v['tarih'] for v in vs}),
            **t,
            'saatlik': round(t['toplam'] / (t['calisma_dk'] / 60), 1) if t['calisma_dk'] > 0 else 0,
            'hedef_gercek': round(t['toplam'] / t['hedef'] * 100, 1) if t['hedef'] > 0 else None,
            'beklenen_oee': round(beklenen, 1),
            'hat_farki': round(t['oee'] - beklenen, 1),
        })
    tablo.sort(key=lambda r: (-r['oee'], -r['toplam']))
    return tablo


def hat_performans(vardiyalar):
    """Hat/makine bazlı özet — operatör tablosunun karşılığı."""
    grup = defaultdict(list)
    for v in vardiyalar:
        grup[(v['lokasyon'], v['bolum'], v['hat'] or '—')].append(v)
    tablo = []
    for (lok, bol, hat), vs in grup.items():
        t = _topla(vs)
        tablo.append({
            'lokasyon': lok, 'bolum': bol, 'hat': hat,
            'operatorler': sorted({v['operator'] for v in vs if v['operator']}),
            'gun': len({v['tarih'] for v in vs}),
            **t,
        })
    tablo.sort(key=lambda r: (-r['toplam']))
    return tablo


# ═════════════════════════════════════════════════════════════════════════════
#  BULGULAR
# ═════════════════════════════════════════════════════════════════════════════

def _bin(n):
    """Binlik ayraçlı sayı: 25038 → '25.038'.

    f'{n:,}'.replace(',', '.') KALIBI KULLANILMAZ: replace metnin TAMAMINA
    uygulandığı için cümledeki normal virgülleri de noktaya çevirir
    ("HİÇ girmiyor. dolayısıyla" — 2026-08-21'de mailde yaşandı)."""
    return '{:,}'.format(int(n)).replace(',', '.')


def _b(kod, siddet, baslik, detay, oneri='', **sayisal):
    return {'kod': kod, 'siddet': siddet, 'baslik': baslik, 'detay': detay,
            'oneri': oneri, 'sayisal': sayisal}


def ct_yok_kirilimi(vardiyalar):
    """Hedef süresi TANIMSIZ üretimin referans kırılımı — adede göre azalan.

    Kullanıcı 2026-08-26: "üretimin %32'sinde süre tanımlı değil uyarısı verip
    bundan dolayı OEE düşük diyor, tıklayıp bu kodları görebileyim."
    Uyarı doğruydu ama eyleme dönüşmüyordu: hangi kodların doldurulacağı
    görünmeden 'referans listesini gözden geçirin' demek işe yaramaz.

    Anahtar (kod, bölüm, lokasyon): aynı kod iki tesiste/bölümde ayrı referans
    satırıdır ve süresi ayrı ayrı doldurulur — tek satırda toplanırsa hangisinin
    eksik olduğu kaybolur. 'hatlar' ve 'operatorler' listeleri kimin nerede
    ürettiğini gösterir (süreyi kim tanımlayacak sorusu buradan cevaplanır)."""
    kirilim = {}
    for v in vardiyalar:
        for r in v['kayitlar']:
            ct = r['guncel_ct'] if (r['guncel_ct'] and r['guncel_ct'] > 0) else (r['cycle_time_sn'] or 0)
            if ct > 0:
                continue
            adet = (r['ok_adet'] or 0) + (r['nok_adet'] or 0)
            if adet <= 0:
                continue          # adetsiz satır 'eksik tanım' listesini şişirmesin
            kod = (r['referans_kodu'] or '').strip() or '—'
            anahtar = (kod, v['bolum'] or '', v['lokasyon'] or '')
            d = kirilim.get(anahtar)
            if d is None:
                d = kirilim[anahtar] = {
                    'referans': kod, 'bolum': v['bolum'] or '',
                    'lokasyon': v['lokasyon'] or '', 'adet': 0, 'kayit_sayisi': 0,
                    '_hatlar': set(), '_operatorler': set(), '_gunler': set()}
            d['adet'] += adet
            d['kayit_sayisi'] += 1
            d['_hatlar'].add(v['hat'] or '')
            d['_operatorler'].add(v['operator'] or '')
            d['_gunler'].add(v['tarih'])
    out = []
    for d in kirilim.values():
        out.append({**{k: x for k, x in d.items() if not k.startswith('_')},
                    'hatlar': sorted(h for h in d['_hatlar'] if h),
                    'operatorler': sorted(o for o in d['_operatorler'] if o),
                    'gun_sayisi': len(d['_gunler'])})
    out.sort(key=lambda x: (-x['adet'], x['referans']))
    return out


def _bulgu_cycle_tanimsiz(vardiyalar, hatlar):
    """SÜRESİ TANIMSIZ REFERANSLAR — en sinsi bulgu.

    Performans = üretilen adet × cycle time / çalışma süresi. Cycle time 0 ise o
    adet performansa HİÇ katkı vermez: operatör tam kapasite çalışsa bile OEE
    düşük çıkar. 'Hattın OEE'si düşük' diye bakılıp saatler harcanabilir; sebep
    üretimde değil, referans listesindeki boş süre alanındadır. Bu yüzden diğer
    OEE bulgularından ÖNCE raporlanır."""
    ct_yok = sum(v['ct_yok_adet'] for v in vardiyalar)
    toplam = sum(v['toplam'] for v in vardiyalar)
    if ct_yok <= 0 or toplam <= 0:
        return []
    oran = ct_yok / toplam * 100
    if oran < 5:
        return []
    # Kırılım TEK YERDEN (ct_yok_kirilimi): panelin 'tıkla ve kodları gör'
    # listesiyle bu bulgunun listesi aynı kaynaktan gelsin, zamanla ayrışmasın.
    kirilim = ct_yok_kirilimi(vardiyalar)
    kodlar = {(k['referans'], k['bolum'], k['lokasyon']): k['adet'] for k in kirilim}
    ilk = [(k['referans'], k['adet']) for k in kirilim[:8]]
    return [_b('cycle_tanimsiz', 'kritik' if oran >= 20 else 'uyari',
               f'Üretimin %{oran:.0f}\'ında hedef süre tanımsız — OEE olduğundan düşük görünüyor',
               f'{_bin(ct_yok)} adet ({len(kodlar)} referans) hedef cycle time olmadan '
               f'kaydedilmiş. Bu adetler performans hesabına HİÇ girmiyor, dolayısıyla '
               f'bu hatların OEE\'si gerçekte olduğundan düşük çıkıyor.',
               'Referans listesinde bu kodların hedef süresini doldurun; OEE kendiliğinden düzelir.',
               adet=ct_yok, oran=round(oran, 1), referans_sayisi=len(kodlar),
               referanslar=[{'kod': k, 'adet': a} for k, a in ilk])]


def _bulgu_oee(hatlar, genel):
    """Dönem ortalamasının belirgin altında kalan hatlar."""
    out = []
    for h in hatlar:
        if h['vardiya'] < ESIK['min_vardiya'] or h['net_plan_dk'] <= 0:
            continue
        if h['oee'] >= ESIK['oee_uyari']:
            continue
        # En zayıf bileşeni söyle — "OEE düşük" tek başına yön göstermez
        bilesenler = {'Kullanılabilirlik (duruş)': h['availability'],
                      'Performans (hız)': min(h['performance'], 100.0),
                      'Kalite (hurda)': h['quality']}
        zayif = min(bilesenler, key=bilesenler.get)
        out.append(_b('oee_dusuk',
                      'kritik' if h['oee'] < ESIK['oee_kritik'] else 'uyari',
                      f"{h['hat']} ({BOLUM_AD.get(h['bolum'], h['bolum'])} · {h['lokasyon']}) OEE %{h['oee']}",
                      f"{h['vardiya']} vardiyada OEE %{h['oee']} — dönem geneli %{genel['oee']}. "
                      f"En zayıf bileşen: {zayif} %{bilesenler[zayif]:.0f}. "
                      f"Plansız duruş {h['plansiz_dk']} dk, üretim {h['toplam']} adet.",
                      f"{zayif} tarafına bakın." if zayif else '',
                      hat=h['hat'], bolum=h['bolum'], lokasyon=h['lokasyon'],
                      oee=h['oee'], genel_oee=genel['oee'], vardiya=h['vardiya'],
                      availability=h['availability'], performance=h['performance'],
                      quality=h['quality'], zayif_bilesen=zayif))
    out.sort(key=lambda x: x['sayisal']['oee'])
    return out[:8]


def _bulgu_durus(vardiyalar):
    """Duruş Pareto'su + tekrar eden sebep-hat çiftleri."""
    out = []
    sebep_dk = defaultdict(float)
    sebep_adet = defaultdict(int)
    cift = defaultdict(lambda: {'dk': 0.0, 'adet': 0})
    toplam_plansiz = 0.0
    for v in vardiyalar:
        for d in v['duruslar']:
            if d['tip'] == 'planli':
                continue
            toplam_plansiz += d['dk']
            sebep_dk[d['sebep']] += d['dk']
            sebep_adet[d['sebep']] += 1
            cift[(d['sebep'], v['hat'], v['bolum'])]['dk'] += d['dk']
            cift[(d['sebep'], v['hat'], v['bolum'])]['adet'] += 1
    if toplam_plansiz <= 0:
        return out

    for sebep, dk in sorted(sebep_dk.items(), key=lambda x: -x[1])[:3]:
        oran = dk / toplam_plansiz * 100
        if oran < ESIK['durus_pareto']:
            continue
        out.append(_b('durus_pareto', 'uyari',
                      f'Plansız duruşun %{oran:.0f}\'ı tek sebepten: {sebep}',
                      f'{sebep_adet[sebep]} kez, toplam {dk:.0f} dk '
                      f'({dk/60:.1f} saat). Dönemdeki tüm plansız duruş {toplam_plansiz:.0f} dk.',
                      'Tek sebebe yoğunlaşmak toplam duruşun dörtte birinden fazlasını geri kazandırır.',
                      sebep=sebep, dk=round(dk), oran=round(oran, 1), tekrar=sebep_adet[sebep]))

    for (sebep, hat, bol), d in sorted(cift.items(), key=lambda x: -x[1]['dk'])[:5]:
        if d['adet'] < ESIK['durus_tekrar']:
            continue
        out.append(_b('durus_tekrar', 'uyari',
                      f'{hat}: "{sebep}" {d["adet"]} kez tekrarladı',
                      f'{BOLUM_AD.get(bol, bol)} · {hat} hattında aynı sebep {d["adet"]} kez, '
                      f'toplam {d["dk"]:.0f} dk. Tekrarlayan duruş tek seferlik arızadan farklıdır — '
                      f'kök sebep giderilmemiş olabilir.',
                      'Kök neden analizi yapın; tekil arıza gibi kapatılmasın.',
                      sebep=sebep, hat=hat, bolum=bol, tekrar=d['adet'], dk=round(d['dk'])))
    return out


def _bulgu_kalite(vardiyalar):
    """Hurda oranı yüksek referans ve hatlar."""
    out = []
    ref = defaultdict(lambda: {'ok': 0, 'nok': 0})
    hat = defaultdict(lambda: {'ok': 0, 'nok': 0})
    for v in vardiyalar:
        for r in v['kayitlar']:
            ref[r['referans_kodu']]['ok'] += r['ok_adet'] or 0
            ref[r['referans_kodu']]['nok'] += r['nok_adet'] or 0
            hat[(v['hat'], v['bolum'])]['ok'] += r['ok_adet'] or 0
            hat[(v['hat'], v['bolum'])]['nok'] += r['nok_adet'] or 0

    def _tara(d, etiket_fn, kod):
        sonuc = []
        for k, s in d.items():
            top = s['ok'] + s['nok']
            if top < ESIK['min_adet'] or s['nok'] == 0:
                continue
            oran = s['nok'] / top * 100
            if oran < ESIK['hurda_uyari']:
                continue
            sonuc.append((oran, k, s, top))
        sonuc.sort(key=lambda x: -x[0])
        return [_b(kod, 'kritik' if o >= ESIK['hurda_kritik'] else 'uyari',
                   f'{etiket_fn(k)} hurda oranı %{o:.1f}',
                   f'{s["nok"]} hatalı / {top} toplam. Dönem hedefi %{ESIK["hurda_uyari"]} altı.',
                   'Ölçüm/ayar kontrolü ve ilk parça onayı gözden geçirilmeli.',
                   ad=etiket_fn(k), nok=s['nok'], toplam=top, oran=round(o, 2))
                for o, k, s, top in sonuc[:5]]

    out += _tara(ref, lambda k: k, 'hurda_referans')
    out += _tara(hat, lambda k: f'{k[0]} ({BOLUM_AD.get(k[1], k[1])})', 'hurda_hat')
    return out


def _bulgu_hedef(vardiyalar):
    """Hedef girilmiş ama tutturulamamış işler."""
    ref = defaultdict(lambda: {'uretim': 0, 'hedef': 0, 'vardiya': 0})
    for v in vardiyalar:
        for r in v['kayitlar']:
            if (r['hedef_adet'] or 0) <= 0:
                continue
            ref[r['referans_kodu']]['uretim'] += (r['ok_adet'] or 0) + (r['nok_adet'] or 0)
            ref[r['referans_kodu']]['hedef'] += r['hedef_adet']
            ref[r['referans_kodu']]['vardiya'] += 1
    out = []
    for k, s in sorted(ref.items(), key=lambda x: x[1]['uretim'] / max(x[1]['hedef'], 1)):
        if s['hedef'] < ESIK['min_adet']:
            continue
        oran = s['uretim'] / s['hedef'] * 100
        if oran >= ESIK['hedef_gercek']:
            continue
        out.append(_b('hedef_alti', 'uyari',
                      f'{k}: hedefin %{oran:.0f}\'ı yapıldı',
                      f'{s["vardiya"]} kayıtta hedef {s["hedef"]}, gerçekleşen {s["uretim"]}. '
                      f'Eksik {s["hedef"] - s["uretim"]} adet.',
                      'Hedef gerçekçi değilse güncelleyin; değilse darboğazı arayın.',
                      referans=k, hedef=s['hedef'], uretim=s['uretim'], oran=round(oran, 1)))
        if len(out) >= 5:
            break
    return out


def _bulgu_vardiya_farki(vardiyalar):
    """Gündüz / gece farkı — makine aynıysa fark yönetilebilir bir şeydir."""
    grup = defaultdict(list)
    for v in vardiyalar:
        if v['vardiya_turu']:
            grup[v['vardiya_turu']].append(v)
    if len(grup) < 2:
        return []
    ozet = {k: _topla(vs) for k, vs in grup.items() if _topla(vs)['vardiya'] >= ESIK['min_vardiya']}
    if len(ozet) < 2:
        return []
    en_iyi = max(ozet, key=lambda k: ozet[k]['oee'])
    en_kotu = min(ozet, key=lambda k: ozet[k]['oee'])
    fark = ozet[en_iyi]['oee'] - ozet[en_kotu]['oee']
    if fark < ESIK['vardiya_fark']:
        return []
    return [_b('vardiya_farki', 'uyari',
               f'{en_kotu} vardiyası {en_iyi} vardiyasından {fark:.0f} puan geride',
               f'{en_iyi}: OEE %{ozet[en_iyi]["oee"]} ({ozet[en_iyi]["vardiya"]} vardiya) · '
               f'{en_kotu}: OEE %{ozet[en_kotu]["oee"]} ({ozet[en_kotu]["vardiya"]} vardiya). '
               f'Makineler aynı olduğuna göre fark yöntem, yetkinlik veya vardiya '
               f'başlangıç/bitiş kayıplarından geliyor.',
               'İki vardiyanın duruş dökümünü yan yana koyun.',
               iyi=en_iyi, kotu=en_kotu, fark=round(fark, 1),
               iyi_oee=ozet[en_iyi]['oee'], kotu_oee=ozet[en_kotu]['oee'])]


def _bulgu_operator(op_tablo):
    """Kendi hattının ortalamasından belirgin sapan operatörler (iki yönlü).

    Ham OEE ile sıralamak haksızdır — yavaş makinedeki kişi hep sonda çıkar.
    Bu yüzden ölçüt 'hat_farki': aynı hatlarda beklenen OEE ile arasındaki puan."""
    out = []
    for o in op_tablo:
        if o['vardiya'] < ESIK['min_vardiya'] or o['net_plan_dk'] <= 0:
            continue
        if o['hat_farki'] <= -ESIK['operator_fark']:
            out.append(_b('operator_dusuk', 'uyari',
                          f"{o['operator']} kendi hatlarının {abs(o['hat_farki']):.0f} puan altında",
                          f"OEE %{o['oee']}, aynı hatlarda beklenen %{o['beklenen_oee']} "
                          f"({o['vardiya']} vardiya, {o['toplam']} adet). "
                          f"Karşılaştırma kişinin çalıştığı hatlara göre düzeltilmiştir.",
                          'Eğitim/yöntem farkı olabilir — duruş kayıtlarına birlikte bakın.',
                          operator=o['operator'], oee=o['oee'], beklenen=o['beklenen_oee'],
                          fark=o['hat_farki'], vardiya=o['vardiya']))
        elif o['hat_farki'] >= ESIK['operator_fark']:
            out.append(_b('operator_yuksek', 'bilgi',
                          f"{o['operator']} kendi hatlarının {o['hat_farki']:.0f} puan üstünde",
                          f"OEE %{o['oee']}, aynı hatlarda beklenen %{o['beklenen_oee']} "
                          f"({o['vardiya']} vardiya). Yönteminin yayılması diğer vardiyaları da yukarı çeker.",
                          'Ne farklı yaptığını sorun — iyi uygulama olarak yayılabilir.',
                          operator=o['operator'], oee=o['oee'], beklenen=o['beklenen_oee'],
                          fark=o['hat_farki'], vardiya=o['vardiya']))
    out.sort(key=lambda x: x['sayisal']['fark'])
    return out[:6]


def _bulgu_veri_kalitesi(conn, vardiyalar, bas, bit):
    """İnsanın gözden kaçırdığı KAYIT hataları — analiz güvenilirliğinin ön koşulu."""
    out = []
    bugun = datetime.now().strftime('%Y-%m-%d')

    acik = [v for v in vardiyalar if v['durum'] != 'kapali' and v['tarih'] < bugun]
    if acik:
        ornek = sorted(acik, key=lambda v: v['tarih'])[:5]
        out.append(_b('acik_vardiya', 'uyari',
                      f'{len(acik)} vardiya kapatılmadan kalmış',
                      'Kapatılmayan vardiyanın süresi ve OEE\'si güvenilir değildir; '
                      'sayaç da kapanışta dondurulmadığı için adet bozulabilir. '
                      + ', '.join(f"{v['tarih']} {v['operator']} ({v['hat']})" for v in ornek),
                      'Panelden kapatın; tekrarlıyorsa operatöre hatırlatın.',
                      sayi=len(acik),
                      ornekler=[{'tarih': v['tarih'], 'operator': v['operator'],
                                 'hat': v['hat']} for v in ornek]))

    uzun = [v for v in vardiyalar if v['toplam_dk'] > ESIK['uzun_vardiya_dk']]
    if uzun:
        out.append(_b('uzun_vardiya', 'bilgi',
                      f'{len(uzun)} vardiya {ESIK["uzun_vardiya_dk"]//60} saatten uzun görünüyor',
                      'Muhtemelen kapanış saati geç girilmiş. Uzun süre OEE paydasını büyütür → '
                      'o vardiyanın (ve ait olduğu operatörün) OEE\'si olduğundan düşük çıkar.',
                      'Saatleri düzeltin; bu kayıtlar ortalamayı aşağı çekiyor.',
                      sayi=len(uzun),
                      ornekler=[{'tarih': v['tarih'], 'operator': v['operator'],
                                 'dk': v['toplam_dk']} for v in sorted(uzun, key=lambda x: -x['toplam_dk'])[:5]]))

    # Üretimi olan ama HİÇ duruş kaydı olmayan uzun vardiyalar: mola bile
    # girilmemişse duruş disiplini yok demektir, availability yanıltıcı çıkar.
    durussuz = [v for v in vardiyalar
                if v['toplam_dk'] >= 240 and not v['duruslar'] and v['toplam'] > 0]
    if len(durussuz) >= 3:
        out.append(_b('durus_girilmemis', 'uyari',
                      f'{len(durussuz)} uzun vardiyada hiç duruş kaydı yok',
                      '4 saatten uzun bir vardiyada mola dahil hiçbir duruş girilmemişse '
                      'kullanılabilirlik %100 görünür ve gerçek kayıplar raporda kaybolur.',
                      'Duruş girişini hatırlatın — OEE\'nin en büyük kaybı burada saklanıyor.',
                      sayi=len(durussuz),
                      ornekler=[{'tarih': v['tarih'], 'operator': v['operator'],
                                 'hat': v['hat'], 'dk': v['toplam_dk']} for v in durussuz[:5]]))

    # Üretimi 0 olan kapalı vardiyalar
    sifir = [v for v in vardiyalar if v['durum'] == 'kapali' and v['toplam'] == 0 and v['toplam_dk'] >= 120]
    if len(sifir) >= 3:
        out.append(_b('uretimsiz_vardiya', 'uyari',
                      f'{len(sifir)} vardiyada hiç üretim kaydı yok',
                      '2 saatten uzun sürmüş ama tek adet girilmemiş vardiyalar. Ya üretim '
                      'gerçekten olmadı (duruş olarak kaydedilmeli) ya da adet girilmeyi unuttu.',
                      'Sebebi belirlenmeli — bu vardiyalar OEE ortalamasını sıfıra doğru çekiyor.',
                      sayi=len(sifir),
                      ornekler=[{'tarih': v['tarih'], 'operator': v['operator'],
                                 'hat': v['hat'], 'dk': v['toplam_dk']} for v in sifir[:5]]))
    return out


def _bulgu_sinyal(vardiyalar, bas, bit):
    """SENSÖR ile KAYIT karşılaştırması — 'insanın yakalayamayacağı' kısım.

    Sayaç modülü olan bir hatta sinyal saatlerce gelmemişse makine durmuştur.
    O aralıkta duruş kaydı YOKSA kayıp görünmez olur: operatör vardiyayı normal
    kapatır, OEE yüksek çıkar, kimse fark etmez. Burada iki veri kaynağı
    (pilot.db sinyalleri ↔ uretim.db duruş kayıtları) karşılaştırılır."""
    if not os.path.exists(PILOT_DB):
        return []
    out = []
    try:
        pc = sqlite3.connect(PILOT_DB, timeout=10.0)
        pc.row_factory = sqlite3.Row
        pc.execute('PRAGMA busy_timeout=5000')
        try:
            # Cihaz sağlığı: dönemde tekrarlayan bağlantı kopmaları
            saglik = pc.execute(
                "SELECT robot_no, COUNT(*) AS n FROM saglik_olaylari "
                "WHERE ts >= ? AND ts <= ? GROUP BY robot_no ORDER BY n DESC LIMIT 5",
                (bas + ' 00:00:00', bit + ' 23:59:59')).fetchall()
            for s in saglik:
                if s['n'] < 5:
                    continue
                out.append(_b('cihaz_saglik', 'uyari',
                              f'{s["robot_no"]} sayaç modülü dönemde {s["n"]} kez kendini kurtardı',
                              'Sağlık olayı = modülün WiFi/bellek sorunundan yeniden başlaması. '
                              'Her olayda o anki sinyaller gecikir veya kaybolur; adet eksik sayılabilir.',
                              'Modülün sinyal gücüne ve besleme hattına bakın.',
                              cihaz=s['robot_no'], olay=s['n']))

            # Hayalet sayım: hiç vardiya olmayan günlerde gelen sinyal
            gunler = {v['tarih'] for v in vardiyalar}
            hayalet = pc.execute(
                "SELECT robot_no, substr(ts,1,10) AS gun, COUNT(*) AS n FROM sayac_olaylari "
                "WHERE ts >= ? AND ts <= ? GROUP BY robot_no, gun HAVING n > 20",
                (bas + ' 00:00:00', bit + ' 23:59:59')).fetchall()
            hat_gun = {(v['hat'], v['tarih']) for v in vardiyalar}
            kayip = [h for h in hayalet if (h['robot_no'], h['gun']) not in hat_gun and h['gun'] in gunler]
            if kayip:
                top = sum(h['n'] for h in kayip)
                out.append(_b('sayilmayan_sinyal', 'uyari',
                              f'{len(kayip)} makine-günde sinyal var ama vardiya yok — {top} sinyal kayda geçmedi',
                              'Sayaç sinyal üretmiş fakat o gün o hatta açılmış vardiya bulunamadı. '
                              'Sayım aktif vardiyaya kapılı olduğu için bu üretim hiçbir rapora girmez.',
                              'Vardiya açılmadan çalışılıyorsa operatöre hatırlatın; '
                              'makine boşta sinyal üretiyorsa parazit olabilir.',
                              makine_gun=len(kayip), sinyal=top,
                              ornekler=[{'hat': h['robot_no'], 'gun': h['gun'], 'sinyal': h['n']}
                                        for h in sorted(kayip, key=lambda x: -x['n'])[:5]]))

            # ── Sinyal boşluğu ↔ duruş kaydı karşılaştırması ──────────────────
            # TEK SORGU: dönemin tüm sinyalleri bir kerede çekilip (makine, gün)
            # bazında gruplanır. Vardiya başına sorgu atmak 23 günlük analizde
            # yüzlerce sorgu demekti (ölçüm: 8.5 sn → 1 sn altı).
            ilgili_hatlar = {v['hat'] for v in vardiyalar if v['hat']}
            zaman = defaultdict(list)
            if ilgili_hatlar:
                for r in pc.execute(
                        "SELECT robot_no, ts FROM sayac_olaylari "
                        "WHERE ts >= ? AND ts <= ? ORDER BY ts",
                        (bas + ' 00:00:00', bit + ' 23:59:59')):
                    if r['robot_no'] in ilgili_hatlar:
                        zaman[(r['robot_no'], r['ts'][:10])].append(r['ts'])

            gorulen = set()      # (hat, gün) — aynı boşluk iki operatör için iki kez çıkmasın
            for v in vardiyalar:
                if v['toplam'] <= 0 or not v['hat'] or (v['hat'], v['tarih']) in gorulen:
                    continue
                sinyaller = zaman.get((v['hat'], v['tarih']), [])
                if len(sinyaller) < 10:
                    continue
                # YALNIZ İKİ SİNYAL ARASINDAKİ boşluğa bakılır. Vardiya başındaki /
                # sonundaki sessizlik masum olabilir (hazırlık, temizlik, o makineye
                # geç geçilmesi); iki üretim arasındaki uzun sessizlik olamaz.
                en_uzun, nerede = 0.0, None
                for i in range(1, len(sinyaller)):
                    try:
                        t0 = datetime.strptime(sinyaller[i - 1], '%Y-%m-%d %H:%M:%S')
                        t1 = datetime.strptime(sinyaller[i], '%Y-%m-%d %H:%M:%S')
                    except ValueError:
                        continue
                    fark = (t1 - t0).total_seconds() / 60.0
                    if fark > en_uzun:
                        en_uzun, nerede = fark, t0.strftime('%H:%M')
                if en_uzun < ESIK['sinyal_bosluk_dk']:
                    continue
                # O makine-gündeki TÜM vardiyaların duruşu sayılır (iki operatör
                # aynı makinede olabilir; biri duruşu girmişse kayıp kayda geçmiştir)
                kayitli = sum(d['dk'] for w in vardiyalar
                              if w['hat'] == v['hat'] and w['tarih'] == v['tarih']
                              for d in w['duruslar'])
                if kayitli >= en_uzun * 0.6:      # boşluk zaten kayda geçmiş
                    continue
                gorulen.add((v['hat'], v['tarih']))
                out.append(_b('sinyal_bosluk', 'kritik' if en_uzun >= 120 else 'uyari',
                              f'{v["hat"]} — {v["tarih"]}: {en_uzun:.0f} dk sinyal yok, duruş kaydı {kayitli:.0f} dk',
                              f'Sayaç {nerede} civarında {en_uzun:.0f} dakika hiç sinyal üretmedi. '
                              f'O gün o hatta girilmiş toplam duruş {kayitli:.0f} dk — aradaki fark '
                              f'kayıt dışı duruş demek. Bu süre kullanılabilirliğe kayıp olarak '
                              f'yansımadığı için OEE olduğundan yüksek görünüyor.',
                              'Operatörle o saati konuşun; tekrarlıyorsa duruş girişi zorunlu hale gelmeli.',
                              hat=v['hat'], tarih=v['tarih'], bosluk_dk=round(en_uzun),
                              kayitli_dk=round(kayitli), saat=nerede, operator=v['operator']))
        finally:
            pc.close()
    except Exception as e:
        out.append(_b('sinyal_okunamadi', 'bilgi', 'Sinyal analizi yapılamadı',
                      f'pilot.db okunamadı: {e}', ''))
    out.sort(key=lambda x: -(x['sayisal'].get('bosluk_dk') or 0))
    return out[:10]


# ═════════════════════════════════════════════════════════════════════════════
#  ANA GİRİŞ
# ═════════════════════════════════════════════════════════════════════════════

def analiz_yap(bas, bit, lokasyon=None, bolum=None, yorum=False, conn=None):
    """Dönem analizi. yorum=True ise Claude API yorum katmanı da çalıştırılır."""
    kendi = conn is None
    conn = conn or _baglan()
    try:
        vardiyalar = vardiya_metrikleri(conn, bas, bit, lokasyon, bolum)
        genel = _topla(vardiyalar)
        op_tablo = operator_performans(vardiyalar)
        hat_tablo = hat_performans(vardiyalar)

        bulgular = []
        if vardiyalar:
            # SIRA ÖNEMLİ: veri kalitesi ve tanımsız süre bulguları ÖNCE gelir —
            # OEE yorumlanmadan önce sayının güvenilir olup olmadığı bilinmeli.
            bulgular += _bulgu_cycle_tanimsiz(vardiyalar, hat_tablo)
            bulgular += _bulgu_veri_kalitesi(conn, vardiyalar, bas, bit)
            bulgular += _bulgu_sinyal(vardiyalar, bas, bit)
            bulgular += _bulgu_oee(hat_tablo, genel)
            bulgular += _bulgu_durus(vardiyalar)
            bulgular += _bulgu_kalite(vardiyalar)
            bulgular += _bulgu_hedef(vardiyalar)
            bulgular += _bulgu_vardiya_farki(vardiyalar)
            bulgular += _bulgu_operator(op_tablo)
        bulgular.sort(key=lambda b: SIDDET_SIRA.get(b['siddet'], 9))

        sonuc = {
            'donem': {'bas': bas, 'bit': bit, 'lokasyon': lokasyon or 'HEPSİ',
                      'bolum': bolum or 'HEPSİ',
                      # Hangi bölümlerin dışarıda bırakıldığı EKRANDA yazsın:
                      # sessiz eksiltme, yanlış sayıdan daha kötüdür.
                      'haric_bolumler': ([] if bolum else list(OEE_DISI_BOLUMLER)),
                      'gun': len({v['tarih'] for v in vardiyalar})},
            'ozet': genel,
            'operatorler': op_tablo,
            'hatlar': hat_tablo,
            'bulgular': bulgular,
            'sayim': {s: sum(1 for b in bulgular if b['siddet'] == s)
                      for s in ('kritik', 'uyari', 'bilgi')},
            'yorum': None,
            'yorum_hata': '',
        }
        if yorum:
            metin, hata = yorum_uret(sonuc)
            sonuc['yorum'] = metin
            sonuc['yorum_hata'] = hata
        return sonuc
    finally:
        if kendi:
            conn.close()


def gunluk_ozet(gun=None, lokasyon=None, conn=None):
    """Tek günün özeti + en önemli bulguları (fabrika özeti alt bloğu / mail için).
    Yorum katmanı ÇALIŞTIRILMAZ — bu her sayfa açılışında çağrılabilir."""
    gun = gun or (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
    a = analiz_yap(gun, gun, lokasyon=lokasyon, yorum=False, conn=conn)
    a['bulgular'] = [b for b in a['bulgular'] if b['siddet'] in ('kritik', 'uyari')][:6]
    return a


# ═════════════════════════════════════════════════════════════════════════════
#  YORUM KATMANI (Claude API)
# ═════════════════════════════════════════════════════════════════════════════

VARSAYILAN_AI = {
    'etkin': False,
    # SAĞLAYICI (2026-08-21, kullanıcı: "kredi yüklemek istemiyorum"):
    #   anthropic → Claude API (kredi ister, en kaliteli yorum)
    #   gemini    → Google Gemini ücretsiz katman (aistudio.google.com'dan
    #               anahtar, kredi kartı GEREKMEZ; veri Google'a gider ve
    #               ücretsiz katmanda üründe kullanılabilir)
    #   ollama    → sunucuda yerel model (veri fabrikadan ÇIKMAZ, tamamen
    #               ücretsiz; RAM ister, CPU'da yavaştır)
    # gemini/ollama YALNIZ requests kullanır — sunucuya paket kurmak gerekmez.
    'saglayici': 'anthropic',
    'api_anahtari': '',
    'model': 'claude-opus-5',
    'max_tokens': 8000,
    'effort': 'medium',
    'dil': 'Türkçe',
    # 'gemini-flash-latest' TAKMA AD: her zaman guncel flash modeline isaret
    # eder. Sabit surum adi KULLANMAYIN — Google eski modelleri yeni hesaplara
    # kapatiyor (2026-08-21: 'gemini-2.5-flash' listede gorunuyordu ama
    # generateContent 404 'no longer available to new users' donuyordu).
    'gemini_model': 'gemini-flash-latest',
    'ollama_url': 'http://localhost:11434',
    'ollama_model': 'qwen2.5:7b',
    'ollama_timeout_sn': 420,     # CPU'da 7B model bir özeti dakikalarca yazabilir
}

SISTEM_PROMPT = """Sen bir imalat tesisinin üretim analistisin. Sana bir dönemin
üretim özeti, operatör/hat performans tablosu ve otomatik çıkarılmış bulgular
JSON olarak veriliyor.

Görevin: fabrika müdürüne okunacak KISA bir yönetici özeti yazmak.

Kurallar:
- Türkçe yaz, sade ve doğrudan. Süsleme yok.
- SADECE verilen veriye dayan. Veride olmayan sayıyı ASLA uydurma.
- Bulguları tekrar etmek yerine ARALARINDAKİ İLİŞKİYİ kur: aynı hatta hem hurda
  hem duruş varsa bunu birbirine bağla; bir operatör farkı makine farkıyla
  açıklanabiliyorsa söyle.
- Veri kalitesi bulgusu (tanımsız süre, kapatılmamış vardiya, girilmemiş duruş)
  varsa ÖNCE onu söyle: bu sorunlar diğer tüm sayıları şüpheli hale getirir.
- En fazla 4 kısa paragraf. Sonunda "Öncelik" başlığı altında en fazla 3 madde:
  bu hafta yapılırsa en çok kazandıracak işler.
- Emin olmadığın yerde "veri yetersiz" de. Kesin konuşmak zorunda değilsin."""


def ai_config(hata_ile=False):
    """ai_config.json — yoksa/bozuksa varsayılan (kapalı) döner.

    utf-8-SIG BİLİNÇLİ (2026-08-21): Not Defteri "UTF-8" ile kaydederken dosya
    başına BOM koyar; düz utf-8 ile okunduğunda json.load İLK KARAKTERDE patlar
    ve config sessizce varsayılana düşer — panelde "API anahtarı tanımlı değil"
    görünür, oysa anahtar dosyada durmaktadır. utf-8-sig BOM'u yutar, BOM'suz
    dosyayı da aynı şekilde okur.

    hata_ile=True → (cfg, hata_metni). Hatayı yutup sessizce varsayılana düşmek
    tanıyı imkânsız kılıyordu; artık sebep panele kadar taşınabiliyor."""
    cfg = dict(VARSAYILAN_AI)
    hata = ''
    anahtarlar = []
    try:
        if not os.path.exists(AI_CONFIG):
            hata = 'dosya yok'
        else:
            with open(AI_CONFIG, encoding='utf-8-sig') as f:
                ham = json.load(f)
            if not isinstance(ham, dict):
                hata = 'dosyanın içeriği JSON nesnesi değil'
            else:
                anahtarlar = [k for k in ham if not k.startswith('_')]
                cfg.update({k: v for k, v in ham.items() if not k.startswith('_')})
    except Exception as e:
        hata = f'{type(e).__name__}: {e}'
        print(f'[analiz] ai_config.json okunamadı: {e}')
    cfg['_alanlar'] = anahtarlar
    return (cfg, hata) if hata_ile else cfg


def _yorum_girdisi(analiz):
    """LLM'e giden veri — HAM DEĞİL, ÖZET.

    Tüm üretim kayıtlarını göndermek hem gereksiz (model zaten toplamları
    yorumlayacak) hem de dışarı çıkan veriyi büyütür. Referans kodları ve operatör
    adları bulgular için gerekli olduğundan kalır; başka kişisel veri yok."""
    return {
        'donem': analiz['donem'],
        'ozet': analiz['ozet'],
        'bulgular': [{k: b[k] for k in ('kod', 'siddet', 'baslik', 'detay', 'sayisal')}
                     for b in analiz['bulgular'][:25]],
        'operatorler': [{k: o[k] for k in ('operator', 'vardiya', 'toplam', 'oee',
                                           'beklenen_oee', 'hat_farki', 'hurda_oran')}
                        for o in analiz['operatorler'][:20]],
        'hatlar': [{k: h[k] for k in ('hat', 'bolum', 'vardiya', 'toplam', 'oee',
                                      'availability', 'performance', 'quality', 'plansiz_dk')}
                   for h in analiz['hatlar'][:20]],
    }


def yorum_uret(analiz, cfg=None):
    """Bulguları seçili sağlayıcıya yorumlatır. (sonuc, hata_mesaji) döner.

    HER HATA YUTULUR: yorum katmanı analizin SÜSÜ, kendisi değil. Sağlayıcı
    hangi sebeple cevap veremezse versin, yerel bulgular yine de gösterilir.

    Üç sağlayıcı da AYNI sistem promptunu ve AYNI özet girdiyi alır; dönüş
    sözleşmesi de aynı: {'metin','model','sure_sn','girdi_token','cikti_token'}.
    """
    cfg = cfg or ai_config()
    if not cfg.get('etkin'):
        return None, 'AI yorumu kapalı (ai_config.json → etkin: true yapın)'
    girdi = ('Aşağıdaki analiz çıktısını yorumla:\n\n'
             + json.dumps(_yorum_girdisi(analiz), ensure_ascii=False, indent=1))
    saglayici = (cfg.get('saglayici') or 'anthropic').strip().lower()
    try:
        if saglayici == 'anthropic':
            return _yorum_anthropic(cfg, girdi)
        if saglayici == 'gemini':
            return _yorum_gemini(cfg, girdi)
        if saglayici == 'ollama':
            return _yorum_ollama(cfg, girdi)
        return None, f'Tanınmayan sağlayıcı: {saglayici} (anthropic | gemini | ollama)'
    except Exception as e:
        return None, f'{type(e).__name__}: {e}'


def _yorum_anthropic(cfg, girdi):
    """Claude API — kredi ister; en kaliteli yorum. Paket: pip install anthropic."""
    anahtar = (cfg.get('api_anahtari') or os.environ.get('ANTHROPIC_API_KEY') or '').strip()
    if not anahtar:
        return None, 'API anahtarı tanımlı değil (ai_config.json → api_anahtari)'
    try:
        import anthropic
    except ImportError:
        return None, "anthropic paketi kurulu değil — sunucuda: pip install anthropic"
    try:
        istemci = anthropic.Anthropic(api_key=anahtar, timeout=120.0)
        istek = {
            'model': cfg.get('model') or 'claude-opus-5',
            'max_tokens': int(cfg.get('max_tokens') or 8000),
            'system': SISTEM_PROMPT,
            'messages': [{'role': 'user', 'content': girdi}],
        }
        if cfg.get('effort'):
            istek['output_config'] = {'effort': cfg['effort']}
        t0 = datetime.now()
        cevap = istemci.messages.create(**istek)
        if cevap.stop_reason == 'refusal':
            return None, 'Model isteği reddetti (güvenlik filtresi)'
        metin = '\n'.join(b.text for b in cevap.content if b.type == 'text').strip()
        if not metin:
            return None, 'Model boş yanıt döndü'
        return {'metin': metin, 'model': cevap.model,
                'sure_sn': round((datetime.now() - t0).total_seconds(), 1),
                'girdi_token': cevap.usage.input_tokens,
                'cikti_token': cevap.usage.output_tokens}, ''
    except anthropic.AuthenticationError:
        return None, 'API anahtarı geçersiz'
    except anthropic.RateLimitError:
        return None, 'API hız sınırı — birazdan tekrar deneyin'
    except anthropic.APIStatusError as e:
        return None, f'API hatası ({e.status_code}): {getattr(e, "message", e)}'
    except anthropic.APIConnectionError:
        return None, 'API bağlantısı kurulamadı (sunucunun internet erişimi var mı?)'


def _yorum_gemini(cfg, girdi):
    """Google Gemini — ÜCRETSİZ KATMAN (2026-08-21).

    Anahtar aistudio.google.com'dan Google hesabıyla alınır, KREDİ KARTI
    GEREKMEZ. Günlük kota günde-bir-rapor kullanımına fazlasıyla yeter.
    ⚠ VERİ GOOGLE'A GİDER ve ücretsiz katmanda Google bu veriyi ürün
    geliştirmede KULLANABİLİR. Gönderilen şey ham kayıt değil, özet tablolar +
    bulgu metinleridir (operatör adları ve referans kodları dahil) — şirket
    politikasına uygunluğuna kullanıcı karar verir.
    Yalnız requests kullanır; sunucuya paket kurmak gerekmez."""
    import requests
    anahtar = (cfg.get('api_anahtari') or os.environ.get('GEMINI_API_KEY') or '').strip()
    if not anahtar:
        return None, ('Gemini API anahtarı tanımlı değil (ai_config.json → api_anahtari). '
                      'Ücretsiz anahtar: aistudio.google.com → Get API key')
    model = (cfg.get('gemini_model') or 'gemini-flash-latest').strip()
    govde = {
        'systemInstruction': {'parts': [{'text': SISTEM_PROMPT}]},
        'contents': [{'role': 'user', 'parts': [{'text': girdi}]}],
        'generationConfig': {'maxOutputTokens': int(cfg.get('max_tokens') or 8000)},
    }
    t0 = datetime.now()

    def _cagri(m):
        return requests.post(
            f'https://generativelanguage.googleapis.com/v1beta/models/{m}:generateContent',
            params={'key': anahtar}, json=govde, timeout=120)

    try:
        r = _cagri(model)
        # 404'TE TAKMA ADA DUS (2026-08-21): Google eski model adlarini yeni
        # hesaplara KAPATIYOR — ListModels'ta gorunse bile generateContent 404
        # 'no longer available to new users' donebiliyor. Config'de bayat bir ad
        # kalsa da rapor uretilmeye devam etsin; kullanilan model yanit
        # sozlesmesindeki 'model' alaninda zaten gorunur.
        if r.status_code == 404 and model != 'gemini-flash-latest':
            model = 'gemini-flash-latest'
            r = _cagri(model)
        # 503'TE ONCE BEKLE-TEKRARLA, SONRA LITE MODELE DUS (2026-08-21):
        # ucretsiz katmanda "high demand" yogunluk hatasi geliyor ve genelde
        # saniyeler icinde geciyor. Ayni anda lite model cogunlukla bos
        # (canli olcum: flash-latest 503 verirken flash-lite-latest cevap verdi).
        # Gunluk rapor icin lite'in kalitesi fazlasiyla yeterli; hicbir insanin
        # beklemedigi 17:00 mail kosusunun bos donmemesi daha onemli.
        if r.status_code == 503:
            import time as _t
            _t.sleep(4)
            r = _cagri(model)
        if r.status_code == 503 and model != 'gemini-flash-lite-latest':
            model = 'gemini-flash-lite-latest'
            r = _cagri(model)
    except requests.exceptions.ConnectionError:
        return None, 'Gemini bağlantısı kurulamadı (sunucunun internet erişimi var mı?)'
    except requests.exceptions.Timeout:
        return None, 'Gemini zaman aşımı — birazdan tekrar deneyin'
    if r.status_code == 429:
        return None, 'Gemini kota/hız sınırı — ücretsiz katman dakikalık sınırlıdır, birazdan tekrar deneyin'
    if r.status_code == 503:
        return None, 'Gemini şu an çok yoğun (ücretsiz katman) — birkaç dakika sonra tekrar deneyin'
    if r.status_code in (401, 403):
        return None, 'Gemini API anahtarı geçersiz ya da yetkisiz'
    if r.status_code == 404:
        return None, (f'Gemini modeli bulunamadı: {model}. Google eski adları yeni '
                      f'hesaplara kapatabiliyor — ai_config.json → gemini_model '
                      f'değerini gemini-flash-latest yapın (her zaman güncel modele işaret eder)')
    if r.status_code != 200:
        try:
            mesaj = r.json()['error']['message']
        except Exception:
            mesaj = r.text[:200]
        return None, f'Gemini hatası ({r.status_code}): {mesaj}'
    d = r.json()
    adaylar = d.get('candidates') or []
    metin = ''
    if adaylar:
        parcalar = (adaylar[0].get('content') or {}).get('parts') or []
        metin = '\n'.join(p.get('text', '') for p in parcalar).strip()
    if not metin:
        sebep = ((adaylar[0].get('finishReason') if adaylar else None)
                 or (d.get('promptFeedback') or {}).get('blockReason') or 'boş yanıt')
        return None, f'Gemini içerik döndürmedi ({sebep})'
    um = d.get('usageMetadata') or {}
    return {'metin': metin, 'model': model,
            'sure_sn': round((datetime.now() - t0).total_seconds(), 1),
            'girdi_token': um.get('promptTokenCount') or 0,
            'cikti_token': um.get('candidatesTokenCount') or 0}, ''


def _yorum_ollama(cfg, girdi):
    """Ollama — SUNUCUDA YEREL MODEL (2026-08-21).

    Veri fabrikadan HİÇ ÇIKMAZ, tamamen ücretsiz, internet gerekmez.
    Kurulum (sunucuda): ollama.com'dan indir → `ollama pull qwen2.5:7b`.
    ⚠ CPU'da yavaştır: 7B model bir özeti dakikalarca yazabilir (günlük mail
    için sorun değil; panelde bekletir). ~6-8 GB boş RAM ister."""
    import requests
    url = (cfg.get('ollama_url') or 'http://localhost:11434').rstrip('/')
    model = cfg.get('ollama_model') or 'qwen2.5:7b'
    govde = {'model': model, 'stream': False,
             'messages': [{'role': 'system', 'content': SISTEM_PROMPT},
                          {'role': 'user', 'content': girdi}],
             'options': {'num_predict': int(cfg.get('max_tokens') or 8000)}}
    t0 = datetime.now()
    try:
        r = requests.post(url + '/api/chat', json=govde,
                          timeout=int(cfg.get('ollama_timeout_sn') or 420))
    except requests.exceptions.ConnectionError:
        return None, (f'Ollama çalışmıyor ({url}) — sunucuda kurulu ve açık olmalı '
                      f'(ollama.com, sonra: ollama pull {model})')
    except requests.exceptions.Timeout:
        return None, ('Ollama zaman aşımı — CPU için model büyük olabilir; '
                      'daha küçük model deneyin (örn. qwen2.5:3b) ya da '
                      'ollama_timeout_sn değerini artırın')
    if r.status_code == 404:
        return None, f'Model yüklü değil — sunucuda: ollama pull {model}'
    if r.status_code != 200:
        return None, f'Ollama hatası ({r.status_code}): {r.text[:200]}'
    d = r.json()
    metin = ((d.get('message') or {}).get('content') or '').strip()
    if not metin:
        return None, 'Ollama boş yanıt döndü'
    return {'metin': metin, 'model': model,
            'sure_sn': round((datetime.now() - t0).total_seconds(), 1),
            'girdi_token': d.get('prompt_eval_count') or 0,
            'cikti_token': d.get('eval_count') or 0}, ''


# ═════════════════════════════════════════════════════════════════════════════
#  TANI / ELLE TEST
# ═════════════════════════════════════════════════════════════════════════════
# Sunucuda panel girişi yapmadan tüm zinciri sınamak için:
#     python analiz.py                      → dün
#     python analiz.py 2026-08-01 2026-08-21 → aralık
# Yerel motoru çalıştırır, sonra AI katmanını dener ve NEREDE takıldığını yazar
# (config yok / paket yok / anahtar geçersiz / internet yok / kredi yok).

if __name__ == '__main__':
    import sys
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

    _bas = sys.argv[1] if len(sys.argv) > 1 else (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
    _bit = sys.argv[2] if len(sys.argv) > 2 else _bas

    print('=' * 70)
    print(f'ANALİZ TANI  ·  {_bas} → {_bit}')
    print('=' * 70)

    # 1) Yerel motor
    _t0 = datetime.now()
    _a = analiz_yap(_bas, _bit, yorum=False)
    _o = _a['ozet']
    print(f"\n[1] YEREL MOTOR  ({(datetime.now() - _t0).total_seconds():.1f} sn)")
    print(f"    vardiya={_o['vardiya']}  üretim={_bin(_o['toplam'])}  OEE=%{_o['oee']}  "
          f"(A%{_o['availability']} P%{_o['performance']} Q%{_o['quality']})")
    print(f"    bulgu: {_a['sayim']['kritik']} kritik · {_a['sayim']['uyari']} uyarı · "
          f"{_a['sayim']['bilgi']} bilgi")
    for _b_ in _a['bulgular'][:8]:
        print(f"      [{_b_['siddet']:6s}] {_b_['baslik'][:80]}")
    if not _o['vardiya']:
        print('    ! Bu aralıkta vardiya yok — AI testi için veri olan bir tarih verin.')

    # 2) Yapılandırma
    _cfg, _cfg_hata = ai_config(hata_ile=True)
    print('\n[2] AI YAPILANDIRMA')
    print(f"    ai_config.json      : {'VAR' if os.path.exists(AI_CONFIG) else 'YOK  → ' + AI_CONFIG}")
    if _cfg_hata and _cfg_hata != 'dosya yok':
        print(f"    ! DOSYA OKUNAMADI   : {_cfg_hata}")
        print("      JSON bozuk olabilir. (Not Defteri BOM'u artık tolere ediliyor.)")
    elif os.path.exists(AI_CONFIG):
        print(f"    dosyadaki alanlar   : {', '.join(_cfg.get('_alanlar') or []) or '(yok)'}")
    print(f"    etkin               : {_cfg.get('etkin')}")
    _sag = (_cfg.get('saglayici') or 'anthropic').strip().lower()
    print(f"    saglayici           : {_sag}   (anthropic | gemini | ollama)")
    _anh = (_cfg.get('api_anahtari') or os.environ.get(
        'ANTHROPIC_API_KEY' if _sag == 'anthropic' else 'GEMINI_API_KEY') or '').strip()
    # Anahtarın KENDİSİ yazdırılmaz — yalnız var mı ve şekli doğru mu.
    if _sag != 'ollama':
        print(f"    api_anahtari        : {'var (' + _anh[:7] + '…' + str(len(_anh)) + ' karakter)' if _anh else 'YOK'}")
    if _sag == 'anthropic':
        print(f"    model               : {_cfg.get('model')}   effort: {_cfg.get('effort')}")
        try:
            import anthropic as _ant
            print(f"    anthropic paketi    : {_ant.__version__}")
        except ImportError:
            print('    anthropic paketi    : KURULU DEĞİL  → pip install anthropic')
    elif _sag == 'gemini':
        print(f"    model               : {_cfg.get('gemini_model')}   (ücretsiz anahtar: aistudio.google.com)")
    else:
        print(f"    model               : {_cfg.get('ollama_model')}   url: {_cfg.get('ollama_url')}")

    # 3) Yorum katmanı
    print('\n[3] AI YORUM DENEMESİ')
    _t0 = datetime.now()
    _y, _hata = yorum_uret(_a)
    if _y:
        print(f"    ✓ BAŞARILI  ·  {_y['model']}  ·  {_y['sure_sn']} sn  ·  "
              f"{_y['girdi_token']} girdi + {_y['cikti_token']} çıktı token")
        print('    ' + '-' * 66)
        for _sat in _y['metin'].splitlines():
            print('    ' + _sat)
    else:
        print(f"    ✗ ALINAMADI: {_hata}")
        print(f"      ({(datetime.now() - _t0).total_seconds():.1f} sn sonra)")
        print('      Bulgular etkilenmez — yalnız yönetici özeti yazılmaz.')
    print()
