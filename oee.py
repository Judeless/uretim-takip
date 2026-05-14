from database import get_db

# Planli durus kategorileri
PLANLI_DURUS_KATEGORILER = {
    'Yemek Molasi',
    'Cay Molasi',
    'Planli Bakim',
    'Planli Temizlik',
}


def pair_cycle_hesapla(ist1_kod, ist2_kod):
    """Robot kaynak hattında iki istasyona atanan referansların gerçek paralel
    çevrim süresini hesaplar.

    Mantık:
      - Robot bir istasyonda kaynak yaparken operatör diğer istasyonda söktak yapar.
      - Her yarı cycle = max(robot_yapacağı_süre, operatörün_yapacağı_süre)
      - Toplam = max(K1, S2) + max(K2, S1)

    Tek istasyon kullanılıyorsa (diğer kod None/boş): tek-istasyon cycle = K + S

    Args:
      ist1_kod, ist2_kod: Referans kodları (str). None veya '' ise o istasyon boş.

    Returns:
      dict: {
        'pair_cycle_sn': toplam çevrim (sn),
        'ist1_kaynak_sn', 'ist1_soktak_sn',
        'ist2_kaynak_sn', 'ist2_soktak_sn',
        'ist1_bekleme_sn': operatör boş kaldığı süre (istasyon 1),
        'ist2_bekleme_sn': operatör boş kaldığı süre (istasyon 2),
        'verim_pct': bekleme/toplam oranı tersi (100%=hiç bekleme yok),
      }
    """
    def _ref(kod):
        if not kod:
            return None
        conn = get_db()
        r = conn.execute(
            "SELECT kaynak_suresi_sn, soktak_suresi_sn, hedef_cycle_time_sn "
            "FROM referans_listesi WHERE UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ',''))",
            (str(kod),)
        ).fetchone()
        conn.close()
        if not r:
            return None
        k = float(r['kaynak_suresi_sn'] or 0)
        s = float(r['soktak_suresi_sn'] or 0)
        if k == 0 and s == 0 and r['hedef_cycle_time_sn']:
            # Eski referans — split yok, varsayılan %40/%60
            cyc = float(r['hedef_cycle_time_sn'])
            k = round(cyc * 0.4, 1)
            s = round(cyc * 0.6, 1)
        return {'kaynak': k, 'soktak': s, 'toplam': round(k + s, 2)}

    r1 = _ref(ist1_kod)
    r2 = _ref(ist2_kod)

    if r1 and not r2:
        return {
            'pair_cycle_sn': r1['toplam'],
            'ist1_kaynak_sn': r1['kaynak'], 'ist1_soktak_sn': r1['soktak'],
            'ist2_kaynak_sn': 0, 'ist2_soktak_sn': 0,
            'ist1_bekleme_sn': 0, 'ist2_bekleme_sn': 0,
            'verim_pct': 100.0,
            'mod': 'tek_istasyon_1',
        }
    if r2 and not r1:
        return {
            'pair_cycle_sn': r2['toplam'],
            'ist1_kaynak_sn': 0, 'ist1_soktak_sn': 0,
            'ist2_kaynak_sn': r2['kaynak'], 'ist2_soktak_sn': r2['soktak'],
            'ist1_bekleme_sn': 0, 'ist2_bekleme_sn': 0,
            'verim_pct': 100.0,
            'mod': 'tek_istasyon_2',
        }
    if not r1 and not r2:
        return {
            'pair_cycle_sn': 0,
            'ist1_kaynak_sn': 0, 'ist1_soktak_sn': 0,
            'ist2_kaynak_sn': 0, 'ist2_soktak_sn': 0,
            'ist1_bekleme_sn': 0, 'ist2_bekleme_sn': 0,
            'verim_pct': 0,
            'mod': 'bos',
        }

    # İki istasyon paralel
    K1, S1 = r1['kaynak'], r1['soktak']
    K2, S2 = r2['kaynak'], r2['soktak']

    # Yarı cycle 1: Robot İst.1'de kaynak (K1), Op İst.2'de söktak (S2)
    yari1 = max(K1, S2)
    # Yarı cycle 2: Robot İst.2'de kaynak (K2), Op İst.1'de söktak (S1)
    yari2 = max(K2, S1)
    pair = yari1 + yari2

    # Bekleme süreleri (verim göstergesi)
    # Yarı 1'de robot İst.1'de kaynak yapıyor, op İst.2'de söktak. Hangisi
    # daha kısaysa o boş kalır.
    op_bos_yari1 = max(0, K1 - S2)   # Robot daha uzun → op boş bekledi (İst.2'de)
    rob_bos_yari1 = max(0, S2 - K1)  # Op daha uzun → robot boş bekledi
    op_bos_yari2 = max(0, K2 - S1)   # İst.1'de op boş
    rob_bos_yari2 = max(0, S1 - K2)

    toplam_bekleme = op_bos_yari1 + rob_bos_yari1 + op_bos_yari2 + rob_bos_yari2
    # Verim = (gerçek iş süresi) / (toplam toplam) ; toplam toplam = pair × 2 (hem robot hem op)
    toplam_kapasite = pair * 2
    verim = max(0, min(100, round((toplam_kapasite - toplam_bekleme) / toplam_kapasite * 100, 1))) \
            if toplam_kapasite > 0 else 0

    return {
        'pair_cycle_sn': round(pair, 2),
        'ist1_kaynak_sn': K1, 'ist1_soktak_sn': S1,
        'ist2_kaynak_sn': K2, 'ist2_soktak_sn': S2,
        'ist1_bekleme_sn': round(rob_bos_yari2 + op_bos_yari2, 1),  # İst.1'de op'un + robotun beklemesi
        'ist2_bekleme_sn': round(rob_bos_yari1 + op_bos_yari1, 1),
        'verim_pct': verim,
        'mod': 'paralel',
    }


def durus_tipi_belirle(sebep):
    """Durus sebebine gore planli/plansiz tipini belirle."""
    return 'planli' if sebep in PLANLI_DURUS_KATEGORILER else 'plansiz'


def hesapla_oee(vardiya_id):
    """Tek bir vardiya icin OEE hesapla."""
    conn = get_db()
    c = conn.cursor()

    vardiya = c.execute('SELECT * FROM vardiyalar WHERE id = ?', (vardiya_id,)).fetchone()
    if not vardiya:
        conn.close()
        return None

    toplam_vardiya_dk = vardiya['toplam_sure_dk']

    # Planli duruslar (mola, planli bakim) — Net Plan Suresi'nden dusulur
    planli_row = c.execute(
        "SELECT COALESCE(SUM(sure_dk), 0) as toplam FROM duruslar WHERE vardiya_id = ? AND durus_tipi = 'planli'",
        (vardiya_id,)
    ).fetchone()
    planli_durus_dk = planli_row['toplam']

    # Plansiz duruslar (ariza, setup) — Availability kaybi olarak gorunur
    plansiz_row = c.execute(
        "SELECT COALESCE(SUM(sure_dk), 0) as toplam FROM duruslar WHERE vardiya_id = ? AND durus_tipi = 'plansiz'",
        (vardiya_id,)
    ).fetchone()
    plansiz_durus_dk = plansiz_row['toplam']

    # Net Planli Uretim Suresi = Toplam Vardiya - Planli Duruslar
    net_plan_dk = max(0, toplam_vardiya_dk - planli_durus_dk)

    # Fiili Calisma Suresi = Net Plan - Plansiz Duruslar (ariza kayiplari)
    calisma_suresi_dk = max(0, net_plan_dk - plansiz_durus_dk)

    # Kullanilabilirlik (Availability) = Fiili Calisma / Net Plan
    availability = (calisma_suresi_dk / net_plan_dk) if net_plan_dk > 0 else 0

    # Uretim verileri (Referans listesi ile join yaparak guncel cycle time'i al)
    uretim_rows = c.execute('''
        SELECT u.*, r.hedef_cycle_time_sn as guncel_ct
        FROM uretim_kayitlari u
        LEFT JOIN referans_listesi r ON u.referans_kodu = r.referans_kodu
        WHERE u.vardiya_id = ?
    ''', (vardiya_id,)).fetchall()

    toplam_ok = sum(r['ok_adet'] for r in uretim_rows)
    toplam_nok = sum(r['nok_adet'] for r in uretim_rows)
    toplam_uretim = toplam_ok + toplam_nok
    toplam_hedef = sum(r['hedef_adet'] for r in uretim_rows)

    calisma_suresi_sn = calisma_suresi_dk * 60

    # Gercek uretim suresi (sn) = her referans icin adet x cycle_time
    gercek_uretim_sn = 0.0
    
    for r in uretim_rows:
        # Oncelikle referans listesindeki guncel sureyi kullan (retroaktif duzeltme)
        # Eger listede yoksa uretim kayidindaki eski sureyi kullan
        ct = r['guncel_ct'] if (r['guncel_ct'] and r['guncel_ct'] > 0) else (r['cycle_time_sn'] or 0)
        
        adet = r['ok_adet'] + r['nok_adet']
        if ct > 0:
            gercek_uretim_sn += adet * ct
        else:
            # Fallback: Eger hic sure tanimli degilse 1 birim = 1 sn sayilmasin (performans 0 cikar)
            # Ama kullanici performansi dusuk cikacak dediği için 0 birakiyoruz ki sure tanimlandiginda duzelsin
            gercek_uretim_sn += 0 

    # Performans = Gercek uretim suresi / Fiili calisma suresi
    # 100% asimi mumkun
    performance = (gercek_uretim_sn / calisma_suresi_sn) if calisma_suresi_sn > 0 else 0

    # Kalite (Quality)
    quality = (toplam_ok / toplam_uretim) if toplam_uretim > 0 else 0

    # OEE = A x P(cap 1.0) x Q
    oee = availability * min(performance, 1.0) * quality

    try:
        durum = vardiya['durum'] or 'kapali'
    except Exception:
        durum = 'kapali'

    conn.close()

    return {
        'vardiya_id': vardiya_id,
        'operator': vardiya['operator_adi'],
        'robot_no': vardiya['robot_no'],
        'baslangic_saati': vardiya['baslangic_saati'],
        'bitis_saati': vardiya['bitis_saati'],
        'tarih': vardiya['tarih'],
        'vardiya_turu': vardiya['vardiya_turu'],
        'durum': durum,
        'toplam_vardiya_dk': toplam_vardiya_dk,
        'planli_durus_dk': planli_durus_dk,
        'plansiz_durus_dk': plansiz_durus_dk,
        'net_plan_dk': net_plan_dk,
        'calisma_suresi_dk': calisma_suresi_dk,
        'toplam_ok': toplam_ok,
        'toplam_nok': toplam_nok,
        'toplam_uretim': toplam_uretim,
        'toplam_hedef': toplam_hedef,
        'gercek_uretim_sn': round(gercek_uretim_sn, 1),
        'availability': round(availability * 100, 1),
        'performance': round(performance * 100, 1),
        'quality': round(quality * 100, 1),
        'oee': round(oee * 100, 1),
    }


def hesapla_oee_ozet(tarih_baslangic=None, tarih_bitis=None, robot_no=None, bolum=None):
    """Tarih araligi, robot ve bolum icin OEE ozeti hesapla."""
    conn = get_db()
    c = conn.cursor()

    query = 'SELECT id FROM vardiyalar WHERE 1=1'
    params = []

    if tarih_baslangic:
        query += ' AND tarih >= ?'
        params.append(tarih_baslangic)
    if tarih_bitis:
        query += ' AND tarih <= ?'
        params.append(tarih_bitis)
    if robot_no:
        query += ' AND robot_no = ?'
        params.append(robot_no)
    if bolum:
        query += " AND COALESCE(bolum, 'kaynak') = ?"
        params.append(bolum)

    query += ' ORDER BY tarih DESC, id DESC'
    vardiya_ids = [row['id'] for row in c.execute(query, params).fetchall()]
    conn.close()

    sonuclar = []
    for vid in vardiya_ids:
        oee_data = hesapla_oee(vid)
        if oee_data:
            sonuclar.append(oee_data)

    if not sonuclar:
        return {
            'vardiya_sayisi': 0,
            'ort_availability': 0,
            'ort_performance': 0,
            'ort_quality': 0,
            'ort_oee': 0,
            'vardiyalar': []
        }

    return {
        'vardiya_sayisi': len(sonuclar),
        'ort_availability': round(sum(s['availability'] for s in sonuclar) / len(sonuclar), 1),
        'ort_performance': round(sum(s['performance'] for s in sonuclar) / len(sonuclar), 1),
        'ort_quality': round(sum(s['quality'] for s in sonuclar) / len(sonuclar), 1),
        'ort_oee': round(sum(s['oee'] for s in sonuclar) / len(sonuclar), 1),
        'vardiyalar': sonuclar
    }
