from database import get_db

# Planli durus kategorileri
PLANLI_DURUS_KATEGORILER = {
    'Yemek Molasi',
    'Cay Molasi',
    'Planli Bakim',
    'Planli Temizlik',
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


def hesapla_oee_ozet(tarih_baslangic=None, tarih_bitis=None, robot_no=None):
    """Tarih araligi ve/veya robot icin OEE ozeti hesapla."""
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
