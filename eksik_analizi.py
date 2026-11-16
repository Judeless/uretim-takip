# -*- coding: utf-8 -*-
"""EKSİK ANALİZİ (kullanıcı 2026-11-16).

"Satış bölümünden mailde tablo ya da Excel hâlinde listeler geliyor; bu listelere plan
iletiyorum, üretim için eksiği olan mekanizmalar için eksik belirtiyorum. Forge'de ayrı bir
modül verdiğim referansların eksiklerini çıkartsın, tablo olarak kopyalayıp ürünlerin yanına
alabileyim. Eksikleri satın alma ve üretim eksikleri olarak ayrı belirt. P parçaları biz
üretiyoruz; P parçayı üretmek için bir A parça eksikse onu da belirtmemiz gerekiyor — örnek
LTK.733'ü üretmemiz gerekiyor, 300.5200 eksik, onun için de 300.2278 sac eksik."

AKIŞ
  1. Liste (yapıştırılan tablo ya da Excel) → satırlar: kod + ihtiyaç (eksi sayının mutlağı).
     Satır SIRASI korunur: çıktı aynı sırayla kopyalanıp Excel'de ürünlerin yanına yapıştırılır.
  2. Her ürün için ağaç açılır (hayali ara kodlar şeffaf — plan modülleriyle aynı:
     kp.urun_agaci_hayali). Ürünün AÇIK LAUNCH'ı ihtiyaçtan düşülür: launch alınınca alt
     parçalar REP'e çekilir, o kısım için serbest stoğa bakılmaz.
  3. Alt parça stoğu plan modüllerinin kuralıyla (kp.eldeki_stok, tesisin depo kümesi) okunur
     ve LİSTE SIRASIYLA tahsis edilir: aynı parçayı kullanan iki ürün aynı stoğu iki kez
     saymaz — üstteki satır önce alır (listedeki sıra = aciliyet).
  4. Eksik parça:
       A (satın alma) → SATIN ALMA eksiği (+ tedarikçide açık sipariş miktarı)
       P (üretim)     → ÜRETİM eksiği; parçanın açık launch'ı düşülür, KALAN için o parçanın
                        da ağacı açılır (300.5200 → 300.2278) — azami_derinlik kadar.
     Teyit/plan kuralları: '93.' teller TK2'de alt parça olarak sayılır ama ağaçlarına
     inilmez (TK1/fason üretir — montaj profiliyle aynı karar).
Bu modül ERP'ye YAZMAZ; yalnız okur. ERP okuması 'kaynak' nesnesi üzerinden (ErpKaynak) —
test sahte kaynakla yapılır."""
import io
import re

import kaynak_plan_kontrol as kp
import kod_sorgu as KS

KOD_RE = re.compile(r'^\d{2,3}\.[A-Za-zÇĞİÖŞÜçğıöşü0-9][A-Za-zÇĞİÖŞÜçğıöşü0-9./\-]*$')
BASLIK_KOD = re.compile(r'(code|kod|referans|article|ürün|urun)', re.I)
BASLIK_IHTIYAC = re.compile(r'(final|availab|ihtiya|minimum|eksik adet|gereken|need|qty|miktar)', re.I)
BASLIK_TESIS = re.compile(r'(bölüm|bolum|tesis|plant)', re.I)
BIRIMSIZ = ('AD', 'NR', 'PZ', 'PC', 'PCS', '')
TARIH_RE = re.compile(r'^\d{1,2}[./]\d{1,2}[./]\d{2,4}$')


def _baslik(c):
    """Türkçe büyük İ/I güvenli küçük harf ('MİNİMUM İHTİYAÇ' → 'minimum ihtiyaç')."""
    return str(c or '').strip().replace('İ', 'I').replace('ı', 'i').lower()


# Kısa yazım: 'LTK.733', 'dtc.472w' (önek harfli) — tam koda app tarafında çözülür
KISA_RE = re.compile(r'^[A-Za-z]{2,4}\.\d[A-Za-z0-9./\-]*$')


def _tarih_mi(c):
    m = TARIH_RE.match(c)
    if not m:
        return False
    g, a, y = (int(x) for x in re.split(r'[./]', c))
    return 1 <= g <= 31 and 1 <= a <= 12 and (y < 100 or 1990 <= y <= 2099)


def _kod_mu(c):
    c = str(c or '').strip().replace(' ', '')
    return bool(c) and bool(KOD_RE.match(c) or KISA_RE.match(c)) and not _tarih_mi(c)


def kod_normal(kod):
    return KS.kod_normal(kod)


def kisa_kod(kod, kisa=True):
    """'10.300.5200' → '300.5200' (kullanıcının yazım alışkanlığı); diğer önekler aynen."""
    return kod[3:] if kisa and kod.startswith('10.') else kod


def _sayi(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace('−', '-').replace(' ', '')
    if not s or not re.match(r'^-?[\d.,]+$', s):
        return None
    # 1.250 / 1,250 binlik; 12,5 ondalık — liste adetleri tam sayı olduğundan binlik varsay
    if re.match(r'^-?\d{1,3}([.,]\d{3})+$', s):
        s = s.replace('.', '').replace(',', '')
    else:
        s = s.replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


def _tesis_yazimi(v):
    s = str(v or '').upper().replace(' ', '').replace('-', '')
    return 'TK1' if s == 'TK1' else 'TK2' if s == 'TK2' else ''


def satirlari_ayikla(tablo):
    """tablo: [[hücre, ...], ...] (yapıştırılan metin ya da Excel'in ilk sayfası).
    Döner: (satirlar, bilgi) — satirlar: [{'sira', 'kod', 'ihtiyac', 'tesis', 'ham'}]
    Başlık satırı varsa sütunlar oradan (Code / Final availab. / MİNİMUM İHTİYAÇ ADETİ /
    Bölüm), yoksa her satırda ilk kod-benzeri hücre = kod, ilk EKSİ sayı (yoksa son sayı)
    = ihtiyaç. Kodsuz satır da SIRADA kalır (kopyalanan çıktı hizası bozulmasın)."""
    kod_s = iht_s = tes_s = None
    bas = 0
    for i, r in enumerate(tablo[:15]):
        hucre = [_baslik(c) for c in r]
        k = next((j for j, c in enumerate(hucre) if c and BASLIK_KOD.search(c) and not _kod_mu(c)), None)
        if k is not None:
            kod_s = k
            iht_s = next((j for j, c in enumerate(hucre) if c and BASLIK_IHTIYAC.search(c)), None)
            tes_s = next((j for j, c in enumerate(hucre) if c and BASLIK_TESIS.search(c)), None)
            bas = i + 1
            break
    satirlar = []
    for r in tablo[bas:]:
        hucre = list(r)
        metin = [str(c if c is not None else '').strip() for c in hucre]
        if not any(metin):
            # BOŞ SATIR DA SIRADA KALIR: çıktı Excel'deki aralığın yanına satır satır
            # yapıştırılıyor; arada atlanan satır alttaki her şeyi bir satır kaydırırdı
            satirlar.append({'sira': len(satirlar) + 1, 'kod': '', 'ihtiyac': None, 'tesis': '', 'ham': ''})
            continue
        kod = ''
        if kod_s is not None and kod_s < len(metin) and _kod_mu(metin[kod_s]):
            kod = metin[kod_s]
        if not kod and kod_s is None:
            kod = next((c for c in metin if _kod_mu(c)), '')
        ihtiyac = None
        if iht_s is not None and iht_s < len(hucre):
            ihtiyac = _sayi(hucre[iht_s])
        if ihtiyac is None:
            sayilar = [_sayi(c) for j, c in enumerate(hucre) if j != kod_s and str(c or '').strip() != kod]
            sayilar = [x for x in sayilar if x is not None]
            eksi = [x for x in sayilar if x < 0]
            ihtiyac = eksi[0] if eksi else (sayilar[-1] if sayilar else None)
        tesis = _tesis_yazimi(metin[tes_s]) if tes_s is not None and tes_s < len(metin) else ''
        satirlar.append({'sira': len(satirlar) + 1, 'kod': kod_normal(kod) if kod else '',
                         'ihtiyac': abs(ihtiyac) if ihtiyac is not None else None,
                         'tesis': tesis, 'ham': ' · '.join(c for c in metin if c)[:160]})
    while satirlar and not satirlar[-1]['ham']:          # sondaki boş satırlar (Excel'in fazlası)
        satirlar.pop()
    bilgi = {'baslik': bas > 0, 'kod_sutun': kod_s, 'ihtiyac_sutun': iht_s, 'tesis_sutun': tes_s}
    return satirlar, bilgi


def metinden_tablo(metin):
    """Excel / Outlook tablosundan kopyalanan metin: satırlar \\n, hücreler TAB. TAB yoksa
    (elle yazılmış 'kod adet' satırları) boşluklardan bölünür."""
    tablo = []
    for satir in str(metin or '').replace('\r\n', '\n').replace('\r', '\n').split('\n'):
        if '\t' in satir:
            tablo.append(satir.split('\t'))
        elif satir.strip():
            tablo.append(re.split(r'[;\s]+', satir.strip()))
        else:
            tablo.append([''])          # boş satır da sırada kalır (bkz. satirlari_ayikla)
    return tablo


def excelden_tablo(veri):
    """Yüklenen .xlsx / .xlsm'in İLK sayfası (formül sonuçlarıyla)."""
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(veri), data_only=True, read_only=True)
    ws = wb.worksheets[0]
    tablo = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return tablo, ws.title


class ErpKaynak:
    """AS400 okumaları (yalnız SELECT). Önbellekli: aynı kod bir analizde bir kez sorulur."""

    def __init__(self, cn):
        self.cn = cn
        self._kart, self._stok, self._siparis, self._launch, self._agac = {}, {}, {}, {}, {}

    def kartlar(self, kodlar):
        yeni = [k for k in dict.fromkeys(kodlar) if k not in self._kart]
        if yeni:
            bulunan = KS.urun_kartlari(self.cn, yeni)
            for k in yeni:
                self._kart[k] = bulunan.get(k)
        return {k: self._kart.get(k) for k in kodlar}

    def stoklar(self, kodlar):
        """{kod: {depo: bakiye}} + tedarikçide açık sipariş (S0ORDF)."""
        yeni = [k for k in dict.fromkeys(kodlar) if k not in self._stok and kp.GUVENLI_KOD.match(k)]
        cu = self.cn.cursor()
        for i in range(0, len(yeni), 50):
            grup = yeni[i:i + 50]
            for k in grup:
                self._stok[k], self._siparis[k] = {}, 0.0
            cu.execute("SELECT TRIM(S0ARTI), TRIM(S0MGCD), SUM(S0GIAD), SUM(S0ORDF) FROM TKC0300F.BSMAF0 "
                       f"WHERE S0SOCI='01' AND S0ARTI IN ({','.join('?' * len(grup))}) "
                       "GROUP BY TRIM(S0ARTI), TRIM(S0MGCD)", grup)
            for art, depo, q, sip in cu.fetchall():
                self._stok.setdefault(art, {})[depo] = float(q or 0)
                self._siparis[art] = self._siparis.get(art, 0.0) + float(sip or 0)
        return {k: self._stok.get(k, {}) for k in kodlar}

    def siparisler(self, kodlar):
        self.stoklar(kodlar)
        return {k: self._siparis.get(k, 0.0) for k in kodlar}

    def launchlar(self, kodlar):
        """{kod: açık launch kalanı (durum 40/45/50)}."""
        yeni = [k for k in dict.fromkeys(kodlar) if k not in self._launch]
        if yeni:
            o = kp.opr_ihtiyaclari(self.cn, yeni, ufuk_gun=3650)
            for k in yeni:
                self._launch[k] = float((o.get(k) or {}).get('launch_adet') or 0)
        return {k: self._launch.get(k, 0.0) for k in kodlar}

    def agac(self, kodlar, acilmayan_onekler=()):
        """{kod: ([(alt, birim, um)], {alt: iz})} — hayali ara kodlar açılmış."""
        anahtar = tuple(acilmayan_onekler or ())
        yeni = [k for k in dict.fromkeys(kodlar) if (k, anahtar) not in self._agac]
        if yeni:
            a, iz = kp.urun_agaci_hayali(self.cn, yeni, 3, anahtar)
            for k in yeni:
                self._agac[(k, anahtar)] = (a.get(k) or [], iz.get(k) or {})
        return {k: self._agac[(k, anahtar)] for k in kodlar}


def _kisit_disi_mi(kod, kisit_disi):
    return any(kod == k or (k.endswith('*') and kod.startswith(k[:-1])) for k in kisit_disi)


def analiz(kaynak, satirlar, ayar):
    """satirlar: satirlari_ayikla çıktısı (tesis'i boş olanlara çağıran Forge'dan doldurabilir).
    ayar: {'depolar': {'TK2': (sayilan, eksi_dusulen), 'TK1': (...)},
           'acilmayan': {'TK2': ('93.',), 'TK1': ()}, 'kisit_disi': (...), 'azami_derinlik': 4,
           'varsayilan_tesis': 'TK2'}
    Döner: satırlar (her biri: durum, satin_alma[], uretim[] ağaç, yapilabilir…) + özetler."""
    derin = int(ayar.get('azami_derinlik') or 4)
    kisit_disi = tuple(ayar.get('kisit_disi') or ())
    vars_tesis = ayar.get('varsayilan_tesis') or 'TK2'

    def depo(tesis):
        return ayar['depolar'].get(tesis) or ayar['depolar'][vars_tesis]

    def acilmayan(tesis):
        return tuple((ayar.get('acilmayan') or {}).get(tesis) or ())

    urunler = [s for s in satirlar if s.get('kod')]
    kart = kaynak.kartlar([s['kod'] for s in urunler])
    launch = kaynak.launchlar([s['kod'] for s in urunler if (kart.get(s['kod']) or {}).get('prov') == 'P'])

    havuz, launch_havuz = {}, {}            # (tesis, kod) → kalan serbest stok / kalan launch

    def serbest(tesis, kod, depolar):
        k = (tesis, kod)
        if k not in havuz:
            sayilan, eksi = depo(tesis)
            havuz[k] = kp.eldeki_stok(depolar.get(kod, {}), sayilan, eksi)
        return havuz[k]

    sonuc = []
    talepler = []                            # (satır, kod, adet, tesis, derinlik, düğüm)
    for s in satirlar:
        r = {'sira': s['sira'], 'kod': s.get('kod') or '', 'ihtiyac': s.get('ihtiyac'),
             'ham': s.get('ham', ''), 'tesis': s.get('tesis') or vars_tesis,
             'satin_alma': {}, 'uretim': [], 'notlar': [], 'durum': 'tamam',
             'aciklama': '', 'prov': '', 'launch': 0.0, 'net': None, 'yapilabilir': None}
        sonuc.append(r)
        if not r['kod']:
            r['durum'] = 'kod_yok'
            continue
        k = kart.get(r['kod'])
        if not k:
            r['durum'] = 'bulunamadi'
            continue
        r['aciklama'], r['prov'] = k.get('aciklama', ''), k.get('prov', '')
        if not r['ihtiyac']:
            r['durum'] = 'ihtiyac_yok'
            continue
        if r['prov'] != 'P':
            r['durum'] = 'satin_alinan'
            r['satin_alma'][r['kod']] = {'kod': r['kod'], 'eksik': r['ihtiyac'], 'um': '',
                                         'siparis': None, 'ust': set(), 'kendisi': True}
            continue
        r['launch'] = launch.get(r['kod'], 0.0)
        r['net'] = max(0.0, r['ihtiyac'] - r['launch'])
        if r['net'] <= 0:
            r['notlar'].append(f"açık launch ({r['launch']:g}) ihtiyacı karşılıyor")
            continue
        kok_dugum = {'kod': r['kod'], 'eksik': r['net'], 'cocuk': []}
        r['_kok'] = kok_dugum
        talepler.append((r, r['kod'], r['net'], r['tesis'], 0, kok_dugum))

    for seviye in range(derin + 1):
        if not talepler:
            break
        # Bu seviyede açılacak kodların ağacı (tesise göre '93.' açılmaz), alt parçaların
        # kartı / stoğu / launch'ı TOPLU okunur
        agaclar = {}
        for tesis in {t[3] for t in talepler}:
            agaclar[tesis] = kaynak.agac([t[1] for t in talepler if t[3] == tesis], acilmayan(tesis))
        altlar = sorted({a for tesis, ag in agaclar.items() for lst, _iz in ag.values() for a, _q, _u in lst})
        akart = kaynak.kartlar(altlar) if altlar else {}
        astok = kaynak.stoklar(altlar) if altlar else {}
        asip = kaynak.siparisler(altlar) if altlar else {}
        alaunch = kaynak.launchlar([a for a in altlar if (akart.get(a) or {}).get('prov') == 'P']) if altlar else {}
        sonraki = []
        for r, kod, adet, tesis, d, dugum in talepler:
            parcalar, iz = agaclar[tesis].get(kod, ([], {}))
            if not parcalar:
                if d == 0:
                    r['durum'] = 'agac_yok'
                else:
                    dugum['not'] = 'ürün ağacı yok'
                continue
            yapilabilir = None
            for alt, birim, um in sorted(parcalar):
                if birim <= 0 or _kisit_disi_mi(alt, kisit_disi) or (iz.get(alt) or {}).get('hayali'):
                    continue                # hayali (stoksuz) / kısıt dışı: plan modülleriyle aynı
                gerek = adet * birim
                eldeki = serbest(tesis, alt, astok)
                al = min(eldeki, gerek)
                havuz[(tesis, alt)] = eldeki - al
                if d == 0:
                    kac = int(al // birim + 1e-9)
                    yapilabilir = kac if yapilabilir is None else min(yapilabilir, kac)
                eksik = gerek - al
                if eksik <= 1e-6:
                    continue
                ak = akart.get(alt) or {}
                um = (um or '').strip().upper()
                if ak.get('prov') == 'P':
                    lk = (tesis, alt)
                    if lk not in launch_havuz:
                        launch_havuz[lk] = alaunch.get(alt, 0.0)
                    kapsar = min(launch_havuz[lk], eksik)
                    launch_havuz[lk] -= kapsar
                    cocuk = {'kod': alt, 'eksik': eksik, 'um': um, 'launch': kapsar,
                             'aciklama': ak.get('aciklama', ''), 'cocuk': []}
                    dugum['cocuk'].append(cocuk)
                    kalan = eksik - kapsar
                    if alt.startswith(acilmayan(tesis)):
                        cocuk['not'] = 'TK1 / fason (ağacına inilmedi)'
                    elif kalan > 1e-6:
                        if d + 1 <= derin:
                            sonraki.append((r, alt, kalan, tesis, d + 1, cocuk))
                        else:
                            cocuk['not'] = 'derinlik sınırı — alt parçalarına bakılmadı'
                else:
                    sa = r['satin_alma'].setdefault(alt, {'kod': alt, 'eksik': 0.0, 'um': um,
                                                         'siparis': asip.get(alt, 0.0), 'ust': set(),
                                                         'aciklama': ak.get('aciklama', ''),
                                                         'tanimsiz': not ak.get('prov')})
                    sa['eksik'] += eksik
                    if kod != r['kod']:
                        sa['ust'].add(kod)
            if d == 0:
                r['yapilabilir'] = yapilabilir
        talepler = sonraki

    for r in sonuc:
        kok = r.pop('_kok', None)
        if kok:
            r['uretim'] = kok['cocuk']
        r['satin_alma'] = sorted(({**v, 'ust': sorted(v['ust'])} for v in r['satin_alma'].values()),
                                 key=lambda x: x['kod'])
        if r['durum'] == 'tamam' and (r['satin_alma'] or r['uretim']):
            r['durum'] = 'eksik'
    return {'satirlar': sonuc, 'ozet': ozet(sonuc)}


def _adet(x, um=''):
    s = f'{x:,.0f}'.replace(',', '.') if abs(x - round(x)) < 1e-6 else f'{x:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    return s + ('' if (um or '').upper() in BIRIMSIZ else ' ' + um.lower())


def _dugum_metni(c, kisa):
    ek = [_adet(c['eksik'], c.get('um'))]
    if c.get('launch'):
        ek.append(f"launch {_adet(c['launch'])}")
    t = f"{kisa_kod(c['kod'], kisa)} ({'; '.join(ek)})"
    alt = c.get('cocuk') or []
    if alt:
        t += ' ← ' + ' + '.join((f'[{_dugum_metni(x, kisa)}]' if x.get('cocuk') and len(alt) > 1
                                  else _dugum_metni(x, kisa)) for x in alt)
    return t


def uretim_metni(dugumler, kisa=True):
    """'300.5200 (9; launch 4) ← 300.2278 (3); 93.TK.555 (7)'
    '←' = soldakini üretmek için eksik olan; '+' = aynı parçanın birden çok eksiği;
    ';' = ürünün ayrı üretim eksikleri; [ ] = kendi alt eksiği olan kardeş."""
    return '; '.join(_dugum_metni(c, kisa) for c in dugumler)


def satin_alma_metni(liste, kisa=True):
    """'50.200.002 (10; 300.5200 için; siparişte 500)'"""
    out = []
    for x in liste:
        if x.get('kendisi'):
            out.append(f"ürünün kendisi ({_adet(x['eksik'])})")
            continue
        ek = [_adet(x['eksik'], x.get('um'))]
        if x.get('ust'):
            ek.append(', '.join(kisa_kod(u, kisa) for u in x['ust']) + ' için')
        if x.get('siparis'):
            ek.append(f"siparişte {_adet(x['siparis'])}")
        out.append(f"{kisa_kod(x['kod'], kisa)} ({'; '.join(ek)})")
    return ', '.join(out)


DURUM_METNI = {'bulunamadi': "AS400'de bulunamadı", 'agac_yok': 'ürün ağacı yok',
               'ihtiyac_yok': '', 'kod_yok': '', 'tamam': '', 'satin_alinan': ''}


def kopya_hucreleri(r, kisa=True):
    """(satın alma, üretim, tek sütun) — Excel'e yapıştırılacak hücre metinleri."""
    if r['durum'] in ('bulunamadi', 'agac_yok'):
        return '', '', DURUM_METNI[r['durum']]
    sa = satin_alma_metni(r['satin_alma'], kisa)
    ur = uretim_metni(r['uretim'], kisa)
    tek = ' | '.join(x for x in (('Satınalma: ' + sa) if sa else '', ('Üretim: ' + ur) if ur else '') if x)
    return sa, ur, tek


def ozet(satirlar):
    """Liste geneli: satın alınacak ve üretilecek parçalar (kod başına toplam + etkilenen ürünler)."""
    sa, ur = {}, {}
    for r in satirlar:
        for x in r['satin_alma']:
            if x.get('kendisi'):
                continue
            d = sa.setdefault(x['kod'], {'kod': x['kod'], 'eksik': 0.0, 'um': x.get('um', ''),
                                         'siparis': x.get('siparis'), 'aciklama': x.get('aciklama', ''),
                                         'urunler': []})
            d['eksik'] += x['eksik']
            if r['kod'] not in d['urunler']:
                d['urunler'].append(r['kod'])

        def gez(dugumler, ust):
            for c in dugumler:
                d = ur.setdefault(c['kod'], {'kod': c['kod'], 'eksik': 0.0, 'um': c.get('um', ''),
                                             'launch': 0.0, 'aciklama': c.get('aciklama', ''),
                                             'urunler': [], 'ust': []})
                d['eksik'] += c['eksik']
                d['launch'] += c.get('launch') or 0
                if r['kod'] not in d['urunler']:
                    d['urunler'].append(r['kod'])
                if ust != r['kod'] and ust not in d['ust']:
                    d['ust'].append(ust)
                gez(c.get('cocuk') or [], c['kod'])
        gez(r['uretim'], r['kod'])
    durumlar = {}
    for r in satirlar:
        durumlar[r['durum']] = durumlar.get(r['durum'], 0) + 1
    return {'satin_alma': sorted(sa.values(), key=lambda x: -x['eksik']),
            'uretim': sorted(ur.values(), key=lambda x: -x['eksik']),
            'durumlar': durumlar}
