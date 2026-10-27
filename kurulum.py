# -*- coding: utf-8 -*-
"""KURULUM PROFİLİ — aynı kodun başka bir firmaya markalı kurulumu (kullanıcı 2026-10-02).

"Bir tedarikçimize Forge'u kurmak istiyoruz; adı Forge olmayacak, kendilerine has bir
program olacak. Bölüm olarak metal enjeksiyon, CNC işleme, montaj olacak."

Karar: TEK kod tabanı. Firmaya özgü olan her şey (ad, logo, bölümler, makineler, açık
modüller) bu profil dosyasından gelir; Forge'a yapılan düzeltmeler o kuruluma da gider.

Dosya: data/kurulum.json (ya da COFLE_KURULUM ortam değişkeni ile başka bir yol).
DOSYA YOKSA bu Cofle'nin kendi kurulumudur ve HER ŞEY bugünkü gibi çalışır — canlı
Cofle sunucusunda bu dosya YOKTUR; varsayılanlar bilerek Cofle'nin bugünkü hâlidir.
Örnek: ornek_kurulum/tedarikci.json. Değişiklik uygulamanın yeniden başlatılmasını ister.
"""
import copy
import json
import os

KOK = os.path.dirname(os.path.abspath(__file__))
YOL = os.environ.get('COFLE_KURULUM') or os.path.join(KOK, 'data', 'kurulum.json')

# Açılıp kapatılabilen modüller. Cofle'de hepsi açık.
MODULLER = {
    'as400':        'AS400 entegrasyonu (teyit, CFI, transfer, EOQ, planlama)',
    'planlar':      'Kaynak / montaj / metal planları (AS400 verisiyle)',
    'kapasite':     'Kapasite modülü (AS400 ürün havuzu)',
    'bakim':        'Bakım programı entegrasyonu + arıza onay kuyruğu',
    'proje':        'Proje takip',
    'tk1':          'TK1 yan tesis (ayrı operatör adresi ve andon)',
    'sayac':        'Saha sayaç cihazları (ESP32) ve sinyal analizi',
    'test_cihaz':   'Test cihazı sayaç entegrasyonu',
    'andon':        'Andon ekranları',
    'is_yonetimi':  'İş yönetimi (iş emri / öncelik takibi)',
    'mail':         'Günlük üretim raporu maili',
    'analiz':       'Analiz sayfası: otomatik bulgular, yönetici özeti, AI yorumu',
    'excel_senkron': 'Referans/duruş listesinin data/uretim_verileri.xlsx ile senkronu',
    'eski_sayfalar': 'Eski panel/andon/operatör sayfaları (legacy, önizleme)',
}
TUM_BOLUMLER = ('kaynak', 'montaj', 'metal', 'isleme', 'lazer', 'pres', 'plastik', 'tel', 'depo', 'kalite')
# FAALİYET BÖLÜMLERİ (kullanıcı 2026-10-27, Aspanda): "depo ve sevkiyat bir bölüm — operatörler
# gün içinde neler yapıyor görelim; kalite personeli günlük ne kadar ölçüm yaptığını yazsın".
# Makine / süre / OEE yok: kayıt = İŞ TÜRÜ (depo) ya da ÖLÇÜM TÜRÜ (kalite) + adet; türler
# referans listesinde durur (panel → Referanslar). Yalnız profilin 'bolumler' listesinde
# varsa açılır — Cofle'nin tam listesine (bolumler None) girmez.
FAALIYET_BOLUMLERI = ('depo', 'kalite')

VARSAYILAN = {
    'kimlik': 'cofle',
    'marka': {
        'ad': 'Cofle Forge',               # başlıklar, giriş ekranı, andon şeridi
        'kisa': 'Cofle',                   # ana ekran kısayolu adları ("Cofle Panel")
        'alt_baslik': 'Üretim Takip Sistemi',
        'alan_adi': 'https://coflemanage.online',
        'logo': '',                        # boş = static/logo.png (Cofle)
        'logo_koyu': '',                   # boş = static/logo_koyu.png
    },
    # None = kodun içindeki Cofle tanımları (TK1/TK2, sekiz bölüm, sabit makineler)
    'lokasyonlar': None,                   # [{'kod': 'TK2', 'ad': 'Merkez'}]
    'bolumler': None,                      # ['metal', 'isleme', 'montaj']
    'bolum_adlari': {},                    # {'isleme': 'CNC İşleme'}
    'makineler': {},                       # {'metal': ['50T-1', ...], 'isleme': [...]}
    'moduller': {m: True for m in MODULLER},
    # OTOMATİK MOLALAR (2026-10-26). None = kodun içindeki Cofle saatleri (app._bolum_molalari).
    # {'cay': [['10:00', 10], ...], 'yemek_sebep': 'Yemek Molası',
    #  'yemek_gruplari': {'1': {'saat': '12:00', 'dk': 30, 'ad': '...'}, ...}}
    # Yemek grubu operatöre bağlıdır (operatorler.yemek_grubu, panelden admin seçer).
    'molalar': None,
    # Veritabanı ilk kurulumunda bilinen şifreli Cofle yöneticisi açılsın mı?
    # Yalnız Cofle'de: internete açık yeni bir kurulumda bilinen şifre = açık kapı.
    'varsayilan_yonetici': True,
}

_onbellek = None


def _birlestir(taban, ust):
    for k, v in (ust or {}).items():
        if isinstance(v, dict) and isinstance(taban.get(k), dict):
            _birlestir(taban[k], v)
        else:
            taban[k] = v
    return taban


def yukle(tazele=False):
    """Profil sözlüğü. Dosya bozuksa SESSİZ geçmez: hata basılır ve uygulama Cofle
    varsayılanıyla açılmaz — yanlış firmanın adıyla yayın yapmaktansa durmak iyidir."""
    global _onbellek
    if _onbellek is not None and not tazele:
        return _onbellek
    p = copy.deepcopy(VARSAYILAN)
    if os.path.exists(YOL):
        with open(YOL, encoding='utf-8-sig') as f:
            dosya = json.load(f)
        # İKİ KATMAN (2026-10-26): repodaki ornek_kurulum/<kimlik>.json TABAN, sunucudaki
        # data/kurulum.json ÜSTÜNE yazar. data/kurulum.json kurulumda bir kez kopyalanır ve
        # güncellemelerde (Aspanda_Guncelle.bat) dokunulmaz — profile sonradan eklenen
        # ayarlar (ör. molalar) ancak bu tabanla sunucuya ulaşır.
        _kim = str(dosya.get('kimlik') or '')
        _taban = os.path.join(KOK, 'ornek_kurulum', _kim + '.json')
        _taban_bl = None
        if _kim and _kim != 'cofle' and _kim.replace('_', '').isalnum() and os.path.exists(_taban):
            with open(_taban, encoding='utf-8-sig') as f:
                _tb = json.load(f)
            _birlestir(p, _tb)
            _taban_bl = _tb.get('bolumler')
        _birlestir(p, dosya)
        # 'bolumler' bir LİSTE: üst katman tabanı tümden ezer → profile sonradan eklenen
        # bölüm (depo / kalite, 2026-10-27) sunucunun eski kopyasının arkasında kalırdı.
        # Taban + yerel BİRLEŞİMİ (taban sırası). Bölüm kaldırmak = repodaki profilden.
        if isinstance(_taban_bl, list) and isinstance(dosya.get('bolumler'), list):
            p['bolumler'] = list(_taban_bl) + [b for b in dosya['bolumler'] if b not in _taban_bl]
        p['_dosya'] = YOL
        # Profil dosyası olan kurulum Cofle değildir: kimlik verilmemişse 'ozel'
        if p.get('kimlik') == 'cofle' and 'kimlik' not in dosya:
            p['kimlik'] = 'ozel'
    else:
        p['_dosya'] = ''
    bilinmeyen = [b for b in (p.get('bolumler') or []) if b not in TUM_BOLUMLER]
    if bilinmeyen:
        raise ValueError(f'kurulum.json: bilinmeyen bölüm {bilinmeyen} — geçerli: {TUM_BOLUMLER}')
    _onbellek = p
    return p


def cofle_mi():
    return yukle().get('kimlik') == 'cofle'


def modul(ad):
    """Modül açık mı? Bilinmeyen ad → açık (yeni özellik kurulumu bozmasın)."""
    return bool((yukle().get('moduller') or {}).get(ad, True))


def bolumler():
    """Bu kurulumun bölümleri; None = Cofle'nin tam listesi (kod içi tanımlar)."""
    return yukle().get('bolumler')


def bolum_var(b):
    bl = bolumler()
    return (b not in FAALIYET_BOLUMLERI) if bl is None else b in bl


def makineler(bolum):
    """Profilde tanımlı makine listesi ya da None (kod içi varsayılan)."""
    m = (yukle().get('makineler') or {}).get(bolum)
    return list(m) if m else None


def marka():
    return yukle()['marka']


def molalar():
    """Profildeki mola tanımı ya da None (Cofle'nin kod içi saatleri)."""
    m = yukle().get('molalar')
    return m if isinstance(m, dict) and (m.get('cay') or m.get('yemek_gruplari')) else None


def istemci_ozeti():
    """Şablonlara / tarayıcıya giden kısım (dosya yolu ve iç alanlar hariç)."""
    p = yukle()
    return {
        'kimlik': p['kimlik'], 'cofle': p['kimlik'] == 'cofle',
        'marka': {k: v for k, v in p['marka'].items() if k not in ('logo', 'logo_koyu')},
        'lokasyonlar': p.get('lokasyonlar'), 'bolumler': p.get('bolumler'),
        'bolum_adlari': p.get('bolum_adlari') or {}, 'makineler': p.get('makineler') or {},
        'moduller': {m: modul(m) for m in MODULLER},
        'molalar': molalar(),
    }
