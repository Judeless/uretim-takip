# -*- coding: utf-8 -*-
"""AS400 (S.I. COFLETK) baglanti ayarlari — ORTAK sabitler.
ONEMLI: Bu dosyada ASLA sifre YOKTUR. Sifre Windows Kimlik Bilgileri
Yoneticisi'nde (keyring / DPAPI sifreli) durur; runtime'da okunur.
Sifreyi kaydetmek icin: python as400/kaydet_sifre.py
"""

# Windows Kimlik Bilgileri Yoneticisi anahtarlari
KEYRING_SERVICE = 'cofle_as400'
ROBOT_VARSAYILAN = 'EMREDTK'         # 5250 ekran robotu (RPR/COP) — oturum_config.json 'kullanici' ile degisir
IMPORT_KULLANICI = 'COFLEFORGE'      # COFLEFORGE.BMMAF0I INSERT profili (Italya IT, 2026-09-09)
KULLANICILAR     = ('EMREDTK', 'COFLEFORGE')   # kasadan okunabilecek profiller (beyaz liste)


def _odbc_kullanici():
    """ODBC (QZDASOINIT) profili: as400/odbc_config.json {"kullanici": "..."}.
    Simone Rota 2026-09-10: COFLEFORGE 'transaction' kullanicisidir, 5250 menusu
    yok; ODBC baglantilarinda EMREDTK yerine o kullanilmali. Dosya yoksa EMREDTK
    (eski davranis). Degisiklik RESTART ister (modul yuklenirken okunur):
    cofle-app servisi + teyit-agent. Once: python as400/odbc_profil_test.py COFLEFORGE"""
    import os, json
    yol = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'odbc_config.json')
    try:
        with open(yol, encoding='utf-8-sig') as f:
            ku = str((json.load(f) or {}).get('kullanici') or '').strip().upper()
        return ku if ku in KULLANICILAR else ''
    except FileNotFoundError:
        return ''
    except Exception as e:
        print(f'[as400_config] odbc_config.json okunamadi ({e}) — {ROBOT_VARSAYILAN} kullaniliyor')
        return ''


# AS400 ODBC kullanici adi (sifre DEGIL): tum SELECT'ler (launch listesi, hareket
# kontrolu, dogrulama) bu profille acilir; import modulu IMPORT_KULLANICI ile.
DB_KULLANICI = _odbc_kullanici() or ROBOT_VARSAYILAN

# Baglanti
HOST   = '192.168.1.1'
DRIVER = 'iSeries Access ODBC Driver'   # 64-bit (Client Access 32-bit DEGIL)

# CANLI uretim emri kaynagi (2026-07-17 katalog kesfiyle bulundu).
# BPROF0 = OP ana dosyasi (~160k; durum 70=kapali cogunluk). XPRO90 = yalniz
# ACIK emirler gorunumu (durum 10/40/45/50, ~5k satir) — hizli ve daima guncel.
# Eski kaynak XB530W0 work-file'i (10>05>03>29+F6 raporu) yalniz GECIKMIS
# launch'lari icerdigi icin TERK EDILDI (F6/Excel adimi artik gereksiz).
KAYNAK_TABLO = 'tkc0301F.XPRO90'

# Kolonlar (BPROF0/XPRO90 ortak):
#   Q0RED1/Q0RED2/Q0RENU = OP no bilesenleri (SS-YY-Nr → launch numarasi)
#   Q0ARTI = article (referans), Q0AVAN = durum:
#       10 = OPR var, launch ALINMAMIS; 40 = launch acik (teyit verilebilir);
#       45/50 = TK1 akisi; 70 = kapali (yalniz BPROF0'da)
#   Q0QTOR = launch adedi, Q0QTRI = simdiye kadar teyit verilen (rientrata)
SORGU = (
    "SELECT Q0RED1, Q0RED2, Q0RENU, Q0ARTI, Q0AVAN, Q0QTOR, Q0QTRI "
    "FROM " + KAYNAK_TABLO
)

# ── TEYIT (rientro) HAREKETLERI ──
# Depo hareket ana dosyasi. MGCACD='RPR' = rientro produzione = uretim teyidi.
# Ekranda gorunen "Re-entry summary" satirlarinin kaynagi (dogrulandi 2026-07-17).
# ONEMLI: hareket tarihi = teyidin GIRILDIGI tarih (dunun uretimi bugun girilir),
# uretim tarihi DEGIL. Mukerrer kontrolu bunu hesaba katmalidir.
HAREKET_TABLO = 'tkc0301F.BMMAF0'
# ARTICLE bazli sorgu (launch bazli DEGIL): operator ayni urunu BASKA bir launch'a
# girip S ile kapatmis olabilir (o launch durum 70 → acik listede gorunmez).
# Mukerrer kontrolu bu yuzden article uzerinden yapilmali (2026-07-17 dersi:
# 1847A → operator 140636'ya girdi+kapatti; bizim listede zombi 17809 vardi).
# RPR = launch teyidi; CFI = launch'siz uretimin depo girisi (07>01>F1, kullanici
# 2026-07-20 ogretti) — ikisi de "bu uretim ERP'ye islendi" kaniti sayilir.
HAREKET_SORGU = (
    "SELECT MGARCD, MGDSSO, MGDAAO, MGDMMO, MGDGGO, MGQTA, MGRED2, MGRENU, MGCACD "
    "FROM " + HAREKET_TABLO + " WHERE MGCACD IN ('RPR','CFI') AND MGARCD IN ({yer}) "
)


def baglanti_dizesi(sifre, kullanici=None):
    """pyodbc connection string uretir. sifre yalnizca runtime'da,
    bellekte kullanilir — hicbir yere yazilmaz. kullanici: varsayilan
    DB_KULLANICI; import modulu IMPORT_KULLANICI ile cagirir."""
    return (
        f"DRIVER={{{DRIVER}}};"
        f"SYSTEM={HOST};"
        f"UID={kullanici or DB_KULLANICI};"
        f"PWD={sifre};"
        # Metin cevirisi: AS400 EBCDIC → Unicode (Turkce alanlar dogru gelsin)
        "CCSID=1208;TRANSLATE=1;"
    )


# ── AS400 SIFRESINI AL (merkezi) ──
# Onceligi: (1) keyring — laptop, ya da servis promanage hesabinda kosuyorsa.
# (2) FALLBACK: teyit-agent (127.0.0.1:5010/sifre). Server'da cofle-app NSSM
# servisi LocalSystem'da kosar → keyring (promanage kasasi) BOS gorunur; ama agent
# RDP/promanage oturumunda kosar, keyring'i gorur ve sifreyi yalniz localhost'a verir.
# Boylece servis hesabini degistirmeye (log-on-as-service tuzagi) gerek kalmaz.
# Sifre keyring'de zaten sifreli durur (kaydet_sifre.py) — bu fonksiyon onu yalniz
# bellekte tasir, hicbir yere yazmaz.
AGENT_URL = 'http://127.0.0.1:5010'


def sifre_al(kullanici=None):
    """AS400 sifresi: once keyring, olmazsa teyit-agent (localhost). None=bulunamadi.
    kullanici: varsayilan DB_KULLANICI; import icin IMPORT_KULLANICI."""
    ku = (kullanici or DB_KULLANICI).strip().upper()
    try:
        import keyring
        pw = keyring.get_password(KEYRING_SERVICE, ku)
        if pw:
            return pw
    except Exception:
        pass
    try:
        import requests
        # Kullanici HER ZAMAN acikca yollanir (2026-09-15): eskiden ku == DB_KULLANICI ise
        # parametresiz soruluyordu → agent'in kendi varsayilani farkliysa (config/restart
        # farki) BASKA kullanicinin sifresi bu kullanici adiyla AS400'e gidebiliyordu.
        r = requests.get(AGENT_URL + '/sifre', params={'kullanici': ku}, timeout=2)
        if r.status_code == 200:
            return (r.json() or {}).get('sifre') or None
    except Exception:
        pass
    return None


# ── BAĞLANTI KİLİDİ (devre kesici, 2026-09-15) ──
# OLAY: COFLEFORGE profili AS400'de 'disabilitato' oldu (CWBSY0011). IBM i, art arda
# başarısız girişte (QMAXSIGN) profili kapatır; bizim otomatik koşular (erken teyit
# 2-3 dk'da bir, launch listesi, doğrulamalar) kimlik hatasına rağmen denemeye devam
# ederse profil AÇILDIKTAN dakikalar sonra yine kilitlenir. Ayrıca şifre alınamayınca
# 'PWD=None' ile giriş denemesi yapılıyordu (o da başarısız giriş sayılır).
# KURAL: kimlik hatası (SQLSTATE 28000 / CWBSY…) → o kullanıcı için KİLİT (dosyada;
# cofle-app servisi ve teyit-agent aynı dosyayı görür) → sonraki otomatik bağlantılar
# AS400'e HİÇ gitmeden BaglantiKilitli ile düşer. Kilidi kaldıran: kaydet_sifre.py <ku>
# (yeni şifre) ya da odbc_profil_test.py <ku> (tek elle deneme, başarılıysa).
# İletişim hataları (CWBCO… / 08001: ağ, sunucu kapalı) kilit KOYMAZ.
KILIT_DOSYA = __import__('os').path.join(__import__('os').path.dirname(__import__('os').path.abspath(__file__)),
                                         'baglanti_kilidi.json')


class BaglantiKilitli(RuntimeError):
    """Kimlik hatası sonrası otomatik bağlantı durduruldu (AS400'e gidilmedi)."""


def _kilitler():
    import json
    try:
        with open(KILIT_DOSYA, encoding='utf-8') as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except (FileNotFoundError, ValueError):
        return {}
    except Exception:
        return {}


def kilit_durumu(kullanici=None):
    """Kilit varsa {'ts','mesaj'} döner, yoksa None."""
    return _kilitler().get((kullanici or DB_KULLANICI).strip().upper())


def kilit_koy(kullanici, mesaj):
    import json, time
    ku = (kullanici or DB_KULLANICI).strip().upper()
    d = _kilitler()
    d[ku] = {'ts': time.strftime('%Y-%m-%d %H:%M:%S'), 'mesaj': str(mesaj)[:500]}
    try:
        with open(KILIT_DOSYA, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, indent=1)
    except Exception as e:
        print(f'[as400_config] baglanti kilidi yazilamadi: {e}')


def kilit_temizle(kullanici):
    import json
    ku = (kullanici or DB_KULLANICI).strip().upper()
    d = _kilitler()
    if ku in d:
        d.pop(ku, None)
        try:
            with open(KILIT_DOSYA, 'w', encoding='utf-8') as f:
                json.dump(d, f, ensure_ascii=False, indent=1)
        except Exception as e:
            print(f'[as400_config] baglanti kilidi silinemedi: {e}')
        return True
    return False


def kimlik_hatasi_mi(e):
    """Giriş/kimlik hatası mı? (yanlış şifre, profil devre dışı, şifre süresi dolmuş,
    bilinmeyen kullanıcı) — SQLSTATE 28000 ya da Client Access güvenlik kodu CWBSY."""
    s = str(e)
    return ('28000' in s) or ('CWBSY' in s)


def baglan(sifre=None, kullanici=None, timeout=20, kilidi_yoksay=False):
    """AS400 ODBC bağlantısı — TEK MERKEZ (autocommit). Kilitliyse AS400'e gitmez;
    şifre yoksa denemez; kimlik hatasında kilit koyar. kilidi_yoksay yalnız elle
    yapılan tek deneme içindir (odbc_profil_test.py)."""
    import pyodbc
    ku = (kullanici or DB_KULLANICI).strip().upper()
    if not kilidi_yoksay:
        k = kilit_durumu(ku)
        if k:
            raise BaglantiKilitli(
                f"AS400 bağlantısı DURDURULDU ({ku}): {k.get('ts', '')} tarihli giriş denemesi kimlik hatası verdi "
                f"— profil yeniden kilitlenmesin diye otomatik denemeler kapalı. İtalya IT profili açıp şifreyi "
                f"teyit edince sunucuda tek deneme: python as400\\odbc_profil_test.py {ku}  "
                f"(şifre değiştiyse önce: python as400\\kaydet_sifre.py {ku}) | son hata: {str(k.get('mesaj', ''))[:220]}")
    if sifre is None:
        sifre = sifre_al(ku)
    if not sifre:
        raise RuntimeError(f'AS400 şifresi alınamadı ({ku}) — kasada yok ya da teyit-agent kapalı '
                           f'(kaydet_sifre.py {ku}); boş şifreyle giriş DENENMEDİ')
    try:
        return pyodbc.connect(baglanti_dizesi(sifre, ku), timeout=timeout, autocommit=True)
    except Exception as e:
        if kimlik_hatasi_mi(e):
            kilit_koy(ku, e)
        raise
