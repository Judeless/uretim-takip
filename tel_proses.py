# -*- coding: utf-8 -*-
"""
tel_proses.py — TK1 tel üretiminin proses adımı mantığı (2026-08-04).

Neden ayrı modül: bu kuralı app.py, mail_raporu.py ve oee.py'nin üçü de kullanıyor.
mail_raporu/oee, app.py'yi import EDEMEZ (Flask uygulamasını ayağa kaldırır, döngüsel
bağımlılık olur) — ortak mantık burada durur, üçü de buradan alır.

TEMEL KURAL
-----------
Bir tel referansı sırayla birden çok HATTAN geçer (halat kesme → yarı/tam otomat →
kapama → son montaj) ve her adımı FARKLI operatör kendi vardiyasında kaydeder.
Adımlar aynı kodla toplanırsa aynı 100 adetlik iş 4 kez sayılır.

ÇÖZÜM: her adım kendi EKİYLE kaydedilir — '93.TK.464 KESIM', '93.TK.464 KAPAMA'…
Böylece adımlar ayrı kodlarda toplanır; ne çoklu sayım olur ne de kimin ne iş
yaptığı kaybolur.

NEDEN "SON ADIM" TANIMI YOK (2026-08-04 revizyonu):
Kapaması ya da son montajı dışarıda yapılacak ürünler SABİT DEĞİL — üretim
yoğunluğuna ve tedarikçi durumuna göre değişiyor. Bu yüzden bir referansın son
adımı önceden tanımlanamaz; 'son adım base kodu alır' kuralı terk edildi.
Ayrıca TK1'de üretim teyidini KALİTE departmanı verdiği için base kodun ERP
article'ıyla birebir kalma zorunluluğu da yok.
"""

import re

# Üretim sırasında proses adımları.
TEL_ADIMLARI = ('Halat Kesme', 'Yarı Otomatik', 'Tam Otomatik', 'Kapama', 'Son Montaj')

# Sıra numarası: yarı ve tam otomatik AYNI adım sayılır (2) — biri diğerinin
# alternatifi, ikisi arka arkaya uygulanmaz.
TEL_ADIM_SIRA = {'Halat Kesme': 1, 'Yarı Otomatik': 2, 'Tam Otomatik': 2,
                 'Kapama': 3, 'Son Montaj': 4}

# FİZİKSEL HATLAR (kullanıcı 2026-08-04): kesim 3 · yarı otomatik 2 · tam otomatik 1
# · kapama 12 · son montaj 4 = 22 hat. Tek makineli adımlarda numara YOK.
# KAPAMA presleri 2026-08-18'de SAHA KODLARINA çevrildi — aşağıdaki nota bakın.
# YARI OTOMATİK 1 → 2 (kullanıcı 2026-08-18): sahada iki makine olduğu ortaya çıktı.
# İkisinin de sayacı ESP32 DEĞİL, SVP test cihazı (app.TEST_CIHAZ_ESLEME:
# 'Yarı Otomatik 1'→41, 'Yarı Otomatik 2'→43) — bu hatlara saha modülü takılmayacak.
# KAPAMA 4 → 12 (2026-08-07): kapama hattında 12 pres var; sayaç modülleri
# takılırken gerçek makine sayısı ortaya çıktı (eski 4 rakamı eksikti). Presler
# üçerli ortak panolarda; 4 ESP32 modülü 3'er presi okur (bkz. pilot/firmware/
# _templates/tel_kapama.ino.tpl). Hat adları firmware'deki robot_no ile BİREBİR.
# ── KAPAMA PRESLERİ: SAHA KODU (kullanıcı 2026-08-18) ───────────────────────
# Modüller takılınca sistemdeki sıra numarası ile presin ÜRETİM SAHASINDAKİ kodu
# tutmadığı görüldü: sistemde 'Kapama 1' olan modülün bağlı olduğu presin saha
# kodu 5. Operatör sahada gördüğü kodu seçmeli → hat adları SAHA KODUNA çevrildi.
#
# FIRMWARE YENİDEN YÜKLENMEDİ: modüller hâlâ robot_no='Kapama 1..6' gönderiyor.
# Hat adı → cihaz eşlemesi app.HAT_SAYAC_CIHAZI'nda (LF-LFP→YF1 ile aynı kalıp).
#   saha 5  ← cihaz Kapama 1      saha 30 ← cihaz Kapama 4
#   saha 12 ← cihaz Kapama 2      saha 28 ← cihaz Kapama 5
#   saha 4  ← cihaz Kapama 3      saha 29 ← cihaz Kapama 6
#
# 'Kapama 9xx' = SAHA KODU HENÜZ VERİLMEMİŞ presler (modül montajı yapılmadı,
# kullanıcı kodları bildirecek). 9xx bilinçli: gerçek bir saha kodu olamaz ve
# yeni 'Kapama 12' ile ÇAKIŞMAZ. Son iki hane firmware kanalıdır (907 → cihaz
# 'Kapama 7'). Bunlar operatör listesinde GİZLİDİR (bkz. TEL_GIZLI_HATLAR).
#
# SIRA DEĞİŞMEZ — YALNIZ AD DEĞİŞTİ: pozisyon = uretim_kayitlari.istasyon.
# Liste kısaltılsaydı Son Montaj hatları 19..22'den 13..16'ya kayar ve yazılmış
# kayıtlar başka hattı göstermeye başlardı. Yeni hat HEP SONA eklenir.
TEL_HATLARI = (
    ['Halat Kesme %d' % i for i in range(1, 4)]
    + ['Yarı Otomatik %d' % i for i in range(1, 3)]
    + ['Tam Otomatik']
    + ['Kapama 5', 'Kapama 12', 'Kapama 4', 'Kapama 30', 'Kapama 28', 'Kapama 29']
    + ['Kapama %d' % i for i in range(907, 913)]
    + ['Son Montaj %d' % i for i in range(1, 5)]
)

# Saha kodu bekleyen (modülü henüz takılmamış) kapama presleri — operatör
# listesinde gösterilmez. Kod geldikçe yukarıdaki adı değiştirip buradan çıkarın;
# POZİSYONU KORUYUN (istasyon numarası kaymasın).
TEL_GIZLI_HATLAR = frozenset('Kapama %d' % i for i in range(907, 913))

_SON_NUMARA = re.compile(r'\s*\d+$')


def tel_hat_adimi(robot_no):
    """Hat adı → PROSES ADIMI. 'Kapama 3' → 'Kapama', 'Yarı Otomatik' → kendisi.

    Son adım karşılaştırması hat adına DEĞİL adım tipine bakar; aksi hâlde
    'Son Montaj 2'de çalışan operatörün kaydı 'Son Montaj' ile eşleşmez ve o
    referansın üretimi hiç sayılmazdı. Tanınmayan hat → None."""
    ad = str(robot_no or '').strip()
    if not ad:
        return None
    if ad in TEL_ADIM_SIRA:
        return ad
    kok = _SON_NUMARA.sub('', ad).strip()
    return kok if kok in TEL_ADIM_SIRA else None


# ── PROSES ADIMI EKİ (kullanıcı 2026-08-04) ─────────────────────────────────
# "Kesim hattındayken seçilen referansın sonuna KESIM yazılır; böylece son
#  operasyon olmadığı anlaşılır ve diğer hatlardaki işlemlerle karışmaz."
#
# HER ADIM EK ALIR — 'son adım ek almaz' kuralı YOK (2026-08-04 revizyonu).
# Sebep (kullanıcı): "Kapaması ya da son montajı dışarıda olacak ürünler belli
# değil; üretim yoğunluğuna ve tedarikçi durumuna göre değişiyor." Yani bir
# referansın SON adımı sabit değildir, referans bazında tanımlanamaz.
# İkinci sebep: TK1'de üretim teyidini KALİTE departmanı veriyor → base kodun
# ERP'deki article ile birebir kalması gerekmiyor, ek serbestçe kullanılabilir.
#
#   93.TK.464 @ Halat Kesme 1 → '93.TK.464 KESIM'
#   93.TK.464 @ Kapama 2      → '93.TK.464 KAPAMA'
#   93.TK.464 @ Son Montaj 3  → '93.TK.464 SON MONTAJ'
#
# Sonuç: her adım kendi kodunda toplanır, hatlar birbirine karışmaz, aynı iş
# birden çok kez "üretim" sayılmaz ve hangi ürünün nerede bittiği raporda
# kendiliğinden görünür — sabit bir proses tanımına ihtiyaç kalmaz.
TEL_ADIM_EKI = {
    'Halat Kesme':   'KESIM',
    'Yarı Otomatik': 'YARI OTOMAT',
    'Tam Otomatik':  'TAM OTOMAT',
    'Kapama':        'KAPAMA',
    'Son Montaj':    'SON MONTAJ',
}
# Ek zaten yazılmış mı? (operatör elle yazdıysa iki kez eklenmesin)
_EK_DESEN = re.compile(
    r'\s+(KESIM|YARI\s*OTOMAT|TAM\s*OTOMAT|KAPAMA|SON\s*MONTAJ)\s*$', re.IGNORECASE)


def tel_ek_ayikla(referans_kodu):
    """Koddaki ara-operasyon ekini atar: '93.TK.464 KESIM' → '93.TK.464'."""
    return _EK_DESEN.sub('', str(referans_kodu or '').strip()).strip()


# Ek → adım ters haritası. Rapor katmanı adımı KODDAN çözer (hat/istasyon
# gerekmez): mail_raporu.py app.py'yi import EDEMEZ ama kodu görebilir.
_EK_ADIM = {ek: adim for adim, ek in TEL_ADIM_EKI.items()}


def tel_koddan_adim(referans_kodu):
    """'93.TK.464 KAPAMA' → 'Kapama'. Ek yoksa None (tel dışı ya da eksiz kayıt).

    Adım bazlı rapor kırılımının temeli: aynı ürünün farklı adımları AYRI kodlarda
    durduğu için, base kod + adım ikilisi 'aynı 100 parça 4 adımdan geçti'yi
    '400 üretim' gibi göstermeden çözer."""
    m = _EK_DESEN.search(str(referans_kodu or ''))
    if not m:
        return None
    # Ek yazımı esnek olabilir ('SON  MONTAJ') → boşlukları tekille, büyüt
    ek = ' '.join(m.group(1).split()).upper()
    return _EK_ADIM.get(ek)


def tel_referans_kodu(referans_kodu, robot_no):
    """Bu hatta kaydedilecek NİHAİ referans kodunu üretir: '93.TK.464 KAPAMA'.

    Her proses adımı kendi ekini alır (bkz. TEL_ADIM_EKI açıklaması). Operatör
    eki elle yazdıysa tekrarlanmaz — önce ayıklanır, sonra hattın doğru eki
    konur; yanlış hatta yazılmış ek de böylece kendiliğinden düzelir.
    Hat tanınmıyorsa (tel dışı bir robot_no) koda DOKUNULMAZ."""
    base = tel_ek_ayikla(referans_kodu)
    if not base:
        return str(referans_kodu or '').strip()
    ek = TEL_ADIM_EKI.get(tel_hat_adimi(robot_no))
    return f'{base} {ek}' if ek else base


def tel_adim_etiketi(robot_no):
    """Hat → rapor/panel için okunabilir adım etiketi ('Kapama 3' → 'Kapama').
    Tel dışı hatlarda '' döner."""
    return tel_hat_adimi(robot_no) or ''
