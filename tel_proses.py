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

# FİZİKSEL HATLAR (kullanıcı 2026-08-04): kesim 3 · yarı otomatik 1 · tam otomatik 1
# · kapama 12 · son montaj 4 = 21 hat. Tek makineli adımlarda numara YOK.
# KAPAMA 4 → 12 (2026-08-07): kapama hattında 12 pres var; sayaç modülleri
# takılırken gerçek makine sayısı ortaya çıktı (eski 4 rakamı eksikti). Presler
# üçerli ortak panolarda; 4 ESP32 modülü 3'er presi okur (bkz. pilot/firmware/
# _templates/tel_kapama.ino.tpl). Hat adları firmware'deki robot_no ile BİREBİR.
TEL_HATLARI = (
    ['Halat Kesme %d' % i for i in range(1, 4)]
    + ['Yarı Otomatik', 'Tam Otomatik']
    + ['Kapama %d' % i for i in range(1, 13)]
    + ['Son Montaj %d' % i for i in range(1, 5)]
)

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
