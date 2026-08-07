# TK1 — TEL ÜRETİMİ · Proses Akışı ve Sayım Modeli

**Durum:** 2026-08-04 · TK1'de ayrı bölüm (`bolum='tel'`, `lokasyon='TK1'`)
**AS400 teyidi:** yok (şimdilik yalnız takip — kullanıcı kararı)

---

## 1. Akış

Bir tel referansı sırayla birden çok hattan geçer. **Her adım ayrı hatta, ayrı
operatör tarafından** kaydedilir:

```
   launch alınan referans
        │
   ┌────▼─────────┐   ┌──────────────┐   ┌──────────┐   ┌─────────────┐
   │ HALAT KESME  │──►│ YARI / TAM   │──►│  KAPAMA  │──►│ SON MONTAJ  │
   │  3 makine    │   │  OTOMATİK    │   │ 4 makine │   │  4 makine   │
   │              │   │  1 + 1 makine│   │          │   │  (HER ÜRÜNDE│
   └──────────────┘   └──────────────┘   └──────────┘   │   YOK)      │
    operatör A          operatör B        operatör C    └─────────────┘
                                                          operatör D
```

**Son montaj her üründe yoktur** — fasona verilebilir ya da o telin prosesinde
hiç olmayabilir. O ürünlerde akış kapamada biter.

**Yarı ve tam otomatik aynı sıradadır** (2. adım): biri diğerinin alternatifidir,
ikisi arka arkaya uygulanmaz.

### Hatlar (13)

| Adım | Makine sayısı | Hat adları |
|---|---|---|
| Halat Kesme | 3 | `Halat Kesme 1` · `Halat Kesme 2` · `Halat Kesme 3` |
| Yarı Otomatik | 1 | `Yarı Otomatik` |
| Tam Otomatik | 1 | `Tam Otomatik` |
| Kapama | 4 | `Kapama 1` … `Kapama 4` |
| Son Montaj | 4 | `Son Montaj 1` … `Son Montaj 4` |

Tek makineli adımlarda numara yoktur. Yeni makine eklenirse **sona** eklenir ve
adım adı korunur (`Kapama 5`) — sistem hat adının sonundaki numarayı atarak adım
tipini bulur (`tel_hat_adimi`).

---

## 2. Çoklu sayım nasıl önleniyor — proses adımı eki

Aynı tel dört adımda kaydedilir. Adımlar **aynı kodla** toplansaydı 100 adetlik iş
raporda 385 görünürdü. Çözüm: **her adım kendi ekiyle kaydedilir.**

```
93.TK.464 @ Halat Kesme 1  →  93.TK.464 KESIM
93.TK.464 @ Yarı Otomatik  →  93.TK.464 YARI OTOMAT
93.TK.464 @ Kapama 2       →  93.TK.464 KAPAMA
93.TK.464 @ Son Montaj 3   →  93.TK.464 SON MONTAJ
```

Ek **otomatik** eklenir (operatör elle yazmaz — yazım hatası referansı bambaşka bir
koda çevirip kaydı kaybederdi). Operatör yine de yazarsa tekrarlanmaz; yanlış hatta
yazılmış ek de hattın doğrusuna çevrilir.

Sonuç: her adım ayrı kodda toplanır → ne çoklu sayım olur ne de kimin ne yaptığı
kaybolur. Aynı adımın **farklı makinelerine bölünmüş** iş ise doğal olarak tek kodda
birleşir (`Kapama 1` + `Kapama 4` → `… KAPAMA`).

### Neden referans bazlı "son adım" tanımı YOK

İlk tasarımda üretim yalnız referansın son adımından sayılıyor ve akış
`referans_listesi.tel_adimlar`'da tanımlanıyordu. **Bu terk edildi** (2026-08-04):

> Kapaması ya da son montajı dışarıda yapılacak ürünler **belli ve sabit değil** —
> üretim yoğunluğuna ve tedarikçi durumuna göre değişiyor.

Yani bir referansın son adımı önceden tanımlanamaz. Ayrıca TK1'de üretim teyidini
**kalite departmanı** verdiği için base kodun ERP article'ıyla birebir kalma
zorunluluğu da yok — ek serbestçe kullanılabiliyor.

`tel_adimlar` kolonu veritabanında duruyor ama **kullanılmıyor**; Excel'de
`Tel Referanslar` sayfasına da gerek yok.

---

## 3. Referans ve operatörler nereden geliyor

Ek bir sayfa gerekmez, ikisi de kuralla ayrışır:

- **Referans:** TK1'de `93.` ile başlayan **her kod** tel referansıdır
  (içe aktarımda 1459 kod tel bölümüne alındı).
- **Operatör:** ana `Operatörler` sayfasındaki isimlerden `BİRCAN KILIÇ` ve
  `OSMAN İMAT` montaj, **kalan herkes tel** (23 kişi). Plastik operatörleri kendi
  sayfasından gelir.

Kural `import_excel.py`'de; içe aktarım **idempotent** (tekrar çalıştırmak güvenli).

---

## 4. TK1 bölüm yapısı (2026-08-04 itibarıyla)

| Bölüm | Hatlar | Not |
|---|---|---|
| **Montaj** | `LF-LFP` · `Iveco` · `TK1-M1…M7` masaları | Masalarda buton+buzzer sayaç modülü |
| **Plastik Enjeksiyon** | `320T` · `407T` · `Yapistirma` · `Sizdirmazlik Test` | Sızdırmazlık adedi Cofle SVP'den |
| **Tel Üretimi** | 13 proses hattı (yukarıda) | Sayaç yok, elle giriş |

> **Pull / Push-Pull** montaj hat listesinden çıkarıldı (2026-08-04): bunlar
> **tel ürünü adları**, hat değil. Geçmiş vardiyalar bu isimlerle kayıtlı kaldığı
> için `TK1_ROBOT_NOLARI` kümesinde duruyorlar — eski kayıtlar filtrelerden düşmez.

---

## 5. Raporlama

- **Günlük mail raporu**: her operatörün her adımdaki işi ayrı satır olarak görünür
  (`KESIMCI A · 93.TK.464 KESIM · 200`). Süzgeç yok — ayrışma kodun kendisinden gelir.
- **Üretim & OEE → Operatör Bazlı Üretim**: operatör × makine × proses × referans
  kırılımı. Tel'de "Makine" ve "Proses" ayrı sütun (`Kapama 3` / `Kapama`).
- **KPI "Toplam İşlem (tüm adımlar)"**: tel bölümünde bu sayı bitmiş ürün değil,
  tüm adımlarda yapılan toplam iştir — etiket bunu açıkça söyler.
- **OEE**: vardiya bazında hesaplanır, yani her operatör **kendi adımının**
  verimliliğiyle ölçülür. Son-adım süzgeci uygulanmaz (kesim operatörünün üretimini
  sıfırlamak yanlış olurdu).

## 6. Kısıtlar ve eksikler

- **Kapama = tek operatör tek iş**: aynı gün aynı referansa ikinci kapama kaydı
  409 ile reddedilir. Diğer adımlar bölünebilir.
- Tel hatlarında **sayaç/andon yok** — tüm girişler elle.
- **İş takibi** ve **AS400 teyidi** tel bölümünde yok (teyidi kalite departmanı verir).
- Tel referanslarında **cycle time tanımlı değil** → OEE performans bileşeni 0 çıkar.
  Adım bazlı süre gerekirse (kesim 30 sn, kapama 60 sn…) ayrı tanım şart olur.
