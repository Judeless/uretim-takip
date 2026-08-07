# TK1 — PLASTİK ENJEKSİYON BÖLÜMÜ · Veri Akışı ve Bağlantı Şeması

**Durum:** 2026-08-04 · TK1'de montajdan **ayrı bölüm** (`bolum='plastik'`, `lokasyon='TK1'`)

Bu bölümün altında **4 makine** var. Yapıştırma ve sızdırmazlık testi **ayrı bölüm değildir**,
bu bölümün makineleridir — operatör vardiya açarken hangisinde çalışacağını seçer.

---

## 1. Makineler ve adet kaynağı

| Makine (`robot_no`) | Ne yapar | Adet nereden gelir | Firmware / kaynak |
|---|---|---|---|
| `320T` | Plastik enjeksiyon | ESP32 sayaç (röle sinyali) | `cofle_sayac_320t` |
| `407T` | Plastik enjeksiyon | ESP32 sayaç (röle sinyali) | `cofle_sayac_407t` |
| `Yapistirma` | Yapıştırma | ESP32 sayaç (**kısa** pulse) | `cofle_sayac_yapistirma` |
| `Sizdirmazlik Test` | Sızdırmazlık testi | **Saha sayacı YOK** → Cofle SVP | `cofle_test.py` (API poller) |

> `robot_no` değerleri **ASCII** yazılır (`Yapistirma`, `Sizdirmazlik Test`). Türkçe karakter
> firmware klasör adlarında, ODBC sorgularında ve rapor kırılımlarında sorun çıkarıyor.

### Referans kodları

Bölümün referansları `data/Tk1 Veriler.xlsx` → **`Plastik Enjeksiyon Referanslar`** sayfasında;
operatörleri **`Plastik Enjeksyion Operatörler`** sayfasında (sayfa adındaki yazım hatası
Excel'de böyle — kod da bu adı arıyor, düzeltilirse `import_excel.py` de güncellenmeli).

**Sonu `G` ile biten kodlar yapıştırma makinesinin ürünleridir:**

```
10.300.1281G   10.300.3746G   10.300.3581G
10.300.3747G   10.300.3294G   10.300.2452G
```

Kalan kodlar (`10.300.1444`, `10.300.2452A/B`, `10.300.3292/3293`, `10.300.3412/3413`, …)
enjeksiyon makinelerinin ürünleri. Şu an sistem bu ayrımı **zorlamıyor**: operatör hangi
makinede olursa olsun listedeki her kodu seçebiliyor. Makineye göre kod filtresi istenirse
ayrıca eklenir (kararı sende — yanlış makineye kayıt riskini azaltır ama esnekliği kısar).

---

## 2. Sızdırmazlık test makinesi — SVP veri akışı

Bu makinede sahaya sensör takılmıyor. Adet, **başarılı test = 1 adet** kuralıyla
Cofle SVP'den (wicow) geliyor. Zincir zaten kurulu ve montajda çalışıyor:

```
 ┌──────────────────┐   operatör testi yapar    ┌────────────────────────┐
 │ Sızdırmazlık     │ ─────────────────────────►│  Cofle SVP (wicow)     │
 │ test cihazı      │                           │  beta.coflesvp.com     │
 └──────────────────┘                           └───────────┬────────────┘
                                                            │ REST (Bearer token)
                                                            │ apibeta.wicow.io/cofle
                                                            ▼
                                          ┌──────────────────────────────┐
                                          │ cofle_test.py  (poller)      │
                                          │ scheduler.py · 20 sn'de bir  │
                                          │ yalnız AKTİF oturumların     │
                                          │ cihazlarını çeker            │
                                          └───────────┬──────────────────┘
                                                      │ UPSERT (test_id benzersiz)
                                                      ▼
                                          ┌──────────────────────────────┐
                                          │ uretim.db · test_sonuclari   │
                                          └───────────┬──────────────────┘
                                                      │ _test_basari_say()
                                                      │ (referans başlangıcından beri PASSED)
                                                      ▼
                                          ┌──────────────────────────────┐
                                          │ uretim_kayitlari             │
                                          │  · test_cihaz_id = <cihaz>   │
                                          │  · sayac_otomatik = 1        │
                                          │  · ok_adet ← canlı sayım     │
                                          └──────────────────────────────┘
```

**Operatör akışı (/tk1):**
1. Bölüm: **Plastik Enjeksiyon** → Makine: **Sizdirmazlik Test** → vardiyayı başlat
2. Üretim kaydı ekle → referans seç → **"Test cihazı ile çalışıyorum"** kutusunu işaretle
3. Açılan listeden sızdırmazlık test cihazını seç → **ADET boş bırakılır**, canlı dolar
4. Operatör ADET'i elle değiştirirse sayaç o kayıt için **donar** (manuel düzeltme kabul edilir)

**Ayar dosyası:** `cofle_test_config.json` (git'e girmez) — `api_base`, `token`,
`poll_aralik_sn`, `etkin`. Token'ın süresi dolarsa beta.coflesvp.com'dan yenisi alınıp
buraya yazılır; poller token'sız sessizce boş döner, uygulama çökmez.

### ⏳ Senden beklenen tek bilgi

SVP tarafında **hangi cihazın** sızdırmazlık test makinesi olduğu. Cihaz listesi
`/api/test_cihazlari` ile geliyor ve operatör dropdown'ında görünüyor; doğru cihaz adını
söylersen:
- listede **yalnız o cihaz** ön-seçili gelecek şekilde bağlarım (yanlış cihaz seçme riski biter),
- istersen o makinede test cihazı seçimini **zorunlu** yaparım (adet elle girilemez).

Şu an bağlama yapılmadığı için operatör listeden **doğru cihazı kendisi seçmeli**.

---

## 3. Enjeksiyon + yapıştırma makineleri (saha sayacı)

Bu üç makine ESP32 sayaç modülüyle çalışır — bağlantı ve firmware ayrıntıları
[BAGLANTI_SEMASI_ABKANT_PLASTIK_YAPISTIRMA.md](BAGLANTI_SEMASI_ABKANT_PLASTIK_YAPISTIRMA.md)
dosyasında. Özet:

```
  MAKİNE RÖLESİ / SİNYALİ            ESP32-WROOM-32U           SUNUCU
 ┌────────────────────┐            ┌────────────────┐        ┌──────────────┐
 │ kuru kontak / 24V  │──[opto]───►│ GPIO25  (sayım)│───WiFi►│ pilot :5001  │
 │ çevrim sinyali     │            │ GND            │        │ sayac_olaylari│
 └────────────────────┘            └────────────────┘        └──────┬───────┘
                                                                    │ bolum+robot_no
                                                                    ▼
                                                            uretim_kayitlari.ok_adet
```

> **Sayım aktif vardiyaya kapılıdır:** o makinede açık vardiya yoksa gelen pulse üretime
> yazılmaz. Makine kapalıyken hat parazitinin hayalet sayım üretmesi böyle engellenir.

---

## 4. Sistem tarafında yapılanlar (2026-08-04)

- **Operatör mobil (`/tk1`)**: Plastik Enjeksiyon bölümü seçilebiliyor; makine listesi
  `320T / 407T / Yapistirma / Sizdirmazlik Test`. Test cihazı bloğu artık montajın yanı sıra
  **plastik** bölümünde de görünüyor.
- **Excel içe aktarma**: `import_tk1()` tek çağrıda montaj **ve** plastik sayfalarını
  aktarıyor (`bolum='plastik'`, `lokasyon='TK1'`). Başlık satırı içerikten tanınıyor —
  `Parça Kodu` / `Ürün Kodu` gibi başlıklar referans olarak yazılmıyor, `TEST` veya
  `ÖZEL ÜRÜN` gibi **rakamsız ama gerçek** kodlar korunuyor.
- **Dashboard**: TK1 artık iki bölümlü. Fabrika Özeti iki bölümü de gösteriyor, bölüm
  seçici TK1'de görünür oldu, Üretim & OEE / Kayıtlar / Referanslar / Operatörler sayfaları
  plastik bölümüne geçilebiliyor. Sinyal Analizi'nde TK1 cihaz listesi bölüme göre geliyor
  (montaj: YF1 + TK1-M1…M7 · plastik: 320T / 407T / Yapistirma).
- **İş takibi ve andon** plastikte **yok** — bunlar kaynak/montaj/metal bölümlerine özel.
