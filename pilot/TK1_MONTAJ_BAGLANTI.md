# TK1 — Montaj Masası Sayaç Bağlantı Şeması (buton + buzzer)

> Röleli makine şemaları için: [BAGLANTI_SEMASI_ABKANT_PLASTIK_YAPISTIRMA.md](BAGLANTI_SEMASI_ABKANT_PLASTIK_YAPISTIRMA.md) ·
> TK1 plastik/yapıştırma: [TK1_PLASTIK_YAPISTIRMA_BAGLANTI.md](TK1_PLASTIK_YAPISTIRMA_BAGLANTI.md)

TK1 (yan tesis) montaj masalarında **makine sinyali yoktur** — üretimi operatör kendisi bildirir.
Bu yüzden şema röleli makinelerden farklıdır: **1 buton + 1 buzzer**. Operatör her parçayı
bitirdiğinde butona basar, cihaz "bip" ile onaylar ve sunucuya 1 pulse yazar.

## Cihazlar (7 masa)

| Masa | `robot_no` | `cihaz_id` (OTA adı) | Firmware klasörü |
|---|---|---|---|
| 1 | `TK1-M1` | `MONTAJ-TK1-M1` | `firmware/cofle_sayac_tk1_m1` |
| 2 | `TK1-M2` | `MONTAJ-TK1-M2` | `firmware/cofle_sayac_tk1_m2` |
| 3 | `TK1-M3` | `MONTAJ-TK1-M3` | `firmware/cofle_sayac_tk1_m3` |
| 4 | `TK1-M4` | `MONTAJ-TK1-M4` | `firmware/cofle_sayac_tk1_m4` |
| 5 | `TK1-M5` | `MONTAJ-TK1-M5` | `firmware/cofle_sayac_tk1_m5` |
| 6 | `TK1-M6` | `MONTAJ-TK1-M6` | `firmware/cofle_sayac_tk1_m6` |
| 7 | `TK1-M7` | `MONTAJ-TK1-M7` | `firmware/cofle_sayac_tk1_m7` |

- Bölüm: `montaj` · sinyal tipi: `kaynak_tip="buton"` · POST `/api/sinyal`, heartbeat 30 sn.
- **Her masaya KENDİ .ino'su yüklenir** — yanlış dosya yüklenirse üretim başka masaya yazılır.
- Sahadaki eski **YF1** modülü yerinde kalır (LF-LFP hattına bağlı), bu 7 cihaz onun yanına eklenir.
- Firmware `_templates/montaj.ino.tpl` + `generate.py` ile üretilir — tek tek `.ino` düzenleme.

---

## 1. GÜÇ — 5 V USB adaptör

Montaj masasında makine panosu/24 V yok; cihaz **prizden 5 V USB adaptörle** beslenir.

```
   220V PRIZ            USB ADAPTOR (5V 1A+)         ESP32-WROOM-32U
 ┌──────────┐          ┌──────────────────┐        ┌────────────────┐
 │          ┼─────────►│  micro-USB kablo ┼───────►│ USB girisi     │
 └──────────┘          └──────────────────┘        └────────────────┘
```

- Telefon şarj adaptörü yeterli (**5 V / en az 1 A**). ESP32 ~250 mA, buzzer anlık ~30 mA çeker.
- Masada 24 V varsa LM2596 buck ile de beslenebilir (bkz. röleli şema §1) — **pot önce 5.0 V'a ayarlanır**.
- ❌ 5 V'u **asla `3V3` pinine verme.**
- Adaptörü masanın **kesintiye uğramayan** prizine tak; makine şalteriyle sönen prize takarsan vardiya
  ortasında cihaz kapanır. (Kapanma veri kaybettirmez — sayım NVS'te, açılışta gönderilmemiş
  pulse'lar tekrar gönderilir — ama kapalıyken basılan butonlar kaydedilmez.)

---

## 2. BUTON — GPIO25

```
   BUTON (NO, momentary)                 ESP32-WROOM-32U
 ┌─────────────────────┐              ┌────────────────┐
 │  kontak ucu 1 ───────┼────kablo────►│ GPIO25         │  INPUT_PULLUP (bosta HIGH=3.3V)
 │  kontak ucu 2 ───────┼────kablo────►│ GND            │
 └─────────────────────┘              └────────────────┘
        (basinca GPIO25 GND'ye iner -> LOW -> 1 sayim, buzzer "bip")
```

- **NO (normalde açık) anlık buton** — kalıcı/kilitli (latching) buton **KULLANMA**, basılı kalırsa
  tek sayım verir ve operatör bir daha sayamaz.
- Ø22 pano tipi metal buton ideal (montaj masasına vidalanır, eldivenle basılır).
- Polarite yok; iki uç ters takılabilir.
- Kablo 2×0.5 mm² yeter. Uzun (>3 m) çekilecekse motor/kaynak kablolarıyla aynı kanaldan geçirme;
  istersen GPIO25–GND arasına **100 nF** kondansatör (firmware integratörü zaten filtreliyor).

---

## 3. BUZZER MODÜLÜ (3 pinli, VCC / I-O / GND)

Aldığımız hazır buzzer kartı (YL-44 tipi): üzerinde transistör var, GPIO'yu yormaz ve
**5 V ile beslenince sesi yüksek çıkar**. Kartın üzerinde `VCC · I/O · GND` yazar.

```
   BUZZER MODUL                          ESP32-WROOM-32U
 ┌──────────────┐                     ┌────────────────┐
 │  VCC  ────────┼────────────────────►│ 3V3            │   <-- 5V DEGIL (asagidaki nedene bak)
 │  I/O  ────────┼────────────────────►│ GPIO27         │   <-- tetik (basista ~200ms)
 │  GND  ────────┼────────────────────►│ GND            │
 └──────────────┘                     └────────────────┘
```

> ⚠️ **VCC = `3V3` pini — `5V` DEĞİL** (sahada belirlendi, 2026-07-31).
> Bu kart **aktif-LOW**: girişi HIGH iken susar. 5 V'tan beslendiğinde "HIGH" eşiği 5 V'a
> göre kurulur, ESP32'nin 3.3 V'luk HIGH'ı girişi tam kapatamaz ve buzzer **sürekli öter**.
> VCC `3V3`'e alınınca HIGH = VCC olur, giriş kesin kapanır. Ses bir miktar kısılır ama
> montaj masasında yeterli — TK1-M1'de böyle çalışıyor.
>
> **Sesin yetmediği gürültülü masa olursa:** VCC'yi `5V`'a geri al, **I/O ile 5V arasına
> 2.2 kΩ pull-up** tak ve GPIO27'yi *open-drain* sür (firmware'de cihaz bazlı bayrak
> gerekir — istenirse eklenir). Direnci takmadan open-drain sürüm yüklenmez: kapalıyken
> hat boşta kalır ve rastgele ötebilir.

**Pin sırasına dikkat:** bu kartların bir kısmında silkscreen sırası `I/O · VCC · GND`
(ortadaki pin VCC). Teli konuma göre değil, **kartın üzerindeki yazıya göre** tak — I/O ile
VCC yer değiştirirse modül sürekli beslenir ve firmware ne yaparsa yapsın öter.

### Polarite — TK1 modülleri **aktif-LOW** (sahada belirlendi, 2026-07-31)

İlk denemede modül sürekli öttü → aldığımız kart **aktif-LOW** (I/O LOW iken öter).
Firmware'de bu ayar **cihaz bazlıdır**, `generate.py` içinde:

```python
('MONTAJ-TK1-M1', 'TK1-M1', {'BUZZER_AKTIF': 'false'}),   # aktif-LOW modul
```

> ⚠️ **`montaj.ino.tpl`'deki sabiti topluca değiştirme.** Aynı şablondan TK2'nin M1–M12
> ve YF1 modülleri de üretiliyor; onların buzzer'ı **aktif-HIGH** çalışıyor. Global
> çevirirsen ilk OTA'da sahadaki 13 modül sürekli ötmeye başlar.

| Gözlem | Anlamı | Yapılacak |
|---|---|---|
| Butona basınca **bip**, boşta sessiz | Doğru | Bir şey yapma |
| Boşta **sürekli ötüyor**, firmware `aktif-HIGH` diyor | Polarite ters | O cihazın `BUZZER_AKTIF` değerini çevir, `python generate.py`, tekrar yükle |
| Firmware `aktif-LOW` olduğu hâlde **sürekli ötüyor** | VCC 5V'ta; 3.3 V HIGH girişi kapatamıyor | Modül `VCC`'sini **`3V3`** pinine al — TK1'de çözüm bu oldu |
| Hiç ötmüyor | I/O yanlış pinde ya da VCC bağlı değil | §3 şemasını + pin sırasını kontrol et |

Çalışan sürüm Serial'da (115200) şunu yazar — yükleme sonrası tek bakışta doğrulanır:
`Buzzer: aktif-LOW (pin GPIO27, bosta HIGH surulur, desen 2x300ms)`

**Ayırt edici test (kablo hatası mı, polarite mi):** I/O ucunu GPIO27'den çıkarıp doğrudan
modülün `VCC`'sine değdir. Susuyorsa kart aktif-LOW'dur (yukarıdaki ayar); hâlâ ötüyorsa
sorun kartta/kabloda — VCC ile I/O ters bağlanmış olabilir.

Açılışta `setup()` pini sürene kadar (<1 sn) kısa bir ses gelebilir, normaldir.

### Bip deseni (2026-07-31 — cihaz bazlı, fw 2.7.3+)

3V3 beslemede ses kısıldığı için TK1 masaları **çift bip** ile onaylar; TK2 değişmedi:

| Cihazlar | Desen | Neden |
|---|---|---|
| TK1-M1 … TK1-M7 | **2 bip × 300 ms** (120 ms ara) | 3V3'te ses kısık → daha uzun + tekrarlı, gürültüde seçilir |
| TK2 M1–M12 + YF1 | 1 bip × 200 ms | Sahadaki alışkanlık; 5 V/aktif-HIGH kurulumda ses zaten yüksek |

Değerler `generate.py` DEVICES satırında (`BUZZER_BEEP_ADET` / `BUZZER_BEEP_MS`) — şablonda
topluca değiştirilmez. Desen donanım timer'ında yürür (HTTP gecikmesi süreyi uzatmaz) ve
toplamı (720 ms) `MIN_PULSE_GAP_MS`'ten (1000 ms) kısadır: bir sonraki sayım gelmeden desen
bitmiş olur, bip'ler üst üste binmez.

---

## 4. PİN ÖZETİ

| ESP32 Pin | İşlev |
|---|---|
| **GPIO25** | Buton — `INPUT_PULLUP`, aktif LOW |
| **GPIO27** | Buzzer modül `I/O` girişi — TK1 modüllerinde **aktif-LOW** |
| **3V3** | Buzzer modül `VCC` (**5V değil** — bkz. §3) |
| **GND** | Buton 2. ucu + buzzer `GND` (ortak) |
| **GPIO26** | Boş (`INPUT_PULLUP`) |
| **GPIO2** | Dahili LED — WiFi yokken yanıp söner |

---

## 5. FİRMWARE FİLTRESİ (neden yanlış saymaz)

| Parametre | Değer | Anlamı |
|---|---|---|
| `INTEG_YUKSEK_ESIK` | 15 (~75 ms) | Bu kadar net basılmayan temas sıçraması sayılmaz |
| `INTEG_DUSUK_ESIK` | 4 (~20 ms) | Basış bitti kabulü (histerezis) |
| `MIN_PULSE_GAP_MS` | 1000 ms | İki sayım arası en az 1 sn → çift basış tek sayılır (en fazla 60 parça/dk) |

Operatör 1 sn'den hızlı basarsa sayılmaz ve heartbeat'te `tani_erken` artar; sahada gerçekten
daha hızlı üretim varsa bu değer düşürülür (şablonda, sonra `generate.py`).

---

## 6. MALZEME LİSTESİ (masa başına)

| # | Malzeme | Not |
|---|---|---|
| 1 | ESP32-WROOM-32U DevKit (38-pin) | Ana kart |
| 2 | Buzzer modülü (3 pin, VCC/I-O/GND) | **3V3'ten** beslenir (bkz. §3 — 5 V'ta sürekli öter) |
| 3 | Ø22 NO anlık buton | Kalıcı/latching OLMAYACAK |
| 4 | 5 V USB adaptör + kablo | En az 1 A |
| 5 | Kablo 2×0.5 mm² | Buton hattı |
| 6 | Küçük plastik kutu | Masaya sabitlenir, buton dışarıda |

7 masa için: 7 ESP32 + 7 buzzer + 7 buton + 7 adaptör.

---

## 7. KURULUM + TEST

1. Masanın numarasını belirle → **o masanın** `.ino`'sunu aç (§ tablosu).
2. Arduino IDE → ESP32-WROOM-32U (DOIT ESP32 DEVKIT V1) → USB'den yükle.
3. Serial Monitor (115200) — şunları görmelisin:
   `COFLE PILOT SAYAC — MONTAJ-TK1-M3` · `MONTAJ BUTON MODU` · `[WiFi] OK · IP=...` · `[READY]`
4. **Kuru test:** GPIO25 ↔ GND arasına tel değdir → Serial'da `[PULSE] Ist.1 #1`, buzzer bip.
5. Butonu bağla, 5 kez bas → Serial'da 5 pulse, dashboard **Saha Cihazları**'nda `MONTAJ-TK1-M3`
   satırı **bağlı** ve bugünkü sayımı 5 olmalı.
6. Operatör tarafı: `/tk1` mobilinde bölüm **Montaj**, hat listesinden **`TK1-M3 Masası (sayaçlı)`**
   seçilir. Vardiya açıldıktan sonra üretim kaydında adet **otomatik** gelir.
7. Sonraki yüklemeler OTA ile: `firmware/OTA_Yukle.bat` (cihaz adı = `cihaz_id`, şifre `cofle-ota-2026`).

> **Ağ:** SSID `COFLE-TK`, sunucu `http://192.168.21.155:5001`. TK1'de zaten YF1 modülü bu ağla
> çalışıyor; yeni masalarda RSSI zayıfsa (heartbeat'te `wifi_rssi` < −80 dBm) AP mesafesini kontrol et.

---

## 8. SUNUCU TARAFI (kayıtlı ayarlar — bilgi)

Bu 7 cihaz sisteme tanıtıldı; ek ayar gerekmez:

| Yer | Ne yapar |
|---|---|
| `app.py` · `SAYAC_AUTO_CIHAZLAR` | Operatör mobilinde **otomatik sayaç** modunu açar |
| `app.py` · `TK1_ROBOT_NOLARI` | Dashboard'da TK1 tesisine ait sayılır |
| `app.py` · `BEKLENEN['montaj']` | Saha Cihazları listesinde `lokasyon=TK1` ile görünür |
| `app.py` · TK1 hat listesi + `mobile_v2.html` | Operatör hat seçiminde `TK1-M1..M7` çıkar |

Masa = hat modeli (TK2'deki M1..M12 ile aynı): vardiyanın hattı cihazın `robot_no`'su ile birebir
eşleştiği için ayrıca hat→cihaz eşlemesi gerekmez. Eski 4 hat adı (Pull · Push-Pull · Iveco ·
LF-LFP) listede kaldı — geçmiş vardiyalar ve hat bazlı çalışma için.
