# TK1 — Plastik Enjeksiyon + Yapıştırma Sayaç Bağlantı Şemaları

> Abkant dahil birleşik şema: [BAGLANTI_SEMASI_ABKANT_PLASTIK_YAPISTIRMA.md](BAGLANTI_SEMASI_ABKANT_PLASTIK_YAPISTIRMA.md)

ESP32-WROOM-32U + makine rölesi ile üretim sinyali sayımı.
**Güç: LM2596 buck modülü** (harici USB adaptör DEĞİL — makinenin 24V'undan beslenir).

| Makine | Firmware | Sinyal tipi | Ayar |
|---|---|---|---|
| **Plastik enjeksiyon** (320T, 407T) | `cofle_sayac_320t/407t.ino` | Uzun cevrim rölesi (~sn'ler LOW) | INTEG_YUKSEK=800, GAP=18s |
| **Yapıştırma** | `cofle_sayac_yapistirma.ino` | **Kısa pulse** (~50–200ms) | INTEG_YUKSEK=15, GAP=1500ms |

> Bölüm ikisi de `plastik` (dashboard'da plastik hattı; makine = robot_no ile ayrılır).

---

## ⚠️ 0. ÖNCE ÖLÇ (en kritik kural)

Röle çıkışını multimetreyle ölç — **bağlamadan önce**:

- **Kuru kontak** (potansiyelsiz; aktifken sadece süreklilik/0V) → **OPTOSUZ** şema (§2 / §4).
- **Gerilim çıkışı** (aktifken 12V/24V okuyor) → **OPTOLU** şema (§3). Doğrudan bağlarsan **ESP32 yanar**.

Bizim döküm/metal makinelerinde röle kuru kontak; plastikte de büyük ihtimalle öyle ama **teyit et.**

---

## 1. LM2596 GÜÇ (her iki makinede AYNI)

LM2596 = ayarlanabilir buck (düşürücü) konvertör. Makinenin 24V'unu 5V'a indirip ESP32'yi besler.

```
   MAKINE 24V DC                 LM2596 MODUL                 ESP32-WROOM-32U
 ┌──────────────┐            ┌───────────────────┐         ┌────────────────┐
 │  24V (+) ─────┼───────────►│ IN+          OUT+ ┼────────►│ 5V  (VIN)      │
 │  0V  (−) ─────┼───────────►│ IN−          OUT− ┼────────►│ GND            │
 └──────────────┘            │   [pot: 5.0V'a]   │         └────────────────┘
                             └───────────────────┘
```

**KURULUM SIRASI (şart):**
1. LM2596 girişine 24V ver, çıkışa **hiçbir şey bağlama.**
2. Multimetreyle **OUT+ / OUT−** ölç, potu çevirerek **5.0V** yap (5.0–5.2V arası kabul; **>5.5V ESP32'yi bozar**).
3. Güç kes, sonra OUT+ → ESP32 **`5V`/`VIN`** pinine, OUT− → **`GND`** pinine bağla.
4. ❌ **5V'u asla `3V3` pinine verme** (ESP32 yanar). ESP32'nin dahili regülatörü 5V→3.3V yapar.

> LM2596 giriş aralığı 4.5–40V → 24V uygun. ESP32 ~250mA çeker, modül 2–3A verir (bol marj).

---

## 2. PLASTİK ENJEKSİYON — OPTOSUZ (röle KURU KONTAK)

En basit ve tercih edilen bağlantı. Röle zaten izole (kuru kontak), araya bir şey gerekmez.

```
  MAKINE ROLESI (kuru kontak)            ESP32
 ┌───────────────────────┐           ┌────────────────┐
 │  NO (norm. acik) ──────┼──kablo───►│ GPIO25         │  INPUT_PULLUP (bosta HIGH=3.3V)
 │  COM (ortak)     ──────┼──kablo───►│ GND            │
 └───────────────────────┘           └────────────────┘
        (röle her cevrimde NO–COM'u KISA/UZUN kapatır)
```

- Röle kapanınca GPIO25 → GND → **LOW** → firmware sayar. Açıkken pull-up HIGH tutar.
- NO/COM hattı için **shielded (blendajlı) kablo** kullan (motor/ısıtıcı paraziti).
- İstersen ekstra: GPIO25–GND arası **100nF kondansatör** (parazit yumuşatır; firmware integratörü zaten filtreler).

**Pin tablosu:**

| ESP32 | Bağlanan |
|---|---|
| GPIO25 | Röle **NO** |
| GND | Röle **COM** + LM2596 OUT− |
| 5V (VIN) | LM2596 OUT+ |

---

## 3. PLASTİK ENJEKSİYON — OPTOLU (röle 24V GERİLİM veriyorsa)

Röle kuru kontak değil de **24V darbe** veriyorsa (PLC/transistör çıkışı gibi) veya azami koruma isteniyorsa **PC817 optokuplör** ile ara.

```
   24V SINYAL           R = 2.2kΩ          PC817 (DIP-4)              ESP32
 ┌────────────┐        (1/4W)          ┌──────────────────┐      ┌───────────┐
 │ Sinyal(+) ──┼───────[ R ]──────────►│1 anot    kolek. 4├─────►│ GPIO25    │  INPUT_PULLUP
 │             │                       │                  │      │           │
 │ Sinyal 0V ──┼──────────────────────►│2 katot   emit.  3├─────►│ GND       │
 └────────────┘                        └──────────────────┘      └───────────┘
```

- Sinyal 24V gelince: LED yanar → foto-transistör iletir → GPIO25 **LOW** → sayar (aktif-LOW, firmware ile uyumlu).
- **Direnç:** 24V → **2.2kΩ**; 12V → **1kΩ** (LED akımı ~10mA olsun).
- **PC817 pinleri** (nokta/çentik 1. pin): `1=anot 2=katot 3=emitter 4=kollektör`.
- **Polarite:** sinyal "üretimde 24V, boşta 0V" varsayıldı. Makine tersse (boşta 24V) LED yönünü ölçüp doğrula.
- Not: ESP32'yi de aynı 24V'dan (LM2596) besliyoruz; bu durumda opto galvanik izolasyon değil **güvenli 24V→3.3V seviye çevirici** işlevi görür (GPIO'yu 24V'dan korur). Tam izolasyon istersen ESP32'yi AYRI 5V'la besle.

**Ek malzeme (optolu):** 1× PC817, 1× 2.2kΩ direnç.

---

## 4. YAPIŞTIRMA — OPTOSUZ (kısa pulse)

**Elektriksel bağlantı §2 ile AYNI** (röle NO → GPIO25, COM → GND, LM2596 güç). Tek fark **firmware**: yapıştırma rölesi kısa pulse verdiği için `cofle_sayac_yapistirma.ino` kısa-pulse eşiğiyle derlenir (INTEG_YUKSEK=15 ≈ 75ms).

```
  YAPISTIRMA ROLESI (kuru kontak)        ESP32
 ┌───────────────────────┐           ┌────────────────┐
 │  NO ───────────────────┼──kablo───►│ GPIO25         │  INPUT_PULLUP
 │  COM ──────────────────┼──kablo───►│ GND            │
 └───────────────────────┘           └────────────────┘
     (her yapıştırmada KISA GND darbesi = 1 sayım)
```

- Röle 24V veriyorsa yine §3 optolu şeması geçerli (aynı PC817 devresi) — sadece firmware yapıştırma olur.
- **Saha ayarı:** çift/eksik sayarsa `MIN_PULSE_GAP_MS` (1500) ve `INTEG_YUKSEK_ESIK` (15) sahada ayarlanır; heartbeat TANI (`tani_sayildi / tani_parazit / tani_erken`) sunucudan izlenir.

---

## 5. MALZEME LİSTESİ (cihaz başı)

| # | Malzeme | Not |
|---|---|---|
| 1 | ESP32-WROOM-32U DevKit (38-pin) | Ana kart |
| 2 | **LM2596 buck modül** | 24V→5V, pot ile ayarlanır |
| 3 | Blendajlı sinyal kablosu (2–4×0.5mm²) | Röle → ESP32 |
| 4 | DIN-ray kutu / pano | Montaj |
| 5 | *(optolu ise)* PC817 + 2.2kΩ direnç | 24V sinyalde |
| 6 | *(ops.)* 100nF kondansatör | Parazit |

---

## 6. PIN ÖZETİ (her iki firmware ortak)

| ESP32 Pin | İşlev |
|---|---|
| **GPIO25** | Üretim sinyali (röle NO / opto çıkışı) — INPUT_PULLUP, aktif LOW |
| **GPIO26, GPIO27** | Boş (kullanılmıyor) |
| **GPIO2** | Dahili LED (durum) |
| **5V / VIN** | LM2596 OUT+ (5.0V) |
| **GND** | LM2596 OUT− + röle COM (+ optolu ise PC817 emitter) |

---

## 7. KURULUM + TEST

1. `pilot/firmware/cofle_sayac_yapistirma/cofle_sayac_yapistirma.ino` (veya 320t/407t) dosyasını aç.
   *(Firmware generate.py ile üretildi; ayar gerekirse `_templates/yapistirma.ino.tpl`'i düzenle, tekrar üret — tek tek .ino'yu elle düzenleme.)*
2. LM2596'yı **önce 5.0V'a ayarla** (§1), sonra ESP32'yi besle.
3. USB'den ilk flash + Serial Monitor (115200): `[READY]` ve `[NVS] Kayitli sayim: N` görmelisin.
4. **Test:** GPIO25 ile GND arasını bir telle kısa süre kısa devre yap → Serial'da `[PULSE] #1` görmeli, dashboard'da sayım artmalı.
5. Röleyi bağla, birkaç gerçek cevrim izle; TANI sayaçlarını (heartbeat) sunucudan kontrol et.
6. Sonraki flash'ları OTA ile (`OTA_Yukle.bat`) yap.

> **Güvenlik hatırlatma:** Röle kuru kontak DEĞİLSE (24V) optosuz bağlama — §3 optolu şart. Şüphede ölç.
