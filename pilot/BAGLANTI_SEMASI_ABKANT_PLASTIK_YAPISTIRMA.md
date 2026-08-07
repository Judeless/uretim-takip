# Bağlantı Şeması — Abkant · Pres · Plastik Enjeksiyon · Yapıştırma

ESP32-WROOM-32U sayaç modülünün makine rölesine bağlanması.
Tüm makine tipleri **aynı donanımı** kullanır; fark yalnızca **firmware eşikleri**ndedir.

| Makine | Bölüm | Firmware | Sinyal karakteri | Eşikler |
|---|---|---|---|---|
| **Abkant 1/2/3** | `pres` | `cofle_sayac_abkant_1..3` | ProManage rölesi, her bükümde kısa kapama | INTEG_YÜKSEK=15 (~75 ms) · GAP=1500 ms |
| **Eksantrik Pres 1–5** | `pres` | `cofle_sayac_pres_1..5` | Makinenin **vuruş sayacına** paralel röle | INTEG_YÜKSEK=15 (~75 ms) · GAP=600 ms |
| **Hidrolik Pres** | `pres` | `cofle_sayac_hidrolik_pres` | Makinenin **vuruş sayacına** paralel röle | INTEG_YÜKSEK=15 (~75 ms) · GAP=600 ms |
| **Broş (`Bros`)** | `pres` | `cofle_sayac_bros` | Makinenin **vuruş sayacına** paralel röle | INTEG_YÜKSEK=15 (~75 ms) · GAP=600 ms |
| **Plastik enj. 320T / 407T** | `plastik` | `cofle_sayac_320t` / `_407t` | Uzun çevrim rölesi (saniyeler LOW) | INTEG_YÜKSEK=800 (~4 s) · GAP=18000 ms |
| **Yapıştırma** | `plastik` | `cofle_sayac_yapistirma` | Kısa pulse (~50–200 ms) | INTEG_YÜKSEK=15 (~75 ms) · GAP=1500 ms |

> Firmware `_templates/*.ino.tpl` + `generate.py` ile üretilir — tek tek `.ino` düzenleme.
> Abkant+pres `pres`, plastik+yapıştırma `plastik` bölümüne sinyal basar; makine `robot_no` ile ayrılır.

**Pres GAP'i neden 600 ms?** Abkant'ta iki büküm arası uzundur (1500 ms güvenli), ama eksantrik
pres sürekli modda 60+ vuruş/dk yapabilir; 1500 ms filtre vuruş **kaybettirir**. 600 ms en fazla
100 vuruş/dk geçirir. Kontak sıçramasını zaten 75 ms'lik integratör eler. Sahada ayarlanır (§6).

**Röleyi nereye takıyoruz?** Makinenin kendi mekanik/elektronik vuruş sayacına **paralel** bir röle
(sayaç bobini enerjilenince röle de çeker). Böylece "makine bir vuruş yaptı" bilgisi, PLC'ye hiç
dokunmadan kuru kontak olarak alınır. Sayaç 24 V ile sürülüyorsa röle bobini 24 V seçilir; ESP32
tarafına yalnız rölenin **NO–COM kuru kontağı** gider.

---

## ⚠️ 0. ÖNCE ÖLÇ — bağlamadan önce (en kritik kural)

Röle çıkışını multimetreyle ölç:

- **Kuru kontak** (potansiyelsiz; aktifken sadece süreklilik / 0 V) → **OPTOSUZ** şema (§2).
- **Gerilim çıkışı** (aktifken 12 V / 24 V okuyor) → **OPTOLU** şema (§3). Direkt bağlarsan **ESP32 yanar.**

Kaynak/metal makinelerimizde röle kuru kontak; abkanttaki ProManage rölesinde ve plastikte de büyük ihtimalle öyle — **yine de her makinede teyit et.**

---

## 1. GÜÇ — LM2596 buck (üç makinede de AYNI)

Makinenin 24 V'unu 5 V'a indirip ESP32'yi besler. Harici USB adaptör kullanma.

```
   MAKINE 24V DC                 LM2596 MODUL                 ESP32-WROOM-32U
 ┌──────────────┐            ┌───────────────────┐         ┌────────────────┐
 │  24V (+) ─────┼───────────►│ IN+          OUT+ ┼────────►│ 5V  (VIN)      │
 │  0V  (−) ─────┼───────────►│ IN−          OUT− ┼────────►│ GND            │
 └──────────────┘            │   [pot: 5.0V'a]   │         └────────────────┘
                             └───────────────────┘
```

**Kurulum sırası (şart):**
1. LM2596 girişine 24 V ver, **çıkışa hiçbir şey bağlama.**
2. OUT+ / OUT− arasını ölç, potu çevirip **5.0 V** yap (5.0–5.2 V kabul; **>5.5 V ESP32'yi bozar**).
3. Gücü kes → OUT+ → ESP32 **`5V`/`VIN`**, OUT− → **`GND`**.
4. ❌ 5 V'u **asla `3V3` pinine verme.** Dahili regülatör 5 V→3.3 V yapar.

> Giriş aralığı 4.5–40 V (24 V uygun). ESP32 ~250 mA çeker, modül 2–3 A verir.

---

## 2. SİNYAL — OPTOSUZ (röle KURU KONTAK) · tercih edilen

Üç makinede de **elektriksel bağlantı birebir aynı.** Röle zaten izole olduğu için araya bir şey girmez.

```
  MAKINE ROLESI (kuru kontak)            ESP32-WROOM-32U
 ┌───────────────────────┐           ┌────────────────┐
 │  NO (norm. acik) ──────┼──kablo───►│ GPIO25         │  INPUT_PULLUP (bosta HIGH=3.3V)
 │  COM (ortak)     ──────┼──kablo───►│ GND            │
 └───────────────────────┘           └────────────────┘
        (röle her cevrimde NO–COM'u kapatir -> GPIO25 LOW -> 1 sayim)
```

Makine bazında **sadece kapanma süresi** değişir — firmware bunu ayırır:

```
  ABKANT        ──┐ ┌──────────   her bukumde ~100-300ms LOW   (INTEG 15 / GAP 1500ms)
                  └─┘

  PRES (sayac)  ──┐┌──┐┌──┐┌───   her vuruste kisa pulse       (INTEG 15 / GAP 600ms)
                  └┘  └┘  └┘      (surekli modda 60+ vurus/dk)

  YAPISTIRMA    ──┐┌───────────   ~50-200ms kisa pulse         (INTEG 15 / GAP 1500ms)
                  └┘

  PLASTIK ENJ.  ──┐        ┌───   saniyeler suren LOW          (INTEG 800 / GAP 18s)
                  └────────┘
```

- NO–COM hattında **blendajlı (shielded) kablo** kullan; blendajı **pano tarafında tek uçtan** toprakla (motor/ısıtıcı paraziti).
- Sinyal kablosunu güç/motor kablolarıyla **aynı kanaldan geçirme.**
- İsteğe bağlı: GPIO25–GND arası **100 nF** kondansatör (firmware integratörü zaten filtreliyor).

---

## 3. SİNYAL — OPTOLU (röle 24 V GERİLİM veriyorsa)

Çıkış kuru kontak değil de PLC/transistör çıkışı gibi **24 V darbe** veriyorsa **PC817 optokuplör** şart.

```
   24V SINYAL           R = 2.2kΩ          PC817 (DIP-4)              ESP32
 ┌────────────┐        (1/4W)          ┌──────────────────┐      ┌───────────┐
 │ Sinyal(+) ──┼───────[ R ]──────────►│1 anot    kolek. 4├─────►│ GPIO25    │  INPUT_PULLUP
 │             │                       │                  │      │           │
 │ Sinyal 0V ──┼──────────────────────►│2 katot   emit.  3├─────►│ GND       │
 └────────────┘                        └──────────────────┘      └───────────┘
```

- 24 V gelince LED yanar → foto-transistör iletir → GPIO25 **LOW** → sayar (aktif-LOW, firmware ile uyumlu).
- **Direnç:** 24 V → **2.2 kΩ**; 12 V → **1 kΩ** (LED akımı ~10 mA).
- **PC817 pin sırası** (nokta/çentik 1. pin): `1=anot · 2=katot · 3=emitter · 4=kollektör`.
- **Polarite:** "üretimde 24 V, boşta 0 V" varsayıldı. Makine tersse (boşta 24 V) LED yönünü ölçüp doğrula.
- ESP32'yi aynı 24 V'dan besliyorsan opto galvanik izolasyon değil, **24 V→3.3 V güvenli seviye çevirici** işlevi görür. Tam izolasyon isteniyorsa ESP32'ye **ayrı 5 V** besleme ver.

**Ek malzeme (optolu):** 1× PC817 + 1× 2.2 kΩ.

---

## 4. PİN ÖZETİ (üç firmware ortak)

| ESP32 Pin | İşlev |
|---|---|
| **GPIO25** | Üretim sinyali — röle NO veya PC817 kollektörü · `INPUT_PULLUP`, aktif LOW |
| **GND** | Röle COM + LM2596 OUT− (+ optoluysa PC817 emitter) |
| **5V / VIN** | LM2596 OUT+ (5.0 V) |
| **GPIO26** | Boş (`INPUT_PULLUP`) |
| **GPIO27** | Abkant/pres firmware'inde buzzer çıkışı var ama **buzzer TAKILMAZ**; plastik/yapıştırmada tanımsız |
| **GPIO2** | Dahili LED (durum göstergesi) |

Sunucu tarafı: abkant+pres `kaynak_tip="role"`, plastik/yapıştırma `kaynak_tip="makine_io"` gönderir; POST `/api/sinyal`, heartbeat 30 sn.

---

## 5. MALZEME LİSTESİ (cihaz başı)

| # | Malzeme | Not |
|---|---|---|
| 1 | ESP32-WROOM-32U DevKit (38-pin) | Ana kart |
| 2 | LM2596 buck modül | 24 V→5 V, pot ile ayarlanır |
| 3 | Blendajlı sinyal kablosu (2×0.5 mm² yeter) | Röle → ESP32 |
| 4 | DIN-ray kutu / pano | Montaj |
| 5 | *(optolu ise)* PC817 + 2.2 kΩ direnç | 24 V sinyalde |
| 6 | *(ops.)* 100 nF kondansatör | Parazit |

---

## 6. KURULUM + TEST

1. İlgili `.ino`'yu aç: `pilot/firmware/cofle_sayac_abkant_1` · `..._pres_1..5` ·
   `..._hidrolik_pres` · `..._bros` · `..._320t` / `..._407t` · `..._yapistirma`.
2. LM2596'yı **önce 5.0 V'a ayarla** (§1), sonra ESP32'yi besle.
3. İlk flash USB'den + Serial Monitor (115200): `[READY]` ve `[NVS] Kayitli sayim: N` görmelisin.
4. **Kuru test:** GPIO25 ↔ GND arasını bir telle kısa devre yap → Serial'da `[PULSE] #1`, dashboard'da sayım artmalı.
5. Röleyi bağla, birkaç gerçek çevrim izle. Heartbeat TANI sayaçlarını sunucudan kontrol et:
   `tani_sayildi` / `tani_parazit` (ham gürültü kenarı) / `tani_erken` (GAP'e takılan).
6. **Saha ayarı:**
   - Çift sayıyorsa → `MIN_PULSE_GAP_MS` yükselt.
   - Eksik sayıyorsa → `INTEG_YUKSEK_ESIK` düşür (veya plastikte çevrim süresini ölçüp GAP'i düşür).
   - Değişiklik `_templates/*.ino.tpl`'de yapılır, `python generate.py` ile yeniden üretilir.
7. Sonraki flash'lar OTA ile: `pilot/firmware/OTA_Yukle.bat`.

> **Güvenlik:** Röle kuru kontak DEĞİLSE optosuz bağlama — §3 şart. Şüphede ölç.
