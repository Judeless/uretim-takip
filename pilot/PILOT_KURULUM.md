# Cofle Pilot Sayaç — Donanım Kurulum Rehberi

ESP32-WROOM-32U + 24V Röle ile robot DO çıkışından sayaç pulse alma.
Pilot olarak **ABB1 robotu / İstasyon 1** için.

---

## 1. Malzeme listesi (1 nokta için)

| # | Parça | Açıklama |
|---|-------|----------|
| 1 | **Olimex ESP32-WROOM-32U DevKit** (1 saha + 1 yedek) | U.FL anten konnektörlü |
| 2 | **2.4 GHz harici WiFi anten** + **U.FL → RP-SMA pigtail** | Kabin dışına anten |
| 3 | **Endüstriyel röle 24V DC bobinli** (kuru kontak NO çıkışlı) | Bakımdan; Omron G2R-1-SN-24V veya muadili |
| 4 | **Röle soketi DIN-ray** (varsa pluggable röle için) | Bakımdan |
| 5 | **DIN-ray plastik kabin** ~100×80×40 mm (Altınkaya RT-504 vb.) | IP54 |
| 6 | **5V/2A USB adaptör** + **micro-USB kablo** (1-1.5m kaliteli) | ESP32 beslemesi |
| 7 | **4×0.5mm² shielded sinyal kablosu** | Robot DO → Röle bobini |
| 8 | **Vidalı klemens** (6 yollu) | Kabin içi bağlantı |
| 9 | **PG7 kablo rakor**ı x 2 | Kabin yan duvar girişleri |
| 10 | **220V priz** (kabin yakını) | USB adaptör için |

---

## 2. Bağlantı şeması (mantıksal akış)

```
┌────────────────────┐                                ┌─────────────────────────┐
│   ABB ROBOT        │                                │   FABRİKA WIFI AĞI      │
│   (RAPID kodda)    │                                │   (etiket yazıcılarla   │
│                    │                                │    aynı SSID)           │
│   SetDO ist1, 1   ─┼──┐                             └──────────▲──────────────┘
│   WaitTime 0.2     │  │ 24V DC                                 │ WiFi 2.4 GHz
│   SetDO ist1, 0    │  │ (parça çıkışında)                       │
└────────────────────┘  │                                         │
                        │                                         │
                        ▼                                         │
              ┌─────────────────────┐                             │
              │   24V RÖLE          │                             │
              │   (bobin DC 24V)    │                             │
              │                     │                             │
              │   Bobin (+/−) ◄─────┼──── 24V DO + 0V             │
              │                     │                             │
              │   Kuru kontak (NO)  │ → ── ── ── ── ── ── ── ──┐ │
              └─────────────────────┘                          │ │
                                                               │ │
                          ┌────────────────────────────────────┼─┼──────┐
                          │  COFLE PİLOT KABİN                 │ │      │
                          │  (DIN-ray plastik IP54)            │ │      │
                          │                                    │ │      │
                          │  ┌──────────────────────┐          │ │      │
                          │  │ ESP32-WROOM-32U      │          │ │      │
                          │  │                      │◄─────────┘ │      │
                          │  │  GPIO25 (SAYAC_PIN) ◄┼── Röle NO  │      │
                          │  │  GND                ◄┼── Röle COM │      │
                          │  │                      │            │      │
                          │  │  micro-USB ◄─────────┼─ 5V/2A     │      │
                          │  │                      │   adaptör  │      │
                          │  │  U.FL ────────────┐  │            │      │
                          │  └───────────────────┼──┘            │      │
                          │                       │              │      │
                          │  ┌─ pigtail ──────────┘              │      │
                          │  │                                   │      │
                          └──┼───────────────────────────────────┘      │
                             │                                          │
                             └─►  Harici WiFi anten (kabin DIŞINA)      │
                                                                        │
                                                                        ▼
                                                       ┌────────────────────────┐
                                                       │  COFLE MANAGE PC       │
                                                       │  (LAN'da 192.168.x.x)  │
                                                       │  pilot_app.py :5001    │
                                                       │  /api/sinyal           │
                                                       └────────────────────────┘
```

---

## 3. Pin bağlantıları (ESP32 → Röle)

ESP32-WROOM-32U DevKit pin numaraları (38-pin board, micro-USB üstte):

| ESP32 Pin | Bağlanan | Açıklama |
|-----------|----------|----------|
| **GPIO25** (sağ taraf, üstten 11. pin) | Röle kontağı **NO** ucu | INPUT_PULLUP — kontak kapanınca LOW okur |
| **GND** (her iki yanda var) | Röle kontağı **COM** ucu | Ortak toprak |
| **micro-USB** | 5V adaptör | Güç |
| **U.FL konnektör** | Pigtail → harici anten | WiFi |

> **GPIO25 neden?** Boot sırasında pull-up/down olmayan, INPUT_PULLUP destekleyen "güvenli" pin. ADC2 değil. Eğer GPIO25 kullanılmışsa GPIO26 da aynı şekilde uygundur.

---

## 4. Robot tarafı (ABB RAPID kodu)

Her parça çıkışında 200ms HIGH pulse atan bir DO ekle. Robot programcısıyla şu satırları parça çıkış noktasına yerleştirin:

```rapid
! Parça istasyon 1'de bittiğinde
SetDO  do_sayac_ist1, 1;
WaitTime 0.2;
SetDO  do_sayac_ist1, 0;
```

`do_sayac_ist1` adında bir DO yoksa robot terminal bloğunda yeni bir DO çıkışı tahsis et.

**Voltaj kontrolü:** Multimetre ile DO bacağı ile 0V arasını ölç — robot çevrim yaptığında 24V görmeli, beklenirken 0V.

---

## 5. Röle bağlantısı detayı

### Bobin tarafı (robot çıkışı → röle giriş)

```
Robot panosu                    Röle (24V bobinli, örn. Omron G2R-1)
┌─────────────┐                 ┌────────────────┐
│  do_sayac_  │                 │                │
│  ist1 (24V)─┼─── kablo ──────►│ A1 (bobin +)   │
│             │                 │                │
│  0V (GND)  ─┼─── kablo ──────►│ A2 (bobin −)   │
└─────────────┘                 │                │
```

> Eğer rölede **dahili flyback diyot YOKSA** (bobin uçlarına paralel 1N4007 ekle, anot A2'ye katot A1'e — bobinin geri-EMK'sını söndürür ve robot DO'sunu korur).

### Kontak tarafı (röle çıkışı → ESP32)

```
Röle kontak tarafı              ESP32
┌────────────────┐              ┌─────────────┐
│   COM (com)   ─┼─── kablo ───►│ GND         │
│                │              │             │
│   NO (n.open) ─┼─── kablo ───►│ GPIO25      │
└────────────────┘              └─────────────┘

NC (normally closed) ucunu BAĞLAMA — kullanılmıyor.
```

ESP32 INPUT_PULLUP modunda GPIO25'i HIGH tutar. Röle kapanınca kontak GND'ye düşer → GPIO25 LOW okur → firmware bunu pulse olarak sayar.

---

## 6. Sahaya hazırlık adımları

### A) Atölye prototip testi (sahaya gitmeden önce)

1. ESP32'yi USB ile PC'ye tak, Arduino IDE'yi aç
2. `pilot/firmware/cofle_sayac.ino` dosyasını aç
3. **YAPILANDIRMA** kısmındaki değerleri doldur:
   - `CIHAZ_ID = "ABB1-IST1"`
   - `WIFI_SSID`, `WIFI_PASS` (etiket yazıcılarının ağı)
   - `SUNUCU_HOST = "http://<PC-LAN-IP>:5001"`
4. **Tools → Board → ESP32 Dev Module** seç, doğru COM portunu seç
5. **Sketch → Include Library → Manage Libraries** → "ArduinoJson" v6.x kur
6. Upload (►) bas
7. **Serial Monitor** (Ctrl+Shift+M, 115200 baud) — bağlantı log'larını izle:
   - `[WiFi] OK · IP=192.168.1.X · RSSI=-55 dBm`
   - `[NVS] Kayitli toplam pulse: 0`
   - `[READY] Sayac aktif...`
8. Test butonu (kapı zili tipi) ile GPIO25 ile GND arasını kısa devre yap → Serial'da `[PULSE] #1` görmelisin
9. Pilot UI'da (`http://<PC-IP>:5001/`) sayaç +1 olmalı

### B) Saha montajı

1. **DIN-ray kabini** robot kontrol panosunun yanına monte et (≥1m mesafe — kaynak arkı EMI azaltma)
2. Cat6 kablosu **kullanma**, sadece 220V uzatma + sinyal kablosu yeter (WiFi'la konuşacak)
3. Robot panosundan röle kabinine **24V + 0V kabloyu** çek (4×0.5mm² shielded)
4. Röleyi DIN-ray'e tak, bobini robot 24V'una bağla
5. Röle kontağından ESP32 kabinine kısa kablo (NO + COM)
6. ESP32 kabinine **anteni dışına çıkar** (PG7 rakorla yandan, pigtail uzunluğu yetiyorsa)
7. **220V adaptör** kabinin yakındaki prize takılı dursun (uzatma kullanma)
8. ESP32'yi USB ile besle, çalıştığını LED'den teyit et:
   - **Hızlı blink**: WiFi bağlanıyor
   - **2 kısa blink**: WiFi OK, hazır
   - **Söndü**: çalışıyor (idle)
   - **Çift kısa blink (her pulse'ta)**: pulse algılandı

### C) İlk doğrulama

1. Pilot PC'sinde `Pilot_Baslat.bat`'ı çalıştır → backend port 5001'de açılır
2. Tarayıcıdan `http://<PC-IP>:5001/` aç → Cofle Pilot Canlı Sayaç sayfası
3. **Bağlı Cihazlar** tablosunda `ABB1-IST1` görünmeli (durum: **● ONLINE**, son heartbeat: birkaç saniye önce)
4. Robot çevrim yapsın → her parçada sayaç +1 olmalı, **Son Sinyaller** log'unda satırlar görünmeli
5. WiFi kablosunu (router'dan) çek → ESP32 buffer'a almaya devam eder. Tekrar tak → birikmiş pulse'lar otomatik gönderilir

---

## 7. Sorun giderme

| Belirti | Olası neden | Çözüm |
|---------|-------------|-------|
| Cihaz **OFFLINE** görünüyor | WiFi yok / yanlış SSID-parola | Serial Monitor'da log oku, kimlikleri doğrula |
| Sayaç **artmıyor** | Röle kontağı kapanmıyor | Röle bobinine 24V geliyor mu multimetre ile bak; kontak NO ucu doğru mu |
| Sayaç **2 katı sayıyor** | Debounce yetersiz / röle çok titrek | `DEBOUNCE_MS`'i 50→100 yap firmware'de |
| Sayaç **bazen kaçırıyor** | Pulse genişliği <50ms (DEBOUNCE'tan kısa) | Robot RAPID'de `WaitTime` 0.2'den 0.3-0.5'e çıkar |
| ESP32 **sürekli reset** | Güç yetersiz / kötü USB kablo | Kaliteli adaptör + kısa USB; LED brownout işareti olur |
| WiFi **kopuk-bağlı** | Sinyal zayıf, kaynak arkı | Anteni kabin dışına çıkar; sinyal -75 dBm altıysa AP yakına ekle |
| Çift sayım **(idempotency)** | Backend duplicate kabul ediyor | UNIQUE constraint OK çalışıyor — Serial'da `[POST] HATA seq=N (HTTP 409)` görmek normal, kayıp değil |

---

## 8. Genel ağ planı

```
┌──────────────────┐                       ┌──────────────────┐
│  Cofle PC        │                       │  Fabrika WiFi    │
│  (sunucu)        │  192.168.1.50         │  Router/AP       │
│  pilot:5001      │◄──────────────────────│                  │
│  ana:5000        │  LAN                  │  SSID: FABRIKA   │
└──────────────────┘                       └────────┬─────────┘
                                                    │ 2.4 GHz
                                          ┌─────────┼─────────┐
                                          ▼         ▼         ▼
                                    ┌────────┐ ┌────────┐ ┌────────┐
                                    │ ESP32  │ │ Etiket │ │ Mobil  │
                                    │ ABB1   │ │ yazıcı │ │ telefon│
                                    │ ISLAR  │ │        │ │        │
                                    └────────┘ └────────┘ └────────┘
```

ESP32 → PC IP+port 5001'e doğrudan HTTP POST. Aynı LAN olduğu için firewall genelde sorun çıkarmaz (eğer Windows Defender uyarı verirse "Allow access" → "Private network" seç).

---

## 9. Sonraki adımlar

Pilot 1 robotta 1 hafta stabil çalıştığında:
- Diğer ABB robotları için aynı setupu çoğalt (her biri için ayrı `CIHAZ_ID`)
- Backend'i ana sisteme entegre et (`pilot_app.py` → `app.py` içine taşı, vardiya bağlama eklendi)
- Andon ekranlarına **canlı sayaç widget'ı** eklendi (mevcut robot kartının üstünde)
- Gün sonu mutabakat: operatör girdiği OK adet vs. sayaçtan gelen toplam karşılaştırması
