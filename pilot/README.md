# Cofle Pilot Sayaç

ESP32 ile robot DO çıkışından otomatik üretim sayımı. **Ana Cofle Manage sistemine
hiç dokunmadan** kendi başına çalışan bir pilot uygulaması (port 5001, ayrı SQLite).

## Klasör yapısı

```
pilot/
├── pilot_app.py          ← Flask backend (port 5001)
├── pilot.db              ← Otomatik oluşur — sayaç + cihaz kayıtları
├── templates/
│   └── pilot.html        ← Canlı sayaç UI'ı
├── firmware/
│   └── cofle_sayac.ino   ← ESP32 Arduino sketch
├── PILOT_KURULUM.md      ← Donanım bağlantı şeması, opto/röle, RAPID kodu
├── Pilot_Baslat.bat      ← Windows başlatıcı (port 5001'i çalıştırır)
└── README.md             ← Bu dosya
```

## Nasıl çalıştırılır

### 1. Backend (PC tarafı)

```bash
cd pilot
python pilot_app.py
```

Veya Windows'ta:
```
pilot\Pilot_Baslat.bat
```

Backend port **5001**'de açılır. Konsolda göreceksin:
```
  Port      : 5001
  Canlı UI  : http://<bu-pc-ip>:5001/
  ESP32 URL : http://<bu-pc-ip>:5001/api/sinyal
```

Tarayıcıdan `http://localhost:5001/` aç → Cofle Pilot Canlı Sayaç sayfası

### 2. ESP32 firmware (cihaz tarafı)

`pilot/firmware/cofle_sayac.ino` — Arduino IDE ile aç, en üstteki **YAPILANDIRMA**
bölümünde şu üç şeyi düzenle:

```cpp
const char* CIHAZ_ID    = "ABB1-IST1";              // benzersiz cihaz adı
const char* WIFI_SSID   = "FABRIKA_WIFI";           // fabrika ağı
const char* WIFI_PASS   = "wifi_parolasi_buraya";
const char* SUNUCU_HOST = "http://192.168.1.50:5001"; // PC'nin LAN IP'si
```

Upload → Serial Monitor (115200) → "Sayaç aktif" mesajını gör → kabloyu sahaya çek.

### 3. Donanım bağlantısı

[PILOT_KURULUM.md](PILOT_KURULUM.md) dosyasına bak — şema, RAPID kodu, sorun giderme.

## Test akışı (donanım gelmeden)

ESP32 olmadan da pilot UI'ı test edebilirsin: curl ile manuel pulse at.

```bash
# 1. Backend'i başlat (pilot_app.py)
# 2. Test pulse:
curl -X POST http://localhost:5001/api/sinyal ^
  -H "Authorization: Bearer cofle-pilot-2026" ^
  -H "Content-Type: application/json" ^
  -d "{\"cihaz_id\":\"TEST\",\"bolum\":\"kaynak\",\"robot_no\":\"ABB1\",\"istasyon\":1,\"idempotency_key\":\"TEST_1\"}"

# 3. Tarayıcıda http://localhost:5001/ → sayaç +1 olmalı
# 4. Aynı idempotency_key tekrar gönder → backend reddeder (çift sayım yok)
```

## Pilot başarılı olduğunda

`pilot/` klasöründeki kod ana sisteme taşınır:
- `pilot.db.sayac_olaylari` ve `cihaz_kayitlari` tabloları → `database.py`'a migration
- API endpoint'leri → `app.py`'a `/api/sinyal/*` olarak entegre
- Pilot UI'daki **canlı sayaç widget'ı** → andon ekranına ek
- Vardiya bağlama: pulse alındığında o anki açık vardiyaya otomatik kayıt
- Gün sonu mutabakat: operatör OK adet vs. sayaç toplamı karşılaştırma

Bu taşıma ~2-3 saatlik bir iştir; o aşamaya gelince yapılır.

## Güvenlik

- ESP32 → backend HTTP POST'larında `Authorization: Bearer cofle-pilot-2026` header'ı
- LAN-only kullanım için yeterli (fabrika içi ağ, dışa açık değil)
- Token'ı `pilot_app.py` içinde `API_TOKEN` değişkeninden değiştirebilirsin (firmware'de aynı değer)

## İletişim portu

- Pilot: **5001** (TCP, HTTP)
- Ana sistem: **5000** (TCP, HTTP) — pilot çalışırken aynı anda çalışmaya devam eder
- ESP32 cihazlarının kullanacağı port sadece 5001
