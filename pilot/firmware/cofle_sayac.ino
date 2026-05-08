/* ============================================================
 *  COFLE PILOT SAYAÇ — ESP32-WROOM-32U firmware (v1.0)
 * ============================================================
 *
 *  HEDEF:  Robot DO çıkışından alınan röle kuru kontağını
 *          sayar, her parça çıkışında Cofle backend'ine
 *          HTTP POST ile pulse atar.
 *
 *  ÖZELLİKLER:
 *   - WiFi auto-reconnect (kopunca otomatik bağlanır)
 *   - Offline buffer (200 adet pulse, RAM'de — ağ koparsa kaybetmez)
 *   - Watchdog timer (30sn'de kilitlenirse kendini reset eder)
 *   - Idempotency-key (cihaz_id + sayı) — çift sayım önler
 *   - Heartbeat (30sn'de bir backend'e "yaşıyorum" sinyali)
 *   - LED gösterge (hareket / hata / iyi durumu)
 *   - Kontak debounce (mekanik röle sıçramasını önler)
 *   - NVS'de toplam pulse sayısı saklanır (yeniden başlatınca kaybolmaz)
 *
 *  KURULUM (Arduino IDE):
 *   1. File > Preferences > Additional Boards URL:
 *        https://espressif.github.io/arduino-esp32/package_esp32_index.json
 *   2. Tools > Board > Boards Manager > "esp32" arat > yükle
 *   3. Tools > Board > "ESP32 Dev Module" seç
 *   4. Tools > Upload Speed > 921600
 *   5. Tools > Flash Size > 4MB
 *   6. Tools > Partition Scheme > "Default 4MB with spiffs"
 *   7. Tools > Port > [ESP32'nin bağlı olduğu COM portu]
 *   8. Aşağıdaki YAPILANDIRMA kısmını doldur, Upload (►)
 *
 *  GEREKLİ KÜTÜPHANELER (otomatik gelir genelde):
 *   - WiFi.h, HTTPClient.h, Preferences.h, esp_task_wdt.h
 *   - ArduinoJson (Sketch > Include Library > Manage Libraries > "ArduinoJson" v6.x)
 *
 *  BAĞLANTI ŞEMASI: PILOT_KURULUM.md dosyasına bakın.
 * ============================================================ */

#include <WiFi.h>
#include <HTTPClient.h>
#include <Preferences.h>
#include <ArduinoJson.h>
#include <esp_task_wdt.h>

// ════════════════════════════════════════════════════════════
//   YAPILANDIRMA — Sahaya göre düzenle
// ════════════════════════════════════════════════════════════

// Bu cihazın benzersiz adı (her ESP32'de farklı olmalı)
const char* CIHAZ_ID  = "ABB1-IST1";

// Bu cihazın izlediği bölüm/robot/istasyon
const char* BOLUM     = "kaynak";    // "kaynak" | "montaj" | "metal"
const char* ROBOT_NO  = "ABB1";       // andon_robot_ayarlari'nda eşleşen ad
const int   ISTASYON  = 1;            // 1 veya 2 (kaynak için), diğerlerinde 0

// WiFi bilgileri (etiket yazıcılarının kullandığı ağ)
const char* WIFI_SSID = "FABRIKA_WIFI";
const char* WIFI_PASS = "wifi_parolasi_buraya";

// Cofle Pilot backend'in URL'si (PC'nin LAN IP'si + port 5001)
const char* SUNUCU_HOST = "http://192.168.1.50:5001";

// Backend'le paylaşılan sır (pilot_app.py'da API_TOKEN ile aynı)
const char* API_TOKEN = "cofle-pilot-2026";

// Pin atamaları (PILOT_KURULUM.md'deki şemaya göre)
const int  SAYAC_PIN  = 25;   // Röle kuru kontağı buraya bağlı (NO)
const int  LED_PIN    =  2;   // Built-in LED (durum göstergesi)

// Diğer ayarlar
const int  DEBOUNCE_MS    = 50;     // Mekanik röle sıçraması filtresi
const int  HEARTBEAT_MS   = 30000;  // 30sn'de bir cihaz heartbeat'i
const int  RETRY_MS       = 3000;   // Buffer'daki bekleyen pulse'ları bu aralıkla deneme
const int  WIFI_TIMEOUT_S = 30;     // WiFi bağlantı timeout
const int  WDT_TIMEOUT_S  = 30;     // Watchdog reset süresi
const int  BUFFER_MAX     = 200;    // RAM'de tutulan max bekleyen pulse
const char* FIRMWARE_VER  = "1.0.0";

// ════════════════════════════════════════════════════════════
//   GLOBAL DURUM
// ════════════════════════════════════════════════════════════

Preferences prefs;
unsigned long pulseSayisi = 0;       // Toplam pulse (NVS'de kalıcı)
int  buffer_kuyruk[BUFFER_MAX];      // Bekleyen sequence numaraları
int  buf_bas = 0, buf_son = 0, buf_dolu = 0;

unsigned long lastPulseRead   = 0;
unsigned long lastHeartbeat   = 0;
unsigned long lastRetry       = 0;
int  lastPinState  = HIGH;
unsigned long lastDebounceMs  = 0;
unsigned long bootMs          = 0;

// ════════════════════════════════════════════════════════════
//   YARDIMCI FONKSİYONLAR
// ════════════════════════════════════════════════════════════

void ledYakBlink(int adet, int sure_ms = 80) {
  for (int i = 0; i < adet; i++) {
    digitalWrite(LED_PIN, HIGH); delay(sure_ms);
    digitalWrite(LED_PIN, LOW);  delay(sure_ms);
  }
}

void wifiBaglan() {
  Serial.printf("[WiFi] '%s' agina baglaniliyor...\n", WIFI_SSID);
  WiFi.mode(WIFI_STA);
  WiFi.begin(WIFI_SSID, WIFI_PASS);
  unsigned long basla = millis();
  while (WiFi.status() != WL_CONNECTED && (millis() - basla) < (WIFI_TIMEOUT_S * 1000)) {
    delay(500); Serial.print('.');
    digitalWrite(LED_PIN, !digitalRead(LED_PIN)); // hızlı blink
    esp_task_wdt_reset();
  }
  if (WiFi.status() == WL_CONNECTED) {
    Serial.printf("\n[WiFi] OK · IP=%s · RSSI=%d dBm\n",
                  WiFi.localIP().toString().c_str(), WiFi.RSSI());
    digitalWrite(LED_PIN, LOW);
    ledYakBlink(2, 100);
  } else {
    Serial.println("\n[WiFi] BAGLANTI BASARISIZ — sonra tekrar denenecek");
  }
}

bool wifiHazir() {
  return WiFi.status() == WL_CONNECTED;
}

// Bekleyen pulse'lardan en eskiyi gönder (varsa). Başarılıysa kuyruktan çıkar.
bool bufferdanBirGonder() {
  if (buf_dolu == 0) return false;
  int seq = buffer_kuyruk[buf_bas];

  HTTPClient http;
  String url = String(SUNUCU_HOST) + "/api/sinyal";
  http.begin(url);
  http.setTimeout(5000);
  http.addHeader("Content-Type", "application/json");
  http.addHeader("Authorization", String("Bearer ") + API_TOKEN);

  StaticJsonDocument<256> doc;
  doc["cihaz_id"] = CIHAZ_ID;
  doc["bolum"]    = BOLUM;
  doc["robot_no"] = ROBOT_NO;
  doc["istasyon"] = ISTASYON;
  doc["kaynak_tip"] = "robot_io";
  String idem = String(CIHAZ_ID) + "_" + String(seq);
  doc["idempotency_key"] = idem;

  String payload;
  serializeJson(doc, payload);

  int rc = http.POST(payload);
  http.end();

  if (rc == 200 || rc == 201) {
    Serial.printf("[POST] OK seq=%d (HTTP %d)\n", seq, rc);
    // Başarılı → kuyruktan çıkar
    buf_bas = (buf_bas + 1) % BUFFER_MAX;
    buf_dolu--;
    return true;
  } else {
    Serial.printf("[POST] HATA seq=%d (HTTP %d) — kuyrukta kal\n", seq, rc);
    return false;
  }
}

void sinyaliKaydet() {
  pulseSayisi++;
  prefs.putULong("pulse", pulseSayisi);  // NVS'de kalıcı

  // Buffer'a ekle (sequence = pulseSayisi)
  if (buf_dolu >= BUFFER_MAX) {
    // Buffer dolu → en eskiyi at, yenisini koy (eski uçtu)
    Serial.println("[BUFFER] DOLU! En eski pulse at, yenisini ekle");
    buf_bas = (buf_bas + 1) % BUFFER_MAX;
    buf_dolu--;
  }
  buffer_kuyruk[buf_son] = (int)pulseSayisi;
  buf_son = (buf_son + 1) % BUFFER_MAX;
  buf_dolu++;

  Serial.printf("[PULSE] #%lu (kuyruk=%d) — gonderiliyor...\n", pulseSayisi, buf_dolu);
  ledYakBlink(2, 50);

  // Eğer WiFi varsa hemen göndermeyi dene
  if (wifiHazir()) {
    bufferdanBirGonder();
  }
}

void heartbeatGonder() {
  if (!wifiHazir()) return;
  HTTPClient http;
  http.begin(String(SUNUCU_HOST) + "/api/sinyal/heartbeat");
  http.setTimeout(5000);
  http.addHeader("Content-Type", "application/json");
  http.addHeader("Authorization", String("Bearer ") + API_TOKEN);

  StaticJsonDocument<384> doc;
  doc["cihaz_id"]      = CIHAZ_ID;
  doc["bolum"]         = BOLUM;
  doc["robot_no"]      = ROBOT_NO;
  doc["firmware_ver"]  = FIRMWARE_VER;
  doc["ip_adresi"]     = WiFi.localIP().toString();
  doc["mac_adresi"]    = WiFi.macAddress();
  doc["wifi_rssi"]     = WiFi.RSSI();
  doc["buffer_kuyruk"] = buf_dolu;
  doc["uptime_sn"]     = (millis() - bootMs) / 1000;
  doc["free_heap"]     = (int)ESP.getFreeHeap();

  String payload; serializeJson(doc, payload);
  int rc = http.POST(payload);
  http.end();
  Serial.printf("[HEART] HTTP %d · RSSI=%d · kuyruk=%d\n", rc, WiFi.RSSI(), buf_dolu);
}

// ════════════════════════════════════════════════════════════
//   SETUP / LOOP
// ════════════════════════════════════════════════════════════

void setup() {
  Serial.begin(115200);
  delay(300);
  bootMs = millis();

  Serial.println("\n╔════════════════════════════════════════════╗");
  Serial.printf( "║  COFLE PILOT SAYAC — %-21s ║\n", CIHAZ_ID);
  Serial.println("╚════════════════════════════════════════════╝");
  Serial.printf("FW: %s · Bolum: %s · Robot: %s · Ist: %d\n",
                FIRMWARE_VER, BOLUM, ROBOT_NO, ISTASYON);
  Serial.printf("MAC: %s\n", WiFi.macAddress().c_str());

  // GPIO
  pinMode(SAYAC_PIN, INPUT_PULLUP);  // Röle kontağı GND'ye kapanır → LOW = pulse
  pinMode(LED_PIN, OUTPUT);
  digitalWrite(LED_PIN, LOW);
  lastPinState = digitalRead(SAYAC_PIN);

  // NVS — kalıcı pulse sayısı
  prefs.begin("cofle", false);
  pulseSayisi = prefs.getULong("pulse", 0);
  Serial.printf("[NVS] Kayitli toplam pulse: %lu\n", pulseSayisi);

  // Watchdog (ESP32 core 3.x ve 2.x ile uyumlu)
#if ESP_ARDUINO_VERSION_MAJOR >= 3
  // Core 3.x: yeni API — config struct
  esp_task_wdt_config_t wdt_config = {
    .timeout_ms     = (uint32_t)(WDT_TIMEOUT_S * 1000),
    .idle_core_mask = (1 << portNUM_PROCESSORS) - 1,
    .trigger_panic  = true,
  };
  esp_err_t wdt_err = esp_task_wdt_init(&wdt_config);
  if (wdt_err == ESP_ERR_INVALID_STATE) {
    // TWDT zaten init edilmiş (Arduino loop'u için) — sadece reconfigure et
    esp_task_wdt_reconfigure(&wdt_config);
  }
#else
  // Core 2.x: eski API — timeout (saniye) + panic flag
  esp_task_wdt_init(WDT_TIMEOUT_S, true);
#endif
  esp_task_wdt_add(NULL);

  wifiBaglan();
  delay(500);

  // İlk heartbeat
  heartbeatGonder();
  lastHeartbeat = millis();
  lastRetry = millis();

  Serial.println("[READY] Sayac aktif — robot DO sinyali bekleniyor...\n");
  ledYakBlink(3, 60);
}

void loop() {
  esp_task_wdt_reset();
  unsigned long now = millis();

  // ── 1. WiFi kontrolu — koptuysa yeniden bagla ──
  if (!wifiHazir()) {
    digitalWrite(LED_PIN, (now / 200) % 2);  // hızlı blink = WiFi yok
    if ((now - lastHeartbeat) > 15000) {     // 15sn'de bir tekrar dene
      wifiBaglan();
      lastHeartbeat = now;
    }
  }

  // ── 2. Sayac pin oku (debounce ile) ──
  int cur = digitalRead(SAYAC_PIN);
  if (cur != lastPinState && (now - lastDebounceMs) > DEBOUNCE_MS) {
    lastDebounceMs = now;
    if (cur == LOW) {  // Röle kapandı → parça çıktı
      sinyaliKaydet();
    }
    lastPinState = cur;
  }

  // ── 3. Buffer'da bekleyen pulse'ları gönder ──
  if (wifiHazir() && buf_dolu > 0 && (now - lastRetry) > RETRY_MS) {
    bufferdanBirGonder();
    lastRetry = now;
  }

  // ── 4. Heartbeat ──
  if (wifiHazir() && (now - lastHeartbeat) > HEARTBEAT_MS) {
    heartbeatGonder();
    lastHeartbeat = now;
  }

  delay(5);  // CPU rahat etsin
}
