/* ============================================================
 *  COFLE PILOT SAYAC — ESP32-WROOM-32U firmware (v2.0)
 * ============================================================
 *
 *  HEDEF:  Robot'un 3 dijital çıkışını izler:
 *            - İstasyon 1 sayaç (her parça = 1 pulse)
 *            - İstasyon 2 sayaç (her parça = 1 pulse)
 *            - Robot çalışıyor durumu (HIGH=çalışıyor)
 *          Her pulse Cofle pilot backend'ine POST edilir.
 *
 *  ÖZELLİKLER:
 *   - 3 ayrı GPIO girişi, debounce'lu
 *   - WiFi auto-reconnect
 *   - Offline buffer (her pulse istasyon bilgisiyle saklanır)
 *   - Watchdog (30sn'de kilitlenirse kendini reset)
 *   - Idempotency-key — çift sayım önler
 *   - Heartbeat (30sn'de bir cihaz + robot durumu)
 *   - LED gösterge
 *   - NVS'de istasyon başına ayrı pulse sayacı (yeniden başlatınca korunur)
 *
 *  PIN ATAMALARI (sahaya göre, üst sıra):
 *   - GPIO25  → Röle 1 NO  (İstasyon 1 sayaç)
 *   - GPIO26  → Röle 2 NO  (İstasyon 2 sayaç)
 *   - GPIO27  → Röle 3 NO  (Robot çalışıyor durumu)
 *   - GND     → her 3 rölenin COM uçları paralel
 *
 *  BAĞLANTI ŞEMASI: PILOT_KURULUM.md
 * ============================================================ */

#include <WiFi.h>
#include <HTTPClient.h>
#include <Preferences.h>
#include <ArduinoJson.h>
#include <esp_task_wdt.h>
#include <ArduinoOTA.h>

// ════════════════════════════════════════════════════════════
//   YAPILANDIRMA — Sahaya göre düzenle
// ════════════════════════════════════════════════════════════

// Bu cihazın benzersiz adı (her ESP32'de farklı olmalı)
const char* CIHAZ_ID  = "ABB1-IO";

// Bu cihazın izlediği bölüm/robot
const char* BOLUM     = "kaynak";    // "kaynak" | "montaj" | "metal"
const char* ROBOT_NO  = "ABB1";      // andon_robot_ayarlari'nda eşleşen ad
// (İstasyon: her sinyal kendi pin'ine göre 1 veya 2 olarak otomatik etiketlenir)

// WiFi bilgileri
const char* WIFI_SSID = "FABRIKA_WIFI";
const char* WIFI_PASS = "wifi_parolasi_buraya";

// Cofle Pilot backend'in URL'si
const char* SUNUCU_HOST = "http://192.168.1.50:5001";

// Backend'le paylaşılan sır (pilot_app.py'da API_TOKEN ile aynı)
const char* API_TOKEN = "cofle-pilot-2026";

// OTA güncelleme parolası (Arduino IDE upload sırasında sorulur)
const char* OTA_PASS = "cofle-ota-2026";

// Pin atamaları
const int PIN_IST1_SAYAC      = 25;   // Röle 1 NO — İstasyon 1
const int PIN_IST2_SAYAC      = 26;   // Röle 2 NO — İstasyon 2
const int PIN_ROBOT_CALISIYOR = 27;   // Röle 3 NO — Robot durumu
const int PIN_LED             =  2;   // Built-in LED

// Diğer ayarlar
const int  DEBOUNCE_MS         = 100;     // Mekanik röle sıçraması
const int  PARAZIT_SAMPLE_N    = 15;      // LOW gördükten sonra kaç kez teyit
const int  PARAZIT_SAMPLE_GAP  = 5;       // Teyit okumaları arası ms (toplam 75ms — 200ms pulse'un 1/3'ü)
const unsigned long MIN_PULSE_GAP_MS = 3000;  // İki pulse arası minimum (3sn — robot cycle bundan hızlı olamaz)
const int  HEARTBEAT_MS   = 30000;
const int  RETRY_MS       = 3000;
const int  WIFI_TIMEOUT_S = 30;
const int  WDT_TIMEOUT_S  = 30;
const int  BUFFER_MAX     = 200;
const char* FIRMWARE_VER  = "2.1.0";

// ════════════════════════════════════════════════════════════
//   GLOBAL DURUM
// ════════════════════════════════════════════════════════════

Preferences prefs;

// İstasyon başına ayrı sayaç (NVS'de kalıcı)
uint32_t pulseIst1 = 0;
uint32_t pulseIst2 = 0;

// Robot durumu (anlık)
bool robotCalisiyor = false;

// Buffer — her giriş seq + istasyon tutuyor
struct PulseKaydi {
  uint32_t seq;
  uint8_t  istasyon;   // 1 veya 2
};
PulseKaydi buffer[BUFFER_MAX];
int buf_bas = 0, buf_son = 0, buf_dolu = 0;

// Pin durumları (debounce için)
int lastIst1State = HIGH;
int lastIst2State = HIGH;
int lastRobotState = HIGH;
unsigned long lastDebounceIst1 = 0;
unsigned long lastDebounceIst2 = 0;
unsigned long lastDebounceRobot = 0;

// Parazit filtresi — son geçerli pulse zamanı (her istasyon için ayrı)
unsigned long lastValidPulseIst1 = 0;
unsigned long lastValidPulseIst2 = 0;

// Robot durumu periyodik stabilite kontrolü
unsigned long lastRobotCheck = 0;
const unsigned long ROBOT_CHECK_INTERVAL_MS = 500;  // Her 500ms'de bir
const int           ROBOT_SAMPLE_N          = 10;   // 10 örnek
const int           ROBOT_SAMPLE_THRESHOLD  = 7;    // 7+ örnek aynıysa kabul (yaklaşık %70)
unsigned long lastRobotStateChangeMs = 0;

unsigned long lastHeartbeat   = 0;
unsigned long lastRetry       = 0;
unsigned long bootMs          = 0;

// ════════════════════════════════════════════════════════════
//   YARDIMCI FONKSİYONLAR
// ════════════════════════════════════════════════════════════

void ledYakBlink(int adet, int sure_ms = 80) {
  for (int i = 0; i < adet; i++) {
    digitalWrite(PIN_LED, HIGH); delay(sure_ms);
    digitalWrite(PIN_LED, LOW);  delay(sure_ms);
  }
}

void wifiBaglan() {
  Serial.printf("[WiFi] '%s' agina baglaniliyor...\n", WIFI_SSID);
  WiFi.mode(WIFI_STA);
  WiFi.begin(WIFI_SSID, WIFI_PASS);
  unsigned long basla = millis();
  while (WiFi.status() != WL_CONNECTED && (millis() - basla) < (WIFI_TIMEOUT_S * 1000)) {
    delay(500); Serial.print('.');
    digitalWrite(PIN_LED, !digitalRead(PIN_LED));
    esp_task_wdt_reset();
  }
  if (WiFi.status() == WL_CONNECTED) {
    Serial.printf("\n[WiFi] OK · IP=%s · RSSI=%d dBm\n",
                  WiFi.localIP().toString().c_str(), WiFi.RSSI());
    digitalWrite(PIN_LED, LOW);
    ledYakBlink(2, 100);
  } else {
    Serial.println("\n[WiFi] BAGLANTI BASARISIZ — sonra tekrar denenecek");
  }
}

bool wifiHazir() {
  return WiFi.status() == WL_CONNECTED;
}

// Pin gerçekten LOW mu? Parazit elemek için N kez peş peşe oku, hepsi LOW olmalı
bool pinGercektenLOW(int pin) {
  for (int i = 0; i < PARAZIT_SAMPLE_N; i++) {
    if (digitalRead(pin) != LOW) return false;
    delay(PARAZIT_SAMPLE_GAP);
  }
  return true;
}

// Bekleyen pulse'lardan en eskiyi gönder
bool bufferdanBirGonder() {
  if (buf_dolu == 0) return false;
  PulseKaydi p = buffer[buf_bas];

  HTTPClient http;
  String url = String(SUNUCU_HOST) + "/api/sinyal";
  http.begin(url);
  http.setTimeout(5000);
  http.addHeader("Content-Type", "application/json");
  http.addHeader("Authorization", String("Bearer ") + API_TOKEN);

  StaticJsonDocument<256> doc;
  doc["cihaz_id"]   = CIHAZ_ID;
  doc["bolum"]      = BOLUM;
  doc["robot_no"]   = ROBOT_NO;
  doc["istasyon"]   = p.istasyon;
  doc["kaynak_tip"] = "robot_io";
  // idempotency_key: cihaz_i<istasyon>_<seq> — istasyon başına ayrı seri
  String idem = String(CIHAZ_ID) + "_i" + String(p.istasyon) + "_" + String(p.seq);
  doc["idempotency_key"] = idem;

  String payload;
  serializeJson(doc, payload);

  int rc = http.POST(payload);
  http.end();

  if (rc == 200 || rc == 201) {
    Serial.printf("[POST] OK Ist.%d seq=%lu (HTTP %d)\n", p.istasyon, p.seq, rc);
    buf_bas = (buf_bas + 1) % BUFFER_MAX;
    buf_dolu--;
    return true;
  } else {
    Serial.printf("[POST] HATA Ist.%d seq=%lu (HTTP %d) — kuyrukta kal\n",
                  p.istasyon, p.seq, rc);
    return false;
  }
}

// Sinyali kaydet (istasyon bazlı sayaç + buffer + NVS)
void istasyonSinyali(uint8_t istasyon) {
  uint32_t seq;
  if (istasyon == 1) {
    pulseIst1++;
    seq = pulseIst1;
    prefs.putULong("pulse_i1", pulseIst1);
  } else if (istasyon == 2) {
    pulseIst2++;
    seq = pulseIst2;
    prefs.putULong("pulse_i2", pulseIst2);
  } else {
    return;
  }

  // Buffer'a ekle (dolu ise en eskiyi at)
  if (buf_dolu >= BUFFER_MAX) {
    Serial.println("[BUFFER] DOLU! En eski pulse atildi");
    buf_bas = (buf_bas + 1) % BUFFER_MAX;
    buf_dolu--;
  }
  buffer[buf_son].seq      = seq;
  buffer[buf_son].istasyon = istasyon;
  buf_son = (buf_son + 1) % BUFFER_MAX;
  buf_dolu++;

  Serial.printf("[PULSE] Ist.%d #%lu (kuyruk=%d) — gonderiliyor...\n",
                istasyon, seq, buf_dolu);
  ledYakBlink(2, 50);

  if (wifiHazir()) bufferdanBirGonder();
}

void heartbeatGonder() {
  if (!wifiHazir()) return;
  HTTPClient http;
  http.begin(String(SUNUCU_HOST) + "/api/sinyal/heartbeat");
  http.setTimeout(5000);
  http.addHeader("Content-Type", "application/json");
  http.addHeader("Authorization", String("Bearer ") + API_TOKEN);

  StaticJsonDocument<512> doc;
  doc["cihaz_id"]        = CIHAZ_ID;
  doc["bolum"]           = BOLUM;
  doc["robot_no"]        = ROBOT_NO;
  doc["firmware_ver"]    = FIRMWARE_VER;
  doc["ip_adresi"]       = WiFi.localIP().toString();
  doc["mac_adresi"]      = WiFi.macAddress();
  doc["wifi_rssi"]       = WiFi.RSSI();
  doc["buffer_kuyruk"]   = buf_dolu;
  doc["uptime_sn"]       = (millis() - bootMs) / 1000;
  doc["free_heap"]       = (int)ESP.getFreeHeap();
  doc["pulse_ist1"]      = pulseIst1;
  doc["pulse_ist2"]      = pulseIst2;
  doc["robot_calisiyor"] = robotCalisiyor;

  String payload; serializeJson(doc, payload);
  int rc = http.POST(payload);
  http.end();
  Serial.printf("[HEART] HTTP %d · RSSI=%d · kuyruk=%d · ist1=%lu · ist2=%lu · robot=%s\n",
                rc, WiFi.RSSI(), buf_dolu, pulseIst1, pulseIst2,
                robotCalisiyor ? "ON" : "OFF");
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
  Serial.printf("FW: %s · Bolum: %s · Robot: %s · 2 istasyon + durum\n",
                FIRMWARE_VER, BOLUM, ROBOT_NO);
  Serial.printf("Pinler: Ist1=GPIO%d · Ist2=GPIO%d · Robot=GPIO%d\n",
                PIN_IST1_SAYAC, PIN_IST2_SAYAC, PIN_ROBOT_CALISIYOR);
  Serial.printf("MAC: %s\n", WiFi.macAddress().c_str());

  // GPIO — 3 giriş input-pullup
  pinMode(PIN_IST1_SAYAC,      INPUT_PULLUP);
  pinMode(PIN_IST2_SAYAC,      INPUT_PULLUP);
  pinMode(PIN_ROBOT_CALISIYOR, INPUT_PULLUP);
  pinMode(PIN_LED,             OUTPUT);
  digitalWrite(PIN_LED, LOW);

  lastIst1State  = digitalRead(PIN_IST1_SAYAC);
  lastIst2State  = digitalRead(PIN_IST2_SAYAC);
  lastRobotState = digitalRead(PIN_ROBOT_CALISIYOR);
  robotCalisiyor = (lastRobotState == LOW);  // röle kapalı = robot çalışıyor

  // NVS — istasyon başına kalıcı sayaç
  prefs.begin("cofle", false);
  pulseIst1 = prefs.getULong("pulse_i1", 0);
  pulseIst2 = prefs.getULong("pulse_i2", 0);
  Serial.printf("[NVS] Kayitli sayaclar: Ist1=%lu · Ist2=%lu\n", pulseIst1, pulseIst2);

  // Watchdog (ESP32 core 3.x ve 2.x ile uyumlu)
#if ESP_ARDUINO_VERSION_MAJOR >= 3
  esp_task_wdt_config_t wdt_config = {
    .timeout_ms     = (uint32_t)(WDT_TIMEOUT_S * 1000),
    .idle_core_mask = (1 << portNUM_PROCESSORS) - 1,
    .trigger_panic  = true,
  };
  esp_err_t wdt_err = esp_task_wdt_init(&wdt_config);
  if (wdt_err == ESP_ERR_INVALID_STATE) {
    esp_task_wdt_reconfigure(&wdt_config);
  }
#else
  esp_task_wdt_init(WDT_TIMEOUT_S, true);
#endif
  esp_task_wdt_add(NULL);

  wifiBaglan();
  delay(500);

  // OTA — WiFi üzerinden firmware güncelleme
  ArduinoOTA.setHostname(CIHAZ_ID);
  ArduinoOTA.setPassword(OTA_PASS);
  ArduinoOTA.onStart([]() {
    Serial.println("\n[OTA] Guncelleme basliyor — pin okumalari duruyor");
  });
  ArduinoOTA.onEnd([]() {
    Serial.println("\n[OTA] Tamam, yeniden baslatiliyor");
  });
  ArduinoOTA.onProgress([](unsigned int p, unsigned int t) {
    Serial.printf("[OTA] %u%%\r", (p / (t / 100)));
  });
  ArduinoOTA.onError([](ota_error_t e) {
    Serial.printf("[OTA] HATA %u\n", e);
  });
  ArduinoOTA.begin();
  Serial.printf("[OTA] Aktif — IDE Network Port: %s @ %s\n",
                CIHAZ_ID, WiFi.localIP().toString().c_str());

  // İlk heartbeat
  heartbeatGonder();
  lastHeartbeat = millis();
  lastRetry = millis();

  Serial.println("[READY] 3 girisli sayac aktif — robot DO sinyalleri bekleniyor...\n");
  ledYakBlink(3, 60);
}

void loop() {
  esp_task_wdt_reset();
  ArduinoOTA.handle();  // OTA güncellemesi bekliyorsa işle
  unsigned long now = millis();

  // ── 1. WiFi kontrolu ──
  if (!wifiHazir()) {
    digitalWrite(PIN_LED, (now / 200) % 2);
    if ((now - lastHeartbeat) > 15000) {
      wifiBaglan();
      lastHeartbeat = now;
    }
  }

  // ── 2. İstasyon 1 sayaç (debounce + multi-sample + min interval) ──
  int curIst1 = digitalRead(PIN_IST1_SAYAC);
  if (curIst1 != lastIst1State && (now - lastDebounceIst1) > DEBOUNCE_MS) {
    lastDebounceIst1 = now;
    if (curIst1 == LOW) {
      // FILTRE 1: Multi-sample teyit
      if (pinGercektenLOW(PIN_IST1_SAYAC)) {
        // FILTRE 2: Son pulse'tan en az MIN_PULSE_GAP_MS geçti mi?
        if ((now - lastValidPulseIst1) >= MIN_PULSE_GAP_MS) {
          istasyonSinyali(1);
          lastValidPulseIst1 = now;
        } else {
          Serial.printf("[FILTRE] Ist.1 erken (gap=%lums < %lums) - SAYILMADI\n",
                        now - lastValidPulseIst1, MIN_PULSE_GAP_MS);
        }
      } else {
        Serial.println("[FILTRE] Ist.1 parazit (multi-sample basarisiz) - SAYILMADI");
      }
    }
    lastIst1State = curIst1;
  }

  // ── 3. İstasyon 2 sayaç ──
  int curIst2 = digitalRead(PIN_IST2_SAYAC);
  if (curIst2 != lastIst2State && (now - lastDebounceIst2) > DEBOUNCE_MS) {
    lastDebounceIst2 = now;
    if (curIst2 == LOW) {
      if (pinGercektenLOW(PIN_IST2_SAYAC)) {
        if ((now - lastValidPulseIst2) >= MIN_PULSE_GAP_MS) {
          istasyonSinyali(2);
          lastValidPulseIst2 = now;
        } else {
          Serial.printf("[FILTRE] Ist.2 erken (gap=%lums < %lums) - SAYILMADI\n",
                        now - lastValidPulseIst2, MIN_PULSE_GAP_MS);
        }
      } else {
        Serial.println("[FILTRE] Ist.2 parazit (multi-sample basarisiz) - SAYILMADI");
      }
    }
    lastIst2State = curIst2;
  }

  // ── 4. Robot çalışıyor durumu (periyodik stabilite + çoğunluk oylaması) ──
  // Edge-triggered yerine her 500ms'de 10 örnek alınır, çoğunluk hangi yöndeyse
  // gerçek durum kabul edilir. EMI ile gelen anlık değişimler süzülür.
  if (now - lastRobotCheck >= ROBOT_CHECK_INTERVAL_MS) {
    lastRobotCheck = now;
    int lowSayim = 0;
    for (int i = 0; i < ROBOT_SAMPLE_N; i++) {
      if (digitalRead(PIN_ROBOT_CALISIYOR) == LOW) lowSayim++;
      delay(2);
    }
    bool yeniDurum = (lowSayim >= ROBOT_SAMPLE_THRESHOLD);  // LOW çoğunluksa ÇALIŞIYOR
    if (yeniDurum != robotCalisiyor) {
      robotCalisiyor = yeniDurum;
      lastRobotStateChangeMs = now;
      Serial.printf("[ROBOT] Durum: %s (low/total=%d/%d)\n",
                    robotCalisiyor ? "ÇALIŞIYOR" : "DURDU",
                    lowSayim, ROBOT_SAMPLE_N);
      // State değişiminde anında heartbeat gönder ki UI hemen güncellensin
      if (wifiHazir()) heartbeatGonder();
    }
  }

  // ── 5. Buffer'da bekleyen pulse'ları gönder ──
  if (wifiHazir() && buf_dolu > 0 && (now - lastRetry) > RETRY_MS) {
    bufferdanBirGonder();
    lastRetry = now;
  }

  // ── 6. Heartbeat ──
  if (wifiHazir() && (now - lastHeartbeat) > HEARTBEAT_MS) {
    heartbeatGonder();
    lastHeartbeat = now;
  }

  delay(5);
}
