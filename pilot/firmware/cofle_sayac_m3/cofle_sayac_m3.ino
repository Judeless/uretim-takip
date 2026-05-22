/* ============================================================
 *  COFLE PILOT SAYAC — MONTAJ-M3 firmware (v2.1)
 *  >>> OTOMATIK URETILDI: generate.py — manuel duzenleme!
 * ============================================================
 *
 *  HEDEF:  Montaj masasi operatorunun butonuna her basisinda
 *          1 uretim pulse'u olarak kaydeder.
 *
 *  PIN ATAMALARI:
 *   - GPIO25  -> BUTON (NO momentary, GND'ye kapanir)
 *   - GPIO26  -> BOS (INPUT_PULLUP)
 *   - GPIO27  -> BOS (INPUT_PULLUP)
 *   - GND     -> Buton 2. pini
 *
 *  Bu dosya pilot/firmware/_templates/montaj.ino.tpl'den uretildi.
 * ============================================================ */

#include <WiFi.h>
#include <HTTPClient.h>
#include <Preferences.h>
#include <ArduinoJson.h>
#include <esp_task_wdt.h>
#include <ArduinoOTA.h>

// ════════════════════════════════════════════════════════════
//   YAPILANDIRMA — M3
// ════════════════════════════════════════════════════════════

const char* CIHAZ_ID  = "MONTAJ-M3";
const char* BOLUM     = "montaj";
const char* ROBOT_NO  = "M3";

const char* WIFI_SSID = "COFLE-TK";
const char* WIFI_PASS = "internet2011!";
const char* SUNUCU_HOST = "http://192.168.21.155:5001";
const char* API_TOKEN = "cofle-pilot-2026";
const char* OTA_PASS = "cofle-ota-2026";

const int PIN_IST1_SAYAC      = 25;   // Buton (NO temas)
const int PIN_IST2_SAYAC      = 26;   // BOS
const int PIN_ROBOT_CALISIYOR = 27;   // BOS
const int PIN_LED             =  2;

// Pulse algilama: GPIO CHANGE interrupt + pulse genislik teyidi (state machine)
// Cross-talk spike'lari ve glitch'ler filtrelenir; sadece tamamlanmis pulse'lar sayilir.
const unsigned long ISR_DEBOUNCE_US  = 1000;   // 1ms ISR seviyesinde minimal debounce
const unsigned long MIN_PULSE_US     = 5000;   // 5ms minimum buton basis suresi (spike koruma)
const unsigned long MIN_PULSE_GAP_MS = 500;    // Montaj: hizli operator (500ms back-to-back kabul)
const int  HEARTBEAT_MS   = 30000;
const int  RETRY_MS       = 3000;
const int  WIFI_TIMEOUT_S = 30;
const int  WDT_TIMEOUT_S  = 30;
const int  BUFFER_MAX     = 200;
const char* FIRMWARE_VER  = "2.2.1-m3";

// ════════════════════════════════════════════════════════════
//   GLOBAL DURUM
// ════════════════════════════════════════════════════════════

Preferences prefs;

uint32_t pulseIst1 = 0;
uint32_t pulseIst2 = 0;

bool robotCalisiyor = false;

struct PulseKaydi {
  uint32_t seq;
  uint8_t  istasyon;
};
PulseKaydi buffer[BUFFER_MAX];
int buf_bas = 0, buf_son = 0, buf_dolu = 0;

unsigned long lastValidPulseIst1 = 0;

// Hardware CHANGE interrupt + state machine — buton pulse pini (PIN_IST1_SAYAC).
// state: 0=bekliyor, 1=basili (LOW), 2=birakildi ve gecerli (loop sayar)
// Cross-talk veya kisa glitch'ler (<MIN_PULSE_US) reddedilir.
volatile uint8_t ist1_state = 0;
volatile unsigned long ist1_pulse_start_us = 0;
volatile uint32_t ist1_isr_count = 0;  // toplam ISR tetiklenme (debug)

unsigned long lastHeartbeat = 0;
unsigned long lastRetry     = 0;
unsigned long bootMs        = 0;

// Self-heal sayaclari — Smart Counter'in donanim watchdog'unu yazilim ile taklit
int httpFailCount = 0;
unsigned long lastSuccessHttp = 0;
int disconnectCount = 0;
unsigned long lastDisconnectMs = 0;
const unsigned long UPTIME_RESET_MS = 24UL * 3600UL * 1000UL;

// ════════════════════════════════════════════════════════════
//   YARDIMCI FONKSIYONLAR
// ════════════════════════════════════════════════════════════

void ledYakBlink(int adet, int sure_ms = 80) {
  for (int i = 0; i < adet; i++) {
    digitalWrite(PIN_LED, HIGH); delay(sure_ms);
    digitalWrite(PIN_LED, LOW);  delay(sure_ms);
  }
}

void wifiBaglan() {
  Serial.printf("[WiFi] '%s' agina baglaniliyor...\n", WIFI_SSID);
  WiFi.persistent(false);
  WiFi.setAutoReconnect(true);
  WiFi.mode(WIFI_STA);
  WiFi.setSleep(false);
  WiFi.setTxPower(WIFI_POWER_19_5dBm);
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

// ════════════════════════════════════════════════════════════
//   INTERRUPT SERVICE ROUTINE (CHANGE — hem rising hem falling)
// ════════════════════════════════════════════════════════════
// CHANGE interrupt + state machine: pulse'un basini ve sonunu izler.
// Sadece >=MIN_PULSE_US suren pulse'lar gecerli sayilir (spike filtresi).
void IRAM_ATTR onIst1Change() {
  unsigned long now_us = micros();
  ist1_isr_count++;
  if (digitalRead(PIN_IST1_SAYAC) == LOW) {
    // Falling — buton basildi
    if (ist1_state == 0 && (now_us - ist1_pulse_start_us) > ISR_DEBOUNCE_US) {
      ist1_state = 1;
      ist1_pulse_start_us = now_us;
    }
  } else {
    // Rising — buton birakildi
    if (ist1_state == 1) {
      unsigned long genislik = now_us - ist1_pulse_start_us;
      if (genislik >= MIN_PULSE_US) {
        ist1_state = 2;
      } else {
        ist1_state = 0;  // spike, atla
      }
    }
  }
}

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
  doc["kaynak_tip"] = "buton";
  String mac6 = WiFi.macAddress();
  mac6.replace(":", "");
  if (mac6.length() > 6) mac6 = mac6.substring(mac6.length() - 6);
  String idem = String(CIHAZ_ID) + "_" + mac6 + "_i" + String(p.istasyon) + "_" + String(p.seq);
  doc["idempotency_key"] = idem;

  String payload;
  serializeJson(doc, payload);

  int rc = http.POST(payload);
  http.end();

  if (rc == 200 || rc == 201) {
    Serial.printf("[POST] OK Ist.%d seq=%lu (HTTP %d)\n", p.istasyon, p.seq, rc);
    buf_bas = (buf_bas + 1) % BUFFER_MAX;
    buf_dolu--;
    httpFailCount = 0;
    lastSuccessHttp = millis();
    return true;
  } else {
    Serial.printf("[POST] HATA Ist.%d seq=%lu (HTTP %d) — kuyrukta kal\n",
                  p.istasyon, p.seq, rc);
    httpFailCount++;
    return false;
  }
}

void istasyonSinyali(uint8_t istasyon) {
  uint32_t seq;
  if (istasyon == 1) {
    pulseIst1++;
    seq = pulseIst1;
    prefs.putULong("pulse_i1", pulseIst1);
  } else {
    return;
  }

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
  // ISR raw count — MIN_PULSE_GAP filtresinden onceki donanim pulse sayisi
  doc["isr_count_ist1"]  = ist1_isr_count;
  doc["robot_calisiyor"] = robotCalisiyor;

  String payload; serializeJson(doc, payload);
  int rc = http.POST(payload);
  http.end();
  if (rc == 200 || rc == 201) {
    httpFailCount = 0;
    lastSuccessHttp = millis();
  } else {
    httpFailCount++;
  }
  Serial.printf("[HEART] HTTP %d · RSSI=%d · kuyruk=%d · buton_sayim=%lu · fail=%d\n",
                rc, WiFi.RSSI(), buf_dolu, pulseIst1, httpFailCount);
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
  Serial.printf("FW: %s · Bolum: %s · Masa: %s · MONTAJ BUTON MODU\n",
                FIRMWARE_VER, BOLUM, ROBOT_NO);
  Serial.printf("Buton pini: GPIO%d (GND'ye basinca pulse)\n", PIN_IST1_SAYAC);
  Serial.printf("Min pulse araligi: %lums (operator hizi siniri)\n", MIN_PULSE_GAP_MS);
  Serial.printf("MAC: %s\n", WiFi.macAddress().c_str());

  pinMode(PIN_IST1_SAYAC,      INPUT_PULLUP);
  pinMode(PIN_IST2_SAYAC,      INPUT_PULLUP);
  pinMode(PIN_ROBOT_CALISIYOR, INPUT_PULLUP);
  pinMode(PIN_LED,             OUTPUT);
  digitalWrite(PIN_LED, LOW);

  robotCalisiyor = false;  // Montajda robot durumu kullanilmiyor

  // Hardware CHANGE interrupt — hem buton basis hem birakma yakalanir,
  // sadece >=MIN_PULSE_US suren basislar gecerli sayilir (spike/glitch filtresi).
  attachInterrupt(digitalPinToInterrupt(PIN_IST1_SAYAC), onIst1Change, CHANGE);
  Serial.printf("[ISR] Buton=GPIO%d — CHANGE interrupt + pulse genislik teyidi (>=%lums)\n",
                PIN_IST1_SAYAC, MIN_PULSE_US / 1000);

  prefs.begin("cofle", false);
  pulseIst1 = prefs.getULong("pulse_i1", 0);
  pulseIst2 = prefs.getULong("pulse_i2", 0);
  Serial.printf("[NVS] Kayitli buton sayimi: %lu\n", pulseIst1);

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

  // WiFi event handler — disconnect olayinda sayim
  WiFi.onEvent([](WiFiEvent_t event, WiFiEventInfo_t info) {
    if (event == ARDUINO_EVENT_WIFI_STA_DISCONNECTED) {
      Serial.printf("[WiFi] DISCONNECTED — sebep=%d\n",
                    info.wifi_sta_disconnected.reason);
      unsigned long now = millis();
      if ((now - lastDisconnectMs) > 600000UL) {
        disconnectCount = 0;
      }
      disconnectCount++;
      lastDisconnectMs = now;
    }
  });

  wifiBaglan();
  delay(500);

  ArduinoOTA.setHostname(CIHAZ_ID);
  ArduinoOTA.setPassword(OTA_PASS);
  ArduinoOTA.onStart([]() { Serial.println("\n[OTA] Guncelleme basliyor"); });
  ArduinoOTA.onEnd([]()   { Serial.println("\n[OTA] Tamam, yeniden baslatiliyor"); });
  ArduinoOTA.onProgress([](unsigned int p, unsigned int t) {
    Serial.printf("[OTA] %u%%\r", (p / (t / 100)));
  });
  ArduinoOTA.onError([](ota_error_t e) { Serial.printf("[OTA] HATA %u\n", e); });
  ArduinoOTA.begin();
  Serial.printf("[OTA] Aktif — IDE Network Port: %s @ %s\n",
                CIHAZ_ID, WiFi.localIP().toString().c_str());

  heartbeatGonder();
  lastHeartbeat = millis();
  lastRetry = millis();

  Serial.println("[READY] CHANGE interrupt + pulse genislik teyidi aktif — operator basisi bekleniyor...");
  Serial.println("        (Cross-talk ve glitch koruma: buton basisi >=5ms olmali)\n");
  ledYakBlink(3, 60);
}

void loop() {
  esp_task_wdt_reset();
  ArduinoOTA.handle();
  unsigned long now = millis();

  if (!wifiHazir()) {
    digitalWrite(PIN_LED, (now / 200) % 2);
    if ((now - lastHeartbeat) > 15000) {
      wifiBaglan();
      lastHeartbeat = now;
    }
  }

  // Buton pulse — ISR state machine tarafindan 2'ye getirildi (tamamlanmis gecerli pulse)
  if (ist1_state == 2) {
    ist1_state = 0;
    if ((now - lastValidPulseIst1) >= MIN_PULSE_GAP_MS) {
      istasyonSinyali(1);
      lastValidPulseIst1 = now;
    } else {
      Serial.printf("[FILTRE] Buton erken (gap=%lums < %lums) - SAYILMADI\n",
                    now - lastValidPulseIst1, MIN_PULSE_GAP_MS);
    }
  }

  if (wifiHazir() && buf_dolu > 0 && (now - lastRetry) > RETRY_MS) {
    bufferdanBirGonder();
    lastRetry = now;
  }

  if (wifiHazir() && (now - lastHeartbeat) > HEARTBEAT_MS) {
    heartbeatGonder();
    lastHeartbeat = now;
  }

  // ─── SELF-HEAL kontrolleri ───
  if ((now - bootMs) > UPTIME_RESET_MS) {
    Serial.println("\n[SELFHEAL] 24 saat uptime doldu — koruyucu reset");
    delay(500);
    ESP.restart();
  }
  if (httpFailCount >= 5 && lastSuccessHttp > 0
      && (now - lastSuccessHttp) > 60000UL) {
    Serial.printf("\n[SELFHEAL] %d ardisik HTTP fail + 60sn sessizlik — RESET\n",
                  httpFailCount);
    delay(500);
    ESP.restart();
  }
  if (disconnectCount >= 3 && (now - lastDisconnectMs) < 600000UL) {
    Serial.printf("\n[SELFHEAL] 10dk icinde %d disconnect — RESET\n",
                  disconnectCount);
    delay(500);
    ESP.restart();
  }

  delay(5);
}
