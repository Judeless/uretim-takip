/* ============================================================
 *  COFLE PILOT SAYAC — ABB5-IO firmware (v2.1)
 *  >>> OTOMATIK URETILDI: generate.py — manuel duzenleme!
 * ============================================================
 *
 *  HEDEF:  ABB5 robotunun 3 dijital cikisini izler:
 *            - Istasyon 1 sayac (her parca = 1 pulse)
 *            - Istasyon 2 sayac (her parca = 1 pulse)
 *            - Robot calisiyor durumu (HIGH=calisiyor)
 *
 *  PIN ATAMALARI:
 *   - GPIO25  -> Role 1 NO  (Istasyon 1 — BEYAZ tel)
 *   - GPIO26  -> Role 2 NO  (Istasyon 2 — GRI tel)
 *   - GPIO27  -> Role 3 NO  (Robot durumu — KAHVE tel)
 *   - GND     -> 3 role COM ortak (SARI tel)
 *
 *  Bu dosya pilot/firmware/_templates/kaynak.ino.tpl'den uretildi.
 * ============================================================ */

#include <WiFi.h>
#include <HTTPClient.h>
#include <Preferences.h>
#include <ArduinoJson.h>
#include <esp_task_wdt.h>
#include <ArduinoOTA.h>

// ════════════════════════════════════════════════════════════
//   YAPILANDIRMA — ABB5
// ════════════════════════════════════════════════════════════

const char* CIHAZ_ID  = "ABB5-IO";
const char* BOLUM     = "kaynak";
const char* ROBOT_NO  = "ABB5";

const char* WIFI_SSID = "COFLE-TK";
const char* WIFI_PASS = "internet2011!";
const char* SUNUCU_HOST = "http://192.168.21.155:5001";
const char* API_TOKEN = "cofle-pilot-2026";
const char* OTA_PASS = "cofle-ota-2026";

const int PIN_IST1_SAYAC      = 25;
const int PIN_IST2_SAYAC      = 26;
const int PIN_ROBOT_CALISIYOR = 27;
const int PIN_LED             =  2;

// Pulse algilama: GPIO interrupt (hardware) — multi-sample state polling
// loop bloke olursa (HTTP timeout, OTA, WiFi reconnect) yine de pulse yakalanir
const unsigned long ISR_DEBOUNCE_US  = 30000;  // 30 ms ISR seviyesinde donanim debounce
const unsigned long MIN_PULSE_GAP_MS = 800;    // Loop seviyesinde back-to-back filtre
                                               // (kaynak robot cycle min 8sn, 800ms aralik yeterli buffer)
const int  HEARTBEAT_MS   = 30000;
const int  RETRY_MS       = 3000;
const int  WIFI_TIMEOUT_S = 30;
const int  WDT_TIMEOUT_S  = 30;
const int  BUFFER_MAX     = 200;
const char* FIRMWARE_VER  = "2.2.0-abb5";

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

int lastRobotState = HIGH;
unsigned long lastDebounceRobot = 0;

unsigned long lastValidPulseIst1 = 0;
unsigned long lastValidPulseIst2 = 0;

// Hardware interrupt — ist1 + ist2 pulse pinleri.
// ISR: FALLING edge'i mikrosaniyeler icinde yakalar, 30ms donanim debounce uygular,
// flag set eder. Loop flag'i gorur ve MIN_PULSE_GAP_MS filtresinden gecirip sayar.
// Bu yontemle loop ne kadar bloke olursa olsun (HTTP, OTA, WiFi) pulse kaybolmaz.
volatile bool ist1_pulse_flag = false;
volatile bool ist2_pulse_flag = false;
volatile unsigned long ist1_isr_last_us = 0;
volatile unsigned long ist2_isr_last_us = 0;
volatile uint32_t ist1_isr_count = 0;  // toplam ISR tetiklenme (debug/dogrulama)
volatile uint32_t ist2_isr_count = 0;

unsigned long lastRobotCheck = 0;
const unsigned long ROBOT_CHECK_INTERVAL_MS = 500;
const int           ROBOT_SAMPLE_N          = 10;
const int           ROBOT_SAMPLE_THRESHOLD  = 7;
unsigned long lastRobotStateChangeMs = 0;

unsigned long lastHeartbeat = 0;
unsigned long lastRetry     = 0;
unsigned long bootMs        = 0;

// Self-heal sayaclari — Smart Counter'in donanim watchdog'unu yazilim ile taklit
// eder. Manuel reset bagimliligi azalir.
int httpFailCount = 0;                   // Ardisik HTTP -1 / fail sayisi
unsigned long lastSuccessHttp = 0;       // Son basarili POST/heartbeat zamani
int disconnectCount = 0;                 // 10dk penceredeki disconnect sayisi
unsigned long lastDisconnectMs = 0;      // Son disconnect zamani
const unsigned long UPTIME_RESET_MS = 24UL * 3600UL * 1000UL;  // 24 saat koruyucu

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
  WiFi.persistent(false);              // Flash spam onlenir
  WiFi.setAutoReconnect(true);         // Kopunca otomatik reconnect
  WiFi.mode(WIFI_STA);
  WiFi.setSleep(false);                // KRITIK: power save kapat (RSSI dusmesin)
  WiFi.setTxPower(WIFI_POWER_19_5dBm); // Max TX gucu
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
//   INTERRUPT SERVICE ROUTINES (FALLING edge yakalama)
// ════════════════════════════════════════════════════════════
// IRAM_ATTR: ISR fonksiyonu IRAM'de saklanmali (cache miss sirasinda da calismali)
// 30ms ISR debounce: kontaktor titremesini kabaca filtreler; gercek tekrar
// pulse loop seviyesinde MIN_PULSE_GAP_MS ile kontrol edilir.
void IRAM_ATTR onIst1Falling() {
  unsigned long now_us = micros();
  if (now_us - ist1_isr_last_us > ISR_DEBOUNCE_US) {
    ist1_pulse_flag = true;
    ist1_isr_last_us = now_us;
    ist1_isr_count++;
  }
}
void IRAM_ATTR onIst2Falling() {
  unsigned long now_us = micros();
  if (now_us - ist2_isr_last_us > ISR_DEBOUNCE_US) {
    ist2_pulse_flag = true;
    ist2_isr_last_us = now_us;
    ist2_isr_count++;
  }
}

// (Legacy multi-sample helper — artik kullanilmiyor, interrupt-based algilamaya gecildi)
// bool pinGercektenLOW(int pin) — kaldirildi

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
  } else if (istasyon == 2) {
    pulseIst2++;
    seq = pulseIst2;
    prefs.putULong("pulse_i2", pulseIst2);
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
  // ISR raw count — sayim'a donen pulse'lardan farkli (MIN_PULSE_GAP filtresinden dusenler)
  // ist1_isr_count - pulseIst1 = "FILTRE" ile reddedilen back-to-back pulse sayisi
  doc["isr_count_ist1"]  = ist1_isr_count;
  doc["isr_count_ist2"]  = ist2_isr_count;
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
  Serial.printf("[HEART] HTTP %d · RSSI=%d · kuyruk=%d · ist1=%lu · ist2=%lu · robot=%s · fail=%d\n",
                rc, WiFi.RSSI(), buf_dolu, pulseIst1, pulseIst2,
                robotCalisiyor ? "ON" : "OFF", httpFailCount);
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

  pinMode(PIN_IST1_SAYAC,      INPUT_PULLUP);
  pinMode(PIN_IST2_SAYAC,      INPUT_PULLUP);
  pinMode(PIN_ROBOT_CALISIYOR, INPUT_PULLUP);
  pinMode(PIN_LED,             OUTPUT);
  digitalWrite(PIN_LED, LOW);

  lastRobotState = digitalRead(PIN_ROBOT_CALISIYOR);
  robotCalisiyor = (lastRobotState == LOW);

  // Hardware interrupt'lari kur — FALLING edge (pin HIGH'tan LOW'a gecince)
  // Robot DO pulse'lari mikrosaniyeler mertebesinde olsa bile yakalanir.
  attachInterrupt(digitalPinToInterrupt(PIN_IST1_SAYAC), onIst1Falling, FALLING);
  attachInterrupt(digitalPinToInterrupt(PIN_IST2_SAYAC), onIst2Falling, FALLING);
  Serial.printf("[ISR] Ist1=GPIO%d, Ist2=GPIO%d — FALLING edge interrupt kuruldu\n",
                PIN_IST1_SAYAC, PIN_IST2_SAYAC);

  prefs.begin("cofle", false);
  pulseIst1 = prefs.getULong("pulse_i1", 0);
  pulseIst2 = prefs.getULong("pulse_i2", 0);
  Serial.printf("[NVS] Kayitli sayaclar: Ist1=%lu · Ist2=%lu\n", pulseIst1, pulseIst2);

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

  // WiFi event handler — disconnect olayinda agresif reconnect + sayim
  WiFi.onEvent([](WiFiEvent_t event, WiFiEventInfo_t info) {
    if (event == ARDUINO_EVENT_WIFI_STA_DISCONNECTED) {
      Serial.printf("[WiFi] DISCONNECTED — sebep=%d\n",
                    info.wifi_sta_disconnected.reason);
      unsigned long now = millis();
      // 10dk'dan uzun sure once disconnect olduysa sayaci sifirla
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

  Serial.println("[READY] Interrupt-based pulse algilama aktif — robot DO sinyalleri bekleniyor...");
  Serial.println("        (Multi-sample state polling kaldirildi, donanim interrupt'i kullaniliyor)\n");
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

  // Ist.1 pulse — ISR tarafindan flag set edildi, biz sadece kontrol + filtreleme yapariz
  if (ist1_pulse_flag) {
    ist1_pulse_flag = false;
    if ((now - lastValidPulseIst1) >= MIN_PULSE_GAP_MS) {
      istasyonSinyali(1);
      lastValidPulseIst1 = now;
    } else {
      Serial.printf("[FILTRE] Ist.1 erken (gap=%lums < %lums) - SAYILMADI\n",
                    now - lastValidPulseIst1, MIN_PULSE_GAP_MS);
    }
  }
  // Ist.2 pulse
  if (ist2_pulse_flag) {
    ist2_pulse_flag = false;
    if ((now - lastValidPulseIst2) >= MIN_PULSE_GAP_MS) {
      istasyonSinyali(2);
      lastValidPulseIst2 = now;
    } else {
      Serial.printf("[FILTRE] Ist.2 erken (gap=%lums < %lums) - SAYILMADI\n",
                    now - lastValidPulseIst2, MIN_PULSE_GAP_MS);
    }
  }

  if (now - lastRobotCheck >= ROBOT_CHECK_INTERVAL_MS) {
    lastRobotCheck = now;
    int lowSayim = 0;
    for (int i = 0; i < ROBOT_SAMPLE_N; i++) {
      if (digitalRead(PIN_ROBOT_CALISIYOR) == LOW) lowSayim++;
      delay(2);
    }
    bool yeniDurum = (lowSayim >= ROBOT_SAMPLE_THRESHOLD);
    if (yeniDurum != robotCalisiyor) {
      robotCalisiyor = yeniDurum;
      lastRobotStateChangeMs = now;
      Serial.printf("[ROBOT] Durum: %s (low/total=%d/%d)\n",
                    robotCalisiyor ? "CALISIYOR" : "DURDU",
                    lowSayim, ROBOT_SAMPLE_N);
      if (wifiHazir()) heartbeatGonder();
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
  // 1) 24 saat uptime → koruyucu reset (Smart Counter benzeri preventive)
  if ((now - bootMs) > UPTIME_RESET_MS) {
    Serial.println("\n[SELFHEAL] 24 saat uptime doldu — koruyucu reset");
    delay(500);
    ESP.restart();
  }
  // 2) 5+ ardisik HTTP fail + 60sn sessizlik → reset (TCP yigini takilmis)
  if (httpFailCount >= 5 && lastSuccessHttp > 0
      && (now - lastSuccessHttp) > 60000UL) {
    Serial.printf("\n[SELFHEAL] %d ardisik HTTP fail + 60sn sessizlik — RESET\n",
                  httpFailCount);
    delay(500);
    ESP.restart();
  }
  // 3) 10dk icinde 3+ disconnect → reset (WiFi yigini bozulmus)
  if (disconnectCount >= 3 && (now - lastDisconnectMs) < 600000UL) {
    Serial.printf("\n[SELFHEAL] 10dk icinde %d disconnect — RESET\n",
                  disconnectCount);
    delay(500);
    ESP.restart();
  }

  delay(5);
}
