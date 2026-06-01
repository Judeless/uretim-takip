/* ============================================================
 *  COFLE PILOT SAYAC — 550T-IO firmware (v2.1)
 *  >>> OTOMATIK URETILDI: generate.py — manuel duzenleme!
 * ============================================================
 *
 *  HEDEF:  550T metal enjeksiyon makinesinin tek uretim
 *          sinyalini izler. Her parca cikisinda 1 pulse.
 *
 *  PIN ATAMALARI:
 *   - GPIO25  -> Role 1 NO  (Uretim sinyali — BEYAZ tel)
 *   - GPIO26  -> BOS (kullanilmiyor)
 *   - GPIO27  -> BOS (kullanilmiyor)
 *   - GND     -> Role 1 COM (SARI tel)
 *
 *  Bu dosya pilot/firmware/_templates/metal.ino.tpl'den uretildi.
 * ============================================================ */

#include <WiFi.h>
#include <HTTPClient.h>
#include <Preferences.h>
#include <ArduinoJson.h>
#include <esp_task_wdt.h>
#include <ArduinoOTA.h>

// ════════════════════════════════════════════════════════════
//   YAPILANDIRMA — 550T
// ════════════════════════════════════════════════════════════

const char* CIHAZ_ID  = "550T-IO";
const char* BOLUM     = "metal";
const char* ROBOT_NO  = "550T";

const char* WIFI_SSID = "COFLE-TK";
const char* WIFI_PASS = "internet2011!";
const char* SUNUCU_HOST = "http://192.168.21.155:5001";
const char* API_TOKEN = "cofle-pilot-2026";
const char* OTA_PASS = "cofle-ota-2026";

const int PIN_IST1_SAYAC      = 25;   // Uretim sinyali
const int PIN_IST2_SAYAC      = 26;   // BOS
const int PIN_ROBOT_CALISIYOR = 27;   // BOS
const int PIN_LED             =  2;

// Pulse algilama: state polling + multi-sample (v2.1 kanitlanmis yontem).
// Rote'nin oturmus LOW seviyesine 75ms boyunca bakar; kisa parazit/cross-talk
// spike'larini eler. MIN_PULSE_GAP=15sn ile rote 10sn aktif periyodunda tek sayim.
const int  DEBOUNCE_MS         = 100;
const int  PARAZIT_SAMPLE_N    = 15;
const int  PARAZIT_SAMPLE_GAP  = 5;            // 15x5 = 75ms onay
const unsigned long MIN_PULSE_GAP_MS = 15000; // Metal: rote 10sn aktif, en kisa cycle 20sn → 15sn (10+5 margin)
const int  HEARTBEAT_MS   = 30000;
const int  RETRY_MS       = 3000;
const int  WIFI_TIMEOUT_S = 30;
const int  WDT_TIMEOUT_S  = 30;
const int  BUFFER_MAX     = 200;
const char* FIRMWARE_VER  = "2.3.1-550t";

// ─── RSSI tabanli radyo recovery (v2.3) ─────────────────────────
const int           RSSI_ZAYIF_ESIK  = -80;
const unsigned long RSSI_KONTROL_MS  = 60000;
const int           RSSI_ZAYIF_MAX   = 5;
const int           RADYO_YENILE_MAX = 3;

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

int lastIst1State = HIGH;
unsigned long lastDebounceIst1 = 0;
unsigned long lastValidPulseIst1 = 0;

unsigned long lastHeartbeat = 0;
unsigned long lastRetry     = 0;
unsigned long bootMs        = 0;

// Self-heal sayaclari — Smart Counter'in donanim watchdog'unu yazilim ile taklit
int httpFailCount = 0;
unsigned long lastSuccessHttp = 0;
int disconnectCount = 0;
unsigned long lastDisconnectMs = 0;
const unsigned long UPTIME_RESET_MS = 24UL * 3600UL * 1000UL;

// RSSI recovery durumu
unsigned long lastRssiKontrol = 0;
int rssiZayifSayac    = 0;
int radyoYenileSayac  = 0;
int sonRSSI           = 0;
uint32_t radyoYenileToplam = 0;
RTC_DATA_ATTR int rtcRssiHardReset = 0;

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

// Yumusak radyo yenileme — RF stack'i kapatip acar. Reboot etmez, sayac/buffer korunur.
void wifiRadyoYenile() {
  Serial.println("\n[RSSI] Radyo yenileniyor (WIFI_OFF -> WIFI_STA, RF re-init)...");
  WiFi.disconnect(true);
  delay(300);
  WiFi.mode(WIFI_OFF);
  delay(500);
  WiFi.mode(WIFI_STA);
  WiFi.setSleep(false);
  WiFi.setTxPower(WIFI_POWER_19_5dBm);
  WiFi.begin(WIFI_SSID, WIFI_PASS);
  unsigned long basla = millis();
  while (WiFi.status() != WL_CONNECTED && (millis() - basla) < 15000) {
    delay(500); Serial.print('+');
    esp_task_wdt_reset();
  }
  radyoYenileToplam++;
  if (WiFi.status() == WL_CONNECTED) {
    sonRSSI = WiFi.RSSI();
    Serial.printf("\n[RSSI] Radyo yenilendi · yeni RSSI=%d dBm\n", sonRSSI);
  } else {
    Serial.println("\n[RSSI] Radyo yenileme sonrasi baglanamadi");
  }
}

// Multi-sample parazit filtresi — pin PARAZIT_SAMPLE_N kez ust uste LOW mu?
bool pinGercektenLOW(int pin) {
  for (int i = 0; i < PARAZIT_SAMPLE_N; i++) {
    if (digitalRead(pin) != LOW) return false;
    delay(PARAZIT_SAMPLE_GAP);
  }
  return true;
}

bool bufferdanBirGonder() {
  if (buf_dolu == 0) return false;
  PulseKaydi p = buffer[buf_bas];

  HTTPClient http;
  String url = String(SUNUCU_HOST) + "/api/sinyal";
  http.begin(url);
  http.setTimeout(3000);
  http.addHeader("Content-Type", "application/json");
  http.addHeader("Authorization", String("Bearer ") + API_TOKEN);

  StaticJsonDocument<256> doc;
  doc["cihaz_id"]   = CIHAZ_ID;
  doc["bolum"]      = BOLUM;
  doc["robot_no"]   = ROBOT_NO;
  doc["istasyon"]   = p.istasyon;
  doc["kaynak_tip"] = "makine_io";
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
    Serial.printf("[POST] OK seq=%lu (HTTP %d)\n", p.seq, rc);
    buf_bas = (buf_bas + 1) % BUFFER_MAX;
    buf_dolu--;
    httpFailCount = 0;
    lastSuccessHttp = millis();
    return true;
  } else {
    Serial.printf("[POST] HATA seq=%lu (HTTP %d) — kuyrukta kal\n", p.seq, rc);
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

  Serial.printf("[PULSE] #%lu (kuyruk=%d) — gonderiliyor...\n", seq, buf_dolu);
  ledYakBlink(2, 50);

  if (wifiHazir()) bufferdanBirGonder();
}

void heartbeatGonder() {
  if (!wifiHazir()) return;
  HTTPClient http;
  http.begin(String(SUNUCU_HOST) + "/api/sinyal/heartbeat");
  http.setTimeout(3000);
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
  doc["radyo_yenile"]    = radyoYenileToplam;
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
  Serial.printf("[HEART] HTTP %d · RSSI=%d · kuyruk=%d · sayim=%lu · fail=%d\n",
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
  Serial.printf("FW: %s · Bolum: %s · Makine: %s · METAL ENJEKSIYON\n",
                FIRMWARE_VER, BOLUM, ROBOT_NO);
  Serial.printf("Uretim pini: GPIO%d (makine her parcada pulse)\n", PIN_IST1_SAYAC);
  Serial.printf("Min pulse araligi: %lums (makine cycle siniri)\n", MIN_PULSE_GAP_MS);
  Serial.printf("MAC: %s\n", WiFi.macAddress().c_str());

  pinMode(PIN_IST1_SAYAC,      INPUT_PULLUP);
  pinMode(PIN_IST2_SAYAC,      INPUT_PULLUP);
  pinMode(PIN_ROBOT_CALISIYOR, INPUT_PULLUP);
  pinMode(PIN_LED,             OUTPUT);
  digitalWrite(PIN_LED, LOW);

  lastIst1State  = digitalRead(PIN_IST1_SAYAC);
  robotCalisiyor = false;  // Metal'de robot durumu kullanilmiyor

  prefs.begin("cofle", false);
  pulseIst1 = prefs.getULong("pulse_i1", 0);
  pulseIst2 = prefs.getULong("pulse_i2", 0);
  Serial.printf("[NVS] Kayitli sayim: %lu\n", pulseIst1);

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

  Serial.println("[READY] State polling + multi-sample aktif — uretim sinyalleri bekleniyor...");
  Serial.println("        (75ms multi-sample parazit filtresi, MIN_PULSE_GAP=15sn)\n");
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

  // Uretim sinyali — state polling: HIGH->LOW gecisi + 75ms multi-sample + min gap
  int curIst1 = digitalRead(PIN_IST1_SAYAC);
  if (curIst1 != lastIst1State && (now - lastDebounceIst1) > DEBOUNCE_MS) {
    lastDebounceIst1 = now;
    if (curIst1 == LOW) {
      if (pinGercektenLOW(PIN_IST1_SAYAC)) {
        if ((now - lastValidPulseIst1) >= MIN_PULSE_GAP_MS) {
          istasyonSinyali(1);
          lastValidPulseIst1 = now;
        } else {
          Serial.printf("[FILTRE] Uretim erken (gap=%lums < %lums) - SAYILMADI\n",
                        now - lastValidPulseIst1, MIN_PULSE_GAP_MS);
        }
      } else {
        Serial.println("[FILTRE] Uretim parazit (multi-sample basarisiz) - SAYILMADI");
      }
    }
    lastIst1State = curIst1;
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
  // 1) 24 saat uptime → koruyucu reset, AMA makine bos iken (uretim kesilmesin)
  if ((now - bootMs) > UPTIME_RESET_MS) {
    if ((now - lastValidPulseIst1) > 120000UL) {
      Serial.println("\n[SELFHEAL] 24h doldu + makine bos (2dk) — koruyucu reset");
      delay(500);
      ESP.restart();
    }
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
  // 4) RSSI tabanli radyo recovery — connected ama sinyal cok zayifsa
  if (wifiHazir() && (now - lastRssiKontrol) > RSSI_KONTROL_MS) {
    lastRssiKontrol = now;
    sonRSSI = WiFi.RSSI();
    if (sonRSSI < RSSI_ZAYIF_ESIK && sonRSSI < 0) {
      rssiZayifSayac++;
      Serial.printf("[RSSI] Zayif sinyal %d dBm (%d/%d)\n", sonRSSI, rssiZayifSayac, RSSI_ZAYIF_MAX);
      if (rssiZayifSayac >= RSSI_ZAYIF_MAX) {
        rssiZayifSayac = 0;
        radyoYenileSayac++;
        wifiRadyoYenile();
        if (radyoYenileSayac >= RADYO_YENILE_MAX) {
          radyoYenileSayac = 0;
          if (rtcRssiHardReset < 2) {
            rtcRssiHardReset++;
            Serial.printf("\n[SELFHEAL] RSSI zayif — ESP.restart() (#%d)\n", rtcRssiHardReset);
            delay(500);
            ESP.restart();
          } else {
            Serial.println("\n[RSSI] Hard reset limiti — zayif sinyalle devam");
          }
        }
      }
    } else {
      rssiZayifSayac = 0;
      radyoYenileSayac = 0;
    }
  }

  delay(5);
}
