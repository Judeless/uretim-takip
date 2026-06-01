/* ============================================================
 *  COFLE PILOT SAYAC — ABB7-IO firmware (v2.1)
 *  >>> OTOMATIK URETILDI: generate.py — manuel duzenleme!
 * ============================================================
 *
 *  HEDEF:  ABB7 robotunun 3 dijital cikisini izler:
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
//   YAPILANDIRMA — ABB7
// ════════════════════════════════════════════════════════════

const char* CIHAZ_ID  = "ABB7-IO";
const char* BOLUM     = "kaynak";
const char* ROBOT_NO  = "ABB7";

const char* WIFI_SSID = "COFLE-TK";
const char* WIFI_PASS = "internet2011!";
const char* SUNUCU_HOST = "http://192.168.21.155:5001";
const char* API_TOKEN = "cofle-pilot-2026";
const char* OTA_PASS = "cofle-ota-2026";

const int PIN_IST1_SAYAC      = 25;
const int PIN_IST2_SAYAC      = 26;
const int PIN_ROBOT_CALISIYOR = 27;
const int PIN_LED             =  2;

// Pulse algilama: state polling + multi-sample (v2.1 kanitlanmis yontem).
// Pini surekli okur, HIGH->LOW gecisinde 75ms boyunca (15x5ms) hala LOW mu
// dogrular. Bu yontem:
//  - Cross-talk spike'larini eler (kisa spike 75ms LOW kalamaz)
//  - Yavas/zayif pull-up kenarlarinda guvenilir (oturmus seviyeye bakar,
//    interrupt edge timing'ine degil)
//  - MIN_PULSE_GAP ile ayni pulse'in cift sayilmasini onler
const int  DEBOUNCE_MS         = 100;          // edge sonrasi bekleme
const int  PARAZIT_SAMPLE_N    = 15;           // multi-sample adedi
const int  PARAZIT_SAMPLE_GAP  = 5;            // sample arasi ms (15x5 = 75ms onay)
const unsigned long MIN_PULSE_GAP_MS = 3000;   // iki gecerli pulse arasi min (robot cycle 8sn+)
const int  HEARTBEAT_MS   = 30000;
const int  RETRY_MS       = 3000;
const int  WIFI_TIMEOUT_S = 30;
const int  WDT_TIMEOUT_S  = 30;
const int  BUFFER_MAX     = 200;
const char* FIRMWARE_VER  = "2.3.1-abb7";

// ─── RSSI tabanli radyo recovery (v2.3) ─────────────────────────
// Cihaz "connected" gorunse bile sinyal cok zayifsa (ornek -88dBm iken
// komsu cihaz -39dBm) bu cogu zaman ESP32 RF yiginin takilmasidir.
// Yumusak radyo yenileme (WIFI_OFF -> WIFI_STA) RF'i yeniden baslatir;
// ESP.restart()'tan daha hedefli ve pulse kaybetmez (interrupt + buffer korur).
const int           RSSI_ZAYIF_ESIK  = -80;     // dBm — alti "zayif" kabul edilir
const unsigned long RSSI_KONTROL_MS  = 60000;   // her 60 sn RSSI olc
const int           RSSI_ZAYIF_MAX   = 5;       // 5 ardisik zayif (~5 dk) -> radyo yenile
const int           RADYO_YENILE_MAX = 3;       // 3 yenileme hala zayifsa -> 1 kez ESP.restart()

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

int lastIst1State  = HIGH;
int lastIst2State  = HIGH;
int lastRobotState = HIGH;
unsigned long lastDebounceIst1  = 0;
unsigned long lastDebounceIst2  = 0;
unsigned long lastDebounceRobot = 0;

unsigned long lastValidPulseIst1 = 0;
unsigned long lastValidPulseIst2 = 0;

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

// RSSI recovery durumu
unsigned long lastRssiKontrol = 0;
int rssiZayifSayac    = 0;   // ardisik zayif RSSI okuma
int radyoYenileSayac  = 0;   // bu zayif periyotta kac kez radyo yenilendi
int sonRSSI           = 0;   // son olculen RSSI (heartbeat'te raporlanir)
uint32_t radyoYenileToplam = 0;  // boot'tan beri toplam radyo yenileme (debug)
// Hard reset'i sinirla — gercek anten/donanim arizasinda reset loop olmasin.
// RTC bellegi ESP.restart()'i atlatir, sadece power-cycle sifirlar.
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

// Yumusak radyo yenileme — RF stack'i tamamen kapatip acar.
// ESP.restart()'tan farki: cihaz reboot etmez, NVS/sayaclar/buffer korunur,
// pulse'lar interrupt + buffer ile kaybolmaz. Takilmis RSSI durumunu cozer.
void wifiRadyoYenile() {
  Serial.println("\n[RSSI] Radyo yenileniyor (WIFI_OFF -> WIFI_STA, RF re-init)...");
  WiFi.disconnect(true);     // baglantiyi kes + radyoyu kapat
  delay(300);
  WiFi.mode(WIFI_OFF);       // RF tamamen kapat
  delay(500);
  WiFi.mode(WIFI_STA);       // RF yeniden baslat
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
    Serial.println("\n[RSSI] Radyo yenileme sonrasi baglanamadi — sonra tekrar");
  }
}

// Multi-sample parazit filtresi — pin PARAZIT_SAMPLE_N kez ust uste LOW mu?
// Cross-talk spike'lari ve yavas kenarlardaki anlik gurultuyu eler.
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
  http.setTimeout(3000);   // kisa timeout — uzun blok = pulse polling gecikmesi
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
  doc["robot_calisiyor"] = robotCalisiyor;
  // RSSI recovery istatistigi — radyo kac kez yenilendi (yuksekse anten/konum sorunu)
  doc["radyo_yenile"]    = radyoYenileToplam;

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

  lastIst1State  = digitalRead(PIN_IST1_SAYAC);
  lastIst2State  = digitalRead(PIN_IST2_SAYAC);
  lastRobotState = digitalRead(PIN_ROBOT_CALISIYOR);
  robotCalisiyor = (lastRobotState == LOW);

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

  Serial.println("[READY] State polling + multi-sample aktif — robot DO sinyalleri bekleniyor...");
  Serial.println("        (75ms multi-sample cross-talk/parazit filtresi, MIN_PULSE_GAP=3sn)\n");
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

  // Ist.1 pulse — state polling: HIGH->LOW gecisi + 75ms multi-sample + min gap
  int curIst1 = digitalRead(PIN_IST1_SAYAC);
  if (curIst1 != lastIst1State && (now - lastDebounceIst1) > DEBOUNCE_MS) {
    lastDebounceIst1 = now;
    if (curIst1 == LOW) {
      if (pinGercektenLOW(PIN_IST1_SAYAC)) {
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
  // Ist.2 pulse
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
  // 1) 24 saat uptime → koruyucu reset, AMA makine bos iken (uretim kesintiye ugramasin)
  //    Son pulse'tan beri 2 dk gectiyse (iki istasyon da) reset guvenli.
  if ((now - bootMs) > UPTIME_RESET_MS) {
    bool makineBos = (now - lastValidPulseIst1) > 120000UL
                  && (now - lastValidPulseIst2) > 120000UL;
    if (makineBos) {
      Serial.println("\n[SELFHEAL] 24h doldu + makine bos (2dk pulse yok) — koruyucu reset");
      delay(500);
      ESP.restart();
    }
    // Makine calisiyorsa reset ertelenir — bir sonraki bos doneme birakilir
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
  // 4) RSSI tabanli radyo recovery — connected ama sinyal cok zayifsa.
  //    KRITIK: radyo yenileme/reset BLOKLAR -> sadece makine BOS iken yap,
  //    yoksa uretim sirasinda pulse kaybi olur. Zayif sinyal acil degil:
  //    cihaz connected oldugu surece buffer + retry veriyi tasir.
  if (wifiHazir() && (now - lastRssiKontrol) > RSSI_KONTROL_MS) {
    lastRssiKontrol = now;
    sonRSSI = WiFi.RSSI();
    bool makineBosRssi = (now - lastValidPulseIst1) > 120000UL
                      && (now - lastValidPulseIst2) > 120000UL;
    if (sonRSSI < RSSI_ZAYIF_ESIK && sonRSSI < 0) {
      rssiZayifSayac++;
      Serial.printf("[RSSI] Zayif sinyal %d dBm (%d/%d ardisik)%s\n",
                    sonRSSI, rssiZayifSayac, RSSI_ZAYIF_MAX,
                    makineBosRssi ? "" : " — makine calisiyor, recovery ertelendi");
      if (rssiZayifSayac >= RSSI_ZAYIF_MAX && makineBosRssi) {
        rssiZayifSayac = 0;
        radyoYenileSayac++;
        wifiRadyoYenile();   // yumusak RF re-init (sadece bos iken — blok guvenli)
        // Radyo yenileme yeterli sayida denenip hala zayifsa, RF stack degil
        // gercek anten/donanim sorunu olabilir. Power-cycle basina en fazla 2 kez
        // tam reset dene (RTC bellegi sayaci), sonra zayif sinyalle calismaya devam.
        if (radyoYenileSayac >= RADYO_YENILE_MAX) {
          radyoYenileSayac = 0;
          if (rtcRssiHardReset < 2) {
            rtcRssiHardReset++;
            Serial.printf("\n[SELFHEAL] RSSI zayif + radyo yenileme yetersiz — ESP.restart() (#%d)\n",
                          rtcRssiHardReset);
            delay(500);
            ESP.restart();
          } else {
            Serial.println("\n[RSSI] Hard reset limiti — muhtemel anten/donanim sorunu, "
                           "zayif sinyalle devam (buffer + retry koruyor)");
          }
        }
      }
    } else {
      // Sinyal iyi — sayaclari sifirla
      if (rssiZayifSayac > 0 || radyoYenileSayac > 0) {
        Serial.printf("[RSSI] Sinyal toparlandi: %d dBm\n", sonRSSI);
      }
      rssiZayifSayac = 0;
      radyoYenileSayac = 0;
    }
  }

  delay(5);
}
