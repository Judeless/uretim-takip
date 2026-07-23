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

// Pulse algilama: BLOKLAMAYAN INTEGRATOR debounce (v2.5).
// Eski "15 ardisik ornek HEPSI LOW" (75ms) filtresi tek bir parazit HIGH ile
// sifirlandigi icin, kaynak EMI'si pulse aninda denk gelince pulse'i eliyordu (eksik sayim).
// Integrator: her dongude pini ornekle -> LOW'da +1 / HIGH'da -1 (0..MAX).
// integ >= YUKSEK_ESIK -> "gercek pulse" (say); histerezisle DUSUK_ESIK'te biter.
// Esik eski 75ms ile ayni hassasiyette ama tek glitch artik pulse'i oldurmez.
const int  INTEG_YUKSEK_ESIK   = 15;   // ~75ms net LOW -> pulse algilandi (eski filtre ile ayni)
const int  INTEG_MAX           = 40;   // tavan (~200ms) — glitch toleransi tamponu
const int  INTEG_DUSUK_ESIK    = 4;    // ~20ms HIGH -> pulse bitti (histerezis)
const unsigned long MIN_PULSE_GAP_MS = 3000;   // iki gecerli pulse arasi min (robot cycle 8sn+)
const int  HEARTBEAT_MS   = 30000;
const int  RETRY_MS       = 3000;
const int  WIFI_TIMEOUT_S = 30;
const int  WDT_TIMEOUT_S  = 30;
const int  BUFFER_MAX     = 2000;  // kesinti kuyrugu (eskiden 200) — elektrik varken gunlerce pulse tutar
const char* FIRMWARE_VER  = "2.6.2-abb5";   // 2.6.2: WDT+OTA fix, non-blocking reconnect, retry backoff, WiFi-down hard reset, churn onleme

// ─── TANI (diagnostic) — sayim filtresi kararlarini heartbeat ile gonderir ──
const int TANI_MAX          = 30;   // RAM ring buffer
const int TANI_HEARTBEAT_N  = 20;   // tek heartbeat'te en fazla kac olay

// ─── RSSI tabanli radyo recovery (v2.3) ─────────────────────────
// Cihaz "connected" gorunse bile sinyal cok zayifsa (ornek -88dBm iken
// komsu cihaz -39dBm) bu cogu zaman ESP32 RF yiginin takilmasidir.
// Yumusak radyo yenileme (WIFI_OFF -> WIFI_STA) RF'i yeniden baslatir;
// ESP.restart()'tan daha hedefli ve pulse kaybetmez (interrupt + buffer korur).
const int           RSSI_ZAYIF_ESIK  = -80;     // dBm — alti "zayif" kabul edilir
const unsigned long RSSI_KONTROL_MS  = 60000;   // her 60 sn RSSI olc
const int           RSSI_ZAYIF_MAX   = 5;       // 5 ardisik zayif (~5 dk) -> radyo yenile
const int           RADYO_YENILE_MAX = 3;       // 3 yenileme hala zayifsa -> 1 kez ESP.restart()

// ─── Zombie WL_CONNECTED kurtarma + NVS watermark (v2.6) ────────────
// Cihaz AP'ye bagli gorunup (RSSI iyi okunur, disconnect event'i ATISLANMAZ) ama
// L3/uygulama yolu (IP/soket/upstream) kopabilir -> heartbeat POST'lari sessizce fail,
// sunucu offline gosterir, SADECE manuel reset duzeltir. Eski self-heal kuyrukta pulse
// varken REBOOT ETMIYORDU -> sonsuz radyo-tazele livelock. Cozum: kuyruk dolu olsa bile
// 5dk'dir sunucuya ulasilamiyorsa watermark'lari (sent_i1/sent_i2) NVS'e yazip GUVENLE
// reboot et (boot'ta gonderilmemis araliklar tekrar gonderilir; bootId NVS'te stabil
// -> sunucu cift saymaz). Kaynak 2 istasyonlu: watermark her istasyon icin ayri.
const unsigned long SUNUCU_SESSIZ_HARD_RESET_MS = 300000UL;  // 5dk: bu sure sunucuya HIC ulasilamazsa watermark + reboot
const unsigned long WATERMARK_YAZ_MS            = 30000UL;   // NVS watermark yazma throttle (yipranma siniri; crash resend penceresi ~30s)

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
// volatile: WiFi event task'i yazar, loop okur (32-bit hizali erisim — kritik bolge gerekmez)
volatile int disconnectCount = 0;        // 10dk penceredeki disconnect sayisi
volatile unsigned long lastDisconnectMs = 0;  // Son disconnect zamani
unsigned long wifiKoptuMs = 0;           // wifiHazir() false'a dustugu an (0=bagli) — WiFi-down hard reset icin
unsigned long retryGecikmeMs = RETRY_MS; // basarisiz POST'ta ustel artar (cap 60s) — sunucu kapaliyken sampling'i korur
const unsigned long UPTIME_RESET_MS = 24UL * 3600UL * 1000UL;  // 24 saat koruyucu

// Zombie kurtarma + watermark durumu (v2.6) — kaynak 2 istasyonlu: her istasyon icin ayri watermark
unsigned long lastRadyoYenile = 0;   // zombie kurtarma radyo-tazele throttle (lastSuccessHttp'yi EZMEDEN)
unsigned long lastWatermark   = 0;   // NVS watermark periyodik yazma zamani
uint32_t sonYazilanSent1 = 0;        // sent_i1 son yazilan deger (gereksiz NVS yazimini onler)
uint32_t sonYazilanSent2 = 0;        // sent_i2 son yazilan deger

// RSSI recovery durumu
unsigned long lastRssiKontrol = 0;
int rssiZayifSayac    = 0;   // ardisik zayif RSSI okuma
int radyoYenileSayac  = 0;   // bu zayif periyotta kac kez radyo yenilendi
int sonRSSI           = 0;   // son olculen RSSI (heartbeat'te raporlanir)
uint32_t radyoYenileToplam = 0;  // boot'tan beri toplam radyo yenileme (debug)
// Hard reset'i sinirla — gercek anten/donanim arizasinda reset loop olmasin.
// RTC bellegi ESP.restart()'i atlatir, sadece power-cycle sifirlar.
RTC_DATA_ATTR int rtcRssiHardReset = 0;
// Sunucu-bakim churn onleme: bot'tan beri hic basarili HTTP olmadan art arda kac kez
// "kuyruk bos + sessiz" reset atildi. 3+ ise reset esigi 60sn -> 30dk (3dk'da bir reboot
// dongusunu keser). Basarili HTTP'de sifirlanir. RTC: reboot'u atlatir, guc kesintisi sifirlar.
RTC_DATA_ATTR uint32_t rtcSunucuResetSayac = 0;

// ─── TANI durumu ────────────────────────────────────────────────
struct TaniOlay {
  uint32_t seq;        // cihaz ici artan olay no (boot_id ile birlikte idempotency)
  uint32_t uptime_ms;  // olay anindaki millis()
  uint8_t  tip;        // 0=SAYILDI, 2=ERKEN  (gurultu taniParazit sayacinda)
  uint32_t gap_ms;     // onceki gecerli sayimdan fark
  uint32_t low_ms;     // SAYILDI: gercek LOW (pulse) suresi — pulse bitince yazilir
};
TaniOlay taniBuf[TANI_MAX];
int taniBas = 0, taniSon = 0, taniDolu = 0;
uint32_t taniSeq = 0;
uint32_t taniSayildi = 0, taniParazit = 0, taniErken = 0;  // parazit = ham (gurultu) HIGH->LOW kenar
uint32_t bootId = 0;

// ─── Sayim kanali — integrator debounce durumu (Ist.1 / Ist.2) ───
struct Kanal {
  int           integ;          // 0..INTEG_MAX
  bool          stableLow;      // debounced "gercek LOW" durumu
  int           lastRaw;        // ham pin durumu (gurultu kenar sayimi)
  unsigned long lastValidPulse; // son gecerli sayim (MIN_PULSE_GAP)
  bool          taniBekliyor;   // pulse devam ediyor; bitince SAYILDI (low_ms ile) push edilir
  unsigned long taniLowStart;
  uint32_t      taniBekleyenGap;
};
Kanal k1 = { 0, false, HIGH, 0, false, 0, 0 };
Kanal k2 = { 0, false, HIGH, 0, false, 0, 0 };

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

// Loop icinden cagrilan BLOKLAMAYAN reconnect tetigi: denemeyi baslatir, BEKLEMEZ.
// Sonuc sonraki loop turlarinda wifiHazir() ile gorulur. Senkron bekleyen wifiBaglan()
// SADECE setup()'ta kullanilir — loop'ta 30sn senkron bekleme, kesinti boyunca
// sampling'i %66 kor birakip pulse kaybettiriyordu (buffer'a bile girmeden).
void wifiBaglanBaslat() {
  Serial.println("[WiFi] Non-blocking reconnect tetiklendi");
  WiFi.reconnect();
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

// Tani ring buffer'a olay ekle.
int taniPush(uint8_t tip, uint32_t gap_ms, uint32_t low_ms) {
  if (taniDolu >= TANI_MAX) { taniBas = (taniBas + 1) % TANI_MAX; taniDolu--; }
  int idx = taniSon;
  taniBuf[idx].seq       = ++taniSeq;
  taniBuf[idx].uptime_ms = millis();
  taniBuf[idx].tip       = tip;
  taniBuf[idx].gap_ms    = gap_ms;
  taniBuf[idx].low_ms    = low_ms;
  taniSon = (taniSon + 1) % TANI_MAX;
  taniDolu++;
  return idx;
}

bool bufferdanBirGonder() {
  if (buf_dolu == 0) return false;
  PulseKaydi p = buffer[buf_bas];

  HTTPClient http;
  String url = String(SUNUCU_HOST) + "/api/sinyal";
  http.begin(url);
  http.setTimeout(3000);   // kisa timeout — uzun blok = pulse polling gecikmesi
  http.setConnectTimeout(3000);   // TCP connect asamasi da sinirli olsun (core-default'a birakma)
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
  // bootId (her acilista rastgele) idem key'e dahil — NVS sifirlanip seq 1'den
  // baslarsa eski key'lerle CAKISMAZ (yoksa INSERT OR IGNORE yeni pulse'lari atardi).
  String idem = String(CIHAZ_ID) + "_" + mac6 + "_b" + String(bootId) + "_i" + String(p.istasyon) + "_" + String(p.seq);
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
    rtcSunucuResetSayac = 0;   // basarili temas — churn-onleme reset sayacini sifirla
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
  // Sayim yolu BLOKLANMAZ: LED blink (200ms) + in-path HTTP POST (3s'ye kadar)
  // buradan KALDIRILDI — pulse aninda integrator DONMASIN (kaynak'ta 2. kanal kacmasin).
  // Pulse buffer'da; loop'taki retry yolu (RETRY_MS'de bir) gonderir.
}

// Bloklamayan integrator debounce — her istasyon icin loop'ta her dongude cagrilir.
void kanalGuncelle(int pin, uint8_t istasyon, Kanal* k, unsigned long now) {
  int v = digitalRead(pin);

  if (k->lastRaw == HIGH && v == LOW) taniParazit++;   // ham gurultu kenar gostergesi
  k->lastRaw = v;

  if (v == LOW) { if (k->integ < INTEG_MAX) k->integ++; }
  else          { if (k->integ > 0)         k->integ--; }

  if (!k->stableLow && k->integ >= INTEG_YUKSEK_ESIK) {
    k->stableLow = true;
    unsigned long gap = now - k->lastValidPulse;
    if (gap >= MIN_PULSE_GAP_MS) {
      istasyonSinyali(istasyon);
      k->lastValidPulse  = now;
      taniSayildi++;
      k->taniBekliyor    = true;
      k->taniLowStart    = now;
      k->taniBekleyenGap = gap;
    } else {
      Serial.printf("[FILTRE] Ist.%d erken (gap=%lums < %lums) - SAYILMADI\n",
                    istasyon, gap, MIN_PULSE_GAP_MS);
      taniErken++;
      taniPush(2, gap, 0);
      k->taniBekliyor = false;
    }
  } else if (k->stableLow && k->integ <= INTEG_DUSUK_ESIK) {
    k->stableLow = false;
    if (k->taniBekliyor) {
      taniPush(0, k->taniBekleyenGap, now - k->taniLowStart);
      k->taniBekliyor = false;
    }
  }
}

// NVS watermark (2 istasyon) — her istasyon icin sunucuya ULASMIS son seq'i yaz (sent_iX).
// Buffer FIFO; gonderilmemis kayitlar buf_bas'tan itibaren. Her istasyon icin kuyruktaki
// EN DUSUK seq -> last_sent = o-1 (kuyrukta yoksa hepsi gonderildi = pulse_iX). Boot'ta
// [sent_iX+1..pulse_iX] tekrar gonderilir. Yipranma icin sadece DEGISINCE yazar.
void watermarkYaz() {
  uint32_t minUnsent1 = 0, minUnsent2 = 0;   // 0 = o istasyonda gonderilmemis yok
  int idx = buf_bas;
  for (int n = 0; n < buf_dolu; n++) {
    uint8_t  ist = buffer[idx].istasyon;
    uint32_t s   = buffer[idx].seq;
    if (ist == 1) { if (minUnsent1 == 0 || s < minUnsent1) minUnsent1 = s; }
    else if (ist == 2) { if (minUnsent2 == 0 || s < minUnsent2) minUnsent2 = s; }
    idx = (idx + 1) % BUFFER_MAX;
  }
  uint32_t sent1 = (minUnsent1 > 0) ? (minUnsent1 - 1) : pulseIst1;
  uint32_t sent2 = (minUnsent2 > 0) ? (minUnsent2 - 1) : pulseIst2;
  if (sent1 != sonYazilanSent1) { prefs.putULong("sent_i1", sent1); sonYazilanSent1 = sent1; }
  if (sent2 != sonYazilanSent2) { prefs.putULong("sent_i2", sent2); sonYazilanSent2 = sent2; }
}

void heartbeatGonder() {
  if (!wifiHazir()) return;
  HTTPClient http;
  http.begin(String(SUNUCU_HOST) + "/api/sinyal/heartbeat");
  http.setTimeout(3000);
  http.setConnectTimeout(3000);
  http.addHeader("Content-Type", "application/json");
  http.addHeader("Authorization", String("Bearer ") + API_TOKEN);

  DynamicJsonDocument doc(4096);
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
  doc["radyo_yenile"]    = radyoYenileToplam;
  doc["http_fail"]       = httpFailCount;     // sunucu post-mortem: RSSI iyi + artiyorsa zombie
  doc["disconnect"]      = disconnectCount;

  // ─── TANI: kumulatif sayaclar + son olaylar (Ist.1 + Ist.2 ortak) ───
  doc["tani_sayildi"] = taniSayildi;
  doc["tani_parazit"] = taniParazit;   // = ham gurultu kenar sayisi (iki istasyon toplami)
  doc["tani_erken"]   = taniErken;
  int taniGonder = taniDolu;
  if (taniGonder > TANI_HEARTBEAT_N) taniGonder = TANI_HEARTBEAT_N;
  if (taniGonder > 0) {
    JsonArray arr = doc.createNestedArray("tani");
    int i = taniBas;
    for (int kk = 0; kk < taniGonder; kk++) {
      JsonObject o = arr.createNestedObject();
      o["s"] = taniBuf[i].seq;  o["b"] = bootId;            o["u"] = taniBuf[i].uptime_ms;
      o["t"] = taniBuf[i].tip;  o["g"] = taniBuf[i].gap_ms; o["l"] = taniBuf[i].low_ms;
      i = (i + 1) % TANI_MAX;
    }
  }

  String payload; serializeJson(doc, payload);
  int rc = http.POST(payload);
  http.end();
  if (rc == 200 || rc == 201) {
    httpFailCount = 0;
    lastSuccessHttp = millis();
    rtcSunucuResetSayac = 0;   // basarili temas — churn-onleme reset sayacini sifirla
    taniBas = (taniBas + taniGonder) % TANI_MAX;   // gonderilenleri dusur
    taniDolu -= taniGonder;
  } else {
    httpFailCount++;
  }
  Serial.printf("[HEART] HTTP %d · RSSI=%d · kuyruk=%d · ist1=%lu · ist2=%lu · robot=%s · tani(S%lu/P%lu/E%lu) · fail=%d\n",
                rc, WiFi.RSSI(), buf_dolu, pulseIst1, pulseIst2,
                robotCalisiyor ? "ON" : "OFF", taniSayildi, taniParazit, taniErken, httpFailCount);
}

// ════════════════════════════════════════════════════════════
//   SETUP / LOOP
// ════════════════════════════════════════════════════════════

void setup() {
  Serial.begin(115200);
  delay(300);
  bootMs = millis();
  // bootId, NVS acildiktan SONRA atanir (prefs.begin asagida) — reboot'lar arasi STABIL
  // olmali ki watermark resend'i ayni idempotency_key'i uretsin (cift sayim yok).

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

  k1.lastRaw     = digitalRead(PIN_IST1_SAYAC);
  k2.lastRaw     = digitalRead(PIN_IST2_SAYAC);
  lastRobotState = digitalRead(PIN_ROBOT_CALISIYOR);
  robotCalisiyor = (lastRobotState == LOW);

  prefs.begin("cofle", false);
  pulseIst1 = prefs.getULong("pulse_i1", 0);
  pulseIst2 = prefs.getULong("pulse_i2", 0);
  Serial.printf("[NVS] Kayitli sayaclar: Ist1=%lu · Ist2=%lu\n", pulseIst1, pulseIst2);

  // bootId STABIL (reboot'lar arasi ayni). NVS silinirse boot_id de gider -> yeni bootId +
  // sayim 0'dan baslar, eski idempotency key'leriyle CAKISMAZ. (Eski kod her boot'ta
  // esp_random uretiyordu; o, reboot resend'inde ayni pulse'a yeni key verip CIFT saydirirdi.)
  bootId = prefs.getULong("boot_id", 0);
  if (bootId == 0) {
    bootId = esp_random();
    prefs.putULong("boot_id", bootId);
  }

  // Reboot resend (2 istasyon): sunucuya ulasmis son seq = sent_iX. Gonderilememis
  // [sent_iX+1 .. pulse_iX] araligini kuyruga geri koy. Eski firmware'de sent_iX yok ->
  // default pulse_iX (yanlis resend yok). Stabil bootId -> sunucu cift saymaz.
  uint32_t sent1 = prefs.getULong("sent_i1", pulseIst1);
  uint32_t sent2 = prefs.getULong("sent_i2", pulseIst2);
  if (sent1 > pulseIst1) sent1 = pulseIst1;
  if (sent2 > pulseIst2) sent2 = pulseIst2;
  sonYazilanSent1 = sent1;
  sonYazilanSent2 = sent2;
  // Yardimci lambda yerine acik dongu (Arduino .ino oto-prototip uyumu)
  for (uint8_t ist = 1; ist <= 2; ist++) {
    uint32_t toplam = (ist == 1) ? pulseIst1 : pulseIst2;
    uint32_t sent   = (ist == 1) ? sent1     : sent2;
    if (toplam <= sent) continue;
    uint32_t resendBas = sent + 1;
    uint32_t adet      = toplam - sent;
    if ((int)(buf_dolu + adet) > BUFFER_MAX) {       // tavan: en yeni seq'leri tut
      uint32_t tasma = (buf_dolu + adet) - BUFFER_MAX;
      if (tasma >= adet) continue;                   // bu istasyon icin yer kalmadi
      resendBas += tasma;
      adet -= tasma;
    }
    for (uint32_t s = resendBas; s <= toplam; s++) {
      buffer[buf_son].seq      = s;
      buffer[buf_son].istasyon = ist;
      buf_son = (buf_son + 1) % BUFFER_MAX;
      buf_dolu++;
    }
    Serial.printf("[NVS] Ist.%d gonderilememis %lu pulse kuyruga kondu (seq %lu..%lu) — reboot resend\n",
                  ist, (unsigned long)adet, (unsigned long)resendBas, (unsigned long)toplam);
  }

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
  ArduinoOTA.onStart([]() {
    // OTA transferi loop'a donmeden dakikalarca surebilir — 30sn task-WDT
    // yukleme ortasinda panik reset atmasin diye loopTask'i izlemeden cikar.
    esp_task_wdt_delete(NULL);
    Serial.println("\n[OTA] Guncelleme basliyor (WDT askida)");
  });
  ArduinoOTA.onEnd([]()   { Serial.println("\n[OTA] Tamam, yeniden baslatiliyor"); });
  ArduinoOTA.onProgress([](unsigned int p, unsigned int t) {
    Serial.printf("[OTA] %u%%\r", (p / (t / 100)));
  });
  ArduinoOTA.onError([](ota_error_t e) {
    Serial.printf("[OTA] HATA %u\n", e);
    esp_task_wdt_add(NULL);   // OTA basarisiz — normal calisma surecek, WDT korumasini geri tak
  });
  ArduinoOTA.begin();
  Serial.printf("[OTA] Aktif — IDE Network Port: %s @ %s\n",
                CIHAZ_ID, WiFi.localIP().toString().c_str());

  heartbeatGonder();
  lastHeartbeat = millis();
  lastRetry = millis();

  Serial.println("[READY] Integrator debounce + TANI aktif — robot DO sinyalleri bekleniyor...");
  Serial.printf( "        (YUKSEK=%d/MAX=%d/DUSUK=%d · MIN_PULSE_GAP=%lums)\n\n",
                 INTEG_YUKSEK_ESIK, INTEG_MAX, INTEG_DUSUK_ESIK, MIN_PULSE_GAP_MS);
  ledYakBlink(3, 60);
}

void loop() {
  esp_task_wdt_reset();
  ArduinoOTA.handle();
  unsigned long now = millis();

  if (!wifiHazir()) {
    if (wifiKoptuMs == 0) wifiKoptuMs = now;   // kesinti baslangicini isaretle
    digitalWrite(PIN_LED, (now / 200) % 2);
    if ((now - lastHeartbeat) > 15000) {
      wifiBaglanBaslat();     // BLOKLAMAZ — sampling kesintide de devam eder
      lastHeartbeat = now;
    }
  } else {
    wifiKoptuMs = 0;
  }

  // Ist.1 + Ist.2 — bloklamayan integrator debounce (gurultuye dayanikli)
  kanalGuncelle(PIN_IST1_SAYAC, 1, &k1, now);
  kanalGuncelle(PIN_IST2_SAYAC, 2, &k2, now);

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
      // Throttle: role flap'inde (her cevrimde ac/kapa) her degisim 3-6sn'lik senkron
      // POST olmasin — min 5sn araliksa gonder; degilse 30sn periyodik heartbeat tasir.
      if (wifiHazir() && (now - lastHeartbeat) > 5000UL) {
        heartbeatGonder();
        lastHeartbeat = now;
      }
    }
  }

  // Retry: basarisiz gonderimde ustel geri cekilme (3s -> 60s cap). Sunucu kapaliyken
  // her 3sn'de 3-6sn bloklanip sampling'i %50 kor birakmayi engeller; basarida 3sn'e doner.
  if (wifiHazir() && buf_dolu > 0 && (now - lastRetry) > retryGecikmeMs) {
    if (bufferdanBirGonder()) {
      retryGecikmeMs = RETRY_MS;
    } else {
      retryGecikmeMs = (retryGecikmeMs * 2 > 60000UL) ? 60000UL : retryGecikmeMs * 2;
    }
    lastRetry = now;
  }

  // Heartbeat: sunucu cevap vermiyorsa araligi 4x'e cikar (30s->120s) — kor pencereyi kucult
  {
    unsigned long hbAralik = (httpFailCount >= 3) ? (HEARTBEAT_MS * 4) : HEARTBEAT_MS;
    if (wifiHazir() && (now - lastHeartbeat) > hbAralik) {
      heartbeatGonder();
      lastHeartbeat = now;
    }
  }

  // NVS watermark periyodik yazimi (WiFi'den BAGIMSIZ) — beklenmeyen reset/guc kesintisinde
  // tekrar-gonderim penceresini ~30s ile sinirlar. Sadece deger degisince yazar (yipranma).
  if ((now - lastWatermark) > WATERMARK_YAZ_MS) {
    watermarkYaz();
    lastWatermark = now;
  }

  // ─── SELF-HEAL kontrolleri ───
  // 1) 24 saat uptime → koruyucu reset, AMA makine bos iken (uretim kesintiye ugramasin)
  //    Son pulse'tan beri 2 dk gectiyse (iki istasyon da) reset guvenli.
  if ((now - bootMs) > UPTIME_RESET_MS) {
    bool makineBos = (now - k1.lastValidPulse) > 120000UL
                  && (now - k2.lastValidPulse) > 120000UL;
    if (makineBos && buf_dolu == 0) {
      Serial.println("\n[SELFHEAL] 24h doldu + makine bos + kuyruk bos — koruyucu reset");
      delay(500);
      ESP.restart();
    }
    // Makine calisiyorsa reset ertelenir — bir sonraki bos doneme birakilir
  }
  // 2) Sunucuya ulasilamiyor — ZOMBIE WL_CONNECTED dahil (RSSI iyi gorunur, disconnect
  //    event'i ATISLANMAZ; bu yuzden lastSuccessHttp tabanli BAGIMSIZ kurtarma sart).
  if (httpFailCount >= 5) {
    unsigned long sonTemas = (lastSuccessHttp > 0) ? lastSuccessHttp : bootMs;
    unsigned long sessizMs = now - sonTemas;
    if (buf_dolu == 0) {
      // Kuyruk bos: kaybedecek veri yok -> hizli reset. CHURN ONLEME: sunucu uzun
      // bakimda ise art arda 3 basarisiz reset'ten sonra esik 60sn -> 30dk cikar
      // (yoksa cihaz bakim boyunca ~3dk'da bir reboot dongusune girer).
      unsigned long resetEsik = (rtcSunucuResetSayac >= 3) ? 1800000UL : 60000UL;
      if (sessizMs > resetEsik) {
        rtcSunucuResetSayac++;
        Serial.printf("\n[SELFHEAL] %d HTTP fail + kuyruk bos + %lus sessiz — RESET (#%lu)\n",
                      httpFailCount, sessizMs / 1000, (unsigned long)rtcSunucuResetSayac);
        delay(500);
        ESP.restart();
      }
    } else if (sessizMs > SUNUCU_SESSIZ_HARD_RESET_MS) {
      // Kuyruk DOLU ama 5dk'dir sunucuya HIC ulasilamiyor (zombie/upstream) -> watermark
      // (sent_i1/sent_i2) yaz, GUVENLE reboot et. Boot'ta gonderilmemis araliklar tekrar
      // gonderilir (stabil bootId -> cift saymaz). Eski "kuyruk varken reboot yok" livelock'u biter.
      watermarkYaz();
      Serial.printf("\n[SELFHEAL] Sunucuya %lus ulasilamadi (zombie?) — watermark(kuyruk=%d) + REBOOT\n",
                    sessizMs / 1000, buf_dolu);
      delay(500);
      ESP.restart();
    } else if ((now - lastRadyoYenile) > 60000UL) {
      // Once YUMUSAK kurtarma: radyo tazele (gercek RF/L2 hang'i reset'siz toparlar).
      // SADECE MAKINE BOSKEN: wifiRadyoYenile ~16sn bloklar, uretim suruyorken pulse
      // kaybettirir (RSSI dalindaki kuralin aynisi). Uretim varsa atla — 5dk
      // watermark+reboot yolu zaten guvenli kurtarmayi saglar.
      bool makineBosZ = (now - k1.lastValidPulse) > 120000UL
                     && (now - k2.lastValidPulse) > 120000UL;
      if (makineBosZ) {
        Serial.printf("\n[SELFHEAL] %d HTTP fail, %lus sessiz, kuyruk=%d — radyo tazele\n",
                      httpFailCount, sessizMs / 1000, buf_dolu);
        wifiRadyoYenile();
        lastRadyoYenile = now;
      }
    }
  }
  // 2b) WiFi 5dk'dir HIC baglanamiyor (AP kayip / assoc wedge) — httpFailCount'tan
  //     BAGIMSIZ kurtarma. POST hic denenemedigi icin fail sayaci artmaz ve 2) dali
  //     ASLA tetiklenmez (ABB4 tipi RF-katmani kilidinin kor noktasi). Watermark NVS'te
  //     oldugu icin kuyruk doluyken bile reboot GUVENLI (stabil bootId -> cift sayim yok).
  if (wifiKoptuMs != 0 && (now - wifiKoptuMs) > SUNUCU_SESSIZ_HARD_RESET_MS) {
    watermarkYaz();
    Serial.printf("\n[SELFHEAL] WiFi %lus'dir baglanamiyor — watermark(kuyruk=%d) + REBOOT\n",
                  (now - wifiKoptuMs) / 1000, buf_dolu);
    delay(500);
    ESP.restart();
  }
  // 3) 10dk icinde 3+ disconnect → reset (WiFi yigini bozulmus)
  if (disconnectCount >= 3 && (now - lastDisconnectMs) < 600000UL && buf_dolu == 0) {
    // Kuyrukta pulse varken reboot etme — loop zaten wifiBaglan ile reconnect dener
    Serial.printf("\n[SELFHEAL] 10dk icinde %d disconnect + kuyruk bos — RESET\n",
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
    bool makineBosRssi = (now - k1.lastValidPulse) > 120000UL
                      && (now - k2.lastValidPulse) > 120000UL;
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
          if (rtcRssiHardReset < 2 && buf_dolu == 0) {
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
