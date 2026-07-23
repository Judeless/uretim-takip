/* ============================================================
 *  COFLE PILOT SAYAC — PLASTIK-320T firmware (v2.1)
 *  >>> OTOMATIK URETILDI: generate.py — manuel duzenleme!
 * ============================================================
 *
 *  HEDEF:  320T plastik enjeksiyon makinesinin uretim
 *          cevrim sinyalini izler. Her parca/cevrimde 1 pulse.
 *          (Calisma mantigi metal enjeksiyon ile ayni; bolum "plastik", TK1.)
 *
 *  PIN ATAMALARI:
 *   - GPIO25  -> Role NO  (Uretim/cevrim sinyali — kuru kontak)
 *   - GPIO26  -> BOS (kullanilmiyor)
 *   - GPIO27  -> BOS (kullanilmiyor)
 *   - GND     -> Role COM
 *
 *  UYARI: Sinyal KURU KONTAK olmali. 24V ise araya optokuplor (ESP32 3.3V).
 *  AYAR: INTEG esikleri + MIN_PULSE_GAP makinenin sinyal sekline gore sahada
 *        ince ayar isteyebilir (metal 400T'de oldugu gibi — asagi bak).
 *
 *  Bu dosya pilot/firmware/_templates/plastik.ino.tpl'den uretildi.
 * ============================================================ */

#include <WiFi.h>
#include <HTTPClient.h>
#include <Preferences.h>
#include <ArduinoJson.h>
#include <esp_task_wdt.h>
#include <ArduinoOTA.h>

// ════════════════════════════════════════════════════════════
//   YAPILANDIRMA — 320T
// ════════════════════════════════════════════════════════════

const char* CIHAZ_ID  = "PLASTIK-320T";
const char* BOLUM     = "plastik";
const char* ROBOT_NO  = "320T";

const char* WIFI_SSID = "COFLE-TK";
const char* WIFI_PASS = "internet2011!";
const char* SUNUCU_HOST = "http://192.168.21.155:5001";
const char* API_TOKEN = "cofle-pilot-2026";
const char* OTA_PASS = "cofle-ota-2026";

const int PIN_IST1_SAYAC      = 25;   // Uretim sinyali
const int PIN_IST2_SAYAC      = 26;   // BOS
const int PIN_ROBOT_CALISIYOR = 27;   // BOS
const int PIN_LED             =  2;

// Pulse algilama: BLOKLAMAYAN INTEGRATOR debounce (v2.5, metal'e ozel).
// 400T gercek cevrim = "makine calisiyor" rolesi 20-26sn LOW. Hat AGIR gurultulu:
// kisa blipler (<=225ms) + ARA ARA 1-2.4sn'lik SUREKLI LOW patlamalari da var.
//   her dongude pini ornekle -> LOW'da +1 / HIGH'da -1 (0..INTEG_MAX).
//   integ >= YUKSEK_ESIK olunca "gercek LOW basladi" (say); histerezisle DUSUK_ESIK'te biter.
// v2.5.0 (~400ms esik) saha testi: 27 baski -> 36 sayim (1-2s patlamalar sayiliyordu = FAZLA).
// v2.5.1: esik ~4sn'ye cikarildi -> 1-2.4s patlamalari eler, 20-26s gercek cevrim genis marjla gecer.
// (~5ms dongu periyodu: 800 ornek ~= 4sn net LOW.) NOT: bu yuksek esik metal'e ozeldir cunku
// gercek pulse cok uzun; montaj/kaynak'ta gercek pulse ~75ms oldugu icin orada DUSUK tutulur.
const int  INTEG_YUKSEK_ESIK   = 800;   // ~4sn net LOW -> pulse basladi (1-2.4s gurultu patlamalarini eler)
const int  INTEG_MAX           = 1200;  // tavan (~6sn) — pulse ici uzun HIGH glitch tamponu
const int  INTEG_DUSUK_ESIK    = 80;    // ~0.4sn surekli HIGH -> pulse bitti; LOW ici ~5.6sn'ye kadar HIGH patlamasini tolere eder
const unsigned long MIN_PULSE_GAP_MS = 18000; // iki gecerli sayim arasi min (makine cycle siniri)
const int  HEARTBEAT_MS   = 30000;
const int  RETRY_MS       = 3000;
const int  WIFI_TIMEOUT_S = 30;
const int  WDT_TIMEOUT_S  = 30;
const int  BUFFER_MAX     = 2000;  // kesinti kuyrugu (eskiden 200) — elektrik varken gunlerce pulse tutar
const char* FIRMWARE_VER  = "2.6.2-320t";   // 2.6.2: WDT+OTA fix, non-blocking reconnect, retry backoff, WiFi-down hard reset, churn onleme

// ─── TANI (diagnostic) — sayim filtresi kararlarini loglar ──────
// Sayma mantigini DEGISTIRMEZ; sadece her karari gozlemleyip heartbeat ile gonderir.
// Amac: eksik/fazla sayma sebebini sunucudan gormek (parazit=opto, erken=cycle/gap,
// SAYILDI low_ms ~2x = role ardisik parcalari birlestiriyor).
const int TANI_MAX          = 30;   // RAM ring buffer (2+ heartbeat'lik olay)
const int TANI_HEARTBEAT_N  = 20;   // tek heartbeat'te en fazla kac olay gonderilsin

// ─── RSSI tabanli radyo recovery (v2.3) ─────────────────────────
const int           RSSI_ZAYIF_ESIK  = -80;
const unsigned long RSSI_KONTROL_MS  = 60000;
const int           RSSI_ZAYIF_MAX   = 5;
const int           RADYO_YENILE_MAX = 3;

// ─── Zombie WL_CONNECTED kurtarma + NVS watermark (v2.6) ────────────
// Cihaz AP'ye bagli gorunup (RSSI iyi okunur, disconnect event'i ATISLANMAZ) ama
// L3/uygulama yolu (IP/soket/upstream) kopabilir -> heartbeat POST'lari sessizce fail,
// sunucu offline gosterir, SADECE manuel reset duzeltir. Eski self-heal kuyrukta pulse
// varken REBOOT ETMIYORDU -> sonsuz radyo-tazele livelock. Cozum: kuyruk dolu olsa bile
// 5dk'dir sunucuya ulasilamiyorsa watermark'i NVS'e yazip GUVENLE reboot et (boot'ta
// [sent_i1+1..pulse_i1] tekrar gonderilir; bootId NVS'te stabil -> sunucu cift saymaz).
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

unsigned long lastHeartbeat = 0;
unsigned long lastRetry     = 0;
unsigned long bootMs        = 0;

// Self-heal sayaclari — Smart Counter'in donanim watchdog'unu yazilim ile taklit
int httpFailCount = 0;
unsigned long lastSuccessHttp = 0;
// volatile: WiFi event task'i yazar, loop okur (32-bit hizali erisim — kritik bolge gerekmez)
volatile int disconnectCount = 0;
volatile unsigned long lastDisconnectMs = 0;
const unsigned long UPTIME_RESET_MS = 24UL * 3600UL * 1000UL;
unsigned long wifiKoptuMs = 0;           // wifiHazir() false'a dustugu an (0=bagli) — WiFi-down hard reset icin
unsigned long retryGecikmeMs = RETRY_MS; // basarisiz POST'ta ustel artar (cap 60s) — sunucu kapaliyken sampling'i korur

// Zombie kurtarma + watermark durumu (v2.6)
unsigned long lastRadyoYenile = 0;   // zombie kurtarma radyo-tazele throttle (lastSuccessHttp'yi EZMEDEN)
unsigned long lastWatermark   = 0;   // NVS watermark periyodik yazma zamani
uint32_t sonYazilanSent       = 0;   // sent_i1 son yazilan deger — gereksiz NVS yazimini onler

// RSSI recovery durumu
unsigned long lastRssiKontrol = 0;
int rssiZayifSayac    = 0;
int radyoYenileSayac  = 0;
int sonRSSI           = 0;
uint32_t radyoYenileToplam = 0;
RTC_DATA_ATTR int rtcRssiHardReset = 0;
// Sunucu-bakim churn onleme: art arda basarisiz "kuyruk bos + sessiz" reset sayisi.
// 3+ ise reset esigi 60sn -> 30dk. Basarili HTTP'de sifirlanir. RTC: reboot'u atlatir.
RTC_DATA_ATTR uint32_t rtcSunucuResetSayac = 0;

// ─── TANI durumu ────────────────────────────────────────────────
struct TaniOlay {
  uint32_t seq;        // cihaz ici artan olay no (boot_id ile birlikte idempotency)
  uint32_t uptime_ms;  // olay anindaki millis()
  uint8_t  tip;        // 0=SAYILDI, 2=ERKEN  (gurultu olay olarak degil, taniParazit sayacinda)
  uint32_t gap_ms;     // onceki gecerli sayimdan fark
  uint32_t low_ms;     // SAYILDI: gercek LOW suresi (pulse bitince yazilir)
};
TaniOlay taniBuf[TANI_MAX];
int taniBas = 0, taniSon = 0, taniDolu = 0;
uint32_t taniSeq = 0;
uint32_t taniSayildi = 0, taniParazit = 0, taniErken = 0;  // parazit = ham (gurultu) HIGH->LOW kenar
uint32_t bootId = 0;            // her acilista rastgele — reboot sonrasi seq cakismasini onler

// ─── Sayim kanali — integrator debounce durumu (per-istasyon) ───
struct Kanal {
  int           integ;          // 0..INTEG_MAX
  bool          stableLow;      // debounced "gercek LOW" durumu
  int           lastRaw;        // ham pin durumu (gurultu kenar sayimi icin)
  unsigned long lastValidPulse; // son gecerli sayim (MIN_PULSE_GAP icin)
  bool          taniBekliyor;   // pulse devam ediyor; bitince SAYILDI olayi (low_ms ile) push edilecek
  unsigned long taniLowStart;   // mevcut LOW baslangici
  uint32_t      taniBekleyenGap;// o pulse'in gap'i
};
Kanal k1 = { 0, false, HIGH, 0, false, 0, 0 };

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

// Loop icinden cagrilan BLOKLAMAYAN reconnect tetigi: denemeyi baslatir, BEKLEMEZ.
// Sonuc sonraki loop turlarinda wifiHazir() ile gorulur. Senkron bekleyen wifiBaglan()
// SADECE setup()'ta kullanilir — loop'ta 30sn senkron bekleme, kesinti boyunca
// sampling'i %66 kor birakip uretim sinyallerini kaybettiriyordu (buffer'a bile girmeden).
void wifiBaglanBaslat() {
  Serial.println("[WiFi] Non-blocking reconnect tetiklendi");
  WiFi.reconnect();
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

// Tani ring buffer'a olay ekle.
int taniPush(uint8_t tip, uint32_t gap_ms, uint32_t low_ms) {
  if (taniDolu >= TANI_MAX) {           // dolu → en eskiyi dusur (kumulatif sayaclar yine korur)
    taniBas = (taniBas + 1) % TANI_MAX;
    taniDolu--;
  }
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
  http.setTimeout(3000);
  http.setConnectTimeout(3000);   // TCP connect asamasi da sinirli (core-default'a birakma)
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
  // bootId (her acilista rastgele) idem key'e dahil — NVS sifirlanip seq 1'den
  // baslarsa eski key'lerle CAKISMAZ (yoksa INSERT OR IGNORE yeni pulse'lari atardi).
  String idem = String(CIHAZ_ID) + "_" + mac6 + "_b" + String(bootId) + "_i" + String(p.istasyon) + "_" + String(p.seq);
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
    rtcSunucuResetSayac = 0;   // basarili temas — churn-onleme reset sayacini sifirla
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
  // Sayim yolu BLOKLANMAZ: LED blink (200ms) + in-path HTTP POST (3s'ye kadar)
  // buradan KALDIRILDI — pulse aninda integrator DONMASIN (sayim kacmasin).
  // Pulse buffer'da; loop'taki retry yolu (RETRY_MS'de bir) gonderir.
}

// Bloklamayan integrator debounce — loop'ta her dongude bir kez cagrilir.
// Kisa gurultu bliplerini emer, gercek (uzun) LOW'u sayar; tek glitch pulse'i oldurmez.
void kanalGuncelle(int pin, uint8_t istasyon, Kanal* k, unsigned long now) {
  int v = digitalRead(pin);

  // Ham HIGH->LOW kenar = gurultu gostergesi (entegratore takilip elenen blipler dahil)
  if (k->lastRaw == HIGH && v == LOW) taniParazit++;
  k->lastRaw = v;

  // Integrator: LOW'da +1, HIGH'da -1 (0..INTEG_MAX)
  if (v == LOW) { if (k->integ < INTEG_MAX) k->integ++; }
  else          { if (k->integ > 0)         k->integ--; }

  if (!k->stableLow && k->integ >= INTEG_YUKSEK_ESIK) {
    k->stableLow = true;                         // gercek LOW basladi
    unsigned long gap = now - k->lastValidPulse;
    if (gap >= MIN_PULSE_GAP_MS) {
      istasyonSinyali(istasyon);                 // SAY (gercek zamanli)
      k->lastValidPulse  = now;
      taniSayildi++;
      k->taniBekliyor    = true;                 // SAYILDI tani olayi pulse bitince (low_ms hazir) yazilir
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
    k->stableLow = false;                         // gercek LOW bitti
    if (k->taniBekliyor) {
      taniPush(0, k->taniBekleyenGap, now - k->taniLowStart);
      k->taniBekliyor = false;
    }
  }
}

// NVS watermark — sunucuya ULASMIS son seq'i (sent_i1) yaz. Boylece zombie/crash reboot'unda
// kuyrukta kalan pulse'lar boot'ta tekrar gonderilir (bkz. setup). Sayim (pulse_i1) zaten her
// basista yaziliyor; bu sadece "nereye kadar teyitlendi". Yipranma icin yalniz DEGISINCE yazar.
void watermarkYaz() {
  uint32_t gonderilen = (buf_dolu > 0) ? (buffer[buf_bas].seq - 1) : pulseIst1;
  if (gonderilen != sonYazilanSent) {
    prefs.putULong("sent_i1", gonderilen);
    sonYazilanSent = gonderilen;
  }
}

void heartbeatGonder() {
  if (!wifiHazir()) return;
  HTTPClient http;
  http.begin(String(SUNUCU_HOST) + "/api/sinyal/heartbeat");
  http.setTimeout(3000);
  http.setConnectTimeout(3000);   // TCP connect asamasi da sinirli (core-default'a birakma)
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
  doc["radyo_yenile"]    = radyoYenileToplam;
  doc["http_fail"]       = httpFailCount;     // sunucu post-mortem: RSSI iyi + artiyorsa zombie
  doc["disconnect"]      = disconnectCount;
  doc["robot_calisiyor"] = robotCalisiyor;

  // ─── TANI: kumulatif sayaclar + son olaylar (hepsi tamamlanmis) ───
  doc["tani_sayildi"] = taniSayildi;
  doc["tani_parazit"] = taniParazit;   // = ham gurultu kenar sayisi
  doc["tani_erken"]   = taniErken;
  int taniGonder = taniDolu;
  if (taniGonder > TANI_HEARTBEAT_N) taniGonder = TANI_HEARTBEAT_N;
  if (taniGonder > 0) {
    JsonArray arr = doc.createNestedArray("tani");
    int i = taniBas;
    for (int k = 0; k < taniGonder; k++) {
      JsonObject o = arr.createNestedObject();
      o["s"] = taniBuf[i].seq;
      o["b"] = bootId;
      o["u"] = taniBuf[i].uptime_ms;
      o["t"] = taniBuf[i].tip;
      o["g"] = taniBuf[i].gap_ms;
      o["l"] = taniBuf[i].low_ms;
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
    // Gonderilen tani olaylarini buffer'dan dusur (basarisizsa kalir, tekrar denenir)
    taniBas = (taniBas + taniGonder) % TANI_MAX;
    taniDolu -= taniGonder;
  } else {
    httpFailCount++;
  }
  Serial.printf("[HEART] HTTP %d · RSSI=%d · kuyruk=%d · sayim=%lu · tani(S%lu/P%lu/E%lu) · fail=%d\n",
                rc, WiFi.RSSI(), buf_dolu, pulseIst1,
                taniSayildi, taniParazit, taniErken, httpFailCount);
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
  Serial.printf("FW: %s · Bolum: %s · Makine: %s · PLASTIK ENJEKSIYON\n",
                FIRMWARE_VER, BOLUM, ROBOT_NO);
  Serial.printf("Uretim pini: GPIO%d (makine her parcada pulse)\n", PIN_IST1_SAYAC);
  Serial.printf("Min pulse araligi: %lums (makine cycle siniri)\n", MIN_PULSE_GAP_MS);
  Serial.printf("MAC: %s\n", WiFi.macAddress().c_str());

  pinMode(PIN_IST1_SAYAC,      INPUT_PULLUP);
  pinMode(PIN_IST2_SAYAC,      INPUT_PULLUP);
  pinMode(PIN_ROBOT_CALISIYOR, INPUT_PULLUP);
  pinMode(PIN_LED,             OUTPUT);
  digitalWrite(PIN_LED, LOW);

  k1.lastRaw     = digitalRead(PIN_IST1_SAYAC);
  robotCalisiyor = false;  // Plastik'te robot durumu kullanilmiyor

  prefs.begin("cofle", false);
  pulseIst1 = prefs.getULong("pulse_i1", 0);
  pulseIst2 = prefs.getULong("pulse_i2", 0);
  Serial.printf("[NVS] Kayitli sayim: %lu\n", pulseIst1);

  // bootId STABIL (reboot'lar arasi ayni). NVS silinirse boot_id de gider -> yeni bootId +
  // sayim 0'dan baslar, eski idempotency key'leriyle CAKISMAZ. (Eski kod her boot'ta
  // esp_random uretiyordu; o, reboot resend'inde ayni pulse'a yeni key verip CIFT saydirirdi.)
  bootId = prefs.getULong("boot_id", 0);
  if (bootId == 0) {
    bootId = esp_random();
    prefs.putULong("boot_id", bootId);
  }

  // Reboot resend: sunucuya ulasmis son seq = sent_i1. Gonderilememis [sent_i1+1 .. pulse_i1]
  // araligini kuyruga geri koy. Eski firmware'de sent_i1 yok -> default pulse_i1 (yanlis resend yok).
  uint32_t sentSeq = prefs.getULong("sent_i1", pulseIst1);
  if (sentSeq > pulseIst1) sentSeq = pulseIst1;   // tutarsizlik guvenligi
  sonYazilanSent = sentSeq;
  if (pulseIst1 > sentSeq) {
    uint32_t resendBas = sentSeq + 1;
    uint32_t adet      = pulseIst1 - sentSeq;
    if (adet > BUFFER_MAX) {                       // en yeni BUFFER_MAX kadarini tut (eskiyi dusur)
      resendBas = pulseIst1 - BUFFER_MAX + 1;
      adet = BUFFER_MAX;
    }
    for (uint32_t s = resendBas; s <= pulseIst1; s++) {
      buffer[buf_son].seq      = s;
      buffer[buf_son].istasyon = 1;
      buf_son = (buf_son + 1) % BUFFER_MAX;
      buf_dolu++;
    }
    Serial.printf("[NVS] Gonderilememis %lu pulse kuyruga kondu (seq %lu..%lu) — reboot resend\n",
                  (unsigned long)adet, (unsigned long)resendBas, (unsigned long)pulseIst1);
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

  Serial.println("[READY] Integrator debounce + TANI aktif — uretim sinyalleri bekleniyor...");
  Serial.printf( "        (gurultuye dayanikli: YUKSEK=%d/MAX=%d/DUSUK=%d · MIN_PULSE_GAP=%lums)\n\n",
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

  // Uretim sinyali — bloklamayan integrator debounce (gurultuye dayanikli)
  kanalGuncelle(PIN_IST1_SAYAC, 1, &k1, now);

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
  // 1) 24 saat uptime → koruyucu reset, AMA makine bos iken (uretim kesilmesin)
  if ((now - bootMs) > UPTIME_RESET_MS) {
    if ((now - k1.lastValidPulse) > 120000UL && buf_dolu == 0) {
      Serial.println("\n[SELFHEAL] 24h doldu + makine bos + kuyruk bos — koruyucu reset");
      delay(500);
      ESP.restart();
    }
  }
  // 2) Sunucuya ulasilamiyor — ZOMBIE WL_CONNECTED dahil (RSSI iyi gorunur, disconnect
  //    event'i ATISLANMAZ; bu yuzden lastSuccessHttp tabanli BAGIMSIZ kurtarma sart).
  if (httpFailCount >= 5) {
    unsigned long sonTemas = (lastSuccessHttp > 0) ? lastSuccessHttp : bootMs;
    unsigned long sessizMs = now - sonTemas;
    if (buf_dolu == 0) {
      // Kuyruk bos: kaybedecek veri yok -> hizli reset. CHURN ONLEME: sunucu uzun
      // bakimda ise art arda 3 basarisiz reset'ten sonra esik 60sn -> 30dk cikar.
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
      // yaz, GUVENLE reboot et. Boot'ta [sent_i1+1..pulse_i1] tekrar gonderilir (stabil
      // bootId -> sunucu cift saymaz). Eski "kuyruk varken hic reboot etme" livelock'u biter.
      watermarkYaz();
      Serial.printf("\n[SELFHEAL] Sunucuya %lus ulasilamadi (zombie?) — watermark(kuyruk=%d) + REBOOT\n",
                    sessizMs / 1000, buf_dolu);
      delay(500);
      ESP.restart();
    } else if ((now - lastRadyoYenile) > 60000UL) {
      // Once YUMUSAK kurtarma: radyo tazele. SADECE MAKINE BOSKEN — wifiRadyoYenile
      // ~16sn bloklar, uretim suruyorken sinyal kaybettirir (RSSI dalinin kurali).
      // Uretim varsa atla — 5dk watermark+reboot yolu zaten guvenli kurtarma.
      bool makineBosZ = (now - k1.lastValidPulse) > 120000UL;
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
  if (disconnectCount >= 3 && (now - lastDisconnectMs) < 600000UL && buf_dolu == 0) {
    // Kuyrukta pulse varken reboot etme — loop zaten wifiBaglan ile reconnect dener
    Serial.printf("\n[SELFHEAL] 10dk icinde %d disconnect + kuyruk bos — RESET\n",
                  disconnectCount);
    delay(500);
    ESP.restart();
  }
  // 4) RSSI tabanli radyo recovery — sadece makine BOS iken (blok pulse kaybetmesin)
  if (wifiHazir() && (now - lastRssiKontrol) > RSSI_KONTROL_MS) {
    lastRssiKontrol = now;
    sonRSSI = WiFi.RSSI();
    bool makineBosRssi = (now - k1.lastValidPulse) > 120000UL;
    if (sonRSSI < RSSI_ZAYIF_ESIK && sonRSSI < 0) {
      rssiZayifSayac++;
      Serial.printf("[RSSI] Zayif sinyal %d dBm (%d/%d)%s\n", sonRSSI, rssiZayifSayac, RSSI_ZAYIF_MAX,
                    makineBosRssi ? "" : " — makine calisiyor, recovery ertelendi");
      if (rssiZayifSayac >= RSSI_ZAYIF_MAX && makineBosRssi) {
        rssiZayifSayac = 0;
        radyoYenileSayac++;
        wifiRadyoYenile();
        if (radyoYenileSayac >= RADYO_YENILE_MAX) {
          radyoYenileSayac = 0;
          if (rtcRssiHardReset < 2 && buf_dolu == 0) {
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
