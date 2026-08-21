/* ============================================================
 *  COFLE PILOT SAYAC — TEL-KAPAMA-4 firmware (3 KANALLI)
 *  >>> OTOMATIK URETILDI: generate.py — manuel duzenleme!
 * ============================================================
 *
 *  HEDEF:  TK1 kapama hattinda ORTAK PANOYU paylasan 3 presin
 *          rolelerini TEK modulle izler. Her pres kendi kanalinda
 *          kendi makine adiyla (robot_no) sunucuya yazilir:
 *
 *              GPIO25 -> Kapama 10      (istasyon 1)
 *              GPIO26 -> Kapama 11      (istasyon 2)
 *              GPIO27 -> Kapama 12      (istasyon 3)
 *
 *          Bolum "tel" (kapama, tel uretiminin proses adimidir).
 *          Operator mobilde hangi presi seciyorsa (vardiya hatti)
 *          sayac dogrudan o kanalla eslesir — ek eslesme tablosu yok.
 *
 *  PIN ATAMALARI:
 *   - GPIO25 / GPIO26 / GPIO27  -> 3 presin ROLE kontagi (INPUT_PULLUP)
 *   - GND                       -> her 3 rolenin ORTAK ucu (COM)
 *   - GPIO2                     -> dahili LED (durum)
 *   Buzzer YOK — kapama presinde onay sesi kullanilmiyor (abkant/pres gibi).
 *
 *  KABLOLAMA (ortak pano):
 *   Role kontagi KAPANINCA pin GND'ye ceker -> 1 uretim pulse'u.
 *   Uc rolenin COM ucu ORTAK GND barasina baglanir; her rolenin NO ucu
 *   kendi GPIO'suna gider. Modul 5V (VIN) veya USB ile beslenir.
 *
 *   Dahili pull-up ZAYIFTIR (~45k). Pano ici kablo 2-3 metreyi asiyorsa veya
 *   yaninda kontaktor/invertor varsa her hat icin HARICI 4.7k pull-up (GPIO -> 3V3)
 *   ekleyin; ayrica GPIO ile GND arasina 100nF parazit yutucu iyi gelir. Uc hat
 *   ayni panoda yan yana gittigi icin capraz kuplaj riski tek kanalliya gore
 *   YUKSEKTIR — sahada tani sayaclarindan (parazit/erken) takip edin.
 *
 *  !!! UYARI — ONCE OLC:
 *   Role KURU KONTAK (potansiyelsiz) olmali. Panodan 24V'lu bir cikis
 *   (PLC sourcing / role bobin gerilimi) aliniyorsa GPIO'ya DIREKT BAGLAMA:
 *   ESP32 3.3V mantik seviyelidir, 24V pini ANINDA yakar. O durumda araya
 *   optokuplor (PC817 vb.) sart. Baglamadan once kontak uclari arasindaki
 *   gerilimi olcun — 0V okumali (kuru kontak).
 *
 *  Bu dosya pilot/firmware/_templates/tel_kapama.ino.tpl'den uretildi.
 * ============================================================ */

#include <WiFi.h>
#include <HTTPClient.h>
#include <Preferences.h>
#include <ArduinoJson.h>
#include <esp_task_wdt.h>
#include <ArduinoOTA.h>

// ════════════════════════════════════════════════════════════
//   YAPILANDIRMA — TEL-KAPAMA-4
// ════════════════════════════════════════════════════════════

const char* CIHAZ_ID  = "TEL-KAPAMA-4";
const char* BOLUM     = "tel";      // kapama, tel uretiminin proses adimidir

// Heartbeat/saglik kayitlarinda gorunen MODUL etiketi. Sayim sinyalleri
// kanalin KENDI robot_no'su ile gider (KANAL_ROBOT) — bu yalniz telemetri.
const char* HB_ROBOT_NO = "Kapama 10-12";

const char* WIFI_SSID = "COFLE-TK";
const char* WIFI_PASS = "internet2011!";
const char* SUNUCU_HOST = "http://192.168.21.155:5001";
const char* API_TOKEN = "cofle-pilot-2026";
const char* OTA_PASS = "cofle-ota-2026";

// ─── 3 KANAL: pin / makine adi / NVS anahtarlari ───────────────
// Sira DEGISMEZ: kanal indeksi (0,1,2) -> istasyon (1,2,3) -> NVS anahtari.
// Araya kanal eklenirse eski NVS sayaclari baska prese kayar.
const int KANAL_SAYISI = 3;
const int  KANAL_PIN[KANAL_SAYISI]   = { 25, 26, 27 };
const char* KANAL_ROBOT[KANAL_SAYISI] = { "Kapama 10", "Kapama 11", "Kapama 12" };
const char* NVS_PULSE[KANAL_SAYISI]  = { "pulse_i1", "pulse_i2", "pulse_i3" };
const char* NVS_SENT[KANAL_SAYISI]   = { "sent_i1",  "sent_i2",  "sent_i3"  };

const int PIN_LED = 2;

// Pulse algilama: BLOKLAMAYAN INTEGRATOR debounce.
// INTEGRATOR SURE TABANLI (v2.8, 2026-08-21) — pres sablonundan portlandi.
// ESKISI DONGU SAYIYORDU: her loop() turunda +1/-1, "15 ornek ~= 75ms" VARSAYIMIYLA.
// Loop suresi sabit degil (WiFi/HTTP isinde tur uzar) ve UC KANAL ayni turda
// orneklendigi icin sapma daha da buyuktu: gercek bir basis esigi tutturamadan
// bitebiliyordu. Artik gecen SURE (mikrosaniye) toplanir — esik ne yaziyorsa o
// kadar millisaniyedir, dongu hizindan BAGIMSIZ.
//
// ESIK/GAP ARTIK KANAL BAZLI (kullanici 2026-08-21): ayni pano icindeki uc presin
// cevrimi ayni degil. Saha kodu 27 olan preste bir urun icin ARDI ARDINA IKI kez
// basiliyor; 1000ms'lik ortak gap ikinci basisi "role cirpinmasi" sayip eliyordu
// ve kisa basislar 75ms esigini de tutturamiyordu -> o preste sayim HIC olusmuyordu.
// Degerler generate.py'den gelir (cihaz + kanal bazli); belirtilmeyen kanal
// v2.7 davranisini korur (75ms / 1000ms).
//   INTEG_ESIK_MS_n : n. kanalda bir basisin sayilmasi icin gereken net LOW suresi
//   MIN_GAP_MS_n    : n. kanalda iki sayim arasi en az sure
// SAHA AYARI: heartbeat TANI'sina bak — tani_erken artiyorsa GAP cok buyuk,
// KISA olaylari cikiyorsa ESIK cok buyuk, tani_parazit artiyorsa gurultu var.
const unsigned long INTEG_ESIK_US[KANAL_SAYISI] = {
  25UL * 1000UL, 75UL * 1000UL, 75UL * 1000UL };
const unsigned long MIN_PULSE_GAP_MS[KANAL_SAYISI] = {
  300UL, 1000UL, 1000UL };
// ─── ISR KENAR SAYACI (v2.8) — SAYMAZ, yalniz OLCER ───────────────────────
// Loop icindeki HTTP cagrilari 3sn'ye kadar BLOKLAR; o sure boyunca pin hic
// okunmaz ve o pencereye denk gelen vurus GENISLIGINDEN BAGIMSIZ kaybolur.
// Kesme blokede bile calisir -> gercek kenar sayisini verir. Ayrica esigi
// tutturamayan (KISA) basislarin genisligini olcmemizi saglar: sahada "esik kac
// olmali" sorusu tahminle degil OLCUMLE cevaplanir.
// DIKKAT — 2026-06-01 REGRESYONU TEKRARLANMASIN: o zaman CHANGE interrupt'ta
// digitalRead() ile kenar YONU tahmin ediliyordu; zayif pull-up yuzunden yavas
// yukselen kenarda ISR hala LOW okuyup state machine'i kilitliyordu. Buradaki
// ISR yalnizca FALLING'de tetiklenir ve SADECE sayac artirir: pin okumaz, yon
// tahmin etmez, sayim kararina KARISMAZ. Sayim yolu aynen polling+integrator.
volatile uint32_t      isrKenarSayisi[KANAL_SAYISI] = { 0, 0, 0 };
volatile unsigned long isrKenarUs[KANAL_SAYISI]     = { 0, 0, 0 };

const int  HEARTBEAT_MS   = 30000;
const int  RETRY_MS       = 3000;
const int  WIFI_TIMEOUT_S = 30;
const int  WDT_TIMEOUT_S  = 30;
const int  BUFFER_MAX     = 2000;  // kesinti kuyrugu — UC KANAL ORTAK kullanir
const int  RESEND_MAX_KANAL = BUFFER_MAX / KANAL_SAYISI;  // boot resend'de kanal basi tavan (bir kanal kuyrugu tek basina doldurmasin)
const char* FIRMWARE_VER  = "2.8.0-tel_kapama_4";   // 2.7.2 dayaniklilik paketi + 3 kanal

// ─── TANI (diagnostic) — sayim filtresi kararlarini heartbeat ile gonderir ──
// Tek kanalli sablonlarda 30/20; burada UC kanal ayni ringi paylasiyor, yani olay
// akisi 3 KAT. Ayni degerlerle ring iki heartbeat arasinda tasar ve tani akisi
// kullanilamaz hale gelirdi -> 45/30. (Kumulatif sayaclar zaten kesin; tasma
// yalnizca olay DETAYINI kaybettirir.)
const int TANI_MAX          = 45;   // RAM ring buffer (UC KANAL ORTAK)
const int TANI_HEARTBEAT_N  = 30;   // tek heartbeat'te en fazla kac olay

// ─── RSSI tabanli radyo recovery ─────────────────────────
const int           RSSI_ZAYIF_ESIK  = -80;
const unsigned long RSSI_KONTROL_MS  = 60000;
const int           RSSI_ZAYIF_MAX   = 5;
const int           RADYO_YENILE_MAX = 3;

// ─── Zombie WL_CONNECTED kurtarma + NVS watermark ────────────
const unsigned long SUNUCU_SESSIZ_HARD_RESET_MS = 300000UL;  // 5dk: sunucuya HIC ulasilamazsa watermark + reboot
const unsigned long WATERMARK_YAZ_MS            = 30000UL;   // NVS watermark yazma throttle
const unsigned long MAKINE_BOS_MS               = 120000UL;  // 2dk pulse yok = hat bos (bloklayan kurtarma serbest)

// ════════════════════════════════════════════════════════════
//   GLOBAL DURUM
// ════════════════════════════════════════════════════════════

Preferences prefs;

uint32_t pulse[KANAL_SAYISI]          = { 0, 0, 0 };
uint32_t sonYazilanSent[KANAL_SAYISI] = { 0, 0, 0 };

bool robotCalisiyor = false;

// Kuyruk kaydi: hangi KANALIN kacinci pulse'u. Kanal, POST'ta robot_no'yu ve
// istasyonu belirler — uc pres tek kuyrugu paylasir, sirasi bozulmaz.
struct PulseKaydi {
  uint32_t seq;
  uint8_t  kanal;   // 0..2
};
PulseKaydi buffer[BUFFER_MAX];
int buf_bas = 0, buf_son = 0, buf_dolu = 0;

unsigned long lastHeartbeat = 0;
unsigned long lastRetry     = 0;
unsigned long bootMs        = 0;

// Self-heal sayaclari
int httpFailCount = 0;
unsigned long lastSuccessHttp = 0;
// volatile: WiFi event task'i yazar, loop okur
volatile int disconnectCount = 0;
volatile unsigned long lastDisconnectMs = 0;
const unsigned long UPTIME_RESET_MS = 24UL * 3600UL * 1000UL;
unsigned long wifiKoptuMs = 0;           // wifiHazir() false'a dustugu an (0=bagli)
unsigned long retryGecikmeMs = RETRY_MS; // basarisiz POST'ta ustel artar (cap 60s)

unsigned long lastRadyoYenile = 0;
unsigned long lastWatermark   = 0;

// RSSI recovery durumu
unsigned long lastRssiKontrol = 0;
int rssiZayifSayac    = 0;
int radyoYenileSayac  = 0;
int sonRSSI           = 0;
uint32_t radyoYenileToplam = 0;
RTC_DATA_ATTR int rtcRssiHardReset = 0;
RTC_DATA_ATTR uint32_t rtcSunucuResetSayac = 0;

// ─── TANI durumu ────────────────────────────────────────────────
struct TaniOlay {
  uint32_t seq;        // cihaz ici artan olay no (boot_id ile idempotency)
  uint32_t uptime_ms;
  uint8_t  tip;        // 0=SAYILDI, 2=ERKEN
  uint32_t gap_ms;
  uint32_t low_ms;
  uint8_t  kanal;      // 0..2 (sunucu su an okumuyor; seri konsol/ileri analiz icin)
};
TaniOlay taniBuf[TANI_MAX];
int taniBas = 0, taniSon = 0, taniDolu = 0;
uint32_t taniSeq = 0;
uint32_t taniSayildi = 0, taniParazit = 0, taniErken = 0;
uint32_t bootId = 0;

// ─── Sayim kanali — integrator debounce durumu (kanal basina bir tane) ───
struct Kanal {
  unsigned long integUs;        // 0..integMaxUs(kanal) — net LOW birikimi (SURE)
  unsigned long sonOrnekUs;     // bir onceki ornegin micros() degeri
  bool          stableLow;      // debounced "basili" durumu
  int           lastRaw;        // ham pin durumu (gurultu kenar sayimi)
  unsigned long lastValidPulse; // son gecerli sayim (MIN_PULSE_GAP)
  bool          taniBekliyor;   // pulse devam ediyor; bitince SAYILDI push edilir
  unsigned long taniLowStart;
  uint32_t      taniBekleyenGap;
  // KISA PULSE OLCUMU (v2.8): esigi tutturamayan basislari da olcup raporlar.
  uint32_t      sonIslenenKenar;  // en son islenen isrKenarSayisi[kanal] degeri
  bool          kisaTakip;        // bir kenar yakalandi, pulse suruyor mu izleniyor
  unsigned long kisaBasUs;        // o kenarin ISR zamani
};
// TUM ALANLAR SIFIR — setup() gercek baslangic degerlerini yazar. Pozisyonel
// liste KULLANILMAZ: struct'a alan eklenince sessizce kayar ve yanlis alana yazar.
Kanal k[KANAL_SAYISI] = {};

// ISR govdeleri — struct'lardan SONRA (Arduino .ino otomatik-prototip ucu
// prototipleri ILK fonksiyon tanimindan once ekler; struct'tan once bir govde
// koymak "'Kanal' has not been declared" derleme hatasi uretir — 2026-08-21).
// Kanal basina AYRI ISR: tek ortak ISR hangi pinin tetikledigini bilemezdi.
void IRAM_ATTR isrKanal0() { isrKenarSayisi[0]++; isrKenarUs[0] = micros(); }
void IRAM_ATTR isrKanal1() { isrKenarSayisi[1]++; isrKenarUs[1] = micros(); }
void IRAM_ATTR isrKanal2() { isrKenarSayisi[2]++; isrKenarUs[2] = micros(); }

// Tavan (glitch toleransi tamponu) ve histerezis esikten TURETILIR — kanal basina
// ayri sabit tutmak, esik degisince guncellenmeyi unutturur.
// BURADA (struct'lardan SONRA), sabitlerin yaninda DEGIL: Arduino .ino otomatik
// prototip ucu, prototipleri dosyadaki ILK FONKSIYON TANIMINDAN once ekler.
// Bu iki inline sabitlerin yaninda dursaydi ilk tanim en tepeye kayar ve asagida
// tanimli struct'lari kullanan her prototip "has not been declared" ile duserdi
// (ayni tuzak 2026-08-21'de pres sablonunda yasandi).
inline unsigned long integMaxUs(uint8_t kn)   { return INTEG_ESIK_US[kn] * 3UL; }
inline unsigned long integDusukUs(uint8_t kn) { return INTEG_ESIK_US[kn] / 4UL; }

// ════════════════════════════════════════════════════════════
//   YARDIMCI FONKSIYONLAR
// ════════════════════════════════════════════════════════════

void ledYakBlink(int adet, int sure_ms = 80) {
  for (int i = 0; i < adet; i++) {
    digitalWrite(PIN_LED, HIGH); delay(sure_ms);
    digitalWrite(PIN_LED, LOW);  delay(sure_ms);
  }
}

// Uc kanaldan HERHANGI biri son MAKINE_BOS_MS icinde saydi mi?
// Bloklayan kurtarma (radyo tazele ~16sn) SADECE hat tamamen bosken calisir —
// yoksa calisan bir presin pulse'i kacar.
bool hatBos(unsigned long now) {
  for (int i = 0; i < KANAL_SAYISI; i++) {
    if ((now - k[i].lastValidPulse) <= MAKINE_BOS_MS) return false;
  }
  return true;
}

void wifiBaglan() {
  Serial.printf("[WiFi] '%s' agina baglaniliyor...\n", WIFI_SSID);
  WiFi.persistent(false);
  WiFi.setAutoReconnect(true);
  WiFi.mode(WIFI_STA);
  WiFi.setSleep(false);
  WiFi.setTxPower(WIFI_POWER_19_5dBm);
  // Coklu AP + ayni SSID (COFLE-TK): tum kanallari tara, EN GUCLU AP'ye baglan.
  WiFi.setScanMethod(WIFI_ALL_CHANNEL_SCAN);
  WiFi.setSortMethod(WIFI_CONNECT_AP_BY_SIGNAL);
  WiFi.begin(WIFI_SSID, WIFI_PASS);
  unsigned long basla = millis();
  while (WiFi.status() != WL_CONNECTED && (millis() - basla) < (WIFI_TIMEOUT_S * 1000)) {
    delay(500); Serial.print('.');
    digitalWrite(PIN_LED, !digitalRead(PIN_LED));
    esp_task_wdt_reset();
  }
  if (WiFi.status() == WL_CONNECTED) {
    Serial.printf("\n[WiFi] OK · IP=%s · RSSI=%d dBm · AP=%s\n",
                  WiFi.localIP().toString().c_str(), WiFi.RSSI(), WiFi.BSSIDstr().c_str());
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
  WiFi.setScanMethod(WIFI_ALL_CHANNEL_SCAN);
  WiFi.setSortMethod(WIFI_CONNECT_AP_BY_SIGNAL);
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
void taniPush(uint8_t tip, uint32_t gap_ms, uint32_t low_ms, uint8_t kanal) {
  if (taniDolu >= TANI_MAX) { taniBas = (taniBas + 1) % TANI_MAX; taniDolu--; }
  int idx = taniSon;
  taniBuf[idx].seq       = ++taniSeq;
  taniBuf[idx].uptime_ms = millis();
  taniBuf[idx].tip       = tip;
  taniBuf[idx].gap_ms    = gap_ms;
  taniBuf[idx].low_ms    = low_ms;
  taniBuf[idx].kanal     = kanal;
  taniSon = (taniSon + 1) % TANI_MAX;
  taniDolu++;
}

bool bufferdanBirGonder() {
  if (buf_dolu == 0) return false;
  PulseKaydi p = buffer[buf_bas];
  uint8_t kn  = (p.kanal < KANAL_SAYISI) ? p.kanal : 0;
  uint8_t ist = kn + 1;   // istasyon 1..3 — kanali ayirt eder (dashboard sayac sifirlama da bunu kullanir)

  HTTPClient http;
  String url = String(SUNUCU_HOST) + "/api/sinyal";
  http.begin(url);
  http.setTimeout(3000);
  http.setConnectTimeout(3000);
  http.addHeader("Content-Type", "application/json");
  http.addHeader("Authorization", String("Bearer ") + API_TOKEN);

  StaticJsonDocument<256> doc;
  doc["cihaz_id"]   = CIHAZ_ID;
  doc["bolum"]      = BOLUM;
  doc["robot_no"]   = KANAL_ROBOT[kn];   // KANALIN KENDI makinesi — sayim buna gore eslesir
  doc["istasyon"]   = ist;
  doc["kaynak_tip"] = "role";
  String mac6 = WiFi.macAddress();
  mac6.replace(":", "");
  if (mac6.length() > 6) mac6 = mac6.substring(mac6.length() - 6);
  // bootId NVS'te STABIL — watermark resend'inde ayni key uretilir, sunucu cift saymaz.
  String idem = String(CIHAZ_ID) + "_" + mac6 + "_b" + String(bootId) + "_i" + String(ist) + "_" + String(p.seq);
  doc["idempotency_key"] = idem;

  String payload;
  serializeJson(doc, payload);

  int rc = http.POST(payload);
  http.end();

  if (rc == 200 || rc == 201) {
    Serial.printf("[POST] OK %s seq=%lu (HTTP %d)\n", KANAL_ROBOT[kn], p.seq, rc);
    buf_bas = (buf_bas + 1) % BUFFER_MAX;
    buf_dolu--;
    httpFailCount = 0;
    lastSuccessHttp = millis();
    rtcSunucuResetSayac = 0;
    return true;
  } else {
    Serial.printf("[POST] HATA %s seq=%lu (HTTP %d) — kuyrukta kal\n",
                  KANAL_ROBOT[kn], p.seq, rc);
    httpFailCount++;
    return false;
  }
}

void kanalSinyali(uint8_t kanal) {
  if (kanal >= KANAL_SAYISI) return;
  pulse[kanal]++;
  uint32_t seq = pulse[kanal];
  prefs.putULong(NVS_PULSE[kanal], pulse[kanal]);

  if (buf_dolu >= BUFFER_MAX) {
    Serial.println("[BUFFER] DOLU! En eski pulse atildi");
    buf_bas = (buf_bas + 1) % BUFFER_MAX;
    buf_dolu--;
  }
  buffer[buf_son].seq   = seq;
  buffer[buf_son].kanal = kanal;
  buf_son = (buf_son + 1) % BUFFER_MAX;
  buf_dolu++;

  Serial.printf("[PULSE] %s #%lu (kuyruk=%d) — gonderiliyor...\n",
                KANAL_ROBOT[kanal], seq, buf_dolu);
  // Sayim yolu BLOKLANMAZ: LED blink + in-path HTTP POST burada YOK —
  // pulse aninda integrator donmasin (diger iki kanalin sayimi da kacmasin).
  // Pulse kuyrukta; loop'taki retry yolu gonderir.
}

// Bloklamayan integrator debounce — loop'ta her kanal icin bir kez cagrilir.
void kanalGuncelle(uint8_t kanal, unsigned long now) {
  Kanal* kk = &k[kanal];
  int v = digitalRead(KANAL_PIN[kanal]);

  if (kk->lastRaw == HIGH && v == LOW) taniParazit++;   // ham gurultu kenar gostergesi
  kk->lastRaw = v;

  // KISA PULSE TAKIBI: esigi tutturamayan basisi da OLCUP raporlayabilmek icin
  // dusen kenari ISR'den al (HTTP blokesinde bile yakalanir).
  uint32_t kenarSim = isrKenarSayisi[kanal];       // volatile tek okuma
  if (kenarSim != kk->sonIslenenKenar) {
    kk->sonIslenenKenar = kenarSim;
    if (!kk->stableLow) {                          // zaten sayilan bir basisin icinde degiliz
      kk->kisaTakip = true;
      kk->kisaBasUs = isrKenarUs[kanal];
    }
  }

  // Gecen SUREYI birikime kat (dongu hizindan bagimsiz).
  unsigned long esikUs  = INTEG_ESIK_US[kanal];
  unsigned long simdiUs = micros();
  unsigned long dtUs    = simdiUs - kk->sonOrnekUs;   // unsigned: overflow'da da dogru
  kk->sonOrnekUs = simdiUs;
  // KOR PENCERE TAVANI — tek ornek esigi TEK BASINA dolduramaz.
  // Loop HTTP'de 3sn bloklanabiliyor. Blok bitince pin LOW okunursa "3 saniyedir
  // basili" DIYEMEYIZ: o pencerede ne oldugunu bilmiyoruz. Esigin dortte biriyle
  // sinirlanir — gercek basis birkac ornekte yine dogrulanir, tek kor pencere
  // asla tek basina sayim yaratmaz.
  if (dtUs > esikUs / 4) dtUs = esikUs / 4;

  unsigned long tavanUs = integMaxUs(kanal);
  if (v == LOW) { kk->integUs = (kk->integUs + dtUs > tavanUs) ? tavanUs : kk->integUs + dtUs; }
  else          { kk->integUs = (kk->integUs < dtUs) ? 0 : kk->integUs - dtUs; }

  if (!kk->stableLow && kk->integUs >= esikUs) {
    kk->stableLow = true;
    unsigned long gap = now - kk->lastValidPulse;
    if (gap >= MIN_PULSE_GAP_MS[kanal]) {
      kanalSinyali(kanal);
      kk->lastValidPulse  = now;
      taniSayildi++;
      kk->taniBekliyor    = true;
      kk->taniLowStart    = now;
      kk->taniBekleyenGap = gap;
    } else {
      Serial.printf("[FILTRE] %s erken (gap=%lums < %lums) - SAYILMADI\n",
                    KANAL_ROBOT[kanal], gap, MIN_PULSE_GAP_MS[kanal]);
      taniErken++;
      taniPush(2, gap, 0, kanal);
      kk->taniBekliyor = false;
    }
  }

  // Kenar yakalanmisti ve pin HIGH'a dondu: basis SAYILMADAN bitti mi?
  // Bu olay (tip 3 = KISA) sahada esigi ayarlamanin tek saglam yolu — genislik
  // dagilimina bakilir, tahminle deger yazilmaz.
  if (kk->kisaTakip && v == HIGH && !kk->stableLow) {
    unsigned long genislikUs = simdiUs - kk->kisaBasUs;
    kk->kisaTakip = false;
    // 2ms ALTI RAPORLANMAZ: tek kanalli sablonda hepsi raporlanir, ama burada
    // TANI ringi UC KANAL ORTAK (45 slot). Kontak sicramasi mikro-kenarlari
    // ringi doldurup gercek SAYILDI olaylarini disari itiyordu — 2ms alti zaten
    // hicbir gercek presi temsil etmez, kumulatif tani_parazit'te gorunur.
    if (genislikUs >= 2000UL) {
      taniPush(3, (uint32_t)(now - kk->lastValidPulse), (uint32_t)(genislikUs / 1000UL), kanal);
      Serial.printf("[KISA] %s pulse %lu us (esik %lu us) - SAYILMADI\n",
                    KANAL_ROBOT[kanal], genislikUs, esikUs);
    }
  }
  if (kk->stableLow) kk->kisaTakip = false;      // sayildiysa takip gereksiz

  if (kk->stableLow && kk->integUs <= integDusukUs(kanal)) {
    kk->stableLow = false;
    if (kk->taniBekliyor) {
      taniPush(0, kk->taniBekleyenGap, now - kk->taniLowStart, kanal);
      kk->taniBekliyor = false;
    }
  }
}

// NVS watermark — KANAL BASINA sunucuya ulasmis son seq'i (sent_iN) yaz.
// Kuyruk uc kanal karisik oldugu icin her kanalin EN ESKI bekleyen seq'i bulunur;
// ondan bir onceki teyitlenmis demektir. Yipranma icin yalniz degisince yazar.
void watermarkYaz() {
  uint32_t enErken[KANAL_SAYISI] = { 0, 0, 0 };   // 0 = o kanalda bekleyen yok
  int idx = buf_bas;
  for (int n = 0; n < buf_dolu; n++) {
    uint8_t kn = buffer[idx].kanal;
    if (kn < KANAL_SAYISI && enErken[kn] == 0) enErken[kn] = buffer[idx].seq;
    idx = (idx + 1) % BUFFER_MAX;
  }
  for (int i = 0; i < KANAL_SAYISI; i++) {
    uint32_t gonderilen = enErken[i] ? (enErken[i] - 1) : pulse[i];
    if (gonderilen != sonYazilanSent[i]) {
      prefs.putULong(NVS_SENT[i], gonderilen);
      sonYazilanSent[i] = gonderilen;
    }
  }
}

void heartbeatGonder() {
  if (!wifiHazir()) return;
  HTTPClient http;
  http.begin(String(SUNUCU_HOST) + "/api/sinyal/heartbeat");
  http.setTimeout(3000);
  http.setConnectTimeout(3000);
  http.addHeader("Content-Type", "application/json");
  http.addHeader("Authorization", String("Bearer ") + API_TOKEN);

  // 8192: 30 tani olayi x 7 alan (~3,4KB) + telemetri. Tek kanalli sablondaki 4096
  // BU sablonda TASARDI — tasan JSON sessizce kirpilir, sunucu 400 doner, httpFail
  // artar ve gereksiz self-heal reset'i tetiklerdi. Heap gecicidir, bollugu zararsiz.
  DynamicJsonDocument doc(8192);
  doc["cihaz_id"]        = CIHAZ_ID;
  doc["bolum"]           = BOLUM;
  doc["robot_no"]        = HB_ROBOT_NO;   // MODUL etiketi (sayim degil, telemetri)
  doc["firmware_ver"]    = FIRMWARE_VER;
  doc["ip_adresi"]       = WiFi.localIP().toString();
  doc["mac_adresi"]      = WiFi.macAddress();
  doc["wifi_rssi"]       = WiFi.RSSI();
  doc["buffer_kuyruk"]   = buf_dolu;
  doc["uptime_sn"]       = (millis() - bootMs) / 1000;
  doc["free_heap"]       = (int)ESP.getFreeHeap();
  doc["pulse_ist1"]      = pulse[0];
  doc["pulse_ist2"]      = pulse[1];
  doc["pulse_ist3"]      = pulse[2];   // 3. kanal (sunucu su an okumuyor — ileri kullanim)
  doc["robot_calisiyor"] = robotCalisiyor;
  doc["radyo_yenile"]    = radyoYenileToplam;
  doc["http_fail"]       = httpFailCount;
  doc["disconnect"]      = disconnectCount;

  // ─── TANI: kumulatif sayaclar + son olaylar ───
  doc["tani_sayildi"] = taniSayildi;
  doc["tani_parazit"] = taniParazit;
  doc["tani_erken"]   = taniErken;
  // ISR kenar sayisi: HTTP blogu dahil TUM dusen kenarlar (uc kanalin toplami).
  // tani_parazit'ten belirgin BUYUKSE loop kor pencerede kenar kaciriyor demektir.
  doc["isr_kenar"]    = isrKenarSayisi[0] + isrKenarSayisi[1] + isrKenarSayisi[2];
  int taniGonder = taniDolu;
  if (taniGonder > TANI_HEARTBEAT_N) taniGonder = TANI_HEARTBEAT_N;
  if (taniGonder > 0) {
    JsonArray arr = doc.createNestedArray("tani");
    int i = taniBas;
    for (int kk = 0; kk < taniGonder; kk++) {
      JsonObject o = arr.createNestedObject();
      o["s"] = taniBuf[i].seq;  o["b"] = bootId;            o["u"] = taniBuf[i].uptime_ms;
      o["t"] = taniBuf[i].tip;  o["g"] = taniBuf[i].gap_ms; o["l"] = taniBuf[i].low_ms;
      o["k"] = taniBuf[i].kanal + 1;   // istasyon no (1..3) — hangi pres
      i = (i + 1) % TANI_MAX;
    }
  }

  String payload; serializeJson(doc, payload);
  int rc = http.POST(payload);
  http.end();
  if (rc == 200 || rc == 201) {
    httpFailCount = 0;
    lastSuccessHttp = millis();
    rtcSunucuResetSayac = 0;
    taniBas = (taniBas + taniGonder) % TANI_MAX;
    taniDolu -= taniGonder;
  } else {
    httpFailCount++;
  }
  Serial.printf("[HEART] HTTP %d · RSSI=%d · kuyruk=%d · sayim(%lu/%lu/%lu) · tani(S%lu/P%lu/E%lu) · fail=%d\n",
                rc, WiFi.RSSI(), buf_dolu, pulse[0], pulse[1], pulse[2],
                taniSayildi, taniParazit, taniErken, httpFailCount);
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
  Serial.printf("FW: %s · Bolum: %s · 3 KANALLI ROLE MODU\n", FIRMWARE_VER, BOLUM);

  pinMode(PIN_LED, OUTPUT);
  digitalWrite(PIN_LED, LOW);

  for (int i = 0; i < KANAL_SAYISI; i++) {
    pinMode(KANAL_PIN[i], INPUT_PULLUP);
    k[i].integUs        = 0;
    k[i].sonOrnekUs     = micros();     // ilk ornek zamani (dtUs ilk turda patlamasin)
    k[i].stableLow      = false;
    k[i].lastRaw        = digitalRead(KANAL_PIN[i]);
    k[i].lastValidPulse = 0;
    k[i].taniBekliyor   = false;
    k[i].taniLowStart   = 0;
    k[i].taniBekleyenGap = 0;
    k[i].sonIslenenKenar = 0;
    k[i].kisaTakip      = false;
    k[i].kisaBasUs      = 0;
    Serial.printf("  Kanal %d · GPIO%-2d -> %s | esik=%lums gap=%lums\n",
                  i + 1, KANAL_PIN[i], KANAL_ROBOT[i],
                  INTEG_ESIK_US[i] / 1000UL, MIN_PULSE_GAP_MS[i]);
  }
  // ISR yalniz OLCER (sayim kararina KARISMAZ) - bkz. isrKanal0/1/2 aciklamasi.
  // Kanal basina ayri ISR: tek ortak ISR hangi pinin tetikledigini bilemezdi.
  attachInterrupt(digitalPinToInterrupt(KANAL_PIN[0]), isrKanal0, FALLING);
  attachInterrupt(digitalPinToInterrupt(KANAL_PIN[1]), isrKanal1, FALLING);
  attachInterrupt(digitalPinToInterrupt(KANAL_PIN[2]), isrKanal2, FALLING);
  Serial.printf("MAC: %s\n", WiFi.macAddress().c_str());

  robotCalisiyor = false;

  prefs.begin("cofle", false);
  for (int i = 0; i < KANAL_SAYISI; i++) {
    pulse[i] = prefs.getULong(NVS_PULSE[i], 0);
  }
  Serial.printf("[NVS] Kayitli sayimlar: %lu / %lu / %lu\n", pulse[0], pulse[1], pulse[2]);

  // bootId STABIL (reboot'lar arasi ayni) — watermark resend'i ayni idempotency_key'i
  // uretsin diye. NVS silinirse boot_id de gider -> yeni bootId + sayim 0'dan baslar,
  // eski key'lerle CAKISMAZ.
  bootId = prefs.getULong("boot_id", 0);
  if (bootId == 0) {
    bootId = esp_random();
    prefs.putULong("boot_id", bootId);
  }

  // Reboot resend — KANAL BASINA: gonderilememis [sent_iN+1 .. pulse_iN] kuyruga geri kondu.
  // Eski firmware'de sent_iN yok -> default pulse_iN (hepsi gonderildi varsay).
  for (int i = 0; i < KANAL_SAYISI; i++) {
    uint32_t sentSeq = prefs.getULong(NVS_SENT[i], pulse[i]);
    if (sentSeq > pulse[i]) sentSeq = pulse[i];   // tutarsizlik guvenligi
    sonYazilanSent[i] = sentSeq;
    if (pulse[i] > sentSeq) {
      uint32_t resendBas = sentSeq + 1;
      uint32_t adet      = pulse[i] - sentSeq;
      if (adet > (uint32_t)RESEND_MAX_KANAL) {   // en yeniyi tut (bir kanal kuyrugu tek basina doldurmasin)
        resendBas = pulse[i] - RESEND_MAX_KANAL + 1;
        adet = RESEND_MAX_KANAL;
      }
      for (uint32_t s = resendBas; s <= pulse[i]; s++) {
        buffer[buf_son].seq   = s;
        buffer[buf_son].kanal = (uint8_t)i;
        buf_son = (buf_son + 1) % BUFFER_MAX;
        buf_dolu++;
      }
      Serial.printf("[NVS] %s: gonderilememis %lu pulse kuyruga kondu (seq %lu..%lu)\n",
                    KANAL_ROBOT[i], (unsigned long)adet,
                    (unsigned long)resendBas, (unsigned long)pulse[i]);
    }
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
    // OTA transferi dakikalarca surebilir — 30sn task-WDT yukleme ortasinda
    // panik reset atmasin diye loopTask'i izlemeden cikar.
    esp_task_wdt_delete(NULL);
    Serial.println("\n[OTA] Guncelleme basliyor (WDT askida)");
  });
  ArduinoOTA.onEnd([]()   { Serial.println("\n[OTA] Tamam, yeniden baslatiliyor"); });
  ArduinoOTA.onProgress([](unsigned int p, unsigned int t) {
    Serial.printf("[OTA] %u%%\r", (p / (t / 100)));
  });
  ArduinoOTA.onError([](ota_error_t e) {
    Serial.printf("[OTA] HATA %u\n", e);
    esp_task_wdt_add(NULL);   // OTA basarisiz — WDT korumasini geri tak
  });
  ArduinoOTA.begin();
  Serial.printf("[OTA] Aktif — IDE Network Port: %s @ %s\n",
                CIHAZ_ID, WiFi.localIP().toString().c_str());

  heartbeatGonder();
  lastHeartbeat = millis();
  lastRetry = millis();

  Serial.println("[READY] SURE tabanli integrator + TANI aktif - 3 kanal role sinyali bekleniyor...");
  for (int i = 0; i < KANAL_SAYISI; i++) {
    Serial.printf( "        %s: esik=%lums gap=%lums\n",
                   KANAL_ROBOT[i], INTEG_ESIK_US[i] / 1000UL, MIN_PULSE_GAP_MS[i]);
  }
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

  // UC KANAL — bloklamayan integrator debounce (her dongude hepsi ornekleniyor)
  for (uint8_t i = 0; i < KANAL_SAYISI; i++) {
    kanalGuncelle(i, now);
  }

  // Retry: basarisiz gonderimde ustel geri cekilme (3s -> 60s cap).
  if (wifiHazir() && buf_dolu > 0 && (now - lastRetry) > retryGecikmeMs) {
    if (bufferdanBirGonder()) {
      retryGecikmeMs = RETRY_MS;
    } else {
      retryGecikmeMs = (retryGecikmeMs * 2 > 60000UL) ? 60000UL : retryGecikmeMs * 2;
    }
    lastRetry = now;
  }

  // Heartbeat: sunucu cevap vermiyorsa araligi 4x'e cikar (30s->120s)
  {
    unsigned long hbAralik = (httpFailCount >= 3) ? (HEARTBEAT_MS * 4) : HEARTBEAT_MS;
    if (wifiHazir() && (now - lastHeartbeat) > hbAralik) {
      heartbeatGonder();
      lastHeartbeat = now;
    }
  }

  // NVS watermark periyodik yazimi (WiFi'den BAGIMSIZ) — beklenmeyen reset/guc
  // kesintisinde tekrar-gonderim penceresini ~30s ile sinirlar.
  if ((now - lastWatermark) > WATERMARK_YAZ_MS) {
    watermarkYaz();
    lastWatermark = now;
  }

  // ─── SELF-HEAL kontrolleri ───
  // 1) 24 saat uptime → koruyucu reset, AMA hat bos ve kuyruk bosken
  if ((now - bootMs) > UPTIME_RESET_MS) {
    if (hatBos(now) && buf_dolu == 0) {
      Serial.println("\n[SELFHEAL] 24h doldu + hat bos + kuyruk bos — koruyucu reset");
      delay(500);
      ESP.restart();
    }
  }
  // 2) Sunucuya ulasilamiyor — ZOMBIE WL_CONNECTED dahil (RSSI iyi gorunur,
  //    disconnect event'i ATISLANMAZ; lastSuccessHttp tabanli BAGIMSIZ kurtarma sart).
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
      // Kuyruk DOLU ama 5dk'dir sunucuya HIC ulasilamiyor -> watermark yaz, GUVENLE
      // reboot et. Boot'ta kanal basina resend yapilir (stabil bootId -> cift sayim yok).
      watermarkYaz();
      Serial.printf("\n[SELFHEAL] Sunucuya %lus ulasilamadi (zombie?) — watermark(kuyruk=%d) + REBOOT\n",
                    sessizMs / 1000, buf_dolu);
      delay(500);
      ESP.restart();
    } else if ((now - lastRadyoYenile) > 60000UL) {
      // Once YUMUSAK kurtarma: radyo tazele. SADECE HAT BOSKEN — wifiRadyoYenile
      // ~16sn bloklar, calisan bir presin pulse'ini kaybettirir.
      if (hatBos(now)) {
        Serial.printf("\n[SELFHEAL] %d HTTP fail, %lus sessiz, kuyruk=%d — radyo tazele\n",
                      httpFailCount, sessizMs / 1000, buf_dolu);
        wifiRadyoYenile();
        lastRadyoYenile = now;
      }
    }
  }
  // 2b) WiFi 5dk'dir HIC baglanamiyor (AP kayip / assoc wedge) — httpFailCount'tan
  //     BAGIMSIZ kurtarma. POST hic denenemedigi icin fail sayaci artmaz.
  if (wifiKoptuMs != 0 && (now - wifiKoptuMs) > SUNUCU_SESSIZ_HARD_RESET_MS) {
    watermarkYaz();
    Serial.printf("\n[SELFHEAL] WiFi %lus'dir baglanamiyor — watermark(kuyruk=%d) + REBOOT\n",
                  (now - wifiKoptuMs) / 1000, buf_dolu);
    delay(500);
    ESP.restart();
  }
  if (disconnectCount >= 3 && (now - lastDisconnectMs) < 600000UL && buf_dolu == 0) {
    Serial.printf("\n[SELFHEAL] 10dk icinde %d disconnect + kuyruk bos — RESET\n",
                  disconnectCount);
    delay(500);
    ESP.restart();
  }
  // 4) RSSI tabanli radyo recovery — sadece hat BOS iken (blok pulse kaybetmesin)
  if (wifiHazir() && (now - lastRssiKontrol) > RSSI_KONTROL_MS) {
    lastRssiKontrol = now;
    sonRSSI = WiFi.RSSI();
    bool bos = hatBos(now);
    if (sonRSSI < RSSI_ZAYIF_ESIK && sonRSSI < 0) {
      rssiZayifSayac++;
      Serial.printf("[RSSI] Zayif sinyal %d dBm (%d/%d)%s\n", sonRSSI, rssiZayifSayac, RSSI_ZAYIF_MAX,
                    bos ? "" : " — pres aktif, recovery ertelendi");
      if (rssiZayifSayac >= RSSI_ZAYIF_MAX && bos) {
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
