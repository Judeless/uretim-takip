/* ============================================================
 *  KLEMENS SHIELD + GPIO27 TEST — bagimsiz tani sketch'i
 *  (WiFi/sunucu YOK; sadece Seri Monitor + dahili LED)
 * ============================================================
 *
 *  AMAC: GPIO27 klemensini GND klemensine degdirince sinyal
 *        aliniyor mu + sayilabiliyor mu gormek. Pres/abkant
 *        sayacinin gercek okuma yontemi budur (kuru kontak).
 *
 *  BAGLANTI:
 *   - Bir kabloyu  GPIO27 klemensine
 *   - Diger ucunu  GND klemensine  degdir/cek
 *  (Buton varsa: butonun bir bacagini GPIO27, digerini GND'ye.)
 *
 *  BEKLENEN:
 *   - Degmiyorken : seviye = ACIK (HIGH)  · LED sonuk
 *   - Degince     : seviye = KAPALI (LOW) · LED yanar · sayac +1
 *
 *  KULLANIM:
 *   1) Karti sec: "ESP32 Dev Module", dogru COM portu
 *   2) Yukle, sonra Arac > Seri Monitor (115200 baud)
 *   3) GPIO27 <-> GND degdir; her temasta "SAYIM: n" artmali
 *
 *  NOT: INPUT_PULLUP kullaniliyor -> harici direnc GEREKMEZ.
 *       ESP32 GPIO 3.3V mantik; buraya SADECE kuru kontak/buton
 *       baglanir. 24V role/buton cikisi ISE DIREKT BAGLAMA
 *       (once optokuplor) — aksi halde kart yanar.
 * ============================================================ */

const int PIN_SINYAL = 27;   // test edilen GPIO (klemens)
const int PIN_LED    =  2;   // ESP32 dahili LED

// Debounce: mekanik temas sicrar (bounce); ~40ms stabil kalinca gecerli say.
const unsigned long DEBOUNCE_MS = 40;

int           sonKararliSeviye = HIGH;   // en son onaylanan seviye (HIGH=acik)
int           sonOkunan        = HIGH;   // ham okuma
unsigned long sonDegisimTs     = 0;      // ham okumanin en son degistigi an
unsigned long sayim            = 0;      // GPIO27->GND kac kez degdi (dusen kenar)
unsigned long sonYazdirma      = 0;

void setup() {
  Serial.begin(115200);
  delay(300);
  pinMode(PIN_SINYAL, INPUT_PULLUP);   // dahili pull-up: bosta HIGH, GND'ye deginca LOW
  pinMode(PIN_LED, OUTPUT);
  digitalWrite(PIN_LED, LOW);

  Serial.println();
  Serial.println("==== KLEMENS / GPIO27 TEST ====");
  Serial.println("GPIO27 <-> GND degdir. Her temasta SAYIM artmali.");
  Serial.println("Baslangic seviye: ACIK bekleniyor...");
  Serial.println("--------------------------------");
}

void loop() {
  int ham = digitalRead(PIN_SINYAL);

  // Ham okuma degistiyse debounce sayacini basa al
  if (ham != sonOkunan) {
    sonOkunan = ham;
    sonDegisimTs = millis();
  }

  // Ham okuma DEBOUNCE_MS boyunca sabit kaldiysa "kararli" kabul et
  if ((millis() - sonDegisimTs) >= DEBOUNCE_MS && ham != sonKararliSeviye) {
    sonKararliSeviye = ham;
    digitalWrite(PIN_LED, (sonKararliSeviye == LOW) ? HIGH : LOW);

    if (sonKararliSeviye == LOW) {           // GND'ye degdi = dusen kenar = 1 sayim
      sayim++;
      Serial.print("KAPALI (bagli)  ->  SAYIM: ");
      Serial.println(sayim);
    } else {
      Serial.println("ACIK (birakildi)");
    }
  }

  // Her 1 sn'de bir anlik durum (canli oldugunu gormek icin)
  if (millis() - sonYazdirma >= 1000) {
    sonYazdirma = millis();
    Serial.print("[durum] seviye=");
    Serial.print(sonKararliSeviye == LOW ? "KAPALI" : "ACIK");
    Serial.print("  toplam sayim=");
    Serial.println(sayim);
  }
}
