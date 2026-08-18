# SERVER'A AS400 TEYİT TAŞIMA — Kurulum Kılavuzu

Hedef: Teyit robotu (PCOMM Session B) **server'da** (192.168.20.210) koşsun →
laptop bağımlılığı bitsin, teyit her zaman **canlı DB** ile verilsin
(2026-07-24 "83 vs 5600" bayat-veri hatası sınıfı tamamen kapanır).

## Mimari (neden agent var?)

`cofle-app` NSSM **servisi** Session 0'da koşar; PCOMM pencereleri (pcsws.exe)
promanage'ın **RDP oturumunda**. Windows oturum izolasyonu → servisin başlattığı
cscript PCOMM'u GÖREMEZ. Çözüm: RDP oturumunda koşan **`teyit_agent.py`**
(127.0.0.1:5010, yalnız yerel) robot cscript'lerini çalıştırır; ana app robot
çağrılarını otomatik ona devreder (config gerekmez — agent ayaktaysa kullanılır,
değilse yerel subprocess = laptop davranışı).

```
[NSSM cofle-app :5000] --HTTP 127.0.0.1:5010--> [teyit_agent.py] --cscript--> [PCOMM B]
      Session 0                                     RDP oturumu (promanage)
```

## Kurulum adımları (sırayla)

### 1. iSeries Access for Windows V5R4 kur
- PCOMM 5.8.3 + **iSeries Access ODBC Driver** birlikte gelir. Kurulum medyası IT'den
  (laptopta kurulu olan paketin aynısı).
- Kurulum sonrası kontrol (promanage PowerShell):
  `python -c "import pyodbc; print(pyodbc.drivers())"` → listede **iSeries Access ODBC Driver** olmalı.

### 2. Python paketleri
```
pip install pyodbc keyring requests
```
(`requests` requirements.txt'e eklendi; pyodbc+keyring AS400'e özel.)

### 3. Ağ erişimi doğrula
- `ping 192.168.1.1` (AS400 host)
- PCOMM ilk oturum açılışında SSL bağlantısı kurulmalı (sertifika sorarsa kabul et).

### 4. Oturum profilleri (A + B)
- Laptop `Masaüstü\AS400-1.ws` dosyasını server'a kopyala → 2 kopya yap.
- Her ikisinde **WorkStationID'yi DEĞİŞTİR** (laptop SELCUK0001 ile çakışmasın):
  örn. `COFSRV01` / `COFSRV02` (.ws dosyası not defteri ile açılır, `WorkStationID=` satırı).
- **Açılış SIRASI önemli:** ilk açılan pencere = Session **A** (insan/izleme),
  ikinci = Session **B** (robotun sürdüğü). Robot `SetConnectionByName("B")` kullanır.

### 5. AS400 şifresi (keyring) — servisi DEĞİŞTİRME, agent köprüler
> ⚠️ NSSM servis hesabını promanage'a ALMA. Denendi (2026-07-24) → "log on as a
> service" tuzağı + logon-failure → dashboard düştü. cofle-app **LocalSystem'de
> KALIR**. Keyring hesaba özeldir ama şifreyi **teyit-agent** (promanage oturumu)
> okuyup servise localhost'tan verir (`as400_config.sifre_al()` → keyring, olmazsa
> agent `/sifre`). Yani servisin keyring'i görmesine gerek yok.

- promanage RDP oturumunda AS400 şifresini kaydet:
  `cd C:\cofle\uretim_takip\as400 && python kaydet_sifre.py`
  (EMREDTK şifresi Credential Manager'a — hiçbir dosyaya düz metin yazılmaz.)
- Bu şifreyi agent (aynı oturum) görecek ve servise verecek. **Agent'ın çalışıyor
  olması ŞART** — agent kapalıysa servis AS400'ü okuyamaz (zaten teyit için PCOMM
  de gerekir, tutarlı).

### 6. Teyit-agent'ı otomatik başlat
- promanage RDP oturumunda `Win+R` → `shell:startup` → içine
  `C:\cofle\uretim_takip\as400\Teyit_Agent_Baslat.bat` **kısayolu** koy.
- İlk sefer için bat'ı elle çalıştır — konsolda
  `Dinleme: 127.0.0.1:5010 · PCOMM pencere: 2` görmelisin.

### 7. Günlük işletme kuralı
- Server yeniden başlarsa: promanage RDP ile gir → PCOMM A ve B'yi aç, **elle
  oturum aç** (sign-on İNSAN işi — robot şifre girmez, güvenlik kuralı) → agent
  penceresinin kalktığını gör → RDP'yi **Disconnect** et (LOGOFF ETME! Logoff
  oturumu kapatır, PCOMM + agent ölür).

> ### ⚠️ "Pencereler açık" YETERLİ DEĞİL — otomasyonu doğrula (2026-08-17 olayı)
>
> Bir haftalık aradan sonra teyit gönderimi 14 satırın 14'ünde düştü. `pcsws.exe`
> **iki tane çalışıyordu**, agent'la **aynı oturumdaydı** (Session 2), sign-on da
> yapılmıştı — ama `PCOMM.autECLConnList.Refresh()` **0 bağlantı** görüyordu.
> Emülatör pencereleri otomasyon katmanına KAYITLI DEĞİLDİ; robot
> `SetConnectionByName("B")` derken `ECL37110 — emülasyon arayüzü kullanılamıyor`
> alıyordu. cscript bu hatada bile **0 ile çıkar**, hata metni **stderr**'e gider.
>
> **Doğrulama ölçüsü süreç listesi değil, bağlantı sayısıdır:**
>
> ```
> C:\cofle\uretim_takip\as400\Robot_Tani.bat
> ```
>
> "Otomasyonun gördüğü bağlantı sayısı" **2** ve adlar **[A]**, **[B]** olmalı.
> 0 görüyorsan PCOMM'u **kapat**, promanage oturumunda **yönetici olmadan** önce
> A sonra B'yi aç, sign-on yap, tanıyı tekrar çalıştır. (Olay bu şekilde çözüldü.)
>
> Not: `tasklist` bu sunucuda asılıyor — süreç bakmak gerekirse `Get-Process`
> kullan (tanı aracı zaten öyle yapıyor).

### 8. Canlı test
1. Dashboard (server) → AS400 Teyit → listeyi yenile (ODBC + keyring testi).
2. TEK launch seç → Gönder → Session B'yi RDP'den izle (robot testi).
3. Agent konsolunda `[AGENT] teyit_gir.js ...` satırı + app sonucunda doğrulama mesajı.

## Geri dönüş
Laptoptan teyit hâlâ mümkün: laptop'ta app konsol uygulaması olarak koşar,
agent yoktur → yerel subprocess yolu aynen çalışır. (Ama laptop DB'si BAYAT —
yalnız acil durumda ve DB tazeleyerek kullan.)

## 9. Oturum gözcüsü — Session B düşerse otomatik geri getirir (2026-08-17)

Session B uzun süre işlem yapılmayınca düşüyor (AS400 `QINACTITV` ya da bağlantı
kopması); teyit robotu bir daha çalışamıyor. Kullanıcı izindeyken elle sign-on
yapacak kimse olmadığı için gözcü eklendi.

> ⚠️ Bu, madde 7'deki **"sign-on İNSAN işi"** kuralını bilerek gevşetir
> (kullanıcı isteği, 2026-08-17). Denge şöyle korundu:
> - Şifre **yine kasada** (Credential Manager / DPAPI) — kodda, config'de, log'da
>   **yok**. `oturum_config.json` yalnız **kullanıcı adını** tutar (gitignore'da).
> - Şifre alt sürece **yalnız ortam değişkeniyle** geçer (`COFLE_AS400_PW`).
>   Argümanla geçilseydi agent konsoluna ve Windows süreç listesine düşerdi —
>   bu yüzden agent bu script için argümanları da loglamaz.
> - Robot **yanlış şifre denemez**: sign-on reddedilirse durur ve ekran mesajını
>   raporlar (AS400 `QMAXSIGN` profil kilitleme riskine girilmez).
> - Varsayılan **KAPALI**; config yoksa/bozuksa da kapalı.

**Açmak için** (promanage oturumunda, bir kez):

```
cd C:\cofle\uretim_takip\as400
copy oturum_config.json.example oturum_config.json
notepad oturum_config.json      ->  "etkin": true,  "kullanici": "<AS400 kullanici adi>"
python kaydet_sifre.py           (sifre zaten kayitliysa gerekmez)
```

Sonra **agent'ı yeniden başlat** (`Teyit_Agent_Baslat.bat`). Konsolda
`Oturum gözcüsü AÇIK — kullanıcı …, her 300 sn kontrol` yazmalı.

**Önce güvenli test** (hiçbir tuş göndermez, şifreyi kullanmaz):

```
C:\Windows\SysWOW64\cscript.exe //nologo oturum_ac.js <kullanici> TESTBAGLAN
```

Bağlantı yoksa kurar, sign-on ekranını ve **alan konumlarını** raporlar, kendi
kurduğu bağlantıyı geri kapatır. `utente alani = kol N` satırları geliyorsa alan
tespiti doğrudur. `DRYRUN` ise hiç bağlanmaz, yalnız mevcut durumu söyler.

**Sonuçlar:** `SONUC=OK` düşmüştü, giriş yapıldı · `SONUC=ZATEN` sağlıklı ya da
robot başka ekranda bırakmış (**dokunmaz**) · `SONUC=IPTAL` sorun var, konsolda
gerekçesiyle yazar. Gözcü teyit robotu koşarken **araya girmez** (aynı kilidi
bekler) ve yalnız durum değişince ekrana yazar.

**Bilinen tuzak:** `CommStarted=true` oturumun kullanılabilir olduğunu
**göstermez**. Ölçümde bağlantı kuruldu, OIA "hazır" dedi ama host 30 sn boyunca
tek satır göndermedi. Robot bu yüzden kararını **ekran içeriğine** göre verir;
bağlı-ama-boş ekranı arıza sayar ve kendi kurduğu yarım bağlantıyı kapatır.
