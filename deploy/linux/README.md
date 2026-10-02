# Markalı kurulum — bulut sunucu (Ubuntu VPS)

Aynı kod, firmaya özgü ad / bölüm / makine / modül ayarı `data/kurulum.json`
dosyasından gelir (bkz. `kurulum.py`). Cofle'nin kendi sunucusunda bu dosya yoktur.

## Ön koşullar
- **Ubuntu 24.04** VPS — Python 3.12 gerekir (22.04'ün 3.10'u yetmez); 2 vCPU / 2 GB RAM / 20 GB disk yeter
- Alan adının **A kaydı** VPS'in IP adresini göstermeli (HTTPS sertifikası için şart)
- Depo özel ise sunucuya salt okunur erişim (GitHub "deploy key")

## 1. Sunucuyu hazırla (bir kez)
```bash
sudo bash kur.sh uretim.firma.com https://github.com/Judeless/uretim-takip.git
```
Paketler, İstanbul saat dilimi, `uretim` kullanıcısı, `/opt/uretim`, Python ortamı,
systemd servisi, Caddy (otomatik HTTPS), günlük yedek (02:30) ve güvenlik duvarı.

## 2. Profil + veritabanı + yönetici (bir kez)
`ornek_kurulum/tedarikci.json` dosyasını düzenleyip sunucuya kopyala (firma adı,
alan adı, makineler), referans Excel'ini de yükle; sonra:
```bash
cd /opt/uretim
sudo -u uretim .venv/bin/python kurulum_hazirla.py --profil profil.json \
     --excel veri.xlsb --yonetici admin --yonetici-ad "Ad Soyad"
sudo systemctl start uretim-takip
```
Ekrana bir kez yazılanlar: panel yöneticisinin **geçici şifresi** (ilk girişte
değiştirilir) ve mobildeki **Admin operatörünün PIN'i**. Bir yere not alın.

## Günlük işler
| İş | Komut |
|---|---|
| Güncelleme (kod) | `sudo bash /opt/uretim/deploy/linux/guncelle.sh` |
| Durum / log | `systemctl status uretim-takip` · `journalctl -u uretim-takip -f` |
| Elle yedek | `sudo -u uretim /usr/local/bin/uretim-yedek` (→ `/opt/uretim/yedek/`) |
| Profil değişikliği | `data/kurulum.json` düzenle → `sudo systemctl restart uretim-takip` |

## Mail (17:00 günlük rapor)
`mail_config.json.example` → `mail_config.json` (SMTP bilgileri, `"etkin": true`).
Linux'ta Outlook yolu yoktur; SMTP kullanılır.

## Logo
`static/` altına `marka_logo.png` (açık zemin) ve `marka_logo_koyu.png` (andon) koyup
profilde `"logo"` / `"logo_koyu"` alanlarına yolunu yazın. Boşsa firma adından yazı
logosu üretilir — Cofle logosu bu kurulumda hiçbir koşulda görünmez.
