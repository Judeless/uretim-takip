#!/usr/bin/env bash
# ════════════════════════════════════════════════════════════════════════
#  MARKALI KURULUM — Ubuntu 24.04 bulut sunucu (VPS) ilk kurulum (Python 3.12+)
#  Kullanım (root ya da sudo ile):
#     sudo bash kur.sh alanadi.com https://github.com/Judeless/uretim-takip.git
#  Yaptıkları: paketler · saat dilimi (İstanbul) · 'uretim' kullanıcısı ·
#  kod /opt/uretim · Python ortamı · systemd servisi · Caddy (otomatik HTTPS) ·
#  günlük veritabanı yedeği · güvenlik duvarı (yalnız 22/80/443)
#  Profil + veritabanı ADIM 2'de: deploy/linux/README.md
# ════════════════════════════════════════════════════════════════════════
set -euo pipefail

ALAN="${1:?Alan adı gerekli (ör. uretim.firma.com)}"
DEPO="${2:?Git depo adresi gerekli}"
DIZIN=/opt/uretim
KULLANICI=uretim

echo "== 1/7 Paketler"
apt-get update -y
apt-get install -y python3 python3-venv python3-pip git sqlite3 debian-keyring \
    debian-archive-keyring apt-transport-https curl ufw

python3 -c 'import sys; sys.exit(0 if sys.version_info >= (3, 12) else 1)' || {
    echo "DURDU: Python 3.12+ gerekli (Ubuntu 24.04). Bu sistemde: $(python3 --version)"; exit 1; }

echo "== 2/7 Saat dilimi: zamanlanmış işler (16:30 vardiya kapatma, 17:00 mail, 18:00 arşiv) yerel saatle koşar"
timedatectl set-timezone Europe/Istanbul

echo "== 3/7 Kullanıcı + kod"
id -u "$KULLANICI" >/dev/null 2>&1 || useradd --system --create-home --shell /usr/sbin/nologin "$KULLANICI"
if [ ! -d "$DIZIN/.git" ]; then
    git clone "$DEPO" "$DIZIN"
fi
chown -R "$KULLANICI:$KULLANICI" "$DIZIN"

echo "== 4/7 Python ortamı"
sudo -u "$KULLANICI" python3 -m venv "$DIZIN/.venv"
sudo -u "$KULLANICI" "$DIZIN/.venv/bin/pip" install --upgrade pip
sudo -u "$KULLANICI" "$DIZIN/.venv/bin/pip" install -r "$DIZIN/requirements.txt"

echo "== 5/7 systemd servisi"
install -m 0644 "$DIZIN/deploy/linux/uretim-takip.service" /etc/systemd/system/uretim-takip.service
systemctl daemon-reload
systemctl enable uretim-takip

echo "== 6/7 Caddy (HTTPS sertifikası Let's Encrypt'ten otomatik)"
if ! command -v caddy >/dev/null; then
    curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/gpg.key' | gpg --dearmor -o /usr/share/keyrings/caddy-stable-archive-keyring.gpg
    curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/debian.deb.txt' > /etc/apt/sources.list.d/caddy-stable.list
    apt-get update -y && apt-get install -y caddy
fi
sed "s/ALAN_ADI/$ALAN/" "$DIZIN/deploy/linux/Caddyfile" > /etc/caddy/Caddyfile
systemctl reload caddy || systemctl restart caddy

echo "== 7/7 Güvenlik duvarı + günlük yedek (02:30)"
ufw allow OpenSSH && ufw allow 80/tcp && ufw allow 443/tcp && ufw --force enable
install -m 0755 "$DIZIN/deploy/linux/yedek.sh" /usr/local/bin/uretim-yedek
echo "30 2 * * * $KULLANICI /usr/local/bin/uretim-yedek" > /etc/cron.d/uretim-yedek

echo
echo "TAMAM. Sıradaki adım (profil + veritabanı + yönetici):"
echo "  sudo -u $KULLANICI bash -c 'cd $DIZIN && .venv/bin/python kurulum_hazirla.py --profil data_profil.json --excel veri.xlsb --yonetici admin --yonetici-ad \"Ad Soyad\"'"
echo "  sudo systemctl start uretim-takip"
echo "  Adres: https://$ALAN"
