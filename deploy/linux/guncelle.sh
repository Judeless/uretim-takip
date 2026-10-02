#!/usr/bin/env bash
# Kodu GitHub'dan güncelleyip servisi yeniden başlatır (veri dokunulmaz:
# uretim.db, data/ ve yedek/ git'in takibinde değil).
#    sudo bash /opt/uretim/deploy/linux/guncelle.sh
set -euo pipefail
DIZIN=/opt/uretim
/usr/local/bin/uretim-yedek || true                      # güncellemeden önce yedek
sudo -u uretim git -C "$DIZIN" fetch origin main
sudo -u uretim git -C "$DIZIN" reset --hard origin/main
sudo -u uretim "$DIZIN/.venv/bin/pip" install -q -r "$DIZIN/requirements.txt"
systemctl restart uretim-takip
sleep 3
systemctl --no-pager --lines=5 status uretim-takip
