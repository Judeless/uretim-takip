#!/usr/bin/env bash
# Günlük veritabanı yedeği (cron 02:30). SQLite'ın kendi .backup komutu: uygulama
# çalışırken de tutarlı kopya alır. 30 günden eski yedekler silinir.
set -euo pipefail
DIZIN=/opt/uretim
HEDEF="$DIZIN/yedek"
mkdir -p "$HEDEF"
sqlite3 "$DIZIN/uretim.db" ".backup '$HEDEF/uretim_$(date +%F).db'"
cp -f "$DIZIN/data/kurulum.json" "$HEDEF/kurulum.json" 2>/dev/null || true
find "$HEDEF" -name 'uretim_*.db' -mtime +30 -delete
