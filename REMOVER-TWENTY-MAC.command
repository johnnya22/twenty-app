#!/bin/bash
set -euo pipefail

DOMAIN="twenty.co"
LABEL="com.twenty.study-os"
BASE_DIR="/Library/Application Support/Twenty"
CADDY_BIN="$BASE_DIR/bin/caddy"
CADDYFILE="$BASE_DIR/Caddyfile"
DATA_DIR="$BASE_DIR/caddy-data"
CONFIG_DIR="$BASE_DIR/caddy-config"
PLIST="/Library/LaunchDaemons/${LABEL}.plist"
HOSTS_TMP="/tmp/twenty-hosts-remove.$$"

cleanup() { rm -f "$HOSTS_TMP"; }
trap cleanup EXIT

echo ""
echo "A remover o domínio local e o arranque automático da Twenty..."
sudo -v

sudo launchctl bootout system "$PLIST" 2>/dev/null || true
sudo launchctl unload -w "$PLIST" 2>/dev/null || true

if [[ -x "$CADDY_BIN" ]]; then
  sudo env \
    HOME=/var/root \
    XDG_DATA_HOME="$DATA_DIR" \
    XDG_CONFIG_HOME="$CONFIG_DIR" \
    "$CADDY_BIN" untrust --config "$CADDYFILE" --adapter caddyfile >/dev/null 2>&1 || true
fi

sudo rm -f "$PLIST"

sudo awk '
  $0 !~ /(^|[[:space:]])twenty\.co([[:space:]]|$)/ &&
  $0 !~ /(^|[[:space:]])www\.twenty\.co([[:space:]]|$)/ &&
  $0 !~ /^# Twenty Study OS — domínio local$/
' /etc/hosts > "$HOSTS_TMP"
sudo cp "$HOSTS_TMP" /etc/hosts
sudo chown root:wheel /etc/hosts
sudo chmod 644 /etc/hosts

sudo rm -rf "$BASE_DIR"
sudo rm -rf /Library/Logs/Twenty
sudo dscacheutil -flushcache || true
sudo killall -HUP mDNSResponder 2>/dev/null || true

echo ""
echo "✅ Twenty removida do arranque automático."
echo "O endereço twenty.co voltou ao funcionamento normal neste Mac."
echo "A pasta da app e os teus dados do navegador não foram apagados."
echo ""
read -r -p "Carrega Enter para fechar..."
