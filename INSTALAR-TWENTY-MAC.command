#!/bin/bash
set -euo pipefail

DOMAIN="twenty.co"
WWW_DOMAIN="www.twenty.co"
LABEL="com.twenty.study-os"
APP_DIR="$(cd "$(dirname "$0")" && pwd)"
BASE_DIR="/Library/Application Support/Twenty"
BIN_DIR="$BASE_DIR/bin"
CADDY_BIN="$BIN_DIR/caddy"
CADDYFILE="$BASE_DIR/Caddyfile"
DATA_DIR="$BASE_DIR/caddy-data"
CONFIG_DIR="$BASE_DIR/caddy-config"
LOG_DIR="/Library/Logs/Twenty"
PLIST="/Library/LaunchDaemons/${LABEL}.plist"
HOSTS_FILE="/etc/hosts"
HOSTS_TMP="/tmp/twenty-hosts.$$"
PLIST_TMP="/tmp/${LABEL}.plist.$$"
CADDY_TMP="/tmp/twenty-caddy.$$"

cleanup() {
  rm -f "$HOSTS_TMP" "$PLIST_TMP" "$CADDY_TMP" /tmp/twenty-caddy-release.json /tmp/twenty-caddy.tar.gz
}
trap cleanup EXIT

echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "  Twenty · instalação automática no Mac"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
echo "A app será servida diretamente desta pasta:"
echo "$APP_DIR"
echo ""
echo "Depois da instalação, não movas nem apagues esta pasta."
echo "Podes correr novamente este instalador quando quiseres atualizar a configuração."
echo ""

if [[ "$(uname -s)" != "Darwin" ]]; then
  echo "Este instalador foi feito apenas para macOS."
  read -r -p "Carrega Enter para fechar..."
  exit 1
fi

if [[ ! -f "$APP_DIR/index.html" ]]; then
  echo "Erro: não encontrei index.html na pasta da app."
  read -r -p "Carrega Enter para fechar..."
  exit 1
fi

if ! command -v python3 >/dev/null 2>&1; then
  echo "Erro: o Python 3 é necessário apenas para descarregar o Caddy."
  read -r -p "Carrega Enter para fechar..."
  exit 1
fi

if ! command -v curl >/dev/null 2>&1; then
  echo "Erro: não encontrei o curl no macOS."
  read -r -p "Carrega Enter para fechar..."
  exit 1
fi

echo "O macOS vai pedir a tua palavra-passe para instalar o serviço local."
sudo -v

# Mantém a autorização sudo ativa durante a instalação.
while true; do
  sudo -n true
  sleep 50
  kill -0 "$$" 2>/dev/null || exit
 done 2>/dev/null &
SUDO_KEEPALIVE_PID=$!
trap 'kill "$SUDO_KEEPALIVE_PID" 2>/dev/null || true; cleanup' EXIT

sudo mkdir -p "$BIN_DIR" "$DATA_DIR" "$CONFIG_DIR" "$LOG_DIR"

# Usa um Caddy já instalado; caso contrário descarrega a release oficial para esta instalação.
if command -v caddy >/dev/null 2>&1; then
  EXISTING_CADDY="$(command -v caddy)"
  echo "A usar Caddy existente: $EXISTING_CADDY"
  sudo cp "$EXISTING_CADDY" "$CADDY_BIN"
else
  case "$(uname -m)" in
    arm64) CADDY_ARCH="arm64" ;;
    x86_64) CADDY_ARCH="amd64" ;;
    *)
      echo "Arquitetura do Mac não suportada: $(uname -m)"
      read -r -p "Carrega Enter para fechar..."
      exit 1
      ;;
  esac

  echo "A descarregar o Caddy oficial..."
  curl -fsSL \
    -H "Accept: application/vnd.github+json" \
    -H "X-GitHub-Api-Version: 2022-11-28" \
    "https://api.github.com/repos/caddyserver/caddy/releases/latest" \
    -o /tmp/twenty-caddy-release.json

  CADDY_URL="$(python3 - "$CADDY_ARCH" <<'PY'
import json
import sys

arch = sys.argv[1].lower()
with open('/tmp/twenty-caddy-release.json', 'r', encoding='utf-8') as handle:
    release = json.load(handle)

candidates = []
for asset in release.get('assets', []):
    name = asset.get('name', '').lower()
    is_macos = 'mac' in name or 'darwin' in name
    is_arch = arch in name
    is_archive = name.endswith('.tar.gz') or name.endswith('.tgz')
    if is_macos and is_arch and is_archive:
        candidates.append(asset.get('browser_download_url', ''))

print(next((url for url in candidates if url), ''))
PY
)"

  if [[ -z "$CADDY_URL" ]]; then
    echo "Não consegui localizar automaticamente a versão do Caddy para este Mac."
    echo "Instala o Caddy e volta a executar este ficheiro."
    read -r -p "Carrega Enter para fechar..."
    exit 1
  fi

  curl -fL "$CADDY_URL" -o /tmp/twenty-caddy.tar.gz
  mkdir -p "$CADDY_TMP"
  tar -xzf /tmp/twenty-caddy.tar.gz -C "$CADDY_TMP"

  if [[ ! -f "$CADDY_TMP/caddy" ]]; then
    echo "O arquivo descarregado não continha o executável Caddy."
    read -r -p "Carrega Enter para fechar..."
    exit 1
  fi

  sudo cp "$CADDY_TMP/caddy" "$CADDY_BIN"
fi

sudo chown root:wheel "$CADDY_BIN"
sudo chmod 755 "$CADDY_BIN"

# Escapa aspas e barras invertidas para usar o caminho da pasta dentro do Caddyfile.
CADDY_APP_DIR="${APP_DIR//\\/\\\\}"
CADDY_APP_DIR="${CADDY_APP_DIR//\"/\\\"}"

cat > /tmp/twenty-Caddyfile <<EOF_CADDY
{
    admin 127.0.0.1:2020
}

$DOMAIN, $WWW_DOMAIN {
    tls internal
    root * "$CADDY_APP_DIR"
    encode gzip
    file_server

    header {
        X-Content-Type-Options nosniff
        Referrer-Policy no-referrer
    }
}
EOF_CADDY
sudo mv /tmp/twenty-Caddyfile "$CADDYFILE"
sudo chown root:wheel "$CADDYFILE"
sudo chmod 644 "$CADDYFILE"

# Valida a configuração antes de instalar o serviço.
sudo env \
  HOME=/var/root \
  XDG_DATA_HOME="$DATA_DIR" \
  XDG_CONFIG_HOME="$CONFIG_DIR" \
  "$CADDY_BIN" validate --config "$CADDYFILE" --adapter caddyfile

# Remove entradas antigas dos dois domínios e adiciona o redirecionamento local.
sudo awk '
  $0 !~ /(^|[[:space:]])twenty\.co([[:space:]]|$)/ &&
  $0 !~ /(^|[[:space:]])www\.twenty\.co([[:space:]]|$)/
' "$HOSTS_FILE" > "$HOSTS_TMP"
printf '\n# Twenty Study OS — domínio local\n127.0.0.1 twenty.co www.twenty.co\n' >> "$HOSTS_TMP"
sudo cp "$HOSTS_TMP" "$HOSTS_FILE"
sudo chown root:wheel "$HOSTS_FILE"
sudo chmod 644 "$HOSTS_FILE"

cat > "$PLIST_TMP" <<EOF_PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>$LABEL</string>

    <key>ProgramArguments</key>
    <array>
        <string>$CADDY_BIN</string>
        <string>run</string>
        <string>--config</string>
        <string>$CADDYFILE</string>
        <string>--adapter</string>
        <string>caddyfile</string>
    </array>

    <key>WorkingDirectory</key>
    <string>$BASE_DIR</string>

    <key>EnvironmentVariables</key>
    <dict>
        <key>HOME</key>
        <string>/var/root</string>
        <key>XDG_DATA_HOME</key>
        <string>$DATA_DIR</string>
        <key>XDG_CONFIG_HOME</key>
        <string>$CONFIG_DIR</string>
    </dict>

    <key>RunAtLoad</key>
    <true/>
    <key>KeepAlive</key>
    <true/>
    <key>ThrottleInterval</key>
    <integer>5</integer>

    <key>StandardOutPath</key>
    <string>$LOG_DIR/caddy.log</string>
    <key>StandardErrorPath</key>
    <string>$LOG_DIR/caddy-error.log</string>
</dict>
</plist>
EOF_PLIST

sudo cp "$PLIST_TMP" "$PLIST"
sudo chown root:wheel "$PLIST"
sudo chmod 644 "$PLIST"
sudo plutil -lint "$PLIST"

# Reinicia ou instala o LaunchDaemon.
sudo launchctl bootout system "$PLIST" 2>/dev/null || true
if ! sudo launchctl bootstrap system "$PLIST"; then
  echo "O método moderno do launchd falhou; a tentar o método compatível..."
  sudo launchctl load -w "$PLIST"
fi
sudo launchctl enable "system/$LABEL" 2>/dev/null || true
sudo launchctl kickstart -k "system/$LABEL" 2>/dev/null || true

sudo dscacheutil -flushcache || true
sudo killall -HUP mDNSResponder 2>/dev/null || true

# Espera que o servidor gere a autoridade local e fique disponível.
echo "A iniciar o serviço local..."
READY=0
for ((i = 1; i <= 30; i++)); do
  if curl -kfsS --resolve "$DOMAIN:443:127.0.0.1" "https://$DOMAIN/" >/dev/null 2>&1; then
    READY=1
    break
  fi
  sleep 1
done

if [[ "$READY" -ne 1 ]]; then
  echo ""
  echo "O serviço foi instalado, mas ainda não respondeu."
  echo "Consulta o erro em: $LOG_DIR/caddy-error.log"
  echo ""
  read -r -p "Carrega Enter para fechar..."
  exit 1
fi

# Garante que o certificado local é reconhecido pelo macOS.
sudo env \
  HOME=/var/root \
  XDG_DATA_HOME="$DATA_DIR" \
  XDG_CONFIG_HOME="$CONFIG_DIR" \
  "$CADDY_BIN" trust --address 127.0.0.1:2020 >/dev/null 2>&1 || true

sudo dscacheutil -flushcache || true
sudo killall -HUP mDNSResponder 2>/dev/null || true

open "https://$DOMAIN"

echo ""
echo "✅ Twenty está instalada."
echo ""
echo "Abre sempre: https://$DOMAIN"
echo "Arranca automaticamente com o Mac."
echo "Não precisas de executar python3 start-server.py."
echo ""
echo "Importante: enquanto estiver instalada, o verdadeiro site twenty.co"
echo "fica substituído por esta app apenas neste Mac."
echo ""
read -r -p "Carrega Enter para fechar esta janela..."
