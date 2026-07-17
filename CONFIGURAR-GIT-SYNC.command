#!/bin/bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
WORKER_DIR="$ROOT_DIR/cloudflare-worker"
CONFIG_FILE="$WORKER_DIR/wrangler.jsonc"
CREDENTIALS_FILE="$ROOT_DIR/TWENTY-SYNC-CREDENCIAIS.txt"

printf '\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n'
printf '  Twenty · Git Sync (Cloudflare + GitHub)\n'
printf '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n'
printf 'Antes de continuar, cria no GitHub:\n'
printf '  1. Um repositório PRIVADO (ex.: twenty-data)\n'
printf '  2. Um fine-grained token limitado a esse repositório\n'
printf '     com Contents: Read and write\n\n'

read -r -p "Username do GitHub: " GITHUB_OWNER
read -r -p "Nome do repositório [twenty-data]: " GITHUB_REPO
GITHUB_REPO="${GITHUB_REPO:-twenty-data}"
read -r -p "Branch [main]: " GITHUB_BRANCH
GITHUB_BRANCH="${GITHUB_BRANCH:-main}"
read -r -p "Origens permitidas, separadas por vírgulas [https://twenty.co]: " ALLOWED_ORIGINS
ALLOWED_ORIGINS="${ALLOWED_ORIGINS:-https://twenty.co}"

printf 'Token do GitHub (não será mostrado): '
read -r -s GITHUB_TOKEN
printf '\n'

SYNC_KEY="$(openssl rand -hex 24)"

GITHUB_OWNER="$GITHUB_OWNER" GITHUB_REPO="$GITHUB_REPO" GITHUB_BRANCH="$GITHUB_BRANCH" ALLOWED_ORIGINS="$ALLOWED_ORIGINS" CONFIG_FILE="$CONFIG_FILE" python3 <<'PY'
import json, os
path = os.environ["CONFIG_FILE"]
with open(path, "r", encoding="utf-8") as f:
    data = json.load(f)
data["vars"]["GITHUB_OWNER"] = os.environ["GITHUB_OWNER"]
data["vars"]["GITHUB_REPO"] = os.environ["GITHUB_REPO"]
data["vars"]["GITHUB_BRANCH"] = os.environ["GITHUB_BRANCH"]
data["vars"]["ALLOWED_ORIGINS"] = os.environ["ALLOWED_ORIGINS"]
with open(path, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
    f.write("\n")
PY

cd "$WORKER_DIR"

if ! command -v npm >/dev/null 2>&1; then
  printf '\nErro: precisas de instalar o Node.js antes de continuar.\n'
  printf 'Instala em https://nodejs.org e volta a executar este ficheiro.\n'
  read -r -p "Carrega Enter para fechar..."
  exit 1
fi

printf '\nA instalar o Wrangler...\n'
npm install

printf '\nVai abrir o browser para iniciares sessão na Cloudflare.\n'
npx wrangler login

printf '\nA guardar os segredos no Cloudflare Worker...\n'
printf '%s' "$GITHUB_TOKEN" | npx wrangler secret put GITHUB_TOKEN
printf '%s' "$SYNC_KEY" | npx wrangler secret put SYNC_KEY

printf '\nA publicar o Worker...\n'
DEPLOY_LOG="$(mktemp)"
npx wrangler deploy 2>&1 | tee "$DEPLOY_LOG"
WORKER_URL="$(grep -Eo 'https://[^[:space:]]+\.workers\.dev' "$DEPLOY_LOG" | tail -1 || true)"
rm -f "$DEPLOY_LOG"

cat > "$CREDENTIALS_FILE" <<CREDS
TWENTY · GIT SYNC

Worker: ${WORKER_URL:-COPIA_O_URL_MOSTRADO_NO_TERMINAL}
Chave: $SYNC_KEY
GitHub: https://github.com/$GITHUB_OWNER/$GITHUB_REPO

Na app:
Admin & dados → Git como base de dados → Configurar

No telemóvel tens de usar uma versão pública da app; twenty.co só existe localmente no Mac.
CREDS
chmod 600 "$CREDENTIALS_FILE"

printf '\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n'
printf '  Configuração terminada\n'
printf '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n'
printf 'Worker: %s\n' "${WORKER_URL:-consulta o URL acima}"
printf 'Chave:  %s\n\n' "$SYNC_KEY"
printf 'Também ficaram guardados em:\n%s\n\n' "$CREDENTIALS_FILE"
printf 'Agora abre a app → Admin & dados → Configurar.\n'
read -r -p "Carrega Enter para fechar..."
