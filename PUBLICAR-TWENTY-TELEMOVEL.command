#!/bin/bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
BUILD_DIR="$(mktemp -d)"
trap 'rm -rf "$BUILD_DIR"' EXIT

printf '\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n'
printf '  Twenty · publicar para o telemóvel\n'
printf '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n'

if ! command -v npm >/dev/null 2>&1; then
  printf 'Erro: instala primeiro o Node.js em https://nodejs.org\n'
  read -r -p "Carrega Enter para fechar..."
  exit 1
fi

read -r -p "Nome do projeto Cloudflare Pages [twenty-study-os]: " PROJECT_NAME
PROJECT_NAME="${PROJECT_NAME:-twenty-study-os}"

printf '\nA preparar apenas os ficheiros públicos da PWA...\n'
rsync -a \
  --exclude '.backup-*' \
  --exclude 'cloudflare-worker' \
  --exclude '*.command' \
  --exclude 'GUIA-GIT-SYNC.md' \
  --exclude 'TWENTY-SYNC-CREDENCIAIS.txt' \
  --exclude '.DS_Store' \
  "$ROOT_DIR/" "$BUILD_DIR/"

printf '\nVai abrir o browser para iniciares sessão na Cloudflare, se necessário.\n'
npx --yes wrangler@latest login

printf '\nA publicar a PWA...\n'
DEPLOY_LOG="$(mktemp)"
npx --yes wrangler@latest pages deploy "$BUILD_DIR" --project-name "$PROJECT_NAME" 2>&1 | tee "$DEPLOY_LOG"
DEPLOY_URL="$(grep -Eo 'https://[^[:space:]]+\.pages\.dev' "$DEPLOY_LOG" | tail -1 || true)"
rm -f "$DEPLOY_LOG"
PUBLIC_ORIGIN="https://${PROJECT_NAME}.pages.dev"

if [ -f "$ROOT_DIR/cloudflare-worker/wrangler.jsonc" ]; then
  CONFIGURED="$(CONFIG_FILE="$ROOT_DIR/cloudflare-worker/wrangler.jsonc" python3 <<'PY2'
import json, os
with open(os.environ["CONFIG_FILE"], encoding="utf-8") as f:
    d = json.load(f)
print("yes" if d.get("vars", {}).get("GITHUB_OWNER") != "SUBSTITUI_PELO_TEU_USERNAME" else "no")
PY2
)"
  if [ "$CONFIGURED" = "yes" ]; then
    PUBLIC_ORIGIN="$PUBLIC_ORIGIN" CONFIG_FILE="$ROOT_DIR/cloudflare-worker/wrangler.jsonc" python3 <<'PY2'
import json, os
p = os.environ["CONFIG_FILE"]
with open(p, encoding="utf-8") as f:
    d = json.load(f)
origins = [x.strip() for x in d["vars"].get("ALLOWED_ORIGINS", "").split(",") if x.strip()]
origin = os.environ["PUBLIC_ORIGIN"]
if origin not in origins:
    origins.append(origin)
d["vars"]["ALLOWED_ORIGINS"] = ",".join(origins)
with open(p, "w", encoding="utf-8") as f:
    json.dump(d, f, ensure_ascii=False, indent=2)
    f.write("\n")
PY2
    printf '
A autorizar o endereço público no Worker de sincronização...
'
    cd "$ROOT_DIR/cloudflare-worker"
    npm install >/dev/null
    npx wrangler deploy
  fi
fi

printf '
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
'
printf '  Publicação concluída\n'
printf '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n'
printf 'URL: %s\n\n' "${DEPLOY_URL:-consulta o endereço mostrado acima}"
printf 'No telemóvel, abre esse endereço e instala a PWA.\n'
printf 'Depois vai a Admin & dados → Git como base de dados → Configurar.\n\n'
printf 'O endereço público foi autorizado automaticamente no Worker, quando este já estava configurado.\n'
read -r -p "Carrega Enter para fechar..."
