#!/usr/bin/env bash
set -euo pipefail

# ==============================
# Konfiguration (anpassen!)
# ==============================
REPO_DIR="${1:-.}"                       # Optional: Pfad zum Repo als 1. Argument
REMOTE_URL="${2:-}"                      # Optional: Remote URL als 2. Argument
TARGET_BRANCH="${3:-$(git -C "$REPO_DIR" branch --show-current)}"

# Beispiel-Aufruf:
# ./setup_and_push.sh /pfad/zum/repo https://github.com/OWNER/REPO.git work
# oder:
# ./setup_and_push.sh . git@github.com:OWNER/REPO.git work

if [[ -z "${REMOTE_URL}" ]]; then
  echo "❌ Bitte Remote-URL übergeben."
  echo "Beispiel HTTPS: ./setup_and_push.sh . https://github.com/OWNER/REPO.git work"
  echo "Beispiel SSH:   ./setup_and_push.sh . git@github.com:OWNER/REPO.git work"
  exit 1
fi

echo "📁 Repo: $REPO_DIR"
echo "🌐 Remote: $REMOTE_URL"
echo "🌿 Branch: $TARGET_BRANCH"

# ==============================
# Checks
# ==============================
if [[ ! -d "$REPO_DIR/.git" ]]; then
  echo "❌ Kein Git-Repo unter: $REPO_DIR"
  exit 1
fi

cd "$REPO_DIR"

# ==============================
# 1) Cleanup + .gitignore
# ==============================
echo "🧹 Entferne Python-Cache..."
find . -type d -name "__pycache__" -prune -exec rm -rf {} +
find . -type f -name "*.pyc" -delete

if [[ ! -f .gitignore ]]; then
  touch .gitignore
fi

append_if_missing() {
  local line="$1"
  if ! grep -qxF "$line" .gitignore; then
    echo "$line" >> .gitignore
  fi
}

append_if_missing "__pycache__/"
append_if_missing "*.pyc"

# ==============================
# 2) Optional committen
# ==============================
if [[ -n "$(git status --porcelain)" ]]; then
  echo "📝 Änderungen gefunden – committe..."
  git add -A
  git commit -m "chore: cleanup cache and prepare repo for remote push"
else
  echo "ℹ️ Keine lokalen Änderungen zum Committen."
fi

# ==============================
# 3) Remote setzen/aktualisieren
# ==============================
if git remote get-url origin >/dev/null 2>&1; then
  echo "🔁 Remote 'origin' existiert – aktualisiere URL..."
  git remote set-url origin "$REMOTE_URL"
else
  echo "➕ Lege Remote 'origin' an..."
  git remote add origin "$REMOTE_URL"
fi

echo "🔍 Remotes:"
git remote -v

# ==============================
# 4) Branch pushen
# ==============================
echo "🚀 Push: origin/$TARGET_BRANCH"
git push -u origin "$TARGET_BRANCH"

echo "✅ Fertig. Branch '$TARGET_BRANCH' ist auf GitHub."
