#!/usr/bin/env bash
# Automail Desk - one-click installer for macOS / Linux
# Usage: curl -fsSL https://raw.githubusercontent.com/o0OAdonisO0o/automail-desk/main/install.sh | bash
set -euo pipefail

REPO="o0OAdonisO0o/automail-desk"
DEST="${AUTOMAIL_HOME:-$HOME/AutomailDesk}"

echo "== Automail Desk installer =="
echo "target: $DEST"

PY="${PYTHON:-python3}"
if ! command -v "$PY" >/dev/null 2>&1; then
  echo "[ERROR] python3 not found. Install Python 3.9+ first." >&2
  exit 1
fi

TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT
echo "downloading https://github.com/$REPO/archive/refs/heads/main.tar.gz"
if ! curl -fsSL "https://github.com/$REPO/archive/refs/heads/main.tar.gz" -o "$TMP/src.tar.gz"; then
  echo "[ERROR] download failed. Check network or proxy settings." >&2
  exit 1
fi
tar -xzf "$TMP/src.tar.gz" -C "$TMP"
rm -rf "$DEST"
mkdir -p "$(dirname "$DEST")"
mv "$TMP"/automail-desk-* "$DEST"

cd "$DEST"
"$PY" -m venv .venv
. .venv/bin/activate
python -m pip install -q --upgrade pip
python -m pip install -q -r requirements.txt

if [ ! -f config.json ]; then
  cp config.example.json config.json
  echo ""
  echo "Generated $DEST/config.json"
  echo "Edit domain / password / account_prefix / account_count before first use."
fi

echo ""
echo "Done. Start with:"
echo "  cd $DEST && ./run.sh"
