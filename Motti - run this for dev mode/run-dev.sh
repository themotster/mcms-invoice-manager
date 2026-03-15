#!/bin/bash
# Completely self-contained: kill this app's Electron, then start MCMS Invoice Manager in dev mode.
# Run from anywhere (or double-click run-dev.command)—no need to cd or type any commands.

# When launched by double-click, PATH may not include node/npm—load profile so npm is found
[[ -f "$HOME/.zprofile" ]] && source "$HOME/.zprofile" 2>/dev/null
[[ -f "$HOME/.zshrc" ]] && source "$HOME/.zshrc" 2>/dev/null
[[ -f "$HOME/.bash_profile" ]] && source "$HOME/.bash_profile" 2>/dev/null
export PATH="/usr/local/bin:/opt/homebrew/bin:$PATH"

# This folder lives inside the app repo; app root is one level up
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
APP_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

echo "MCMS Invoice Manager — dev mode"
echo "App folder: $APP_ROOT"
echo ""

if ! command -v npm >/dev/null 2>&1; then
  echo "ERROR: npm not found. Install Node.js (e.g. from nodejs.org or Homebrew) and try again."
  echo "Press Enter to close."
  read -r
  exit 1
fi

echo "Stopping any MCMS Electron processes from this project..."
for pid in $(pgrep -f "Electron" 2>/dev/null); do
  cwd=$(lsof -a -p "$pid" -d cwd 2>/dev/null | tail -1 | awk '{print $NF}')
  real_cwd=""
  [[ -n "$cwd" ]] && real_cwd="$(cd "$cwd" 2>/dev/null && pwd)"
  if [[ -n "$real_cwd" && ("$real_cwd" == "$APP_ROOT" || "$real_cwd" == "$APP_ROOT"/*) ]]; then
    kill "$pid" 2>/dev/null || true
  fi
done
sleep 1

echo "Starting dev mode (build + watch + Electron)..."
echo "Keep this window open. The app window will open when ready."
echo ""
cd "$APP_ROOT" || { echo "Could not cd to app folder."; read -r; exit 1; }

# Ensure local node binaries (webpack, etc.) are found when double-clicking
export PATH="$APP_ROOT/node_modules/.bin:$PATH"

# If node_modules missing or incomplete, install first
if [[ ! -f "$APP_ROOT/node_modules/.bin/webpack" ]]; then
  echo "Installing dependencies (first run or after clone)..."
  npm install || { echo "npm install failed."; read -r; exit 1; }
  echo ""
fi

npm run mcms:dev
code=$?
if [[ $code -ne 0 ]]; then
  echo ""
  echo "Dev mode exited with an error (code $code). Check the messages above."
  echo "Press Enter to close."
  read -r
fi
exit $code
