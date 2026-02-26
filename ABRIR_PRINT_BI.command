#!/bin/bash
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
APP_PATH="$PROJECT_DIR/PRINT_BI.app"
SCRIPT_PATH="$PROJECT_DIR/launcher/PRINT_BI_Launcher.js"

if ! command -v osacompile >/dev/null 2>&1; then
  osascript -e 'display dialog "osacompile não encontrado no macOS." buttons {"OK"} default button "OK" with title "PRINT BI"'
  exit 1
fi

if [ ! -f "$SCRIPT_PATH" ]; then
  osascript -e 'display dialog "Script do launcher não encontrado." buttons {"OK"} default button "OK" with title "PRINT BI"'
  exit 1
fi

rm -rf "$APP_PATH"
osacompile -l JavaScript -o "$APP_PATH" "$SCRIPT_PATH"
open "$APP_PATH"
