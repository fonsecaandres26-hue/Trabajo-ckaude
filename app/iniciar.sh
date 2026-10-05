#!/usr/bin/env sh
# Abre Control Logístico (macOS / Linux).
cd "$(dirname "$0")"
python3 -c "import openpyxl" 2>/dev/null || python3 -m pip install --quiet openpyxl
exec python3 logistica.py "$@"
