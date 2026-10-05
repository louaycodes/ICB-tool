#!/bin/bash
# ============================================================
# ICB Desktop — Quick launch script
# Double-click this file or run: ./launch.sh
# ============================================================

set -e

# Script directory (even if launched from elsewhere)
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "==========================================="
echo "  ICB Analysis — Desktop Launch"
echo "==========================================="
echo ""

# 1. Create the venv if it doesn't exist
if [ ! -d "venv" ]; then
    echo "[1/3] Creating virtual environment..."
    python3 -m venv venv
else
    echo "[1/3] Existing virtual environment detected."
fi

# 2. Activate the venv
source venv/bin/activate

# 3. Install dependencies
echo "[2/3] Checking dependencies..."
pip install -q -r requirements.txt

# 4. Launch the desktop application
echo "[3/3] Launching the application..."
echo ""
python3 icb_desktop.py
