#!/bin/bash
# ============================================================
# ICB Desktop — Script de lancement rapide
# Double-cliquez sur ce fichier ou executez : ./launch.sh
# ============================================================

set -e

# Repertoire du script (meme si lance depuis ailleurs)
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "==========================================="
echo "  ICB Analysis — Lancement Desktop"
echo "==========================================="
echo ""

# 1. Creer le venv s'il n'existe pas
if [ ! -d "venv" ]; then
    echo "[1/3] Creation de l'environnement virtuel..."
    python3 -m venv venv
else
    echo "[1/3] Environnement virtuel existant detecte."
fi

# 2. Activer le venv
source venv/bin/activate

# 3. Installer les dependances
echo "[2/3] Verification des dependances..."
pip install -q -r requirements.txt

# 4. Lancer l'application desktop
echo "[3/3] Lancement de l'application..."
echo ""
python3 icb_desktop.py
