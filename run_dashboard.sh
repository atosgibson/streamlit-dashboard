#!/usr/bin/env bash
set -e

APP_DIR="/home/laptop-cer-03/Desktop/MJ eng/mj-dashboard"
LOG_DIR="$APP_DIR/05_reports/logs"
PORT=8501

mkdir -p "$LOG_DIR"
LOG_FILE="$LOG_DIR/streamlit_$(date +%Y%m%d_%H%M%S).log"

cd "$APP_DIR"

# habilita conda no script
source "/home/laptop-cer-03/miniconda3/etc/profile.d/conda.sh"
conda activate atos

# sobe o streamlit em background (salvando log)
nohup python -m streamlit run 03_app/streamlit/app.py --server.port $PORT --server.headless true \
  > "$LOG_FILE" 2>&1 &

# abre no navegador
sleep 2
xdg-open "http://localhost:$PORT" >/dev/null 2>&1 || true
