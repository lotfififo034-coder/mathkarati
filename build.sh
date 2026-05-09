#!/bin/bash
set -e

echo "==> [1/4] Installing system packages..."
apt-get update -qq 2>/dev/null &&   apt-get install -y -qq fontconfig fonts-noto-core 2>/dev/null ||   echo "WARNING: apt-get failed (may be normal on some platforms)"

echo "==> [2/4] Installing Cairo Arabic font..."
FONT_DIR="/usr/local/share/fonts/cairo"
mkdir -p "$FONT_DIR"

if fc-list 2>/dev/null | grep -qi "cairo"; then
  echo "    Cairo font already present: $(fc-list 2>/dev/null | grep -i cairo | head -1)"
else
  echo "    Downloading Cairo font..."
  # Primary source: Google Fonts GitHub
  FONT_URL="https://github.com/google/fonts/raw/main/ofl/cairo/Cairo%5Bslnt%2Cwght%5D.ttf"
  FONT_FALLBACK1="https://github.com/alif-type/cairo/releases/download/1.004/Cairo-1.004.zip"
  
  if curl -fsSL --max-time 30 "$FONT_URL" -o "$FONT_DIR/Cairo.ttf" 2>/dev/null; then
    echo "    Downloaded from Google Fonts GitHub"
  else
    echo "    Primary source failed, trying fallback..."
    # Fallback: direct woff2 (will work as ttf placeholder)
    curl -fsSL --max-time 30       "https://fonts.gstatic.com/s/cairo/v28/SLXgc1nY6HkvalIvTp0zQg.woff2"       -o "$FONT_DIR/Cairo.woff2" 2>/dev/null ||     echo "    WARNING: Cairo font download failed — will use Calibri fallback"
  fi
  
  # Rebuild font cache
  fc-cache -fv "$FONT_DIR" 2>/dev/null || true
  echo "    Font status: $(fc-list 2>/dev/null | grep -i cairo | head -1 || echo 'not found (Calibri will be used)')"
fi

echo "==> [3/4] Installing Python dependencies..."
pip install --no-cache-dir -r requirements.txt

echo "==> [4/4] Installing Node.js dependencies..."
cd node_scripts
if npm install --production --no-audit --no-fund 2>/dev/null; then
  echo "    Node modules installed OK"
else
  echo "    WARNING: npm install failed — Premium engine will fallback to Canva"
fi
cd ..

echo ""
echo "    Python: $(python3 --version)"
echo "    Node:   $(node --version 2>/dev/null || echo 'not found')"
echo "    Cairo:  $(fc-list 2>/dev/null | grep -i cairo | wc -l) font files"
echo ""
echo "==> Build complete ✓"
