#!/bin/bash
# ⚠️  لا تستخدم set -e — نريد البناء يكمل حتى عند أخطاء الخطوط الاختيارية

echo "==> [1/4] Installing system packages..."
apt-get update -qq 2>/dev/null && \
  apt-get install -y -qq fontconfig fonts-noto-core 2>/dev/null || \
  echo "WARNING: apt-get failed (may be normal on some platforms)"

echo "==> [2/4] Installing Cairo Arabic font..."
# استخدام مجلد writable دائماً ($HOME/.fonts يعمل على كل البيئات)
FONT_DIR="${HOME}/.fonts/cairo"
mkdir -p "$FONT_DIR" 2>/dev/null || {
  FONT_DIR="/tmp/fonts/cairo"
  mkdir -p "$FONT_DIR"
}

if fc-list 2>/dev/null | grep -qi "cairo"; then
  echo "    Cairo font already present: $(fc-list 2>/dev/null | grep -i cairo | head -1)"
else
  echo "    Downloading Cairo font..."
  FONT_URL="https://github.com/google/fonts/raw/main/ofl/cairo/Cairo%5Bslnt%2Cwght%5D.ttf"

  if curl -fsSL --max-time 30 "$FONT_URL" -o "$FONT_DIR/Cairo.ttf" 2>/dev/null; then
    echo "    Downloaded Cairo.ttf → $FONT_DIR"
  else
    echo "    Primary source failed, trying Google Fonts API..."
    curl -fsSL --max-time 30 \
      "https://fonts.googleapis.com/css2?family=Cairo&display=swap" \
      -o /dev/null 2>/dev/null || true
    # Fallback: Amiri font (Arabic, widely available)
    curl -fsSL --max-time 30 \
      "https://github.com/google/fonts/raw/main/ofl/amiri/Amiri-Regular.ttf" \
      -o "$FONT_DIR/Amiri-Regular.ttf" 2>/dev/null && \
      echo "    Amiri fallback installed" || \
      echo "    WARNING: font download failed — will use Calibri fallback"
  fi

  # Rebuild font cache (non-fatal)
  fc-cache -fv "$FONT_DIR" 2>/dev/null || true
  echo "    Font status: $(fc-list 2>/dev/null | grep -i 'cairo\|amiri' | head -1 || echo 'not found (Calibri will be used)')"
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
echo "    Amiri:  $(fc-list 2>/dev/null | grep -i amiri | wc -l) font files"
echo ""
echo "==> Build complete ✓"
