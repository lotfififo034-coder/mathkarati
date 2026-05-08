#!/bin/bash
set -e

echo "==> مذكرتي Pro — Build Script v9 ✦"
echo "==> Installing system fonts (Noto + Cairo Arabic)..."

# تثبيت الخطوط عبر apt (الطريقة الأموثق لدى Render)
apt-get update -qq 2>/dev/null || true
apt-get install -y -qq fonts-noto fonts-noto-core fonts-noto-extra 2>/dev/null || true

# تحميل خط Cairo مباشرة من Google Fonts
FONT_DIR="/usr/local/share/fonts/cairo"
if ! fc-list 2>/dev/null | grep -qi "cairo"; then
  mkdir -p "$FONT_DIR"
  echo "==> Downloading Cairo font..."
  curl -fsSL "https://github.com/google/fonts/raw/main/ofl/cairo/Cairo%5Bslnt%2Cwght%5D.ttf" \
       -o "$FONT_DIR/Cairo.ttf" 2>/dev/null \
  || curl -fsSL "https://fonts.gstatic.com/s/cairo/v28/SLXgc1nY6HkvalIvTp0zQg.woff2" \
       -o "$FONT_DIR/Cairo.woff2" 2>/dev/null || true
  fc-cache -fv "$FONT_DIR" 2>/dev/null || true
  echo "==> Cairo font status: $(fc-list 2>/dev/null | grep -i cairo | head -1 || echo 'not found — will fallback to Arial')"
else
  echo "==> Cairo font already available."
fi

echo "==> Installing Python dependencies..."
pip install -r requirements.txt

echo "==> Installing Node.js dependencies..."
cd node_scripts
npm install --production --silent
echo "==> node_modules installed: $(ls node_modules | wc -l) packages"
cd ..

echo "==> Verifying Node.js..."
if node --version; then
  echo "==> Node.js OK ✓"
else
  echo "==> WARNING: Node.js not found — Premium engine will fallback to Canva"
fi

echo "==> Build complete ✓"
