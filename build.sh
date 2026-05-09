#!/bin/bash
set -e

echo "==> Installing system fonts (Cairo Arabic)..."
apt-get update -qq && apt-get install -y -qq fonts-urw-base35 fonts-noto fonts-noto-core 2>/dev/null || true
# Try to install Cairo font via pip package that bundles it
pip install --quiet cairocffi 2>/dev/null || true
# Download Cairo font directly if not available
FONT_DIR="/usr/local/share/fonts/cairo"
if ! fc-list | grep -qi "cairo"; then
  mkdir -p "$FONT_DIR"
  echo "==> Downloading Cairo font..."
  curl -sL "https://github.com/google/fonts/raw/main/ofl/cairo/Cairo%5Bslnt%2Cwght%5D.ttf" \
       -o "$FONT_DIR/Cairo.ttf" 2>/dev/null || \
  curl -sL "https://fonts.gstatic.com/s/cairo/v28/SLXgc1nY6HkvalIvTp0zQg.woff2" \
       -o "$FONT_DIR/Cairo.ttf" 2>/dev/null || true
  fc-cache -fv "$FONT_DIR" 2>/dev/null || true
  echo "==> Cairo font installed: $(fc-list | grep -i cairo | head -1)"
else
  echo "==> Cairo font already available: $(fc-list | grep -i cairo | head -1)"
fi

echo "==> Installing Python dependencies..."
pip install -r requirements.txt

echo "==> Installing Node.js dependencies..."
cd node_scripts && npm install --production && cd ..

echo "==> Verifying Node.js..."
node --version && echo "Node.js OK" || echo "WARNING: Node.js not found — Premium engine will fallback to Canva"

echo "==> Build complete ✓"
