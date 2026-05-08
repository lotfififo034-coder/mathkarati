#!/bin/bash

echo "==> مذكرتي Pro — Build Script v9.1 ✦"

# ── تثبيت مكتبات Python أولاً (الأهم) ──────────────────────────────
echo "==> Installing Python dependencies..."
pip install -r requirements.txt

# ── تثبيت خط Cairo في مجلد المشروع (لا يحتاج صلاحيات root) ──────────
echo "==> Installing Cairo font (user directory)..."
FONT_DIR="$HOME/.fonts/cairo"
mkdir -p "$FONT_DIR"

if ! fc-list 2>/dev/null | grep -qi "cairo"; then
    echo "==> Downloading Cairo font..."
    curl -fsSL "https://github.com/google/fonts/raw/main/ofl/cairo/Cairo%5Bslnt%2Cwght%5D.ttf" \
         -o "$FONT_DIR/Cairo.ttf" 2>/dev/null && \
    fc-cache -f "$FONT_DIR" 2>/dev/null && \
    echo "==> Cairo font installed ✓" || \
    echo "==> Cairo font download failed — will use Arial fallback"
else
    echo "==> Cairo font already available ✓"
fi

# ── تثبيت Node.js dependencies ──────────────────────────────────────
echo "==> Installing Node.js dependencies..."
cd node_scripts
npm install --production --silent && \
    echo "==> node_modules installed: $(ls node_modules 2>/dev/null | wc -l) packages ✓" || \
    echo "==> WARNING: npm install failed — Premium engine will fallback to Canva"
cd ..

# ── التحقق من Node.js ───────────────────────────────────────────────
if node --version 2>/dev/null; then
    echo "==> Node.js OK ✓"
else
    echo "==> WARNING: Node.js not found — Premium engine will use Canva fallback"
fi

echo "==> Build complete ✓"
