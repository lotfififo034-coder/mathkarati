#!/bin/bash
set -e

echo "==> Installing Python dependencies..."
pip install -r requirements.txt

echo "==> Installing Node.js dependencies..."
cd node_scripts && npm install --production && cd ..

echo "==> Build complete ✓"
echo "    Note: Cairo font will fall back to Arial (no system font install needed)"
