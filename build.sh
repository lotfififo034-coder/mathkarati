#!/bin/bash
# مذكرتي Pro — Render Build Script
# لا يحتاج صلاحيات النظام — يعمل على Render Free Plan

echo "==> Installing Python dependencies..."
pip install -r requirements.txt
echo "Python deps OK"

echo "==> Installing Node.js dependencies..."
cd node_scripts
npm install --production
cd ..
echo "Node deps OK"

echo "==> Build complete ✓"
