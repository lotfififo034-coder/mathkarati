"""
Flask API — مذكرتي Pro v17

Thin HTTP adapter. Contains ZERO business logic.
All generation logic lives in engine/pipeline.py.
"""
import base64
import logging
import os
import re
import sys
import time

from flask import Flask, jsonify, make_response, request, send_from_directory

# ── Setup Python path ─────────────────────────────────────────────────
_BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _BASE)

from core.models import PresentationRequest
from engine.pipeline import get_pipeline

app = Flask(__name__, static_folder="public", static_url_path="")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
    stream=sys.stdout,
)
log = logging.getLogger(__name__)


# ── CORS ──────────────────────────────────────────────────────────────
@app.after_request
def _cors(r):
    r.headers["Access-Control-Allow-Origin"] = "*"
    r.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    r.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return r


@app.before_request
def _preflight():
    if request.method == "OPTIONS":
        r = make_response("", 204)
        r.headers["Access-Control-Allow-Origin"] = "*"
        r.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
        r.headers["Access-Control-Allow-Headers"] = "Content-Type"
        return r


# ── Static ────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return send_from_directory("public", "index.html")


# ── Health ────────────────────────────────────────────────────────────
@app.route("/ping")
def ping():
    return "pong", 200


@app.route("/health")
def health():
    pipeline = get_pipeline()
    return jsonify({
        "status": "ok",
        "version": "17.0",
        "python": sys.version.split()[0],
        "engine": "PPTXExportPipeline",
        "font": pipeline._font,
    }), 200


@app.route("/warmup")
def warmup():
    """Non-blocking warmup — pipeline initializes on first call."""
    get_pipeline()  # ensure initialized
    return jsonify({"status": "ready", "modules_ready": True}), 200


# ── Generate ──────────────────────────────────────────────────────────
@app.route("/generate", methods=["POST"])
def generate():
    t0 = time.monotonic()

    raw = request.get_json(force=True, silent=True)
    if not raw:
        return jsonify({"error": "بيانات غير صالحة — أرسل JSON صحيح"}), 400

    req = PresentationRequest.from_dict(raw)
    errors = req.validate()
    if errors:
        return jsonify({"error": " | ".join(errors)}), 400

    pipeline = get_pipeline()
    result = pipeline.build(req)

    if not result.success:
        log.error(f"Build failed: {result.error}")
        return jsonify({"error": result.error}), 500

    # Build filename
    latin = re.sub(r"[^\w]", "_", req.student_name, flags=re.ASCII).strip("_")
    safe_name = latin[:20] if latin else f"prs_{int(time.time())}"
    filename = f"mathkarati_{safe_name}.pptx"

    # Encode as base64 for transport
    b64 = base64.b64encode(result.data).decode("ascii")

    elapsed = time.monotonic() - t0
    log.info(f"✅ Generated {result.slide_count} slides in {elapsed:.2f}s | {len(result.data):,} bytes")

    return jsonify({
        "ok": True,
        "filename": filename,
        "data": b64,
        "size": len(result.data),
        "slides": result.slide_count,
        "font": result.font_used,
        "elapsed": round(elapsed, 2),
    })


# ── Entry point ───────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_ENV") == "development"
    # Warm up on start
    get_pipeline()
    app.run(host="0.0.0.0", port=port, debug=debug, use_reloader=False)
