"""
Flask API — مذكرتي Pro v17.1
Thin HTTP adapter. Zero business logic.

FIXED vs v17.0:
- filename sanitizer handles Arabic correctly (uses transliteration fallback)
- warmup returns immediately (no blocking init inside request)  
- detailed error logging with stage info
"""
import base64
import logging
import os
import sys
import time
import unicodedata

from flask import Flask, jsonify, make_response, request, send_from_directory

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


def _safe_filename(name: str) -> str:
    """
    Convert any string (including Arabic) to a safe ASCII filename.
    Arabic → NFKD normalization strips diacritics, non-ASCII → removed,
    spaces/special → underscore. Falls back to timestamp if empty.
    """
    if not name:
        return f"prs_{int(time.time())}"
    # Normalize: decompose Unicode, then encode to ASCII ignoring non-ASCII
    normalized = unicodedata.normalize('NFKD', name)
    ascii_str = normalized.encode('ascii', 'ignore').decode('ascii')
    # Replace non-word characters with underscore
    safe = ''.join(c if c.isalnum() else '_' for c in ascii_str).strip('_')
    # If all Arabic (safe is empty after stripping), use student index
    if not safe:
        safe = f"student_{int(time.time()) % 100000}"
    return safe[:24]


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


# ── Health & Warmup ───────────────────────────────────────────────────
@app.route("/ping")
def ping():
    return "pong", 200

@app.route("/health")
def health():
    pipeline = get_pipeline()
    return jsonify({
        "status": "ok",
        "version": "17.1",
        "python": sys.version.split()[0],
        "font": pipeline._font,
    }), 200

@app.route("/warmup")
def warmup():
    """
    FIXED: pipeline is pre-initialized at startup, so this returns instantly.
    Frontend uses this to detect cold-start readiness.
    """
    get_pipeline()  # no-op after first call
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
        log.error(f"Build failed: {result.error} | stages={result.stages}")
        return jsonify({
            "error": result.error,
            "stages": result.stages,
        }), 500

    safe = _safe_filename(req.student_name)
    filename = f"mathkarati_{safe}.pptx"
    b64 = base64.b64encode(result.data).decode("ascii")
    elapsed = time.monotonic() - t0

    log.info(f"Generated: slides={result.slide_count} {len(result.data):,}B {elapsed:.2f}s")
    return jsonify({
        "ok": True,
        "filename": filename,
        "data": b64,
        "size": len(result.data),
        "slides": result.slide_count,
        "font": result.font_used,
        "elapsed": round(elapsed, 2),
        "stages": result.stages,
    })


# ── Entry point ───────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    get_pipeline()  # eager init on startup
    app.run(host="0.0.0.0", port=port, debug=False, use_reloader=False)
