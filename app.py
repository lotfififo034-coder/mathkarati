"""
مذكرتي Pro v3 — Flask + Node.js Hybrid
Design Engine: MathKarati PRO v3 (PptxGenJS)
Server: Flask + Gunicorn
Deploy: Render.com ready
"""
import os, sys, json, subprocess, logging, io
from flask import Flask, request, send_file, jsonify, send_from_directory, make_response

app = Flask(__name__, static_folder="public", static_url_path="")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
log = logging.getLogger(__name__)

# Path to the Node.js API generator
GENERATOR = os.path.join(os.path.dirname(__file__), "node_scripts", "generator_api.js")
NODE_MODULES = os.path.join(os.path.dirname(__file__), "node_scripts", "node_modules")


# ── CORS ────────────────────────────────────────────────────────────
@app.after_request
def add_cors(response):
    response.headers["Access-Control-Allow-Origin"]  = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return response

@app.before_request
def handle_options():
    if request.method == "OPTIONS":
        resp = make_response("", 204)
        resp.headers["Access-Control-Allow-Origin"]  = "*"
        resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
        resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
        return resp


# ── ROUTES ──────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory("public", "index.html")

@app.route("/health")
def health():
    return jsonify({"status": "ok", "service": "مذكرتي Pro v3", "engine": "PptxGenJS v3"}), 200

@app.route("/generate", methods=["POST"])
def generate():
    try:
        data = request.get_json(force=True, silent=True)
        if not data:
            return jsonify({"error": "بيانات غير صالحة"}), 400

        if not data.get("studentName") or not data.get("titleAr"):
            return jsonify({"error": "اسم الطالب وعنوان المذكرة مطلوبان"}), 400

        log.info(f"Generating for: {data.get('studentName','unknown')} | theme: {data.get('theme','noir')}")

        # Call Node.js generator
        env = os.environ.copy()
        env["NODE_PATH"] = NODE_MODULES

        result = subprocess.run(
            ["node", GENERATOR],
            input=json.dumps(data, ensure_ascii=False).encode("utf-8"),
            capture_output=True,
            timeout=90,
            cwd=os.path.join(os.path.dirname(__file__), "node_scripts"),
            env=env,
        )

        if result.returncode != 0:
            err = result.stderr.decode("utf-8", errors="replace")
            log.error(f"Node.js error: {err}")
            return jsonify({"error": f"خطأ في المحرك: {err[:300]}"}), 500

        pptx_bytes = result.stdout
        if len(pptx_bytes) < 1000:
            err = result.stderr.decode("utf-8", errors="replace")
            log.error(f"Empty output. stderr: {err}")
            return jsonify({"error": "الملف فارغ — تحقق من المحرك"}), 500

        log.info(f"Generated {len(pptx_bytes):,} bytes. {result.stderr.decode('utf-8','replace').strip()}")

        student = data.get("studentName", "مذكرة").replace(" ", "_")
        filename = f"عرض_{student}.pptx"

        return send_file(
            io.BytesIO(pptx_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=filename,
        )

    except subprocess.TimeoutExpired:
        log.error("Node.js timeout after 90s")
        return jsonify({"error": "انتهت مهلة التوليد (90 ثانية)"}), 504

    except FileNotFoundError:
        log.error("node not found — is Node.js installed?")
        return jsonify({"error": "Node.js غير مثبت على الخادم"}), 500

    except Exception as e:
        log.error(f"Unexpected error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500


# ── ENTRY POINT ─────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_ENV") == "development"
    log.info(f"Starting مذكرتي Pro v3 on port {port}")
    app.run(host="0.0.0.0", port=port, debug=debug)
