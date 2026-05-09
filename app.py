"""
مذكرتي Pro — خادم Flask للإنتاج
Production server — Render deployment ready
"""
import os, sys, json, tempfile, logging, io
from flask import Flask, request, send_file, jsonify, send_from_directory, make_response

# Add scripts dir to path so we can import generator directly
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "scripts"))
from generator import generate_presentation

# ── App setup ──────────────────────────────────────────
app = Flask(__name__, static_folder="public", static_url_path="")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
log = logging.getLogger(__name__)

# ── CORS middleware ─────────────────────────────────────
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

# ── Routes ─────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory("public", "index.html")

@app.route("/health")
def health():
    return jsonify({"status": "ok", "service": "مذكرتي Pro"}), 200

@app.route("/generate", methods=["POST"])
def generate():
    try:
        data = request.get_json(force=True, silent=True)
        if not data:
            return jsonify({"error": "بيانات غير صالحة"}), 400

        # Validate required fields
        if not data.get("studentName") or not data.get("titleAr"):
            return jsonify({"error": "اسم الطالب وعنوان المذكرة مطلوبان"}), 400

        log.info(f"Generating for: {data.get('studentName','unknown')}")

        # Generate in temp dir
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, "presentation.pptx")
            generate_presentation(data, out_path)

            student = data.get("studentName", "مذكرة").replace(" ", "_")
            filename = f"عرض_{student}.pptx"

            # Read file before temp dir is deleted
            with open(out_path, "rb") as f:
                pptx_bytes = f.read()

        # Write to BytesIO for send_file
        return send_file(
            io.BytesIO(pptx_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        log.error(f"Generation error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500


# ── Entry point ─────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_ENV") == "development"
    log.info(f"Starting مذكرتي Pro on port {port}")
    app.run(host="0.0.0.0", port=port, debug=debug)
