"""
مذكرتي Pro v7 — Canva Level
3 محركات: Classic · Canva · Premium(Node)
"""
import os, sys, json, subprocess, tempfile, logging, io
from flask import Flask, request, send_file, jsonify, send_from_directory, make_response

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "scripts"))

app = Flask(__name__, static_folder="public", static_url_path="")
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

NODE_SCRIPT  = os.path.join(os.path.dirname(__file__), "node_scripts", "generator_api.js")
NODE_MODULES = os.path.join(os.path.dirname(__file__), "node_scripts", "node_modules")

CLASSIC_THEMES = {'navy_gold','dark_teal','burgundy','forest','midnight_purple','charcoal_orange','ice_blue','sand_gold'}
PREMIUM_THEMES = {'noir','atlas','sakura'}

@app.after_request
def cors(r):
    r.headers["Access-Control-Allow-Origin"]  = "*"
    r.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    r.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return r

@app.before_request
def preflight():
    if request.method == "OPTIONS":
        r = make_response("", 204)
        r.headers["Access-Control-Allow-Origin"]  = "*"
        r.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
        r.headers["Access-Control-Allow-Headers"] = "Content-Type"
        return r

@app.route("/")
def index():
    return send_from_directory("public", "index.html")

@app.route("/health")
def health():
    return jsonify({"status": "ok", "version": "7.0", "engines": ["canva", "classic", "premium"]}), 200

@app.route("/generate", methods=["POST"])
def generate():
    try:
        data = request.get_json(force=True, silent=True)
        if not data:
            return jsonify({"error": "بيانات غير صالحة"}), 400
        if not data.get("studentName") or not data.get("titleAr"):
            return jsonify({"error": "اسم الطالب وعنوان المذكرة مطلوبان"}), 400

        engine = data.get("engine", "canva")
        theme  = data.get("theme", "navy_gold")
        log.info(f"[{engine}] theme={theme} student={data.get('studentName','?')[:20]}")

        if engine == "premium" or theme in PREMIUM_THEMES:
            return _gen_premium(data)
        elif engine == "classic":
            return _gen_python(data, "generator_classic")
        else:  # canva (default)
            return _gen_python(data, "generator_canva")

    except Exception as e:
        log.error(f"Unexpected: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500


def _gen_python(data, module_name):
    try:
        mod = __import__(module_name)
        # reload to pick up any changes
        import importlib
        mod = importlib.import_module(module_name)

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            path = f.name
        mod.generate_presentation(data, path)
        with open(path, "rb") as f:
            pptx_bytes = f.read()
        os.unlink(path)

        name = data.get("studentName","مذكرة").replace(" ","_")
        return send_file(
            io.BytesIO(pptx_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=f"عرض_{name}.pptx",
        )
    except Exception as e:
        log.error(f"{module_name} error: {e}", exc_info=True)
        return jsonify({"error": f"خطأ في المحرك: {str(e)[:300]}"}), 500


def _gen_premium(data):
    try:
        env = os.environ.copy()
        env["NODE_PATH"] = NODE_MODULES
        result = subprocess.run(
            ["node", NODE_SCRIPT],
            input=json.dumps(data, ensure_ascii=False).encode("utf-8"),
            capture_output=True, timeout=90,
            cwd=os.path.join(os.path.dirname(__file__), "node_scripts"),
            env=env,
        )
        if result.returncode != 0:
            err = result.stderr.decode("utf-8", errors="replace")
            return jsonify({"error": f"خطأ في المحرك: {err[:300]}"}), 500
        pptx_bytes = result.stdout
        if len(pptx_bytes) < 1000:
            return jsonify({"error": "ملف فارغ من المحرك"}), 500
        name = data.get("studentName","مذكرة").replace(" ","_")
        return send_file(
            io.BytesIO(pptx_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=f"عرض_{name}.pptx",
        )
    except subprocess.TimeoutExpired:
        return jsonify({"error": "انتهت مهلة التوليد"}), 504
    except FileNotFoundError:
        return jsonify({"error": "Node.js غير مثبت على الخادم"}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port,
            debug=os.environ.get("FLASK_ENV") == "development")
