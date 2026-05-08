"""
مذكرتي Pro v9 — FIXED
3 محركات: Classic · Canva · Premium(Node)
إصلاحات v9:
- تسجيل أخطاء تفصيلي للـ traceback الكامل
- إصلاح مشكلة PORT على Render
- timeout أطول للاتصالات
- إصلاح MIME type لـ favicon
- health endpoint مُحسَّن
"""
import os, sys, json, subprocess, shutil, tempfile, logging, io, importlib, traceback
from flask import Flask, request, send_file, jsonify, send_from_directory, make_response

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "scripts"))

app = Flask(__name__, static_folder=None)

# ── Logging مُحسَّن ───────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    stream=sys.stdout,   # مهم: stdout وليس stderr حتى يظهر في Render logs
)
log = logging.getLogger(__name__)

NODE_SCRIPT  = os.path.join(os.path.dirname(__file__), "node_scripts", "generator_api.js")
NODE_MODULES = os.path.join(os.path.dirname(__file__), "node_scripts", "node_modules")

CLASSIC_THEMES = {'navy_gold','dark_teal','burgundy','forest','midnight_purple','charcoal_orange','ice_blue','sand_gold'}
PREMIUM_THEMES = {'noir','atlas','sakura'}

# ── فحص Node.js مرة واحدة عند الإقلاع ──────────────────────────────
def _check_node() -> bool:
    if shutil.which("node") is None:
        log.warning("Node.js غير موجود — سيتم الفول-باك تلقائياً على محرك Canva")
        return False
    if not os.path.exists(NODE_SCRIPT):
        log.warning("generator_api.js غير موجود — سيتم الفول-باك تلقائياً على محرك Canva")
        return False
    if not os.path.isdir(NODE_MODULES):
        log.warning("node_modules غير مثبتة — شغل: cd node_scripts && npm install")
        return False
    log.info(f"Node.js متاح ✓  |  node_modules: {len(os.listdir(NODE_MODULES))} packages")
    return True

NODE_AVAILABLE = _check_node()
log.info(f"مذكرتي Pro v9 جاهز | NODE_AVAILABLE={NODE_AVAILABLE}")

# ── CORS ─────────────────────────────────────────────────────────────
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

# ── مسارات ───────────────────────────────────────────────────────────
PUBLIC = os.path.join(os.path.dirname(os.path.abspath(__file__)), "public")

@app.route("/")
def index():
    return send_from_directory(PUBLIC, "index.html")

@app.route("/favicon.ico")
def favicon():
    return make_response("", 204)

@app.route("/health")
def health():
    cairo_ok = None
    try:
        from generator_canva import _CAIRO_OK
        cairo_ok = _CAIRO_OK
    except Exception as e:
        log.warning(f"health: تعذر تحميل generator_canva: {e}")

    return jsonify({
        "status":         "ok",
        "version":        "9.0",
        "engines":        ["canva", "classic", "premium"],
        "node_available": NODE_AVAILABLE,
        "cairo_font":     cairo_ok,
        "python":         sys.version,
    }), 200

# ── التوليد الرئيسي ──────────────────────────────────────────────────
@app.route("/static_pub/<path:filename>")
def static_files(filename):
    try:
        return send_from_directory(PUBLIC, filename)
    except Exception:
        return send_from_directory(PUBLIC, "index.html")

@app.route("/generate", methods=["POST"])
def generate():
    try:
        data = request.get_json(force=True, silent=True)
        if not data:
            log.error("generate: لم يُرسَل JSON صالح")
            return jsonify({"error": "بيانات غير صالحة — تأكد من إرسال JSON صحيح"}), 400
        if not data.get("studentName"):
            return jsonify({"error": "اسم الطالب مطلوب"}), 400
        if not data.get("titleAr"):
            return jsonify({"error": "عنوان المذكرة مطلوب"}), 400

        engine = data.get("engine", "canva")
        theme  = data.get("theme", "navy_gold")
        log.info(f"[generate] engine={engine} theme={theme} student={str(data.get('studentName',''))[:30]}")

        if engine == "premium" or theme in PREMIUM_THEMES:
            if NODE_AVAILABLE:
                return _gen_premium(data)
            else:
                log.warning("Node.js غير متاح — تحويل تلقائي إلى محرك Canva")
                data["_fallback"] = "premium→canva"
                return _gen_python(data, "generator_canva")
        elif engine == "classic":
            return _gen_python(data, "generator_classic")
        else:
            return _gen_python(data, "generator_canva")

    except Exception as e:
        log.error(f"generate: خطأ غير متوقع: {e}\n{traceback.format_exc()}")
        return jsonify({"error": f"خطأ غير متوقع: {str(e)[:300]}"}), 500


def _gen_python(data: dict, module_name: str):
    """يولد PPTX عبر Python (canva أو classic)."""
    path = None
    try:
        log.info(f"[{module_name}] بدء التوليد...")
        mod = importlib.import_module(module_name)
        importlib.reload(mod)

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            path = f.name

        mod.generate_presentation(data, path)

        if not os.path.exists(path):
            log.error(f"[{module_name}] الملف غير موجود بعد التوليد: {path}")
            return jsonify({"error": "فشل إنتاج الملف — الملف غير موجود"}), 500

        size = os.path.getsize(path)
        if size < 500:
            log.error(f"[{module_name}] الملف فارغ — حجمه {size} bytes")
            return jsonify({"error": f"فشل إنتاج الملف — الملف فارغ ({size} bytes)"}), 500

        log.info(f"[{module_name}] تم التوليد بنجاح ✓ حجم={size} bytes")

        with open(path, "rb") as f:
            pptx_bytes = f.read()

        name = data.get("studentName", "مذكرة").replace(" ", "_")
        suffix = "_canva-fallback" if data.get("_fallback") else ""
        return send_file(
            io.BytesIO(pptx_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=f"عرض_{name}{suffix}.pptx",
        )
    except ImportError as e:
        log.error(f"[{module_name}] خطأ في الاستيراد: {e}\n{traceback.format_exc()}")
        return jsonify({"error": f"خطأ في تحميل المحرك '{module_name}': {e}"}), 500
    except Exception as e:
        log.error(f"[{module_name}] خطأ في التوليد: {e}\n{traceback.format_exc()}")
        return jsonify({"error": f"خطأ في المحرك: {str(e)[:300]}"}), 500
    finally:
        if path and os.path.exists(path):
            try:
                os.unlink(path)
            except Exception:
                pass


def _gen_premium(data: dict):
    """يولد PPTX عبر Node.js/pptxgenjs مع fallback تلقائي."""
    try:
        env = os.environ.copy()
        env["NODE_PATH"] = NODE_MODULES

        result = subprocess.run(
            ["node", NODE_SCRIPT],
            input=json.dumps(data, ensure_ascii=False).encode("utf-8"),
            capture_output=True,
            timeout=120,
            cwd=os.path.join(os.path.dirname(__file__), "node_scripts"),
            env=env,
        )

        if result.returncode != 0:
            stderr = result.stderr.decode("utf-8", errors="replace").strip()
            log.error(f"Node.js exit {result.returncode}: {stderr[:500]}")
            log.warning("فشل Node.js — تحويل تلقائي إلى محرك Canva")
            data["_fallback"] = "node-error→canva"
            return _gen_python(data, "generator_canva")

        pptx_bytes = result.stdout
        if len(pptx_bytes) < 1000:
            stderr = result.stderr.decode("utf-8", errors="replace").strip()
            log.error(f"Node.js output فارغ. stderr: {stderr[:300]}")
            return jsonify({"error": "المحرك Premium أنتج ملفاً فارغاً"}), 500

        name = data.get("studentName", "مذكرة").replace(" ", "_")
        log.info(f"[premium/node] تم التوليد بنجاح ✓ حجم={len(pptx_bytes)} bytes")
        return send_file(
            io.BytesIO(pptx_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=f"عرض_{name}.pptx",
        )

    except subprocess.TimeoutExpired:
        log.error("Node.js timeout بعد 120 ثانية")
        return jsonify({"error": "انتهت مهلة التوليد (120 ثانية) — حاول تقليل عدد الشرائح"}), 504
    except FileNotFoundError:
        log.error("node غير موجود في PATH")
        return jsonify({"error": "Node.js غير مثبت على الخادم"}), 500


if __name__ == "__main__":
    port  = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_ENV") == "development"
    log.info(f"مذكرتي Pro v9 — port={port} debug={debug} node={NODE_AVAILABLE}")
    app.run(host="0.0.0.0", port=port, debug=debug)
