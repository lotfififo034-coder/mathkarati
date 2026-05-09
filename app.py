"""
مذكرتي Pro v12 — متوافق مع Python 3.11/3.12/3.13/3.14
الإصلاحات:
- لا import داخل route handlers
- لا --preload (يسبب مشكلة fork+lock في Python 3.14)
- محركات محمّلة عند أول طلب بشكل آمن thread-safe
- keep-alive /ping خفيف
"""
import os, sys, json, subprocess, shutil, tempfile, logging, io, importlib, threading

from flask import Flask, request, send_file, jsonify, send_from_directory, make_response

# ── مسار السكريبتات ──────────────────────────────────────────────────
_BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_BASE, "scripts"))

app = Flask(__name__, static_folder="public", static_url_path="")
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

NODE_SCRIPT  = os.path.join(_BASE, "node_scripts", "generator_api.js")
NODE_MODULES = os.path.join(_BASE, "node_scripts", "node_modules")

CLASSIC_THEMES = {'navy_gold','dark_teal','burgundy','forest','midnight_purple',
                  'charcoal_orange','ice_blue','sand_gold','slate_crimson'}
PREMIUM_THEMES = {'noir','atlas','sakura'}

# ── فحص Node.js ──────────────────────────────────────────────────────
def _check_node() -> bool:
    if shutil.which("node") is None:           return False
    if not os.path.exists(NODE_SCRIPT):        return False
    if not os.path.isdir(NODE_MODULES):        return False
    return True

NODE_AVAILABLE = _check_node()

# ══════════════════════════════════════════════════════════════════════
# Thread-safe module cache — آمن مع Python 3.14 وأي نسخة أخرى
# القاعدة: import يحدث مرة واحدة فقط، في الـ worker نفسه، بدون fork
# ══════════════════════════════════════════════════════════════════════
_mod_lock   = threading.Lock()
_mod_cache: dict = {}
_preload_done = threading.Event()   # ← يُشار إليه عند انتهاء التحميل

def _preload_modules():
    """
    يُحمّل الـ modules في background thread.
    _preload_done.set() عند الانتهاء حتى يعلم /health و /warmup.
    """
    def _load():
        for name in ["generator_canva", "generator_classic"]:
            try:
                _get_module(name)
            except Exception as e:
                log.warning(f"Pre-load failed for {name}: {e}")
        _preload_done.set()
        log.info("✅ All modules pre-loaded and ready")
    t = threading.Thread(target=_load, daemon=True, name="preloader")
    t.start()
    log.info("✅ Background module pre-loading started")

def _get_module(name: str):
    """
    يُحمّل الـ module مرة واحدة ويخزّنه.
    آمن تماماً: القفل لا يُستخدم إذا كان الـ module محمّلاً مسبقاً.
    """
    # fast path — بدون قفل
    m = _mod_cache.get(name)
    if m is not None:
        return m
    # slow path — مع قفل (يحدث مرة واحدة فقط لكل module)
    with _mod_lock:
        m = _mod_cache.get(name)   # double-check
        if m is not None:
            return m
        log.info(f"Loading module: {name}")
        m = importlib.import_module(name)
        _mod_cache[name] = m
        log.info(f"✅ Loaded: {name}")
        return m

# ── CORS ──────────────────────────────────────────────────────────────
@app.after_request
def _cors(r):
    r.headers["Access-Control-Allow-Origin"]  = "*"
    r.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    r.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return r

@app.before_request
def _preflight():
    if request.method == "OPTIONS":
        r = make_response("", 204)
        r.headers["Access-Control-Allow-Origin"]  = "*"
        r.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
        r.headers["Access-Control-Allow-Headers"] = "Content-Type"
        return r

# ── Routes ────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return send_from_directory("public", "index.html")

@app.route("/ping")
def ping():
    """Keep-alive خفيف — لا يلمس أي module"""
    return "pong", 200

@app.route("/health")
def health():
    """
    لا يُنفّذ أي import هنا — كل القراءات من الـ cache فقط.
    يُضيف حقل modules_ready لمعرفة إذا انتهى الـ preload.
    """
    cairo_ok = None
    m = _mod_cache.get("generator_canva")
    if m is not None:
        cairo_ok = getattr(m, "_CAIRO_OK", None)

    return jsonify({
        "status":         "ok",
        "version":        "14.0",
        "python":         sys.version.split()[0],
        "engines":        ["canva", "classic", "premium"],
        "node_available": NODE_AVAILABLE,
        "cairo_font":     cairo_ok,
        "modules_loaded": list(_mod_cache.keys()),
        "modules_ready":  _preload_done.is_set(),   # ← الجديد: هل الـ preload انتهى؟
    }), 200

@app.route("/warmup")
def warmup():
    """
    NON-BLOCKING: يرجع الحالة الحالية فوراً بدون انتظار.
    الـ frontend يستدعيه كل ثانيتين حتى modules_ready=true.
    هذا يمنع blocking الـ worker الوحيد على Render Free.
    """
    ready = _preload_done.is_set()
    return jsonify({
        "status":  "ready" if ready else "loading",
        "modules": list(_mod_cache.keys()),
        "modules_ready": ready,
    }), 200 if ready else 202

# ── التوليد ───────────────────────────────────────────────────────────
@app.route("/generate", methods=["POST"])
def generate():
    try:
        data = request.get_json(force=True, silent=True)
        if not data:
            return jsonify({"error": "بيانات غير صالحة — أرسل JSON صحيح"}), 400
        if not data.get("studentName"):
            return jsonify({"error": "اسم الطالب مطلوب"}), 400
        if not data.get("titleAr"):
            return jsonify({"error": "عنوان المذكرة مطلوب"}), 400

        engine = data.get("engine", "canva")
        theme  = data.get("theme",  "navy_gold")

        if theme not in (CLASSIC_THEMES | PREMIUM_THEMES):
            theme = "navy_gold"
            data["theme"] = theme
            log.warning("Unknown theme → navy_gold")

        log.info(f"[{engine}] theme={theme} student={str(data.get('studentName',''))[:30]}")

        if engine == "premium" or theme in PREMIUM_THEMES:
            if NODE_AVAILABLE:
                return _gen_premium(data)
            data["_fallback"] = "premium→canva"
            return _gen_python(data, "generator_canva")
        elif engine == "classic":
            return _gen_python(data, "generator_classic")
        else:
            return _gen_python(data, "generator_canva")

    except Exception as e:
        log.error(f"Unexpected: {e}", exc_info=True)
        return jsonify({"error": f"خطأ غير متوقع: {str(e)[:300]}"}), 500


def _gen_python(data: dict, module_name: str):
    path = None
    try:
        mod = _get_module(module_name)          # آمن — thread-safe

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            path = f.name

        mod.generate_presentation(data, path)

        if not os.path.exists(path) or os.path.getsize(path) < 2000:
            return jsonify({"error": "الملف فارغ أو تالف"}), 500

        with open(path, "rb") as f:
            pptx_bytes = f.read()

        import re as _re, time as _time
        name   = data.get("studentName", "")
                # اسم الملف: نأخذ الأحرف اللاتينية إن وجدت، وإلا timestamp
        latin = _re.sub(r'[^\w]', '_', name, flags=_re.ASCII).strip('_')
        # Arabic names produce empty latin → use transliterated fallback
        safe_name = latin[:20] if latin else f"prs_{int(_time.time())}"
        # suffix: use 'canva-fallback' so frontend detection works
        file_suffix = "_canva-fallback" if data.get("_fallback") else ""
        resp   = send_file(
            io.BytesIO(pptx_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=f"mathkarati_{safe_name}{file_suffix}.pptx",
        )
        resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
        resp.headers["Pragma"]        = "no-cache"
        return resp

    except ImportError as e:
        log.error(f"Import error [{module_name}]: {e}")
        return jsonify({"error": f"خطأ في تحميل المحرك: {e}"}), 500
    except Exception as e:
        log.error(f"{module_name} error: {e}", exc_info=True)
        return jsonify({"error": f"خطأ في المحرك: {str(e)[:300]}"}), 500
    finally:
        if path and os.path.exists(path):
            try: os.unlink(path)
            except Exception: pass


def _gen_premium(data: dict):
    try:
        env = os.environ.copy()
        env["NODE_PATH"] = NODE_MODULES
        result = subprocess.run(
            ["node", NODE_SCRIPT],
            input=json.dumps(data, ensure_ascii=False).encode("utf-8"),
            capture_output=True, timeout=90,
            cwd=os.path.join(_BASE, "node_scripts"), env=env,
        )
        if result.returncode != 0:
            log.error(f"Node exit {result.returncode}: {result.stderr.decode(errors='replace')[:300]}")
            data["_fallback"] = "node-error→canva"
            return _gen_python(data, "generator_canva")

        pptx_bytes = result.stdout
        if len(pptx_bytes) < 1000:
            return jsonify({"error": "المحرك Premium أنتج ملفاً فارغاً"}), 500

        import re as _re, time as _time
        name = data.get("studentName", "")
        latin = _re.sub(r'[^\w]', '_', name, flags=_re.ASCII).strip('_')
        safe_name = latin[:20] if latin else f"prs_{int(_time.time())}"
        return send_file(
            io.BytesIO(pptx_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=f"mathkarati_{safe_name}_premium.pptx",
        )
    except subprocess.TimeoutExpired:
        return jsonify({"error": "انتهت مهلة التوليد — قلّل عدد الشرائح"}), 504
    except FileNotFoundError:
        return jsonify({"error": "Node.js غير مثبت"}), 500


# ── تحميل مسبق عند بدء التشغيل (في Gunicorn و dev معاً) ─────────────
_preload_modules()

if __name__ == "__main__":
    port  = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_ENV") == "development"
    app.run(host="0.0.0.0", port=port, debug=debug)
