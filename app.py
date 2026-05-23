"""
Flask API — مذكرتي Pro v17.2
+ نظام البيع والتحميل المحمي
+ لوحة الإدارة
+ Preview آمن بدون تحميل
"""
import base64
import functools
import hashlib
import io
import logging
import mimetypes
import os
import sys
import time
import unicodedata
import uuid

from flask import (Flask, Response, abort, jsonify, make_response,
                   request, send_file, send_from_directory)

_BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _BASE)

from core.models import PresentationRequest
from core.payment_models import (Order, OrderStatus, PaymentMethod,
                                  StoredPresentation, PRICE_DZD)
from core.order_store import get_store
from engine.pipeline import get_pipeline

app = Flask(__name__, static_folder="public", static_url_path="")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
    stream=sys.stdout,
)
log = logging.getLogger(__name__)

ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "mathkarati_admin_2025")
ADMIN_TOKEN_HASH = hashlib.sha256(ADMIN_PASSWORD.encode()).hexdigest()
MAX_RECEIPT_MB = 10
ALLOWED_RECEIPT_TYPES = {"image/jpeg", "image/png", "image/webp", "application/pdf"}


def _safe_filename(name):
    if not name:
        return f"prs_{int(time.time())}"
    normalized = unicodedata.normalize("NFKD", name)
    ascii_str = normalized.encode("ascii", "ignore").decode("ascii")
    safe = "".join(c if c.isalnum() else "_" for c in ascii_str).strip("_")
    if not safe:
        safe = f"student_{int(time.time()) % 100000}"
    return safe[:24]


def require_admin(f):
    @functools.wraps(f)
    def wrapper(*args, **kwargs):
        token = request.headers.get("X-Admin-Token", "")
        if hashlib.sha256(token.encode()).hexdigest() != ADMIN_TOKEN_HASH:
            return jsonify({"error": "غير مصرح"}), 401
        return f(*args, **kwargs)
    return wrapper


@app.after_request
def _cors(r):
    r.headers["Access-Control-Allow-Origin"] = "*"
    r.headers["Access-Control-Allow-Methods"] = "GET, POST, PUT, OPTIONS"
    r.headers["Access-Control-Allow-Headers"] = "Content-Type, X-Admin-Token"
    return r

@app.before_request
def _preflight():
    if request.method == "OPTIONS":
        r = make_response("", 204)
        r.headers["Access-Control-Allow-Origin"] = "*"
        r.headers["Access-Control-Allow-Methods"] = "GET, POST, PUT, OPTIONS"
        r.headers["Access-Control-Allow-Headers"] = "Content-Type, X-Admin-Token"
        return r


@app.route("/")
def index():
    return send_from_directory("public", "index.html")

@app.route("/admin")
def admin_page():
    return send_from_directory("public", "admin.html")

@app.route("/ping")
def ping():
    return "pong", 200

@app.route("/health")
def health():
    pipeline = get_pipeline()
    store = get_store()
    return jsonify({
        "status": "ok", "version": "17.2",
        "python": sys.version.split()[0],
        "font": pipeline._font,
        "stats": store.get_stats(),
    }), 200

@app.route("/warmup")
def warmup():
    get_pipeline()
    return jsonify({"status": "ready", "modules_ready": True}), 200


@app.route("/generate", methods=["POST"])
def generate():
    t0 = time.monotonic()
    raw = request.get_json(force=True, silent=True)
    if not raw:
        return jsonify({"error": "بيانات غير صالحة"}), 400

    req = PresentationRequest.from_dict(raw)
    errors = req.validate()
    if errors:
        return jsonify({"error": " | ".join(errors)}), 400

    pipeline = get_pipeline()
    result = pipeline.build(req)

    if not result.success:
        log.error(f"Build failed: {result.error}")
        return jsonify({"error": result.error, "stages": result.stages}), 500

    safe = _safe_filename(req.student_name)
    filename = f"mathkarati_{safe}.pptx"
    presentation_id = str(uuid.uuid4())

    store = get_store()
    prs = StoredPresentation(
        presentation_id=presentation_id,
        filename=filename,
        data_b64=base64.b64encode(result.data).decode("ascii"),
        slide_count=result.slide_count,
        student_name=req.student_name,
        title=req.title_ar,
        engine=req.engine,
        theme=req.theme,
    )
    store.store_presentation(prs)

    elapsed = time.monotonic() - t0
    log.info(f"Generated: {presentation_id} slides={result.slide_count} {elapsed:.2f}s")

    return jsonify({
        "ok": True,
        "presentation_id": presentation_id,
        "slides": result.slide_count,
        "font": result.font_used,
        "elapsed": round(elapsed, 2),
        "stages": result.stages,
        "price": PRICE_DZD,
    })


@app.route("/preview/<presentation_id>")
def preview_info(presentation_id):
    store = get_store()
    prs = store.get_presentation(presentation_id)
    if not prs:
        return jsonify({"error": "العرض غير موجود أو انتهت صلاحيته"}), 404
    return jsonify(prs.to_preview_dict())


@app.route("/preview/<presentation_id>/slide/<int:slide_num>")
def preview_slide(presentation_id, slide_num):
    store = get_store()
    prs = store.get_presentation(presentation_id)
    if not prs:
        return jsonify({"error": "العرض غير موجود"}), 404
    if slide_num < 1 or slide_num > prs.slide_count:
        return jsonify({"error": "رقم الشريحة غير صالح"}), 400

    svg = _make_slide_svg(slide_num, prs.slide_count, prs.title, prs.student_name, prs.theme)
    return Response(svg, mimetype="image/svg+xml",
                    headers={"Cache-Control": "no-store"})


def _make_slide_svg(slide_num, total, title, student, theme):
    theme_colors = {
        "navy_gold": ("#07172F", "#C6A03C"),
        "dark_teal": ("#0F2D2D", "#2DD4BF"),
        "burgundy": ("#2D0A0A", "#DC2626"),
        "midnight_purple": ("#1A0A2E", "#7C3AED"),
        "charcoal_orange": ("#1C1C1C", "#F97316"),
        "ice_blue": ("#1a3a5c", "#0EA5E9"),
        "sand_gold": ("#3d2c00", "#D97706"),
        "forest": ("#0a2e1a", "#22C55E"),
        "slate_crimson": ("#1a1a2e", "#EF4444"),
        "noir": ("#0a0a0a", "#ffffff"),
        "atlas": ("#0c1445", "#3B82F6"),
        "sakura": ("#2d0a1a", "#F472B6"),
    }
    bg, accent = theme_colors.get(theme, ("#07172F", "#C6A03C"))
    t = title[:45].replace("<","&lt;").replace(">","&gt;").replace("&","&amp;")
    s = student[:30].replace("<","&lt;").replace(">","&gt;").replace("&","&amp;")

    slide_labels = {
        1: "الغلاف", 2: "مقدمة", 3: "خطة العمل", 4: "الإشكالية",
        5: "الأهداف", 6: "الأهمية", 7: "المنهجية", 8: "الإحصاءات",
        9: "النتائج", 10: "الخاتمة", 11: "التوصيات", 12: "الآفاق",
        13: "المراجع", 14: "شكر وتقدير",
    }
    label = slide_labels.get(slide_num, f"الشريحة {slide_num}")

    return f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 960 540">
  <defs>
    <pattern id="wm" x="0" y="0" width="220" height="90" patternUnits="userSpaceOnUse" patternTransform="rotate(-30)">
      <text x="5" y="55" font-family="Cairo,Arial" font-size="13" fill="rgba(255,255,255,0.07)" font-weight="700">مذكرتي Pro — للمعاينة فقط</text>
    </pattern>
    <linearGradient id="bg" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" style="stop-color:{bg};stop-opacity:1"/>
      <stop offset="100%" style="stop-color:{bg}ee;stop-opacity:1"/>
    </linearGradient>
  </defs>
  <rect width="960" height="540" fill="url(#bg)"/>
  <rect x="0" y="0" width="6" height="540" fill="{accent}"/>
  <rect x="0" y="520" width="960" height="20" fill="{accent}" opacity="0.25"/>
  <rect x="20" y="18" width="90" height="28" rx="6" fill="{accent}" opacity="0.85"/>
  <text x="65" y="37" font-family="Cairo,Arial" font-size="12" fill="{bg}" text-anchor="middle" font-weight="700">{slide_num} / {total}</text>
  <text x="480" y="160" font-family="Cairo,Arial" font-size="18" fill="{accent}" text-anchor="middle" font-weight="700" opacity="0.6">{label}</text>
  <text x="480" y="215" font-family="Cairo,Arial" font-size="24" fill="white" text-anchor="middle" font-weight="700">{t}</text>
  <text x="480" y="255" font-family="Cairo,Arial" font-size="14" fill="{accent}" text-anchor="middle">{s}</text>
  <rect x="160" y="290" width="640" height="4" rx="2" fill="white" opacity="0.08"/>
  <rect x="200" y="308" width="560" height="4" rx="2" fill="white" opacity="0.06"/>
  <rect x="180" y="326" width="600" height="4" rx="2" fill="white" opacity="0.05"/>
  <rect x="220" y="344" width="520" height="4" rx="2" fill="white" opacity="0.04"/>
  <rect x="160" y="362" width="640" height="4" rx="2" fill="white" opacity="0.03"/>
  <rect width="960" height="540" fill="url(#wm)"/>
  <rect x="260" y="420" width="440" height="72" rx="12" fill="rgba(0,0,0,0.55)" stroke="{accent}" stroke-width="1" stroke-opacity="0.3"/>
  <text x="480" y="448" font-family="Cairo,Arial" font-size="14" fill="{accent}" text-anchor="middle" font-weight="700">🔒 نسخة معاينة محمية</text>
  <text x="480" y="468" font-family="Cairo,Arial" font-size="11" fill="rgba(255,255,255,0.65)" text-anchor="middle">ادفع 800 دج واحصل على كود تحميل الملف الكامل</text>
  <text x="480" y="484" font-family="Cairo,Arial" font-size="10" fill="rgba(255,255,255,0.4)" text-anchor="middle">CCP / BaridiMob</text>
</svg>'''


@app.route("/order", methods=["POST"])
def create_order():
    data = request.get_json(force=True, silent=True) or {}
    presentation_id = data.get("presentation_id", "").strip()
    student_name = data.get("student_name", "").strip()
    student_email = data.get("student_email", "").strip()
    phone = data.get("phone", "").strip()
    payment_method = data.get("payment_method", "ccp").strip()

    errors = []
    if not presentation_id: errors.append("معرف العرض مطلوب")
    if not student_name: errors.append("اسم الطالب مطلوب")
    if not phone: errors.append("رقم الهاتف مطلوب")
    if payment_method not in ["ccp", "baridi"]: errors.append("طريقة الدفع غير صالحة")
    if errors:
        return jsonify({"error": " | ".join(errors)}), 400

    store = get_store()
    prs = store.get_presentation(presentation_id)
    if not prs:
        return jsonify({"error": "العرض غير موجود أو انتهت صلاحيته. أعد توليد العرض."}), 404

    order = Order.create(presentation_id, student_name, student_email, phone, payment_method)
    store.save_order(order)
    log.info(f"New order: {order.order_id}")
    return jsonify({
        "ok": True, "order_id": order.order_id,
        "amount": PRICE_DZD, "payment_method": payment_method,
        "status": order.status.value,
    })


@app.route("/order/<order_id>/receipt", methods=["POST"])
def upload_receipt(order_id):
    store = get_store()
    order = store.get_order(order_id)
    if not order:
        return jsonify({"error": "الطلب غير موجود"}), 404
    if order.status not in [OrderStatus.PENDING, OrderStatus.UPLOADED]:
        return jsonify({"error": "لا يمكن رفع الوصل"}), 400
    if "receipt" not in request.files:
        return jsonify({"error": "الملف مطلوب"}), 400

    f = request.files["receipt"]
    content = f.read()
    if len(content) > MAX_RECEIPT_MB * 1024 * 1024:
        return jsonify({"error": f"حجم الملف يجب أن يكون أقل من {MAX_RECEIPT_MB}MB"}), 400

    mime = f.mimetype or mimetypes.guess_type(f.filename)[0] or ""
    if mime not in ALLOWED_RECEIPT_TYPES:
        return jsonify({"error": "نوع الملف غير مقبول (JPG, PNG, PDF)"}), 400

    path = store.save_receipt(order_id, f.filename, content)
    order.receipt_path = path
    order.receipt_filename = f.filename
    order.status = OrderStatus.UPLOADED
    store.save_order(order)
    return jsonify({"ok": True, "status": order.status.value})


@app.route("/order/<order_id>/status")
def order_status(order_id):
    store = get_store()
    order = store.get_order(order_id)
    if not order:
        return jsonify({"error": "الطلب غير موجود"}), 404
    return jsonify({
        "order_id": order.order_id, "status": order.status.value,
        "has_code": order.download_code is not None,
        "code_used": order.code_used, "amount": order.amount,
    })


@app.route("/order/<order_id>/activate", methods=["POST"])
def activate_code(order_id):
    store = get_store()
    order = store.get_order(order_id)
    if not order:
        return jsonify({"error": "الطلب غير موجود"}), 404

    data = request.get_json(force=True, silent=True) or {}
    code = data.get("code", "").strip()

    if not order.is_code_valid(code):
        if order.code_used:
            return jsonify({"error": "تم استخدام هذا الكود مسبقاً"}), 400
        if order.status != OrderStatus.APPROVED:
            return jsonify({"error": "الطلب لم تتم الموافقة عليه بعد"}), 400
        return jsonify({"error": "كود غير صحيح أو منتهي الصلاحية"}), 400

    prs = store.get_presentation(order.presentation_id)
    if not prs:
        return jsonify({"error": "انتهت صلاحية العرض. تواصل مع الإدارة.", "contact": True}), 410

    ip = request.headers.get("X-Forwarded-For", request.remote_addr or "unknown")
    ua = request.user_agent.string[:200]
    order.mark_downloaded(ip, ua)
    store.save_order(order)
    log.info(f"Download: order={order_id} ip={ip}")

    file_data = base64.b64decode(prs.data_b64)
    return send_file(
        io.BytesIO(file_data),
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        as_attachment=True,
        download_name=prs.filename,
    )


# ── ADMIN API ───────────────────────────────────────────────────────────
@app.route("/admin/auth", methods=["POST"])
def admin_auth():
    data = request.get_json(force=True, silent=True) or {}
    password = data.get("password", "")
    if hashlib.sha256(password.encode()).hexdigest() == ADMIN_TOKEN_HASH:
        return jsonify({"ok": True, "token": password})
    return jsonify({"error": "كلمة المرور غير صحيحة"}), 401


@app.route("/admin/stats")
@require_admin
def admin_stats():
    return jsonify(get_store().get_stats())


@app.route("/admin/orders")
@require_admin
def admin_orders():
    store = get_store()
    status_filter = request.args.get("status")
    if status_filter:
        try:
            orders = store.get_orders_by_status(OrderStatus(status_filter))
        except ValueError:
            orders = store.get_all_orders()
    else:
        orders = store.get_all_orders()
    return jsonify([o.to_dict() for o in orders])


@app.route("/admin/orders/<order_id>/receipt")
@require_admin
def admin_get_receipt(order_id):
    store = get_store()
    path = store.get_receipt_path(order_id)
    if not path or not os.path.exists(path):
        return jsonify({"error": "لا يوجد وصل"}), 404
    mime = mimetypes.guess_type(path)[0] or "application/octet-stream"
    return send_file(path, mimetype=mime)


@app.route("/admin/orders/<order_id>/approve", methods=["POST"])
@require_admin
def admin_approve(order_id):
    store = get_store()
    order = store.get_order(order_id)
    if not order:
        return jsonify({"error": "الطلب غير موجود"}), 404

    data = request.get_json(force=True, silent=True) or {}
    hours = int(data.get("hours_valid", 48))
    note = data.get("note", "")

    code = order.generate_code(hours_valid=hours)
    order.status = OrderStatus.APPROVED
    order.admin_note = note
    store.save_order(order)
    log.info(f"Approved: order={order_id} code={code}")
    return jsonify({
        "ok": True, "order_id": order_id,
        "download_code": code, "expires_hours": hours,
        "status": order.status.value,
    })


@app.route("/admin/orders/<order_id>/reject", methods=["POST"])
@require_admin
def admin_reject(order_id):
    store = get_store()
    order = store.get_order(order_id)
    if not order:
        return jsonify({"error": "الطلب غير موجود"}), 404
    data = request.get_json(force=True, silent=True) or {}
    order.status = OrderStatus.REJECTED
    order.admin_note = data.get("note", "الوصل غير صحيح")
    store.save_order(order)
    return jsonify({"ok": True, "status": order.status.value})


@app.route("/admin/orders/<order_id>/regen_code", methods=["POST"])
@require_admin
def admin_regen_code(order_id):
    store = get_store()
    order = store.get_order(order_id)
    if not order:
        return jsonify({"error": "الطلب غير موجود"}), 404
    if order.status not in [OrderStatus.APPROVED, OrderStatus.DOWNLOADED]:
        return jsonify({"error": "يجب أن يكون الطلب موافقاً عليه"}), 400
    data = request.get_json(force=True, silent=True) or {}
    hours = int(data.get("hours_valid", 48))
    code = order.generate_code(hours_valid=hours)
    order.code_used = False
    order.status = OrderStatus.APPROVED
    store.save_order(order)
    return jsonify({"ok": True, "download_code": code, "expires_hours": hours})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    get_pipeline()
    get_store()
    app.run(host="0.0.0.0", port=port, debug=False, use_reloader=False)
