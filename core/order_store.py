"""
Order Store — مذكرتي Pro v17
تخزين الطلبات والعروض بشكل آمن
"""
from __future__ import annotations
import json
import os
import threading
import time
import logging
from typing import Dict, Optional, List

from core.payment_models import Order, StoredPresentation, OrderStatus

log = logging.getLogger(__name__)

_DATA_DIR = os.environ.get("DATA_DIR", os.path.join(os.path.dirname(__file__), "..", "data"))
_ORDERS_FILE = os.path.join(_DATA_DIR, "orders.json")
_RECEIPTS_DIR = os.path.join(_DATA_DIR, "receipts")


def _ensure_dirs():
    os.makedirs(_DATA_DIR, exist_ok=True)
    os.makedirs(_RECEIPTS_DIR, exist_ok=True)


class OrderStore:
    """Thread-safe order storage"""

    def __init__(self):
        self._lock = threading.RLock()
        self._orders: Dict[str, Order] = {}
        self._presentations: Dict[str, StoredPresentation] = {}  # in-memory only
        _ensure_dirs()
        self._load()

    def _load(self):
        if os.path.exists(_ORDERS_FILE):
            try:
                with open(_ORDERS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                for d in data:
                    o = Order.from_dict(d)
                    self._orders[o.order_id] = o
                log.info(f"Loaded {len(self._orders)} orders")
            except Exception as e:
                log.error(f"Failed to load orders: {e}")

    def _save(self):
        try:
            with open(_ORDERS_FILE, "w", encoding="utf-8") as f:
                json.dump([o.to_dict() for o in self._orders.values()], f,
                          ensure_ascii=False, indent=2)
        except Exception as e:
            log.error(f"Failed to save orders: {e}")

    # ── Orders ────────────────────────────────────────────────────────

    def save_order(self, order: Order):
        with self._lock:
            order.updated_at = time.time()
            self._orders[order.order_id] = order
            self._save()

    def get_order(self, order_id: str) -> Optional[Order]:
        with self._lock:
            return self._orders.get(order_id)

    def get_all_orders(self) -> List[Order]:
        with self._lock:
            return sorted(self._orders.values(), key=lambda o: o.created_at, reverse=True)

    def get_orders_by_status(self, status: OrderStatus) -> List[Order]:
        with self._lock:
            return [o for o in self._orders.values() if o.status == status]

    # ── Receipts ──────────────────────────────────────────────────────

    def save_receipt(self, order_id: str, filename: str, data: bytes) -> str:
        """حفظ وصل الدفع"""
        _ensure_dirs()
        ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else "jpg"
        safe_name = f"{order_id}.{ext}"
        path = os.path.join(_RECEIPTS_DIR, safe_name)
        with open(path, "wb") as f:
            f.write(data)
        return path

    def get_receipt_path(self, order_id: str) -> Optional[str]:
        for ext in ["jpg", "jpeg", "png", "pdf", "webp"]:
            p = os.path.join(_RECEIPTS_DIR, f"{order_id}.{ext}")
            if os.path.exists(p):
                return p
        return None

    # ── Presentations (in-memory, secure) ────────────────────────────

    def store_presentation(self, prs: StoredPresentation):
        """تخزين العرض في الذاكرة فقط (آمن)"""
        with self._lock:
            self._cleanup_expired()
            self._presentations[prs.presentation_id] = prs
            log.info(f"Stored presentation {prs.presentation_id} ({prs.slide_count} slides)")

    def get_presentation(self, presentation_id: str) -> Optional[StoredPresentation]:
        with self._lock:
            prs = self._presentations.get(presentation_id)
            if prs and prs.is_expired():
                del self._presentations[presentation_id]
                return None
            return prs

    def _cleanup_expired(self):
        expired = [k for k, v in self._presentations.items() if v.is_expired()]
        for k in expired:
            del self._presentations[k]

    # ── Stats ─────────────────────────────────────────────────────────

    def get_stats(self) -> dict:
        with self._lock:
            orders = list(self._orders.values())
            return {
                "total": len(orders),
                "pending": sum(1 for o in orders if o.status == OrderStatus.PENDING),
                "uploaded": sum(1 for o in orders if o.status == OrderStatus.UPLOADED),
                "approved": sum(1 for o in orders if o.status == OrderStatus.APPROVED),
                "rejected": sum(1 for o in orders if o.status == OrderStatus.REJECTED),
                "downloaded": sum(1 for o in orders if o.status == OrderStatus.DOWNLOADED),
                "revenue": sum(o.amount for o in orders if o.status in
                               [OrderStatus.APPROVED, OrderStatus.DOWNLOADED]) ,
                "stored_presentations": len(self._presentations),
            }


# Singleton
_store: Optional[OrderStore] = None
_store_lock = threading.Lock()


def get_store() -> OrderStore:
    global _store
    if _store is None:
        with _store_lock:
            if _store is None:
                _store = OrderStore()
    return _store
