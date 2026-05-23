"""
Payment & Download System Models — مذكرتي Pro v17
نماذج نظام البيع والتحميل المحمي
"""
from __future__ import annotations
import uuid
import time
import secrets
import hashlib
from dataclasses import dataclass, field
from typing import Optional
from enum import Enum


class OrderStatus(str, Enum):
    PENDING   = "pending"     # في انتظار رفع الوصل
    UPLOADED  = "uploaded"    # تم رفع الوصل، في انتظار المراجعة
    APPROVED  = "approved"    # تمت الموافقة، الكود متاح
    REJECTED  = "rejected"    # مرفوض
    DOWNLOADED = "downloaded" # تم التحميل


class PaymentMethod(str, Enum):
    CCP      = "ccp"
    BARIDI   = "baridi"


PRICE_DZD = 800


@dataclass
class Order:
    order_id: str
    presentation_id: str        # ID العرض المرتبط
    student_name: str
    student_email: str
    phone: str
    payment_method: PaymentMethod
    status: OrderStatus = OrderStatus.PENDING
    receipt_path: Optional[str] = None    # مسار صورة الوصل
    receipt_filename: Optional[str] = None
    download_code: Optional[str] = None   # كود التفعيل
    code_expires_at: Optional[float] = None  # Unix timestamp
    code_used: bool = False
    created_at: float = field(default_factory=time.time)
    updated_at: float = field(default_factory=time.time)
    admin_note: str = ""
    download_ip: Optional[str] = None
    download_at: Optional[float] = None
    download_user_agent: Optional[str] = None
    amount: int = PRICE_DZD

    @classmethod
    def create(cls, presentation_id: str, student_name: str,
               student_email: str, phone: str,
               payment_method: str) -> "Order":
        return cls(
            order_id=str(uuid.uuid4()),
            presentation_id=presentation_id,
            student_name=student_name,
            student_email=student_email,
            phone=phone,
            payment_method=PaymentMethod(payment_method),
        )

    def generate_code(self, hours_valid: int = 48) -> str:
        """توليد كود تفعيل آمن"""
        code = secrets.token_urlsafe(12).upper()[:16]
        self.download_code = code
        self.code_expires_at = time.time() + (hours_valid * 3600)
        self.code_used = False
        self.status = OrderStatus.APPROVED
        self.updated_at = time.time()
        return code

    def is_code_valid(self, code: str) -> bool:
        if not self.download_code:
            return False
        if self.code_used:
            return False
        if self.code_expires_at and time.time() > self.code_expires_at:
            return False
        if self.status != OrderStatus.APPROVED:
            return False
        return self.download_code == code.strip().upper()

    def mark_downloaded(self, ip: str, user_agent: str):
        self.code_used = True
        self.status = OrderStatus.DOWNLOADED
        self.download_ip = ip
        self.download_at = time.time()
        self.download_user_agent = user_agent
        self.updated_at = time.time()

    def to_dict(self) -> dict:
        return {
            "order_id": self.order_id,
            "presentation_id": self.presentation_id,
            "student_name": self.student_name,
            "student_email": self.student_email,
            "phone": self.phone,
            "payment_method": self.payment_method.value,
            "status": self.status.value,
            "receipt_path": self.receipt_path,
            "receipt_filename": self.receipt_filename,
            "download_code": self.download_code,
            "code_expires_at": self.code_expires_at,
            "code_used": self.code_used,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
            "admin_note": self.admin_note,
            "download_ip": self.download_ip,
            "download_at": self.download_at,
            "download_user_agent": self.download_user_agent,
            "amount": self.amount,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "Order":
        o = cls(
            order_id=d["order_id"],
            presentation_id=d["presentation_id"],
            student_name=d["student_name"],
            student_email=d.get("student_email", ""),
            phone=d.get("phone", ""),
            payment_method=PaymentMethod(d.get("payment_method", "ccp")),
            status=OrderStatus(d.get("status", "pending")),
            receipt_path=d.get("receipt_path"),
            receipt_filename=d.get("receipt_filename"),
            download_code=d.get("download_code"),
            code_expires_at=d.get("code_expires_at"),
            code_used=d.get("code_used", False),
            created_at=d.get("created_at", time.time()),
            updated_at=d.get("updated_at", time.time()),
            admin_note=d.get("admin_note", ""),
            download_ip=d.get("download_ip"),
            download_at=d.get("download_at"),
            download_user_agent=d.get("download_user_agent"),
            amount=d.get("amount", PRICE_DZD),
        )
        return o


@dataclass
class StoredPresentation:
    """العرض المخزن في الذاكرة (محمي)"""
    presentation_id: str
    filename: str
    data_b64: str           # base64 encoded PPTX - مخزن محمي
    slide_count: int
    student_name: str
    title: str
    engine: str
    theme: str
    created_at: float = field(default_factory=time.time)
    expires_at: float = field(default_factory=lambda: time.time() + 86400 * 3)  # 3 days

    def is_expired(self) -> bool:
        return time.time() > self.expires_at

    def to_preview_dict(self) -> dict:
        """معلومات آمنة للمعاينة - بدون البيانات الحقيقية"""
        return {
            "presentation_id": self.presentation_id,
            "slide_count": self.slide_count,
            "student_name": self.student_name,
            "title": self.title,
            "engine": self.engine,
            "theme": self.theme,
        }
