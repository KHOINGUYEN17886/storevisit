# -*- coding: utf-8 -*-
"""
narrator.py — StoreVisit: Humanized AI Narrative Layer
======================================================
Biến dữ liệu kiểm tra cửa hàng (StoreReportData) thành văn phong tiếng Việt
tự nhiên, giọng QLKD/ASM chuyên nghiệp — thay cho các câu template máy móc.

Cascade LLM: Gemini (free) → Groq (free) → Claude Haiku (paid) → rule-based.
Xem reports/ai_client.py. Không có API key → tự động fallback template gốc.

Nguyên tắc GROUNDING (chống bịa):
  - Chỉ diễn đạt lại DỮ LIỆU được cung cấp; tuyệt đối không thêm số liệu/sự kiện mới.
  - Nếu không có dữ liệu → trả về fallback, KHÔNG gọi LLM.

Cache: hash nội dung prompt → file JSON trong temp/ai_cache/ để không gọi lại
cùng một nội dung (tiết kiệm quota + tăng tốc khi render lại).
"""
from __future__ import annotations

import os
import re
import json
import hashlib
from typing import List, Optional

try:
    from reports.ai_client import AIClient
except Exception:  # pragma: no cover - import path fallback khi chạy lẻ
    from ai_client import AIClient  # type: ignore


_SYSTEM_PROMPT = (
    "Bạn là chuyên gia vận hành chuỗi bán lẻ thời trang cao cấp An Phước — Pierre Cardin. "
    "Bạn viết nhận định cho báo cáo công tác kiểm tra cửa hàng do Quản lý Kinh doanh (QLKD/ASM) trình bày. "
    "Văn phong: tiếng Việt chuẩn mực, chuyên nghiệp, súc tích, mang tính quản trị và định hướng hành động — "
    "như một giám đốc vùng dày dạn kinh nghiệm nói với cửa hàng trưởng. "
    "TUYỆT ĐỐI KHÔNG bịa thêm số liệu, tên người, hay sự kiện không có trong dữ liệu. "
    "Không dùng emoji. Không lặp lại nguyên văn dữ liệu thô — hãy diễn giải thành câu văn mạch lạc, có nhận định và khuyến nghị."
)


class StoreNarrator:
    """Sinh văn bản humanized cho báo cáo; an toàn khi không có LLM (fallback)."""

    def __init__(self, cache_dir: Optional[str] = None, enabled: bool = True):
        self.enabled = enabled
        self._client: Optional[AIClient] = None
        self.last_source = "none"
        # Cache dir
        if cache_dir is None:
            base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            cache_dir = os.path.join(base, "temp", "ai_cache")
        self.cache_dir = cache_dir
        try:
            os.makedirs(self.cache_dir, exist_ok=True)
        except Exception:
            pass

    # ── infra ────────────────────────────────────────────────────────────────
    def _get_client(self) -> Optional[AIClient]:
        if not self.enabled:
            return None
        if self._client is None:
            has_key = any(os.environ.get(k) for k in
                          ("GEMINI_API_KEY", "GROQ_API_KEY", "ANTHROPIC_API_KEY"))
            if not has_key:
                self.enabled = False
                return None
            self._client = AIClient()
        return self._client

    def _cache_path(self, key: str) -> str:
        h = hashlib.sha1(key.encode("utf-8")).hexdigest()[:20]
        return os.path.join(self.cache_dir, f"{h}.json")

    def _cached(self, key: str) -> Optional[str]:
        try:
            p = self._cache_path(key)
            if os.path.exists(p):
                with open(p, "r", encoding="utf-8") as f:
                    return json.load(f).get("text")
        except Exception:
            pass
        return None

    def _store_cache(self, key: str, text: str, source: str) -> None:
        try:
            with open(self._cache_path(key), "w", encoding="utf-8") as f:
                json.dump({"text": text, "source": source}, f, ensure_ascii=False)
        except Exception:
            pass

    @staticmethod
    def _clean(text: str) -> str:
        """Bỏ markdown/emoji thừa, chuẩn hoá khoảng trắng."""
        if not text:
            return ""
        text = re.sub(r"[*_`#>]+", "", text)              # markdown
        text = re.sub(r"[ \t]+", " ", text)
        text = re.sub(r"\n{3,}", "\n\n", text).strip()
        return text

    def _run(self, prompt: str, fallback: str, max_tokens: int = 400) -> str:
        client = self._get_client()
        if client is None:
            self.last_source = "rule-based"
            return fallback
        cached = self._cached(prompt)
        if cached is not None:
            self.last_source = "cache"
            return cached
        try:
            out = client.generate(prompt, fallback=fallback,
                                   system=_SYSTEM_PROMPT, max_tokens=max_tokens)
            self.last_source = client.last_source
            if not out or self.last_source == "rule-based":
                return fallback
            out = self._clean(out)
            self._store_cache(prompt, out, self.last_source)
            return out
        except Exception:
            self.last_source = "rule-based"
            return fallback

    # ── public API ───────────────────────────────────────────────────────────
    def executive_summary(self, facts: List[str], store_name: str,
                          fallback: str = "") -> str:
        """
        'Nhận định chung của QLKD' — tổng hợp toàn bộ dữ kiện thành 3-5 câu
        nhận định + định hướng. `facts` là các gạch đầu dòng dữ liệu thô đã grounded.
        """
        facts = [f for f in facts if f and str(f).strip()]
        if not facts:
            return fallback
        fb = fallback or ("Cửa hàng vận hành ổn định tại thời điểm kiểm tra; "
                          "đề nghị cửa hàng trưởng duy trì và bám sát các hạng mục đã trao đổi.")
        prompt = (
            f"Cửa hàng: {store_name}.\n"
            f"Dữ liệu ghi nhận trong buổi kiểm tra (chỉ dùng đúng các dữ kiện này):\n"
            + "\n".join(f"- {f}" for f in facts)
            + "\n\nHãy viết một đoạn NHẬN ĐỊNH CHUNG CỦA QLKD dài 3-5 câu: "
              "đánh giá tổng thể tình hình, nêu bật 1-2 điểm mạnh và 1-2 vấn đề trọng yếu cần ưu tiên, "
              "kết bằng một định hướng hành động ngắn gọn. Viết liền mạch, không gạch đầu dòng."
        )
        return self._run(prompt, fb, max_tokens=400)

    def humanize_findings(self, section_label: str, findings: List[str],
                          rating: str = "", fallback: str = "") -> str:
        """
        Diễn đạt lại danh sách lỗi/nhận xét thô của 1 hạng mục thành 1-3 câu
        văn tự nhiên kèm khuyến nghị. Giữ nguyên bản chất, không thêm lỗi mới.
        """
        findings = [f for f in findings if f and str(f).strip()]
        if not findings:
            return fallback
        fb = fallback or " ".join(findings)
        rating_txt = f" (Kết quả đánh giá: {rating})" if rating else ""
        prompt = (
            f"Hạng mục kiểm tra: {section_label}{rating_txt}.\n"
            f"Các ghi nhận thực tế (chỉ diễn đạt lại đúng nội dung này):\n"
            + "\n".join(f"- {f}" for f in findings)
            + "\n\nHãy viết 1-3 câu nhận xét chuyên nghiệp bằng tiếng Việt: mô tả vấn đề "
              "và nêu hướng khắc phục cụ thể, khả thi cho cửa hàng trưởng. Viết liền mạch, không gạch đầu dòng."
        )
        return self._run(prompt, fb, max_tokens=300)


# ── singleton tiện dụng ─────────────────────────────────────────────────────────
_default: Optional[StoreNarrator] = None


def get_narrator() -> StoreNarrator:
    global _default
    if _default is None:
        _default = StoreNarrator()
    return _default


if __name__ == "__main__":
    n = StoreNarrator()
    print("--- executive_summary ---")
    print(n.executive_summary(
        facts=[
            "Doanh thu tháng đạt 1.2 tỷ VNĐ, hoàn thành 92% chỉ tiêu, tăng 8% so với cùng kỳ.",
            "Mặt tiền: bảng hiệu bám bụi, cần vệ sinh.",
            "Trưng bày An Phước đạt chuẩn; khu Pierre Cardin thiếu size L áo sơ mi.",
            "Nhân sự: đồng phục đúng SOP; thái độ phục vụ tốt.",
        ],
        store_name="Phú Mỹ Hưng 1 - Nguyễn Đức Cảnh",
    ))
    print("source =", n.last_source)
