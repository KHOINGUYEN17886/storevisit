# -*- coding: utf-8 -*-
"""
ai_client.py — Retail Commander: Centralized AI Client
=======================================================
Single entry point cho tất cả AI narrative calls trong hệ thống.

Cascade ưu tiên:
  1. Gemini 2.5 Flash Lite  (google-genai — miễn phí, nhanh)
  2. Groq llama-3.3-70b     (groq package — miễn phí, fallback khi Gemini quota hết)
  3. Claude Haiku 4.5       (Anthropic SDK — trả phí, fallback cuối)
  4. Rule-based             (offline fallback — không cần API key)

USAGE:
    from modules.ai_client import ai_generate, AIClient

    # One-shot:
    text = ai_generate("Phân tích tình trạng tồn kho: ...")

    # Stateful (cache model init):
    client = AIClient()
    text = client.generate("...")
    print(client.last_source)   # 'gemini' | 'claude' | 'rule-based'

DIAGNOSIS:
    Nếu Gemini không hoạt động, chạy:
        python -m modules.ai_client --test
    để xem chi tiết lỗi và source được chọn.
"""
from __future__ import annotations

import os
import time
from typing import Optional

# ── Constants ──────────────────────────────────────────────────────────────────
GEMINI_MODEL   = 'gemini-2.5-flash-lite'       # Tested working 2026-05-13
GROQ_MODEL     = 'llama-3.3-70b-versatile'     # Free tier, fast
CLAUDE_MODEL   = 'claude-haiku-4-5-20251001'
MAX_TOKENS     = 800
TEMPERATURE    = 0.3                            # Low temp = consistent, factual
TIMEOUT_S      = 20
# Prompt caching: system prompts ≥ 1024 tokens get cached (5-min TTL, -90% cost on cache hits)
CACHE_MIN_CHARS = 3800   # ~1024 tokens ≈ 3800 chars — shorter prompts not worth caching


class AIClient:
    """
    Stateful AI client với cascade Gemini → Claude → rule-based.
    Thread-safe cho single-process use (không cần lock).
    """

    def __init__(self,
                 gemini_key: Optional[str] = None,
                 groq_key: Optional[str] = None,
                 anthropic_key: Optional[str] = None,
                 prefer: str = 'gemini'):
        self._gkey    = gemini_key    or os.environ.get('GEMINI_API_KEY',    '')
        self._qkey    = groq_key      or os.environ.get('GROQ_API_KEY',      '')
        self._akey    = anthropic_key or os.environ.get('ANTHROPIC_API_KEY', '')
        self._prefer  = prefer
        self.last_source = 'none'
        self._gclient = None
        self._qclient = None
        self._aclient = None

    # ── Internal: init clients lazily ─────────────────────────────────────────

    def _get_gemini(self):
        if self._gclient is not None:
            return self._gclient
        if not self._gkey:
            return None
        try:
            from google import genai
            self._gclient = genai.Client(api_key=self._gkey)
            return self._gclient
        except ImportError:
            return None

    def _get_groq(self):
        if self._qclient is not None:
            return self._qclient
        if not self._qkey:
            return None
        try:
            from groq import Groq
            self._qclient = Groq(api_key=self._qkey)
            return self._qclient
        except ImportError:
            return None

    def _get_anthropic(self):
        if self._aclient is not None:
            return self._aclient
        if not self._akey:
            return None
        try:
            import anthropic
            self._aclient = anthropic.Anthropic(api_key=self._akey)
            return self._aclient
        except ImportError:
            return None

    # ── Internal: single-provider calls ───────────────────────────────────────

    def _call_gemini(self, prompt: str, max_tokens: Optional[int] = None) -> str:
        client = self._get_gemini()
        if client is None:
            raise RuntimeError('Gemini client not available')
        limit = max_tokens or MAX_TOKENS
        response = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=prompt,
            config={
                'temperature': TEMPERATURE,
                'max_output_tokens': limit,
            }
        )
        return response.text.strip()

    def _call_groq(self, prompt: str, system: Optional[str] = None,
                   max_tokens: Optional[int] = None) -> str:
        client = self._get_groq()
        if client is None:
            raise RuntimeError('Groq client not available')
        messages = []
        if system:
            messages.append({'role': 'system', 'content': system})
        messages.append({'role': 'user', 'content': prompt})
        response = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=messages,
            temperature=TEMPERATURE,
            max_tokens=max_tokens or MAX_TOKENS,
        )
        return response.choices[0].message.content.strip()

    def _call_claude(self, prompt: str,
                     system: Optional[str] = None,
                     max_tokens: Optional[int] = None) -> str:
        client = self._get_anthropic()
        if client is None:
            raise RuntimeError('Anthropic client not available')
        _mt = max_tokens or MAX_TOKENS
        kwargs: dict = {
            'model':      CLAUDE_MODEL,
            'max_tokens': _mt,
            'temperature': TEMPERATURE,
            'messages':   [{'role': 'user', 'content': prompt}],
        }
        if system:
            # Use prompt caching for long system prompts (≥ 1024 tokens, 5-min TTL)
            if len(system) >= CACHE_MIN_CHARS:
                kwargs['system'] = [{'type': 'text', 'text': system,
                                     'cache_control': {'type': 'ephemeral'}}]
            else:
                kwargs['system'] = system
        msg = client.messages.create(**kwargs)
        return msg.content[0].text.strip()

    # ── Public: generate with cascade ─────────────────────────────────────────

    def generate(self, prompt: str,
                 fallback: Optional[str] = None,
                 system: Optional[str] = None,
                 max_tokens: Optional[int] = None) -> str:
        """
        Generate text với cascade Gemini → Claude → fallback string.

        Args:
            prompt:     User prompt (tiếng Việt hoặc EN đều được)
            fallback:   String trả về nếu cả 2 API fail
            system:     System prompt (tùy chọn). Nếu ≥ 3800 ký tự → Claude tự động cache
            max_tokens: Override MAX_TOKENS (mặc định 800). Dùng 2000 cho structured JSON.

        Side effect:
            self.last_source = 'gemini' | 'claude' | 'rule-based'
        """
        # ── Tier 1: Gemini ──────────────────────────────────────────────────
        if self._gkey:
            try:
                _gemini_prompt = f"{system}\n\n{prompt}" if system else prompt
                result = self._call_gemini(_gemini_prompt, max_tokens=max_tokens)
                self.last_source = 'gemini'
                return result
            except Exception as e:
                _err = str(e)
                if 'RESOURCE_EXHAUSTED' in _err or '429' in _err:
                    _hint = ' (quota hết)'
                elif 'NOT_FOUND' in _err or '404' in _err:
                    _hint = f' (model {GEMINI_MODEL} không tồn tại)'
                else:
                    _hint = f' ({_err[:60]})'
                print(f'  ⚠️  Gemini fail{_hint} → fallback Groq')

        # ── Tier 2: Groq (free, fast — fallback khi Gemini hết quota) ───────
        if self._qkey:
            try:
                result = self._call_groq(prompt, system=system, max_tokens=max_tokens)
                self.last_source = 'groq'
                return result
            except Exception as e:
                _err = str(e)
                if '429' in _err or 'rate_limit' in _err.lower():
                    _hint = ' (rate limit)'
                else:
                    _hint = f' ({_err[:60]})'
                print(f'  ⚠️  Groq fail{_hint} → fallback Claude')

        # ── Tier 3: Claude Haiku (with optional prompt caching) ─────────────
        if self._akey:
            try:
                result = self._call_claude(prompt, system=system, max_tokens=max_tokens)
                self.last_source = 'claude'
                return result
            except Exception as e:
                print(f'  ⚠️  Claude fail ({str(e)[:60]}) → rule-based')

        # ── Tier 4: Rule-based fallback ─────────────────────────────────────
        self.last_source = 'rule-based'
        return fallback or '[AI không khả dụng — dùng rule-based fallback]'

    def generate_structured(self, prompt: str,
                             fallback: Optional[dict] = None,
                             system: Optional[str] = None,
                             max_tokens: Optional[int] = None) -> dict:
        """
        Generate JSON-parseable response. Returns dict.
        Nếu parse fail → wrap text vào {'narrative': text}.
        """
        import json
        text = self.generate(prompt, fallback=None, system=system, max_tokens=max_tokens)
        if not text or self.last_source == 'rule-based':
            return fallback or {}
        try:
            # Tìm JSON block trong response
            start = text.find('{')
            end   = text.rfind('}') + 1
            if start >= 0 and end > start:
                return json.loads(text[start:end])
        except Exception:
            pass
        return {'narrative': text}


# ── Module-level singleton ─────────────────────────────────────────────────────
_default_client: Optional[AIClient] = None


def _get_default() -> AIClient:
    global _default_client
    if _default_client is None:
        _default_client = AIClient()
    return _default_client


def ai_generate(prompt: str,
                fallback: Optional[str] = None,
                system: Optional[str] = None,
                max_tokens: Optional[int] = None) -> str:
    """One-shot convenience function. Dùng singleton client.
    system: system prompt — nếu ≥ 3800 ký tự thì Claude tự động cache (giảm ~90% cost).
    max_tokens: override default 800. Dùng 2000 cho structured JSON responses.
    """
    return _get_default().generate(prompt, fallback=fallback,
                                   system=system, max_tokens=max_tokens)


def ai_source() -> str:
    """Trả về source của lần generate gần nhất ('gemini'|'claude'|'rule-based')."""
    return _get_default().last_source


# ── CLI test ──────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    import sys
    print('=== AI Client Diagnostic ===')
    print(f'GEMINI_API_KEY:    {"SET" if os.environ.get("GEMINI_API_KEY") else "MISSING"}')
    print(f'GROQ_API_KEY:      {"SET" if os.environ.get("GROQ_API_KEY") else "MISSING"}')
    print(f'ANTHROPIC_API_KEY: {"SET" if os.environ.get("ANTHROPIC_API_KEY") else "MISSING"}')
    print(f'Gemini model: {GEMINI_MODEL}')
    print(f'Groq model:   {GROQ_MODEL}')
    print(f'Claude model: {CLAUDE_MODEL}')
    print()

    client = AIClient()
    prompt = (
        'Cửa hàng Vincom HCM: Sơ mi chiếm 45% tồn kho nhưng chỉ 20% doanh thu, '
        'SSR=2.25. Quần Chinos: 15% tồn, 30% doanh thu, SSR=0.5, WOS=1.8 tuần. '
        'Viết 2 câu nhận xét ngắn gọn bằng tiếng Việt.'
    )
    print('Prompt:', prompt[:80], '...')
    print()
    t0 = time.time()
    result = client.generate(prompt, fallback='[Không có AI]')
    elapsed = time.time() - t0
    print(f'Source: {client.last_source}  ({elapsed:.1f}s)')
    print('Result:', result)
