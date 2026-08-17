"""
config.py  –  إعدادات EGX AI Scanner v2
==========================================
الاستراتيجية: SMA200 + Breakout/Pullback (مطابق للـ Pine Script)

ملاحظة مهمة: التوكن ومعرّف الشات يُقرأان من متغيرات البيئة (.env)
ولا يجب أبداً كتابتهما هنا مباشرة أو رفعهما على GitHub.
"""

import os
from dotenv import load_dotenv

load_dotenv()  # يقرأ القيم من ملف .env المحلي (غير مرفوع على GitHub)

CONFIG = {
    # ─── تليجرام ──────────────────────────────────────────────────
    "telegram_token":   os.getenv("TELEGRAM_BOT_TOKEN", ""),
    "telegram_chat_id": os.getenv("TELEGRAM_CHAT_ID", ""),

    # ─── Trend Filter ─────────────────────────────────────────────
    "sma_length":          200,    # SMA200 — فلتر الاتجاه الرئيسي

    # ─── Momentum ─────────────────────────────────────────────────
    "rsi_length":          14,
    "rsi_upper":           73,
    "rsi_lower":           42,

    # ─── Risk Management ──────────────────────────────────────────
    "atr_length":          14,
    "atr_mult":            2.8,
    "tp_atr_mult":         4.2,

    # ─── Volume ───────────────────────────────────────────────────
    "vol_lookback":        20,
    "vol_multiplier":      1.6,

    # ─── Breakout ─────────────────────────────────────────────────
    "breakout_lookback":   18,

    # ─── ADX ──────────────────────────────────────────────────────
    "adx_length":          14,
    "adx_threshold":       19,

    # ─── Pullback Mode ────────────────────────────────────────────
    "use_pullback_mode":   True,   # ارتداد من SMA200
    "pullback_sensitivity": 3,
}
