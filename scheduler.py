"""
scheduler.py  –  تشغيل السكانر تلقائياً كل يوم
================================================
يشغّل الفحص بعد إغلاق البورصة المصرية (3:15 مساءً بتوقيت القاهرة)
"""

import asyncio
import logging
import schedule
import time
from datetime import datetime
import pytz

from egx_scanner import main as run_scanner

log = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

CAIRO_TZ = pytz.timezone("Africa/Cairo")


def is_trading_day() -> bool:
    """البورصة المصرية: الأحد → الخميس"""
    now = datetime.now(CAIRO_TZ)
    return now.weekday() < 5  # 0=Mon … 4=Fri → EGX: Sun–Thu mapped as 6,0,1,2,3


def job():
    now = datetime.now(CAIRO_TZ)
    log.info(f"⏰  Scheduler triggered at {now.strftime('%Y-%m-%d %H:%M %Z')}")
    if not is_trading_day():
        log.info("⛔  ليس يوم تداول – تم التخطي.")
        return
    asyncio.run(run_scanner())


def main():
    log.info("🚀  EGX Scheduler started")
    log.info("   سيعمل كل يوم عمل الساعة 3:20 مساءً (بعد إغلاق البورصة)")

    # الساعة 15:20 بتوقيت القاهرة يومياً
    schedule.every().day.at("15:20").do(job)

    # يمكنك إضافة وقت إضافي
    # schedule.every().day.at("10:30").do(job)  # فحص منتصف الجلسة

    log.info("⏳  Waiting for next scheduled run ...")
    while True:
        schedule.run_pending()
        time.sleep(30)


if __name__ == "__main__":
    # تشغيل فوري للتجربة:
    # asyncio.run(run_scanner())

    # تشغيل الـ scheduler:
    main()
