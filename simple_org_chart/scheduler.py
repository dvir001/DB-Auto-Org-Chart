"""Background scheduler management for SimpleOrgChart."""

from __future__ import annotations

import logging
import os
import threading
import time
from datetime import datetime
from typing import Callable, Optional

import schedule

from simple_org_chart.settings import load_settings

logger = logging.getLogger(__name__)

_scheduler_running = False
_scheduler_lock = threading.Lock()
_scheduler_thread: Optional[threading.Thread] = None
_update_callback: Optional[Callable[[], None]] = None


def configure_scheduler(update_callback: Callable[[], None]) -> None:
    """Register the callback used to refresh employee data."""
    global _update_callback
    _update_callback = update_callback


def is_scheduler_running() -> bool:
    """Return True if the background scheduler loop is active."""
    return _scheduler_running


def _ensure_callback() -> Callable[[], None]:
    if _update_callback is None:
        raise RuntimeError("Scheduler update callback has not been configured")
    return _update_callback


def _schedule_loop() -> None:
    global _scheduler_running

    try:
        update_callback = _ensure_callback()
    except RuntimeError as exc:
        logger.error(str(exc))
        _scheduler_running = False
        return

    schedule.clear()
    settings = load_settings()

    if os.environ.get("RUN_INITIAL_UPDATE", "true").lower() == "true":
        logger.info("[%s] Running initial employee data update on startup...", datetime.now())
        update_callback()

    if settings.get("autoUpdateEnabled", True):
        update_time = settings.get("updateTime", "20:00")
        schedule.every().day.at(update_time).do(update_callback)
        logger.info("Scheduled daily updates at %s", update_time)
    else:
        logger.info("Automatic updates are disabled; skipping daily schedule")

    while _scheduler_running:
        schedule.run_pending()
        time.sleep(60)


def start_scheduler() -> None:
    """Start the background scheduler thread if it is not already running."""
    global _scheduler_running, _scheduler_thread

    with _scheduler_lock:
        if _scheduler_running:
            return
        _scheduler_running = True
        _scheduler_thread = threading.Thread(target=_schedule_loop, daemon=True)
        _scheduler_thread.start()
        logger.info("Scheduler started")


def stop_scheduler() -> None:
    """Stop the background scheduler loop."""
    global _scheduler_running

    with _scheduler_lock:
        if not _scheduler_running:
            return
        _scheduler_running = False
        logger.info("Scheduler stopped")


def restart_scheduler() -> None:
    """Restart the scheduler, reloading settings and timings."""
    stop_scheduler()
    time.sleep(2)
    schedule.clear()
    start_scheduler()


__all__ = [
    "configure_scheduler",
    "is_scheduler_running",
    "restart_scheduler",
    "start_scheduler",
    "stop_scheduler",
]
