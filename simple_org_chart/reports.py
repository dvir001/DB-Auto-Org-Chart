"""Report cache helpers for SimpleOrgChart."""

from __future__ import annotations

import json
import logging
import os
from datetime import datetime, timedelta, timezone
from typing import Callable, Iterable, List, Optional, Sequence, Type

import simple_org_chart.config as app_config
from simple_org_chart.msgraph import parse_graph_datetime

logger = logging.getLogger(__name__)


MISSING_MANAGER_FILE = str(app_config.MISSING_MANAGER_FILE)
DISABLED_LICENSE_FILE = str(app_config.DISABLED_LICENSE_FILE)
DISABLED_USERS_FILE = str(app_config.DISABLED_USERS_FILE)
RECENTLY_DISABLED_FILE = str(app_config.RECENTLY_DISABLED_FILE)
RECENTLY_HIRED_FILE = str(app_config.RECENTLY_HIRED_FILE)
LAST_LOGIN_FILE = str(app_config.LAST_LOGIN_FILE)
FILTERED_LICENSE_FILE = str(app_config.FILTERED_LICENSE_FILE)
FILTERED_USERS_FILE = str(app_config.FILTERED_USERS_FILE)


class ReportCacheManager:
    """Centralised helper for loading cached report data."""

    def __init__(self, refresh_callback: Optional[Callable[[], None]] = None) -> None:
        self._refresh_callback = refresh_callback

    def load_json(
        self,
        path: str,
        *,
        refresh: bool = False,
        description: str = "report cache",
        expected_type: Optional[Type] = list,
    ):
        """Load a JSON payload from disk, optionally refreshing first."""
        if not path:
            logger.error("No path provided for %s", description)
            return [] if expected_type is list else None

        if refresh or not os.path.exists(path):
            if refresh:
                logger.info("Refreshing %s", description)
            if self._refresh_callback is not None:
                try:
                    self._refresh_callback()
                except Exception as exc:  # pragma: no cover - defensive
                    logger.error("Failed to refresh %s: %s", description, exc)
            else:
                logger.warning("No refresh callback configured; cannot refresh %s", description)

        if not os.path.exists(path):
            logger.warning("%s not found at %s", description, path)
            return [] if expected_type is list else None

        try:
            with open(path, "r", encoding="utf-8") as handle:
                data = json.load(handle)
        except json.JSONDecodeError as decode_error:
            logger.error("Failed to parse %s at %s: %s", description, path, decode_error)
            return [] if expected_type is list else None
        except Exception as error:  # pragma: no cover - I/O errors
            logger.error("Unexpected error loading %s at %s: %s", description, path, error)
            return [] if expected_type is list else None

        if expected_type is not None and not isinstance(data, expected_type):
            logger.warning("Unexpected payload type for %s; expected %s", description, expected_type.__name__)
            return [] if expected_type is list else None

        return data


def load_missing_manager_data(cache: ReportCacheManager, *, force_refresh: bool = False):
    return cache.load_json(
        MISSING_MANAGER_FILE,
        refresh=force_refresh,
        description="missing manager report cache",
    )


def load_disabled_license_data(cache: ReportCacheManager, *, force_refresh: bool = False):
    return cache.load_json(
        DISABLED_LICENSE_FILE,
        refresh=force_refresh,
        description="disabled licensed users report cache",
    )


def load_disabled_users_data(cache: ReportCacheManager, *, force_refresh: bool = False):
    return cache.load_json(
        DISABLED_USERS_FILE,
        refresh=force_refresh,
        description="disabled users report cache",
    )


def load_recently_disabled_data(cache: ReportCacheManager, *, force_refresh: bool = False):
    return cache.load_json(
        RECENTLY_DISABLED_FILE,
        refresh=force_refresh,
        description="recently disabled employees report cache",
    )


def load_recently_hired_data(cache: ReportCacheManager, *, force_refresh: bool = False):
    return cache.load_json(
        RECENTLY_HIRED_FILE,
        refresh=force_refresh,
        description="recently hired employees report cache",
    )


def load_last_login_data(cache: ReportCacheManager, *, force_refresh: bool = False):
    return cache.load_json(
        LAST_LOGIN_FILE,
        refresh=force_refresh,
        description="last sign-in report cache",
    )


def load_filtered_license_data(cache: ReportCacheManager, *, force_refresh: bool = False):
    return cache.load_json(
        FILTERED_LICENSE_FILE,
        refresh=force_refresh,
        description="filtered licensed users report cache",
    )


def load_filtered_user_data(cache: ReportCacheManager, *, force_refresh: bool = False):
    return cache.load_json(
        FILTERED_USERS_FILE,
        refresh=force_refresh,
        description="filtered users report cache",
    )


def apply_disabled_filters(
    records: Optional[Sequence[dict]],
    *,
    licensed_only: bool = False,
    recent_days: Optional[int] = None,
    include_guests: bool = False,
    include_members: bool = True,
):
    if not records:
        return []

    cutoff = None
    if recent_days and recent_days > 0:
        cutoff = datetime.now(timezone.utc) - timedelta(days=recent_days)

    filtered: List[dict] = []

    for record in records:
        user_type = (record.get("userType") or "").lower()

        if user_type == "guest" and not include_guests:
            continue
        if user_type == "member" and not include_members:
            continue

        if licensed_only and (record.get("licenseCount") or 0) == 0:
            continue

        if cutoff is not None:
            observed = parse_graph_datetime(
                record.get("firstSeenDisabledAt")
                or record.get("disabledDate")
            )
            if not observed or observed < cutoff:
                continue

        filtered.append(record)

    return filtered


def calculate_license_totals(records: Optional[Iterable[dict]]):
    return sum((record.get("licenseCount") or 0) for record in records or [])


def apply_last_login_filters(
    records: Optional[Sequence[dict]],
    *,
    include_enabled: bool = True,
    include_disabled: bool = True,
    include_licensed: bool = True,
    include_unlicensed: bool = True,
    include_members: bool = True,
    include_guests: bool = True,
    include_never_signed_in: bool = True,
    inactive_days: Optional[str] = None,
):
    if not records:
        return []

    inactive_threshold = None
    require_never_signed_in = False

    if inactive_days not in (None, "", "none"):
        if isinstance(inactive_days, str) and inactive_days.lower() == "never":
            require_never_signed_in = True
        else:
            try:
                inactive_threshold = int(inactive_days)  # type: ignore[arg-type]
            except (TypeError, ValueError):
                inactive_threshold = None

    filtered: List[dict] = []

    for record in records:
        account_enabled = record.get("accountEnabled", True)
        if account_enabled and not include_enabled:
            continue
        if not account_enabled and not include_disabled:
            continue

        license_count = record.get("licenseCount") or 0
        if license_count > 0 and not include_licensed:
            continue
        if license_count == 0 and not include_unlicensed:
            continue

        user_type = (record.get("userType") or "").lower()
        if user_type == "member" and not include_members:
            continue
        if user_type == "guest" and not include_guests:
            continue

        never_signed_in = bool(record.get("neverSignedIn"))
        if never_signed_in and not include_never_signed_in:
            continue
        if require_never_signed_in and not never_signed_in:
            continue

        if inactive_threshold is not None:
            days_since = record.get("daysSinceLastActivity")
            if days_since is None or days_since < inactive_threshold:
                continue

        filtered.append(record)

    return filtered


def apply_filtered_user_filters(
    records: Optional[Sequence[dict]],
    *,
    include_enabled: bool = True,
    include_disabled: bool = True,
    include_licensed: bool = True,
    include_unlicensed: bool = True,
    include_members: bool = True,
    include_guests: bool = True,
):
    if not records:
        return []

    filtered: List[dict] = []

    for record in records:
        account_enabled = record.get("accountEnabled", True)
        if account_enabled and not include_enabled:
            continue
        if not account_enabled and not include_disabled:
            continue

        license_count = record.get("licenseCount") or 0
        if license_count > 0 and not include_licensed:
            continue
        if license_count == 0 and not include_unlicensed:
            continue

        user_type = (record.get("userType") or "").lower()

        if user_type == "guest" and not include_guests:
            continue
        if user_type == "member" and not include_members:
            continue

        filtered.append(record)

    return filtered


__all__ = [
    "ReportCacheManager",
    "apply_disabled_filters",
    "apply_filtered_user_filters",
    "apply_last_login_filters",
    "calculate_license_totals",
    "load_disabled_license_data",
    "load_disabled_users_data",
    "load_filtered_license_data",
    "load_filtered_user_data",
    "load_last_login_data",
    "load_missing_manager_data",
    "load_recently_disabled_data",
    "load_recently_hired_data",
]
