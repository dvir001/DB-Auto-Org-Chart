"""File handling helpers for SimpleOrgChart."""

from __future__ import annotations

import logging

try:
    from PIL import Image  # type: ignore
except ImportError:  # pragma: no cover - optional dependency
    Image = None  # type: ignore

logger = logging.getLogger(__name__)


def validate_image_file(file_obj) -> bool:
    """Validate that an uploaded file is a safe image."""
    if not Image:
        logger.warning("Pillow not available, skipping image validation")
        return True

    try:
        image = Image.open(file_obj)
        image.verify()

        file_obj.seek(0)
        image = Image.open(file_obj)
        if image.width > 2000 or image.height > 2000:
            return False

        file_obj.seek(0)
        return True
    except Exception as error:  # noqa: BLE001 - validation best-effort
        logger.error("Image validation failed: %s", error)
        return False


__all__ = ["validate_image_file"]
