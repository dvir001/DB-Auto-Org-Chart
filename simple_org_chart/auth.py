"""Authentication helpers and decorators for SimpleOrgChart."""

import re
from functools import wraps
from typing import Any, Callable, Dict

from flask import jsonify, redirect, request, session, url_for

NextHandler = Callable[..., Any]
_next_path_pattern = re.compile(r"^[A-Za-z0-9_\-/]*$")


def sanitize_next_path(raw_value: str | None) -> str:
    """Sanitize the "next" query parameter to prevent open redirects."""
    if not raw_value:
        return ""

    candidate = raw_value.strip()

    if candidate.startswith(("http://", "https://", "//")):
        return ""

    candidate = candidate.lstrip("/")

    if not _next_path_pattern.fullmatch(candidate):
        return ""

    return candidate


def require_auth(func: NextHandler) -> NextHandler:
    """API decorator that ensures the caller is authenticated."""

    @wraps(func)
    def wrapper(*args: Any, **kwargs: Dict[str, Any]):
        if not session.get("authenticated"):
            return jsonify({"error": "Authentication required"}), 401
        return func(*args, **kwargs)

    return wrapper


def login_required(func: NextHandler) -> NextHandler:
    """Route decorator that redirects to the login page when unauthenticated."""

    @wraps(func)
    def wrapper(*args: Any, **kwargs: Dict[str, Any]):
        if not session.get("authenticated"):
            desired_path = sanitize_next_path(request.path)
            params: Dict[str, Any] = {"next": desired_path} if desired_path else {}
            return redirect(url_for("login", **params))
        return func(*args, **kwargs)

    return wrapper


__all__ = ["login_required", "require_auth", "sanitize_next_path"]
