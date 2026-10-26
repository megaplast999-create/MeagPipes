"""Security, password hashing, and session management module."""

from __future__ import annotations

import secrets
import time
from datetime import datetime, timezone
from typing import Any, Dict, Optional

import bcrypt
import jwt
from fastapi import HTTPException, Request, Response, status

from app.core.config import settings

ALGORITHM = "HS256"


class AuthSecurityError(HTTPException):
    """Generic authentication exception preventing username enumeration."""

    def __init__(self, detail: str = "Invalid email or password"):
        super().__init__(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=detail,
            headers={"WWW-Authenticate": "Bearer"},
        )


class SessionExpiredError(HTTPException):
    """Session idle or absolute expiration exception."""

    def __init__(self, detail: str = "Session has expired"):
        super().__init__(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=detail,
            headers={"WWW-Authenticate": "Bearer"},
        )


class CSRFError(HTTPException):
    """CSRF token mismatch or absent exception."""

    def __init__(self, detail: str = "Invalid or missing CSRF token"):
        super().__init__(
            status_code=status.HTTP_403_FORBIDDEN,
            detail=detail,
        )


def hash_password(password: str) -> str:
    """Hashes a password using bcrypt with a generated salt."""
    salt = bcrypt.gensalt(rounds=12)
    return bcrypt.hashpw(password.encode("utf-8"), salt).decode("utf-8")


def verify_password(plain_password: str, hashed_password: str) -> bool:
    """Verifies a plain password against a bcrypt hash in constant time."""
    try:
        return bcrypt.checkpw(plain_password.encode("utf-8"), hashed_password.encode("utf-8"))
    except Exception:
        return False


def generate_csrf_token() -> str:
    """Generates a cryptographically strong random CSRF token."""
    return secrets.token_hex(24)


def create_session_token(
    user_id: int,
    email: str,
    csrf_token: Optional[str] = None,
    idle_minutes: Optional[int] = None,
    absolute_hours: Optional[int] = None,
) -> tuple[str, str]:
    """Creates a signed session token encoding user claims, absolute expiry, and idle timeout."""
    now = int(time.time())
    idle_limit = (idle_minutes or settings.SESSION_IDLE_TIMEOUT_MINUTES) * 60
    absolute_limit = (absolute_hours or settings.SESSION_ABSOLUTE_LIFETIME_HOURS) * 3600
    csrf = csrf_token or generate_csrf_token()

    payload: Dict[str, Any] = {
        "sub": str(user_id),
        "email": email,
        "iat": now,
        "exp": now + absolute_limit,
        "last_active": now,
        "idle_timeout_seconds": idle_limit,
        "csrf": csrf,
    }

    token = jwt.encode(payload, settings.SECRET_KEY, algorithm=ALGORITHM)
    return token, csrf


def decode_and_validate_session(
    token: str,
    current_time: Optional[int] = None,
) -> Dict[str, Any]:
    """Decodes and validates a session token against absolute expiry and idle timeout."""
    now = current_time if current_time is not None else int(time.time())
    try:
        payload = jwt.decode(
            token,
            settings.SECRET_KEY,
            algorithms=[ALGORITHM],
            options={"verify_exp": False},
        )
    except jwt.InvalidTokenError:
        raise AuthSecurityError("Invalid session credentials")

    exp = payload.get("exp", 0)
    if now > exp:
        raise SessionExpiredError("Session absolute lifetime has expired")

    last_active = payload.get("last_active", 0)
    idle_limit = payload.get("idle_timeout_seconds", settings.SESSION_IDLE_TIMEOUT_MINUTES * 60)

    if (now - last_active) > idle_limit:
        raise SessionExpiredError("Session idle timeout exceeded (30 minutes inactivity)")

    return payload


def update_session_activity(payload: Dict[str, Any], current_time: Optional[int] = None) -> str:
    """Refreshes the last_active timestamp while preserving absolute expiry and user claims."""
    now = current_time if current_time is not None else int(time.time())
    new_payload = dict(payload)
    new_payload["last_active"] = now
    return jwt.encode(new_payload, settings.SECRET_KEY, algorithm=ALGORITHM)


def set_session_cookies(response: Response, session_token: str, csrf_token: str) -> None:
    """Attaches secure HTTP-only session cookie and companion CSRF cookie to response."""
    is_prod = settings.ENVIRONMENT.lower() == "production"
    response.set_cookie(
        key=settings.COOKIE_NAME,
        value=session_token,
        httponly=True,
        secure=settings.SESSION_COOKIE_SECURE or is_prod,
        samesite=settings.SESSION_COOKIE_SAMESITE,
        max_age=settings.SESSION_ABSOLUTE_LIFETIME_HOURS * 3600,
        path="/",
    )
    # The CSRF cookie is readable by frontend Javascript to send back in X-CSRF-Token header
    response.set_cookie(
        key=settings.CSRF_COOKIE_NAME,
        value=csrf_token,
        httponly=False,
        secure=settings.SESSION_COOKIE_SECURE or is_prod,
        samesite=settings.SESSION_COOKIE_SAMESITE,
        max_age=settings.SESSION_ABSOLUTE_LIFETIME_HOURS * 3600,
        path="/",
    )


def clear_session_cookies(response: Response) -> None:
    """Clears authentication and CSRF cookies on logout."""
    is_prod = settings.ENVIRONMENT.lower() == "production"
    response.delete_cookie(
        key=settings.COOKIE_NAME,
        httponly=True,
        secure=settings.SESSION_COOKIE_SECURE or is_prod,
        samesite=settings.SESSION_COOKIE_SAMESITE,
        path="/",
    )
    response.delete_cookie(
        key=settings.CSRF_COOKIE_NAME,
        httponly=False,
        secure=settings.SESSION_COOKIE_SECURE or is_prod,
        samesite=settings.SESSION_COOKIE_SAMESITE,
        path="/",
    )


def validate_csrf(request: Request, session_payload: Dict[str, Any]) -> None:
    """Validates the CSRF token from the X-CSRF-Token header against the session payload for mutating requests."""
    if request.method in ("POST", "PUT", "PATCH", "DELETE"):
        header_csrf = request.headers.get("X-CSRF-Token")
        expected_csrf = session_payload.get("csrf")
        if not header_csrf or not expected_csrf or not secrets.compare_digest(header_csrf, expected_csrf):
            raise CSRFError("Missing or invalid CSRF token")
