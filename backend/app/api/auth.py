"""Authentication and session management API routes."""

from __future__ import annotations

import logging
from typing import Any, Dict, Optional

from fastapi import APIRouter, Depends, HTTPException, Request, Response, status
from pydantic import BaseModel, EmailStr
from sqlalchemy.orm import Session

from app.core.config import settings
from app.core.rate_limit import RateLimiter, get_client_ip
from app.core.security import (
    AuthSecurityError,
    clear_session_cookies,
    create_session_token,
    decode_and_validate_session,
    set_session_cookies,
    update_session_activity,
    validate_csrf,
    verify_password,
)
from app.db.models import AuditEvent, User
from app.db.session import get_db

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/auth", tags=["Authentication"])


class LoginRequest(BaseModel):
    email: EmailStr
    password: str


class UserResponse(BaseModel):
    id: int
    email: str
    full_name: Optional[str] = None
    is_active: bool

    model_config = {"from_attributes": True}


class SessionStatusResponse(BaseModel):
    authenticated: bool
    user: Optional[UserResponse] = None


def get_current_user_optional(
    request: Request,
    response: Response,
    db: Session = Depends(get_db),
) -> Optional[User]:
    """Inspects session cookie or Authorization header and returns User if valid, else None."""
    token = request.cookies.get(settings.COOKIE_NAME)
    if not token:
        auth_header = request.headers.get("Authorization")
        if auth_header and auth_header.startswith("Bearer "):
            token = auth_header.split(" ", 1)[1]

    if not token:
        return None

    try:
        payload = decode_and_validate_session(token)
    except Exception:
        return None

    user_id = int(payload.get("sub", 0))
    user = db.query(User).filter(User.id == user_id, User.is_active == True).first()
    if not user:
        return None

    # Sliding idle activity refresh: issue updated cookie
    try:
        refreshed_token = update_session_activity(payload)
        set_session_cookies(response, refreshed_token, payload.get("csrf", ""))
    except Exception as e:
        logger.debug(f"Failed to refresh session cookie: {e}")

    return user


def get_current_user(
    request: Request,
    response: Response,
    db: Session = Depends(get_db),
) -> User:
    """Dependency that enforces authenticated user and CSRF validation on mutating calls."""
    token = request.cookies.get(settings.COOKIE_NAME)
    if not token:
        auth_header = request.headers.get("Authorization")
        if auth_header and auth_header.startswith("Bearer "):
            token = auth_header.split(" ", 1)[1]

    if not token:
        raise AuthSecurityError("Authentication required")

    payload = decode_and_validate_session(token)

    # Validate CSRF: raises CSRFError (403) directly on mismatch!
    validate_csrf(request, payload)

    user_id = int(payload.get("sub", 0))
    user = db.query(User).filter(User.id == user_id, User.is_active == True).first()
    if not user:
        raise AuthSecurityError("Authentication required")

    try:
        refreshed_token = update_session_activity(payload)
        set_session_cookies(response, refreshed_token, payload.get("csrf", ""))
    except Exception as e:
        logger.debug(f"Failed to refresh session cookie: {e}")

    return user


login_limiter = RateLimiter(
    action="login",
    max_requests=settings.RATE_LIMIT_LOGIN_PER_MINUTE,
    window_seconds=60,
)


@router.post("/login", response_model=UserResponse)
async def login(
    req: LoginRequest,
    request: Request,
    response: Response,
    db: Session = Depends(get_db),
    _limiter: None = Depends(login_limiter),
) -> Any:
    """Authenticates pilot user, issues secure session cookies, and records audit trail."""
    client_ip = get_client_ip(request)

    # Generic lookup: prevents timing and username disclosure
    user = db.query(User).filter(User.email == req.email.lower().strip()).first()
    if not user or not user.is_active or not verify_password(req.password, user.hashed_password):
        # Record failed login attempt in audit
        audit = AuditEvent(
            event_type="LOGIN_FAILED",
            client_ip=client_ip,
            details={"email": req.email.lower().strip()},
        )
        db.add(audit)
        db.commit()
        raise AuthSecurityError("Invalid email or password")

    # Generate session and CSRF tokens
    session_token, csrf_token = create_session_token(
        user_id=user.id,
        email=user.email,
        idle_minutes=settings.SESSION_IDLE_TIMEOUT_MINUTES,
        absolute_hours=settings.SESSION_ABSOLUTE_LIFETIME_HOURS,
    )

    set_session_cookies(response, session_token, csrf_token)

    # Record successful login in audit log
    audit = AuditEvent(
        event_type="LOGIN_SUCCESS",
        user_id=user.id,
        client_ip=client_ip,
        details={"email": user.email},
    )
    db.add(audit)
    db.commit()

    return user


@router.post("/logout")
async def logout(
    request: Request,
    response: Response,
    db: Session = Depends(get_db),
    current_user: Optional[User] = Depends(get_current_user_optional),
) -> Dict[str, str]:
    """Clears session and CSRF cookies, recording logout audit event."""
    client_ip = get_client_ip(request)
    if current_user:
        audit = AuditEvent(
            event_type="LOGOUT",
            user_id=current_user.id,
            client_ip=client_ip,
            details={"email": current_user.email},
        )
        db.add(audit)
        db.commit()

    clear_session_cookies(response)
    return {"message": "Successfully logged out"}


@router.get("/me", response_model=SessionStatusResponse)
async def get_session_status(
    current_user: Optional[User] = Depends(get_current_user_optional),
) -> Any:
    """Returns current authentication state and authorized user profile."""
    if not current_user:
        return SessionStatusResponse(authenticated=False, user=None)
    return SessionStatusResponse(authenticated=True, user=UserResponse.model_validate(current_user))
