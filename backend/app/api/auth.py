"""Authentication and session management API routes."""

from __future__ import annotations

import logging
from typing import Any, Dict, Optional

from fastapi import APIRouter, Depends, HTTPException, Request, Response, status
from pydantic import BaseModel, EmailStr
from sqlalchemy.orm import Session

from app.core.config import settings
from app.core.rate_limit import RateLimiter, get_client_ip, get_redis_client
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
    clean_email = req.email.lower().strip()
    lock_key = f"lockout:user:{clean_email}"
    failed_key = f"failed_attempts:user:{clean_email}"

    redis_client = get_redis_client()

    # 1. Enforce Account Lockout (5 consecutive failures = 15-minute lockout)
    try:
        if redis_client.get(lock_key):
            ttl = redis_client.ttl(lock_key)
            mins = max(1, ttl // 60) if (ttl and ttl > 0) else settings.ACCOUNT_LOCKOUT_DURATION_MINUTES
            raise HTTPException(
                status_code=status.HTTP_423_LOCKED,
                detail=f"Account is temporarily locked due to 5 consecutive failed login attempts. Try again in {mins} minutes.",
            )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Redis lockout check failed: {e}")
        # Fail closed policy on security infrastructure outage
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="Security verification service temporarily unavailable",
        )

    # Generic lookup: prevents timing and username disclosure
    user = db.query(User).filter(User.email == clean_email).first()
    if not user or not user.is_active or not verify_password(req.password, user.hashed_password):
        # Record failed attempt counter in Redis
        try:
            pipe = redis_client.pipeline()
            pipe.incr(failed_key)
            pipe.expire(failed_key, settings.ACCOUNT_LOCKOUT_DURATION_MINUTES * 60)
            res = pipe.execute()
            failures = res[0]
            if failures >= settings.ACCOUNT_LOCKOUT_MAX_FAILURES:
                redis_client.set(lock_key, "locked", ex=settings.ACCOUNT_LOCKOUT_DURATION_MINUTES * 60)
                redis_client.delete(failed_key)
                logger.warning(
                    f"Account locked: {clean_email} reached {failures} failed login attempts from IP {client_ip}."
                )
                raise HTTPException(
                    status_code=status.HTTP_423_LOCKED,
                    detail="Account is temporarily locked due to 5 consecutive failed login attempts. Try again in 15 minutes.",
                )
        except HTTPException:
            raise
        except Exception as e:
            logger.error(f"Failed to record lockout attempt in Redis: {e}")

        # Record failed login attempt in audit
        audit = AuditEvent(
            event_type="LOGIN_FAILED",
            client_ip=client_ip,
            details={"email": clean_email},
        )
        db.add(audit)
        db.commit()
        raise AuthSecurityError("Invalid email or password")

    # Clear lockout counters upon successful authentication
    try:
        redis_client.delete(failed_key)
        redis_client.delete(lock_key)
    except Exception as e:
        logger.warning(f"Failed to clear lockout keys in Redis: {e}")

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
