"""FastAPI main application entrypoint for Mega Pipes Quotation System."""

from __future__ import annotations

import logging
from contextlib import asynccontextmanager
from typing import AsyncGenerator

from fastapi import FastAPI, Request, Response, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from app.api.auth import router as auth_router
from app.api.quotations import router as quotations_router
from app.api.logistics import router as logistics_router
from app.core.config import settings, setup_logging
from app.core.security import hash_password
from app.db.models import Base, User
from app.db.session import SessionLocal, engine

logger = logging.getLogger(__name__)


def seed_pilot_user() -> None:
    """Provisions the initial pilot user if not already present in the database."""
    if not settings.PILOT_USER_EMAIL:
        return
    db = SessionLocal()
    try:
        existing = db.query(User).filter(User.email == settings.PILOT_USER_EMAIL.lower().strip()).first()
        if not existing:
            # If a password hash is provided in env, use it; otherwise generate from default for dev
            pwd_hash = settings.PILOT_USER_PASSWORD_HASH
            if not pwd_hash:
                pwd_hash = hash_password("PilotMegaPipes2026!")

            pilot = User(
                email=settings.PILOT_USER_EMAIL.lower().strip(),
                hashed_password=pwd_hash,
                full_name=settings.PILOT_USER_FULL_NAME,
                is_active=True,
                is_superuser=True,
            )
            db.add(pilot)
            db.commit()
            logger.info(f"Provisioned pilot user: {settings.PILOT_USER_EMAIL}")
    except Exception as e:
        logger.error(f"Error provisioning pilot user: {e}")
        db.rollback()
    finally:
        db.close()


@asynccontextmanager
async def lifespan(app: FastAPI) -> AsyncGenerator[None, None]:
    """Lifespan context manager for startup and shutdown hooks."""
    # 1. Setup logging and log filter
    setup_logging()
    logger.info("Initializing Mega Pipes Quotation System API...")

    # 2. Validate production security settings
    settings.validate_production_settings()

    # 3. Create tables in development / test environments if needed
    if settings.ENVIRONMENT.lower() in ("development", "test"):
        Base.metadata.create_all(bind=engine)

    # 4. Seed pilot user if needed
    seed_pilot_user()

    yield

    logger.info("Shutting down Mega Pipes Quotation System API...")


app = FastAPI(
    title="Mega Pipes Quotation System API",
    description="Secure, production-grade calculation and quotation platform for Mega Pipes.",
    version="1.0.0",
    docs_url="/docs" if settings.ENVIRONMENT.lower() != "production" else None,
    redoc_url="/redoc" if settings.ENVIRONMENT.lower() != "production" else None,
    lifespan=lifespan,
)

# CORS Middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.CORS_ORIGINS,
    allow_credentials=True,
    allow_methods=["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition", "X-CSRF-Token"],
)


# Security Headers Middleware
@app.middleware("http")
async def add_security_headers(request: Request, call_next) -> Response:
    """Adds robust HTTP security headers to all responses."""
    response = await call_next(request)
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    response.headers["Content-Security-Policy"] = "default-src 'self'"
    return response


# Global Exception Handler: prevents leakage of stack traces or sensitive credentials
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception) -> JSONResponse:
    logger.error(f"Unhandled error processing {request.method} {request.url.path}: {exc}", exc_info=True)
    return JSONResponse(
        status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
        content={"detail": "An internal server error occurred. Please contact system administrator."},
    )


# Health Check
@app.get("/health", tags=["Monitoring"])
async def health_check() -> dict:
    return {
        "status": "healthy",
        "service": "megapipes-quotation-api",
        "environment": settings.ENVIRONMENT,
        "version": "1.0.0",
    }


# Include Routers
app.include_router(auth_router)
app.include_router(quotations_router)
app.include_router(logistics_router)
