"""Application configuration and settings using Pydantic Settings."""

from __future__ import annotations

import json
import logging
import re
from typing import Any, List
from pydantic import field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Central configuration for Mega Pipes Quotation System backend."""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=True,
        extra="ignore",
    )

    ENVIRONMENT: str = "development"
    LOG_LEVEL: str = "INFO"

    # PostgreSQL Database URL
    DATABASE_URL: str = "postgresql://postgres:postgres@localhost:5432/megapipes"

    # Redis URL
    REDIS_URL: str = "redis://localhost:6379/0"

    # Secret Key for Signing Sessions / JWT Tokens
    SECRET_KEY: str = "pilot-insecure-dev-secret-key-MegaPipes-999-not-for-production"

    # Session Constraints
    SESSION_IDLE_TIMEOUT_MINUTES: int = 30
    SESSION_ABSOLUTE_LIFETIME_HOURS: int = 8

    # Security & Cookie Attributes
    COOKIE_NAME: str = "megapipes_session"
    CSRF_COOKIE_NAME: str = "megapipes_csrf"
    SESSION_COOKIE_SECURE: bool = False  # Set True in production
    SESSION_COOKIE_SAMESITE: str = "lax"

    # CORS Allowed Origins
    CORS_ORIGINS: List[str] = ["http://localhost:3000", "http://localhost:5173"]

    # Trusted Proxies for Client IP Resolution (Supports IPs, CIDR ranges, and wildcards)
    TRUSTED_PROXIES: List[str] = [
        "127.0.0.1",
        "::1",
        "10.0.0.0/8",
        "172.16.0.0/12",
        "192.168.0.0/16",
    ]

    # Rate Limiting Parameters (Requests per minute)
    RATE_LIMIT_LOGIN_PER_MINUTE: int = 10
    RATE_LIMIT_CALCULATE_PER_MINUTE: int = 30
    RATE_LIMIT_EXPORT_PER_MINUTE: int = 20

    # Account Lockout Security (Consecutive failed attempts and duration)
    ACCOUNT_LOCKOUT_MAX_FAILURES: int = 5
    ACCOUNT_LOCKOUT_DURATION_MINUTES: int = 15

    # Pilot User Provisioning
    PILOT_USER_EMAIL: str = "pilot@megapipes.internal"
    PILOT_USER_PASSWORD_HASH: str = ""
    PILOT_USER_FULL_NAME: str = "Mega Pipes Pilot User"

    @field_validator("CORS_ORIGINS", mode="before")
    @classmethod
    def assemble_cors_origins(cls, v: Any) -> List[str]:
        if isinstance(v, str):
            if v.strip().startswith("[") and v.strip().endswith("]"):
                return json.loads(v)
            return [i.strip() for i in v.split(",") if i.strip()]
        if isinstance(v, list):
            return v
        return []

    @field_validator("TRUSTED_PROXIES", mode="before")
    @classmethod
    def assemble_trusted_proxies(cls, v: Any) -> List[str]:
        if isinstance(v, str):
            if v.strip().startswith("[") and v.strip().endswith("]"):
                return json.loads(v)
            return [i.strip() for i in v.split(",") if i.strip()]
        if isinstance(v, list):
            return v
        return ["127.0.0.1", "::1"]

    def validate_production_settings(self) -> None:
        """Fails startup safely if mandatory production security settings are absent or insecure."""
        if self.ENVIRONMENT.lower() == "production":
            insecure_keys = [
                "change-this",
                "insecure",
                "default",
                "pilot-insecure",
                "MegaPipes-999",
            ]
            if len(self.SECRET_KEY) < 32 or any(k in self.SECRET_KEY.lower() for k in insecure_keys):
                raise ValueError(
                    "FATAL: Application cannot start in production with an insecure or default SECRET_KEY. "
                    "Provide a high-entropy secret of at least 32 characters via the SECRET_KEY environment variable."
                )
            if "localhost" in self.DATABASE_URL or "postgres:postgres" in self.DATABASE_URL:
                raise ValueError(
                    "FATAL: Application cannot start in production with a default localhost DATABASE_URL. "
                    "Configure a valid production PostgreSQL instance."
                )


settings = Settings()


class SensitiveDataLogFilter(logging.Filter):
    """Log filter to scrub passwords, tokens, cookies, and secrets from all log output."""

    SENSITIVE_PATTERNS = [
        (re.compile(r"(password|pwd|secret|token|hash)[\s:=]+['\"]?([^'\",\s&]+)", re.IGNORECASE), r"\1=[REDACTED]"),
        (re.compile(r"(Authorization:\s*Bearer\s+)[^\s]+", re.IGNORECASE), r"\1[REDACTED]"),
        (re.compile(r"(Cookie:\s*[^;\r\n]*megapipes_session=)[^;\r\n]+", re.IGNORECASE), r"\1[REDACTED]"),
        (re.compile(r"postgres(?:ql)?://([^:]+):([^@]+)@", re.IGNORECASE), r"postgresql://\1:*****@"),
    ]

    def filter(self, record: logging.LogRecord) -> bool:
        if isinstance(record.msg, str):
            message = record.msg
            for pattern, repl in self.SENSITIVE_PATTERNS:
                message = pattern.sub(repl, message)
            record.msg = message
        return True


def setup_logging() -> None:
    """Configures application-wide logging with sensitive information redaction."""
    log_level = getattr(logging, settings.LOG_LEVEL.upper(), logging.INFO)
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )
    scrub_filter = SensitiveDataLogFilter()
    for handler in logging.root.handlers:
        handler.addFilter(scrub_filter)
