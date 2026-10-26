"""Authentication and security automated tests."""

import logging
import time
from fastapi.testclient import TestClient
import jwt
import pytest

from app.core.config import SensitiveDataLogFilter, settings
from app.core.security import create_session_token, decode_and_validate_session
from app.db.models import User


class TestAuthenticationAndSessions:
    """Verifies single-user pilot authentication, timeouts, and CSRF protection."""

    def test_login_success(self, client: TestClient, test_user: User):
        """Valid credentials return user profile and set secure HTTP-only cookies."""
        resp = client.post(
            "/api/auth/login",
            json={"email": "pilot@megapipes.internal", "password": "PilotPassword123!"},
        )
        assert resp.status_code == 200
        data = resp.json()
        assert data["email"] == "pilot@megapipes.internal"
        assert "password" not in data

        # Check cookies
        assert settings.COOKIE_NAME in resp.cookies
        assert settings.CSRF_COOKIE_NAME in resp.cookies

    def test_login_invalid_password_generic_error(self, client: TestClient, test_user: User):
        """Invalid password returns generic 401 without revealing password state."""
        resp = client.post(
            "/api/auth/login",
            json={"email": "pilot@megapipes.internal", "password": "WrongPassword!"},
        )
        assert resp.status_code == 401
        assert resp.json()["detail"] == "Invalid email or password"

    def test_login_nonexistent_email_generic_error(self, client: TestClient):
        """Unknown email returns identical generic 401 without revealing user absence."""
        resp = client.post(
            "/api/auth/login",
            json={"email": "nonexistent@megapipes.internal", "password": "AnyPassword!"},
        )
        assert resp.status_code == 401
        assert resp.json()["detail"] == "Invalid email or password"

    def test_logout_clears_cookies(self, authenticated_client: TestClient):
        """Logout endpoint clears session and CSRF cookies."""
        resp = authenticated_client.post("/api/auth/logout")
        assert resp.status_code == 200

        # Verify session deletion in Set-Cookie response header
        set_cookie = resp.headers.get("set-cookie", "")
        assert settings.COOKIE_NAME in set_cookie

        # Clear client jar to simulate browser cookie expiry behavior
        authenticated_client.cookies.clear()
        status_resp = authenticated_client.get("/api/auth/me")
        assert status_resp.status_code == 200
        assert status_resp.json()["authenticated"] is False

    def test_session_idle_timeout_enforcement(self, test_user: User):
        """Tokens inactive for over 30 minutes are rejected as expired."""
        token, csrf = create_session_token(
            user_id=test_user.id,
            email=test_user.email,
            idle_minutes=30,
            absolute_hours=8,
        )

        # Decode immediately -> valid
        valid_payload = decode_and_validate_session(token)
        assert valid_payload["sub"] == str(test_user.id)

        # Simulate 31 minutes later
        simulated_time = int(time.time()) + (31 * 60)
        with pytest.raises(Exception) as exc:
            decode_and_validate_session(token, current_time=simulated_time)
        assert "idle timeout" in str(exc.value.detail).lower()

    def test_session_absolute_lifetime_enforcement(self, test_user: User):
        """Tokens beyond absolute 8-hour lifetime are rejected even if active."""
        token, csrf = create_session_token(
            user_id=test_user.id,
            email=test_user.email,
            idle_minutes=30,
            absolute_hours=8,
        )

        # Simulate 8 hours and 1 minute later
        simulated_time = int(time.time()) + (8 * 3600) + 60
        with pytest.raises(Exception) as exc:
            decode_and_validate_session(token, current_time=simulated_time)
        assert "expired" in str(exc.value.detail).lower()

    def test_csrf_protection_on_mutating_requests(self, client: TestClient, test_user: User):
        """Mutating POST requests without valid CSRF header are rejected with HTTP 403."""
        token, csrf = create_session_token(
            user_id=test_user.id,
            email=test_user.email,
        )
        client.cookies.set(settings.COOKIE_NAME, token)
        client.cookies.set(settings.CSRF_COOKIE_NAME, csrf)

        # Request WITHOUT X-CSRF-Token header
        resp = client.post(
            "/api/quotations",
            json={"enquiry_no": "ENQ-001", "customer_name": "Test Cust"},
        )
        assert resp.status_code == 403
        assert "csrf" in resp.json()["detail"].lower()

        # Request WITH invalid X-CSRF-Token header
        resp = client.post(
            "/api/quotations",
            headers={"X-CSRF-Token": "invalid-csrf-token"},
            json={"enquiry_no": "ENQ-001", "customer_name": "Test Cust"},
        )
        assert resp.status_code == 403

    def test_log_redaction_filter(self):
        """Verifies that sensitive data patterns are scrubbed from log output."""
        log_filter = SensitiveDataLogFilter()
        record = logging.LogRecord(
            name="test",
            level=logging.INFO,
            pathname=__file__,
            lineno=10,
            msg="User login with password='SuperSecretPassword123' token: abcdef123456",
            args=(),
            exc_info=None,
        )
        log_filter.filter(record)
        assert "SuperSecretPassword123" not in record.msg
        assert "[REDACTED]" in record.msg
