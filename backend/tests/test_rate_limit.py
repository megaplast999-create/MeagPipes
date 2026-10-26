"""Redis-backed rate limiting and Fail-Closed security tests."""

from unittest.mock import MagicMock
from fastapi.testclient import TestClient
import pytest
import redis

from app.core.config import settings
from app.core.rate_limit import RateLimiter, get_client_ip, set_redis_client
from app.db.models import User


class TestRateLimitingAndFailClosed:
    """Verifies Redis sliding-window enforcement and fail-closed security policy."""

    def test_login_rate_limiting_enforced(self, client: TestClient, test_user: User):
        """Reaching the login limit triggers HTTP 429 with Retry-After."""
        limit = settings.RATE_LIMIT_LOGIN_PER_MINUTE

        # Perform allowed requests across unique emails to test IP rate limiting
        for i in range(limit):
            resp = client.post(
                "/api/auth/login",
                json={"email": f"pilot_attempt_{i}@megapipes.internal", "password": "WrongPassword!"},
            )
            assert resp.status_code == 401

        # The next attempt exceeds the limit and must be rejected with 429
        exceeded_resp = client.post(
            "/api/auth/login",
            json={"email": "pilot_exceeded@megapipes.internal", "password": "WrongPassword!"},
        )
        assert exceeded_resp.status_code == 429
        assert "Too many requests" in exceeded_resp.json()["detail"]
        assert "Retry-After" in exceeded_resp.headers

    def test_redis_fail_closed_policy(self, client: TestClient, fake_redis):
        """If Redis is unavailable or throws an error, rate-limited operations return HTTP 503."""
        broken_redis = MagicMock()
        broken_redis.pipeline.side_effect = redis.ConnectionError("Redis cluster unreachable")
        broken_redis.get.side_effect = redis.ConnectionError("Redis cluster unreachable")
        set_redis_client(broken_redis)

        try:
            resp = client.post(
                "/api/auth/login",
                json={"email": "pilot@megapipes.internal", "password": "AnyPassword"},
            )
            # MUST fail closed: 503 Service Unavailable, NEVER silent pass
            assert resp.status_code == 503
            assert "Service temporarily unavailable" in resp.json()["detail"] or "unavailable" in resp.json()["detail"].lower()
        finally:
            set_redis_client(fake_redis)

    def test_redis_recovery(self, client: TestClient, fake_redis):
        """When Redis connectivity is restored, operations resume normally."""
        # 1. Break Redis
        broken_redis = MagicMock()
        broken_redis.pipeline.side_effect = redis.ConnectionError("Connection lost")
        broken_redis.get.side_effect = redis.ConnectionError("Connection lost")
        set_redis_client(broken_redis)

        resp1 = client.post("/api/auth/login", json={"email": "p@test.com", "password": "p"})
        assert resp1.status_code == 503

        # 2. Restore Redis
        set_redis_client(fake_redis)

        resp2 = client.post("/api/auth/login", json={"email": "p@test.com", "password": "p"})
        # Fails with 401 (normal auth failure), NOT 503
        assert resp2.status_code == 401

    def test_proxy_ip_spoofing_prevention(self):
        """Only explicitly trusted proxies can forward client IP addresses."""
        class MockRequest:
            def __init__(self, host: str, xff: str | None = None, x_real: str | None = None):
                self.client = MagicMock(host=host)
                headers = {}
                if xff:
                    headers["X-Forwarded-For"] = xff
                if x_real:
                    headers["X-Real-IP"] = x_real
                self.headers = headers

        # 1. Untrusted client IP attempting spoofing
        req_untrusted = MockRequest(host="198.51.100.1", xff="203.0.113.50")
        assert get_client_ip(req_untrusted) == "198.51.100.1"

        # 2. Trusted proxy forwarding legitimate client IP
        req_trusted = MockRequest(host="127.0.0.1", xff="203.0.113.50, 10.0.0.1")
        assert get_client_ip(req_trusted) == "203.0.113.50"

    def test_real_header_chain_render_and_vercel(self):
        """Tests real multi-hop proxy chains behind Vercel Edge and Render reverse proxies."""
        class MockRequest:
            def __init__(self, host: str, headers: dict):
                self.client = MagicMock(host=host)
                self.headers = headers

        # Scenario 1: Vercel frontend rewrite -> Render internal load balancer (10.0.4.15 in 10.0.0.0/8)
        # Header chain: client IP (198.51.100.77), Vercel edge proxy (76.76.21.21), Render internal proxy (10.0.4.15)
        headers_vercel_render = {
            "X-Forwarded-For": "198.51.100.77, 76.76.21.21, 10.0.4.15",
            "X-Forwarded-Proto": "https",
            "X-Forwarded-Host": "megapipes.vercel.app",
        }
        req_render = MockRequest(host="10.0.4.15", headers=headers_vercel_render)
        assert get_client_ip(req_render) == "198.51.100.77"

        # Scenario 2: Vercel providing X-Real-IP behind trusted proxy (172.17.0.2 in 172.16.0.0/12)
        headers_x_real = {
            "X-Real-IP": "203.0.113.195",
            "X-Forwarded-Proto": "https",
        }
        req_real = MockRequest(host="172.17.0.2", headers=headers_x_real)
        assert get_client_ip(req_real) == "203.0.113.195"

        # Scenario 3: Untrusted direct connection (e.g. 8.8.8.8) attempting to fake Vercel headers
        headers_spoof = {
            "X-Forwarded-For": "198.51.100.77, 76.76.21.21",
            "X-Real-IP": "1.2.3.4",
        }
        req_spoof = MockRequest(host="8.8.8.8", headers=headers_spoof)
        assert get_client_ip(req_spoof) == "8.8.8.8"

    def test_account_lockout_after_five_failed_attempts(self, client: TestClient, test_user: User):
        """5 consecutive failed login attempts lock the account for 15 minutes (HTTP 423)."""
        target_email = "target_lockout@megapipes.internal"

        # 1. First 4 failed attempts return standard 401
        for attempt in range(1, 5):
            resp = client.post(
                "/api/auth/login",
                json={"email": target_email, "password": "WrongPassword123!"},
            )
            assert resp.status_code == 401, f"Attempt {attempt} should return 401"
            assert "Invalid email or password" in resp.json()["detail"]

        # 2. 5th failed attempt triggers HTTP 423 Locked
        fifth_resp = client.post(
            "/api/auth/login",
            json={"email": target_email, "password": "WrongPassword123!"},
        )
        assert fifth_resp.status_code == 423
        assert "Account is temporarily locked" in fifth_resp.json()["detail"]
        assert "15 minutes" in fifth_resp.json()["detail"]

        # 3. 6th attempt while locked also immediately returns HTTP 423
        locked_resp = client.post(
            "/api/auth/login",
            json={"email": target_email, "password": "AnyPassword"},
        )
        assert locked_resp.status_code == 423
        assert "Account is temporarily locked" in locked_resp.json()["detail"]

        # 4. Another unrelated account is NOT locked out
        other_resp = client.post(
            "/api/auth/login",
            json={"email": "pilot@megapipes.internal", "password": "PilotPassword123!"},
        )
        assert other_resp.status_code == 200
        assert other_resp.json()["email"] == "pilot@megapipes.internal"
