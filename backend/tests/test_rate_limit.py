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

        # Perform allowed requests
        for _ in range(limit):
            resp = client.post(
                "/api/auth/login",
                json={"email": "pilot@megapipes.internal", "password": "WrongPassword!"},
            )
            assert resp.status_code == 401

        # The next attempt exceeds the limit and must be rejected with 429
        exceeded_resp = client.post(
            "/api/auth/login",
            json={"email": "pilot@megapipes.internal", "password": "WrongPassword!"},
        )
        assert exceeded_resp.status_code == 429
        assert "Too many requests" in exceeded_resp.json()["detail"]
        assert "Retry-After" in exceeded_resp.headers

    def test_redis_fail_closed_policy(self, client: TestClient, fake_redis):
        """If Redis is unavailable or throws an error, rate-limited operations return HTTP 503."""
        broken_redis = MagicMock()
        broken_redis.pipeline.side_effect = redis.ConnectionError("Redis cluster unreachable")
        set_redis_client(broken_redis)

        try:
            resp = client.post(
                "/api/auth/login",
                json={"email": "pilot@megapipes.internal", "password": "AnyPassword"},
            )
            # MUST fail closed: 503 Service Unavailable, NEVER silent pass
            assert resp.status_code == 503
            assert "Rate limit service unavailable" in resp.json()["detail"]
        finally:
            set_redis_client(fake_redis)

    def test_redis_recovery(self, client: TestClient, fake_redis):
        """When Redis connectivity is restored, operations resume normally."""
        # 1. Break Redis
        broken_redis = MagicMock()
        broken_redis.pipeline.side_effect = redis.ConnectionError("Connection lost")
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
            def __init__(self, host: str, xff: str | None = None):
                self.client = MagicMock(host=host)
                self.headers = {"X-Forwarded-For": xff} if xff else {}

        # 1. Untrusted client IP attempting spoofing
        req_untrusted = MockRequest(host="198.51.100.1", xff="203.0.113.50")
        assert get_client_ip(req_untrusted) == "198.51.100.1"

        # 2. Trusted proxy forwarding legitimate client IP
        req_trusted = MockRequest(host="127.0.0.1", xff="203.0.113.50, 10.0.0.1")
        assert get_client_ip(req_trusted) == "203.0.113.50"
