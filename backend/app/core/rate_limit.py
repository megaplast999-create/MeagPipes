"""Redis-backed sliding-window rate limiter with strict Fail-Closed security."""

from __future__ import annotations

import logging
import time
import uuid
from typing import Callable, Optional

from fastapi import HTTPException, Request, Response, status
import redis

from app.core.config import settings

logger = logging.getLogger(__name__)

# Global Redis client connection pool (lazily initialized or injected for testing)
_redis_client: Optional[redis.Redis] = None


def get_redis_client() -> redis.Redis:
    """Returns the shared Redis client or creates one from configuration."""
    global _redis_client
    if _redis_client is None:
        _redis_client = redis.from_url(
            settings.REDIS_URL,
            decode_responses=True,
            socket_timeout=2.0,
            socket_connect_timeout=2.0,
        )
    return _redis_client


def set_redis_client(client: Optional[redis.Redis]) -> None:
    """Sets or overrides the shared Redis client (used for tests and mocking)."""
    global _redis_client
    _redis_client = client


import ipaddress


def is_trusted_proxy(ip_str: str) -> bool:
    """Checks whether the connecting IP is a configured trusted reverse proxy or within a trusted subnet."""
    if not ip_str:
        return False
    if "*" in settings.TRUSTED_PROXIES:
        return True
    try:
        ip_obj = ipaddress.ip_address(ip_str)
    except ValueError:
        return False

    for trusted in settings.TRUSTED_PROXIES:
        trusted = trusted.strip()
        if not trusted:
            continue
        try:
            if "/" in trusted:
                if ip_obj in ipaddress.ip_network(trusted, strict=False):
                    return True
            elif ip_obj == ipaddress.ip_address(trusted):
                return True
        except ValueError:
            continue
    return False


def get_client_ip(request: Request) -> str:
    """Extracts client IP safely, trusting forwarded headers only from configured trusted proxies (e.g. Render/Vercel)."""
    direct_ip = request.client.host if request.client else "127.0.0.1"

    # Only inspect forwarded headers if the direct connecting host is a trusted proxy
    if is_trusted_proxy(direct_ip):
        # 1. Standard X-Forwarded-For chain: client, proxy1, proxy2...
        forwarded = request.headers.get("X-Forwarded-For")
        if forwarded:
            ips = [ip.strip() for ip in forwarded.split(",") if ip.strip()]
            if ips:
                return ips[0]

        # 2. X-Real-IP header (passed by Vercel/Nginx)
        real_ip = request.headers.get("X-Real-IP")
        if real_ip and real_ip.strip():
            return real_ip.strip()

        # 3. Cloudflare / Edge proxy header
        cf_ip = request.headers.get("CF-Connecting-IP")
        if cf_ip and cf_ip.strip():
            return cf_ip.strip()

    return direct_ip


class RateLimiter:
    """Sliding-window rate limiter dependency with strict Fail-Closed enforcement."""

    def __init__(
        self,
        action: str,
        max_requests: int,
        window_seconds: int = 60,
        key_generator: Optional[Callable[[Request], str]] = None,
    ):
        self.action = action
        self.max_requests = max_requests
        self.window_seconds = window_seconds
        self.key_generator = key_generator

    async def __call__(self, request: Request, response: Response) -> None:
        client_ip = get_client_ip(request)
        if self.key_generator:
            ident = self.key_generator(request)
        else:
            ident = client_ip

        rate_key = f"rate_limit:{self.action}:{ident}"
        now = time.time()
        clear_before = now - self.window_seconds

        client = get_redis_client()

        try:
            # Fail closed: ping Redis to verify connectivity
            # Execute sliding-window transaction in Redis
            pipe = client.pipeline()
            pipe.zremrangebyscore(rate_key, "-inf", clear_before)
            pipe.zcard(rate_key)
            results = pipe.execute()

            current_count = results[1]

            if current_count >= self.max_requests:
                # Add headers for rate limit info
                response.headers["Retry-After"] = str(self.window_seconds)
                logger.warning(
                    f"Rate limit exceeded for action '{self.action}' by identifier {ident}. "
                    f"Count={current_count}, Limit={self.max_requests}"
                )
                raise HTTPException(
                    status_code=status.HTTP_429_TOO_MANY_REQUESTS,
                    detail=f"Too many requests for action '{self.action}'. Please try again later.",
                    headers={"Retry-After": str(self.window_seconds)},
                )

            # Record this request in the sliding window
            item_id = f"{now}:{uuid.uuid4().hex[:8]}"
            pipe = client.pipeline()
            pipe.zadd(rate_key, {item_id: now})
            pipe.expire(rate_key, self.window_seconds + 5)
            pipe.execute()

        except HTTPException:
            # Re-raise explicit HTTP exceptions (e.g. 429)
            raise
        except (redis.RedisError, ConnectionError, TimeoutError, OSError) as e:
            # FAIL CLOSED: If Redis is unavailable, fail closed with HTTP 503
            logger.error(f"Redis rate limiting failure for action '{self.action}': {e}. Failing closed.")
            raise HTTPException(
                status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
                detail="Rate limit service unavailable; operation rejected for security",
            )
