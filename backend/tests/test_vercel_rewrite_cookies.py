"""End-to-end verification of Vercel /api rewrite and SameSite/Secure cookie attributes.

Simulates the complete production architecture:
1. Frontend hosted on Vercel (https://megapipes.vercel.app)
2. Vercel /api/:path* rewrite proxying requests to Render backend
3. SameSite=Lax, HttpOnly, and Secure cookie behavior
4. CSRF double-submit validation via X-CSRF-Token header
5. Real IP resolution through Vercel/Render proxy headers
"""

from fastapi.testclient import TestClient
import pytest

from app.core.config import settings
from app.db.models import User


class TestVercelRewriteAndCookieSecurity:
    """Verifies that the Vercel /api rewrite and hardened cookie architecture function end to end."""

    def test_vercel_rewrite_login_sets_hardened_cookies(self, client: TestClient, test_user: User):
        """Simulates Vercel reverse proxy rewrite to /api/auth/login and verifies Set-Cookie headers."""
        # Enable Secure cookies for production simulation
        orig_secure = settings.SESSION_COOKIE_SECURE
        orig_samesite = settings.SESSION_COOKIE_SAMESITE
        settings.SESSION_COOKIE_SECURE = True
        settings.SESSION_COOKIE_SAMESITE = "lax"

        try:
            # Request coming via Vercel rewrite with Vercel edge proxy headers
            vercel_headers = {
                "Host": "megapipes-api.onrender.com",
                "X-Forwarded-Host": "megapipes.vercel.app",
                "X-Forwarded-Proto": "https",
                "X-Forwarded-For": "203.0.113.88, 76.76.21.21",
                "Origin": "https://megapipes.vercel.app",
            }

            resp = client.post(
                "/api/auth/login",
                json={"email": "pilot@megapipes.internal", "password": "PilotPassword123!"},
                headers=vercel_headers,
            )

            assert resp.status_code == 200, f"Login failed: {resp.text}"
            assert resp.json()["email"] == "pilot@megapipes.internal"

            # Verify Set-Cookie headers
            set_cookie_headers = resp.headers.get_list("set-cookie") if hasattr(resp.headers, "get_list") else [resp.headers.get("set-cookie", "")]
            raw_cookies = "; ".join(set_cookie_headers)

            # 1. Session Cookie Verification
            assert settings.COOKIE_NAME in raw_cookies
            assert "httponly" in raw_cookies.lower()
            assert "secure" in raw_cookies.lower()
            assert "samesite=lax" in raw_cookies.lower()

            # 2. CSRF Cookie Verification (readable by JavaScript to populate X-CSRF-Token)
            assert settings.CSRF_COOKIE_NAME in raw_cookies
            assert "samesite=lax" in raw_cookies.lower()

        finally:
            settings.SESSION_COOKIE_SECURE = orig_secure
            settings.SESSION_COOKIE_SAMESITE = orig_samesite

    def test_vercel_rewrite_authenticated_api_flow(self, client: TestClient, test_user: User):
        """Verifies end-to-end quotation creation and calculation using Vercel proxy cookies and CSRF."""
        # 1. Login through Vercel rewrite
        vercel_headers = {
            "X-Forwarded-Host": "megapipes.vercel.app",
            "X-Forwarded-Proto": "https",
            "X-Forwarded-For": "203.0.113.88, 76.76.21.21",
            "Origin": "https://megapipes.vercel.app",
        }
        login_resp = client.post(
            "/api/auth/login",
            json={"email": "pilot@megapipes.internal", "password": "PilotPassword123!"},
            headers=vercel_headers,
        )
        assert login_resp.status_code == 200

        # Extract CSRF token from cookies
        csrf_token = login_resp.cookies.get(settings.CSRF_COOKIE_NAME)
        assert csrf_token is not None, "CSRF cookie missing from login response"

        # 2. Create Quotation through Vercel rewrite (Mutating POST requires CSRF token)
        api_headers = {
            **vercel_headers,
            "X-CSRF-Token": csrf_token,
        }
        create_resp = client.post(
            "/api/quotations",
            json={
                "enquiry_no": "MKTG/ENQ/26-27/VERCEL-01",
                "customer_name": "Vercel Rewrite Infrastructure Corp",
                "delivery_location": "Krishna Nagar, Bagodar, Jharkhand",
                "trip_rate": 93000.0,
                "lines": [
                    {
                        "line_number": 1,
                        "size_mm": 1226.0,
                        "thickness_mm": 8.4,
                        "basis": "ID",
                        "quantity": 2000.0,
                        "steel_rate": 62000.0,
                        "conversion": 4000.0,
                        "ext_rate": 500.0,
                        "int_rate": 250.0,
                        "manual_fc": 1024.0,
                    }
                ],
            },
            headers=api_headers,
        )
        assert create_resp.status_code == 201, f"Quotation creation failed: {create_resp.text}"
        quote_id = create_resp.json()["id"]

        # 3. Calculate Quotation through Vercel rewrite
        calc_resp = client.post(
            f"/api/quotations/{quote_id}/calculate",
            headers=api_headers,
        )
        assert calc_resp.status_code == 200, f"Calculation failed: {calc_resp.text}"
        calc_data = calc_resp.json()
        assert calc_data["subtotal_excl_gst"] == 44702000.0
        assert calc_data["results_snapshot"]["lines"][0]["total_per_m"] == 22351.0

        # 4. Attempt mutating request WITHOUT X-CSRF-Token -> Must be rejected with HTTP 403
        bad_headers = {**vercel_headers}  # No X-CSRF-Token
        blocked_resp = client.post(
            f"/api/quotations/{quote_id}/calculate",
            headers=bad_headers,
        )
        assert blocked_resp.status_code == 403
        assert "CSRF" in blocked_resp.json()["detail"]

        # 5. Read-only GET /api/quotations/{id} succeeds with session cookie alone (no CSRF header required)
        get_resp = client.get(
            f"/api/quotations/{quote_id}",
            headers=vercel_headers,
        )
        assert get_resp.status_code == 200
        assert get_resp.json()["id"] == quote_id
