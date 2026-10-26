"""Quotation management, calculation lifecycle, and cross-user authorization tests."""

from fastapi.testclient import TestClient
import pytest
from sqlalchemy.orm import Session

from app.core.config import settings
from app.core.security import create_session_token
from app.db.models import Quotation, QuotationLine, User


class TestQuotationLifecycleAndAuthorization:
    """Verifies quotation creation, updating, calculation, snapshots, and authorization boundaries."""

    @pytest.fixture
    def sample_quote_payload(self):
        return {
            "enquiry_no": "MKTG/ENQ/26-27/001",
            "customer_name": "Antigravity Engineering Ltd",
            "project_name": "Pipeline Water Supply",
            "delivery_location": "Bagodar, Jharkhand",
            "distance_km": 782.0,
            "trip_rate": 93000.0,
            "calculation_mode": "REFERENCE",
            "lines": [
                {
                    "line_number": 1,
                    "size_mm": 1226.0,
                    "thickness_mm": 8.4,
                    "basis": "ID",
                    "quantity": 2000.0,
                    "unit": "Metres",
                    "pipe_length_m": 12.0,
                    "pipes_per_trailer": 4.0,
                    "steel_rate": 62000.0,
                    "conversion": 4000.0,
                    "ext_rate": 500.0,
                    "int_rate": 250.0,
                    "trip_rate": 93000.0,
                    "handling": 100.0,
                    "financial_pct": 1.0,
                    "manual_fc": 1024.0,
                }
            ],
        }

    def test_create_and_retrieve_quotation(self, authenticated_client: TestClient, sample_quote_payload):
        """User can create draft quotation and retrieve it."""
        # 1. Create Draft
        resp = authenticated_client.post("/api/quotations", json=sample_quote_payload)
        assert resp.status_code == 201
        data = resp.json()
        quote_id = data["id"]
        assert data["enquiry_no"] == "MKTG/ENQ/26-27/001"
        assert data["status"] == "DRAFT"
        assert len(data["lines"]) == 1

        # 2. Retrieve Specific Quotation
        get_resp = authenticated_client.get(f"/api/quotations/{quote_id}")
        assert get_resp.status_code == 200
        assert get_resp.json()["id"] == quote_id

        # 3. List Quotations
        list_resp = authenticated_client.get("/api/quotations")
        assert list_resp.status_code == 200
        assert any(q["id"] == quote_id for q in list_resp.json())

    def test_unauthorized_cross_user_access_blocked(
        self,
        authenticated_client: TestClient,
        client: TestClient,
        test_other_user: User,
        sample_quote_payload,
    ):
        """User B cannot access or modify User A's quotation, even with known ID."""
        # User A creates quotation
        resp = authenticated_client.post("/api/quotations", json=sample_quote_payload)
        quote_id = resp.json()["id"]

        # Authenticate client as User B
        other_token, other_csrf = create_session_token(
            user_id=test_other_user.id,
            email=test_other_user.email,
        )
        client.cookies.clear()
        client.cookies.set(settings.COOKIE_NAME, other_token)
        client.cookies.set(settings.CSRF_COOKIE_NAME, other_csrf)
        client.headers.update({"X-CSRF-Token": other_csrf})

        # User B attempts to view User A's quote -> HTTP 403 Forbidden
        view_resp = client.get(f"/api/quotations/{quote_id}")
        assert view_resp.status_code == 403

        # User B attempts to calculate User A's quote -> HTTP 403 Forbidden
        calc_resp = client.post(f"/api/quotations/{quote_id}/calculate")
        assert calc_resp.status_code == 403

    def test_quotation_server_side_calculation_and_snapshot(
        self, authenticated_client: TestClient, sample_quote_payload
    ):
        """Calculating quotation runs server-side engine and creates immutable snapshot."""
        # 1. Create Draft
        resp = authenticated_client.post("/api/quotations", json=sample_quote_payload)
        quote_id = resp.json()["id"]

        # 2. Calculate Quotation
        calc_resp = authenticated_client.post(f"/api/quotations/{quote_id}/calculate")
        assert calc_resp.status_code == 200
        snap = calc_resp.json()

        assert snap["quotation_id"] == quote_id
        assert snap["version"] == 2
        # Verify Unipro Line 1 value: 44,702,000
        assert snap["subtotal_excl_gst"] == 44702000.0
        assert snap["grand_total"] > 44702000.0

        # 3. Retrieve Latest Snapshot
        snap_resp = authenticated_client.get(f"/api/quotations/{quote_id}/snapshot")
        assert snap_resp.status_code == 200
        assert snap_resp.json()["version"] == 2

    def test_input_validation_constraints(self, authenticated_client: TestClient, sample_quote_payload):
        """Invalid inputs (negative thickness, invalid basis) are rejected with HTTP 422."""
        invalid_payload = dict(sample_quote_payload)
        invalid_payload["lines"] = [
            {
                "line_number": 1,
                "size_mm": 1226.0,
                "thickness_mm": -8.4,  # Negative thickness
                "basis": "INVALID_BASIS",
                "quantity": 2000.0,
            }
        ]
        resp = authenticated_client.post("/api/quotations", json=invalid_payload)
        assert resp.status_code == 422

    def test_auto_logistics_and_manual_price_override(self, authenticated_client: TestClient):
        """Zero distance and zero trip rate auto-resolve logistics, and manual selling price overrides formula."""
        # 1. Create with zero distance and zero trip rate (auto-resolved from Bagodar PIN 825322)
        payload = {
            "enquiry_no": "MKTGENQ26-27594",
            "customer_name": "Konar, Jharkhand",
            "project_name": "Unipro Techno Infrastructure",
            "delivery_location": "Bagodar, Jharkhand - 825322",
            "distance_km": 0.0,
            "trip_rate": 0.0,
            "calculation_mode": "REFERENCE",
            "lines": [
                {
                    "line_number": 1,
                    "size_mm": 1226.0,
                    "thickness_mm": 8.4,
                    "basis": "ID",
                    "quantity": 2000.0,
                    "unit": "Metres",
                    "pipe_length_m": 12.0,
                    "pipes_per_trailer": 4.0,
                    "has_external": True,
                    "has_internal": True,
                    "steel_rate": 62000.0,
                    "conversion": 4000.0,
                    "manual_selling_price": 0.0,
                    "override_reason": "",
                }
            ],
        }
        res = authenticated_client.post("/api/quotations", json=payload)
        assert res.status_code == 201
        quote = res.json()
        assert quote["distance_km"] == 782.0
        assert quote["trip_rate"] >= 93000.0
        quote_id = quote["id"]

        # 2. Calculate without 500 error
        calc_res = authenticated_client.post(f"/api/quotations/{quote_id}/calculate")
        assert calc_res.status_code == 200
        snap = calc_res.json()
        assert snap["grand_total"] > 0

        # 3. Apply manual price override
        update_payload = {
            "lines": [
                {
                    "line_number": 1,
                    "size_mm": 1226.0,
                    "thickness_mm": 8.4,
                    "basis": "ID",
                    "quantity": 2000.0,
                    "unit": "Metres",
                    "pipe_length_m": 12.0,
                    "pipes_per_trailer": 4.0,
                    "has_external": True,
                    "has_internal": True,
                    "steel_rate": 62000.0,
                    "conversion": 4000.0,
                    "manual_selling_price": 22353.0,
                    "override_reason": "Target tender selling price",
                }
            ]
        }
        update_res = authenticated_client.put(f"/api/quotations/{quote_id}", json=update_payload)
        assert update_res.status_code == 200

        # 4. Recalculate with manual override
        recalc_res = authenticated_client.post(f"/api/quotations/{quote_id}/calculate")
        assert recalc_res.status_code == 200
        recalc_snap = recalc_res.json()
        line_result = recalc_snap["results_snapshot"]["lines"][0]
        assert line_result["price_source"] == "MANUAL_OVERRIDE"
        assert line_result["applied_unit_price"] == 22353.0

