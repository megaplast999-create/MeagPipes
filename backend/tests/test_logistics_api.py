"""Tests for logistics and pincode calculation API endpoints."""

import pytest
from fastapi.testclient import TestClient

from app.main import app

client = TestClient(app)


def test_pincode_lookup_valid():
    """Test lookup of Bagodar reference pincode 825322."""
    res = client.get("/api/logistics/pincode/825322")
    assert res.status_code == 200
    data = res.json()
    assert data["valid"] is True
    assert data["pin"] == "825322"
    assert data["distance_km"] == 782.0
    assert data["trip_rate"] >= 93000.0
    assert "Bagodar" in data["office_name"] or "GIRIDIH" in data["district"]


def test_pincode_lookup_extracted_from_address():
    """Test extracting 6-digit pin from address string."""
    res = client.get("/api/logistics/pincode/Pune, Maharashtra - 411001")
    assert res.status_code == 200
    data = res.json()
    assert data["valid"] is True
    assert data["pin"] == "411001"
    assert data["distance_km"] > 0


def test_pincode_lookup_invalid():
    """Test rejection of non-pincode text."""
    res = client.get("/api/logistics/pincode/INVALID")
    assert res.status_code == 400


def test_resolve_logistics_endpoint():
    """Test POST /api/logistics/resolve."""
    res = client.post("/api/logistics/resolve", json={
        "delivery_location": "Bagodar, Jharkhand 825322",
        "calculation_mode": "REFERENCE",
    })
    assert res.status_code == 200
    data = res.json()
    assert data["valid"] is True
    assert data["distance_km"] == 782.0
