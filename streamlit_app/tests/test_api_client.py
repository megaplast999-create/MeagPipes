"""Unit tests for Streamlit API client, session isolation, and CSRF token propagation."""

from __future__ import annotations

from unittest.mock import MagicMock, patch
import pytest
import requests
import streamlit as st

from streamlit_app.api_client import (
    APIError,
    api_request,
    get_user_session,
    reset_user_session,
    _get_csrf_token,
)
from streamlit_app.config import CSRF_COOKIE_NAME


@pytest.fixture(autouse=True)
def clean_session_state() -> None:
    """Ensures each test has an isolated session state."""
    if "api_session" in st.session_state:
        del st.session_state["api_session"]
    yield
    if "api_session" in st.session_state:
        del st.session_state["api_session"]


def test_session_isolation_in_session_state() -> None:
    """Verifies requests.Session is stored strictly per-user in st.session_state."""
    s1 = get_user_session()
    assert isinstance(s1, requests.Session)
    assert st.session_state["api_session"] is s1

    # Second call returns the same session instance
    s2 = get_user_session()
    assert s1 is s2

    # Reset clears session instance
    reset_user_session()
    assert "api_session" not in st.session_state

    # Next call creates fresh session
    s3 = get_user_session()
    assert s3 is not s1


def test_csrf_token_extraction() -> None:
    session = requests.Session()
    # No cookies initially
    assert _get_csrf_token(session) is None

    # Attach CSRF cookie
    session.cookies.set(CSRF_COOKIE_NAME, "test-csrf-token-12345")
    assert _get_csrf_token(session) == "test-csrf-token-12345"


def test_api_request_injects_csrf_header_on_post() -> None:
    session = get_user_session()
    session.cookies.set(CSRF_COOKIE_NAME, "mocked_csrf_secret")

    mock_response = MagicMock()
    mock_response.ok = True
    mock_response.status_code = 200
    mock_response.json.return_value = {"status": "ok"}

    with patch.object(session, "request", return_value=mock_response) as mock_req:
        res = api_request("/api/test", method="POST", json_data={"hello": "world"})
        assert res == {"status": "ok"}

        # Verify call arguments
        mock_req.assert_called_once()
        _, kwargs = mock_req.call_args
        assert kwargs["method"] == "POST"
        assert kwargs["headers"].get("X-CSRF-Token") == "mocked_csrf_secret"
        assert kwargs["json"] == {"hello": "world"}


def test_api_request_handles_422_validation_error() -> None:
    session = get_user_session()

    mock_response = MagicMock()
    mock_response.ok = False
    mock_response.status_code = 422
    mock_response.json.return_value = {
        "detail": [
            {"loc": ["body", "customer_name"], "msg": "Field required", "type": "missing"}
        ]
    }

    with patch.object(session, "request", return_value=mock_response):
        with pytest.raises(APIError) as exc_info:
            api_request("/api/quotations", method="POST", json_data={})

        assert exc_info.value.status_code == 422
        assert "Field required" in exc_info.value.message


def test_api_request_handles_stream_excel_download() -> None:
    session = get_user_session()

    mock_response = MagicMock()
    mock_response.ok = True
    mock_response.status_code = 200
    mock_response.content = b"PK\x03\x04mocked_excel_content"
    mock_response.headers = {
        "Content-Disposition": 'attachment; filename="ENQ-26-27-001_R0.xlsx"'
    }

    with patch.object(session, "request", return_value=mock_response):
        content, filename = api_request("/api/quotations/1/export", method="GET", stream=True)
        assert content == b"PK\x03\x04mocked_excel_content"
        assert filename == "ENQ-26-27-001_R0.xlsx"
