"""API client for Mega Pipes FastAPI backend with strict per-user session isolation and CSRF handling."""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Tuple
import requests
import streamlit as st

try:
    from .config import (
        BACKEND_URL,
        CSRF_COOKIE_NAME,
        SESSION_COOKIE_NAME,
    )
except (ImportError, ValueError):
    from config import (
        BACKEND_URL,
        CSRF_COOKIE_NAME,
        SESSION_COOKIE_NAME,
    )



class APIError(Exception):
    """Custom structured exception for API errors."""

    def __init__(self, message: str, status_code: int = 0, details: Any = None):
        super().__init__(message)
        self.message = message
        self.status_code = status_code
        self.details = details


def get_user_session() -> requests.Session:
    """Retrieves or initializes an isolated requests.Session per Streamlit browser session.
    
    CRITICAL SECURITY REQUIREMENT:
    Never use a module-level global session. Storing in st.session_state ensures
    complete session isolation between concurrent users.
    """
    if "api_session" not in st.session_state:
        session = requests.Session()
        session.headers.update({
            "Accept": "application/json",
            "User-Agent": "MegaPipes-Streamlit-Client/1.0",
        })
        st.session_state["api_session"] = session
    return st.session_state["api_session"]


def reset_user_session() -> None:
    """Closes and resets the current user's session."""
    if "api_session" in st.session_state:
        try:
            st.session_state["api_session"].close()
        except Exception:
            pass
        del st.session_state["api_session"]


def _get_csrf_token(session: requests.Session) -> Optional[str]:
    """Extracts CSRF token from current session cookies."""
    for cookie in session.cookies:
        if cookie.name == CSRF_COOKIE_NAME:
            return cookie.value
    return None


def api_request(
    endpoint: str,
    method: str = "GET",
    json_data: Optional[Dict[str, Any]] = None,
    params: Optional[Dict[str, Any]] = None,
    stream: bool = False,
    timeout: int = 30,
) -> Any:
    """Executes an authenticated HTTP request against the backend.
    
    Automatically attaches CSRF header for mutating requests (POST, PUT, PATCH, DELETE).
    """
    session = get_user_session()
    url = f"{BACKEND_URL}/{endpoint.lstrip('/')}"
    method = method.upper()

    headers: Dict[str, str] = {}

    # Attach CSRF token on mutating requests if present in cookies
    if method in ("POST", "PUT", "PATCH", "DELETE"):
        csrf = _get_csrf_token(session)
        if csrf:
            headers["X-CSRF-Token"] = csrf

    try:
        response = session.request(
            method=method,
            url=url,
            json=json_data,
            params=params,
            headers=headers,
            stream=stream,
            timeout=timeout,
        )
    except requests.RequestException as exc:
        raise APIError(f"Network error connecting to {url}: {exc}", status_code=0) from exc

    # Streamed response handling (e.g. binary Excel export)
    if stream:
        if not response.ok:
            try:
                err_json = response.json()
                msg = err_json.get("detail", f"Export failed with status {response.status_code}")
            except Exception:
                msg = f"Export failed with status {response.status_code}"
            raise APIError(msg, status_code=response.status_code)

        # Parse filename from Content-Disposition header
        disposition = response.headers.get("Content-Disposition", "")
        filename = "Mega_Pipes_Quotation.xlsx"
        match = re.search(r'filename\*?=(?:UTF-8\'\')?["\']?([^";\n]+)["\']?', disposition)
        if match:
            filename = match.group(1).strip()

        return response.content, filename

    # Standard JSON handling
    try:
        data = response.json()
    except Exception:
        data = response.text

    if not response.ok:
        msg = f"Request failed with status {response.status_code}"
        if isinstance(data, dict) and "detail" in data:
            detail = data["detail"]
            if isinstance(detail, list):
                msg = "; ".join([d.get("msg", str(d)) for d in detail if isinstance(d, dict)])
            else:
                msg = str(detail)
        elif isinstance(data, str) and data:
            msg = data
        raise APIError(msg, status_code=response.status_code, details=data)

    return data


# --- High-level API Service functions ---

def api_login(email: str, password: str) -> Dict[str, Any]:
    """Authenticates with backend and establishes session cookies."""
    return api_request(
        "/api/auth/login",
        method="POST",
        json_data={"email": email, "password": password},
    )


def api_get_me() -> Optional[Dict[str, Any]]:
    """Checks session validity with GET /api/auth/me."""
    try:
        res = api_request("/api/auth/me", method="GET")
        if res.get("authenticated") and res.get("user"):
            return res["user"]
        return None
    except APIError:
        return None


def api_logout() -> None:
    """Terminates session on backend and clears local session state."""
    try:
        api_request("/api/auth/logout", method="POST")
    except Exception:
        pass
    finally:
        reset_user_session()


def api_list_quotations() -> List[Dict[str, Any]]:
    """Retrieves all quotations belonging to authorized user."""
    return api_request("/api/quotations", method="GET")


def api_get_quotation(quotation_id: int) -> Dict[str, Any]:
    """Retrieves a single quotation by ID."""
    return api_request(f"/api/quotations/{quotation_id}", method="GET")


def api_create_quotation(payload: Dict[str, Any]) -> Dict[str, Any]:
    """Creates a new quotation draft."""
    return api_request("/api/quotations", method="POST", json_data=payload)


def api_update_quotation(quotation_id: int, payload: Dict[str, Any]) -> Dict[str, Any]:
    """Updates an existing quotation draft."""
    return api_request(f"/api/quotations/{quotation_id}", method="PUT", json_data=payload)


def api_calculate_quotation(quotation_id: int) -> Dict[str, Any]:
    """Triggers server-side calculation and creates snapshot."""
    return api_request(f"/api/quotations/{quotation_id}/calculate", method="POST")


def api_get_snapshot(quotation_id: int) -> Dict[str, Any]:
    """Retrieves latest calculation snapshot for quotation."""
    return api_request(f"/api/quotations/{quotation_id}/snapshot", method="GET")


def api_export_excel(quotation_id: int) -> Tuple[bytes, str]:
    """Downloads authoritative Excel quotation workbook."""
    return api_request(f"/api/quotations/{quotation_id}/export", method="GET", stream=True)
