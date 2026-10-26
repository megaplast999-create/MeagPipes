"""Authentication handling and login UI for Streamlit app."""

from __future__ import annotations

import streamlit as st
from typing import Optional, Dict, Any
try:
    from .api_client import api_login, api_get_me, api_logout, APIError
except (ImportError, ValueError):
    from api_client import api_login, api_get_me, api_logout, APIError



def check_authentication() -> Optional[Dict[str, Any]]:
    """Checks whether the current Streamlit session has an active, valid authentication state."""
    if "user" in st.session_state and st.session_state["user"]:
        return st.session_state["user"]

    # Verify session with backend
    user = api_get_me()
    if user:
        st.session_state["user"] = user
        return user

    return None


def render_login_view() -> None:
    """Renders the professional Mega Pipes login interface."""
    st.markdown(
        """
        <div style="text-align: center; margin-bottom: 2rem;">
            <h1 style="color: #3b82f6; margin-bottom: 0.25rem;">MEGA PIPES</h1>
            <p style="color: #94a3b8; font-size: 1.1rem;">Commercial Quotation & Cost-Sheet Management</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.container(border=True):
            st.subheader("Authorized Staff Sign-In")
            st.caption("Secure session authentication via Render FastAPI Backend")

            email = st.text_input(
                "Work Email",
                placeholder="pilot@megapipes.internal",
                key="login_email_input",
            )
            password = st.text_input(
                "Password",
                type="password",
                placeholder="••••••••••••",
                key="login_password_input",
            )

            def _fill_demo() -> None:
                st.session_state["login_email_input"] = "pilot@megapipes.internal"
                st.session_state["login_password_input"] = "PilotMegaPipes2026!"

            col_btn1, col_btn2 = st.columns([2, 1])
            with col_btn1:
                submit = st.button("Sign In to Mega Pipes", type="primary", use_container_width=True)
            with col_btn2:
                st.button("Fill Demo", on_click=_fill_demo, use_container_width=True)


            if submit:
                clean_email = email.strip()
                if not clean_email or not password:
                    st.error("Please enter both email and password.")
                    return

                with st.spinner("Authenticating with backend..."):
                    try:
                        user = api_login(clean_email, password)
                        st.session_state["user"] = user
                        st.success(f"Welcome back, {user.get('full_name') or clean_email}!")
                        st.rerun()
                    except APIError as err:
                        st.error(f"Authentication failed: {err.message}")
                    except Exception as err:
                        st.error("Unexpected error during authentication. Please check network connection.")


def render_user_header(user: Dict[str, Any]) -> None:
    """Renders top navigation and user session bar with logout option."""
    col_brand, col_status, col_logout = st.columns([3, 2, 1])

    with col_brand:
        st.markdown(
            f"**MEGA PIPES** &bull; *Commercial Quotation Portal* | User: `{user.get('email')}`",
        )

    with col_status:
        st.caption("Backend: `meagpipes-1.onrender.com` (PostgreSQL)")

    with col_logout:
        if st.button("Logout", type="secondary", use_container_width=True):
            api_logout()
            st.session_state.clear()
            st.rerun()

    st.divider()
