"""Mega Pipes Web Pilot - Streamlit Commercial Quotation Application.

Production Streamlit web client communicating with FastAPI PostgreSQL backend.
"""

from __future__ import annotations

import streamlit as st
from .auth import check_authentication, render_login_view, render_user_header
from .quotation_history import render_dashboard
from .quotation_form import render_quotation_editor
from .calculation_results import render_calculation_results


def main() -> None:
    # Set page configuration
    st.set_page_config(
        page_title="Mega Pipes Quotation System",
        page_icon="🏭",
        layout="wide",
        initial_sidebar_state="collapsed",
    )

    # 1. Enforce Authentication
    user = check_authentication()
    if not user:
        render_login_view()
        return

    # 2. Render Header with Session Identity & Logout
    render_user_header(user)

    # 3. View Routing State
    if "view" not in st.session_state:
        st.session_state["view"] = "dashboard"
    if "active_quote_id" not in st.session_state:
        st.session_state["active_quote_id"] = None
    if "active_snapshot" not in st.session_state:
        st.session_state["active_snapshot"] = None

    current_view = st.session_state["view"]

    # Navigation callbacks
    def go_dashboard() -> None:
        st.session_state["view"] = "dashboard"
        st.session_state["active_quote_id"] = None
        st.session_state["active_snapshot"] = None
        st.rerun()

    def go_new_quote() -> None:
        st.session_state["view"] = "new_quote"
        st.session_state["active_quote_id"] = None
        st.session_state["active_snapshot"] = None
        st.rerun()

    def go_edit_quote(qid: int) -> None:
        st.session_state["view"] = "edit_quote"
        st.session_state["active_quote_id"] = qid
        st.session_state["active_snapshot"] = None
        st.rerun()

    def go_results(qid: int, snapshot: dict) -> None:
        st.session_state["view"] = "results"
        st.session_state["active_quote_id"] = qid
        st.session_state["active_snapshot"] = snapshot
        st.rerun()

    # View dispatcher
    if current_view == "dashboard":
        render_dashboard(
            on_new_quote=go_new_quote,
            on_edit_quote=go_edit_quote,
            on_view_results=go_results,
        )
    elif current_view == "new_quote":
        render_quotation_editor(
            quotation_id=None,
            on_back=go_dashboard,
            on_view_results=go_results,
        )
    elif current_view == "edit_quote":
        render_quotation_editor(
            quotation_id=st.session_state["active_quote_id"],
            on_back=go_dashboard,
            on_view_results=go_results,
        )
    elif current_view == "results":
        render_calculation_results(
            quotation_id=st.session_state["active_quote_id"],
            snapshot=st.session_state["active_snapshot"],
            on_back=go_dashboard,
            on_edit=go_edit_quote,
        )

    # Footer
    st.markdown(
        """
        <div style="margin-top: 4rem; text-align: center; color: #64748b; font-size: 0.85rem; border-top: 1px solid #1e293b; padding-top: 1.5rem;">
            Mega Pipes Quotation System &bull; Production Streamlit Client &bull; Render FastAPI Backend
        </div>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
