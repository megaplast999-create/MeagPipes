"""Quotation Dashboard & History view for Streamlit app."""

from __future__ import annotations

import streamlit as st
from typing import List, Dict, Any, Callable
try:
    from .api_client import (
        api_list_quotations,
        api_calculate_quotation,
        api_get_snapshot,
        api_export_excel,
        APIError,
    )
    from .config import format_number
except (ImportError, ValueError):
    from api_client import (
        api_list_quotations,
        api_calculate_quotation,
        api_get_snapshot,
        api_export_excel,
        APIError,
    )
    from config import format_number



def render_dashboard(
    on_new_quote: Callable[[], None],
    on_edit_quote: Callable[[int], None],
    on_view_results: Callable[[int, Dict[str, Any]], None],
) -> None:
    """Renders the quotation management dashboard with list, search, and quick actions."""
    st.subheader("Commercial Quotations Dashboard")

    col_actions, col_refresh, col_new = st.columns([4, 1, 1.5])
    with col_new:
        if st.button("+ New Quotation", type="primary", use_container_width=True):
            on_new_quote()
            return
    with col_refresh:
        refresh = st.button("Refresh", use_container_width=True)

    # Fetch quotations from backend
    try:
        with st.spinner("Loading quotations from server..."):
            quotations: List[Dict[str, Any]] = api_list_quotations()
    except APIError as err:
        st.error(f"Failed to retrieve quotations: {err.message}")
        return
    except Exception as err:
        st.error(f"Unexpected error loading quotations: {err}")
        return

    # Metrics Summary
    total = len(quotations)
    calculated = sum(1 for q in quotations if q.get("status") == "CALCULATED")
    drafts = sum(1 for q in quotations if q.get("status") == "DRAFT")

    m_col1, m_col2, m_col3 = st.columns(3)
    with m_col1:
        st.metric("Total Quotations", total)
    with m_col2:
        st.metric("Calculated & Verified", calculated)
    with m_col3:
        st.metric("Draft Status", drafts)

    st.write("")

    if not quotations:
        st.info("No quotations found. Click '+ New Quotation' above to create your first quotation.")
        return

    # Search filter
    search_query = st.text_input("Search quotations", placeholder="Filter by Enquiry #, Customer, Project, or Location...")
    filtered_quotes = quotations
    if search_query.strip():
        q_lower = search_query.lower()
        filtered_quotes = [
            q for q in quotations
            if q_lower in str(q.get("enquiry_no", "")).lower()
            or q_lower in str(q.get("customer_name", "")).lower()
            or q_lower in str(q.get("project_name", "")).lower()
            or q_lower in str(q.get("delivery_location", "")).lower()
        ]

    st.caption(f"Showing {len(filtered_quotes)} of {total} quotations")

    # Render each quotation card / row
    for q in filtered_quotes:
        q_id = q["id"]
        enquiry = q.get("enquiry_no", "N/A")
        customer = q.get("customer_name", "Unknown")
        project = q.get("project_name") or "Standard Project"
        location = q.get("delivery_location") or "Ex-Works"
        mode = q.get("calculation_mode", "REFERENCE")
        status = q.get("status", "DRAFT")
        version = q.get("version", 1)
        line_count = len(q.get("lines", []))
        dist_km = q.get("distance_km", 0)
        trip_rate = q.get("trip_rate", 98000)

        with st.container(border=True):
            r_col1, r_col2, r_col3, r_col4 = st.columns([3, 2, 2, 3])

            with r_col1:
                st.markdown(f"**{enquiry}** &bull; `{customer}`")
                st.caption(f"Project: {project} | Location: {location}")

            with r_col2:
                st.write(f"Mode: `{mode}`")
                st.caption(f"{format_number(dist_km, 0)} km @ ₹{format_number(trip_rate, 0)}/trip")

            with r_col3:
                status_color = "green" if status == "CALCULATED" else "orange"
                st.markdown(f":{status_color}[**{status}**] (v{version})")
                st.caption(f"{line_count} line{'s' if line_count != 1 else ''}")

            with r_col4:
                btn_cols = st.columns(3)
                with btn_cols[0]:
                    if st.button("Edit", key=f"edit_{q_id}", use_container_width=True):
                        on_edit_quote(q_id)
                        return

                with btn_cols[1]:
                    if status == "CALCULATED":
                        if st.button("Results", key=f"res_{q_id}", use_container_width=True):
                            try:
                                snapshot = api_get_snapshot(q_id)
                                on_view_results(q_id, snapshot)
                                return
                            except APIError as e:
                                st.error(f"Could not load snapshot: {e.message}")
                    else:
                        if st.button("Calc", key=f"calc_{q_id}", type="primary", use_container_width=True):
                            try:
                                with st.spinner("Calculating..."):
                                    snapshot = api_calculate_quotation(q_id)
                                    st.success(f"Calculated {enquiry}!")
                                    on_view_results(q_id, snapshot)
                                    return
                            except APIError as e:
                                st.error(f"Calculation failed: {e.message}")

                with btn_cols[2]:
                    # Export button triggers download
                    try:
                        xlsx_bytes, filename = api_export_excel(q_id)
                        st.download_button(
                            label="Excel",
                            data=xlsx_bytes,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_{q_id}",
                            use_container_width=True,
                        )
                    except Exception:
                        st.button("Excel", key=f"dl_err_{q_id}", disabled=True, use_container_width=True)
