"""Authoritative calculation results view and commercial analysis for Streamlit app."""

from __future__ import annotations

import streamlit as st
import pandas as pd
from typing import Dict, Any, Callable
from .config import format_inr, format_number
from .api_client import api_export_excel, APIError


def render_calculation_results(
    quotation_id: int,
    snapshot: Dict[str, Any],
    on_back: Callable[[], None],
    on_edit: Callable[[int], None],
) -> None:
    """Displays authoritative server-side commercial calculations and Excel export."""
    st.subheader(f"Commercial Calculation Results (Snapshot v{snapshot.get('version', 1)})")

    # Action navigation bar
    col_back, col_edit, col_export = st.columns([2, 2, 3])
    with col_back:
        if st.button("← Dashboard", use_container_width=True):
            on_back()
            return
    with col_edit:
        if st.button("Edit Quotation", use_container_width=True):
            on_edit(quotation_id)
            return
    with col_export:
        try:
            xlsx_bytes, filename = api_export_excel(quotation_id)
            st.download_button(
                label="📥 Download Commercial Excel (.xlsx)",
                data=xlsx_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        except APIError as e:
            st.error(f"Excel export failed: {e.message}")

    # Main KPI Metrics
    grand_total = snapshot.get("grand_total", 0.0)
    subtotal = snapshot.get("subtotal_excl_gst", 0.0)
    gst = snapshot.get("gst_amount", 0.0)
    metres = snapshot.get("total_metres", 0.0)
    mt = snapshot.get("total_mt", 0.0)
    trips = snapshot.get("total_trips", 0.0)
    mode = snapshot.get("calculation_mode", "REFERENCE")

    st.markdown(
        f"<span style='background-color:#1e3a8a;color:#93c5fd;padding:4px 10px;border-radius:6px;font-size:0.85rem;font-weight:600;'>MODE: {mode}</span>",
        unsafe_allow_html=True,
    )

    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    with kpi1:
        st.metric("Grand Total (INR)", format_inr(grand_total, 2))
    with kpi2:
        st.metric("Basic Subtotal", format_inr(subtotal, 2))
    with kpi3:
        st.metric("GST Amount (18%)", format_inr(gst, 2))
    with kpi4:
        st.metric("Total Mass", f"{format_number(mt, 3)} MT")
    with kpi5:
        st.metric("Logistics Trips", f"{format_number(trips, 2)} Trips")

    st.divider()

    # Detailed line-by-line results table
    results_snapshot = snapshot.get("results_snapshot", {})
    lines = results_snapshot.get("lines", [])

    st.markdown(f"#### Line-by-Line Commercial Breakdown ({len(lines)} lines)")

    if lines:
        table_rows = []
        for idx, line in enumerate(lines):
            table_rows.append({
                "Line #": idx + 1,
                "OD (mm)": round(line.get("od_mm", 0.0), 1),
                "ID (mm)": round(line.get("inner_d_mm", 0.0), 1),
                "Mass (kg/m)": round(line.get("mass_kg_per_m", 0.0), 2),
                "Length (m)": round(line.get("metres", 0.0), 1),
                "Weight (MT)": round(line.get("weight_mt", 0.0), 3),
                "Trips": round(line.get("freight_trips", 0.0), 2),
                "Bare / m": f"₹{format_number(line.get('bare_per_m'), 2)}",
                "Ext / m": f"₹{format_number(line.get('external_per_m'), 2)}",
                "Int / m": f"₹{format_number(line.get('internal_per_m'), 2)}",
                "Freight / m": f"₹{format_number(line.get('freight_per_m'), 2)}",
                "Total Rate / m": f"₹{format_number(line.get('total_per_m'), 2)}",
                "Total Rate / MT": f"₹{format_number(line.get('total_per_mt'), 2)}",
                "Line Value (INR)": format_inr(line.get("line_value"), 2),
            })

        df = pd.DataFrame(table_rows)
        st.dataframe(df, use_container_width=True, hide_index=True)

    # Cost Sheet Weighted Breakdown
    weighted = results_snapshot.get("weighted")
    if weighted:
        st.markdown("#### Weighted Cost Sheet Summary (₹ / MT)")
        w_cols = st.columns(6)
        with w_cols[0]:
            st.metric("Bare Pipe / MT", f"₹{format_number(weighted.get('bare_pipe_cost_per_mt'), 2)}")
        with w_cols[1]:
            st.metric("Ext Coating / MT", f"₹{format_number(weighted.get('external_cost_per_mt'), 2)}")
        with w_cols[2]:
            st.metric("Int Coating / MT", f"₹{format_number(weighted.get('internal_cost_per_mt'), 2)}")
        with w_cols[3]:
            st.metric("Freight / MT", f"₹{format_number(weighted.get('freight_per_mt'), 2)}")
        with w_cols[4]:
            st.metric("Financial Cost / MT", f"₹{format_number(weighted.get('financial_cost_per_mt'), 2)}")
        with w_cols[5]:
            st.metric("Blended Rate / MT", f"₹{format_number(weighted.get('total_per_mt'), 2)}")
