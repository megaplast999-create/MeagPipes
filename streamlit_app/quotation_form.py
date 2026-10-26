"""Quotation creation and editing form for Streamlit app."""

from __future__ import annotations

import streamlit as st
from typing import Dict, Any, List, Optional, Callable
from .api_client import (
    api_get_quotation,
    api_create_quotation,
    api_update_quotation,
    api_calculate_quotation,
    APIError,
)
from .config import DEFAULT_LINE, DEFAULT_CALCULATION_MODE, DEFAULT_TRIP_RATE


def render_quotation_editor(
    quotation_id: Optional[int],
    on_back: Callable[[], None],
    on_view_results: Callable[[int, Dict[str, Any]], None],
) -> None:
    """Renders quotation header and line items editor."""
    is_edit = quotation_id is not None
    title = f"Edit Quotation (ID: {quotation_id})" if is_edit else "Create New Commercial Quotation"
    st.subheader(title)

    # Initialize form state in st.session_state if not already present
    state_key = f"form_quote_{quotation_id or 'new'}"
    if state_key not in st.session_state:
        if is_edit:
            try:
                with st.spinner("Fetching quotation..."):
                    q_data = api_get_quotation(quotation_id)
                st.session_state[state_key] = {
                    "enquiry_no": q_data.get("enquiry_no", ""),
                    "customer_name": q_data.get("customer_name", ""),
                    "project_name": q_data.get("project_name") or "",
                    "delivery_location": q_data.get("delivery_location") or "",
                    "distance_km": float(q_data.get("distance_km", 0.0)),
                    "trip_rate": float(q_data.get("trip_rate", DEFAULT_TRIP_RATE)),
                    "calculation_mode": q_data.get("calculation_mode", DEFAULT_CALCULATION_MODE),
                    "lines": [dict(DEFAULT_LINE, **l) for l in q_data.get("lines", [])] or [dict(DEFAULT_LINE)],
                }
            except APIError as err:
                st.error(f"Failed to fetch quotation: {err.message}")
                if st.button("Back to Dashboard"):
                    on_back()
                return
        else:
            # Generate suggested enquiry code for new quotation
            import time
            code = f"ENQ-26-27-{int(time.time()) % 10000:04d}"
            st.session_state[state_key] = {
                "enquiry_no": code,
                "customer_name": "",
                "project_name": "",
                "delivery_location": "",
                "distance_km": 0.0,
                "trip_rate": DEFAULT_TRIP_RATE,
                "calculation_mode": DEFAULT_CALCULATION_MODE,
                "lines": [dict(DEFAULT_LINE)],
            }

    data = st.session_state[state_key]

    # Top action bar
    col_nav, col_space, col_save, col_calc = st.columns([2, 3, 2, 2.5])
    with col_nav:
        if st.button("← Back to Dashboard", use_container_width=True):
            on_back()
            return

    # Quotation Header Section
    with st.container(border=True):
        st.markdown("#### 1. Header & Commercial Logistics")
        h_col1, h_col2, h_col3 = st.columns(3)

        with h_col1:
            data["enquiry_no"] = st.text_input("Enquiry Number *", value=data["enquiry_no"])
            data["customer_name"] = st.text_input("Customer Name *", value=data["customer_name"])

        with h_col2:
            data["project_name"] = st.text_input("Project Name", value=data["project_name"])
            data["delivery_location"] = st.text_input("Delivery Location", value=data["delivery_location"])

        with h_col3:
            data["distance_km"] = st.number_input(
                "Distance (km)", min_value=0.0, value=float(data["distance_km"]), step=10.0
            )
            data["trip_rate"] = st.number_input(
                "Trip Rate (₹ / Trailer)", min_value=0.0, value=float(data["trip_rate"]), step=1000.0
            )

        mode_options = ["REFERENCE", "NORMAL", "LEGACY"]
        current_mode_idx = mode_options.index(data["calculation_mode"]) if data["calculation_mode"] in mode_options else 0
        data["calculation_mode"] = st.selectbox(
            "Calculation Mode Policy",
            options=mode_options,
            index=current_mode_idx,
            help="REFERENCE mode reproduces exact commercial spreadsheet parity.",
        )

    # Quotation Lines Section
    lines: List[Dict[str, Any]] = data["lines"]

    st.markdown(f"#### 2. Quotation Line Items ({len(lines)})")

    line_to_delete = None
    line_to_duplicate = None

    for idx, line in enumerate(lines):
        line_num = idx + 1
        line["line_number"] = line_num

        with st.container(border=True):
            header_col, action_col = st.columns([5, 2])
            with header_col:
                st.markdown(
                    f"**Line #{line_num}:** `{line.get('size_mm')} mm` ({line.get('basis')}) &times; `{line.get('thickness_mm')} mm WT` &bull; `{line.get('quantity')} {line.get('unit')}`"
                )
            with action_col:
                btn_c1, btn_c2 = st.columns(2)
                with btn_c1:
                    if st.button("Duplicate", key=f"dup_{idx}", use_container_width=True):
                        line_to_duplicate = idx
                with btn_c2:
                    if len(lines) > 1 and st.button("Delete", key=f"del_{idx}", use_container_width=True):
                        line_to_delete = idx

            # Primary specification inputs
            l_col1, l_col2, l_col3, l_col4 = st.columns(4)
            with l_col1:
                line["size_mm"] = st.number_input(
                    "Size / OD (mm)",
                    min_value=1.0,
                    max_value=5000.0,
                    value=float(line.get("size_mm", 508.0)),
                    step=1.0,
                    key=f"size_{idx}",
                )
                line["basis"] = st.selectbox(
                    "Sizing Basis",
                    options=["ID", "OD"],
                    index=0 if line.get("basis", "ID").upper() == "ID" else 1,
                    key=f"basis_{idx}",
                )

            with l_col2:
                line["thickness_mm"] = st.number_input(
                    "Thickness (mm)",
                    min_value=0.1,
                    max_value=100.0,
                    value=float(line.get("thickness_mm", 6.4)),
                    step=0.1,
                    key=f"thk_{idx}",
                )
                unit_opts = ["Metres", "MT", "Nos"]
                curr_unit = line.get("unit", "Metres")
                curr_unit_idx = unit_opts.index(curr_unit) if curr_unit in unit_opts else 0
                line["unit"] = st.selectbox(
                    "Quantity Unit",
                    options=unit_opts,
                    index=curr_unit_idx,
                    key=f"unit_{idx}",
                )

            with l_col3:
                line["quantity"] = st.number_input(
                    "Quantity",
                    min_value=1.0,
                    value=float(line.get("quantity", 1000.0)),
                    step=10.0,
                    key=f"qty_{idx}",
                )
                line["pipe_length_m"] = st.number_input(
                    "Pipe Length (m)",
                    min_value=6.0,
                    max_value=13.0,
                    value=float(line.get("pipe_length_m", 12.0)),
                    step=0.5,
                    key=f"len_{idx}",
                )

            with l_col4:
                line["pipes_per_trailer"] = st.number_input(
                    "Pipes / Trailer",
                    min_value=1.0,
                    max_value=20.0,
                    value=float(line.get("pipes_per_trailer", 4.0)),
                    step=1.0,
                    key=f"trailer_{idx}",
                )
                line["has_external"] = st.checkbox(
                    "External 3LPE Coating",
                    value=bool(line.get("has_external", True)),
                    key=f"ext_chk_{idx}",
                )
                line["has_internal"] = st.checkbox(
                    "Internal Liquid Epoxy",
                    value=bool(line.get("has_internal", True)),
                    key=f"int_chk_{idx}",
                )

            # Collapsible Advanced Commercial Rates & Overrides
            with st.expander(f"Commercial Rates & Manual Overrides (Line #{line_num})", expanded=False):
                cr_col1, cr_col2, cr_col3 = st.columns(3)
                with cr_col1:
                    line["steel_rate"] = st.number_input(
                        "Steel Rate (₹/MT)", min_value=0.0, value=float(line.get("steel_rate", 61000.0)), key=f"steel_{idx}"
                    )
                    line["grade_extra"] = st.number_input(
                        "Grade Extra (₹/MT)", min_value=0.0, value=float(line.get("grade_extra", 0.0)), key=f"grade_{idx}"
                    )
                    line["yield_pct"] = st.number_input(
                        "Yield Burning (%)", min_value=0.0, max_value=20.0, value=float(line.get("yield_pct", 2.0)), key=f"yield_{idx}"
                    )
                    line["conversion"] = st.number_input(
                        "Conversion Cost (₹/MT)", min_value=0.0, value=float(line.get("conversion", 5000.0)), key=f"conv_{idx}"
                    )

                with cr_col2:
                    line["ext_rate"] = st.number_input(
                        "Ext Coating Rate (₹/m²)", min_value=0.0, value=float(line.get("ext_rate", 500.0)), key=f"extrate_{idx}"
                    )
                    line["int_rate"] = st.number_input(
                        "Int Coating Rate (₹/m²)", min_value=0.0, value=float(line.get("int_rate", 275.0)), key=f"intrate_{idx}"
                    )
                    line["handling"] = st.number_input(
                        "Handling (₹/MT)", min_value=0.0, value=float(line.get("handling", 100.0)), key=f"hand_{idx}"
                    )
                    line["financial_pct"] = st.number_input(
                        "Financial Cost (%)", min_value=0.0, max_value=20.0, value=float(line.get("financial_pct", 1.5)), key=f"fin_{idx}"
                    )

                with cr_col3:
                    line["manual_fc"] = st.number_input(
                        "Manual FC Override (₹)", min_value=0.0, value=float(line.get("manual_fc", 0.0)), key=f"mfc_{idx}"
                    )
                    line["manual_selling_price"] = st.number_input(
                        "Manual Price Override (₹)", min_value=0.0, value=float(line.get("manual_selling_price", 0.0)), key=f"msp_{idx}"
                    )
                    line["override_reason"] = st.text_input(
                        "Override Reason", value=str(line.get("override_reason", "")), key=f"reason_{idx}"
                    )

    # Handle line modifications
    if line_to_duplicate is not None:
        new_item = dict(lines[line_to_duplicate])
        lines.append(new_item)
        st.rerun()

    if line_to_delete is not None:
        lines.pop(line_to_delete)
        st.rerun()

    if st.button("+ Add Pipe Line", use_container_width=False):
        lines.append(dict(DEFAULT_LINE))
        st.rerun()

    st.write("")
    st.divider()

    # Form Submission Actions
    def build_payload() -> Optional[Dict[str, Any]]:
        enq = data["enquiry_no"].strip()
        cust = data["customer_name"].strip()
        if not enq or not cust:
            st.error("Enquiry Number and Customer Name are required.")
            return None

        # Build clean payload according to backend QuotationCreateRequest
        return {
            "enquiry_no": enq,
            "customer_name": cust,
            "project_name": data["project_name"].strip() or None,
            "delivery_location": data["delivery_location"].strip() or None,
            "distance_km": float(data["distance_km"]),
            "trip_rate": float(data["trip_rate"]),
            "calculation_mode": data["calculation_mode"],
            "lines": [
                {
                    "line_number": idx + 1,
                    "size_mm": float(l["size_mm"]),
                    "thickness_mm": float(l["thickness_mm"]),
                    "basis": str(l["basis"]).upper(),
                    "quantity": float(l["quantity"]),
                    "unit": str(l["unit"]),
                    "pipe_length_m": float(l.get("pipe_length_m", 12.0)),
                    "pipes_per_trailer": float(l.get("pipes_per_trailer", 4.0)),
                    "has_external": bool(l.get("has_external", True)),
                    "has_internal": bool(l.get("has_internal", True)),
                    "steel_rate": float(l.get("steel_rate", 61000.0)),
                    "grade_extra": float(l.get("grade_extra", 0.0)),
                    "yield_pct": float(l.get("yield_pct", 2.0)),
                    "conversion": float(l.get("conversion", 5000.0)),
                    "testing": float(l.get("testing", 0.0)),
                    "margin": float(l.get("margin", 0.0)),
                    "ext_rate": float(l.get("ext_rate", 500.0)),
                    "int_rate": float(l.get("int_rate", 275.0)),
                    "handling": float(l.get("handling", 100.0)),
                    "financial_pct": float(l.get("financial_pct", 1.5)),
                    "gst_pct": float(l.get("gst_pct", 18.0)),
                    "manual_fc": float(l.get("manual_fc", 0.0)),
                    "manual_selling_price": float(l.get("manual_selling_price", 0.0)),
                    "override_reason": str(l.get("override_reason", "")),
                }
                for idx, l in enumerate(lines)
            ],
        }

    b_col1, b_col2 = st.columns([1, 1])
    with b_col1:
        if st.button("Save Draft", type="secondary", use_container_width=True):
            payload = build_payload()
            if payload:
                try:
                    with st.spinner("Saving quotation..."):
                        if is_edit:
                            res = api_update_quotation(quotation_id, payload)
                        else:
                            res = api_create_quotation(payload)
                        st.success(f"Quotation {res['enquiry_no']} saved successfully!")
                        # Clear form state and navigate back
                        del st.session_state[state_key]
                        on_back()
                except APIError as e:
                    st.error(f"Failed to save quotation: {e.message}")

    with b_col2:
        if st.button("Save & Calculate Quotation", type="primary", use_container_width=True):
            payload = build_payload()
            if payload:
                try:
                    with st.spinner("Saving and calculating commercial rates..."):
                        if is_edit:
                            saved = api_update_quotation(quotation_id, payload)
                        else:
                            saved = api_create_quotation(payload)

                        target_id = saved["id"]
                        snapshot = api_calculate_quotation(target_id)
                        st.success(f"Calculation complete for {saved['enquiry_no']}!")
                        del st.session_state[state_key]
                        on_view_results(target_id, snapshot)
                except APIError as e:
                    st.error(f"Operation failed: {e.message}")
