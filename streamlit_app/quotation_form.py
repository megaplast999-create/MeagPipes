"""Quotation creation and editing form for Streamlit app with Pincode Intelligence and Line-Level Pricing."""

from __future__ import annotations

import re
import streamlit as st
from typing import Dict, Any, List, Optional, Callable

try:
    from .api_client import (
        api_get_quotation,
        api_create_quotation,
        api_update_quotation,
        api_calculate_quotation,
        api_resolve_logistics,
        APIError,
    )
    from .config import (
        DEFAULT_LINE,
        DEFAULT_CALCULATION_MODE,
        DEFAULT_TRIP_RATE,
        DEFAULT_STEEL_RATE,
        DEFAULT_CONVERSION,
        DEFAULT_DELIVERY_PIN,
        format_inr,
    )
except (ImportError, ValueError):
    from api_client import (
        api_get_quotation,
        api_create_quotation,
        api_update_quotation,
        api_calculate_quotation,
        api_resolve_logistics,
        APIError,
    )
    from config import (
        DEFAULT_LINE,
        DEFAULT_CALCULATION_MODE,
        DEFAULT_TRIP_RATE,
        DEFAULT_STEEL_RATE,
        DEFAULT_CONVERSION,
        DEFAULT_DELIVERY_PIN,
        format_inr,
    )


def _extract_pin(text: Optional[str]) -> Optional[str]:
    """Helper to extract first 6-digit Indian PIN code from address text."""
    if not text:
        return None
    match = re.search(r"\b([1-9][0-9]{5})\b", str(text))
    return match.group(1) if match else None


def render_quotation_editor(
    quotation_id: Optional[int],
    on_back: Callable[[], None],
    on_view_results: Callable[[int, Dict[str, Any]], None],
) -> None:
    """Renders quotation header and line items editor with automatic logistics and prominent line pricing."""
    is_edit = quotation_id is not None
    title = f"Edit Quotation (ID: {quotation_id})" if is_edit else "Create Commercial Quotation"
    st.subheader(title)

    # Initialize form state in st.session_state if not already present
    state_key = f"form_quote_{quotation_id or 'new'}"
    if state_key not in st.session_state:
        if is_edit:
            try:
                with st.spinner("Fetching quotation..."):
                    q_data = api_get_quotation(quotation_id)
                
                delivery_loc = q_data.get("delivery_location") or ""
                extracted_pin = _extract_pin(delivery_loc) or DEFAULT_DELIVERY_PIN
                dist = float(q_data.get("distance_km", 0.0))
                trate = float(q_data.get("trip_rate", 0.0))

                if dist <= 0 or trate <= 0:
                    logistics = api_resolve_logistics(pincode=extracted_pin, address=delivery_loc)
                    dist = logistics.get("distance_km", 782.0)
                    trate = logistics.get("trip_rate", DEFAULT_TRIP_RATE)

                st.session_state[state_key] = {
                    "enquiry_no": q_data.get("enquiry_no", ""),
                    "customer_name": q_data.get("customer_name", ""),
                    "project_name": q_data.get("project_name") or "",
                    "delivery_location": delivery_loc,
                    "delivery_pin": extracted_pin,
                    "distance_km": dist,
                    "trip_rate": trate,
                    "calculation_mode": q_data.get("calculation_mode", DEFAULT_CALCULATION_MODE),
                    "lines": [dict(DEFAULT_LINE, **l) for l in q_data.get("lines", [])] or [dict(DEFAULT_LINE)],
                }
            except APIError as err:
                st.error(f"Failed to fetch quotation: {err.message}")
                if st.button("Back to Dashboard"):
                    on_back()
                return
        else:
            # Generate suggested enquiry code for new quotation matching Excel format
            import time
            code = f"MKTGENQ26-{int(time.time()) % 100000:05d}"
            # Resolve initial Bagodar defaults
            init_logistics = api_resolve_logistics(pincode=DEFAULT_DELIVERY_PIN)
            st.session_state[state_key] = {
                "enquiry_no": code,
                "customer_name": "",
                "project_name": "",
                "delivery_location": "Bagodar, Jharkhand",
                "delivery_pin": DEFAULT_DELIVERY_PIN,
                "distance_km": init_logistics.get("distance_km", 782.0),
                "trip_rate": init_logistics.get("trip_rate", DEFAULT_TRIP_RATE),
                "calculation_mode": DEFAULT_CALCULATION_MODE,
                "lines": [dict(DEFAULT_LINE)],
            }

    data = st.session_state[state_key]

    # Top action bar
    col_nav, col_space = st.columns([3, 7])
    with col_nav:
        if st.button("← Back to Dashboard", use_container_width=True):
            on_back()
            return

    # Quotation Header Section
    with st.container(border=True):
        st.markdown("#### 1. Header & Commercial Logistics")
        st.caption("Enter normal customer & site details. Road distance and freight rates are calculated automatically from the Pincode.")

        h_col1, h_col2, h_col3 = st.columns(3)

        with h_col1:
            data["enquiry_no"] = st.text_input("Enquiry Number *", value=data["enquiry_no"])
            data["customer_name"] = st.text_input("Customer Name *", value=data["customer_name"], placeholder="e.g. Acme Infrastructure Ltd.")

        with h_col2:
            data["project_name"] = st.text_input("Project Name", value=data["project_name"], placeholder="e.g. Water Supply Project")
            data["delivery_location"] = st.text_input("Delivery Place / Site Address", value=data["delivery_location"], placeholder="e.g. Pune, Maharashtra")

        with h_col3:
            prev_pin = data.get("delivery_pin", DEFAULT_DELIVERY_PIN)
            new_pin = st.text_input(
                "Delivery Pincode (6 digits) *",
                value=prev_pin,
                max_chars=6,
                help="6-digit Indian PIN code. System computes road distance from MPPL supplying plant.",
            ).strip()

            mode_options = ["REFERENCE", "NORMAL", "LEGACY"]
            curr_mode = data.get("calculation_mode", DEFAULT_CALCULATION_MODE)
            current_mode_idx = mode_options.index(curr_mode) if curr_mode in mode_options else 0
            selected_mode = st.selectbox(
                "Calculation Mode Policy",
                options=mode_options,
                index=current_mode_idx,
                help="REFERENCE: Authoritative Excel model. NORMAL: Dynamic route costing.",
            )
            data["calculation_mode"] = selected_mode

        # Trigger automatic logistics calculation if PIN or Mode changed
        if new_pin != prev_pin or data.get("distance_km", 0.0) <= 0:
            data["delivery_pin"] = new_pin
            if len(new_pin) == 6 and new_pin.isdigit() and not new_pin.startswith("0"):
                log_res = api_resolve_logistics(pincode=new_pin, address=data.get("delivery_location"), mode=selected_mode)
                data["distance_km"] = float(log_res.get("distance_km", 782.0))
                data["trip_rate"] = float(log_res.get("trip_rate", DEFAULT_TRIP_RATE))
                # Auto-fill delivery location if empty
                if not data.get("delivery_location"):
                    data["delivery_location"] = log_res.get("formatted_address", "")
                st.session_state[f"logistics_cache_{new_pin}"] = log_res

        # Display rich logistics summary badge (No manual asking of distance / trip!)
        log_info = st.session_state.get(f"logistics_cache_{data.get('delivery_pin')}") or api_resolve_logistics(
            pincode=data.get("delivery_pin", DEFAULT_DELIVERY_PIN), mode=data["calculation_mode"]
        )

        st.info(
            f"📍 **Logistics (Auto-Calculated from PIN {data.get('delivery_pin')})**: "
            f"**Location:** {log_info.get('formatted_address', 'India')} • "
            f"**Distance:** `{data['distance_km']:.1f} km` • "
            f"**Freight Rate:** `{format_inr(data['trip_rate'])} / trailer` "
            f"*(Supplying Plant: {log_info.get('plant_name', 'Mega Pipes Jabalpur/Bagodar')})*"
        )

        # Advanced Logistics Overrides (Optional expander - kept collapsed so users are not asked)
        with st.expander("⚙️ Advanced Logistics & Freight Overrides (Optional)", expanded=False):
            st.caption("Distance and trip rates are auto-calculated from your PIN code above. You may manually fine-tune them here if needed.")
            log_c1, log_c2 = st.columns(2)
            with log_c1:
                data["distance_km"] = st.number_input(
                    "Distance (km) [Auto-populated]",
                    min_value=1.0,
                    value=float(data["distance_km"]),
                    step=10.0,
                )
            with log_c2:
                data["trip_rate"] = st.number_input(
                    "Trip Rate (₹ / Trailer) [Auto-populated]",
                    min_value=1000.0,
                    value=float(data["trip_rate"]),
                    step=1000.0,
                )

    # Quotation Lines Section
    lines: List[Dict[str, Any]] = data["lines"]

    st.markdown(f"#### 2. Quotation Line Items ({len(lines)})")
    st.caption("Configure pipe geometry and specify unit prices or custom pricing overrides per line, matching the Excel Quotation model.")

    line_to_delete = None
    line_to_duplicate = None

    for idx, line in enumerate(lines):
        line_num = idx + 1
        line["line_number"] = line_num

        with st.container(border=True):
            header_col, action_col = st.columns([5, 2])
            with header_col:
                msp = float(line.get("manual_selling_price", 0.0))
                price_badge = f" | 🏷️ Selling Price: **₹{msp:,.2f}** (Override)" if msp > 0 else " | ⚙️ Price: **Auto-Formula**"
                st.markdown(
                    f"**Line #{line_num}:** `{line.get('size_mm')} mm` ({line.get('basis')}) &times; `{line.get('thickness_mm')} mm WT` &bull; `{line.get('quantity')} {line.get('unit')}`{price_badge}"
                )
            with action_col:
                btn_c1, btn_c2 = st.columns(2)
                with btn_c1:
                    if st.button("Duplicate", key=f"dup_{idx}", use_container_width=True):
                        line_to_duplicate = idx
                with btn_c2:
                    if len(lines) > 1 and st.button("Delete", key=f"del_{idx}", use_container_width=True):
                        line_to_delete = idx

            # Section A: Primary Technical Specifications
            l_col1, l_col2, l_col3, l_col4 = st.columns(4)
            with l_col1:
                line["size_mm"] = st.number_input(
                    "Size / OD (mm)",
                    min_value=1.0,
                    max_value=5000.0,
                    value=float(line.get("size_mm", 1226.0)),
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
                    value=float(line.get("thickness_mm", 8.4)),
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
                    value=float(line.get("quantity", 2000.0)),
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

            # Section B: PROMINENT PRICE ENTRY & COMMERCIAL SPECIFICATION (Matching Excel QUOTE sheet)
            st.markdown("##### 💰 Commercial Price Entry & Steel Costing")
            p_col1, p_col2, p_col3 = st.columns(3)

            with p_col1:
                line["manual_selling_price"] = st.number_input(
                    "Selling Price Override (₹)",
                    min_value=0.0,
                    value=float(line.get("manual_selling_price", 0.0)),
                    step=100.0,
                    key=f"msp_{idx}",
                    help="Custom unit price. Leave 0 to calculate automatically from the steel & conversion formula.",
                )

            with p_col2:
                line["steel_rate"] = st.number_input(
                    "Steel Base Rate (₹/MT)",
                    min_value=0.0,
                    value=float(line.get("steel_rate", DEFAULT_STEEL_RATE)),
                    step=500.0,
                    key=f"steel_{idx}",
                    help="Base steel delivered cost per MT (default ₹62,000 from Excel SETTINGS).",
                )

            with p_col3:
                line["margin"] = st.number_input(
                    "Margin (₹/MT)",
                    min_value=0.0,
                    value=float(line.get("margin", 0.0)),
                    step=100.0,
                    key=f"margin_{idx}",
                    help="Commercial profit margin per MT (default ₹0).",
                )

            if float(line["manual_selling_price"]) > 0:
                line["override_reason"] = st.text_input(
                    "Price Override Reason *",
                    value=str(line.get("override_reason", "")),
                    placeholder="e.g. Approved client tender rate / Special project pricing",
                    key=f"reason_{idx}",
                )
                st.caption(f"ℹ️ Manual price active: Unit price locked at **₹{line['manual_selling_price']:,.2f}** per {line['unit']}.")
            else:
                st.caption(f"ℹ️ Formula costing active: Unit price calculated dynamically from Steel Rate (₹{line['steel_rate']:,.0f}) + Coating + Freight.")

            # Section C: Collapsible Advanced Commercial Details (Conversion, Coatings, Yield)
            with st.expander(f"⚙️ Advanced Cost Parameters (Line #{line_num})", expanded=False):
                cr_col1, cr_col2, cr_col3 = st.columns(3)
                with cr_col1:
                    line["conversion"] = st.number_input(
                        "Conversion Cost (₹/MT)", min_value=0.0, value=float(line.get("conversion", DEFAULT_CONVERSION)), key=f"conv_{idx}"
                    )
                    line["yield_pct"] = st.number_input(
                        "Yield Burning (%)", min_value=0.0, max_value=20.0, value=float(line.get("yield_pct", 2.0)), key=f"yield_{idx}"
                    )
                    line["grade_extra"] = st.number_input(
                        "Grade Extra (₹/MT)", min_value=0.0, value=float(line.get("grade_extra", 0.0)), key=f"grade_{idx}"
                    )

                with cr_col2:
                    line["ext_rate"] = st.number_input(
                        "Ext Coating Rate (₹/m²)", min_value=0.0, value=float(line.get("ext_rate", 500.0)), key=f"extrate_{idx}"
                    )
                    line["int_rate"] = st.number_input(
                        "Int Coating Rate (₹/m²)", min_value=0.0, value=float(line.get("int_rate", 250.0)), key=f"intrate_{idx}"
                    )
                    line["handling"] = st.number_input(
                        "Handling (₹/MT)", min_value=0.0, value=float(line.get("handling", 100.0)), key=f"hand_{idx}"
                    )

                with cr_col3:
                    line["financial_pct"] = st.number_input(
                        "Financial Cost (%)", min_value=0.0, max_value=20.0, value=float(line.get("financial_pct", 1.0)), key=f"fin_{idx}"
                    )
                    line["gst_pct"] = st.number_input(
                        "GST (%)", min_value=0.0, max_value=100.0, value=float(line.get("gst_pct", 18.0)), key=f"gst_{idx}"
                    )
                    line["manual_fc"] = st.number_input(
                        "Manual FC Override (₹)", min_value=0.0, value=float(line.get("manual_fc", 0.0)), key=f"mfc_{idx}"
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

        # Build clean delivery location including PIN if not already present
        deliv = data["delivery_location"].strip()
        pin = data.get("delivery_pin", "").strip()
        if pin and pin not in deliv:
            deliv = f"{deliv} - {pin}".strip(" -")

        # Ensure positive distance and trip rate
        dist_km = float(data.get("distance_km", 0.0))
        trip_rate = float(data.get("trip_rate", 0.0))
        if dist_km <= 0 or trip_rate <= 0:
            logistics = api_resolve_logistics(pincode=pin, address=deliv, mode=data["calculation_mode"])
            dist_km = logistics.get("distance_km", 782.0)
            trip_rate = logistics.get("trip_rate", DEFAULT_TRIP_RATE)

        # Build clean payload according to backend QuotationCreateRequest
        return {
            "enquiry_no": enq,
            "customer_name": cust,
            "project_name": data["project_name"].strip() or None,
            "delivery_location": deliv or None,
            "distance_km": float(dist_km),
            "trip_rate": float(trip_rate),
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
                    "steel_rate": float(l.get("steel_rate", DEFAULT_STEEL_RATE)),
                    "grade_extra": float(l.get("grade_extra", 0.0)),
                    "yield_pct": float(l.get("yield_pct", 2.0)),
                    "conversion": float(l.get("conversion", DEFAULT_CONVERSION)),
                    "testing": float(l.get("testing", 0.0)),
                    "margin": float(l.get("margin", 0.0)),
                    "ext_rate": float(l.get("ext_rate", 500.0)),
                    "int_rate": float(l.get("int_rate", 250.0)),
                    "handling": float(l.get("handling", 100.0)),
                    "financial_pct": float(l.get("financial_pct", 1.0)),
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
                    with st.spinner("Saving quotation draft..."):
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
                    with st.spinner("Saving quotation and executing commercial calculation..."):
                        if is_edit:
                            saved = api_update_quotation(quotation_id, payload)
                        else:
                            saved = api_create_quotation(payload)

                        target_id = saved["id"]
                        snapshot = api_calculate_quotation(target_id)
                        st.success(f"Quotation calculated successfully! Grand Total: {format_inr(snapshot.get('grand_total'))}")
                        del st.session_state[state_key]
                        on_view_results(target_id, snapshot)
                except APIError as e:
                    st.error(f"Operation failed: {e.message}")
