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
        PRODUCT_STANDARDS,
        PIPE_FAMILIES,
        DIAMETER_BASES,
        STEEL_GRADES,
        EXTERNAL_COATINGS,
        INTERNAL_COATINGS,
        QUANTITY_UNITS,
        CUSTOM_PIPE_SIZE_OPTIONS,
        PRICE_UNITS,
        STANDARD_SIZES,
        STANDARD_THICKNESSES,
        GRADES_BY_FAMILY,
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
        PRODUCT_STANDARDS,
        PIPE_FAMILIES,
        DIAMETER_BASES,
        STEEL_GRADES,
        EXTERNAL_COATINGS,
        INTERNAL_COATINGS,
        QUANTITY_UNITS,
        CUSTOM_PIPE_SIZE_OPTIONS,
        PRICE_UNITS,
        STANDARD_SIZES,
        STANDARD_THICKNESSES,
        GRADES_BY_FAMILY,
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

            # Section A: 13 Pipe Configuration Options from Excel QUOTE sheet
            st.markdown(f"##### 📋 Pipe Line Specifications (Line #{line_num})")

            # Row 1: Standards, Family, Sizing Basis & Custom Sizing
            r1_c1, r1_c2, r1_c3, r1_c4 = st.columns(4)
            with r1_c1:
                std_opts = PRODUCT_STANDARDS
                curr_std = line.get("product_standard", "IS-3589-2001")
                idx_std = std_opts.index(curr_std) if curr_std in std_opts else 0
                line["product_standard"] = st.selectbox("Product Standard", std_opts, index=idx_std, key=f"std_{idx}")

            with r1_c2:
                fam_opts = PIPE_FAMILIES
                curr_fam = line.get("pipe_family", "HSAW")
                idx_fam = fam_opts.index(curr_fam) if curr_fam in fam_opts else 0
                line["pipe_family"] = st.selectbox("Pipe Family", fam_opts, index=idx_fam, key=f"fam_{idx}")

            with r1_c3:
                basis_opts = DIAMETER_BASES
                curr_basis = line.get("basis", "ID").upper()
                idx_basis = basis_opts.index(curr_basis) if curr_basis in basis_opts else 0
                line["basis"] = st.selectbox("Diameter Basis", basis_opts, index=idx_basis, key=f"basis_{idx}")

            with r1_c4:
                cust_opts = CUSTOM_PIPE_SIZE_OPTIONS
                curr_cust = line.get("custom_pipe_size", "NO").upper()
                idx_cust = cust_opts.index(curr_cust) if curr_cust in cust_opts else 0
                line["custom_pipe_size"] = st.selectbox("Custom Pipe Size", cust_opts, index=idx_cust, key=f"cust_{idx}")

            # Row 2: Dimensions & Material Grade (Standard Dropdowns or Custom Input matching Excel)
            r2_c1, r2_c2, r2_c3, r2_c4 = st.columns(4)
            is_custom_size = (line.get("custom_pipe_size", "NO").upper() == "YES")

            with r2_c1:
                if not is_custom_size:
                    size_key = (line.get("pipe_family", "HSAW"), line.get("basis", "ID").upper())
                    avail_sizes = STANDARD_SIZES.get(size_key, [508.0, 610.0, 762.0, 1226.0, 1400.0])
                    curr_val = float(line.get("size_mm", avail_sizes[0]))
                    if curr_val not in avail_sizes:
                        avail_sizes = sorted(avail_sizes + [curr_val])
                    s_idx = avail_sizes.index(curr_val)
                    line["size_mm"] = st.selectbox(
                        "Size (mm) *",
                        options=avail_sizes,
                        index=s_idx,
                        key=f"size_sel_{idx}",
                        help="Standard diameter from Excel master tblSizes",
                    )
                else:
                    line["size_mm"] = st.number_input(
                        "Custom Size (mm) *",
                        min_value=1.0,
                        max_value=5000.0,
                        value=float(line.get("size_mm", 1226.0)),
                        step=1.0,
                        key=f"size_num_{idx}",
                    )

            with r2_c2:
                if not is_custom_size:
                    thk_opts = STANDARD_THICKNESSES.get(line.get("pipe_family", "HSAW"), [6.4, 7.1, 7.9, 8.4, 9.5])
                    curr_thk = float(line.get("thickness_mm", thk_opts[0]))
                    if curr_thk not in thk_opts:
                        thk_opts = sorted(thk_opts + [curr_thk])
                    t_idx = thk_opts.index(curr_thk)
                    line["thickness_mm"] = st.selectbox(
                        "Wall Thickness (mm) *",
                        options=thk_opts,
                        index=t_idx,
                        key=f"thk_sel_{idx}",
                        help="Standard thickness from Excel master tblThickness",
                    )
                else:
                    line["thickness_mm"] = st.number_input(
                        "Custom Wall Thickness (mm) *",
                        min_value=0.1,
                        max_value=100.0,
                        value=float(line.get("thickness_mm", 8.4)),
                        step=0.1,
                        key=f"thk_num_{idx}",
                    )

            with r2_c3:
                fam = line.get("pipe_family", "HSAW")
                grade_opts = GRADES_BY_FAMILY.get(fam, STEEL_GRADES)
                curr_grade = line.get("steel_grade", "Fe 410")
                if curr_grade not in grade_opts:
                    grade_opts = [curr_grade] + grade_opts
                idx_grade = grade_opts.index(curr_grade)
                line["steel_grade"] = st.selectbox("Steel Grade", grade_opts, index=idx_grade, key=f"grade_{idx}")

            with r2_c4:
                line["pipe_length_m"] = st.number_input(
                    "Pipe Length (m)",
                    min_value=6.0,
                    max_value=13.0,
                    value=float(line.get("pipe_length_m", 12.0)),
                    step=0.5,
                    key=f"len_{idx}",
                )

            # Row 3: Coatings & Quantities
            r3_c1, r3_c2, r3_c3, r3_c4 = st.columns(4)
            with r3_c1:
                ext_opts = EXTERNAL_COATINGS
                curr_ext = line.get("external_coating", "3LPE")
                idx_ext = ext_opts.index(curr_ext) if curr_ext in ext_opts else 0
                line["external_coating"] = st.selectbox("External Coating", ext_opts, index=idx_ext, key=f"ext_{idx}")
                line["has_external"] = line["external_coating"].upper() not in ("BARE", "NONE", "")

            with r3_c2:
                int_opts = INTERNAL_COATINGS
                curr_int = line.get("internal_coating", "Liquid Epoxy")
                idx_int = int_opts.index(curr_int) if curr_int in int_opts else 0
                line["internal_coating"] = st.selectbox("Internal Coating", int_opts, index=idx_int, key=f"int_{idx}")
                line["has_internal"] = line["internal_coating"].upper() not in ("NONE", "BARE", "")

            with r3_c3:
                line["quantity"] = st.number_input(
                    "Quantity *",
                    min_value=1.0,
                    value=float(line.get("quantity", 2000.0)),
                    step=10.0,
                    key=f"qty_{idx}",
                )

            with r3_c4:
                unit_opts = QUANTITY_UNITS
                curr_unit = line.get("unit", "Metres")
                idx_unit = unit_opts.index(curr_unit) if curr_unit in unit_opts else 0
                line["unit"] = st.selectbox("Quantity Unit", unit_opts, index=idx_unit, key=f"unit_{idx}")

            # Row 4: Commercial Pricing & Selling Price Override (Direct Price Entry Matching Excel)
            r4_c1, r4_c2, r4_c3, r4_c4 = st.columns(4)
            with r4_c1:
                line["manual_selling_price"] = st.number_input(
                    "Manual Selling Price (Optional)",
                    min_value=0.0,
                    value=float(line.get("manual_selling_price", 0.0)),
                    step=100.0,
                    key=f"msp_{idx}",
                    help="Custom target selling price. Leave 0 to calculate automatically from cost formula.",
                )

            with r4_c2:
                punit_opts = PRICE_UNITS
                curr_punit = line.get("price_unit", "₹/m")
                idx_punit = punit_opts.index(curr_punit) if curr_punit in punit_opts else 0
                line["price_unit"] = st.selectbox("Price Unit", punit_opts, index=idx_punit, key=f"punit_{idx}")

            with r4_c3:
                line["steel_rate"] = st.number_input(
                    "Steel Base Rate (₹/MT)",
                    min_value=0.0,
                    value=float(line.get("steel_rate", DEFAULT_STEEL_RATE)),
                    step=500.0,
                    key=f"steel_{idx}",
                    help="Base steel delivered cost per MT (default ₹61,000 / ₹62,000 from Excel SETTINGS).",
                )

            with r4_c4:
                line["pipes_per_trailer"] = st.number_input(
                    "Pipes / Trailer",
                    min_value=1.0,
                    max_value=20.0,
                    value=float(line.get("pipes_per_trailer", 4.0)),
                    step=1.0,
                    key=f"trailer_{idx}",
                )

            if float(line["manual_selling_price"]) > 0:
                line["override_reason"] = st.text_input(
                    "Price Override Reason *",
                    value=str(line.get("override_reason", "")),
                    placeholder="e.g. Approved client tender rate / Special project pricing",
                    key=f"reason_{idx}",
                )
                st.caption(f"ℹ️ Manual price active: Unit price locked at **₹{line['manual_selling_price']:,.2f}** ({line['price_unit']}).")
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
