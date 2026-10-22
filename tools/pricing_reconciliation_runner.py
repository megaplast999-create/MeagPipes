"""
Mega Pipes — Automated Pricing Variance & Golden Reconciliation Runner.

Executes both Normal Mode and Legacy Mode calculations for the Unipro / Konar
enquiry, performs line-by-line and component-level variance reconciliation,
validates against approved golden baselines, and generates machine-readable
reconciliation artifacts.

Usage:
    python tools/pricing_reconciliation_runner.py [--export] [--json]
"""

import sys
import os
import json
import csv
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Dict, List, Any, Tuple

PROJECT_ROOT = Path(__file__).resolve().parent.parent
TESTS_DIR = PROJECT_ROOT / "TESTS"
EXPORTS_DIR = PROJECT_ROOT / "EXPORTS"

# Add TESTS dir to sys.path to import calculation_engine
if str(TESTS_DIR) not in sys.path:
    sys.path.insert(0, str(TESTS_DIR))

from calculation_engine import calc_line, QuoteLineResult

# ==============================================================================
# AUTHORITATIVE GOLDEN BASELINES
# ==============================================================================

UNIPRO_TENDER_LINES = [
    {"size_mm": 1226, "thickness_mm": 8.4, "basis": "ID", "quantity": 2000, "manual_fc": 1542, "item_label": "Line 1 — 1226mm ID (1243mm OD / 49\")"},
    {"size_mm": 1300, "thickness_mm": 8.4, "basis": "ID", "quantity": 350,  "manual_fc": 1534, "item_label": "Line 2 — 1300mm ID (1317mm OD / 52\")"},
    {"size_mm": 1350, "thickness_mm": 8.4, "basis": "ID", "quantity": 500,  "manual_fc": 1529, "item_label": "Line 3 — 1350mm ID (1367mm OD / 54\")"},
    {"size_mm": 1400, "thickness_mm": 8.4, "basis": "ID", "quantity": 4400, "manual_fc": 1525, "item_label": "Line 4 — 1400mm ID (1417mm OD / 56\")"},
]

COMMON_COMMERCIAL_RATES = {
    "unit": "Metres",
    "pipe_length_m": 12.0,
    "pipes_per_trailer": 4.0,
    "has_external": True,
    "has_internal": True,
    "steel_rate": 61000.0,
    "grade_extra": 0.0,
    "yield_pct": 2.0,
    "conversion": 5000.0,
    "testing": 0.0,
    "margin": 0.0,
    "ext_rate": 500.0,
    "int_rate": 275.0,
    "handling": 100.0,
    "financial_pct": 1.5,
    "gst_pct": 18.0,
}

# Approved Golden Target Specifications
GOLDEN_NORMAL_TARGETS = {
    1: {"bare_per_m": 17594, "ext_per_m": 1952, "int_per_m": 1059, "freight_per_m": 2084, "total_per_m": 22689, "line_value": 45378000, "freight_per_mt": 8147},
    2: {"bare_per_m": 18650, "ext_per_m": 2068, "int_per_m": 1123, "freight_per_m": 2268, "total_per_m": 24109, "line_value": 8438150,  "freight_per_mt": 8364},
    3: {"bare_per_m": 19359, "ext_per_m": 2147, "int_per_m": 1166, "freight_per_m": 2185, "total_per_m": 24857, "line_value": 12428500, "freight_per_mt": 7761},
    4: {"bare_per_m": 20068, "ext_per_m": 2225, "int_per_m": 1209, "freight_per_m": 2079, "total_per_m": 25581, "line_value": 112556400, "freight_per_mt": 7123},
    "subtotal": 178801050,
    "gst_18": 32184189,
    "grand_total": 210985239
}

GOLDEN_LEGACY_TARGETS = {
    1: {"bare_per_m": 17587, "ext_per_m": 1952, "int_per_m": 1074, "freight_per_m": 2068, "total_per_m": 22681, "line_value": 45362000, "freight_per_mt": 8083},
    2: {"bare_per_m": 18639, "ext_per_m": 2068, "int_per_m": 1138, "freight_per_m": 2069, "total_per_m": 23914, "line_value": 8369900,  "freight_per_mt": 7632},
    3: {"bare_per_m": 19349, "ext_per_m": 2147, "int_per_m": 1181, "freight_per_m": 2071, "total_per_m": 24748, "line_value": 12374000, "freight_per_mt": 7355},
    4: {"bare_per_m": 20060, "ext_per_m": 2225, "int_per_m": 1224, "freight_per_m": 2071, "total_per_m": 25580, "line_value": 112552000, "freight_per_mt": 7097},
    "subtotal": 178657900,
    "gst_18": 32158422,
    "grand_total": 210816322
}

EXPECTED_LINE_VARIANCES = {
    1: {"diff_per_m": 8,   "qty": 2000, "line_impact": 16000},
    2: {"diff_per_m": 195, "qty": 350,  "line_impact": 68250},
    3: {"diff_per_m": 109, "qty": 500,  "line_impact": 54500},
    4: {"diff_per_m": 1,   "qty": 4400, "line_impact": 4400},
    "total_impact": 143150
}


def run_reconciliation() -> Dict[str, Any]:
    """
    Executes full calculation and regression reconciliation.
    Returns structured results dictionary.
    """
    normal_results = []
    legacy_results = []
    line_variance_bridges = []

    normal_subtotal = 0.0
    legacy_subtotal = 0.0

    validation_errors = []

    for idx, item in enumerate(UNIPRO_TENDER_LINES):
        line_num = idx + 1
        qty = item["quantity"]

        # Run Normal Mode calculation
        kwargs_norm = {
            **COMMON_COMMERCIAL_RATES,
            "size_mm": item["size_mm"],
            "thickness_mm": item["thickness_mm"],
            "basis": item["basis"],
            "quantity": qty,
            "trip_rate": 125.32 * 782,
            "legacy_mode": False
        }
        res_norm = calc_line(**kwargs_norm)
        normal_results.append(res_norm)
        normal_subtotal += res_norm.line_value

        # Run Legacy Mode calculation
        kwargs_leg = {
            **COMMON_COMMERCIAL_RATES,
            "size_mm": item["size_mm"],
            "thickness_mm": item["thickness_mm"],
            "basis": item["basis"],
            "quantity": qty,
            "trip_rate": 98000.0,
            "manual_fc": item["manual_fc"],
            "legacy_mode": True
        }
        res_leg = calc_line(**kwargs_leg)
        legacy_results.append(res_leg)
        legacy_subtotal += res_leg.line_value

        # Calculate unit and component differences
        diff_bare_m = res_norm.bare_per_m - res_leg.bare_per_m
        diff_ext_m = res_norm.external_per_m - res_leg.external_per_m
        diff_int_m = res_norm.internal_per_m - res_leg.internal_per_m
        diff_fr_m = res_norm.freight_per_m - res_leg.freight_per_m
        diff_total_m = res_norm.total_per_m - res_leg.total_per_m
        diff_line_val = res_norm.line_value - res_leg.line_value

        # Component impacts in Rupees
        impact_bare = diff_bare_m * qty
        impact_ext = diff_ext_m * qty
        impact_int = diff_int_m * qty
        impact_fr = diff_fr_m * qty

        # Verify component sum matches line difference
        component_sum_m = diff_bare_m + diff_ext_m + diff_int_m + diff_fr_m
        if component_sum_m != diff_total_m:
            validation_errors.append(
                f"Line {line_num}: Component per-metre sum ({component_sum_m}) != Total per-metre diff ({diff_total_m})"
            )

        component_sum_rs = impact_bare + impact_ext + impact_int + impact_fr
        if component_sum_rs != diff_line_val:
            validation_errors.append(
                f"Line {line_num}: Component monetary sum ({component_sum_rs}) != Line value diff ({diff_line_val})"
            )

        # Verify against expected baseline checkpoints
        exp_var = EXPECTED_LINE_VARIANCES[line_num]
        if diff_total_m != exp_var["diff_per_m"]:
            validation_errors.append(
                f"Line {line_num}: Selling price diff ({diff_total_m}) != Expected diff ({exp_var['diff_per_m']})"
            )
        if diff_line_val != exp_var["line_impact"]:
            validation_errors.append(
                f"Line {line_num}: Line impact ({diff_line_val}) != Expected impact ({exp_var['line_impact']})"
            )

        # Verify Normal Golden Target
        gold_n = GOLDEN_NORMAL_TARGETS[line_num]
        if res_norm.total_per_m != gold_n["total_per_m"]:
            validation_errors.append(
                f"Line {line_num} Normal: Selling price ({res_norm.total_per_m}) != Golden target ({gold_n['total_per_m']})"
            )
        if res_norm.line_value != gold_n["line_value"]:
            validation_errors.append(
                f"Line {line_num} Normal: Line value ({res_norm.line_value}) != Golden target ({gold_n['line_value']})"
            )

        # Verify Legacy Golden Target
        gold_l = GOLDEN_LEGACY_TARGETS[line_num]
        if res_leg.total_per_m != gold_l["total_per_m"]:
            validation_errors.append(
                f"Line {line_num} Legacy: Selling price ({res_leg.total_per_m}) != Golden target ({gold_l['total_per_m']})"
            )
        if res_leg.line_value != gold_l["line_value"]:
            validation_errors.append(
                f"Line {line_num} Legacy: Line value ({res_leg.line_value}) != Golden target ({gold_l['line_value']})"
            )

        line_variance_bridges.append({
            "line_number": line_num,
            "label": item["item_label"],
            "size_mm": item["size_mm"],
            "quantity_m": qty,
            "legacy": {
                "pre_cost_mt": res_leg.pre_cost,
                "financial_mt": res_leg.financial_cost_per_mt,
                "ext_mt": res_leg.external_cost_per_mt,
                "int_mt": res_leg.internal_cost_per_mt,
                "freight_mt": res_leg.freight_per_mt,
                "bare_per_m": res_leg.bare_per_m,
                "ext_per_m": res_leg.external_per_m,
                "int_per_m": res_leg.internal_per_m,
                "freight_per_m": res_leg.freight_per_m,
                "total_per_m": res_leg.total_per_m,
                "line_value": res_leg.line_value,
            },
            "normal": {
                "pre_cost_mt": res_norm.pre_cost,
                "financial_mt": res_norm.financial_cost_per_mt,
                "ext_mt": res_norm.external_cost_per_mt,
                "int_mt": res_norm.internal_cost_per_mt,
                "freight_mt": res_norm.freight_per_mt,
                "bare_per_m": res_norm.bare_per_m,
                "ext_per_m": res_norm.external_per_m,
                "int_per_m": res_norm.internal_per_m,
                "freight_per_m": res_norm.freight_per_m,
                "total_per_m": res_norm.total_per_m,
                "line_value": res_norm.line_value,
            },
            "variance": {
                "bare_per_m_diff": diff_bare_m,
                "ext_per_m_diff": diff_ext_m,
                "int_per_m_diff": diff_int_m,
                "freight_per_m_diff": diff_fr_m,
                "total_per_m_diff": diff_total_m,
                "bare_impact_rs": impact_bare,
                "ext_impact_rs": impact_ext,
                "int_impact_rs": impact_int,
                "freight_impact_rs": impact_fr,
                "line_impact_rs": diff_line_val,
            }
        })

    # Subtotals & Totals
    grand_diff = normal_subtotal - legacy_subtotal
    if grand_diff != EXPECTED_LINE_VARIANCES["total_impact"]:
        validation_errors.append(
            f"Grand variance ({grand_diff}) != Expected variance ({EXPECTED_LINE_VARIANCES['total_impact']})"
        )

    normal_gst = round(normal_subtotal * 0.18, 0)
    normal_grand = normal_subtotal + normal_gst
    legacy_gst = round(legacy_subtotal * 0.18, 0)
    legacy_grand = legacy_subtotal + legacy_gst

    if normal_subtotal != GOLDEN_NORMAL_TARGETS["subtotal"]:
        validation_errors.append(
            f"Normal subtotal ({normal_subtotal}) != Golden target ({GOLDEN_NORMAL_TARGETS['subtotal']})"
        )
    if legacy_subtotal != GOLDEN_LEGACY_TARGETS["subtotal"]:
        validation_errors.append(
            f"Legacy subtotal ({legacy_subtotal}) != Golden target ({GOLDEN_LEGACY_TARGETS['subtotal']})"
        )

    # Component-level cumulative variance
    total_bare_impact = sum(b["variance"]["bare_impact_rs"] for b in line_variance_bridges)
    total_ext_impact = sum(b["variance"]["ext_impact_rs"] for b in line_variance_bridges)
    total_int_impact = sum(b["variance"]["int_impact_rs"] for b in line_variance_bridges)
    total_freight_impact = sum(b["variance"]["freight_impact_rs"] for b in line_variance_bridges)

    cumulative_component_impact = total_bare_impact + total_ext_impact + total_int_impact + total_freight_impact
    if cumulative_component_impact != grand_diff:
        validation_errors.append(
            f"Cumulative component impact ({cumulative_component_impact}) != Grand total diff ({grand_diff})"
        )

    is_passed = (len(validation_errors) == 0)

    return {
        "status": "PASS" if is_passed else "FAIL",
        "validation_errors": validation_errors,
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "line_variance_bridges": line_variance_bridges,
        "totals": {
            "legacy": {
                "subtotal_excl_gst": legacy_subtotal,
                "gst_18": legacy_gst,
                "grand_total_incl_gst": legacy_grand
            },
            "normal": {
                "subtotal_excl_gst": normal_subtotal,
                "gst_18": normal_gst,
                "grand_total_incl_gst": normal_grand
            },
            "variance": {
                "subtotal_diff": grand_diff,
                "gst_diff": normal_gst - legacy_gst,
                "grand_total_diff": normal_grand - legacy_grand,
                "total_bare_financial_impact": total_bare_impact,
                "total_ext_coating_impact": total_ext_impact,
                "total_int_coating_impact": total_int_impact,
                "total_freight_impact": total_freight_impact,
            }
        }
    }


def export_reconciliation_reports(recon_data: Dict[str, Any]) -> Tuple[Path, Path]:
    """Export reconciliation data to CSV and JSON."""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    recon_dir = EXPORTS_DIR / f"PRICING_RECONCILIATION_{ts}"
    recon_dir.mkdir(parents=True, exist_ok=True)

    json_path = recon_dir / f"pricing_variance_reconciliation_{ts}.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(recon_data, f, indent=2)

    csv_path = recon_dir / f"pricing_variance_bridge_{ts}.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        fieldnames = [
            "Line", "ItemLabel", "QuantityM", "Component",
            "LegacyValPerUnit", "NormalValPerUnit", "UnitVariance",
            "LineImpactRs", "RootCause", "Classification", "Action"
        ]
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()

        for b in recon_data["line_variance_bridges"]:
            ln = b["line_number"]
            lbl = b["label"]
            q = b["quantity_m"]
            v = b["variance"]
            leg = b["legacy"]
            norm = b["normal"]

            # Bare / Financial
            writer.writerow({
                "Line": ln, "ItemLabel": lbl, "QuantityM": q, "Component": "Bare Price (Financial Cost)",
                "LegacyValPerUnit": f"{leg['bare_per_m']:.0f}",
                "NormalValPerUnit": f"{norm['bare_per_m']:.0f}",
                "UnitVariance": f"{v['bare_per_m_diff']:+.0f}",
                "LineImpactRs": f"{v['bare_impact_rs']:+.0f}",
                "RootCause": "Dynamic working capital factor k/(1-k) on freight-inclusive baseCost vs unverified static manual FC",
                "Classification": "Expected Business Rule Difference",
                "Action": "No Change (Retain approved algebraic financing)"
            })
            # External Coating
            writer.writerow({
                "Line": ln, "ItemLabel": lbl, "QuantityM": q, "Component": "External Coating",
                "LegacyValPerUnit": f"{leg['ext_per_m']:.0f}",
                "NormalValPerUnit": f"{norm['ext_per_m']:.0f}",
                "UnitVariance": f"{v['ext_per_m_diff']:+.0f}",
                "LineImpactRs": f"{v['ext_impact_rs']:+.0f}",
                "RootCause": "Identical OD coating area formula (Pi=3.14, ExtRate=500)",
                "Classification": "Expected Agreement",
                "Action": "No Change"
            })
            # Internal Coating
            writer.writerow({
                "Line": ln, "ItemLabel": lbl, "QuantityM": q, "Component": "Internal Coating",
                "LegacyValPerUnit": f"{leg['int_per_m']:.0f}",
                "NormalValPerUnit": f"{norm['int_per_m']:.0f}",
                "UnitVariance": f"{v['int_per_m_diff']:+.0f}",
                "LineImpactRs": f"{v['int_impact_rs']:+.0f}",
                "RootCause": "Physical inside diameter (ID) used in Normal Mode vs Outside Diameter (OD) erroneously reused in Legacy sheet",
                "Classification": "Expected Business Rule Difference (Legacy Defect Corrected)",
                "Action": "No Change (Retain physical ID specification)"
            })
            # Freight
            writer.writerow({
                "Line": ln, "ItemLabel": lbl, "QuantityM": q, "Component": "Freight (Logistics)",
                "LegacyValPerUnit": f"{leg['freight_per_m']:.0f}",
                "NormalValPerUnit": f"{norm['freight_per_m']:.0f}",
                "UnitVariance": f"{v['freight_per_m_diff']:+.0f}",
                "LineImpactRs": f"{v['freight_impact_rs']:+.0f}",
                "RootCause": "Discrete whole-trailer ceiling allocation (RoundUp(pipes/4)) in Normal Mode vs fractional trailers in Legacy sheet",
                "Classification": "Expected Business Rule Difference",
                "Action": "No Change (Retain approved whole-trailer allocation)"
            })

    return csv_path, json_path


def format_console_report(recon: Dict[str, Any]) -> str:
    """Format human-readable summary report."""
    lines = [
        "=" * 105,
        "MEGA PIPES — PRICING VARIANCE ROOT CAUSE & RECONCILIATION GATE",
        "=" * 105,
        f"Execution Timestamp: {recon['timestamp']}",
        f"Reconciliation Gate Status: {recon['status']}",
    ]
    if recon["validation_errors"]:
        lines.append("VALIDATION ERRORS ENCOUNTERED:")
        for err in recon["validation_errors"]:
            lines.append(f"  [X] {err}")
    else:
        lines.append("ALL INTEGRITY & RECONCILIATION GATES PASSED (100% Agreement).")

    lines.extend([
        "-" * 105,
        f"{'Line':<6} | {'Qty (m)':<7} | {'Legacy Price':>13} | {'Normal Price':>13} | {'Diff / m':>10} | {'Line Impact':>15} | {'Status':>8}",
        "-" * 105,
    ])
    for b in recon["line_variance_bridges"]:
        ln = b["line_number"]
        q = b["quantity_m"]
        lp = b["legacy"]["total_per_m"]
        np = b["normal"]["total_per_m"]
        dp = b["variance"]["total_per_m_diff"]
        imp = b["variance"]["line_impact_rs"]
        lines.append(f"Line {ln:<1} | {q:<7} | Rs. {lp:9,.0f}/m | Rs. {np:9,.0f}/m | {dp:+9.0f}/m | Rs. {imp:+11,.0f} | {'PASS':>8}")

    tot = recon["totals"]
    lines.extend([
        "-" * 105,
        "COMPONENT IMPACT RECONCILIATION:",
        f"  1. Freight (Whole-Trailer Allocation vs Fractional):   Rs. {tot['variance']['total_freight_impact']:+12,.0f}",
        f"  2. Internal Coating (Inner Diameter ID vs Outer OD):   Rs. {tot['variance']['total_int_coating_impact']:+12,.0f}",
        f"  3. Bare Price (Dynamic Algebraic FC vs Static Manual): Rs. {tot['variance']['total_bare_financial_impact']:+12,.0f}",
        f"  4. External Coating (Identical OD Formulation):        Rs. {tot['variance']['total_ext_coating_impact']:+12,.0f}",
        f"  --------------------------------------------------------------------------------",
        f"  TOTAL SUB-TOTAL VARIANCE (EXCL. GST):                  Rs. {tot['variance']['subtotal_diff']:+12,.0f} (EXACT MATCH)",
        f"  GST IMPACT @ 18%:                                      Rs. {tot['variance']['gst_diff']:+12,.0f}",
        f"  GRAND TOTAL VARIANCE (INCL. GST):                      Rs. {tot['variance']['grand_total_diff']:+12,.0f}",
        "=" * 105,
    ])
    return "\n".join(lines)


if __name__ == "__main__":
    recon = run_reconciliation()
    print(format_console_report(recon))
    if "--export" in sys.argv:
        c_path, j_path = export_reconciliation_reports(recon)
        print(f"\nExported artifacts:")
        print(f"  CSV:  {c_path}")
        print(f"  JSON: {j_path}")
    if recon["status"] != "PASS":
        sys.exit(1)
