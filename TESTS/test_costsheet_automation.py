"""Cost-sheet automation: engine vs. the reference sheets the sales team uses.

Expected numbers are transcribed from:
  * MKTGENQ26-27-594 - Unipro 2.xlsx  -> 'CS Stacking 10-09-26' (stacking, 93,000/trip)
  * same workbook                      -> 'CS NESTING10-09-2026' (pooled 103 trips, 98,000/trip)
"""
import math
import pytest
from calculation_engine import (
    REFERENCE_POLICY, calc_costsheet, calc_line, financial_cost, freight_detail,
)

UNIPRO = [(1226, 2000), (1300, 350), (1350, 500), (1400, 4400)]


def _lines(**over):
    base = dict(basis="ID", thickness_mm=8.4, steel_rate=62000.0, yield_pct=2.0, conversion=4000.0,
                ext_rate=500.0, int_rate=250.0, trip_rate=93000.0, handling=100.0, financial_pct=1.0)
    base.update(over)
    return [{**base, "size_mm": s, "quantity": q} for s, q in UNIPRO]


# ---------------------------------------------------------------- stacking sheet
def test_stacking_sheet_exact_with_typed_fc():
    """Typed FC 1024/1014/1011/1008 -> the sheet's own figures, to the rupee."""
    fcs = [1024, 1014, 1011, 1008]
    lines = [{**l, "manual_fc": fc} for l, fc in zip(_lines(), fcs)]
    cs = calc_costsheet(lines)
    assert [r.freight_per_mt for r in cs.lines] == [7676, 7248, 6985, 6740]
    assert [r.bare_per_m for r in cs.lines] == [17459, 18503, 19209, 19915]
    assert [r.external_per_m for r in cs.lines] == [1952, 2068, 2147, 2225]
    assert [r.internal_per_m for r in cs.lines] == [976, 1034, 1074, 1113]
    assert [r.freight_per_m for r in cs.lines] == [1964, 1965, 1966, 1967]
    assert [r.total_per_m for r in cs.lines] == [22351, 23570, 24396, 25220]
    assert cs.subtotal_excl_gst == 176_117_500
    assert cs.total_trips == pytest.approx(151.0417, abs=1e-4)      # sheet: "151 trips"


def test_stacking_sheet_weighted_total_column():
    fcs = [1024, 1014, 1011, 1008]
    cs = calc_costsheet([{**l, "manual_fc": fc} for l, fc in zip(_lines(), fcs)])
    assert cs.weighted["freight_per_mt"] == pytest.approx(7016.435335, abs=1e-3)   # G37
    assert cs.weighted["financial_cost_per_mt"] == pytest.approx(1012.5177, abs=1e-3)  # G32
    assert cs.weighted["subtotal_per_mt"] == pytest.approx(86713.1436, abs=1e-3)   # H49


def test_auto_fc_is_fixed_point_of_sheet_check_cell():
    """Sheet row 41 (Grand total incl GST x 1%) must equal the FC that fed it."""
    r = calc_line(**_lines()[0], **REFERENCE_POLICY)
    assert r.financial_cost_per_mt == pytest.approx(1031.23, abs=0.01)
    sub_total = r.pre_cost + r.financial_cost_per_mt + r.external_cost_per_mt \
        + r.internal_cost_per_mt + r.freight_per_mt
    assert sub_total * 1.18 * 0.01 == pytest.approx(r.financial_cost_per_mt, abs=1e-6)


def test_auto_fc_no_manual_entry_needed():
    cs = calc_costsheet(_lines())
    assert [r.total_per_m for r in cs.lines] == [22353, 23573, 24400, 25224]
    assert cs.subtotal_excl_gst == 176_142_150


def test_simple_fc_method_understates_fc():
    base = 67240 + 7630.37 + 3815.19 + 7676
    simple = financial_cost(67240, 7630.37, 3815.19, 7676, 1.0, 18.0, method="SIMPLE")
    algebraic = financial_cost(67240, 7630.37, 3815.19, 7676, 1.0, 18.0, method="ALGEBRAIC")
    assert simple == round(base * 1.18 * 0.01) == 1019
    assert algebraic == pytest.approx(1031.23, abs=0.01)


# ---------------------------------------------------------------- nesting sheet
def test_nesting_sheet_pooled_trips():
    """103 manually planned trips pooled over all lines; conv/FC differ per line."""
    conv = [5000, 10000, 10000, 5000]
    fcs = [1480, 1569, 1569, 1480]
    lines = [{**l, "steel_rate": 61000.0, "conversion": c, "manual_fc": f, "trip_rate": 98000.0,
              "financial_pct": 1.5} for l, c, f in zip(_lines(), conv, fcs)]
    cs = calc_costsheet(lines, pooled_trips=103)
    assert [r.freight_per_mt for r in cs.lines] == [5070] * 4           # sheet C37:F37
    assert [r.bare_per_m for r in cs.lines] == [17571, 20004, 20768, 20047]
    assert [r.freight_per_m for r in cs.lines] == [1297, 1375, 1427, 1480]
    assert [r.total_per_m for r in cs.lines] == [21796, 24481, 25416, 24865]
    assert cs.subtotal_excl_gst == 174_274_350                          # sheet G49


# ---------------------------------------------------------------- freight rules
def test_fractional_vs_whole_trailers_line2():
    kw = dict(trip_rate=93000, metres=350, pipe_length_m=12, mass_mt_per_m_val=0.27108512676,
              pipes_per_trailer=4, handling_rs_per_mt=100)
    phys, trips, fr = freight_detail(**kw, whole_trailers=False)
    assert (phys, round(trips, 4), fr) == (8, 7.2917, 7248)
    phys, trips, fr = freight_detail(**kw, whole_trailers=True)
    assert (phys, trips, fr) == (8, 8, 7942)


def test_manual_trips_override_wins():
    _, trips, fr = freight_detail(93000, 350, 12, 0.27108512676, 4, 100, trips_override=7.5)
    assert trips == 7.5 and fr == math.ceil(93000 * 7.5 / (350 * 0.27108512676)) + 100


def test_policy_is_explicit_not_inferred_from_steel_rate():
    """Old engine flipped method when steel_rate happened to be 62000."""
    a = calc_line(**_lines(steel_rate=62000.0)[0], **REFERENCE_POLICY)
    b = calc_line(**_lines(steel_rate=62001.0)[0], **REFERENCE_POLICY)
    assert a.freight_per_mt == b.freight_per_mt == 7676
    assert a.internal_cost_per_mt == pytest.approx(b.internal_cost_per_mt, rel=1e-6)
