"""P4 Line Agreement and Golden Quotation Validation Suite.

Validates that Python reference calculations agree with:
1. The authoritative Option 1 golden targets in DOCS/ANTIGRAVITY_MASTER_PROMPT.md.
2. The verified VBA CalcLine output recorded in TESTS/vba_p3_golden_20261008.txt.

Tolerance requirement: Absolute difference <= 1 rupee across all components and totals.
Exact requirement: Exact integer match on approved rounded selling prices and line values.
"""

from pathlib import Path
import re
import pytest

from calculation_engine import (
    CalculationError,
    QuoteLineResult,
    calc_line,
    calc_quote,
    geometry,
)

ROOT_DIR = Path(__file__).resolve().parents[1]
VBA_FIXTURE_PATH = ROOT_DIR / "TESTS" / "vba_p3_golden_20261008.txt"

# Approved Option 1 Golden Specifications
GOLDEN_INPUTS = [
    {"id": 1226, "wt": 8.4, "qty": 2000, "legacy_fc": 1542},
    {"id": 1300, "wt": 8.4, "qty": 350, "legacy_fc": 1534},
    {"id": 1350, "wt": 8.4, "qty": 500, "legacy_fc": 1529},
    {"id": 1400, "wt": 8.4, "qty": 4400, "legacy_fc": 1525},
]

EXPECTED_NORMAL = [
    {
        "line": 1, "od": 1243, "id": 1226, "wt": 8.4, "metres": 2000,
        "bare_per_m": 17594, "ext_per_m": 1952, "int_per_m": 1059, "freight_per_m": 2084,
        "total_per_m": 22689, "line_value": 45378000,
        "freight_per_mt": 8147,
    },
    {
        "line": 2, "od": 1317, "id": 1300, "wt": 8.4, "metres": 350,
        "bare_per_m": 18650, "ext_per_m": 2068, "int_per_m": 1123, "freight_per_m": 2268,
        "total_per_m": 24109, "line_value": 8438150,
        "freight_per_mt": 8364,
    },
    {
        "line": 3, "od": 1367, "id": 1350, "wt": 8.4, "metres": 500,
        "bare_per_m": 19359, "ext_per_m": 2147, "int_per_m": 1166, "freight_per_m": 2185,
        "total_per_m": 24857, "line_value": 12428500,
        "freight_per_mt": 7761,
    },
    {
        "line": 4, "od": 1417, "id": 1400, "wt": 8.4, "metres": 4400,
        "bare_per_m": 20068, "ext_per_m": 2225, "int_per_m": 1209, "freight_per_m": 2079,
        "total_per_m": 25581, "line_value": 112556400,
        "freight_per_mt": 7123,
    },
]

EXPECTED_NORMAL_GRAND_TOTAL = 178801050

EXPECTED_LEGACY = [
    {
        "line": 1, "od": 1243, "id": 1226, "wt": 8.4, "metres": 2000,
        "bare_per_m": 17587, "ext_per_m": 1952, "int_per_m": 1074, "freight_per_m": 2068,
        "total_per_m": 22681, "line_value": 45362000,
        "freight_per_mt": 8083, "financial_per_mt": 1542,
    },
    {
        "line": 2, "od": 1317, "id": 1300, "wt": 8.4, "metres": 350,
        "bare_per_m": 18639, "ext_per_m": 2068, "int_per_m": 1138, "freight_per_m": 2069,
        "total_per_m": 23914, "line_value": 8369900,
        "freight_per_mt": 7632, "financial_per_mt": 1534,
    },
    {
        "line": 3, "od": 1367, "id": 1350, "wt": 8.4, "metres": 500,
        "bare_per_m": 19349, "ext_per_m": 2147, "int_per_m": 1181, "freight_per_m": 2071,
        "total_per_m": 24748, "line_value": 12374000,
        "freight_per_mt": 7355, "financial_per_mt": 1529,
    },
    {
        "line": 4, "od": 1417, "id": 1400, "wt": 8.4, "metres": 4400,
        "bare_per_m": 20060, "ext_per_m": 2225, "int_per_m": 1224, "freight_per_m": 2071,
        "total_per_m": 25580, "line_value": 112552000,
        "freight_per_mt": 7097, "financial_per_mt": 1525,
    },
]

EXPECTED_LEGACY_GRAND_TOTAL = 178657900


def get_normal_line_results() -> list[QuoteLineResult]:
    trip_rate = 125.32 * 782
    results = []
    for item in GOLDEN_INPUTS:
        res = calc_line(
            size_mm=item["id"],
            thickness_mm=item["wt"],
            basis="ID",
            quantity=item["qty"],
            unit="Metres",
            trip_rate=trip_rate,
            legacy_mode=False,
        )
        results.append(res)
    return results


def get_legacy_line_results() -> list[QuoteLineResult]:
    trip_rate = 98000.0
    results = []
    for item in GOLDEN_INPUTS:
        res = calc_line(
            size_mm=item["id"],
            thickness_mm=item["wt"],
            basis="ID",
            quantity=item["qty"],
            unit="Metres",
            trip_rate=trip_rate,
            manual_fc=item["legacy_fc"],
            legacy_mode=True,
        )
        results.append(res)
    return results


def parse_vba_fixture():
    assert VBA_FIXTURE_PATH.is_file(), f"VBA fixture not found: {VBA_FIXTURE_PATH}"
    content = VBA_FIXTURE_PATH.read_text(encoding="utf-8")
    lines_data = {}
    current_key = None

    for line in content.splitlines():
        line = line.strip()
        m_line = re.match(r"LINE=(\d+)\s+MODE=(\w+)", line)
        if m_line:
            idx = int(m_line.group(1))
            mode = m_line.group(2)
            current_key = (idx, mode)
            lines_data[current_key] = {}
            continue
        if current_key and "=" in line:
            parts = line.split("=")
            if len(parts) == 2:
                lines_data[current_key][parts[0].strip()] = parts[1].strip()

    return lines_data


# --- TESTS ---

@pytest.mark.parametrize("line_idx", [0, 1, 2, 3])
def test_normal_mode_per_line_against_golden_option1(line_idx):
    results = get_normal_line_results()
    actual = results[line_idx]
    expected = EXPECTED_NORMAL[line_idx]

    # Tolerance check: abs diff <= 1 rupee
    assert abs(actual.bare_per_m - expected["bare_per_m"]) <= 1
    assert abs(actual.external_per_m - expected["ext_per_m"]) <= 1
    assert abs(actual.internal_per_m - expected["int_per_m"]) <= 1
    assert abs(actual.freight_per_m - expected["freight_per_m"]) <= 1
    assert abs(actual.total_per_m - expected["total_per_m"]) <= 1
    assert abs(actual.line_value - expected["line_value"]) <= 1

    # Exact integer assertions
    assert actual.bare_per_m == expected["bare_per_m"]
    assert actual.external_per_m == expected["ext_per_m"]
    assert actual.internal_per_m == expected["int_per_m"]
    assert actual.freight_per_m == expected["freight_per_m"]
    assert actual.total_per_m == expected["total_per_m"]
    assert actual.line_value == expected["line_value"]
    assert actual.od_mm == expected["od"]
    assert actual.inner_d_mm == expected["id"]


def test_normal_mode_grand_total_option1():
    results = get_normal_line_results()
    grand_total = sum(r.line_value for r in results)
    assert grand_total == EXPECTED_NORMAL_GRAND_TOTAL
    assert abs(grand_total - EXPECTED_NORMAL_GRAND_TOTAL) <= 1


@pytest.mark.parametrize("line_idx", [0, 1, 2, 3])
def test_legacy_mode_per_line_against_golden(line_idx):
    results = get_legacy_line_results()
    actual = results[line_idx]
    expected = EXPECTED_LEGACY[line_idx]

    # Tolerance check: abs diff <= 1 rupee
    assert abs(actual.bare_per_m - expected["bare_per_m"]) <= 1
    assert abs(actual.external_per_m - expected["ext_per_m"]) <= 1
    assert abs(actual.internal_per_m - expected["int_per_m"]) <= 1
    assert abs(actual.freight_per_m - expected["freight_per_m"]) <= 1
    assert abs(actual.total_per_m - expected["total_per_m"]) <= 1
    assert abs(actual.line_value - expected["line_value"]) <= 1

    # Exact integer assertions
    assert actual.bare_per_m == expected["bare_per_m"]
    assert actual.external_per_m == expected["ext_per_m"]
    assert actual.internal_per_m == expected["int_per_m"]
    assert actual.freight_per_m == expected["freight_per_m"]
    assert actual.total_per_m == expected["total_per_m"]
    assert actual.line_value == expected["line_value"]


def test_legacy_mode_grand_total():
    results = get_legacy_line_results()
    grand_total = sum(r.line_value for r in results)
    assert grand_total == EXPECTED_LEGACY_GRAND_TOTAL
    assert abs(grand_total - EXPECTED_LEGACY_GRAND_TOTAL) <= 1


@pytest.mark.parametrize("line_idx", [0, 1, 2, 3])
def test_python_vs_vba_normal_mode_agreement(line_idx):
    """Compares Python calculation against verified VBA CalcLine fixture."""
    vba_data = parse_vba_fixture()
    key = (line_idx + 1, "CURRENT")
    assert key in vba_data, f"Line {line_idx + 1} CURRENT mode missing in VBA fixture"
    vba_row = vba_data[key]

    results = get_normal_line_results()
    py_res = results[line_idx]

    vba_bare = float(vba_row["BarePerM"])
    vba_ext = float(vba_row["ExtPerM"])
    vba_int = float(vba_row["IntPerM"])
    vba_fr = float(vba_row["FreightPerM"])
    vba_total_m = float(vba_row["TotalPerM"])
    vba_val = float(vba_row["LineValue"])

    # Agreement within 1 rupee
    assert abs(py_res.bare_per_m - vba_bare) <= 1
    assert abs(py_res.external_per_m - vba_ext) <= 1
    assert abs(py_res.internal_per_m - vba_int) <= 1
    assert abs(py_res.freight_per_m - vba_fr) <= 1
    assert abs(py_res.total_per_m - vba_total_m) <= 1
    assert abs(py_res.line_value - vba_val) <= 1

    # Exact equality
    assert py_res.bare_per_m == vba_bare
    assert py_res.external_per_m == vba_ext
    assert py_res.internal_per_m == vba_int
    assert py_res.freight_per_m == vba_fr
    assert py_res.total_per_m == vba_total_m
    assert py_res.line_value == vba_val


@pytest.mark.parametrize("line_idx", [0, 1, 2, 3])
def test_python_vs_vba_legacy_mode_agreement(line_idx):
    """Compares Python calculation against verified VBA CalcLine fixture."""
    vba_data = parse_vba_fixture()
    key = (line_idx + 1, "LEGACY")
    assert key in vba_data, f"Line {line_idx + 1} LEGACY mode missing in VBA fixture"
    vba_row = vba_data[key]

    results = get_legacy_line_results()
    py_res = results[line_idx]

    vba_bare = float(vba_row["BarePerM"])
    vba_ext = float(vba_row["ExtPerM"])
    vba_int = float(vba_row["IntPerM"])
    vba_fr = float(vba_row["FreightPerM"])
    vba_total_m = float(vba_row["TotalPerM"])
    vba_val = float(vba_row["LineValue"])

    # Agreement within 1 rupee
    assert abs(py_res.bare_per_m - vba_bare) <= 1
    assert abs(py_res.external_per_m - vba_ext) <= 1
    assert abs(py_res.internal_per_m - vba_int) <= 1
    assert abs(py_res.freight_per_m - vba_fr) <= 1
    assert abs(py_res.total_per_m - vba_total_m) <= 1
    assert abs(py_res.line_value - vba_val) <= 1

    # Exact equality
    assert py_res.bare_per_m == vba_bare
    assert py_res.external_per_m == vba_ext
    assert py_res.internal_per_m == vba_int
    assert py_res.freight_per_m == vba_fr
    assert py_res.total_per_m == vba_total_m
    assert py_res.line_value == vba_val


def test_invalid_geometry_guards():
    with pytest.raises(CalculationError):
        calc_line(10, 8.4, "ID", 1)

    with pytest.raises(CalculationError):
        calc_line(100, 60, "OD", 1)
