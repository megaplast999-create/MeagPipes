"""Isolated unit tests for Synopsis Column C, I, M, and P precision and presentation settings.

Verifies:
1. VBA source in modSynopsis.bas uses unrounded =L/K with '#,##0.00' format for Column M.
2. VBA source in modSynopsis.bas uses unrounded =12*J/1000 with '#,##0.00' format for Column P.
3. No integer ROUNDUP is present in Column M or Column P definitions.
4. Column C defaults to nominal whole-inch (=ROUND(D/25.4, 0)) and supports Exact (=D/25.4).
5. Column I defaults to physical dispatch joints (=CEILING(H/12, 1)) and supports 12m Equivalent (=H/12).
6. Mathematical precision matches Lodha reference unrounded values for all 5 lines.
7. Display formatting produces proper decimal places.
8. Diagnostic columns have zero impact on commercial pricing or grand totals.
"""

from decimal import Decimal
from pathlib import Path
import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
VBA_SYNOPSIS = REPO_ROOT / "VBA" / "Modules" / "modSynopsis.bas"
VBA_PRICING = REPO_ROOT / "VBA" / "Modules" / "modPricing.bas"
CALC_ENGINE = REPO_ROOT / "TESTS" / "calculation_engine.py"


def test_vba_synopsis_column_m_unrounded_ratio():
    """Verify Column M formula in modSynopsis.bas is =L/K formatted as #,##0.00."""
    content = VBA_SYNOPSIS.read_text(encoding="utf-8")
    
    assert 'ws.Cells(outRow, 13).Formula = "=L" & outRow & "/K" & outRow' in content, (
        "Column M must be defined as unrounded '=L{row}/K{row}'"
    )
    assert 'ws.Cells(outRow, 13).NumberFormat = "#,##0.00"' in content, (
        "Column M must have NumberFormat '#,##0.00'"
    )
    assert 'ROUNDUP(L" & outRow & "/K' not in content, (
        "Column M must NOT use integer ROUNDUP"
    )


def test_vba_synopsis_column_p_unrounded_weight():
    """Verify Column P formula in modSynopsis.bas is =12*J/1000 formatted as #,##0.00."""
    content = VBA_SYNOPSIS.read_text(encoding="utf-8")
    
    assert 'ws.Cells(outRow, 16).Formula = "=12*J" & outRow & "/1000"' in content, (
        "Column P must be defined as unrounded '=12*J{row}/1000'"
    )
    assert 'ws.Cells(outRow, 16).NumberFormat = "#,##0.00"' in content, (
        "Column P must have NumberFormat '#,##0.00'"
    )
    assert 'ROUNDUP(12*J" & outRow' not in content, (
        "Column P must NOT use integer ROUNDUP"
    )


def test_vba_synopsis_column_c_presentation_setting():
    """Verify Column C supports both nominal whole-inch default and exact converted inches."""
    content = VBA_SYNOPSIS.read_text(encoding="utf-8")
    
    # Must support default nominal whole-inch
    assert 'ws.Cells(outRow, 3).Formula = "=ROUND(D" & outRow & "/25.4,0)"' in content
    assert 'ws.Cells(outRow, 3).NumberFormat = "0"' in content
    
    # Must support exact converted inches
    assert 'ws.Cells(outRow, 3).Formula = "=D" & outRow & "/25.4"' in content
    assert 'ws.Cells(outRow, 3).NumberFormat = "#,##0.00"' in content
    assert 'SynopsisInchDisplay' in content


def test_vba_synopsis_column_i_presentation_setting():
    """Verify Column I supports both physical dispatch joints default and 12m equivalent."""
    content = VBA_SYNOPSIS.read_text(encoding="utf-8")
    
    # Must support default dispatch joints
    assert 'ws.Cells(outRow, 9).Formula = "=CEILING(H" & outRow & "/12,1)"' in content
    assert 'ws.Cells(outRow, 9).NumberFormat = "#,##0"' in content
    
    # Must support 12m equivalent
    assert 'ws.Cells(outRow, 9).Formula = "=H" & outRow & "/12"' in content
    assert 'ws.Cells(outRow, 9).NumberFormat = "#,##0.00"' in content
    assert 'SynopsisPipeCountDisplay' in content


def test_vba_synopsis_settings_readers_and_setters():
    """Verify modSynopsis.bas includes helper functions to read and toggle settings."""
    content = VBA_SYNOPSIS.read_text(encoding="utf-8")
    
    assert "Public Function GetSynopsisSetting(" in content or "Private Function GetSynopsisSetting(" in content
    assert "Public Sub SetSynopsisPresentationSettings(" in content


def test_lodha_five_lines_synopsis_precision():
    """Verify precision of Column M and Column P against Lodha 5 reference lines."""
    lodha_lines = [
        # (line, od, wt, qty_m, ref_cbm_ratio, ref_pipe_wt_mt)
        (1, 1050.0, 12.0, 365.0, 3.5891, 3.686204),
        (2, 1074.0, 12.0, 25035.0, 3.6701, 3.771434),
        (3, 900.0, 12.0, 120.0, 3.0823, 3.153515),
        (4, 700.0, 12.0, 80.0, 2.4066, 2.443264),
        (5, 1000.0, 12.0, 100.0, 3.4201, 3.508641),
    ]

    for line, od, wt, qty_m, ref_cbm, ref_pipe_wt in lodha_lines:
        d_od = Decimal(str(od))
        d_wt = Decimal(str(wt))
        d_qty = Decimal(str(qty_m))

        # Mass kg/m = (OD - WT) * WT * 0.0246615
        d_mass_kg = (d_od - d_wt) * d_wt * Decimal("0.0246615")

        # Tonnage MT = Qty * mass / 1000
        d_mt = (d_qty * d_mass_kg) / Decimal(1000)

        # Volume CBM = (OD/1000)^2 * Qty
        vol_cbm = ((od / 1000.0) ** 2) * qty_m

        # Column M: unrounded ratio = Volume / MT
        calc_cbm_ratio = vol_cbm / float(d_mt)
        assert abs(calc_cbm_ratio - ref_cbm) < 1e-4, (
            f"Line {line} CBM ratio mismatch: calc={calc_cbm_ratio}, ref={ref_cbm}"
        )

        # Column P: unrounded 12m pipe weight MT = 12 * mass / 1000
        calc_pipe_wt = float(Decimal("12") * d_mass_kg / Decimal(1000))
        assert abs(calc_pipe_wt - ref_pipe_wt) < 1e-6, (
            f"Line {line} pipe wt mismatch: calc={calc_pipe_wt}, ref={ref_pipe_wt}"
        )

        # Verify 2 decimal place string rendering
        str_cbm = f"{calc_cbm_ratio:,.2f}"
        str_pipe_wt = f"{calc_pipe_wt:,.2f}"
        assert len(str_cbm.split(".")[1]) == 2
        assert len(str_pipe_wt.split(".")[1]) == 2


def test_synopsis_columns_m_and_p_isolated_from_pricing():
    """Verify that synopsis columns are strictly informational and do not feed into pricing modules."""
    if VBA_PRICING.exists():
        pricing_code = VBA_PRICING.read_text(encoding="utf-8")
        assert "SYNOPSIS" not in pricing_code
        assert "CBM Ratio" not in pricing_code
        assert "Weight of pipe" not in pricing_code
    
    calc_engine_code = CALC_ENGINE.read_text(encoding="utf-8")
    assert "SYNOPSIS" not in calc_engine_code
    assert "cbm_ratio" not in calc_engine_code
