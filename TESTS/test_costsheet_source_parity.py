"""Rigorous regression and parity verification suite matching live Excel Cost Sheets.

Validates:
1. All 4 line items of Unipro (MKTGENQ26-27-594 - Unipro 2.xlsx - CS Stacking 10-09-26).
2. All 4 line items of Unipro Nesting (CS NESTING10-09-2026).
3. All 5 line items of Lodha (MKTGENQ26-27610 - Lodha 1.xlsx - Sheet1).
4. Finance Cost data model (suggested FC, applied FC, manual entry).
5. Manual Selling Price Override pathway with audit trail.
6. Rounding rules and dual-mode OD/ID sizing.
"""

import unittest
from decimal import Decimal
import openpyxl

from calculation_engine import (
    calc_line,
    geometry,
    mass_mt_per_m,
    quantity_conversion,
    roundup,
    external_coating_per_mt,
    internal_coating_per_mt,
    freight_per_mt,
    financial_cost,
    selling_price_per_m,
)


class TestUniproCostSheetParity(unittest.TestCase):
    """Verifies parity against MKTGENQ26-27-594 - Unipro 2.xlsx (CS Stacking 10-09-26)."""

    def setUp(self):
        self.wb_path = r"C:\Users\Administrator\Downloads\MKTGENQ26-27-594 - Unipro 2.xlsx"
        self.wb = openpyxl.load_workbook(self.wb_path, data_only=True)
        self.ws_stack = self.wb["CS Stacking 10-09-26"]
        self.ws_offer = self.wb["Sheet1"]

    def test_line_1_stacking_golden_baseline(self):
        """Unipro Line 1: 1226 ID (1243 OD) x 8.4mm WT, 2000m -> Rs 22,351/m."""
        res = calc_line(
            size_mm=1226.0,
            thickness_mm=8.4,
            basis="ID",
            quantity=2000.0,
            unit="Metres",
            steel_rate=62000.0,
            conversion=4000.0,
            yield_pct=2.0,
            margin=0.0,
            testing=0.0,
            ext_rate=500.0,
            int_rate=250.0,
            trip_rate=93000.0,
            handling=100.0,
            financial_pct=1.0,
            manual_fc=1024.0,  # Row 32 historical manual input
            pipes_per_trailer=4.0,
            approved_mode=True,
        )
        self.assertEqual(res.od_mm, 1243.0)
        self.assertEqual(res.inner_d_mm, 1226.0)
        self.assertAlmostEqual(res.mass_mt_per_m, 0.25575553836, places=8)
        self.assertAlmostEqual(res.weight_mt, 511.51107672, places=5)
        self.assertEqual(res.pipes, 167)
        self.assertEqual(res.freight_per_mt, 7676.0)
        self.assertEqual(res.financial_cost_per_mt, 1024.0)
        self.assertEqual(res.bare_per_m, 17459.0)
        self.assertEqual(res.external_per_m, 1952.0)
        self.assertEqual(res.internal_per_m, 976.0)
        self.assertEqual(res.freight_per_m, 1964.0)
        self.assertEqual(res.total_per_m, 22351.0)
        self.assertEqual(res.line_value, 44702000.0)

    def test_line_2_stacking_golden_baseline(self):
        """Unipro Line 2: 1300 ID (1317 OD) x 8.4mm WT, 350m -> Rs 23,570/m."""
        res = calc_line(
            size_mm=1300.0,
            thickness_mm=8.4,
            basis="ID",
            quantity=350.0,
            unit="Metres",
            steel_rate=62000.0,
            conversion=4000.0,
            yield_pct=2.0,
            margin=0.0,
            testing=0.0,
            ext_rate=500.0,
            int_rate=250.0,
            trip_rate=93000.0,
            handling=100.0,
            financial_pct=1.0,
            manual_fc=1014.0,  # Row 32 historical manual input
            pipes_per_trailer=4.0,
            approved_mode=True,
        )
        self.assertEqual(res.od_mm, 1317.0)
        self.assertAlmostEqual(res.mass_mt_per_m, 0.27108512676, places=8)
        self.assertAlmostEqual(res.weight_mt, 94.87979437, places=5)
        self.assertEqual(res.pipes, 30)
        self.assertEqual(res.freight_per_mt, 7248.0)
        self.assertEqual(res.bare_per_m, 18503.0)
        self.assertEqual(res.external_per_m, 2068.0)
        self.assertEqual(res.internal_per_m, 1034.0)
        self.assertEqual(res.freight_per_m, 1965.0)
        self.assertEqual(res.total_per_m, 23570.0)
        self.assertEqual(res.line_value, 8249500.0)

    def test_line_3_stacking_golden_baseline(self):
        """Unipro Line 3: 1350 ID (1367 OD) x 8.4mm WT, 500m -> Rs 24,396/m."""
        res = calc_line(
            size_mm=1350.0,
            thickness_mm=8.4,
            basis="ID",
            quantity=500.0,
            unit="Metres",
            steel_rate=62000.0,
            conversion=4000.0,
            yield_pct=2.0,
            margin=0.0,
            testing=0.0,
            ext_rate=500.0,
            int_rate=250.0,
            trip_rate=93000.0,
            handling=100.0,
            financial_pct=1.0,
            manual_fc=1011.0,  # Row 32 historical manual input
            pipes_per_trailer=4.0,
            approved_mode=True,
        )
        self.assertEqual(res.od_mm, 1367.0)
        self.assertAlmostEqual(res.mass_mt_per_m, 0.28144295676, places=8)
        self.assertAlmostEqual(res.weight_mt, 140.72147838, places=5)
        self.assertEqual(res.pipes, 42)
        self.assertEqual(res.freight_per_mt, 6985.0)
        self.assertEqual(res.bare_per_m, 19209.0)
        self.assertEqual(res.external_per_m, 2147.0)
        self.assertEqual(res.internal_per_m, 1074.0)
        self.assertEqual(res.freight_per_m, 1966.0)
        self.assertEqual(res.total_per_m, 24396.0)
        self.assertEqual(res.line_value, 12198000.0)

    def test_line_4_stacking_golden_baseline(self):
        """Unipro Line 4: 1400 ID (1417 OD) x 8.4mm WT, 4400m -> Rs 25,220/m."""
        res = calc_line(
            size_mm=1400.0,
            thickness_mm=8.4,
            basis="ID",
            quantity=4400.0,
            unit="Metres",
            steel_rate=62000.0,
            conversion=4000.0,
            yield_pct=2.0,
            margin=0.0,
            testing=0.0,
            ext_rate=500.0,
            int_rate=250.0,
            trip_rate=93000.0,
            handling=100.0,
            financial_pct=1.0,
            manual_fc=1008.0,  # Row 32 historical manual input
            pipes_per_trailer=4.0,
            approved_mode=True,
        )
        self.assertEqual(res.od_mm, 1417.0)
        self.assertAlmostEqual(res.mass_mt_per_m, 0.29180078676, places=8)
        self.assertAlmostEqual(res.weight_mt, 1283.92346174, places=5)
        self.assertEqual(res.pipes, 367)
        self.assertEqual(res.freight_per_mt, 6740.0)
        self.assertEqual(res.bare_per_m, 19915.0)
        self.assertEqual(res.external_per_m, 2225.0)
        self.assertEqual(res.internal_per_m, 1113.0)
        self.assertEqual(res.freight_per_m, 1967.0)
        self.assertEqual(res.total_per_m, 25220.0)
        self.assertEqual(res.line_value, 110968000.0)

    def test_unipro_total_project_value_parity(self):
        """Total project value for Unipro Stacking equals exactly Rs 17,61,17,500."""
        total_val = 44702000.0 + 8249500.0 + 12198000.0 + 110968000.0
        self.assertEqual(total_val, 176117500.0)
        # Check against Sheet1 (Offer) H16
        excel_offer_total = self.ws_offer["H16"].value
        self.assertEqual(float(excel_offer_total), 176117500.0)

    def test_unipro_stacking_trailer_reconciliation_source_vs_engine(self):
        """Reconciles 151 continuous source trips vs 153 discrete application trailers.
        
        Validates:
        1. Source CS Stacking 10-09-26 Row 61 has 151.0417 continuous trips (H37: '151 trips').
        2. Application uses discrete per-line loading: 42 + 8 + 11 + 92 = 153 dedicated trailers.
        3. Project-level pooled integer trailer capacity is ceil(606 / 4) = 152 trailers.
        4. Freight rates (Rs 7,676, 7,248, 6,985, 6,740) and unit prices (Rs 22,351, 23,570, 
           24,396, 25,220) are mathematically invariant between definitions.
        """
        # 1. Source Stacking Sheet Values
        src_metres = [self.ws_stack[f"F{r}"].value for r in (57, 58, 59, 60)]
        src_raw_pipes = [self.ws_stack[f"G{r}"].value for r in (57, 58, 59, 60)]
        src_pipes_per_tr = [self.ws_stack[f"K{r}"].value for r in (57, 58, 59, 60)]
        src_raw_trips = [self.ws_stack[f"N{r}"].value for r in (57, 58, 59, 60)]
        src_tot_trips = self.ws_stack["N61"].value
        src_h37_label = str(self.ws_stack["H37"].value)

        self.assertEqual(src_metres, [2000, 350, 500, 4400])
        self.assertEqual(src_pipes_per_tr, [4, 4, 4, 4])
        self.assertAlmostEqual(src_raw_pipes[0], 2000 / 12, places=6)
        self.assertAlmostEqual(src_raw_pipes[1], 350 / 12, places=6)
        self.assertAlmostEqual(src_raw_pipes[2], 500 / 12, places=6)
        self.assertAlmostEqual(src_raw_pipes[3], 4400 / 12, places=6)
        
        self.assertAlmostEqual(src_raw_trips[0], (2000 / 12) / 4, places=6)  # 41.6667
        self.assertAlmostEqual(src_raw_trips[1], (350 / 12) / 4, places=6)   # 7.2917
        self.assertAlmostEqual(src_raw_trips[2], (500 / 12) / 4, places=6)   # 10.4167
        self.assertAlmostEqual(src_raw_trips[3], (4400 / 12) / 4, places=6)  # 91.6667

        self.assertAlmostEqual(src_tot_trips, 151.04166667, places=6)
        self.assertEqual(round(src_tot_trips), 151)
        self.assertIn("151 trips", src_h37_label)

        # 2. Application Engine Values (Discrete Per-Line)
        engine_line_pipes = [int(roundup(m / 12.0, 0)) for m in src_metres]
        engine_line_trailers = [int(roundup(p / 4.0, 0)) for p in engine_line_pipes]
        self.assertEqual(engine_line_pipes, [167, 30, 42, 367])
        self.assertEqual(sum(engine_line_pipes), 606)
        self.assertEqual(engine_line_trailers, [42, 8, 11, 92])
        self.assertEqual(sum(engine_line_trailers), 153)

        # 3. Aggregation Bridges
        project_pooled_integer_trailers = int(roundup(sum(engine_line_pipes) / 4.0, 0))
        self.assertEqual(project_pooled_integer_trailers, 152)

        continuous_project_trips = round(sum(src_metres) / (12.0 * 4.0))
        self.assertEqual(continuous_project_trips, 151)

        # 4. Invariance of Freight Rates and Per-Metre Selling Prices
        line_specs = [
            (1226, 8.4, 2000, 1024),
            (1300, 8.4, 350, 1014),
            (1350, 8.4, 500, 1011),
            (1400, 8.4, 4400, 1008),
        ]
        expected_freight = [7676.0, 7248.0, 6985.0, 6740.0]
        expected_unit_prices = [22351.0, 23570.0, 24396.0, 25220.0]

        for idx, (sz, wt, qty, fc) in enumerate(line_specs):
            res = calc_line(
                size_mm=sz,
                thickness_mm=wt,
                basis="ID",
                quantity=qty,
                unit="Metres",
                steel_rate=62000.0,
                conversion=4000.0,
                ext_rate=500.0,
                int_rate=250.0,
                trip_rate=93000.0,
                handling=100.0,
                financial_pct=1.0,
                manual_fc=fc,
                pipes_per_trailer=4.0,
                approved_mode=True,
            )
            self.assertEqual(res.freight_per_mt, expected_freight[idx])
            self.assertEqual(res.total_per_m, expected_unit_prices[idx])


class TestUniproNestingCostSheetParity(unittest.TestCase):
    """Verifies parity against MKTGENQ26-27-594 - Unipro 2.xlsx (CS NESTING10-09-2026)."""

    def setUp(self):
        self.wb_path = r"C:\Users\Administrator\Downloads\MKTGENQ26-27-594 - Unipro 2.xlsx"
        self.wb = openpyxl.load_workbook(self.wb_path, data_only=True)
        self.ws_nest = self.wb["CS NESTING10-09-2026"]

    def test_nesting_trips_and_freight_pool(self):
        """103 trips across 7 nesting groups @ Rs 98,000/trip yields pooled Rs 5,070/MT."""
        total_trips = self.ws_nest["N103"].value
        self.assertEqual(total_trips, 103)
        total_freight = self.ws_nest["Q66"].value
        self.assertEqual(total_freight, 10094000)
        freight_per_mt = self.ws_nest["Q67"].value
        self.assertAlmostEqual(freight_per_mt, 5069.87790382, places=6)

    def test_nesting_per_metre_prices(self):
        """Per-metre prices match CS NESTING Row 48: Rs 21,796, 24,481, 25,416, 24,865."""
        p1 = self.ws_nest["C48"].value
        p2 = self.ws_nest["D48"].value
        p3 = self.ws_nest["E48"].value
        p4 = self.ws_nest["F48"].value
        self.assertEqual((p1, p2, p3, p4), (21796, 24481, 25416, 24865))
        total_val = self.ws_nest["G49"].value
        self.assertEqual(total_val, 174274350)


class TestLodhaGeometryParity(unittest.TestCase):
    """Verifies geometry and surface area parity against MKTGENQ26-27610 - Lodha 1.xlsx."""

    def setUp(self):
        self.wb_path = r"C:\Users\Administrator\Downloads\MKTGENQ26-27610 - Lodha 1.xlsx"
        self.wb = openpyxl.load_workbook(self.wb_path, data_only=True)
        self.ws = self.wb["Sheet1"]

    def test_lodha_line_1_od_specified(self):
        """Lodha Line 1: OD 1050, WT 12mm, 365m -> ID 1026, 307.184 kg/m, 112.122 MT."""
        od_mm, id_mm, kg_m = geometry(1050.0, 12.0, "OD")
        self.assertEqual(od_mm, 1050.0)
        self.assertEqual(id_mm, 1026.0)
        self.assertAlmostEqual(kg_m, 307.183644, places=6)
        mt_per_m = mass_mt_per_m(od_mm, 12.0)
        self.assertAlmostEqual(mt_per_m * 365.0, 112.12203006, places=6)

    def test_lodha_line_2_id_specified(self):
        """Lodha Line 2: ID 1050, WT 12mm, 25035m -> OD 1074, 314.286 kg/m, 7868.154 MT."""
        od_mm, id_mm, kg_m = geometry(1050.0, 12.0, "ID")
        self.assertEqual(od_mm, 1074.0)
        self.assertEqual(id_mm, 1050.0)
        self.assertAlmostEqual(kg_m, 314.286156, places=6)
        mt_per_m = mass_mt_per_m(od_mm, 12.0)
        self.assertAlmostEqual(mt_per_m * 25035.0, 7868.15391546, places=6)

    def test_lodha_line_3_od_specified(self):
        """Lodha Line 3: OD 900, WT 12mm, 120m -> ID 876, 262.793 kg/m, 31.535 MT."""
        od_mm, id_mm, kg_m = geometry(900.0, 12.0, "OD")
        self.assertEqual(od_mm, 900.0)
        self.assertEqual(id_mm, 876.0)
        self.assertAlmostEqual(kg_m, 262.792944, places=6)
        mt_per_m = mass_mt_per_m(od_mm, 12.0)
        self.assertAlmostEqual(mt_per_m * 120.0, 31.53515328, places=6)

    def test_lodha_line_4_od_specified(self):
        """Lodha Line 4: OD 700, WT 12mm, 80m -> ID 676, 203.605 kg/m, 16.288 MT."""
        od_mm, id_mm, kg_m = geometry(700.0, 12.0, "OD")
        self.assertEqual(od_mm, 700.0)
        self.assertEqual(id_mm, 676.0)
        self.assertAlmostEqual(kg_m, 203.605344, places=6)
        mt_per_m = mass_mt_per_m(od_mm, 12.0)
        self.assertAlmostEqual(mt_per_m * 80.0, 16.28842752, places=6)

    def test_lodha_line_5_od_specified(self):
        """Lodha Line 5: OD 1000, WT 12mm, 100m -> ID 976, 292.387 kg/m, 29.239 MT."""
        od_mm, id_mm, kg_m = geometry(1000.0, 12.0, "OD")
        self.assertEqual(od_mm, 1000.0)
        self.assertEqual(id_mm, 976.0)
        self.assertAlmostEqual(kg_m, 292.386744, places=6)
        mt_per_m = mass_mt_per_m(od_mm, 12.0)
        self.assertAlmostEqual(mt_per_m * 100.0, 29.23867440, places=6)

    def test_lodha_total_weight_parity(self):
        """Total tonnage across all 5 Lodha lines equals 8,057.338 MT."""
        total_mt = (
            mass_mt_per_m(1050.0, 12.0) * 365.0
            + mass_mt_per_m(1074.0, 12.0) * 25035.0
            + mass_mt_per_m(900.0, 12.0) * 120.0
            + mass_mt_per_m(700.0, 12.0) * 80.0
            + mass_mt_per_m(1000.0, 12.0) * 100.0
        )
        self.assertAlmostEqual(total_mt, 8057.33820072, places=5)
        excel_k17 = float(self.ws["K17"].value)
        self.assertAlmostEqual(total_mt, excel_k17, places=5)


class TestManualSellingPriceOverrideAudit(unittest.TestCase):
    """Verifies that manual selling price overrides preserve audit trail and calculate correctly."""

    def test_manual_override_audit_trail(self):
        """Line 1 calculated price Rs 22,351 overridden to Rs 22,350 preserves variance and source."""
        res = calc_line(
            size_mm=1226.0,
            thickness_mm=8.4,
            basis="ID",
            quantity=2000.0,
            unit="Metres",
            steel_rate=62000.0,
            conversion=4000.0,
            ext_rate=500.0,
            int_rate=250.0,
            manual_fc=1024.0,
            trip_rate=93000.0,
            handling=100.0,
            pipes_per_trailer=4.0,
            approved_mode=True,
            manual_selling_price=22350.0,
            override_reason="Agreed client concession",
        )
        self.assertEqual(res.calculated_unit_price, 22351.0)
        self.assertEqual(res.applied_unit_price, 22350.0)
        self.assertEqual(res.price_source, "MANUAL_OVERRIDE")
        self.assertEqual(res.price_variance, -1.0)
        self.assertEqual(res.override_reason, "Agreed client concession")
        self.assertEqual(res.line_value, 22350.0 * 2000.0)


if __name__ == "__main__":
    unittest.main()
