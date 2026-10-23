"""Automated Regression Test Suite for Workflow Reset and Sheet Visibility.

Covers Scenarios A through G:
- Test A: Clear All Inputs (clears customer, project, line inputs, overrides, lines, outputs; preserves master data, settings, history)
- Test B: New Quotation (resets inputs cleanly without hardcoded Unipro data, hides OFFER/SYNOPSIS, preserves numbering rules)
- Test C: Calculate Quotation (keeps OFFER/SYNOPSIS hidden before calc, validates, calculates, reveals outputs with current data)
- Test D: Validation Failure (empty customer or 0 lines keeps outputs hidden and displays validation error)
- Test E: Preview Before Calculation Blocked (attempting to preview OFFER or SYNOPSIS before calculation is politely blocked)
- Test F: Recalculate Existing Quotation (preserves valid inputs and overrides)
- Test G: Repeatability (subsequent quotations do not retain remnants of previous quotations in reports)
"""

import os
import sys
import shutil
import time
import win32com.client as win32

# Path definitions
ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
BASE_DEV_WB = os.path.join(ROOT_DIR, "Mega_Pipes_Quotation_DEV.xlsm")
DISPOSABLE_DIR = os.path.join(ROOT_DIR, "..", "disposable_parity_test")
DISPOSABLE_WB = os.path.join(DISPOSABLE_DIR, "Mega_Pipes_Workflow_Test.xlsm")
VBA_MODULES_DIR = os.path.join(ROOT_DIR, "VBA", "Modules")

MODULES_TO_IMPORT = ["modPricing", "modQuote", "modQuoteLines", "modExport", "modUI", "modSynopsis"]

xlSheetVisible = -1
xlSheetHidden = 0
xlSheetVeryHidden = 2


def run_all_scenarios():
    print("=" * 75)
    print("MEGA PIPES WORKFLOW & VISIBILITY REGRESSION TEST SUITE (SCENARIOS A - G)")
    print("=" * 75)
    print(f"Base Workbook:       {BASE_DEV_WB}")
    print(f"Disposable Workbook: {DISPOSABLE_WB}")

    # 1. Prepare disposable copy
    os.makedirs(DISPOSABLE_DIR, exist_ok=True)
    if os.path.exists(DISPOSABLE_WB):
        try:
            os.remove(DISPOSABLE_WB)
        except Exception:
            pass
    shutil.copy2(BASE_DEV_WB, DISPOSABLE_WB)

    excel = win32.Dispatch("Excel.Application")
    try:
        excel.Visible = False
    except Exception:
        pass
    excel.DisplayAlerts = False

    wb = None
    try:
        print("\n[Setup] Opening disposable workbook and importing updated VBA modules...")
        wb = excel.Workbooks.Open(os.path.abspath(DISPOSABLE_WB))
        vbp = wb.VBProject

        for mod_name in MODULES_TO_IMPORT:
            try:
                comp = vbp.VBComponents(mod_name)
                vbp.VBComponents.Remove(comp)
            except Exception:
                pass
            bas_path = os.path.join(VBA_MODULES_DIR, f"{mod_name}.bas")
            if os.path.exists(bas_path):
                vbp.VBComponents.Import(bas_path)
                print(f"  + Imported {mod_name}.bas")

        parity_path = os.path.join(VBA_MODULES_DIR, "modCostSheetParity.bas")
        if os.path.exists(parity_path):
            try:
                comp = vbp.VBComponents("modCostSheetParity")
                vbp.VBComponents.Remove(comp)
            except Exception:
                pass
            vbp.VBComponents.Import(parity_path)

        # Trigger initial sheet setup
        excel.Run("modUI.SetupQuoteSheetUI")

        wsQ = wb.Worksheets("QUOTE")
        wsLines = wb.Worksheets("QUOTE_LINES")
        wsOff = wb.Worksheets("OFFER")
        wsSyn = wb.Worksheets("SYNOPSIS")
        wsLog = wb.Worksheets("QUOTE_LOG")
        tblLog = wsLog.ListObjects("tblQuoteLog")
        tblLines = wsLines.ListObjects("tblQuoteLines")

        # -------------------------------------------------------------
        # TEST A — Clear All Inputs
        # -------------------------------------------------------------
        print("\n[Running Test A] Clear All Inputs...")
        initial_log_count = tblLog.ListRows.Count

        # Populate user inputs and line
        wsQ.Range("B8").Value = "Custom Corp Ltd"
        wsQ.Range("B9").Value = "Jane Doe"
        wsQ.Range("B10").Value = "Visakhapatnam Port"
        wsQ.Range("B11").Value = "Project Pipeline 2026"
        wsQ.Range("B12").Value = "Sales Exec 1"
        wsQ.Range("B13").Value = "P02"
        wsQ.Range("B23").Value = 1226
        wsQ.Range("B24").Value = 8.4
        wsQ.Range("B28").Value = 1500
        wsQ.Range("B31").Value = "YES"
        wsQ.Range("C31").Value = "1226"
        wsQ.Range("B32").Value = 24500
        wsQ.Range("B33").Value = "Client special rate"

        excel.Run("modUI.UI_AddItem")
        assert tblLines.ListRows.Count >= 1, "Failed to add quotation line for Test A"

        # Execute Clear All Inputs routine
        excel.Run("modUI.UI_ClearAllInputs", False)

        # Verify intended inputs are cleared
        assert str(wsQ.Range("B8").Value or "").strip() == "", "Customer B8 was not cleared"
        assert str(wsQ.Range("B9").Value or "").strip() == "", "Contact Person B9 was not cleared"
        assert str(wsQ.Range("B10").Value or "").strip() == "", "Delivery Place B10 was not cleared"
        assert str(wsQ.Range("B11").Value or "").strip() == "", "Project Ref B11 was not cleared"
        assert str(wsQ.Range("B12").Value or "").strip() == "", "Salesperson B12 was not cleared"
        assert float(wsQ.Range("B23").Value or 0) == 0.0, "Form size B23 was not cleared"
        assert float(wsQ.Range("B24").Value or 0) == 0.0, "Form thickness B24 was not cleared"
        assert float(wsQ.Range("B28").Value or 0) == 0.0, "Form quantity B28 was not cleared"
        assert str(wsQ.Range("B31").Value or "").strip().upper() == "NO", "Custom toggle B31 not reset"
        assert str(wsQ.Range("B32").Value or "").strip() == "", "Price override B32 not cleared"
        assert str(wsQ.Range("B33").Value or "").strip() == "", "Override reason B33 not cleared"

        active_lines = 0
        for r in tblLines.ListRows:
            pid = str(r.Range.Cells(1, tblLines.ListColumns("ProductID").Index).Value or "").strip()
            if len(pid) > 0:
                active_lines += 1
        assert active_lines == 0, f"Expected 0 active lines after clearing, found {active_lines}"

        assert float(wsQ.Range("E8").Value or 0) == 0.0, "Subtotal E8 was not 0 after clear"
        assert float(wsQ.Range("E10").Value or 0) == 0.0, "Grand Total E10 was not 0 after clear"
        assert str(wsQ.Range("E11").Value or "").strip() == "Draft", "Status E11 was not 'Draft'"
        assert wsOff.Visible == xlSheetHidden, "OFFER sheet was not hidden after Clear All Inputs"
        assert wsSyn.Visible == xlSheetHidden, "SYNOPSIS sheet was not hidden after Clear All Inputs"
        assert wb.Worksheets("MASTERS").ListObjects("tblProducts").ListRows.Count > 0, "Masters damaged"
        assert tblLog.ListRows.Count == initial_log_count, "Historical logs were altered"
        print("  -> PASS: Test A (All inputs, lines, outputs cleared; master data preserved)")

        # -------------------------------------------------------------
        # TEST B — New Quotation
        # -------------------------------------------------------------
        print("\n[Running Test B] New Quotation Reset...")
        # Populate and calculate first
        wsQ.Range("B8").Value = "Existing Customer"
        wsQ.Range("B23").Value = 1226
        wsQ.Range("B24").Value = 8.4
        wsQ.Range("B28").Value = 500
        excel.Run("modUI.UI_AddItem")
        excel.Run("modUI.UI_CalculateQuotation")

        assert wsOff.Visible == xlSheetVisible, "OFFER should be visible after calculation"
        assert wsSyn.Visible == xlSheetVisible, "SYNOPSIS should be visible after calculation"

        # Trigger New Quotation
        excel.Run("modUI.UI_NewQuotation")

        cust_val = str(wsQ.Range("B8").Value or "").strip()
        assert cust_val == "", f"Expected blank customer on New Quotation, got '{cust_val}'"
        assert float(wsQ.Range("B23").Value or 0) == 0.0, "Size B23 must be 0 on New Quotation"
        assert float(wsQ.Range("B24").Value or 0) == 0.0, "Thickness B24 must be 0 on New Quotation"
        assert float(wsQ.Range("B28").Value or 0) == 0.0, "Quantity B28 must be 0 on New Quotation"
        assert str(wsQ.Range("B32").Value or "").strip() == "", "Price override B32 must be blank"
        assert str(wsQ.Range("B33").Value or "").strip() == "", "Override reason B33 must be blank"
        assert wsOff.Visible == xlSheetHidden, "OFFER must be hidden on New Quotation"
        assert wsSyn.Visible == xlSheetHidden, "SYNOPSIS must be hidden on New Quotation"

        q_num = str(wsQ.Range("B4").Value or "").strip()
        assert q_num.startswith("MKTG/ENQ/"), f"Quotation number invalid format: {q_num}"
        print(f"  -> PASS: Test B (New quotation allocated '{q_num}', clean inputs, outputs hidden)")

        # -------------------------------------------------------------
        # TEST C — Calculate Quotation Controls OFFER and SYNOPSIS
        # -------------------------------------------------------------
        print("\n[Running Test C] Calculate Quotation Visibility and Generation...")
        assert wsOff.Visible == xlSheetHidden, "OFFER must be hidden before calc"
        assert wsSyn.Visible == xlSheetHidden, "SYNOPSIS must be hidden before calc"

        test_cust = "Apex Engineering Works"
        wsQ.Range("B8").Value = test_cust
        wsQ.Range("B9").Value = "Vikas Patel"
        wsQ.Range("B10").Value = "Dahej, Gujarat"
        wsQ.Range("B11").Value = "Dahej Water Pipeline"
        wsQ.Range("B23").Value = 1226
        wsQ.Range("B24").Value = 8.4
        wsQ.Range("B28").Value = 1000

        excel.Run("modUI.UI_AddItem")

        assert wsOff.Visible == xlSheetHidden, "OFFER must remain hidden before calculation button is clicked"
        assert wsSyn.Visible == xlSheetHidden, "SYNOPSIS must remain hidden before calculation button is clicked"

        excel.Run("modUI.UI_CalculateQuotation")

        assert wsOff.Visible == xlSheetVisible, "OFFER should be visible after successful calculation"
        assert wsSyn.Visible == xlSheetVisible, "SYNOPSIS should be visible after successful calculation"

        offer_purchaser = str(wsOff.Range("C7").Value or "").strip()
        assert test_cust in offer_purchaser, f"OFFER purchaser expected '{test_cust}', got '{offer_purchaser}'"

        syn_cust = str(wsSyn.Range("D6").Value or "").strip()
        assert test_cust in syn_cust, f"SYNOPSIS customer expected '{test_cust}', got '{syn_cust}'"
        print(f"  -> PASS: Test C (OFFER and SYNOPSIS revealed post-calc with customer '{test_cust}')")

        # -------------------------------------------------------------
        # TEST D — Validation Failure Keeps Outputs Hidden
        # -------------------------------------------------------------
        print("\n[Running Test D] Validation Failures...")
        # Empty customer validation
        excel.Run("modUI.UI_NewQuotation")
        wsQ.Range("B8").Value = ""  # Empty customer
        wsQ.Range("B23").Value = 1226
        wsQ.Range("B24").Value = 8.4
        wsQ.Range("B28").Value = 500
        excel.Run("modUI.UI_AddItem")

        excel.Run("modUI.UI_CalculateQuotation")

        assert wsOff.Visible == xlSheetHidden, "OFFER must stay hidden when customer validation fails"
        assert wsSyn.Visible == xlSheetHidden, "SYNOPSIS must stay hidden when customer validation fails"
        banner_text = str(wsQ.Range("A34").Value or "")
        assert "Validation Error" in banner_text, f"Expected validation error banner, got '{banner_text}'"

        # 0 lines validation
        wsQ.Range("B8").Value = "Valid Customer Name"
        excel.Run("modQuote.ClearQuoteLines")

        excel.Run("modUI.UI_CalculateQuotation")
        assert wsOff.Visible == xlSheetHidden, "OFFER must stay hidden when line items validation fails"
        assert wsSyn.Visible == xlSheetHidden, "SYNOPSIS must stay hidden when line items validation fails"
        print("  -> PASS: Test D (Validation failures keep OFFER & SYNOPSIS hidden)")

        # -------------------------------------------------------------
        # TEST E — Preview Before Calculation Blocked
        # -------------------------------------------------------------
        print("\n[Running Test E] Previewing Hidden Outputs Before Calc...")
        excel.Run("modUI.UI_NewQuotation")
        assert wsOff.Visible == xlSheetHidden
        assert wsSyn.Visible == xlSheetHidden

        excel.Run("modUI.UI_PreviewOffer")
        assert wsOff.Visible == xlSheetHidden, "UI_PreviewOffer must not unhide OFFER if not calculated"

        excel.Run("modUI.UI_PreviewSynopsis")
        assert wsSyn.Visible == xlSheetHidden, "UI_PreviewSynopsis must not unhide SYNOPSIS if not calculated"
        print("  -> PASS: Test E (Preview before calculation blocked safely)")

        # -------------------------------------------------------------
        # TEST F — Recalculate Preserves Valid Inputs & Overrides
        # -------------------------------------------------------------
        print("\n[Running Test F] Recalculate Preserves Inputs & Overrides...")
        excel.Run("modUI.UI_NewQuotation")
        wsQ.Range("B8").Value = "Preserve Inputs Ltd"
        wsQ.Range("B9").Value = "Key Contact Person"
        wsQ.Range("B10").Value = "Destination Station"
        wsQ.Range("B23").Value = 1226
        wsQ.Range("B24").Value = 8.4
        wsQ.Range("B28").Value = 1200

        excel.Run("modUI.UI_AddItem")
        excel.Run("modUI.UI_CalculateQuotation")

        qty_before = float(tblLines.ListRows(1).Range.Cells(1, tblLines.ListColumns("Quantity").Index).Value)
        unit_price_before = float(tblLines.ListRows(1).Range.Cells(1, tblLines.ListColumns("UnitPrice").Index).Value)
        total_before = float(wsQ.Range("E10").Value)

        assert qty_before == 1200.0, "Initial quantity was not set"
        assert unit_price_before > 0.0, "Initial unit price was not calculated"
        assert total_before > 0.0, "Grand total was not calculated"

        # Recalculate quotation
        excel.Run("modUI.UI_Recalculate")

        qty_after = float(tblLines.ListRows(1).Range.Cells(1, tblLines.ListColumns("Quantity").Index).Value)
        unit_price_after = float(tblLines.ListRows(1).Range.Cells(1, tblLines.ListColumns("UnitPrice").Index).Value)
        total_after = float(wsQ.Range("E10").Value)

        assert str(wsQ.Range("B8").Value).strip() == "Preserve Inputs Ltd", "Recalculation wiped Customer"
        assert str(wsQ.Range("B9").Value).strip() == "Key Contact Person", "Recalculation wiped Contact Person"
        assert str(wsQ.Range("B10").Value).strip() == "Destination Station", "Recalculation wiped Delivery Place"
        assert qty_after == qty_before, "Recalculation changed or wiped line quantity"
        assert unit_price_after == unit_price_before, "Recalculation changed line unit price"
        assert total_after == total_before, "Recalculation changed quotation grand total"
        print("  -> PASS: Test F (Valid inputs, customer, line quantity 1,200, unit price & grand total preserved)")

        # -------------------------------------------------------------
        # TEST G — Repeatability (No Leaked Data Across Quotes)
        # -------------------------------------------------------------
        print("\n[Running Test G] Repeatability Across Multiple Quotes...")
        # First quote
        excel.Run("modUI.UI_NewQuotation")
        wsQ.Range("B8").Value = "First Client Infra"
        wsQ.Range("B23").Value = 1226
        wsQ.Range("B24").Value = 8.4
        wsQ.Range("B28").Value = 2000
        excel.Run("modUI.UI_AddItem")
        excel.Run("modUI.UI_CalculateQuotation")

        q1_offer_purchaser = str(wsOff.Range("C7").Value or "")
        assert "First Client Infra" in q1_offer_purchaser

        # Second quote
        excel.Run("modUI.UI_NewQuotation")
        assert wsOff.Visible == xlSheetHidden
        assert wsSyn.Visible == xlSheetHidden

        wsQ.Range("B8").Value = "Second Client Energy"
        wsQ.Range("B23").Value = 1300
        wsQ.Range("B24").Value = 8.4
        wsQ.Range("B28").Value = 750
        excel.Run("modUI.UI_AddItem")
        excel.Run("modUI.UI_CalculateQuotation")

        q2_offer_purchaser = str(wsOff.Range("C7").Value or "")
        assert "Second Client Energy" in q2_offer_purchaser
        assert "First Client" not in q2_offer_purchaser, "Quotation 1 purchaser leaked into Quotation 2 OFFER!"

        q2_syn_cust = str(wsSyn.Range("D6").Value or "")
        assert "Second Client Energy" in q2_syn_cust
        assert "First Client" not in q2_syn_cust, "Quotation 1 customer leaked into Quotation 2 SYNOPSIS!"

        q2_offer_qty = float(wsOff.Range("F12").Value or 0)
        assert q2_offer_qty == 750.0, f"Expected Q2 quantity 750, got {q2_offer_qty}"
        print("  -> PASS: Test G (Repeatability: Zero leaked data across sequential quotations)")

        print("\n" + "=" * 75)
        print("ALL TESTS A THROUGH G PASSED EMPIRICALLY IN MICROSOFT EXCEL!")
        print("=" * 75)
        return True

    finally:
        if wb:
            try:
                wb.Close(False)
            except Exception:
                pass
        excel.Quit()


def test_suite():
    assert run_all_scenarios() is True


if __name__ == "__main__":
    run_all_scenarios()
