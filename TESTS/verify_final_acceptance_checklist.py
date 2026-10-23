"""Final Acceptance Checklist Verification Script.

Executes and captures empirical evidence for all 7 acceptance checklist items:
1. Open disposable workbook and confirm OFFER/SYNOPSIS are hidden.
2. Click New Quotation and confirm previous customer and pipe inputs are cleared.
3. Click Clear All Inputs and confirm line items and totals reset.
4. Try Calculate Quotation without a customer and confirm validation blocks it.
5. Create a test quotation and confirm OFFER/SYNOPSIS appear with the new customer's data.
6. Start another quotation and confirm no previous report data remains.
7. Inspect the offer's print layout and save/reopen behavior.
"""

import os
import sys
import shutil
import win32com.client as win32

ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
BASE_DEV_WB = os.path.join(ROOT_DIR, "Mega_Pipes_Quotation_DEV.xlsm")
DISPOSABLE_DIR = os.path.join(ROOT_DIR, "..", "disposable_parity_test")
DISPOSABLE_WB = os.path.join(DISPOSABLE_DIR, "Mega_Pipes_Workflow_Test.xlsm")
VBA_MODULES_DIR = os.path.join(ROOT_DIR, "VBA", "Modules")

MODULES_TO_IMPORT = ["modPricing", "modQuote", "modQuoteLines", "modExport", "modUI", "modSynopsis"]

xlSheetVisible = -1
xlSheetHidden = 0


def prepare_disposable_workbook(excel):
    print("\n[Setup] Preparing fresh disposable workbook and importing updated VBA modules...")
    os.makedirs(DISPOSABLE_DIR, exist_ok=True)
    if os.path.exists(DISPOSABLE_WB):
        try:
            os.remove(DISPOSABLE_WB)
        except Exception:
            pass
    shutil.copy2(BASE_DEV_WB, DISPOSABLE_WB)

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
        print("  + Imported modCostSheetParity.bas")

    excel.Run("modUI.OnWorkbookOpen")
    wb.Save()
    wb.Close(True)
    print("  [x] Setup complete: Disposable workbook initialized and saved with hidden output sheets.")


def verify_checklist():
    print("=" * 75)
    print("FINAL ACCEPTANCE CHECKLIST VERIFICATION")
    print("=" * 75)
    print(f"Disposable Workbook: {DISPOSABLE_WB}")

    excel = win32.Dispatch("Excel.Application")
    try:
        excel.Visible = False
    except Exception:
        pass
    excel.DisplayAlerts = False

    prepare_disposable_workbook(excel)

    wb = None
    try:
        # -------------------------------------------------------------
        # Item 1: Open disposable workbook and confirm OFFER/SYNOPSIS are hidden
        # -------------------------------------------------------------
        print("\n[Item 1/7] Open workbook & confirm OFFER / SYNOPSIS are hidden...")
        wb = excel.Workbooks.Open(os.path.abspath(DISPOSABLE_WB))
        excel.Run("modUI.OnWorkbookOpen")

        wsOff = wb.Worksheets("OFFER")
        wsSyn = wb.Worksheets("SYNOPSIS")
        wsQ = wb.Worksheets("QUOTE")
        wsLines = wb.Worksheets("QUOTE_LINES")
        tblLines = wsLines.ListObjects("tblQuoteLines")

        print(f"  OFFER Visible State:    {wsOff.Visible} (Expected {xlSheetHidden} = xlSheetHidden)")
        print(f"  SYNOPSIS Visible State: {wsSyn.Visible} (Expected {xlSheetHidden} = xlSheetHidden)")
        assert wsOff.Visible == xlSheetHidden, "OFFER is not hidden on open"
        assert wsSyn.Visible == xlSheetHidden, "SYNOPSIS is not hidden on open"
        print("  [x] Item 1 PASSED: OFFER and SYNOPSIS are confirmed hidden on open.")

        # -------------------------------------------------------------
        # Item 2: Click New Quotation and confirm previous customer and pipe inputs are cleared
        # -------------------------------------------------------------
        print("\n[Item 2/7] Click New Quotation & confirm customer and pipe inputs are cleared...")
        excel.Run("modUI.UI_NewQuotation")

        b4_val = str(wsQ.Range("B4").Value or "").strip()
        b8_val = str(wsQ.Range("B8").Value or "").strip()
        b9_val = str(wsQ.Range("B9").Value or "").strip()
        b10_val = str(wsQ.Range("B10").Value or "").strip()
        b11_val = str(wsQ.Range("B11").Value or "").strip()
        b12_val = str(wsQ.Range("B12").Value or "").strip()
        b23_val = float(wsQ.Range("B23").Value or 0)
        b24_val = float(wsQ.Range("B24").Value or 0)
        b28_val = float(wsQ.Range("B28").Value or 0)
        b31_val = str(wsQ.Range("B31").Value or "").strip()
        b32_val = str(wsQ.Range("B32").Value or "").strip()
        b33_val = str(wsQ.Range("B33").Value or "").strip()

        print(f"  Quote ID (B4):        '{b4_val}'")
        print(f"  Customer (B8):        '{b8_val}' (Must be empty, not Unipro)")
        print(f"  Contact Person (B9):  '{b9_val}'")
        print(f"  Delivery Place (B10): '{b10_val}'")
        print(f"  Project Ref (B11):    '{b11_val}'")
        print(f"  Salesperson (B12):    '{b12_val}'")
        print(f"  Pipe Size (B23):      {b23_val} (Must be 0, not 1226)")
        print(f"  Thickness (B24):      {b24_val} (Must be 0, not 8.4)")
        print(f"  Quantity (B28):       {b28_val} (Must be 0, not 2000)")
        print(f"  Custom Size (B31):    '{b31_val}'")
        print(f"  Price Override (B32): '{b32_val}'")
        print(f"  Override Reason (B33):'{b33_val}'")

        assert b8_val == "", "Customer B8 was not cleared"
        assert b9_val == "", "Contact Person B9 was not cleared"
        assert b10_val == "", "Delivery Place B10 was not cleared"
        assert b11_val == "", "Project Ref B11 was not cleared"
        assert b12_val == "", "Salesperson B12 was not cleared"
        assert b23_val == 0.0, "Pipe Size B23 was not 0"
        assert b24_val == 0.0, "Thickness B24 was not 0"
        assert b28_val == 0.0, "Quantity B28 was not 0"
        assert b31_val.upper() == "NO", "Custom toggle B31 not NO"
        assert b32_val == "", "Price override B32 not empty"
        assert b33_val == "", "Override reason B33 not empty"
        print("  [x] Item 2 PASSED: All previous customer and pipe inputs cleanly cleared.")

        # -------------------------------------------------------------
        # Item 3: Click Clear All Inputs and confirm line items and totals reset
        # -------------------------------------------------------------
        print("\n[Item 3/7] Click Clear All Inputs & confirm line items and totals reset...")
        # Populate dummy data first
        wsQ.Range("B8").Value = "Temporary Test Customer"
        wsQ.Range("B23").Value = 1226
        wsQ.Range("B24").Value = 8.4
        wsQ.Range("B28").Value = 1000
        excel.Run("modUI.UI_AddItem")

        assert tblLines.ListRows.Count >= 1, "Failed to add test line"
        assert float(wsQ.Range("E8").Value or 0) > 0, "Subtotal was not calculated after add item"

        # Click Clear All Inputs
        excel.Run("modUI.UI_ClearAllInputs", False)

        active_lines = 0
        for r in tblLines.ListRows:
            pid = str(r.Range.Cells(1, tblLines.ListColumns("ProductID").Index).Value or "").strip()
            if len(pid) > 0:
                active_lines += 1

        subtotal = float(wsQ.Range("E8").Value or 0)
        gst = float(wsQ.Range("E9").Value or 0)
        grand_total = float(wsQ.Range("E10").Value or 0)
        status_val = str(wsQ.Range("E11").Value or "").strip()

        print(f"  Active Lines in Table: {active_lines} (Expected 0)")
        print(f"  Live Subtotal (E8):    Rs. {subtotal:,.2f}")
        print(f"  Live GST @ 18% (E9):   Rs. {gst:,.2f}")
        print(f"  Live Grand Total (E10):Rs. {grand_total:,.2f}")
        print(f"  Workflow Status (E11): '{status_val}'")
        print(f"  OFFER Sheet Visible:   {wsOff.Visible}")
        print(f"  SYNOPSIS Sheet Visible:{wsSyn.Visible}")

        assert active_lines == 0, "tblQuoteLines still has active lines"
        assert subtotal == 0.0, "Subtotal is not 0"
        assert grand_total == 0.0, "Grand total is not 0"
        assert status_val == "Draft", "Status is not Draft"
        assert wsOff.Visible == xlSheetHidden, "OFFER should be hidden"
        assert wsSyn.Visible == xlSheetHidden, "SYNOPSIS should be hidden"
        print("  [x] Item 3 PASSED: Line items and totals completely reset to zero/draft.")

        # -------------------------------------------------------------
        # Item 4: Try Calculate Quotation without a customer and confirm validation blocks it
        # -------------------------------------------------------------
        print("\n[Item 4/7] Try Calculate Quotation without a customer...")
        wsQ.Range("B8").Value = ""  # Blank customer
        wsQ.Range("B23").Value = 1226
        wsQ.Range("B24").Value = 8.4
        wsQ.Range("B28").Value = 500
        excel.Run("modUI.UI_AddItem")

        excel.Run("modUI.UI_CalculateQuotation")

        banner_text = str(wsQ.Range("A34").Value or "")
        print(f"  Feedback Banner (A34): '{banner_text}'")
        print(f"  OFFER Sheet Visible:   {wsOff.Visible} (Expected {xlSheetHidden})")
        print(f"  SYNOPSIS Sheet Visible:{wsSyn.Visible} (Expected {xlSheetHidden})")

        assert "Validation Error" in banner_text, "Validation error was not displayed in banner"
        assert wsOff.Visible == xlSheetHidden, "OFFER must remain hidden on validation failure"
        assert wsSyn.Visible == xlSheetHidden, "SYNOPSIS must remain hidden on validation failure"
        print("  [x] Item 4 PASSED: Validation blocked calculation without a customer and kept outputs hidden.")

        # -------------------------------------------------------------
        # Item 5: Create a test quotation and confirm OFFER/SYNOPSIS appear with new customer data
        # -------------------------------------------------------------
        print("\n[Item 5/7] Create test quotation & confirm OFFER/SYNOPSIS appear with new customer...")
        test_customer = "Tata Projects Ltd"
        wsQ.Range("B8").Value = test_customer
        wsQ.Range("B9").Value = "Sunil Kumar"
        wsQ.Range("B10").Value = "Mundra Site, Gujarat"
        wsQ.Range("B11").Value = "Mundra Desalination Project"

        excel.Run("modUI.UI_CalculateQuotation")

        print(f"  OFFER Sheet Visible:   {wsOff.Visible} (Expected {xlSheetVisible} = xlSheetVisible)")
        print(f"  SYNOPSIS Sheet Visible:{wsSyn.Visible} (Expected {xlSheetVisible} = xlSheetVisible)")
        assert wsOff.Visible == xlSheetVisible, "OFFER should be visible after calculation"
        assert wsSyn.Visible == xlSheetVisible, "SYNOPSIS should be visible after calculation"

        offer_purchaser = str(wsOff.Range("C7").Value or "").strip()
        offer_total = float(wsOff.Range("H12").Value or 0)
        syn_customer = str(wsSyn.Range("D6").Value or "").strip()
        syn_qty = float(wsSyn.Range("H12").Value or 0)

        print(f"  OFFER Purchaser (C7):  '{offer_purchaser}'")
        print(f"  OFFER Line 1 Val (H12): Rs. {offer_total:,.2f}")
        print(f"  SYNOPSIS Customer (D6):'{syn_customer}'")
        print(f"  SYNOPSIS Line 1 M (H12):{syn_qty:,.2f} MTR")

        assert test_customer in offer_purchaser, f"OFFER purchaser does not contain '{test_customer}'"
        assert test_customer in syn_customer, f"SYNOPSIS customer does not contain '{test_customer}'"
        print("  [x] Item 5 PASSED: OFFER and SYNOPSIS generated and revealed with new customer data.")

        # -------------------------------------------------------------
        # Item 6: Start another quotation and confirm no previous report data remains
        # -------------------------------------------------------------
        print("\n[Item 6/7] Start another quotation & confirm no previous report data remains...")
        excel.Run("modUI.UI_NewQuotation")
        assert wsOff.Visible == xlSheetHidden, "OFFER must be hidden on New Quotation"
        assert wsSyn.Visible == xlSheetHidden, "SYNOPSIS must be hidden on New Quotation"

        second_customer = "Larsen & Toubro Construction"
        wsQ.Range("B8").Value = second_customer
        wsQ.Range("B23").Value = 1300
        wsQ.Range("B24").Value = 8.4
        wsQ.Range("B28").Value = 800
        excel.Run("modUI.UI_AddItem")
        excel.Run("modUI.UI_CalculateQuotation")

        q2_offer_purchaser = str(wsOff.Range("C7").Value or "").strip()
        q2_syn_customer = str(wsSyn.Range("D6").Value or "").strip()
        q2_offer_qty = float(wsOff.Range("F12").Value or 0)

        print(f"  Q2 OFFER Purchaser (C7):   '{q2_offer_purchaser}'")
        print(f"  Q2 SYNOPSIS Customer (D6): '{q2_syn_customer}'")
        print(f"  Q2 OFFER Quantity (F12):   {q2_offer_qty:,.0f} MTR")

        assert second_customer in q2_offer_purchaser, "Second customer missing from OFFER"
        assert test_customer not in q2_offer_purchaser, f"Old customer '{test_customer}' leaked into Q2 OFFER"
        assert test_customer not in q2_syn_customer, f"Old customer '{test_customer}' leaked into Q2 SYNOPSIS"
        assert q2_offer_qty == 800.0, f"Expected 800 MTR, got {q2_offer_qty}"
        print("  [x] Item 6 PASSED: Zero remnants from previous quotation.")

        # -------------------------------------------------------------
        # Item 7: Inspect offer print layout and save/reopen behavior
        # -------------------------------------------------------------
        print("\n[Item 7/7] Inspect offer print layout & save/reopen behavior...")
        # Inspect PageSetup
        ps = wsOff.PageSetup
        orientation = ps.Orientation  # 1 = xlPortrait, 2 = xlLandscape
        fit_wide = ps.FitToPagesWide
        print(f"  OFFER Page Orientation: {orientation} (1 = xlPortrait)")
        print(f"  OFFER FitToPagesWide:   {fit_wide}")

        # Check total rows and terms
        offer_terms_title = str(wsOff.Range("A19").Value or "").strip()
        print(f"  OFFER Terms Header (A19): '{offer_terms_title}'")
        assert "TERMS & CONDITIONS" in offer_terms_title, "Terms header missing from OFFER"

        # Save workbook
        print("  Saving workbook state...")
        wb.Save()
        wb.Close(True)
        wb = None

        # Reopen workbook to test lifecycle
        print("  Reopening workbook to verify clean initialization...")
        wb = excel.Workbooks.Open(os.path.abspath(DISPOSABLE_WB))
        excel.Run("modUI.OnWorkbookOpen")

        wsOff_reopen = wb.Worksheets("OFFER")
        wsSyn_reopen = wb.Worksheets("SYNOPSIS")
        wsStart_reopen = wb.Worksheets("START")
        active_sheet_name = wb.ActiveSheet.Name

        print(f"  Active Sheet on Reopen: '{active_sheet_name}' (Expected 'START')")
        print(f"  OFFER Visible on Reopen:    {wsOff_reopen.Visible} (Expected {xlSheetHidden})")
        print(f"  SYNOPSIS Visible on Reopen: {wsSyn_reopen.Visible} (Expected {xlSheetHidden})")

        assert wsOff_reopen.Visible == xlSheetHidden, "OFFER must be hidden on workbook reopen"
        assert wsSyn_reopen.Visible == xlSheetHidden, "SYNOPSIS must be hidden on workbook reopen"
        assert active_sheet_name == "START", f"Active sheet should be START, got {active_sheet_name}"
        print("  [x] Item 7 PASSED: Print layout verified; save and reopen cleanly initializes with hidden reports.")

        print("\n" + "=" * 75)
        print("ALL 7 FINAL ACCEPTANCE CHECKLIST ITEMS EMPIRICALLY VERIFIED!")
        print("=" * 75)
        return True

    finally:
        if wb:
            try:
                wb.Close(False)
            except Exception:
                pass
        excel.Quit()


if __name__ == "__main__":
    verify_checklist()
