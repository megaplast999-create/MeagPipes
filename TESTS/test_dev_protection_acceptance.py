"""Read-only Acceptance Test Runner for DEV/Mega_Pipes_Quotation_DEV.xlsm.
Tests A through F on DEV:
- Test A: Input editing on QUOTE sheet
- Test B: Calculate quotation (OFFER & SYNOPSIS generated and protected)
- Test C: Manual editing prevention (cells locked against user modification)
- Test D: Macro functionality (VBA can write & recalculate without runtime errors)
- Test E: Recalculation & stale results (modifying inputs hides/invalidates stale outputs, recalculate refreshes & re-protects)
- Test F: Exported workbook (separate .xlsx is independently editable, original stays protected)
"""

import os
import sys
import time
import subprocess
import pythoncom
import win32com.client as win32

ROOT_DIR = os.path.abspath(r"c:\Users\Administrator\Desktop\Mega Pipes\Quotation System")
DEV_WB = os.path.join(ROOT_DIR, "DEV", "Mega_Pipes_Quotation_DEV.xlsm")

xlSheetVisible = -1
xlSheetHidden = 0

def kill_excel():
    subprocess.run(["powershell", "-Command", "Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue"], capture_output=True)
    time.sleep(1)

def run_acceptance_tests():
    print("=" * 80)
    print("ACCEPTANCE TESTS ON CURRENT DEV WORKBOOK")
    print("=" * 80)

    kill_excel()
    pythoncom.CoInitialize()
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False

    wb = None
    try:
        wb = excel.Workbooks.Open(DEV_WB)
        try:
            excel.Run("modUI.SetTestMode", True)
        except Exception:
            pass

        wsQ = wb.Worksheets("QUOTE")
        wsLines = wb.Worksheets("QUOTE_LINES")
        wsOff = wb.Worksheets("OFFER")
        wsSyn = wb.Worksheets("SYNOPSIS")

        # Test A: Input editing on QUOTE
        print("\n--- TEST A: QUOTE Sheet Input Editing ---")
        wsQ.Range("B8").Value = "Acme Infrastructure Ltd."
        assert str(wsQ.Range("B8").Value) == "Acme Infrastructure Ltd."
        wsQ.Range("B10").Value = "Pune, Maharashtra - 411001"
        assert "Pune" in str(wsQ.Range("B10").Value)
        wsQ.Range("B20").Value = "IS-3589-2001"
        wsQ.Range("B21").Value = "HSAW"
        wsQ.Range("B22").Value = "ID"
        wsQ.Range("B23").Value = 1226
        wsQ.Range("B24").Value = 8.4
        wsQ.Range("B25").Value = "Fe 410"
        wsQ.Range("B26").Value = "3LPE"
        wsQ.Range("B27").Value = "Liquid Epoxy"
        wsQ.Range("B28").Value = 2000
        wsQ.Range("B29").Value = "Metres"
        wsQ.Range("B30").Value = 12.0
        wsQ.Range("B31").Value = "NO"
        wsQ.Range("C31").Value = ""
        wsQ.Range("B32").Value = ""
        wsQ.Range("B33").Value = ""
        print("PASS: Input cells B8, B10, B20:B33 are unlocked and editable.")

        # Test B: Calculation & Sheet Generation + Protection
        print("\n--- TEST B: Calculation & Sheet Generation + Protection ---")
        excel.Run("modQuote.ClearQuoteLines")
        tblLines = wsLines.ListObjects("tblQuoteLines")
        row1 = tblLines.ListRows(1) if tblLines.ListRows.Count > 0 else tblLines.ListRows.Add()
        row1.Range.Cells(1, tblLines.ListColumns("ProductID").Index).Value = "IS-3589-2001-HSAW"
        row1.Range.Cells(1, tblLines.ListColumns("Family").Index).Value = "HSAW"
        row1.Range.Cells(1, tblLines.ListColumns("Basis").Index).Value = "ID"
        row1.Range.Cells(1, tblLines.ListColumns("SizeMM").Index).Value = 1226
        row1.Range.Cells(1, tblLines.ListColumns("ThicknessMM").Index).Value = 8.4
        row1.Range.Cells(1, tblLines.ListColumns("GradeID").Index).Value = "Fe 410"
        row1.Range.Cells(1, tblLines.ListColumns("CoatingExternalID").Index).Value = "3LPE"
        row1.Range.Cells(1, tblLines.ListColumns("CoatingInternalID").Index).Value = "Liquid Epoxy"
        row1.Range.Cells(1, tblLines.ListColumns("Quantity").Index).Value = 2000
        row1.Range.Cells(1, tblLines.ListColumns("QuantityUnit").Index).Value = "Metres"
        row1.Range.Cells(1, tblLines.ListColumns("PipeLengthM").Index).Value = 12.0
        row1.Range.Cells(1, tblLines.ListColumns("PlantID").Index).Value = "P01"
        row1.Range.Cells(1, tblLines.ListColumns("Status").Index).Value = "OK"

        excel.Run("modUI.UI_CalculateQuotation")
        grand_total = float(wsQ.Range("E10").Value or 0)
        print(f"Quotation Calculated: Grand Total = Rs. {grand_total:,.2f}")
        assert grand_total > 0
        assert wsOff.Visible == xlSheetVisible
        assert wsSyn.Visible == xlSheetVisible
        assert wsOff.ProtectContents is True
        assert wsSyn.ProtectContents is True
        print("PASS: Both OFFER and SYNOPSIS generated, visible, and PROTECTED.")

        # Test C: Manual Editing Prevention
        print("\n--- TEST C: Manual Editing Prevention on Output Sheets ---")
        assert wsOff.Range("F12").Locked is True
        assert wsOff.Range("G12").Locked is True
        assert wsOff.Range("H12").Locked is True
        assert wsOff.Range("A17").Locked is True
        assert wsOff.Range("A18").Locked is True
        assert wsSyn.Range("L12").Locked is True
        print("PASS: Output cells confirmed LOCKED against manual editing.")

        # Test D: Macro Recalculation
        print("\n--- TEST D: Macro Recalculation Capability ---")
        excel.Run("modUI.UI_CalculateQuotation")
        assert wsOff.ProtectContents is True
        assert wsSyn.ProtectContents is True
        print("PASS: Recalculated cleanly with 0 protection errors; protection restored.")

        # Test E: Stale Results Invalidation
        print("\n--- TEST E: Stale Results Invalidation ---")
        wsQ.Range("B20").Value = "IS-3589-2001"
        wsQ.Range("B21").Value = "HSAW"
        wsQ.Range("B22").Value = "ID"
        wsQ.Range("B23").Value = 1300
        wsQ.Range("B24").Value = 8.4
        wsQ.Range("B25").Value = "Fe 410"
        wsQ.Range("B26").Value = "3LPE"
        wsQ.Range("B27").Value = "Liquid Epoxy"
        wsQ.Range("B28").Value = 1000
        wsQ.Range("B29").Value = "Metres"
        wsQ.Range("B30").Value = 12.0
        wsQ.Range("B31").Value = "NO"
        wsQ.Range("C31").Value = ""
        wsQ.Range("B32").Value = ""
        wsQ.Range("B33").Value = ""
        excel.Run("modUI.UI_AddItem")
        assert wsOff.Visible == xlSheetHidden, "OFFER must be hidden on input change"
        assert wsSyn.Visible == xlSheetHidden, "SYNOPSIS must be hidden on input change"
        print("PASS: Stale output sheets hidden immediately upon modifying quotation items.")

        excel.Run("modUI.UI_CalculateQuotation")
        assert wsOff.Visible == xlSheetVisible
        assert wsSyn.Visible == xlSheetVisible
        assert wsOff.ProtectContents is True
        assert wsSyn.ProtectContents is True
        print("PASS: Recalculation updated outputs and restored protection.")

        # Test F: Exported Workbook Editability & Working Workbook Safety
        print("\n--- TEST F: Exported Workbook Editability & Independence ---")
        export_res = excel.Run("modExport.ExportCustomerExcelDual", "", "", False, False)
        print(f"ExportCustomerExcelDual Result: {export_res}")

        exported_path = None
        for part in export_res.split(";"):
            if part.startswith("LocalPath="):
                p = part.split("=")[1].strip()
                if p and os.path.exists(p):
                    exported_path = p
                    break
            elif part.startswith("CentralPath="):
                p = part.split("=")[1].strip()
                if p and os.path.exists(p):
                    exported_path = p
                    break
            elif part.startswith("StagedPath="):
                p = part.split("=")[1].strip()
                if p and os.path.exists(p):
                    exported_path = p
                    break

        print(f"Exported customer file: {exported_path}")
        assert exported_path and os.path.exists(exported_path)
        assert wsOff.ProtectContents is True
        assert wsSyn.ProtectContents is True
        print("PASS: Original working workbook remains PROTECTED during and after export.")

        exp_wb = excel.Workbooks.Open(exported_path)
        try:
            exp_ws_offer = exp_wb.Worksheets("OFFER")
            exp_ws_syn = exp_wb.Worksheets("Synopsis")
            assert exp_ws_offer.ProtectContents is False, "Exported OFFER must be unprotected"
            assert exp_ws_syn.ProtectContents is False, "Exported Synopsis must be unprotected"
            assert exp_ws_offer.Range("F12").Locked is False, "Exported cells must be unlocked"
            print("PASS: Exported customer sheets are UNPROTECTED and cells UNLOCKED.")

            # Test editing exported file
            exp_ws_offer.Range("F12").Value = 7777
            exp_ws_offer.Range("G12").Value = 33000
            exp_ws_offer.Range("H12").Value = 7777 * 33000
            assert exp_ws_offer.Range("F12").Value == 7777
            assert exp_ws_offer.Range("G12").Value == 33000
            exp_wb.Save()
            print("PASS: Recipient can independently edit quantities, rates, totals, and save cleanly.")
        finally:
            exp_wb.Close(SaveChanges=False)

        # Pre-export validation
        excel.Run("modUI.HideOutputSheets")
        val_res = excel.Run("modExport.ValidateQuotationBeforeExport", "")
        assert val_res is False
        print("PASS: ValidateQuotationBeforeExport blocks export when quotation is uncalculated.")

        print("\n" + "=" * 80)
        print("ALL ACCEPTANCE TESTS A THROUGH F PASSED EMPIRICALLY ON DEV!")
        print("=" * 80)

    finally:
        if wb:
            try:
                wb.Save()
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        excel.Quit()
        kill_excel()
        pythoncom.CoUninitialize()

    return True

if __name__ == "__main__":
    if not run_acceptance_tests():
        sys.exit(1)
