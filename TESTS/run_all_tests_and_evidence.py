import os
import sys
import shutil
import win32com.client as win32
from pypdf import PdfWriter
import pymupdf

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

def run_tests():
    root = os.path.abspath(".")
    dev_wb_path = os.path.join(root, "DEV", "Mega_Pipes_Quotation_DEV.xlsm")
    test_wb_path = os.path.join(root, "DEV", "Mega_Pipes_Quotation_TEST_COPY.xlsm")
    shutil.copy2(dev_wb_path, test_wb_path)
    print(f"Created disposable test workbook: {test_wb_path}")

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    test_results = {}
    pdf_pages = []

    try:
        wb = excel.Workbooks.Open(test_wb_path)
        ws_quote = wb.Worksheets("QUOTE")
        ws_lines = wb.Worksheets("QUOTE_LINES")
        tbl_lines = ws_lines.ListObjects("tblQuoteLines")

        # ---------------------------------------------------------
        # TEST A: Empty quotation
        # ---------------------------------------------------------
        print("\n--- TEST A: Empty Quotation ---")
        excel.Run("modUI.UI_NewQuotation")
        
        # Check empty state message and line count
        empty_val = str(ws_quote.Range("D21").Value)
        act_lines_tbl = 0
        for r in range(1, tbl_lines.ListRows.Count + 1):
            if tbl_lines.ListRows(r).Range.Cells(1, 3).Value:
                act_lines_tbl += 1
                
        cap_val = str(ws_quote.Range("F14").Value)
        print(f"Empty cell D21: {empty_val}")
        print(f"Active table lines: {act_lines_tbl}")
        print(f"Capacity badge text: {repr(cap_val)}")

        assert "No quotation lines added yet" in empty_val, "Empty message missing in D21"
        assert act_lines_tbl == 0, "Table lines should be 0"
        assert "0 of 10" in cap_val, "Capacity indicator should be 0 of 10"
        
        page1_pdf = os.path.join(root, "TESTS", "acceptance_page1_empty.pdf")
        ws_quote.ExportAsFixedFormat(0, page1_pdf)
        pdf_pages.append(page1_pdf)
        test_results["Test A (Empty Quotation)"] = "PASS"
        print("Test A: PASS")

        # ---------------------------------------------------------
        # TEST B: Add one line
        # ---------------------------------------------------------
        print("\n--- TEST B: Add One Line ---")
        ws_quote.Range("B21").Value = "HSAW"
        ws_quote.Range("B22").Value = "ID"
        ws_quote.Range("B23").Value = 1226
        ws_quote.Range("B24").Value = 8.4
        ws_quote.Range("B25").Value = "Fe 410"
        ws_quote.Range("B26").Value = "PU 1000 micron"
        ws_quote.Range("B27").Value = "Food grade epoxy 406 micron"
        ws_quote.Range("B28").Value = 2000
        ws_quote.Range("B29").Value = "Metres"

        excel.Run("modUI.UI_AddItem")

        line1_text = str(ws_quote.Range("D21").Value)
        tot_text = str(ws_quote.Range("D22").Value)
        qty_input = ws_quote.Range("B28").Value
        print(f"Row 21: {line1_text}")
        print(f"Row 22: {tot_text}")
        print(f"Quantity input B28 after add: {qty_input}")

        assert "L0001: HSAW 1226mm" in line1_text, "Line 1 text mismatch"
        assert "TABLE TOTAL (1 of 10 sizes)" in tot_text, "Dynamic total row should be at Row 22"
        assert qty_input == 0, "Quantity cell B28 should be cleared to 0"
        test_results["Test B (Add One Line)"] = "PASS"
        print("Test B: PASS")

        # ---------------------------------------------------------
        # TEST C: Add multiple sizes (4 sizes)
        # ---------------------------------------------------------
        print("\n--- TEST C: Add Multiple Sizes ---")
        # Add Line 2
        ws_quote.Range("B23").Value = 1300
        ws_quote.Range("B24").Value = 8.4
        ws_quote.Range("B28").Value = 350
        excel.Run("modUI.UI_AddItem")

        # Add Line 3
        ws_quote.Range("B23").Value = 1350
        ws_quote.Range("B24").Value = 8.4
        ws_quote.Range("B28").Value = 500
        excel.Run("modUI.UI_AddItem")

        # Add Line 4
        ws_quote.Range("B23").Value = 1400
        ws_quote.Range("B24").Value = 8.4
        ws_quote.Range("B28").Value = 4400
        excel.Run("modUI.UI_AddItem")

        # Check rows 21 to 24
        l1 = str(ws_quote.Range("D21").Value)
        l2 = str(ws_quote.Range("D22").Value)
        l3 = str(ws_quote.Range("D23").Value)
        l4 = str(ws_quote.Range("D24").Value)
        tot_r25 = str(ws_quote.Range("D25").Value)
        val_r25 = str(ws_quote.Range("I25").Text)

        print(f"Row 21: {l1}")
        print(f"Row 22: {l2}")
        print(f"Row 23: {l3}")
        print(f"Row 24: {l4}")
        print(f"Row 25 (Total): {tot_r25} | Total Val: {val_r25}")

        assert "1226mm" in l1 and "1300mm" in l2 and "1350mm" in l3 and "1400mm" in l4, "Lines 1-4 mismatch"
        assert "TABLE TOTAL (4 of 10 sizes)" in tot_r25, "Dynamic total row should be at Row 25"
        assert "17,88,01,050" in val_r25 or "178,801,050" in val_r25, f"Golden Total mismatch: {val_r25}"

        page2_pdf = os.path.join(root, "TESTS", "acceptance_page2_four_sizes.pdf")
        ws_quote.ExportAsFixedFormat(0, page2_pdf)
        pdf_pages.append(page2_pdf)
        test_results["Test C (Add Multiple Sizes)"] = "PASS"
        print("Test C: PASS")

        # ---------------------------------------------------------
        # TEST D: Full capacity (10 lines) & 11th line block
        # ---------------------------------------------------------
        print("\n--- TEST D: Full Capacity (10 lines) ---")
        extra_sizes = [
            (1050, 8.0, 500),
            (900, 7.5, 400),
            (800, 7.0, 300),
            (700, 6.5, 250),
            (600, 6.0, 200),
            (500, 5.5, 150)
        ]
        for sz, wt, qty in extra_sizes:
            ws_quote.Range("B23").Value = sz
            ws_quote.Range("B24").Value = wt
            ws_quote.Range("B28").Value = qty
            excel.Run("modUI.UI_AddItem")

        cap_10 = str(ws_quote.Range("F14").Value)
        tot_r31 = str(ws_quote.Range("D31").Value)
        print(f"Capacity at 10 lines: {repr(cap_10)}")
        print(f"Row 31 (Total): {tot_r31}")

        assert "10 of 10" in cap_10, "Capacity should show 10 of 10"
        assert "MAXIMUM CAPACITY REACHED" in cap_10, "Capacity badge should indicate maximum reached"
        assert "TABLE TOTAL (10 of 10 sizes)" in tot_r31, "Total row should be at Row 31 for 10 lines"

        page3_pdf = os.path.join(root, "TESTS", "acceptance_page3_full_capacity.pdf")
        ws_quote.ExportAsFixedFormat(0, page3_pdf)
        pdf_pages.append(page3_pdf)

        # Attempt 11th line
        ws_quote.Range("B23").Value = 400
        ws_quote.Range("B24").Value = 5.0
        ws_quote.Range("B28").Value = 100
        excel.Run("modUI.UI_AddItem")
        
        # Confirm line count is still 10
        tot_r31_after = str(ws_quote.Range("D31").Value)
        assert "TABLE TOTAL (10 of 10 sizes)" in tot_r31_after, "11th line should be blocked"
        test_results["Test D (Full Capacity 10 Lines)"] = "PASS"
        print("Test D: PASS")

        # ---------------------------------------------------------
        # TEST E: Edit a line
        # ---------------------------------------------------------
        print("\n--- TEST E: Edit a Line ---")
        # Select Line 4 (Row 24)
        excel.Run("modUI.HandleLineSelection", 24)
        excel.Run("modUI.UI_EditItem")

        btn_text = ""
        for s in ws_quote.Shapes:
            if s.Name == "btnQuoteAdd":
                btn_text = s.TextFrame2.TextRange.Text
                break
        form_sz = ws_quote.Range("B23").Value
        form_qty = ws_quote.Range("B28").Value
        print(f"Button text in edit mode: {btn_text}")
        print(f"Form Size loaded: {form_sz}, Form Qty: {form_qty}")

        assert "UPDATE LINE (L0004)" in btn_text, f"Button text expected UPDATE LINE (L0004), got {btn_text}"
        assert form_sz == 1400, "Line 4 size should be 1400"
        assert form_qty == 4400, "Line 4 qty should be 4400"

        page4_pdf = os.path.join(root, "TESTS", "acceptance_page4_edit_workflow.pdf")
        ws_quote.ExportAsFixedFormat(0, page4_pdf)
        pdf_pages.append(page4_pdf)

        # Modify quantity to 4500 and save
        ws_quote.Range("B28").Value = 4500
        excel.Run("modUI.UI_AddItem")

        # Verify line 4 updated and button returned to normal
        for s in ws_quote.Shapes:
            if s.Name == "btnQuoteAdd":
                btn_text = s.TextFrame2.TextRange.Text
                break
        l4_updated = str(ws_quote.Range("F24").Value)
        print(f"Button text after save: {btn_text}")
        print(f"Line 4 updated quantity cell F24: {l4_updated}")

        assert "+ ADD ITEM TO QUOTE" in btn_text, "Button text should revert to + ADD ITEM TO QUOTE"
        assert "4,500" in l4_updated, "Line 4 quantity should reflect 4,500"
        assert "TABLE TOTAL (10 of 10 sizes)" in str(ws_quote.Range("D31").Value), "Line count must remain 10"
        test_results["Test E (Edit Line)"] = "PASS"
        print("Test E: PASS")

        # ---------------------------------------------------------
        # TEST F: Duplicate a line
        # ---------------------------------------------------------
        print("\n--- TEST F: Duplicate a Line ---")
        # Remove lines 5-10 to leave 4 lines
        for l_num in range(10, 4, -1):
            excel.Run("modQuote.DeleteQuoteLine", f"L{l_num:04d}")
        excel.Run("modQuote.RecalculateQuote")
        excel.Run("modUI.RefreshQuoteLinesDisplay")

        # Duplicate Line 2 (L0002)
        excel.Run("modUI.HandleLineSelection", 22)
        excel.Run("modUI.UI_DuplicateItem")

        tot_after_dup = str(ws_quote.Range("D26").Value)
        l5_text = str(ws_quote.Range("D25").Value)
        print(f"Line 5 created: {l5_text}")
        print(f"Total row after duplicate: {tot_after_dup}")

        assert "L0005: HSAW 1300mm" in l5_text, "Duplicated line 5 should match L0002 specs"
        assert "TABLE TOTAL (5 of 10 sizes)" in tot_after_dup, "Total row should show 5 sizes"
        test_results["Test F (Duplicate Line)"] = "PASS"
        print("Test F: PASS")

        # ---------------------------------------------------------
        # TEST G: Remove a line
        # ---------------------------------------------------------
        print("\n--- TEST G: Remove a Line ---")
        # Remove L0005
        excel.Run("modUI.HandleLineSelection", 25)
        excel.Run("modUI.UI_RemoveItem")

        tot_after_rem = str(ws_quote.Range("D25").Value)
        print(f"Total row after removal: {tot_after_rem}")
        assert "TABLE TOTAL (4 of 10 sizes)" in tot_after_rem, "Total row should show 4 sizes at Row 25"
        test_results["Test G (Remove Line)"] = "PASS"
        print("Test G: PASS")

        # ---------------------------------------------------------
        # TEST H: Clear behaviors
        # ---------------------------------------------------------
        print("\n--- TEST H: Clear Behaviors ---")
        # Test CLEAR INPUTS
        ws_quote.Range("B28").Value = 9999
        excel.Run("modUI.UI_ClearLine")
        b28_val = ws_quote.Range("B28").Value
        l1_val = str(ws_quote.Range("D21").Value)
        print(f"B28 after ClearLine: {b28_val}, Line 1: {l1_val}")
        assert b28_val == 0, "B28 should be 0"
        assert "L0001" in l1_val, "Quote lines should be preserved during ClearLine"

        # Test CLEAR ALL LINES
        excel.Run("modUI.UI_ClearAllItems")
        empty_check = str(ws_quote.Range("D21").Value)
        print(f"D21 after ClearAllItems: {empty_check}")
        assert "No quotation lines added yet" in empty_check, "All lines should be cleared"
        test_results["Test H (Clear Behaviors)"] = "PASS"
        print("Test H: PASS")

        # ---------------------------------------------------------
        # TEST I: Recalculation and persistence
        # ---------------------------------------------------------
        print("\n--- TEST I: Recalculation & Persistence ---")
        excel.Run("modSelfTest.PopulateGoldenReferenceQuote")
        excel.Run("modQuote.RecalculateQuote")
        excel.Run("modUI.RefreshQuoteLinesDisplay")

        grand_tot_before = str(ws_quote.Range("I25").Text)
        print(f"Grand total before save: {grand_tot_before}")
        assert "17,88,01,050" in grand_tot_before or "178,801,050" in grand_tot_before, "Golden Total mismatch"

        page5_pdf = os.path.join(root, "TESTS", "acceptance_page5_recalculated.pdf")
        ws_quote.ExportAsFixedFormat(0, page5_pdf)
        pdf_pages.append(page5_pdf)

        wb.Save()
        wb.Close(False)

        # Reopen workbook to verify persistence
        wb2 = excel.Workbooks.Open(test_wb_path)
        ws_quote2 = wb2.Worksheets("QUOTE")
        grand_tot_after = str(ws_quote2.Range("I25").Text)
        l1_after = str(ws_quote2.Range("D21").Value)
        print(f"Grand total after reopen: {grand_tot_after}")
        print(f"Line 1 after reopen: {l1_after}")
        assert "17,88,01,050" in grand_tot_after or "178,801,050" in grand_tot_after, "Reopened total mismatch"
        assert "L0001" in l1_after, "Reopened line 1 mismatch"
        wb2.Close(False)
        test_results["Test I (Recalculation & Persistence)"] = "PASS"
        print("Test I: PASS")

        # ---------------------------------------------------------
        # TEST J: Layout and Acceptance PDF Export
        # ---------------------------------------------------------
        print("\n--- TEST J: PDF Acceptance Package ---")
        combined_pdf_path = os.path.join(root, "TESTS", "quote_sheet_ui_acceptance.pdf")
        merger = PdfWriter()
        for p in pdf_pages:
            merger.append(p)
        merger.write(combined_pdf_path)
        merger.close()
        print(f"Combined acceptance PDF generated: {combined_pdf_path} ({len(pdf_pages)} pages)")

        # Convert each page to high-res PNG for visual inspection
        doc = pymupdf.open(combined_pdf_path)
        for i, page in enumerate(doc):
            pix = page.get_pixmap(dpi=150)
            png_out = os.path.join(root, "TESTS", f"acceptance_page_{i+1}.png")
            pix.save(png_out)
            print(f"Saved PNG page {i+1}: {png_out}")
        doc.close()

        test_results["Test J (Layout & Acceptance PDF)"] = "PASS"
        print("Test J: PASS")

        # ---------------------------------------------------------
        # TEST K: Regression Suites & Golden Values
        # ---------------------------------------------------------
        print("\n--- TEST K: Regression Suites ---")
        wb_reg = excel.Workbooks.Open(dev_wb_path)
        selftest_log = excel.Run("modSelfTest.RunAllSelfTests")
        print("\n--- modSelfTest.RunAllSelfTests Output ---")
        print(selftest_log)
        assert selftest_log == "PASS" or "PASS" in selftest_log, f"Self-test failed: {selftest_log}"
        unified_file = os.path.join(root, "TESTS", "unified_selftest_p24.txt")
        with open(unified_file, "r", encoding="utf-8", errors="ignore") as f:
            log_content = f.read()
        assert "Failed: 0" in log_content, f"Self-test failures detected: {log_content}"
        assert "[PASS] P3/P5 Golden Legacy Contract" in log_content, "Legacy Golden failed"
        assert "[PASS] P3/P5 Golden Normal Option 1" in log_content, "Normal Golden failed"
        wb_reg.Close(False)

        test_results["Test K (Regression Suites & Golden Values)"] = "PASS"
        print("Test K: PASS")

    finally:
        excel.Quit()
        if os.path.exists(test_wb_path):
            try:
                os.remove(test_wb_path)
            except:
                pass

    print("\n=======================================================")
    print("ALL TESTS COMPLETED SUCCESSFULLY!")
    print("=======================================================")
    for k, v in test_results.items():
        print(f"{k}: {v}")

if __name__ == "__main__":
    run_tests()
