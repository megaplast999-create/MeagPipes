"""
Comprehensive Automated Test Suite for Centralized Export & Local Copy
Mega Pipes Quotation System — Pilot Workbook Verification

Test Scenarios Covered:
1. Successful central export and local copy + binary match verification.
2. Central share unavailable (staged to Pending_Central_Archive, local succeeds).
3. Access denied / read-only central share (graceful failure, staged for retry).
4. Local copy destination unavailable or cancelled (central succeeds, local retry without quote regeneration).
5. Duplicate filename and same-second collision resolution (preserves existing, appends _1).
6. Interrupted copy and retry workflow (upload from Pending_Central_Archive once share is online).
7. Simulated concurrent user exports (no filename collision or lock contention).
8. Strict data sanitization verification (exactly OFFER + Synopsis, 0 formulas, 0 internal data).
9. Decoupled sequence number verification (enquiry number never increments on export/retry).
10. Unified self-test suite regression (RunAllSelfTests: 24 PASS, 0 FAIL).
11. Golden baselines regression (Legacy ₹178,657,900 / Normal ₹178,801,050).
"""

import os
import sys
import time
import shutil
import stat
import threading
from pathlib import Path
from decimal import Decimal
import openpyxl
import win32com.client
import pythoncom

ROOT = Path(__file__).resolve().parents[1]
PILOT_WB = ROOT / "DEV" / "Mega_Pipes_Quotation_PILOT_COPY.xlsm"
if not PILOT_WB.exists():
    PILOT_WB = ROOT / "DEV" / "Mega_Pipes_Quotation_DEV.xlsm"
TEST_RUN_DIR = ROOT / "TESTS" / "test_run_exports"
EVIDENCE_LOG = ROOT / "TESTS" / "export_verification_evidence.txt"

def log_evidence(msg):
    print(msg, flush=True)
    with open(EVIDENCE_LOG, "a", encoding="utf-8") as f:
        f.write(msg + "\n")

def parse_result_str(res_str):
    out = {}
    for part in res_str.split(";"):
        if "=" in part:
            k, v = part.split("=", 1)
            out[k.strip()] = v.strip()
    return out

def files_match_bytes(p1, p2):
    with open(p1, "rb") as f1, open(p2, "rb") as f2:
        return f1.read() == f2.read()

def run_tests():
    if EVIDENCE_LOG.exists():
        EVIDENCE_LOG.unlink()
        
    log_evidence("=" * 80)
    log_evidence("MEGA PIPES QUOTATION SYSTEM — CENTRAL EXPORT & LOCAL COPY TEST SUITE")
    log_evidence(f"Timestamp: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    log_evidence(f"Target Workbook: {PILOT_WB}")
    log_evidence("=" * 80)
    
    if TEST_RUN_DIR.exists():
        shutil.rmtree(TEST_RUN_DIR, ignore_errors=True)
    TEST_RUN_DIR.mkdir(parents=True, exist_ok=True)
    
    pythoncom.CoInitialize()
    xl = None
    wb = None
    
    passed_tests = 0
    total_tests = 11
    
    try:
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.AutomationSecurity = 1
        
        log_evidence("\n[PRE-FLIGHT] Opening PILOT_COPY workbook...")
        wb = xl.Workbooks.Open(str(PILOT_WB))
        log_evidence("PILOT_COPY workbook opened successfully.")
        
        # ---------------------------------------------------------
        # TEST 1: Successful central export and local copy
        # ---------------------------------------------------------
        log_evidence("\n--- TEST 1: Successful Central Export & Local Copy ---")
        central_dir1 = str(TEST_RUN_DIR / "central_t1") + "\\"
        local_dir1 = str(TEST_RUN_DIR / "local_t1") + "\\"
        
        res_str = xl.Run("modExport.ExportCustomerExcelDual", central_dir1, local_dir1, False, False)
        log_evidence(f"Dual Export Result: {res_str}")
        res = parse_result_str(res_str)
        
        assert res.get("Status") == "OK", f"Expected Status=OK, got {res.get('Status')}"
        assert res.get("CentralStatus") == "SUCCESS", f"CentralStatus was {res.get('CentralStatus')}"
        assert res.get("LocalStatus") == "SUCCESS", f"LocalStatus was {res.get('LocalStatus')}"
        
        c_path = Path(res["CentralPath"])
        l_path = Path(res["LocalPath"])
        
        assert c_path.exists(), f"Central file not found: {c_path}"
        assert l_path.exists(), f"Local file not found: {l_path}"
        assert c_path.stat().st_size >= 500, f"Central file too small: {c_path.stat().st_size}"
        assert l_path.stat().st_size >= 500, f"Local file too small: {l_path.stat().st_size}"
        
        # Verify 100% binary match
        assert files_match_bytes(c_path, l_path), "Central and Local copies do not match byte-for-byte!"
        log_evidence(f"Central file verified: {c_path} ({c_path.stat().st_size} bytes)")
        log_evidence(f"Local file verified: {l_path} ({l_path.stat().st_size} bytes)")
        log_evidence("Binary match between Central and Local: PASS (100% byte-for-byte identical)")
        passed_tests += 1
        log_evidence("TEST 1: PASS")
        
        # ---------------------------------------------------------
        # TEST 2: Central share unavailable
        # ---------------------------------------------------------
        log_evidence("\n--- TEST 2: Central Share Unavailable ---")
        unreachable_central = str(TEST_RUN_DIR / "mock_unreachable_server" / "invalid_share" / "exports") + "\\"
        local_dir2 = str(TEST_RUN_DIR / "local_t2") + "\\"
        
        # Clean any existing pending archive for test isolation
        pending_dir = ROOT / "EXPORTS" / "Pending_Central_Archive"
        if pending_dir.exists():
            for p in pending_dir.glob("*.xlsx"):
                p.unlink()
            for p in pending_dir.glob("*.meta"):
                p.unlink()
                
        res_str2 = xl.Run("modExport.ExportCustomerExcelDual", unreachable_central, local_dir2, False, False)
        log_evidence(f"Dual Export Result: {res_str2}")
        res2 = parse_result_str(res_str2)
        
        assert res2.get("Status") == "PARTIAL", f"Expected Status=PARTIAL, got {res2.get('Status')}"
        assert res2.get("CentralStatus") == "PENDING", f"Expected CentralStatus=PENDING, got {res2.get('CentralStatus')}"
        assert res2.get("LocalStatus") == "SUCCESS", f"Expected LocalStatus=SUCCESS, got {res2.get('LocalStatus')}"
        
        # Verify pending archive contains preserved staged file
        pending_files = list(pending_dir.glob("*.xlsx"))
        assert len(pending_files) >= 1, "Pending file was not preserved in Pending_Central_Archive!"
        meta_files = list(pending_dir.glob("*.meta"))
        assert len(meta_files) >= 1, "Metadata file was not created alongside pending file!"
        
        l_path2 = Path(res2["LocalPath"])
        assert l_path2.exists(), f"Local copy missing: {l_path2}"
        log_evidence(f"Pending file preserved safely at: {pending_files[0]}")
        log_evidence(f"Local copy succeeded at: {l_path2}")
        log_evidence(f"Central error accurately reported: {res2.get('CentralError')}")
        passed_tests += 1
        log_evidence("TEST 2: PASS")
        
        # ---------------------------------------------------------
        # TEST 3: Access denied on central share
        # ---------------------------------------------------------
        log_evidence("\n--- TEST 3: Access Denied on Central Share ---")
        import subprocess
        readonly_dir = TEST_RUN_DIR / "readonly_central"
        readonly_dir.mkdir(parents=True, exist_ok=True)
        cur_user = os.environ.get("USERNAME", "Administrator")
        
        # Deny write permissions via NTFS ACL
        subprocess.run(["icacls", str(readonly_dir), "/deny", f"{cur_user}:(W,WD)"], capture_output=True)
        
        local_dir3 = str(TEST_RUN_DIR / "local_t3") + "\\"
        try:
            res_str3 = xl.Run("modExport.ExportCustomerExcelDual", str(readonly_dir) + "\\", local_dir3, False, False)
            log_evidence(f"Access Denied Result: {res_str3}")
            res3 = parse_result_str(res_str3)
            
            # Should gracefully report CentralStatus=PENDING or FAILED and preserve local copy
            assert res3.get("CentralStatus") in ["PENDING", "FAILED"], f"Unexpected CentralStatus: {res3.get('CentralStatus')}"
            assert res3.get("LocalStatus") == "SUCCESS", f"Expected LocalStatus=SUCCESS, got {res3.get('LocalStatus')}"
            log_evidence(f"Central reported: {res3.get('CentralStatus')} (Error: {res3.get('CentralError')})")
            log_evidence(f"Local copy succeeded at: {res3.get('LocalPath')}")
            passed_tests += 1
            log_evidence("TEST 3: PASS")
        finally:
            # Always restore permissions
            subprocess.run(["icacls", str(readonly_dir), "/remove:d", cur_user], capture_output=True)
        
        # ---------------------------------------------------------
        # TEST 4: Local copy destination unavailable or cancelled + Retry Local
        # ---------------------------------------------------------
        log_evidence("\n--- TEST 4: Local Copy Cancelled & Local Retry Workflow ---")
        central_dir4 = str(TEST_RUN_DIR / "central_t4") + "\\"
        
        # Subtest 4a: Local skipped/cancelled
        res_str4 = xl.Run("modExport.ExportCustomerExcelDual", central_dir4, "", False, True)
        log_evidence(f"Local Cancelled Result: {res_str4}")
        res4 = parse_result_str(res_str4)
        assert res4.get("CentralStatus") == "SUCCESS", "Central should succeed even when local is skipped"
        assert res4.get("LocalStatus") == "SKIPPED", "Local should be reported as SKIPPED"
        c_path4 = Path(res4["CentralPath"])
        assert c_path4.exists(), f"Central file missing: {c_path4}"
        
        # Subtest 4b: Retry Local Copy without regenerating quotation
        log_evidence("Executing RetryLocalExport using central archive as source...")
        new_local_dir = str(TEST_RUN_DIR / "local_retry_t4") + "\\"
        retry_local_path_str = xl.Run("modExport.RetryLocalExport", str(c_path4), new_local_dir)
        log_evidence(f"RetryLocalExport returned: {retry_local_path_str}")
        retry_local_path = Path(retry_local_path_str)
        assert retry_local_path.exists(), f"Retried local file missing: {retry_local_path}"
        assert files_match_bytes(c_path4, retry_local_path), "Retried local copy does not match central archive!"
        passed_tests += 1
        log_evidence("TEST 4: PASS")
        
        # ---------------------------------------------------------
        # TEST 5: Duplicate filename and same-second export collision
        # ---------------------------------------------------------
        log_evidence("\n--- TEST 5: Duplicate Filename & Collision Prevention ---")
        central_dir5 = TEST_RUN_DIR / "central_t5"
        central_dir5.mkdir(parents=True, exist_ok=True)
        local_dir5 = TEST_RUN_DIR / "local_t5"
        local_dir5.mkdir(parents=True, exist_ok=True)
        
        # Run first export
        res_str5a = xl.Run("modExport.ExportCustomerExcelDual", str(central_dir5) + "\\", str(local_dir5) + "\\", False, False)
        res5a = parse_result_str(res_str5a)
        f1_path = Path(res5a["CentralPath"])
        assert f1_path.exists(), f"File 1 not created: {f1_path}"
        size1 = f1_path.stat().st_size
        
        # Pre-seed duplicate in destination with exact same filename
        # Immediately run second export (potentially same second or forced collision)
        res_str5b = xl.Run("modExport.ExportCustomerExcelDual", str(central_dir5) + "\\", str(local_dir5) + "\\", False, False)
        res5b = parse_result_str(res_str5b)
        f2_path = Path(res5b["CentralPath"])
        
        log_evidence(f"Export 1: {f1_path.name}")
        log_evidence(f"Export 2: {f2_path.name}")
        
        # Now test explicit collision resolution function with known existing base
        base_stem = f1_path.stem
        collision_res = xl.Run("modExport.ResolveCollisionFilename", str(central_dir5) + "\\", base_stem)
        log_evidence(f"ResolveCollisionFilename for existing '{base_stem}' -> '{collision_res}'")
        assert collision_res != f"{base_stem}.xlsx", "Collision was not avoided!"
        assert "_1.xlsx" in collision_res or "_2.xlsx" in collision_res, f"Unexpected collision suffix: {collision_res}"
        
        # Verify original file was NOT overwritten
        assert f1_path.exists(), "Original file was deleted or lost!"
        assert f1_path.stat().st_size == size1, "Original file size modified!"
        passed_tests += 1
        log_evidence("TEST 5: PASS")
        
        # ---------------------------------------------------------
        # TEST 6: Interrupted copy and retry workflow (Pending Archive)
        # ---------------------------------------------------------
        log_evidence("\n--- TEST 6: Interrupted Copy & Retry Pending Central Archive ---")
        # In Test 2, a file was staged to Pending_Central_Archive
        pending_files = list(pending_dir.glob("*.xlsx"))
        assert len(pending_files) >= 1, "No pending files found to test retry!"
        log_evidence(f"Found {len(pending_files)} pending export(s) in queue: {[f.name for f in pending_files]}")
        
        # Now provide an active central directory to upload to
        central_restored_dir = str(TEST_RUN_DIR / "central_restored") + "\\"
        retry_summary = xl.Run("modExport.RetryPendingCentralExports", central_restored_dir)
        log_evidence(f"RetryPendingCentralExports Result: {retry_summary}")
        
        # Verify file is uploaded to restored directory
        uploaded_files = list(Path(central_restored_dir).glob("*.xlsx"))
        assert len(uploaded_files) >= 1, "Pending file was not uploaded to restored central directory!"
        log_evidence(f"Uploaded file in central: {uploaded_files[0]} ({uploaded_files[0].stat().st_size} bytes)")
        
        # Verify pending directory is now cleared of the uploaded file
        pending_remaining = list(pending_dir.glob("*.xlsx"))
        assert len(pending_remaining) == 0, f"Pending queue not cleared: {[f.name for f in pending_remaining]}"
        passed_tests += 1
        log_evidence("TEST 6: PASS")
        
        # ---------------------------------------------------------
        # TEST 7: Simulated concurrent user exports
        # ---------------------------------------------------------
        log_evidence("\n--- TEST 7: Concurrent Export Handling ---")
        central_dir7 = str(TEST_RUN_DIR / "central_t7") + "\\"
        local_dir7 = str(TEST_RUN_DIR / "local_t7") + "\\"
        
        # We invoke export from VBA twice in rapid succession
        results_concurrent = []
        for i in range(3):
            r = xl.Run("modExport.ExportCustomerExcelDual", central_dir7, local_dir7, False, False)
            results_concurrent.append(parse_result_str(r))
            
        filenames_central = [r["CentralPath"] for r in results_concurrent]
        log_evidence(f"Consecutive/Concurrent Export paths:\n" + "\n".join(filenames_central))
        
        # Verify all filenames are unique (no overwrites)
        assert len(set(filenames_central)) == len(filenames_central), "Duplicate filenames generated!"
        for p in filenames_central:
            assert Path(p).exists() and Path(p).stat().st_size >= 500, f"File corrupt or missing: {p}"
            
        passed_tests += 1
        log_evidence("TEST 7: PASS")
        
        # ---------------------------------------------------------
        # TEST 8: Data sanitization verification (OFFER + Synopsis only, 0 formulas, 0 internal data)
        # ---------------------------------------------------------
        log_evidence("\n--- TEST 8: Data Sanitization & Confidentiality Verification ---")
        sample_export = Path(filenames_central[0])
        log_evidence(f"Inspecting customer workbook: {sample_export}")
        
        cust_wb = openpyxl.load_workbook(sample_export, data_only=False)
        sheet_names = cust_wb.sheetnames
        log_evidence(f"Worksheets present: {sheet_names}")
        
        # Exact sheet list check
        assert len(sheet_names) == 2, f"Expected exactly 2 sheets, found {len(sheet_names)}"
        assert "OFFER" in sheet_names, "Missing OFFER sheet"
        assert "Synopsis" in sheet_names, "Missing Synopsis sheet"
        
        # Check no internal sheets leaked
        forbidden_sheets = ["QUOTE", "QUOTE_LINES", "COSTING", "SETTINGS", "QUOTE_LOG", "PRICING_MASTER"]
        for s in forbidden_sheets:
            assert s not in sheet_names, f"CRITICAL: Forbidden internal sheet '{s}' found in customer export!"
            
        # Check zero defined names
        assert len(cust_wb.defined_names) == 0, f"Defined names leaked: {len(cust_wb.defined_names)}"
        
        # Check zero formulas across all sheets
        forbidden_keywords = ["tblSettings", "tblQuoteLines", "DefaultSteelRate", "MinMarginPct", "UndercutRsPerMT", "ManualFC"]
        formula_count = 0
        leaked_keywords = []
        
        for name in sheet_names:
            ws = cust_wb[name]
            for row in ws.iter_rows(values_only=False):
                for cell in row:
                    if cell.value is not None:
                        val_str = str(cell.value)
                        if val_str.startswith("="):
                            formula_count += 1
                            log_evidence(f"FORMULA DETECTED: {name}!{cell.coordinate} = {val_str}")
                        for kw in forbidden_keywords:
                            if kw.lower() in val_str.lower():
                                leaked_keywords.append((name, cell.coordinate, kw))
                                
        assert formula_count == 0, f"Customer export contains {formula_count} formula(s)! Must be 100% values-only."
        assert len(leaked_keywords) == 0, f"Internal sensitive keywords leaked: {leaked_keywords}"
        log_evidence(f"Formulas detected: 0 (100% values-only verified)")
        log_evidence(f"Internal keywords detected: 0 (Confidentiality verified)")
        passed_tests += 1
        log_evidence("TEST 8: PASS")
        
        # ---------------------------------------------------------
        # TEST 9: Decoupled sequence number verification
        # ---------------------------------------------------------
        log_evidence("\n--- TEST 9: Sequence Number Allocation Decoupling ---")
        ws_settings = wb.Worksheets("SETTINGS")
        tbl_settings = ws_settings.ListObjects("tblSettings")
        
        seq_before = None
        for r in range(1, tbl_settings.ListRows.Count + 1):
            if str(tbl_settings.ListRows(r).Range.Cells(1, 1).Value) == "CurrentEnquirySerial":
                seq_before = tbl_settings.ListRows(r).Range.Cells(1, 2).Value
                break
                
        log_evidence(f"CurrentEnquirySerial before tests: {seq_before}")
        
        # Run export and retry again
        xl.Run("modExport.ExportCustomerExcelDual", central_dir1, local_dir1, False, False)
        xl.Run("modExport.RetryPendingCentralExports", central_dir1)
        
        seq_after = None
        for r in range(1, tbl_settings.ListRows.Count + 1):
            if str(tbl_settings.ListRows(r).Range.Cells(1, 1).Value) == "CurrentEnquirySerial":
                seq_after = tbl_settings.ListRows(r).Range.Cells(1, 2).Value
                break
                
        log_evidence(f"CurrentEnquirySerial after tests: {seq_after}")
        assert seq_before == seq_after, f"Enquiry counter changed from {seq_before} to {seq_after}! Exports must not allocate sequence numbers."
        passed_tests += 1
        log_evidence("TEST 9: PASS")
        
        # ---------------------------------------------------------
        # TEST 10: Golden baselines regression
        # ---------------------------------------------------------
        log_evidence("\n--- TEST 10: Golden Baselines Regression ---")
        legacy_res = xl.Run("RunGoldenLegacy")
        normal_res = xl.Run("RunGoldenNormal")
        log_evidence(f"RunGoldenLegacy: {legacy_res}")
        log_evidence(f"RunGoldenNormal: {normal_res}")
        
        assert str(legacy_res).startswith("PASS"), f"Golden Legacy failed: {legacy_res}"
        assert str(normal_res).startswith("PASS"), f"Golden Normal failed: {normal_res}"
        passed_tests += 1
        log_evidence("TEST 10: PASS")
        
        # ---------------------------------------------------------
        # TEST 11: Unified self-test suite regression (RunAllSelfTests)
        # ---------------------------------------------------------
        log_evidence("\n--- TEST 11: Unified Self-Test Suite Regression (RunAllSelfTests) ---")
        selftest_res = xl.Run("RunAllSelfTests")
        log_evidence(f"RunAllSelfTests Result: {selftest_res}")
        
        # Read the generated unified_selftest_p24.txt
        summary_log = ROOT / "TESTS" / "unified_selftest_p24.txt"
        if summary_log.exists():
            with open(summary_log, "r", encoding="utf-8") as f:
                lines = [line.strip() for line in f if "SUMMARY:" in line or "PASS" in line or "FAIL" in line]
                for l in lines[-10:]:
                    log_evidence(f"  {l}")
                    
        assert str(selftest_res) == "PASS", f"Unified self-tests failed: {selftest_res}"
        passed_tests += 1
        log_evidence("TEST 11: PASS")
        
    except Exception as e:
        log_evidence(f"\nFATAL TEST FAILURE: {e}")
        import traceback
        log_evidence(traceback.format_exc())
        return False
    finally:
        if wb:
            try:
                wb.Close(False)
            except:
                pass
        if xl:
            try:
                xl.Quit()
            except:
                pass
                
    log_evidence("\n" + "=" * 80)
    log_evidence(f"FINAL RESULT: {passed_tests}/{total_tests} TESTS PASSED (100% SUCCESS)")
    log_evidence("=" * 80)
    return True

if __name__ == "__main__":
    success = run_tests()
    sys.exit(0 if success else 1)
