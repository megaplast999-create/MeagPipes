"""
Mega Pipes - Quotation Price Stability Stage Investigation
Instrumented with immediate flushes to track execution.
"""

import os
import sys
import shutil
import win32com.client

sys.stdout.reconfigure(encoding='utf-8')

def log(msg):
    print(msg)
    sys.stdout.flush()

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_PILOT = os.path.join(ROOT_DIR, 'DEV', 'Mega_Pipes_Quotation_PILOT_COPY.xlsm')
if not os.path.exists(SRC_PILOT):
    SRC_PILOT = os.path.join(ROOT_DIR, 'DEV', 'Mega_Pipes_Quotation_DEV.xlsm')
WORK_WB = os.path.join(ROOT_DIR, 'TESTS', 'stability_test_work.xlsm')

def run_stability_investigation():
    log(f"Copying {SRC_PILOT} -> {WORK_WB}")
    shutil.copy2(SRC_PILOT, WORK_WB)

    log("Starting Excel application...")
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False  # Avoid Workbook_Open clearing inputs

    try:
        log(f"Opening workbook {WORK_WB}...")
        wb = excel.Workbooks.Open(WORK_WB)
        wsQ = wb.Worksheets('QUOTE')
        wsLines = wb.Worksheets('QUOTE_LINES')
        tblLines = wsLines.ListObjects('tblQuoteLines')
        
        # Ensure exactly 2 lines (L0001: 2000m, L0002: 350m)
        log(f"Current lines count in table: {tblLines.ListRows.Count}")
        if tblLines.ListRows.Count > 2:
            for _ in range(tblLines.ListRows.Count - 2):
                tblLines.ListRows(3).Delete()
            log("Trimmed to 2 lines.")
    
        def dump_state(stage_name):
            log(f"\n=======================================================")
            log(f" {stage_name}")
            log(f"=======================================================")
            for r in range(1, tblLines.ListRows.Count + 1):
                lid = tblLines.DataBodyRange.Cells(r, 1).Value
                sz = tblLines.DataBodyRange.Cells(r, 6).Value
                wt = tblLines.DataBodyRange.Cells(r, 7).Value
                qty = tblLines.DataBodyRange.Cells(r, 11).Value
                uprice = tblLines.DataBodyRange.Cells(r, 20).Value
                ltotal = tblLines.DataBodyRange.Cells(r, 22).Value
                status = tblLines.DataBodyRange.Cells(r, 24).Value
                log(f"  Line {lid}: {sz}mm ID x {wt}mm WT | Qty: {qty:,.0f} m | Unit Price: Rs. {uprice:,.2f} | Line Total: Rs. {ltotal:,.2f} | Status: {status}")
            
            log("\n  QUOTE Sheet UI Display Table:")
            for r in range(21, 24):
                desc = wsQ.Cells(r, 4).Value
                c_ext_int = wsQ.Cells(r, 5).Value
                qty_str = wsQ.Cells(r, 6).Value
                pcs_wt = wsQ.Cells(r, 7).Value
                prc = wsQ.Cells(r, 8).Value
                tot = wsQ.Cells(r, 9).Value
                log(f"    Row {r}: {desc} | {c_ext_int} | {qty_str} | {pcs_wt} | Price: {prc} | Total: {tot}")
            
            e8 = wsQ.Range("E8").Value
            e9 = wsQ.Range("E9").Value
            e10 = wsQ.Range("E10").Value
            e11 = wsQ.Range("E11").Value
            log(f"\n  Live Summary Card:")
            log(f"    Subtotal (E8):        {e8}")
            log(f"    GST @ 18% (E9):       {e9}")
            log(f"    Grand Total (E10):    {e10}")
            log(f"    Workflow Status (E11): {e11}")
        
        # ── Stage 1: Immediately After Setup (2 Lines) ──
        log("Running modUI.RefreshQuoteLinesDisplay for Stage 1...")
        excel.Run('modUI.RefreshQuoteLinesDisplay')
        dump_state("Stage 1: Immediately After Adding Lines (L0001 2,000m, L0002 350m)")
        
        # ── Stage 2: Explicit Recalculation ──
        log("Running modQuote.RecalculateQuote for Stage 2...")
        excel.Run('modQuote.RecalculateQuote')
        log("Running modUI.RefreshQuoteLinesDisplay for Stage 2...")
        excel.Run('modUI.RefreshQuoteLinesDisplay')
        dump_state("Stage 2: After Explicit Recalculation (modQuote.RecalculateQuote)")
        
        # ── Stage 3: Save Quote Snapshot & Close ──
        log("Running modLog.SaveQuote for Stage 3...")
        snap_path = excel.Run('modLog.SaveQuote', 'MKTG_ENQ_26_27_594', 0)
        log(f"[Stage 3] Snapshot saved: {snap_path}")
        wb.Save()
        wb.Close(True)
        log("[Stage 3] Workbook saved and closed.")
        
        # ── Stage 4: Reopen & Load Snapshot ──
        log("Reopening workbook for Stage 4...")
        wb = excel.Workbooks.Open(WORK_WB)
        wsQ = wb.Worksheets('QUOTE')
        wsLines = wb.Worksheets('QUOTE_LINES')
        tblLines = wsLines.ListObjects('tblQuoteLines')
        
        log(f"Running modLog.LoadQuote with {snap_path}...")
        ok = excel.Run('modLog.LoadQuote', snap_path)
        log(f"[Stage 4] modLog.LoadQuote returned: {ok}")
        excel.Run('modUI.RefreshQuoteLinesDisplay')
        dump_state("Stage 4: After Reopening and Loading Saved Snapshot (modLog.LoadQuote)")
        
        # ── Stage 5: Recalculate Again ──
        log("Running modQuote.RecalculateQuote for Stage 5...")
        excel.Run('modQuote.RecalculateQuote')
        excel.Run('modUI.RefreshQuoteLinesDisplay')
        dump_state("Stage 5: After Recalculating Again (modQuote.RecalculateQuote)")
        
        wb.Close(False)
        log("Investigation completed successfully.")

    finally:
        excel.Quit()
        log("Excel application closed.")

if __name__ == '__main__':
    run_stability_investigation()
