"""Mega Pipes - Master Migration & Integration Script
Updates Mega_Pipes_Quotation_DEV.xlsm with:
1. Full Location & Territory Intelligence (COMPARE sheet, modCompare, tblPlants, tblDistanceCache, btnQuoteCompare)
2. Null-safe competitor engine (modCompetitors)
3. Automated Cost Sheet Formulas & Sizing parity
4. Manual selling price override & audit trail preservation
"""

import os
import sys
import shutil
import win32com.client
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

def log(msg):
    print(msg)
    sys.stdout.flush()

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PILOT_BACKUP = os.path.join(ROOT_DIR, 'BACKUP', 'Mega_Pipes_Quotation_PILOT_COPY_pre_stability_investigation_20261021_114000.xlsm')
TARGET_WB = os.path.join(ROOT_DIR, 'DEV', 'Mega_Pipes_Quotation_DEV.xlsm')
TEST_TARGET = os.path.join(ROOT_DIR, 'TESTS', 'dev_migration_test.xlsm')

def perform_migration(target_path):
    log(f"\n=======================================================")
    log(f" Starting Migration for: {target_path}")
    log(f"=======================================================")
    
    # 1. First, update openpyxl sheets: PLANTS, DISTANCE_CACHE, SETTINGS, COMPARE
    log("[Step 1] Loading openpyxl workbooks to sync location master data...")
    wb_src = openpyxl.load_workbook(PILOT_BACKUP, data_only=False, keep_vba=True)
    wb_tgt = openpyxl.load_workbook(target_path, data_only=False, keep_vba=True)
    
    # Sync PLANTS data and table references
    log("[Step 1.1] Syncing PLANTS sheet (tblPlants & tblPlantProducts)...")
    ws_src_p = wb_src['PLANTS']
    ws_tgt_p = wb_tgt['PLANTS']
    # Clear existing target area up to row 35
    for r in range(1, 35):
        for c in range(1, 15):
            ws_tgt_p.cell(r, c).value = None
    # Copy from source (rows 1 to 30, cols 1 to 14)
    for r in range(1, 31):
        for c in range(1, 15):
            val = ws_src_p.cell(r, c).value
            if val is not None:
                ws_tgt_p.cell(r, c).value = val
    # Update table references
    if 'tblPlants' in ws_tgt_p.tables:
        ws_tgt_p.tables['tblPlants'].ref = 'A3:M19'
    if 'tblPlantProducts' in ws_tgt_p.tables:
        ws_tgt_p.tables['tblPlantProducts'].ref = 'A23:E29'
    log(f"  tblPlants ref set to {ws_tgt_p.tables['tblPlants'].ref}")
    log(f"  tblPlantProducts ref set to {ws_tgt_p.tables['tblPlantProducts'].ref}")
                
    # Sync DISTANCE_CACHE (rows 1 to 21)
    log("[Step 1.2] Syncing DISTANCE_CACHE sheet (tblDistanceCache)...")
    ws_src_dc = wb_src['DISTANCE_CACHE']
    ws_tgt_dc = wb_tgt['DISTANCE_CACHE']
    for r in range(1, 25):
        for c in range(1, 14):
            ws_tgt_dc.cell(r, c).value = None
    for r in range(1, 22):
        for c in range(1, 14):
            val = ws_src_dc.cell(r, c).value
            if val is not None:
                ws_tgt_dc.cell(r, c).value = val
    if 'tblDistanceCache' in ws_tgt_dc.tables:
        ws_tgt_dc.tables['tblDistanceCache'].ref = 'A3:L20'
    log(f"  tblDistanceCache ref set to {ws_tgt_dc.tables['tblDistanceCache'].ref}")
                
    # Sync SETTINGS: preserve DEV's calibrated cost sheet defaults (TripRateManual=93000, SteelRate=62000, Conversion=4000, FinPct=1)
    # Only append rows 37 and 38 for export folders
    log("[Step 1.3] Syncing SETTINGS sheet (preserving calibrated cost defaults)...")
    ws_tgt_set = wb_tgt['SETTINGS']
    ws_src_set = wb_src['SETTINGS']
    for r in range(37, 39):
        for c in range(1, 8):
            val = ws_src_set.cell(r, c).value
            if val is not None:
                ws_tgt_set.cell(r, c).value = val
    if 'tblSettings' in ws_tgt_set.tables:
        ws_tgt_set.tables['tblSettings'].ref = 'A3:G38'
    log(f"  tblSettings ref set to {ws_tgt_set.tables['tblSettings'].ref}")

    # Sync COMPARE sheet layout & content (A1 to Z65)
    log("[Step 1.4] Syncing COMPARE dashboard layout...")
    ws_src_cmp = wb_src['COMPARE']
    ws_tgt_cmp = wb_tgt['COMPARE']
    for r in range(1, 65):
        for c in range(1, 27):
            val = ws_src_cmp.cell(r, c).value
            if val is not None:
                ws_tgt_cmp.cell(r, c).value = val
                
    wb_tgt.save(target_path)
    wb_tgt.close()
    wb_src.close()
    log("[Step 1] OpenPyXL sheet data synchronization complete.")
    
    # 2. Next, use Excel COM to import VBA modules and create UI navigation buttons
    log("\n[Step 2] Starting Excel COM to update VBA modules & shapes...")
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    
    try:
        wb = excel.Workbooks.Open(target_path)
        vb_proj = wb.VBProject
        
        # 2.1 Import / Replace modCompare
        log("[Step 2.1] Importing/Updating modCompare...")
        mod_compare_src = os.path.join(ROOT_DIR, 'VBA', 'Modules', 'modCompare.bas')
        for comp in vb_proj.VBComponents:
            if comp.Name == "modCompare":
                vb_proj.VBComponents.Remove(comp)
                break
        vb_proj.VBComponents.Import(mod_compare_src)
        log("[Step 2.1] modCompare imported successfully.")
        
        # 2.2 Replace modCompetitors with Null-safe version
        log("[Step 2.2] Updating modCompetitors...")
        mod_comp_src = os.path.join(ROOT_DIR, 'VBA', 'Modules', 'modCompetitors.bas')
        for comp in vb_proj.VBComponents:
            if comp.Name == "modCompetitors":
                vb_proj.VBComponents.Remove(comp)
                break
        vb_proj.VBComponents.Import(mod_comp_src)
        log("[Step 2.2] modCompetitors updated successfully.")
        
        # 2.3 Import / Replace modUI (with btnQuoteCompare included)
        log("[Step 2.3] Updating modUI...")
        mod_ui_src = os.path.join(ROOT_DIR, 'VBA', 'Modules', 'modUI.bas')
        for comp in vb_proj.VBComponents:
            if comp.Name == "modUI":
                vb_proj.VBComponents.Remove(comp)
                break
        vb_proj.VBComponents.Import(mod_ui_src)
        log("[Step 2.3] modUI updated successfully.")

        # 2.4 Setup QUOTE sheet UI template and buttons
        log("[Step 2.4] Setting up QUOTE sheet UI and all 19 buttons...")
        excel.Run('modUI.SetupQuoteSheetUI')
        excel.Run('modSelfTest.PopulateGoldenReferenceQuote')
        log("[Step 2.4] QUOTE sheet UI setup and golden lines populated.")

        # 2.5 Add btnCompareReturn button on COMPARE sheet
        log("[Step 2.5] Checking/Adding btnCompareReturn button on COMPARE...")
        wsCmp = wb.Worksheets('COMPARE')
        has_ret = False
        for i in range(1, wsCmp.Shapes.Count + 1):
            if wsCmp.Shapes(i).Name == "btnCompareReturn":
                has_ret = True
                break
        if not has_ret:
            btn_ret = wsCmp.Shapes.AddShape(1, 10.0, 10.0, 180.0, 30.0)
            btn_ret.Name = "btnCompareReturn"
            btn_ret.TextFrame.Characters().Text = "<- RETURN TO QUOTE"
            btn_ret.TextFrame.Characters().Font.Name = "Segoe UI"
            btn_ret.TextFrame.Characters().Font.Size = 9
            btn_ret.TextFrame.Characters().Font.Bold = True
            btn_ret.TextFrame.Characters().Font.Color = 16777215
            btn_ret.Fill.Solid()
            btn_ret.Fill.ForeColor.RGB = 2121760
            btn_ret.Line.Visible = False
            btn_ret.OnAction = "modUI.UI_BackToHome"
            log("[Step 2.5] btnCompareReturn added on COMPARE.")
        else:
            log("[Step 2.5] btnCompareReturn already present.")

        # 2.6 Build Compare Sheet
        log("[Step 2.6] Building COMPARE sheet...")
        excel.Run('modCompare.BuildCompareSheet', 'KM')

        # 2.6 Save workbook
        wb.Save()
        log("[Step 2] Workbook saved successfully via Excel COM.")
        
        # 3. Test verification
        log("\n[Step 3] Running Self-Tests and Diagnostics...")
        try:
            res_comp = excel.Run('modCompare.RunCompareSelfTests')
            log(f"  modCompare.RunCompareSelfTests: {res_comp}")
        except Exception as e:
            log(f"  modCompare error: {e}")
            
        try:
            res_comp_p10 = excel.Run('modCompetitors.RunCompetitorSelfTests')
            log(f"  modCompetitors.RunCompetitorSelfTests: {res_comp_p10}")
        except Exception as e:
            log(f"  modCompetitors error: {e}")

        try:
            res_all = excel.Run('modSelfTest.RunAllSelfTests')
            log(f"  modSelfTest.RunAllSelfTests: {res_all}")
        except Exception as e:
            log(f"  RunAllSelfTests error: {e}")

        try:
            res_approved = excel.Run('modPricing.ApprovedGoldenTotal')
            log(f"  modPricing.ApprovedGoldenTotal: {res_approved} (Expected: 176117500)")
        except Exception as e:
            log(f"  ApprovedGoldenTotal error: {e}")

        try:
            res_norm = excel.Run('modPricing.P3GoldenTotal', False)
            log(f"  modPricing.P3GoldenTotal(False): {res_norm} (Expected: 178801050)")
        except Exception as e:
            log(f"  P3GoldenTotal(False) error: {e}")

        wb.Save()
        wb.Close(True)
        del wb
        log("[Step 3] Verification complete.")
        
    finally:
        excel.Quit()
        del excel
        log("Excel Application closed.")

if __name__ == '__main__':
    log(f"Migrating directly to master DEV workbook: {TARGET_WB}")
    perform_migration(TARGET_WB)

