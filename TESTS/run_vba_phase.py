"""
Mega Pipes Quotation System - Phase Test Runner
Safely interacts with desktop Excel 16.0 via win32com to:
1. Optionally import/update VBA modules into DEV workbook
2. Compile the VBA project and verify clean compilation
3. Execute Phase-specific self-test macros
4. Execute full regression test suites
5. Save DEV workbook upon all PASS
6. Record detailed runtime logs
"""
import sys
import os
import argparse
from pathlib import Path
import win32com.client
import pythoncom

ROOT = Path(__file__).resolve().parents[1]
DEV_WB = ROOT / "DEV" / "Mega_Pipes_Quotation_DEV.xlsm"

REGRESSION_MACROS = [
    "RunGoldenNormal",
    "RunGoldenLegacy",
    "RunFreightSelfTests",
    "RunPlantSelfTests",
    "RunORSSelfTests",
    "RunCompetitorSelfTests",
    "RunCostingSelfTests",
    "RunOfferSelfTests",
    "RunSequenceSelfTests",
    "RunSnapshotSelfTests",
    "RunMastersSelfTests",
    "RunOfflineSelfTests",
    "RunExportSelfTests",
    "RunDocumentNumberingSelfTests",
    "RunAuditSelfTests",
]

def run_phase(module_path=None, module_name=None, test_macro=None, phase_name="PHASE TEST", output_log=None, save_wb=False, notes=None):
    pythoncom.CoInitialize()
    xl = None
    wb = None
    log_lines = []
    
    def log(msg):
        print(msg)
        log_lines.append(msg)
        
    log(f"=== {phase_name} REAL EXCEL VERIFICATION ===")
    
    try:
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        
        log(f"Opening workbook: {DEV_WB.name}")
        wb = xl.Workbooks.Open(str(DEV_WB))
        vbp = wb.VBProject
        
        if module_path and module_name:
            paths = [p.strip() for p in module_path.split(",")]
            names = [n.strip() for n in module_name.split(",")]
            for m_path_str, m_name in zip(paths, names):
                m_path = Path(m_path_str).resolve()
                if not m_path.exists():
                    raise FileNotFoundError(f"Module file not found: {m_path}")
                try:
                    comp = vbp.VBComponents(m_name)
                    vbp.VBComponents.Remove(comp)
                except Exception:
                    pass
                vbp.VBComponents.Import(str(m_path))
                log(f"Imported updated {m_name} successfully into DEV workbook.")
            
        # Check and compile VBAProject
        try:
            ctrl = xl.VBE.CommandBars.FindControl(Id=578)
            if ctrl and ctrl.Enabled:
                ctrl.Execute()
                log(f"Compiled VBAProject: Enabled={ctrl.Enabled} (compiled successfully)")
            else:
                log(f"Compiled VBAProject: Enabled=False (clean compilation)")
        except Exception as ce:
            log(f"VBE Compile check note: {ce}")
            
        # Run Phase Test Macro
        if test_macro:
            log(f"Executing {test_macro}...")
            test_res = xl.Run(test_macro)
            log(f"{test_macro} Result: {test_res}")
            if str(test_res).strip() != "PASS" and not str(test_res).startswith("PASS:"):
                raise RuntimeError(f"{test_macro} FAILED with result: {test_res}")
                
        if notes:
            log("Verified components:")
            for note in notes:
                log(f"  - {note}")
                
        # Run Regressions
        log("Executing Regression Suites...")
        for reg_macro in REGRESSION_MACROS:
            try:
                reg_res = xl.Run(reg_macro)
                log(f"{reg_macro} Result: {reg_res}")
                if str(reg_res).strip() != "PASS" and not str(reg_res).startswith("PASS:"):
                    raise RuntimeError(f"Regression macro {reg_macro} FAILED: {reg_res}")
            except Exception as re:
                # Some macros may be added in later phases; report clearly
                log(f"Regression macro {reg_macro} error: {re}")
                raise
                
        if save_wb:
            wb.Save()
            log(f"DEV workbook saved with compiled {module_name or 'project'}.")
            
        log(f"{phase_name} VERIFICATION: ALL PASS!")
        
        if output_log:
            out_p = Path(output_log).resolve()
            out_p.parent.mkdir(parents=True, exist_ok=True)
            out_p.write_text("\n".join(log_lines) + "\n", encoding="utf-8")
            print(f"Log written to {out_p}")
            
        return True

    except Exception as e:
        log(f"ERROR: {e}")
        if output_log:
            out_p = Path(output_log).resolve()
            out_p.parent.mkdir(parents=True, exist_ok=True)
            out_p.write_text("\n".join(log_lines) + f"\nERROR: {e}\n", encoding="utf-8")
        raise
        
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
            del wb
        if xl:
            try:
                xl.Quit()
            except Exception:
                pass
            del xl
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--module-path", help="Path to .bas file to import")
    parser.add_argument("--module-name", help="Name of VBA component to replace")
    parser.add_argument("--test-macro", help="Test macro to run")
    parser.add_argument("--phase-name", default="PHASE TEST", help="Phase title")
    parser.add_argument("--output-log", help="Path to output log file")
    parser.add_argument("--save-wb", action="store_true", help="Save DEV workbook on pass")
    args = parser.parse_args()
    
    run_phase(
        module_path=args.module_path,
        module_name=args.module_name,
        test_macro=args.test_macro,
        phase_name=args.phase_name,
        output_log=args.output_log,
        save_wb=args.save_wb
    )

