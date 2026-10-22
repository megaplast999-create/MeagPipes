# Performance and Reliability Report (P28)

**Execution Date:** 2026-10-17  
**System Platform:** Windows 11 (AMD64)  
**Processor / Cores:** Intel64 Family 6 Model 186 Stepping 3, GenuineIntel (12 logical cores)  
**Excel Application:** Microsoft Excel Version 16.0 (Build 20430.0)  
**Target Specification:** All core user interactions (open, recalculation, snapshot save/reload) **under 5.0 seconds**

---

## 1. Measured Performance Results

| Operation | Repetitions | Min (s) | Median (s) | Max (s) | Mean (s) | Target | Result |
|---|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| **Workbook Open** | 3 | 0.7653s | **0.9403s** | 2.0650s | 1.2569s | < 5.0s | **PASS** |
| **Quotation Recalculation** | 5 | 0.2234s | **0.2958s** | 0.6122s | 0.3562s | < 5.0s | **PASS** |
| **Quote Snapshot Save** | 3 | 0.2693s | **0.3776s** | 0.7673s | 0.4714s | < 5.0s | **PASS** |
| **Quote Snapshot Reload** | 3 | 0.4016s | **0.4616s** | 0.6727s | 0.5120s | < 5.0s | **PASS** |
| **Unified Self-Test Suite (27 components)** | 3 | 13.7368s | **14.2939s** | 17.6525s | 15.2277s | < 25.0s | **PASS** |

---

## 2. Bottleneck Analysis & Reliability Observations

1. **Recalculation Efficiency**:
   - `RecalculateQuote` completes in **0.2958s** median time, well below the 5-second target.
   - Vectorized data reads, event suppression (`Application.EnableEvents = False`), and batch array processing avoid unnecessary worksheet recalculation cycles.
2. **Snapshot Saving & Immutable Reload**:
   - Saving an immutable snapshot workbook takes **0.3776s**.
   - Reloading and verifying an existing snapshot takes **0.4616s**.
   - Screen updating suppression (`Application.ScreenUpdating = False`) ensures smooth, flicker-free background execution.
3. **Memory and Resource Stability**:
   - Excel COM handles are deterministically closed and unreferenced.
   - Zero memory leaks or dangling processes observed across repeated execution loops.
