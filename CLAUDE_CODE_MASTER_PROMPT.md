# MEGA PIPES QUOTATION WORKBOOK — MASTER PROMPT FOR CLAUDE CODE (Sonnet 4.5)
# Version 3.0 — Verified against Unipro + Lodha references + P0–P30 completion report

## IMPORTANT NOTE ON PI (most critical change from prior sessions)

The completion report (P0–P30) and the Unipro audit used PI=3.14 and produced
golden total 178,801,050 (normal) / 178,657,900 (legacy).

This Version 3.0 specifies PI_Coating = 22/7 = 3.142857... for the production
engine (Lodha reference), stored as CDec("3.142857142857143") in tblSettings.

The Unipro golden tests STILL use PI=3.14 (legacy and normal both), because the
Unipro reference workbook uses 3.14. All Lodha-specific tests use PI=22/7.

Normal mode authoritative total under PI=3.14:  178,805,450 (prefer) or 178,801,050
Normal mode authoritative total under PI=22/7:  178,815,550 (not a target; Lodha is geometry-only)
Legacy mode authoritative total (PI=3.14):      178,657,900

## HOW TO START

Tell Claude Code / Roo Code:
```
Read CLAUDE_CODE_MASTER_PROMPT.md fully. Check DOCS\PROGRESS.md for current state.
Audit the current workbook, fix any golden-test failures, and continue from
the first incomplete phase. Log everything in DOCS\PROGRESS.md.
```

---

## 1. WHAT YOU ARE BUILDING

A fully automated Windows Excel macro workbook (Mega_Pipes_Quotation_DEV.xlsm)
for Mega Pipes Pvt Ltd. Sales staff fill a structured form and get a complete
commercial quotation: live calculation, nearest plant by road distance,
competitor comparison, cost sheet, offer letter, synopsis, PDF and Excel exports.

Current state: P0-P30 completed and tagged. The VBA self-test reports
178,801,050 for normal mode. The authoritative target is 178,805,450.
The 4,400 Rs gap = 1 Rs/m on Line 4 (IntPerM borderline at PI=3.14).
Both values are accepted; 178,805,450 is preferred per the master prompt.

---

## 2. OPERATING RULES

1. Work only in DEV. Never touch LIVE. Backup to BACKUP\ before every phase.
2. Export embedded VBA before importing. Report any drift.
3. After every VBA import: compile in real Excel via COM, then run a trivial macro.
4. Build loop: write VBA -> import -> compile -> test -> fix -> repeat (max 5 attempts; then NEEDS_HUMAN).
5. Never write PASS without saving actual test output in TESTS\.
6. All constants come from tblSettings -- never hardcode in VBA.
7. Money and weight use CDec() in VBA. Never Double for money.
8. ORS key only from Environ("ORS_API_KEY"). Never print, log or save it.
9. Public gTestMode As Boolean. In test mode, ShowMsg/AskInput write to TESTS\ui_log.txt.
10. Git commit and tag after every verified phase.
11. Keep DOCS\PROGRESS.md updated after every step.

---

## 3. FORMULA ENGINE

### 3.1 VBA helpers

```vba
Function RoundUpD(x As Variant, Optional n As Long = 0) As Variant
    Dim q As Variant, t As Variant
    q = CDec(10) ^ (-n)
    t = CDec(x) / q
    If t = Int(t) Then
        RoundUpD = t * q
    Else
        RoundUpD = (Int(t) + CDec(1)) * q
    End If
End Function

Function MassKgPerM(od As Variant, wt As Variant) As Variant
    MassKgPerM = (CDec(od) - CDec(wt)) * CDec(wt) * CDec("0.0246615")
End Function

Function MassMtPerM(od As Variant, wt As Variant) As Variant
    MassMtPerM = MassKgPerM(od, wt) / CDec(1000)
End Function
```

### 3.2 Geometry and basis rules

OD basis: OD_mm = typed; ID_mm = OD_mm - 2*WT (no rounding)
ID basis: ID_mm = typed; OD_mm = RoundUpD(ID_mm + 2*WT, 0)
InnerD = OD_mm - 2*WT (always the actual bore, both bases)

### 3.3 Mass
kgm = (OD_mm - WT) * WT * 0.0246615
mtm = kgm / 1000

### 3.4 Coating area (synopsis)
PI_Coating = CDec("3.142857142857143")  ' = 22/7; from tblSettings
ext_m2 = RoundUpD(OD_mm * metres * PI_Coating / 1000, 0)
int_m2 = RoundUpD(InnerD * metres * PI_Coating / 1000, 0)

### 3.5 Pre-cost per MT
Pre = (steelRate + gradeExtra) * (1 + yieldPct/100) + conversion + testing + margin

### 3.6 Coating cost per MT
PI used for pricing = CDec(rates.PI) from tblSettings (3.14 for Unipro tests, 22/7 for production)
extCoat = extRate * PI * (OD_mm / 1000) / mtm
intCoat = intRate * PI * (InnerD / 1000) / mtm   ' normal mode
' Legacy only: intCoat uses OD_mm not InnerD

### 3.7 Freight engine
pipesPerTrailer = from tblStacking by OD range
trailers_normal = RoundUpD(pipes_int / pipesPerTrailer, 0)   ' whole trailers
trailers_legacy = CDec(pipes_int) / CDec(pipesPerTrailer)    ' fractional
freightPerMT = RoundUpD(TripRate * trailers / weightMT, 0) + Handling

### 3.8 Financial cost (algebraic, no circular reference)
BaseCost = Pre + extCoat + intCoat + freightPerMT
k  = (finPct / 100) * (1 + GST / 100)
FC = k * BaseCost / (1 - k)
' Legacy mode: FC = ManualFC typed per line

### 3.9 Selling price
BarePerM    = RoundUpD((Pre + FC) * mtm, 0)
ExtPerM     = RoundUpD(extCoat    * mtm, 0)
IntPerM     = RoundUpD(intCoat    * mtm, 0)
FreightPerM = RoundUpD(freightPerMT * mtm, 0)
TotalPerM   = BarePerM + ExtPerM + IntPerM + FreightPerM
LineValue   = TotalPerM * metres
GrandTotal  = SUM(LineValue) * (1 + GST/100)

---

## 4. GOLDEN TEST VALUES

### 4.1 Unipro (PI=3.14, ID basis, WT 8.4, Steel 61000, Yield 2%, Conv 5000, Ext 500, Int 275)

LEGACY (trip 98000, FC typed, fractional trailers, int uses OD):
ID=1226/OD=1243: 22,681/m  -> 45,362,000
ID=1300/OD=1317: 23,914/m  ->  8,369,900
ID=1350/OD=1367: 24,748/m  -> 12,374,000
ID=1400/OD=1417: 25,580/m  -> 112,552,000
TOTAL: 178,657,900

NORMAL (trip=125.32x782=98000.24, FC algebraic, whole trailers, int uses InnerD):
ID=1226/OD=1243: 42 trailers, 22,689/m  -> 45,378,000
ID=1300/OD=1317:  8 trailers, 24,109/m  ->  8,438,150
ID=1350/OD=1367: 11 trailers, 24,857/m  -> 12,428,500
ID=1400/OD=1417: 92 trailers, 25,582/m  -> 112,560,800
TOTAL: 178,805,450  (178,801,050 also accepted; 1 Rs/m borderline on Line 4)

### 4.2 Lodha (geometry and coating only, PI=22/7)

OD=1050/WT=12: kgm=307.1836, ext_m2=1205, int_m2=1177 (qty=365m)
OD= 900/WT=12: kgm=262.7929, ext_m2= 340, int_m2= 331 (qty=100m)
OD= 700/WT=12: kgm=203.6053, ext_m2= 176, int_m2= 170 (qty=55m)
OD=1000/WT=12: kgm=292.3867, ext_m2= 315, int_m2= 307 (qty=100m)
ID=1050/WT=12: OD=1074, kgm=314.2862, ext_m2=84504, int_m2=82616 (qty=25035m)

---

## 5. UI LAYOUT (QUOTE sheet, 9-column grid A-I)

Row 1: MEGA PIPES banner (A1:I1, UI_GREEN_DARK)
Rows 3-18: Quotation Header (A-C) | Live Summary (D-E) | Command Center (F-I)
Row 19: Section labels
Rows 20-30: Input form (A-C) | Line items table (D-I, up to 10 rows)
Rows 31-37: Rate inputs (A-C)
Below: Add/Edit/Remove/Clear buttons (A-C)

No floating shapes over cells. Buttons in column I anchored per row.
Totals row immediately after last active line.

---

## 6. 16-COLUMN SYNOPSIS

A=Sr, B=ID mm, C=OD inch, D=OD mm, E=WT mm, F=Grade, G=Type,
H=Qty m, I=Pipe count, J=kg/m, K=Qty MT, L=Volume m3, M=CBM ratio,
N=Ext m2, O=Int m2, P=12m pipe wt MT

PI for N,O = 22/7 always (Lodha spec)
OD inch: 2dp in grid, integer on offer

---

## 7. COLOUR PALETTE

UI_GREEN_DARK = #206021 (banners, main buttons)
UI_GREEN_MID  = #3C8C3F (subheaders, Add Item)
UI_NAVY       = #102C60 (Executive card banners)
UI_BLUE       = #4682B4 (Edit Line, Open Quote)
UI_TEAL       = #008080 (Synopsis preview)
UI_ORANGE     = #B06C00 (edit mode, custom geometry warning)
UI_RED        = #AA3C32 (Remove Line, critical validation)
GRAND_TOTAL   = #FFF3CD fill / #783C00 text

---

## 8. KEY SETTINGS (tblSettings)

GST_pct=18, PI_Coating=3.142857142857143, PI_Legacy=3.14
FreightMethod=B, RatePerKm=125.32, TripRateManual=98000
Handling_Rs_per_MT=100, TrailerPayloadMT=28
LegacyMode=No, MaxLinesPerQuote=10
DefSteel=61000, DefYield=2, DefConv=5000, DefFin=1.5
ValidityDays=5, EnquirySerialStart=595

---

## 9. CURRENT STATUS AND NEXT STEPS

P0-P30 completed. Git tag: e1b201c (latest).

Key open items per master prompt:
1. Normal golden: VBA self-test shows 178,801,050; target 178,805,450.
   Gap = 1 Rs/m on Line 4 IntPerM (PI borderline). Both accepted.
2. PI_Coating in tblSettings should be updated to 22/7 = 3.142857142857143
   for production. Unipro golden tests must continue to use PI=3.14 explicitly.
3. Lodha geometry golden tests need to be added to the Python test suite.
4. The CLAUDE_CODE_MASTER_PROMPT.md (this file) is now saved in the project root.

To continue: run the Python test suite (pytest TESTS/), check for failures,
then update PI_Coating in tblSettings if not already 22/7.

---

## 10. THINGS ONLY A HUMAN CAN DO

Real coating rates (3LP thickness, Zinc Rich Epoxy); plant exact addresses;
ORS API key (set ORS_API_KEY=...); passwords; shared folder path and permissions;
testing with multiple real users; visual layout approval; training; go-live decision.

---

## 11. VBA MODULES

`modCompare.bas` builds the COMPARE sheet:
1. Reads our selected plant and its km from the current quote.
2. Reads all competitor distances from tblDistanceCache.
3. Reads ex-mill prices from tblCompPrices (matching product family).
4. Sorts by km ascending.
5. Applies the row colour rules at runtime (never hardcoded row numbers).
6. Calculates the recommended offer.
7. Writes all values only (no formulas in the COMPARE sheet).

The COMPARE sheet is refreshed automatically when:
- The distance lookup is completed (Find Distance button).
- A competitor ex-mill price is entered or changed.
- The selected plant is changed.

---

## 12. OUTPUTS

### COMPARE sheet (full competitor intelligence, always visible)

The COMPARE sheet has three sections built by modCompare.bas:

**Section A — Our plants:** one row per eligible plant showing road km, freight/MT,
trips needed and km difference vs the selected plant. Our plants: A9D18E light green.
The selected plant has a ★ badge.

**Section B — Competitor distance table:** one row per competitor plant, sorted by
road km ascending. Columns: Sr.No., Full plant name, City/State, Road km (bold),
km vs our nearest (+ means farther; − means closer in red font), Freight/MT Rs
(using our RatePerKm), Freight diff vs ours Rs, Ex-mill price (user-entered or
pre-filled from tblCompPrices with date badge), Landed price Rs/MT, Landed price
diff vs ours Rs, Status, Distance source.

**Row colour rules (runtime, data-driven, never hardcode row numbers):**
- Our plants: A9D18E light green
- Competitor km > our nearest km: no fill (white)
- Competitor km < our nearest km, diff 1–300 km: FFFF00 yellow, black bold (moderate threat)
- Competitor km < our nearest km, diff > 300 km: FF6B6B light red, 800000 dark red bold (major threat)
- km diff column: C00000 dark red font for any competitor closer than us

**Section C — Recommended offer summary:**
- Cheapest competitor landed price (from those with a price entered)
- Undercut amount (from Settings)
- Recommended offer = cheapest landed − undercut
- Our floor price = full cost build-up + MinMarginRsPerMT
- Status: ✓ safe / ⚠ cannot beat without going below floor
- If floor > recommended: warning and the floor price as the minimum viable offer

**Mechanics:**
- km values: from ONE ORS matrix call that fetches all plant distances in parallel.
  Results cached by (plantId, custLat4dp, custLon4dp) in tblDistanceCache.
  "Refresh distance" button clears the cache for the current customer location.
- Ex-mill prices: entered by clicking the cell in column H. Stored in tblCompPrices
  with competitor name, product family and date. Pre-filled for returning competitors.
  "STALE" badge (orange) after StaleDays.
- The table refreshes automatically when: distance is updated, a price is entered,
  or the selected plant changes.
- Values only, no formulas in the sheet. Every row colour set by modCompare.bas.

