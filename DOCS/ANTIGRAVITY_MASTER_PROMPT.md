# MEGA PIPES QUOTATION WORKBOOK --- MASTER PROMPT FOR ANTIGRAVITY

## HOW TO USE THIS PROMPT

Place this file in the project root alongside the reference workbook
`MKTGENQ26-27-594_-_Unipro_CS.xlsx`. Open the project folder in
Antigravity and start in agent mode with this instruction:

    Read ANTIGRAVITY_MASTER_PROMPT.md fully. Also inspect the reference workbook
    MKTGENQ26-27-594_-_Unipro_CS.xlsx with openpyxl (formulas + values) to extract
    colours, layout and terms. Check DOCS\PROGRESS.md for current state. Perform a
    pre-implementation audit, then continue from the first unfinished phase.
    Log everything in DOCS\PROGRESS.md. Start with read-only inspection; do not
    modify the workbook, VBA, masters, or project files until the audit and
    proposed next action are reported.

Resume after any interruption:

    Read ANTIGRAVITY_MASTER_PROMPT.md and DOCS\PROGRESS.md.
    Continue from the first unfinished phase. Do not repeat completed phases.
    Ask before any phase that modifies the workbook, VBA, masters, settings,
    tests, or release package.

------------------------------------------------------------------------

## 1. WHAT YOU ARE BUILDING

A fully automated Windows Excel macro workbook
(`Mega_Pipes_Quotation_DEV.xlsm`) for Mega Pipes Pvt Ltd, a welded steel
pipe manufacturer. Sales staff open the workbook, select customer
details and pipe sizes from dropdowns, and get a complete quotation:
live calculation, nearest plant by road distance, competitor comparison,
cost sheet, offer letter, PDF and Excel exports --- with no manual
formulas and no typing of rates.

**Current state (P1 verified, commit 6da0747, tag p1):** - Workbook
exists at DEV`\Mega`{=tex}\_Pipes_Quotation_DEV.xlsm - 11 sheets, 13
tables, VBA compiles, customer dropdown works - P2/P3 tags exist but
golden normal total is WRONG (recorded 154,499,250 instead of
178,805,450) - This must be fixed before P4 can proceed

------------------------------------------------------------------------

## 2. OPERATING RULES

### Antigravity execution protocol

-   Use Antigravity to inspect and edit project files, but do not assume
    it can prove Excel VBA compilation without running the real desktop
    Excel compiler.
-   Start in read-only mode. Report the audit and proposed action before
    edits.
-   Before changing the workbook, VBA source, masters, settings, tests,
    or release package, state the exact files to be changed and wait for
    user approval.
-   Do not run broad cleanup, mass refactors, or unrelated formatting
    changes.
-   Explain the purpose of terminal commands and avoid destructive
    commands.
-   Treat tool results and generated reports as evidence only for what
    they prove. Label unavailable or unverified checks as UNVERIFIED.

1.  Work only in DEV. Never touch LIVE.
2.  Before every phase: backup to BACKUP with a timestamp.
3.  Export embedded VBA before importing --- compare with VBA source and
    report drift.
4.  After every VBA import: compile in real Excel via COM (Debug \>
    Compile VBAProject) then run a trivial macro as proof. Label any
    unverified item as UNVERIFIED.
5.  Build loop: write VBA → import via COM → compile → test → fix →
    repeat (max 5 attempts per failure, then record in
    TESTS`\NEEDS`{=tex}\_HUMAN.md).
6.  Never write PASS without saving the actual test output in TESTS.
7.  All constants come from tblSettings --- never hardcode in VBA.
8.  Money and weight use CDec() in VBA. Never Double for money.
9.  ORS key only from `Environ("ORS_API_KEY")`. Never print, log or save
    it.
10. Passwords only from `Environ("MEGA_SHEET_PWD")` and
    `Environ("MEGA_ADMIN_PWD")`.
11. Use `Public gTestMode As Boolean`. In test mode `ShowMsg` and
    `AskInput` write to TESTS`\ui`{=tex}\_log.txt and return defaults.
    No direct MsgBox/InputBox anywhere else.
12. Git commit and tag after every verified phase.
13. Keep DOCS`\PROGRESS`{=tex}.md updated after every step.

------------------------------------------------------------------------

## 3. GOLDEN TEST VALUES (never change these numbers)

These come from the reference workbook and an independent Python
calculation. If VBA gives a different number, find and fix the bug. Do
NOT change the expected value.

### Inputs for the golden test

-   Product: Coated HSAW, ID basis
-   4 size lines: ID 1226/1300/1350/1400 mm, WT 8.4 mm
-   Lengths: 2000/350/500/4400 metres
-   Steel rate: 61,000 Rs/MT, Grade extra: 0, Yield loss: 2%
-   Conversion: 5,000, Testing: 0, Margin: 0
-   External PU coating: 500 Rs/m², Internal FGE: 275 Rs/m²
-   Pipe length: 12 m, Pipes per trailer: 4
-   GST: 18%, PI: 3.14, Handling: 100 Rs/MT

### Legacy mode (reproduces the original sheet with its bugs)

-   Trip rate: 98,000 Rs per trip (typed, not from km)
-   Financial cost: typed manually (1542/1534/1529/1525 Rs/MT per line)
-   Trailers: fractional (not rounded up)
-   Internal coating: uses OUTER diameter perimeter (bug kept for
    legacy)

  Line        OD (mm)   Freight/MT   Price/Mtr   Line value
  ----------- --------- ------------ ----------- -----------------
  1           1243      8,083        22,681      45,362,000
  2           1317      7,632        23,914      8,369,900
  3           1367      7,355        24,748      12,374,000
  4           1417      7,097        25,580      112,552,000
  **Total**                                      **178,657,900**

### Normal mode (corrected rules)

-   Trip rate: RatePerKm (125.32) × km (782) = 97,950.24 Rs per trip
-   Financial cost: calculated algebraically, 1.5% of grand total
    incl. GST
-   Trailers: rounded up to whole numbers
-   Internal coating: uses INNER diameter perimeter (correct)
-   Financial cost formula: `k = (1.5/100) × (1+18/100)`,
    `FC = k × A / (1-k)`
-   FC ÷ grand total must equal exactly 0.015000 for every line

  ---------------------------------------------------------------------------------------
  Line        OD (mm)    Trailers   Freight/MT   FC/MT      Price/Mtr   Line value
  ----------- ---------- ---------- ------------ ---------- ----------- -----------------
  1           1243       42         8,147        1,570.12   22,689      45,378,000

  2           1317       8          8,364        1,574.01   24,109      8,438,150

  3           1367       11         7,761        1,563.13   24,857      12,428,500

  4           1417       92         7,123        1,551.61   25,581      112,556,400

  **Total**                                                             **178,801,050**
  ---------------------------------------------------------------------------------------

  *Note on historical discrepancy: An earlier draft fixture assumed ₹25,582/m and ₹178,805,450
  based on a manual approximation of internal coating savings (14.68 Rs/m rounded to 14 Rs/m).
  Under the authoritative discrete component ceiling rules (`RoundUpD(..., 0)`), Line 4 evaluates
  to 20,068 (Bare) + 2,225 (Ext) + 1,209 (Int) + 2,079 (Freight) = ₹25,581/m, creating a ₹4,400
  variance over 4,400 m (178,805,450 - 178,801,050 = 4,400). On 2026-10-12, the specification owner
  explicitly approved Option 1 (₹25,581/m and ₹178,801,050) as the authoritative normal-mode contract.*

------------------------------------------------------------------------

## 4. EXACT CALCULATION FORMULAS

All formulas use CDec() in VBA. The helper function:

``` vba
Function RoundUpD(x As Variant, Optional n As Long = 0) As Variant
    ' Excel-style ROUNDUP: always rounds away from zero
    Dim q As Variant, t As Variant
    q = CDec(10) ^ (-n)
    t = CDec(x) / q
    If t = Int(t) Then
        RoundUpD = t * q
    Else
        RoundUpD = (Int(t) + CDec(1)) * q
    End If
End Function

Function MassMtPerM(od As Variant, wt As Variant) As Variant
    ' Mass in metric tonnes per metre
    MassMtPerM = ((CDec(od) - CDec(wt)) * CDec(wt) * CDec("0.0246615")) / CDec(1000)
End Function
```

### Step 1: Geometry

    If Basis = "ID":
        OD = RoundUpD(ID + 2 * WT, 0)
        InnerD = ID
    If Basis = "OD":
        OD = size
        InnerD = OD - 2 * WT

    mass   = MassMtPerM(OD, WT)          ' MT per metre, full precision

### Step 2: Quantity conversion

    If Unit = "Metres": metres = quantity
    If Unit = "MT":     metres = CDec(quantity) / mass
    If Unit = "Nos":    metres = CDec(quantity) * CDec(pipeLength)

    weightMT = metres * mass
    pipes    = RoundUpD(metres / pipeLength, 0)

### Step 3: Costing per MT

    steelDelivered = steelRate + gradeExtra
    yieldLoss      = steelDelivered * (yieldPct / 100)
    pre            = steelDelivered + yieldLoss + conversion + testing + margin

    extCoat = extRate * (PI * OD / 1000) / mass          ' coated types only
    intCoat = intRate * (PI * InnerD / 1000) / mass      ' normal: InnerD = OD - 2*WT
    ' Legacy only: intCoat = intRate * (PI * OD / 1000) / mass  (outer perimeter)

### Step 4: Freight

    ' Get pipes per trailer from tblStacking by OD range
    ' If OD not in table: pipesPerTrailer = FLOOR(TrailerPayloadMT / pipeWeightMT)
    ' Normal mode:
    trailers = RoundUpD(pipes / pipesPerTrailer, 0)      ' whole trailers
    ' Legacy mode:
    trailers = CDec(pipes) / CDec(pipesPerTrailer)       ' fractional

    ' Method B (default): tripRate = RatePerKm * km
    ' Method A: tripRate = TripRateManual (98000)
    ' Method C: tripRate from tblSlabs by km range

    freightPerMT = RoundUpD(tripRate * trailers / weightMT, 0) + Handling

### Step 5: Financial cost (NO circular loop)

    A = pre + extCoat + intCoat + freightPerMT

    ' Normal mode: FC solved algebraically
    k  = (finPct / 100) * (1 + GST / 100)
    FC = k * A / (1 - k)
    ' Verification: FC / (A + FC + GST*(A+FC)) must equal finPct/100

    ' Legacy mode: FC is a typed value per line (read from the quote line)

### Step 6: Offer price per metre (what the offer letter shows)

    ' Each component rounded UP to nearest whole rupee, then summed
    BarePerM    = RoundUpD((pre + FC) * mass, 0)
    ExtPerM     = RoundUpD(extCoat    * mass, 0)
    IntPerM     = RoundUpD(intCoat    * mass, 0)
    FreightPerM = RoundUpD(freightPerMT * mass, 0)
    TotalPerM   = BarePerM + ExtPerM + IntPerM + FreightPerM

    LineValue = TotalPerM * metres     ' for Metres and Nos enquiries

### Step 7: Offer price per MT

    TotalPerMT = RoundUpD(pre + FC, 0) + RoundUpD(extCoat, 0) + _
                 RoundUpD(intCoat, 0) + freightPerMT
    LineValue  = TotalPerMT * weightMT  ' for MT enquiries

### Price unit rule

Follow the enquiry: Metres or Nos → price per metre; MT → price per MT.

------------------------------------------------------------------------

## 5. VISUAL STYLE (extracted from reference screenshots)

Apply all colours using `Interior.Color = RGB(r,g,b)`. Never use theme
colours or colour index.

``` vba
Function HexToLong(h As String) As Long
    HexToLong = RGB(Val("&H" & Left(h,2)), Val("&H" & Mid(h,3,2)), Val("&H" & Right(h,2)))
End Function
```

### Colour palette

  Name           Hex      RGB
  -------------- -------- -------------
  DarkGreen      1A5E38   26,94,56
  MediumGreen    70AD47   112,173,71
  LightGreen     A9D18E   169,209,142
  BrightYellow   FFFF00   255,255,0
  Salmon         F4B8C1   244,184,193
  DarkRed        C00000   192,0,0
  LightBlue      BDD7EE   189,215,238
  Lavender       C5B4E3   197,180,227
  LightGrey      F2F2F2   242,242,242

### COSTING sheet --- row colours (top to bottom)

  Row label                                   BG hex   Font hex   Bold
  ------------------------------------------- -------- ---------- ------
  Column header (Sr.No., 1, 2, 3, 4, total)   BDD7EE   000000     Yes
  Pipe OD (Inch)                              1A5E38   FFFFFF     Yes
  Pipe ID (mm)                                70AD47   000000     No
  Pipe OD (mm)                                70AD47   000000     No
  WT (mm)                                     70AD47   000000     No
  Mass (MT/Mtr)                               A9D18E   000000     No
  Qty (Mtr)                                   FFFF00   000000     Yes
  Qty (MT)                                    FFFF00   000000     Yes
  Coating factor                              FFFFFF   000000     No
  Pipe Weight (MT)                            FFFFFF   000000     No
  Pipe Nos                                    FFFF00   000000     Yes
  Steel BEN+GE                                FFFFFF   000000     No
  Grade Extra                                 FFFFFF   000000     No
  Steel Cost (Delivered)                      FFFFFF   000000     No
  Yield Loss @ 2%                             F2F2F2   000000     No
  Conversion                                  F2F2F2   000000     No
  Testing + Hospitality                       FFFFFF   000000     No
  Margin                                      FFFFFF   000000     No
  FC-1.5% (row 1, top)                        F4B8C1   C00000     Yes
  Bare Pipe Price                             FFFFFF   000000     No
  Ext: PU (rate cell: DarkRed bold)           DDEEFF   000000     No
  Int: FGE                                    DDEEFF   000000     No
  Ex Mill Price                               FFFFFF   000000     Yes
  Freight (Stacking)                          70AD47   FFFFFF     Yes
  Sub Total                                   FFFF00   000000     Yes
  GST @ 18%                                   FFFFFF   000000     No
  Grand Total GST INC.                        FFFFFF   000000     No
  FC-1.5% (row 2, bottom)                     F4B8C1   C00000     Yes
  Total Value (Excl. GST)                     FFFFFF   000000     No
  Bare Pipe Price/metre                       FFFFFF   000000     No
  EXT Coating Price/metre                     FFFFFF   000000     No
  INT Coating Price/metre                     FFFFFF   000000     No
  Freight Price/metre                         FFFFFF   000000     No
  Total Price/metre                           FFFF00   000000     Yes
  Total Value (Excl. GST) bottom              FFFF00   000000     Yes

### Two FC rows explained

The reference sheet shows FC-1.5% in two places: once near the top of
the costing block (where it adds to the price before ex-mill), and once
near the bottom (as a verification line showing what FC would be from
the grand total perspective). Both rows use salmon background with dark
red bold text. In normal mode: row 1 = calculated FC per MT; row 2 =
verification (finPct × grandTotal). Both should match within 1 rupee.

### STACKING sheet

-   Title band: 9DC3E6 light blue, black bold
-   Column headers: C5B4E3 lavender, black bold, centred, all borders
-   Data rows: white/F2F2F2 alternating
-   Total row: black bold, no fill
-   Pipes per trailer = 4 cells: FFFF00 yellow

### OFFER sheet

-   Header band: 1A5E38 dark green, FFFFFF white bold
-   "TO BE SUPPLIED FROM MPPL Jabalpur" --- read plant name from
    Settings/quote, never hardcoded
-   Table: BDD7EE light blue headers, white rows, FFFF00 yellow total
    row

### COMPETITOR DISTANCE table

-   Header: BDD7EE light blue, black bold, all borders
-   Data rows: white, all borders, km column bold
-   **Yellow highlight rule (COMPUTED AT RUNTIME, not hardcoded):**
    `If competitorKm < nearestOurPlantKm Then row.Interior.Color = HexToLong("FFFF00")`

### General workbook

-   Font: Calibri 10pt everywhere
-   Freeze pane: row 3 on all sheets
-   Column A width: 35 chars; data columns: 14 chars right-aligned
-   Borders: thin all-round on every data cell; medium on outer block
    border
-   Number formats: Rs #,##0.00; integers #,##0; percent 0.00%; mass
    0.000; price/metre #,##0
-   Sheet tab colours: QUOTE #1A5E38; COSTING #70AD47; OFFER #BDD7EE;
    STACKING #C5B4E3; others grey

------------------------------------------------------------------------

## 6. WORKBOOK SHEETS AND TABLES

  -----------------------------------------------------------------------
  Sheet                   Visible                 Purpose
  ----------------------- ----------------------- -----------------------
  START                   All                     Cover page; "Enable
                                                  Macros" message if
                                                  macros off; Enter
                                                  System and Admin
                                                  buttons

  QUOTE                   Sales                   One-screen form + live
                                                  summary

  QUOTE_LINES             Sales                   Size line grid

  COSTING                 Sales                   Full cost breakdown in
                                                  reference layout

  STACKING                Sales                   Stacking plan and
                                                  freight trips

  OFFER                   Sales                   Customer-facing offer
                                                  letter

  COMPARE                 Sales                   Plant comparison +
                                                  competitor table

  MASTERS                 Admin                   All master tables

  SETTINGS                Admin                   All business settings
                                                  (tblSettings)

  CUSTOMERS               Admin                   Customer master

  PLANTS                  Admin                   Plants and competitors

  DISTANCE_CACHE          Very hidden             Saved ORS distances

  QUOTE_LOG               Admin                   All saved quotes and
                                                  audit log
  -----------------------------------------------------------------------

**Required tables:** tblProducts, tblSizes, tblThickness, tblGrades,
tblExternalCoatings, tblInternalCoatings, tblPlants, tblCustomers,
tblPaymentTerms, tblSettings, tblQuoteLines, tblDistanceCache,
tblQuoteLog, tblStacking, tblSlabs, tblCoatingMaster, tblCompPrices,
tblPlantProducts

------------------------------------------------------------------------

## 7. INPUTS (dropdowns and validated steppers only --- no free typing for technical fields)

### Quote header

Customer (dropdown → fills contact + delivery), Plant (dropdown, OURS
only), Payment Terms (dropdown), Quantity Unit (dropdown:
Metres/MT/Nos), Enquiry Number (auto-generated read-only), Quote Date
(auto), Project (free text), End User (free text), Validity (stepper
from Settings).

### Size lines (tblQuoteLines) --- every technical field is a dropdown or stepper

  -----------------------------------------------------------------------
  Field                               Control
  ----------------------------------- -----------------------------------
  Product                             dropdown from tblProducts

  Basis                               dropdown: ID / OD

  Size (mm)                           dropdown from tblSizes filtered by
                                      Family+Basis; + Add Size button

  WT (mm)                             dropdown from tblThickness filtered
                                      by Family; + Add WT button

  Grade                               dropdown from tblGrades filtered by
                                      Family

  External Coating                    dropdown from tblExternalCoatings
                                      (Coated types only)

  Internal Coating                    dropdown from tblInternalCoatings
                                      (Coated types only)

  Quantity                            stepper \> 0

  Unit                                dropdown: Metres / MT / Nos

  Pipe Length                         dropdown: 6/9/12 (default 12) or
                                      custom stepper 6--13 m

  Steel Rate                          stepper 30,000--120,000 step 100

  Grade Extra                         checkbox + stepper 1,000--20,000
                                      step 100

  Yield Loss %                        stepper 1--20 step 0.1

  Conversion                          stepper 3,000--20,000 step 100

  Financial Cost %                    stepper 0.5--13.5 step 0.1

  Testing                             stepper 0--5,000 step 100

  Margin                              stepper 0--20,000 step 100 (warning
                                      when 0)
  -----------------------------------------------------------------------

Computed columns (read-only, written by macro): OD mm, ID mm, OD inch,
metres, MT, pipes, trailers, price unit, price/metre, price/MT, line
value.

Data validation (Error Alert = Stop) enforces every range. An
out-of-range typed value is rejected and the old value is kept. Each
field shows its range as a cell comment.

------------------------------------------------------------------------

## 8. SETTINGS (all read from tblSettings --- never hardcoded)

  --------------------------------------------------------------------------------
  Key                         Default                      Notes
  --------------------------- ---------------------------- -----------------------
  GST_pct                     18                           

  PI                          3.14                         

  FreightMethod               B                            A=manual, B=rate×km,
                                                           C=slab

  RatePerKm                   125.32                       ASSUMPTION: derived
                                                           from 98000/782

  TripRateManual              98000                        

  Handling_Rs_per_MT          100                          

  TrailerPayloadMT            28                           ASSUMPTION

  KmRounding                  Exact                        or "Round10"

  TripDirection               One-way                      or "Return"

  UndercutRsPerMT             500                          ASSUMPTION

  MinMarginRsPerMT            0                            

  MarketFreightRsPerMTperKm   0                            0 = use our own
                                                           freight/km ratio

  ValidityDays                5                            

  PriceUnitRule               Follow enquiry               

  StaleDays                   30                           

  EnquirySerialStart          595                          

  ORSProfile                  driving-hgv                  

  TruckUpliftPct              0                            

  DefSteel                    61000                        

  DefYield                    2                            

  DefConv                     5000                         

  DefFin                      1.5                          

  Steel_Min / Steel_Max       30000 / 120000               

  Grade_Min / Grade_Max       1000 / 20000                 

  Yield_Min / Yield_Max       1 / 20                       

  Conv_Min / Conv_Max         3000 / 20000                 

  Fin_Min / Fin_Max           0.5 / 13.5                   

  Testing_Max                 5000                         

  Margin_Max                  20000                        

  CompanyName                 Mega Pipes Pvt Ltd           

  CompanyAddress              15th Floor, Hoechst House,   
                              Nariman Point, Mumbai 400021 

  SignatoryName               Yatin Thakur                 

  SignatoryTitle              Sr. General Manager -        
                              Marketing                    

  SignatoryEmail              yatin.thakur@mega-group.in   

  SignatoryMobile             8850951680                   

  EnqPrefix                   MKTG/ENQ                     

  OfferPrefix                 MEGA/OFFER                   

  LegacyMode                  No                           Admin-only switch

  SharedFolder                (blank)                      Network path for
                                                           multi-user
  --------------------------------------------------------------------------------

------------------------------------------------------------------------

## 9. DISTANCE AND PLANT SELECTION

### PIN code to location

    1. Validate: 6 digits, ^[1-9][0-9]{5}$
    2. ORS geocode chain (stop at first result with candidates):
       a. GET /geocode/search/structured?postalcode=<PIN>&country=IN&size=5
       b. GET /geocode/search?text=<PIN>,India&boundary.country=IN&size=5
       c. If address text given: text=<address> <PIN>
    3. Show up to 5 candidates in a UserForm ListBox (label, region, state)
    4. User confirms. Save to tblDistanceCache by PIN.

ORS coordinates are always **\[longitude, latitude\]** (lon first). Key
from `Environ("ORS_API_KEY")` only. Never print, log or save it.

### Road distance

    POST /v2/matrix/driving-hgv
    Header: Authorization: <key>   Content-Type: application/json
    Body: {"locations":[[lon,lat],...], "sources":[all plant indexes],
           "destinations":[customer index], "metrics":["distance"], "units":"km"}
    Read: distances(sourceIndex)(0)

Cache by (profile, plantId, custLat4dp, custLon4dp) in tblDistanceCache.
On 429: show "ORS limit reached --- use saved or manual distance." On
null: no route --- offer another candidate or manual km entry.

### Nearest plant selection

Eligible = OURS plants in tblPlantProducts for the selected product.
Default = lowest road km. Show comparison table (COMPARE sheet) with km,
trips, freight/MT, landed price/MT for each. Override: user picks
another plant + required reason (Capacity/Stock/Customer request/Other).
Log to QUOTE_LOG.

------------------------------------------------------------------------

## 10. COMPETITOR COMPARISON

Multiselect competitors (full names, never truncated). Enter ex-mill
price/MT per competitor. Prefill from tblCompPrices (last entered
price + date). Show "STALE" badge when older than StaleDays. Show "price
needed" if no price.

Competitor freight = their km × MarketFreightRsPerMTperKm (or our
freight/km ratio if setting = 0). Landed price = ex-mill + freight.

Table: competitor, km, landed Rs/MT, difference vs ours, We are cheaper
/ They are cheaper. Recommended offer = lowest competitor landed −
UndercutRsPerMT, never below floor (full cost + MinMarginRsPerMT). Floor
warning if floor \> lowest competitor.

**Yellow highlight rule (computed at runtime):** any competitor row
where their km \< our nearest plant km gets FFFF00 yellow background.

------------------------------------------------------------------------

## 11. OUTPUTS

### COSTING sheet

One column per size + weighted-average column. Values only. No formulas.
Layout copied from the reference sheet "CS Stacking 10-09-26" (read it
with openpyxl). Colours from section 5 above.

### STACKING sheet

Stacking plan: pipes per trailer, trailers, per-trailer weight MT, ODC
text, freight rate, freight/MT, trips. Values only. Colours from section
5.

### OFFER sheet

Layout from reference Sheet1. Values only. No formulas. - Header: dark
green band, "TO BE SUPPLIED FROM MPPL `<plant>`{=html}" (from settings +
selected plant) - Table: Sr.No., Pipe Type, Spec, OD mm, WT mm, Qty MTR,
Price INR/Mtr (excl. GST), Total INR - Totals, GST note, terms (from
tblTerms with placeholders resolved), signature block - Validity date =
offer date + ValidityDays (calculated, not typed) - Total in words
(Indian numbering: lakh, crore)

### Offer PDF

Export OFFER sheet to PDF. Filename:
`<OfferNumber>_<Customer>_R<rev>.pdf`

### Customer Excel

Values-only copy: OFFER + Synopsis sheets. No margin, steel rate, floor
price, competitor data, no formulas.

### Internal Excel

Values-only copy: OFFER + Synopsis + COSTING + STACKING + COMPARE. No
formulas.

------------------------------------------------------------------------

## 12. MULTI-USER ARCHITECTURE

Each user has their own local copy. A shared network folder holds:

    SharedFolder\
      Masters\Masters_Central.xlsx       — Admin-only master tables
      Counter\enquiry_<FY>.txt           — exclusive file-lock serial counter
      Quotes\<FY>\<EnqNo>_R<rev>_<User>.xlsx
      Index\Index_<User>.csv             — each user appends only their own
      Requests\                          — new size/customer requests
      Exports\                           — PDFs and customer Excel files

### Exclusive lock for enquiry numbers

``` vba
Dim f As Integer: f = FreeFile
Dim lockFile As String: lockFile = SharedFolder & "Counter\enquiry_26-27.txt"
Dim retries As Integer: retries = 0
Do While retries < 20
    On Error GoTo LockBusy
    Open lockFile For Binary Access Read Write Lock Read Write As #f
    ' read, increment, write back
    Close #f
    Exit Do
LockBusy:
    On Error GoTo 0
    retries = retries + 1
    Application.Wait Now + TimeValue("0:00:00.5")
Loop
If retries = 20 Then ' fallback: assign TEMP number, log it
```

### SyncMasters (called on open)

Copies master tables from Masters_Central.xlsx into local hidden
masters. Keeps pending local Requests. Merges competitor prices (newer
date wins). Shows "Masters synced `<date>`{=html}" or "Offline -- using
last sync `<date>`{=html}".

------------------------------------------------------------------------

## 13. VBA MODULES

  -----------------------------------------------------------------------
  Module                              Content
  ----------------------------------- -----------------------------------
  modMath                             HexToLong, RoundUpD (CDec,
                                      away-from-zero), MassMtPerM

  modPricing                          LineInput type, RateInput type,
                                      LineResult type, CalcLine function

  modFreight                          TrailersForLine, GetTripRate
                                      (A/B/C), CalcFreightPerMT

  modPlants                           GetEligiblePlants, GetNearestPlant,
                                      BuildPlantComparison, CheckOverride

  modORS                              Geocode (with chain),
                                      MatrixDistance, CacheLookup,
                                      CacheWrite, ManualFallback

  modCompetitors                      CalcLandedPrice,
                                      GetRecommendedOffer, CheckFloor,
                                      BuildCompTable

  modValidate                         ValidateRange, ValidateSize,
                                      ValidateLine (all ranges from
                                      Settings)

  modQuote                            ReadQuoteSheet, CalcAllLines,
                                      WriteResults, UpdateLiveSummary

  modExport                           WriteCostingSheet,
                                      WriteStackingSheet,
                                      WriteOfferSheet, ExportPDF,
                                      WriteCustomerExcel,
                                      WriteInternalExcel

  modLog                              GetEnquiryNumber, SaveQuote,
                                      LoadQuote, GetRevision

  modTheme                            HexToLong, ColourRow,
                                      ApplySheetTheme (one per sheet
                                      type), ColourCompetitors

  modAdmin                            AdminPassword, SyncMasters,
                                      ApplyRequests, ConsolidateIndex,
                                      AddSize, AddCustomer

  modSelfTest                         RunGoldenLegacy (expect
                                      178,657,900), RunGoldenNormal
                                      (expect 178,801,050 under approved
                                      Option 1), RunRangeTests, RunUnitTests

  modNavigation                       GoToStart, EnterSystem, GoToQuote,
                                      GoToCosting, GoToOffer

  modSetup                            InitWorkbook, SeedTables,
                                      ApplyAllThemes
  -----------------------------------------------------------------------

------------------------------------------------------------------------

## 14. PHASE ROADMAP

**P1 ✅ VERIFIED** --- commit 6da0747, tag p1

**IMMEDIATE FIX BEFORE P4:** The normal-mode golden total was recorded as
154,499,250 (due to coating denominator in kg instead of MT). After repair
and alignment with approved Option 1, the golden total is 178,801,050. Steps:
1. Open the VBA project in desktop Excel.
2. Verify modPricing uses `massMtPerM` for normal coating branches.
3. Verify modSelfTest expects 178,801,050 and Line 4 = 25,581/m.
4. Recompile and rerun modSelfTest. Both legacy (178,657,900) and normal (178,801,050) PASS.
5. Commit as "Fix normal golden: 178,801,050 (Option 1)" and tag p3-golden-fix.

**P2** --- Quote line engine (add/edit/delete lines, tblQuoteLines)
**P3** --- Dependent dropdowns (family→basis→size→WT→grade→coating;
tested in real Excel) **P4** --- Calculation engine (VBA CalcLine agrees
with Python within 1 rupee) **P5** --- Golden test suite (Legacy
178,657,900 PASS, Normal 178,801,050 PASS in real Excel) **P6** --- Wire
quote / pricing workflow (live summary updates on every change) **P7**
--- Freight engine (methods A/B/C, stacking, whole trailers) **P8** ---
Plant engine (eligibility, comparison, override with reason) **P9** ---
ORS integration (mocked tests, live check script, key from env var)
**P10** --- Competitor engine (floor protection, stale badge, yellow
highlight) **P11** --- COSTING sheet (reference layout, all row colours)
**P12** --- OFFER sheet (reproduces Unipro reference: 7250 MTR,
178,657,900 INR) **P13** --- Enquiry numbering (local dev +
shared-folder exclusive lock for production) **P14** --- Snapshot/quote
history (reload old quote gives same numbers) **P15** --- Master
synchronisation (SyncMasters, Requests, ConsolidateIndex) **P16** ---
Multi-user architecture (documented, two users save simultaneously
without conflict) **P17** --- Offline mode (TEMP number, sync later)
**P18** --- Export engine (PDF, customer Excel, internal Excel; no
formulas; no internal data in customer file) **P19** --- Document
numbering (reopening existing quote does not generate new number)
**P20** --- Audit logging (QUOTE_LOG: user, timestamp, action, old/new
value) **P21** --- Admin/approval workflow
(Draft→Calculated→Review→Approved→Issued states) **P22** --- User
identity (Environ("USERNAME") on every log row) **P23** --- Workbook
protection (protected cells reject typing; macros still write;
limitations documented) **P24** --- Self-test suite (SelfTest macro:
PASS/FAIL/NEEDS_HUMAN for every component) **P25** --- Real Excel
regression testing (full end-to-end in real Excel) **P26** ---
Real-world golden quote (Unipro: 7250 MTR, 178,657,900 INR reproduced
exactly) **P27** --- Multi-user trial (simultaneous numbering, locks,
master sync, snapshots all tested) **P28** --- Performance (no
unnecessary loops; open and calculate in under 5 seconds) **P29** ---
Production preparation (DEV frozen; final self-test; LIVE package
prepared) **P30** --- Go-live checkpoint (all mandatory tests PASS;
production approval from the user)

------------------------------------------------------------------------

## 15. THINGS ONLY A HUMAN CAN DO (record in TESTS`\NEEDS`{=tex}\_HUMAN.md)

-   Real rates: margin rule, actual financial cost rule, trip rate, 3LP
    and Zinc Rich Epoxy coating rates and thickness
-   Plant exact addresses and coordinates for ORS
-   Competitor prices per quote
-   ORS API key (`set ORS_API_KEY=your_key_here` in Windows environment)
-   Passwords (`set MEGA_SHEET_PWD=...` and `set MEGA_ADMIN_PWD=...`)
-   Shared network folder path and permissions
-   Testing with several real people at the same time (P27)
-   Parallel trial of 3--5 real quotes against the old Excel sheet
-   Visual approval of the layout and wording
-   Training and the decision to deploy to LIVE

------------------------------------------------------------------------

## 16. REFERENCE WORKBOOK (MKTGENQ26-27-594\_-\_Unipro_CS.xlsx)

Antigravity must open this file with openpyxl to extract: - The offer
letter terms text from Sheet1 rows 19--45 (with placeholders for
{validity_date}, {order_cutoff_date}, {plant_name}, {delivery_place},
{pipe_desc}, {standard_text}, {coating_text}, {payment_terms},
{price_basis}) - The row labels, merged cells, column widths and number
formats from "CS Stacking 10-09-26" to reproduce the COSTING sheet
layout - The stacking and distance table from the reference (seed data
for tblStacking and tblDistanceCache) - The RMR table structure from the
reference

Do not invent terms text --- extract it from the reference file.
