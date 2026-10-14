Attribute VB_Name = "modUI"
Option Explicit

' ============================================================
' Mega Pipes Quotation System
' Module: modUI
' Purpose:
'   Complete Redesigned Guided Quotation User Interface
'   Provides commercial estimators with an intuitive, modern,
'   high-contrast, spacious console.
'   No navigation or manual formulas required.
'   Full editing freedom with unlocked inputs and zero password blocks.
' ============================================================

' UI Color Palette (Mega Pipes brand)
Public Const UI_GREEN_DARK  As Long = 2169120   ' RGB(32, 96, 33)   - Primary brand green
Public Const UI_GREEN_MID   As Long = 4166204   ' RGB(60, 140, 63)  - Mid accent green
Public Const UI_GREEN_LIGHT As Long = 13166093  ' RGB(205, 230, 200) - Soft green
Public Const UI_BLUE        As Long = 11829830  ' RGB(70, 130, 180) - Action blue
Public Const UI_NAVY        As Long = 6296336   ' RGB(16, 44, 96)   - Executive navy
Public Const UI_TEAL        As Long = 8421376   ' RGB(0, 128, 128)  - Synopsis teal
Public Const UI_ORANGE      As Long = 27760     ' RGB(176, 108, 0)  - Calc amber
Public Const UI_RED         As Long = 255       ' RGB(255, 0, 0)    - Alert red
Public Const UI_GRAY_DARK   As Long = 5263440   ' RGB(80, 80, 80)   - Secondary gray
Public Const UI_WHITE       As Long = 16777215  ' RGB(255, 255, 255)
Public Const UI_YELLOW      As Long = 65535     ' RGB(255, 255, 0)

' Backend sheets hidden from normal users
Private Const BACKEND_SHEETS As String = "QUOTE_LINES,COSTING,COMPARE,MASTERS,SETTINGS,CUSTOMERS,PLANTS,DISTANCE_CACHE,QUOTE_LOG"

' ============================================================
' 1. WORKBOOK OPEN & LIFECYCLE
' ============================================================

Public Sub OnWorkbookOpen()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Hide internal backend calculation sheets
    HideBackendSheets
    
    ' Ensure no sheet protection prevents editing
    modProtection.UnprotectAllSheets
    
    ' Build and refresh the START screen
    BuildStartScreen
    
    ' Configure guided quotation console on QUOTE
    SetupQuoteSheetUI
    
    ' Activate START dashboard
    ThisWorkbook.Worksheets("START").Activate
    
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Dim saveErrDesc As String
    saveErrDesc = Err.Description
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    SafeMsgBox "Mega Pipes Quotation System initialization: " & saveErrDesc, vbInformation
End Sub

' ============================================================
' 2. START SCREEN BUILDER
' ============================================================

Public Sub BuildStartScreen()
    Dim ws As Worksheet
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean
    
    On Error GoTo ErrorHandler
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Set ws = ThisWorkbook.Worksheets("START")
    modProtection.UnprotectSheet ws
    
    ' Clear all existing cells & shapes to completely wipe any prior ghost text
    ws.Cells.Clear
    Dim sIdx As Long
    For sIdx = ws.Shapes.Count To 1 Step -1
        On Error Resume Next
        ws.Shapes(sIdx).Delete
        On Error GoTo 0
    Next sIdx
    
    ' Page setup for perfect 1-page display and print
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintArea = "$A$1:$G$35"
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
    
    ws.Range("A1:G35").Interior.Color = RGB(248, 250, 248)
    ws.Range("A1:G35").Font.Name = "Calibri"
    
    ' Column widths
    ws.Columns("A").ColumnWidth = 4
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 22
    ws.Columns("F").ColumnWidth = 22
    ws.Columns("G").ColumnWidth = 4
    
    ' Title Banner (B2:F3)
    ws.Range("B2:F3").Merge
    With ws.Range("B2")
        .Value = "MEGA PIPES - COMMERCIAL ESTIMATION SYSTEM"
        .Interior.Color = UI_GREEN_DARK
        .Font.Color = UI_WHITE
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Subtitle (B4:F4)
    ws.Range("B4:F4").Merge
    With ws.Range("B4")
        .Value = "Automated Welded Steel Pipe Quotation Engine | Version 2.0"
        .Interior.Color = UI_GREEN_MID
        .Font.Color = UI_WHITE
        .Font.Size = 9.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Status Console (B6:F12)
    Dim qNum As String, qCust As String, qTotal As String, qMode As String
    qNum = Trim$(CStr(ThisWorkbook.Worksheets("QUOTE").Range("B4").Value))
    If Len(qNum) = 0 Then qNum = "None (Click NEW QUOTATION)"
    qCust = Trim$(CStr(ThisWorkbook.Worksheets("QUOTE").Range("B8").Value))
    If Len(qCust) = 0 Then qCust = "Not Selected"
    qTotal = "Rs. " & Format$(Val(ThisWorkbook.Worksheets("QUOTE").Range("E10").Value), "#,##0")
    qMode = Trim$(CStr(ThisWorkbook.Worksheets("QUOTE").Range("E12").Value))
    If Len(qMode) = 0 Then qMode = "NORMAL (Option 1)"
    
    With ws.Range("B6:F6")
        .Merge
        .Value = "QUOTATION STATUS DASHBOARD"
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = UI_GREEN_DARK
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Range("B8:F12").Interior.Color = RGB(255, 255, 255)
    ws.Range("B8:F12").Font.Size = 9.5
    
    ws.Range("B8").Value = "Active Reference:"
    ws.Range("C8:D8").Merge: ws.Range("C8").Value = qNum: ws.Range("C8").Font.Bold = True
    ws.Range("E8").Value = "Customer:"
    ws.Range("F8").Value = qCust: ws.Range("F8").Font.Bold = True
    
    ws.Range("B10").Value = "Current Grand Total:"
    ws.Range("C10:D10").Merge: ws.Range("C10").Value = qTotal: ws.Range("C10").Font.Bold = True: ws.Range("C10").Font.Color = UI_ORANGE
    ws.Range("E10").Value = "Pricing Mode:"
    ws.Range("F10").Value = qMode
    
    With ws.Range("B8:F12").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 205, 180)
    End With
    
    ' Section Title (B14:F14)
    With ws.Range("B14:F14")
        .Merge
        .Value = "QUOTATION WORKFLOW ACTIONS"
        .Font.Bold = True
        .Font.Size = 10.5
        .Font.Color = UI_GREEN_DARK
        .HorizontalAlignment = xlCenter
    End With
    
    ' Action Buttons on START
    AddModernButton ws, "btnStartNew", "NEW QUOTATION", "modUI.UI_NewQuotation", _
                    ws.Range("B16"), ws.Range("F16"), UI_GREEN_MID, UI_WHITE, True
                    
    AddModernButton ws, "btnStartOpen", "OPEN EXISTING QUOTATION", "modUI.UI_OpenQuotation", _
                    ws.Range("B18"), ws.Range("C18"), UI_BLUE, UI_WHITE, False
                    
    AddModernButton ws, "btnStartHistory", "QUOTATION HISTORY", "modUI.UI_QuotationHistory", _
                    ws.Range("D18"), ws.Range("F18"), RGB(60, 90, 140), UI_WHITE, False
                    
    AddModernButton ws, "btnStartPreview", "PREVIEW OFFER LETTER", "modUI.UI_PreviewOffer", _
                    ws.Range("B20"), ws.Range("C20"), RGB(90, 110, 130), UI_WHITE, False
                    
    AddModernButton ws, "btnStartSynopsis", "PREVIEW SYNOPSIS REPORT", "modUI.UI_PreviewSynopsis", _
                    ws.Range("D20"), ws.Range("F20"), UI_TEAL, UI_WHITE, False
                    
    AddModernButton ws, "btnStartPDF", "EXPORT OFFER (PDF)", "modUI.UI_ExportPDF", _
                    ws.Range("B22"), ws.Range("C22"), RGB(160, 60, 40), UI_WHITE, False
                    
    AddModernButton ws, "btnStartExcel", "EXPORT CUSTOMER EXCEL", "modUI.UI_ExportCustomerExcel", _
                    ws.Range("D22"), ws.Range("F22"), RGB(0, 110, 160), UI_WHITE, False
                    
    AddModernButton ws, "btnStartAdmin", "ADMIN / BACK OFFICE ACCESS", "modUI.UI_AdminUnlock", _
                    ws.Range("B24"), ws.Range("F24"), UI_GRAY_DARK, UI_WHITE, False
                    
    ' Notification Banner
    With ws.Range("B26:F28")
        .Merge
        .Value = "System Status: Online | Master data verified | Quotation engine ready." & vbCrLf & _
                 "To create or update a quotation, click 'NEW QUOTATION' or 'OPEN EXISTING QUOTATION' above."
        .Interior.Color = RGB(240, 245, 240)
        .Font.Size = 8.5
        .Font.Color = RGB(60, 80, 60)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("B26:F28").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(200, 220, 200)
    End With
    
    ' Footer
    With ws.Range("B30:F30")
        .Merge
        .Value = "Mega Pipes Quotation System | Confidential & Proprietary | All Rights Reserved"
        .Font.Size = 7.5
        .Font.Color = RGB(150, 150, 150)
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Range("A1:G35").Locked = False
    
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub

ErrorHandler:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
End Sub

' ============================================================
' 3. GUIDED QUOTATION CONSOLE (QUOTE SHEET)
' ============================================================

Public Sub SetupQuoteSheetUI()
    Dim ws As Worksheet
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean
    Dim rIdx As Long
    
    On Error GoTo ErrorHandler
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Set ws = ThisWorkbook.Worksheets("QUOTE")
    modProtection.UnprotectSheet ws
    
    ' Clear all existing cells & shapes to completely wipe prior ghost blocks and old tables
    ws.Cells.Clear
    Dim sIdx As Long
    For sIdx = ws.Shapes.Count To 1 Step -1
        On Error Resume Next
        ws.Shapes(sIdx).Delete
        On Error GoTo 0
    Next sIdx
    
    ' Configure page setup for crisp single-screen / 1-page landscape fit
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintArea = "$A$1:$I$52"
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.3)
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
    
    ' ── 1. Unified 9-Column Grid (Total Width = 140) ──
    ws.Columns("A").ColumnWidth = 18    ' Line ID / Left card labels
    ws.Columns("B").ColumnWidth = 28    ' Specification / Left card inputs
    ws.Columns("C").ColumnWidth = 22    ' Coating (Ext/Int) / Left card inputs merged B:C
    ws.Columns("D").ColumnWidth = 18    ' Table Quantity / Summary labels
    ws.Columns("E").ColumnWidth = 14    ' Table Pipes / Summary values
    ws.Columns("F").ColumnWidth = 14    ' Table Weight MT / Right card panel
    ws.Columns("G").ColumnWidth = 16    ' Table Unit Price / Right card panel
    ws.Columns("H").ColumnWidth = 18    ' Table Line Total / Right card panel
    ws.Columns("I").ColumnWidth = 12    ' Table Status / Right card panel
    
    ' ── 2. Row Heights ──
    ws.Rows(1).RowHeight = 26           ' Top brand title
    ws.Rows(2).RowHeight = 8            ' Spacer
    ws.Rows(3).RowHeight = 24           ' Block 1 card headers
    For rIdx = 4 To 17
        ws.Rows(rIdx).RowHeight = 21    ' Block 1 card content
    Next rIdx
    ws.Rows(18).RowHeight = 10          ' Spacer
    ws.Rows(19).RowHeight = 24          ' Block 2 card headers
    For rIdx = 20 To 33
        ws.Rows(rIdx).RowHeight = 21    ' Block 2 inputs and button tiers
    Next rIdx
    ws.Rows(34).RowHeight = 10          ' Spacer
    ws.Rows(35).RowHeight = 26          ' Validation feedback banner
    ws.Rows(36).RowHeight = 10          ' Spacer
    ws.Rows(37).RowHeight = 24          ' Section 3: Table title header
    ws.Rows(38).RowHeight = 22          ' Table column titles
    
    ' ── Top Brand Title Banner (A1:I1) ──
    ws.Range("A1:I1").Merge
    With ws.Range("A1")
        .Value = "MEGA PIPES - COMMERCIAL ESTIMATION SYSTEM"
        .Interior.Color = UI_GREEN_DARK
        .Font.Color = UI_WHITE
        .Font.Bold = True
        .Font.Size = 13
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' ── 3. Block 1 Left Card: Quotation Header Details (Rows 3-17) ──
    ws.Range("A3:C3").Merge
    With ws.Range("A3")
        .Value = "1. QUOTATION HEADER DETAILS"
        .Interior.Color = UI_GREEN_DARK
        .Font.Color = UI_WHITE
        .Font.Bold = True
        .Font.Size = 10.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range("A4").Value = "Quotation Number"
    ws.Range("A5").Value = "Enquiry Date"
    ws.Range("A6").Value = "Type of Enquiry"
    ws.Range("A7").Value = "Product Type"
    ws.Range("A8").Value = "Customer"
    ws.Range("A9").Value = "Contact Person"
    ws.Range("A10").Value = "Delivery Place"
    ws.Range("A11").Value = "Project / Tender Ref"
    ws.Range("A12").Value = "Salesperson"
    ws.Range("A13").Value = "Supplying Plant"
    ws.Range("A14").Value = "Payment Terms"
    ws.Range("A15").Value = "Validity (Days)"
    ws.Range("A16").Value = "Price Basis"
    ws.Range("A17").Value = "Pricing Mode"
    
    With ws.Range("A4:A17")
        .Font.Name = "Calibri"
        .Font.Size = 9.5
        .Font.Bold = True
        .Font.Color = RGB(30, 45, 30)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    For rIdx = 4 To 17
        ws.Range("A" & rIdx).Interior.Color = IIf(rIdx Mod 2 = 0, RGB(245, 248, 245), RGB(255, 255, 255))
        ws.Range("B" & rIdx & ":C" & rIdx).Merge
        ws.Range("B" & rIdx).Interior.Color = RGB(255, 255, 255)
        ws.Range("B" & rIdx).Font.Name = "Calibri"
        ws.Range("B" & rIdx).Font.Size = 9.5
        ws.Range("B" & rIdx).HorizontalAlignment = xlLeft
        ws.Range("B" & rIdx).VerticalAlignment = xlCenter
    Next rIdx
    
    With ws.Range("A3:C17").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 210, 180)
    End With
    
    ws.Range("B4").Interior.Color = RGB(242, 245, 248)
    ws.Range("B4").Font.Bold = True
    ws.Range("B4").Font.Color = UI_NAVY
    
    ' ── 4. Block 1 Middle Card: Live Quotation Summary Console (D3:E17) ──
    ws.Range("D3:E3").Merge
    With ws.Range("D3")
        .Value = "LIVE QUOTATION SUMMARY"
        .Interior.Color = RGB(26, 82, 45)
        .Font.Color = UI_WHITE
        .Font.Bold = True
        .Font.Size = 10.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range("D4").Value = "Total Metres"
    ws.Range("D5").Value = "Total Weight (MT)"
    ws.Range("D6").Value = "Total Pipe Count"
    ws.Range("D7").Value = "Total Trailers"
    ws.Range("D8").Value = "Subtotal (Excl. GST)"
    ws.Range("D9").Value = "GST @ 18%"
    ws.Range("D10").Value = "Grand Total (Incl. GST)"
    ws.Range("D11").Value = "Workflow Status"
    ws.Range("D12").Value = "Mode"
    ws.Range("D13").Value = "Supplying Plant"
    ws.Range("D14").Value = "Valid Lines"
    
    With ws.Range("D4:D14")
        .Font.Name = "Calibri"
        .Font.Size = 9.5
        .Font.Bold = False
        .Font.Color = RGB(40, 40, 40)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    For rIdx = 4 To 14
        ws.Range("D" & rIdx & ":E" & rIdx).Interior.Color = IIf(rIdx Mod 2 = 0, RGB(250, 252, 250), RGB(255, 255, 255))
        ws.Range("E" & rIdx).Font.Name = "Calibri"
        ws.Range("E" & rIdx).Font.Size = 9.5
        ws.Range("E" & rIdx).HorizontalAlignment = xlRight
        ws.Range("E" & rIdx).VerticalAlignment = xlCenter
    Next rIdx
    
    ws.Range("D8:E8").Font.Bold = True
    ws.Range("D8:E8").Interior.Color = RGB(242, 248, 242)
    ws.Range("D10:E10").Font.Bold = True
    ws.Range("D10:E10").Font.Size = 11
    ws.Range("D10:E10").Interior.Color = RGB(255, 243, 205)
    ws.Range("D10:E10").Font.Color = RGB(120, 60, 0)
    
    ws.Range("E11").HorizontalAlignment = xlCenter
    ws.Range("E12").HorizontalAlignment = xlCenter
    ws.Range("E14").NumberFormat = "@"
    ws.Range("E14").HorizontalAlignment = xlCenter
    
    ws.Range("D15:E17").Merge
    With ws.Range("D15")
        .Value = "COMMERCIAL STATUS:" & vbCrLf & "QUOTATION CONSOLE READY"
        .Font.Name = "Calibri"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = UI_GREEN_DARK
        .Interior.Color = RGB(238, 248, 238)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    With ws.Range("D3:E17").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(190, 215, 190)
    End With
    
    ' ── 5. Block 1 Right Card: Quotation Capacity & Guidelines (F3:I17) ──
    ws.Range("F3:I3").Merge
    With ws.Range("F3")
        .Value = "QUOTATION CAPACITY & STATUS"
        .Interior.Color = RGB(26, 82, 45)
        .Font.Color = UI_WHITE
        .Font.Bold = True
        .Font.Size = 10.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range("F4:I8").Merge
    With ws.Range("F4")
        .Value = "QUOTATION CAPACITY" & vbCrLf & "0 of 10 sizes added" & vbCrLf & "(10 sizes remaining)"
        .Font.Name = "Calibri"
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(80, 100, 80)
        .Interior.Color = RGB(245, 248, 245)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    ws.Range("F9:I13").Merge
    With ws.Range("F9")
        .Value = "ESTIMATION GUIDELINES:" & vbCrLf & _
                 "- Up to 10 pipe sizes supported per quote." & vbCrLf & _
                 "- Click 'CALCULATE QUOTATION' after modifying lines." & vbCrLf & _
                 "- Review Offer & Synopsis before PDF export."
        .Font.Name = "Calibri"
        .Font.Size = 8.5
        .Font.Color = RGB(60, 75, 60)
        .Interior.Color = RGB(252, 254, 252)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    ws.Range("F14:I17").Merge
    With ws.Range("F14")
        .Value = "SYSTEM CAPABILITIES:" & vbCrLf & _
                 "- Multi-plant freight optimization (P01 / P02)" & vbCrLf & _
                 "- Full support for custom diameters & wall thickness"
        .Font.Name = "Calibri"
        .Font.Size = 8.5
        .Font.Color = RGB(70, 70, 70)
        .Interior.Color = RGB(248, 250, 248)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    With ws.Range("F3:I17").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(190, 215, 190)
    End With
    
    ' ── 6. Block 2 Left Card: Pipe Requirement Input Box (Rows 19-33) ──
    ws.Range("A19:C19").Merge
    With ws.Range("A19")
        .Value = "2. ADD / CONFIGURE PIPE LINE ITEM"
        .Interior.Color = UI_GREEN_DARK
        .Font.Color = UI_WHITE
        .Font.Bold = True
        .Font.Size = 10.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range("A20").Value = "Product Standard"
    ws.Range("A21").Value = "Pipe Family"
    ws.Range("A22").Value = "Diameter Basis"
    ws.Range("A23").Value = "Size (mm)"
    ws.Range("A24").Value = "Wall Thickness (mm)"
    ws.Range("A25").Value = "Steel Grade"
    ws.Range("A26").Value = "External Coating"
    ws.Range("A27").Value = "Internal Coating"
    ws.Range("A28").Value = "Quantity"
    ws.Range("A29").Value = "Quantity Unit"
    ws.Range("A30").Value = "Pipe Length (m)"
    
    With ws.Range("A20:A30")
        .Font.Name = "Calibri"
        .Font.Size = 9.5
        .Font.Bold = True
        .Font.Color = RGB(30, 45, 30)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    For rIdx = 20 To 30
        ws.Range("A" & rIdx).Interior.Color = IIf(rIdx Mod 2 = 0, RGB(248, 250, 248), RGB(255, 255, 255))
        ws.Range("B" & rIdx & ":C" & rIdx).Merge
        ws.Range("B" & rIdx).Interior.Color = RGB(255, 255, 255)
        ws.Range("B" & rIdx).Font.Name = "Calibri"
        ws.Range("B" & rIdx).Font.Size = 9.5
        ws.Range("B" & rIdx).HorizontalAlignment = xlLeft
        ws.Range("B" & rIdx).VerticalAlignment = xlCenter
    Next rIdx
    
    ws.Range("A31:C33").Merge
    With ws.Range("A31")
        .Value = "Tip: Select Family & Basis to populate standard sizes." & vbCrLf & _
                 "For non-standard OD/ID or WT, click 'CUSTOM PIPE SIZE'."
        .Font.Name = "Calibri"
        .Font.Size = 8.5
        .Font.Italic = True
        .Font.Color = RGB(100, 100, 100)
        .Interior.Color = RGB(250, 250, 250)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    With ws.Range("A19:C33").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 210, 180)
    End With
    
    ' ── 7. Block 2 Right Card: Command Center & Actions (Rows 19-33) ──
    ws.Range("D19:I19").Merge
    With ws.Range("D19")
        .Value = "QUOTATION COMMAND CENTER & ACTIONS"
        .Interior.Color = UI_GREEN_DARK
        .Font.Color = UI_WHITE
        .Font.Bold = True
        .Font.Size = 10.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range("D20:I33").Interior.Color = RGB(250, 252, 250)
    With ws.Range("D19:I33").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 210, 180)
    End With
    
    ' Tier 1 (Rows 20-21): Line Item Operations
    AddModernButton ws, "btnQuoteAdd", "+ ADD ITEM TO QUOTE", "modUI.UI_AddItem", _
                    ws.Range("D20"), ws.Range("E21"), UI_GREEN_MID, UI_WHITE, True
    AddModernButton ws, "btnQuoteDuplicate", "DUPLICATE ITEM", "modUI.UI_DuplicateItem", _
                    ws.Range("F20"), ws.Range("G21"), RGB(40, 120, 90), UI_WHITE, True
    AddModernButton ws, "btnQuoteEdit", "EDIT SELECTED LINE", "modUI.UI_EditItem", _
                    ws.Range("H20"), ws.Range("I21"), UI_BLUE, UI_WHITE, True
                    
    ' Tier 2 (Rows 23-24): Line Item Cleaners
    AddModernButton ws, "btnQuoteRemove", "- REMOVE LINE", "modUI.UI_RemoveItem", _
                    ws.Range("D23"), ws.Range("E24"), RGB(170, 60, 50), UI_WHITE, True
    AddModernButton ws, "btnQuoteClearLine", "CLEAR LINE INPUTS", "modUI.UI_ClearLine", _
                    ws.Range("F23"), ws.Range("G24"), RGB(110, 120, 130), UI_WHITE, False
    AddModernButton ws, "btnQuoteClear", "CLEAR ALL LINES", "modUI.UI_ClearAllItems", _
                    ws.Range("H23"), ws.Range("I24"), RGB(90, 90, 90), UI_WHITE, False
                    
    ' Tier 3 (Rows 26-27): Quotation & Calculation
    AddModernButton ws, "btnQuoteCalc", "CALCULATE QUOTATION", "modUI.UI_CalculateQuotation", _
                    ws.Range("D26"), ws.Range("E27"), UI_GREEN_DARK, UI_WHITE, True
    AddModernButton ws, "btnQuoteSave", "SAVE DRAFT", "modUI.UI_SaveDraft", _
                    ws.Range("F26"), ws.Range("G27"), UI_NAVY, UI_WHITE, True
    AddModernButton ws, "btnQuoteReset", "RESET DRAFT", "modUI.UI_ResetDraft", _
                    ws.Range("H26"), ws.Range("I27"), RGB(140, 80, 50), UI_WHITE, False
                    
    ' Tier 4 (Rows 29-30): Document Previews & Custom Sizes
    AddModernButton ws, "btnQuotePreview", "PREVIEW OFFER", "modUI.UI_PreviewOffer", _
                    ws.Range("D29"), ws.Range("E30"), RGB(70, 100, 130), UI_WHITE, True
    AddModernButton ws, "btnQuoteSynopsis", "PREVIEW SYNOPSIS", "modUI.UI_PreviewSynopsis", _
                    ws.Range("F29"), ws.Range("G30"), UI_TEAL, UI_WHITE, True
    AddModernButton ws, "btnQuoteCustomSize", "CUSTOM PIPE SIZE", "modUI.UI_CustomSize", _
                    ws.Range("H29"), ws.Range("I30"), RGB(160, 120, 20), UI_WHITE, True
                    
    ' Tier 5 (Rows 32-33): Exports & Navigation
    AddModernButton ws, "btnQuotePDF", "EXPORT PDF", "modUI.UI_ExportPDF", _
                    ws.Range("D32"), ws.Range("E33"), RGB(160, 45, 35), UI_WHITE, True
    AddModernButton ws, "btnQuoteExcel", "EXPORT EXCEL", "modUI.UI_ExportCustomerExcel", _
                    ws.Range("F32"), ws.Range("G33"), RGB(20, 110, 160), UI_WHITE, True
    AddModernButton ws, "btnQuoteHome", "<- BACK TO START", "modUI.UI_BackToHome", _
                    ws.Range("H32"), ws.Range("I33"), UI_GRAY_DARK, UI_WHITE, True
                    
    ' ── 8. Validation Feedback Banner at Row 35 ──
    ws.Range("A35:I35").Merge
    With ws.Range("A35")
        .Value = "Ready. Select specifications above and click '+ ADD ITEM TO QUOTE'."
        .Interior.Color = RGB(245, 248, 245)
        .Font.Size = 9.5
        .Font.Color = UI_GREEN_DARK
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("A35:I35").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 210, 180)
    End With
    
    ' ── 9. Block 3: Quotation Line Items Table Header (Rows 37-38) ──
    ws.Range("A37:I37").Merge
    With ws.Range("A37")
        .Value = "3. CURRENT QUOTATION LINE ITEMS"
        .Interior.Color = UI_GREEN_DARK
        .Font.Color = UI_WHITE
        .Font.Bold = True
        .Font.Size = 10.5
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range("A38").Value = "Line ID"
    ws.Range("B38").Value = "Pipe Specification"
    ws.Range("C38").Value = "Coating (Ext / Int)"
    ws.Range("D38").Value = "Quantity"
    ws.Range("E38").Value = "Pipes"
    ws.Range("F38").Value = "Weight (MT)"
    ws.Range("G38").Value = "Unit Price"
    ws.Range("H38").Value = "Line Total"
    ws.Range("I38").Value = "Status"
    
    With ws.Range("A38:I38")
        .Interior.Color = UI_GREEN_MID
        .Font.Color = UI_WHITE
        .Font.Bold = True
        .Font.Size = 9.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Populate active lines display
    RefreshQuoteLinesDisplay
    
    ' Set master data validations
    modMasterData.SetupMasterDrivenHeader
    modDependentDropdowns.SetupDependentDropdownArchitecture
    
    ' Ensure all input cells are unlocked
    ws.Range("A1:I120").Locked = False
    
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub

ErrorHandler:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
End Sub

' ============================================================
' 4. QUOTATION LINES DISPLAY TABLE REFRESH
' ============================================================

Public Sub RefreshQuoteLinesDisplay()
    Dim wsQ As Worksheet
    Dim wsLines As Worksheet
    Dim tblLines As ListObject
    Dim rObj As ListRow
    Dim outRow As Long
    
    On Error GoTo ErrorHandler
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    modProtection.UnprotectSheet wsQ
    
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    Set tblLines = wsLines.ListObjects("tblQuoteLines")
    
    ' Clear previous display rows (39 to 75)
    wsQ.Range("A39:I75").UnMerge
    wsQ.Range("A39:I75").ClearContents
    wsQ.Range("A39:I75").Interior.Color = xlNone
    wsQ.Range("A39:I75").Borders.LineStyle = xlNone
    
    outRow = 39
    If tblLines.ListRows.Count > 0 Then
        For Each rObj In tblLines.ListRows
            Dim pId As String
            pId = Trim$(CStr(rObj.Range.Cells(1, tblLines.ListColumns("ProductID").Index).Value))
            If Len(pId) > 0 Then
                Dim fam As String, bas As String, sz As Double, wt As Double, grd As String
                Dim qty As Double, unt As String, pcs As Long, mt As Double, prc As Double, tot As Double
                Dim extC As String, intC As String, st As String
                
                fam = CStr(rObj.Range.Cells(1, tblLines.ListColumns("Family").Index).Value)
                bas = CStr(rObj.Range.Cells(1, tblLines.ListColumns("Basis").Index).Value)
                sz = Val(rObj.Range.Cells(1, tblLines.ListColumns("SizeMM").Index).Value)
                wt = Val(rObj.Range.Cells(1, tblLines.ListColumns("ThicknessMM").Index).Value)
                grd = CStr(rObj.Range.Cells(1, tblLines.ListColumns("GradeID").Index).Value)
                qty = Val(rObj.Range.Cells(1, tblLines.ListColumns("Quantity").Index).Value)
                unt = CStr(rObj.Range.Cells(1, tblLines.ListColumns("QuantityUnit").Index).Value)
                pcs = CLng(Val(rObj.Range.Cells(1, tblLines.ListColumns("PipeCount").Index).Value))
                mt = Val(rObj.Range.Cells(1, tblLines.ListColumns("WeightMT").Index).Value)
                prc = Val(rObj.Range.Cells(1, tblLines.ListColumns("UnitPrice").Index).Value)
                tot = Val(rObj.Range.Cells(1, tblLines.ListColumns("LineTotal").Index).Value)
                st = CStr(rObj.Range.Cells(1, tblLines.ListColumns("Status").Index).Value)
                
                extC = CStr(rObj.Range.Cells(1, tblLines.ListColumns("CoatingExternalID").Index).Value)
                intC = CStr(rObj.Range.Cells(1, tblLines.ListColumns("CoatingInternalID").Index).Value)
                If Len(extC) = 0 Then extC = "Bare"
                If Len(intC) = 0 Then intC = "None"
                
                wsQ.Rows(outRow).RowHeight = 22
                wsQ.Cells(outRow, 1).Value = rObj.Range.Cells(1, tblLines.ListColumns("LineID").Index).Value
                wsQ.Cells(outRow, 2).Value = fam & " " & sz & "mm (" & bas & ") x " & wt & "mm " & grd
                wsQ.Cells(outRow, 3).Value = extC & " / " & intC
                wsQ.Cells(outRow, 4).Value = Format$(qty, "#,##0") & " " & unt
                wsQ.Cells(outRow, 5).Value = pcs
                wsQ.Cells(outRow, 6).Value = mt
                wsQ.Cells(outRow, 6).NumberFormat = "#,##0.00"
                wsQ.Cells(outRow, 7).Value = prc
                wsQ.Cells(outRow, 7).NumberFormat = "[$₹-4009]#,##0"
                wsQ.Cells(outRow, 8).Value = tot
                wsQ.Cells(outRow, 8).NumberFormat = "[$₹-4009]#,##0"
                wsQ.Cells(outRow, 9).Value = st
                
                ' Row styling
                With wsQ.Range(wsQ.Cells(outRow, 1), wsQ.Cells(outRow, 9))
                    .Font.Name = "Calibri"
                    .Font.Size = 9.5
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                    .Borders.Color = RGB(210, 225, 210)
                    .Interior.Color = IIf(outRow Mod 2 = 0, RGB(248, 252, 248), RGB(255, 255, 255))
                End With
                wsQ.Cells(outRow, 1).HorizontalAlignment = xlCenter
                wsQ.Cells(outRow, 2).HorizontalAlignment = xlLeft
                wsQ.Cells(outRow, 3).HorizontalAlignment = xlCenter
                wsQ.Cells(outRow, 4).HorizontalAlignment = xlRight
                wsQ.Cells(outRow, 5).HorizontalAlignment = xlRight
                wsQ.Cells(outRow, 6).HorizontalAlignment = xlRight
                wsQ.Cells(outRow, 7).HorizontalAlignment = xlRight
                wsQ.Cells(outRow, 8).HorizontalAlignment = xlRight
                wsQ.Cells(outRow, 9).HorizontalAlignment = xlCenter
                
                outRow = outRow + 1
            End If
        Next rObj
    End If
    
    If outRow = 39 Then
        wsQ.Rows(39).RowHeight = 26
        wsQ.Range("A39:I39").Merge
        wsQ.Range("A39").Value = "(No quotation lines added yet. Use Section 2 above to select product and click '+ ADD ITEM TO QUOTE')"
        wsQ.Range("A39").Font.Italic = True
        wsQ.Range("A39").Font.Size = 9.5
        wsQ.Range("A39").Font.Color = RGB(120, 120, 120)
        wsQ.Range("A39").HorizontalAlignment = xlCenter
        outRow = 40
    End If
    
    Dim actLines As Long, rL As ListRow
    actLines = 0
    For Each rL In tblLines.ListRows
        If Len(Trim$(CStr(rL.Range.Cells(1, tblLines.ListColumns("ProductID").Index).Value))) > 0 Then
            actLines = actLines + 1
        End If
    Next rL
    
    ' Update Capacity Indicator Console on F4:I8
    If Not wsQ.Range("F4:I8").MergeCells Then wsQ.Range("F4:I8").Merge
    With wsQ.Range("F4")
        If actLines = 0 Then
            .Value = "QUOTATION CAPACITY" & vbCrLf & "0 of 10 sizes added" & vbCrLf & "(10 sizes remaining)"
            .Font.Color = RGB(80, 100, 80)
            .Interior.Color = RGB(245, 248, 245)
        ElseIf actLines < 10 Then
            .Value = "QUOTATION CAPACITY" & vbCrLf & actLines & " of 10 sizes added" & vbCrLf & "(" & (10 - actLines) & " sizes remaining)"
            .Font.Color = UI_GREEN_DARK
            .Interior.Color = RGB(235, 245, 235)
        Else
            .Value = "QUOTATION CAPACITY" & vbCrLf & "10 of 10 sizes added" & vbCrLf & "(MAXIMUM CAPACITY REACHED)"
            .Font.Color = RGB(170, 40, 0)
            .Interior.Color = RGB(255, 242, 230)
        End If
        .Font.Name = "Calibri"
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With wsQ.Range("F4:I8").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 210, 180)
    End With
    
    ' Summary Card E14
    wsQ.Range("E14").NumberFormat = "@"
    wsQ.Range("E14").Value = CStr(actLines) & " of 10"
    wsQ.Range("E14").Font.Bold = True
    
    ' Section 3 Header A37
    wsQ.Range("A37").Value = "3. CURRENT QUOTATION LINE ITEMS (" & CStr(actLines) & " of 10 sizes added)"
    
    wsQ.Range("A1:I120").Locked = False
    Exit Sub
ErrorHandler:
    On Error Resume Next
    modProtection.UnprotectSheet wsQ
End Sub

' ============================================================
' 5. USER WORKFLOW ACTION PROCEDURES
' ============================================================

Public Sub UI_NewQuotation()
    Dim prevEvents As Boolean, prevScreen As Boolean
    
    On Error GoTo ErrorHandler
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim wsQ As Worksheet
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    modProtection.UnprotectSheet wsQ
    
    ' Allocate fresh enquiry number
    modQuoteHeader.InitializeNewQuote forceNew:=True
    
    ' Populate clean guided header defaults
    wsQ.Range("B5").Value = Date
    wsQ.Range("B6").Value = "Tender"
    wsQ.Range("B7").Value = "Coated HSAW"
    wsQ.Range("B8").Value = "Unipro Techno Infrastructure Pvt. Ltd"
    wsQ.Range("B9").Value = "Rakesh Gupta"
    wsQ.Range("B10").Value = "Krishna Nagar, Bagodar, Jharkhand - 825322"
    wsQ.Range("B11").Value = "Konar, Jharkhand Project"
    wsQ.Range("B12").Value = "Commercial Sales Team"
    wsQ.Range("B13").Value = "P02"
    wsQ.Range("B14").Value = "30 Days from Dispatch"
    wsQ.Range("B15").Value = 30
    wsQ.Range("B16").Value = "Metres"
    wsQ.Range("B17").Value = "NORMAL (Option 1)"
    
    ' Clean input cells in product requirement area (Rows 20-30)
    wsQ.Range("B20").Value = "IS-3589-2001"
    wsQ.Range("B21").Value = "HSAW"
    wsQ.Range("B22").Value = "ID"
    wsQ.Range("B23").Value = 1226
    wsQ.Range("B24").Value = 8.4
    wsQ.Range("B25").Value = "Fe 410"
    wsQ.Range("B26").Value = "PU 1000 micron"
    wsQ.Range("B27").Value = "Food grade epoxy 406 micron"
    wsQ.Range("B28").Value = 2000
    wsQ.Range("B29").Value = "Metres"
    wsQ.Range("B30").Value = 12#
    
    ' Clear quote lines table
    modQuote.ClearQuoteLines
    
    ' Refresh dropdowns & display table
    modDependentDropdowns.SetupDependentDropdownArchitecture
    modDependentDropdowns.SetupQuoteLineDropdowns
    RefreshQuoteLinesDisplay
    
    SetFeedbackBanner "New quotation created. Configure product and click '+ ADD ITEM TO QUOTE'.", UI_GREEN_DARK
    
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = True
    
    wsQ.Activate
    On Error Resume Next
    wsQ.Range("B8").Select
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    SafeMsgBox "Could not start new quotation: " & Err.Description, vbExclamation
End Sub

Public Sub UI_OpenQuotation()
    On Error GoTo ErrorHandler
    Dim enquiryNo As String
    If Application.DisplayAlerts Then
        enquiryNo = InputBox("Enter Enquiry Reference number to open:" & vbCrLf & _
                             "(e.g. MKTGENQ26-27-001 or MKTG/ENQ/26-27/594)", _
                             "Mega Pipes - Open Quotation", "")
    End If
                         
    enquiryNo = Trim$(enquiryNo)
    If Len(enquiryNo) = 0 Then Exit Sub
    
    Application.ScreenUpdating = False
    Dim ok As Boolean
    ok = modLog.LoadQuote(enquiryNo)
    
    If ok Then
        RefreshQuoteLinesDisplay
        SetFeedbackBanner "Quotation loaded: " & enquiryNo, UI_GREEN_DARK
        ThisWorkbook.Worksheets("QUOTE").Activate
        Application.ScreenUpdating = True
        SafeMsgBox "Quotation " & enquiryNo & " loaded successfully.", vbInformation
    Else
        Application.ScreenUpdating = True
        SafeMsgBox "Could not find quotation snapshot for reference: " & enquiryNo, vbExclamation
    End If
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    SafeMsgBox "Error opening quotation: " & Err.Description, vbCritical
End Sub

Public Sub UI_AddItem()
    Dim wsQ As Worksheet
    Dim quoteId As String, fam As String, bas As String
    Dim sz As Double, wt As Double, grd As String
    Dim extC As String, intC As String, qty As Double, unt As String
    Dim prodId As String, plt As String, plantId As String
    
    On Error GoTo ErrorHandler
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    
    quoteId = Trim$(CStr(wsQ.Range("B4").Value))
    If Len(quoteId) = 0 Then
        quoteId = modSequence.AllocateEnquiryNumber()
        wsQ.Range("B4").Value = quoteId
    End If
    
    ' Pre-check 10-line capacity
    Dim tblLinesCheck As ListObject
    Set tblLinesCheck = ThisWorkbook.Worksheets("QUOTE_LINES").ListObjects("tblQuoteLines")
    Dim activeCount As Long, rC As ListRow
    activeCount = 0
    For Each rC In tblLinesCheck.ListRows
        If Len(Trim$(CStr(rC.Range.Cells(1, tblLinesCheck.ListColumns("ProductID").Index).Value))) > 0 Then
            activeCount = activeCount + 1
        End If
    Next rC
    If activeCount >= 10 Then
        SetFeedbackBanner "Limit reached: Maximum 10 sizes allowed per quotation. Remove a line before adding.", UI_RED
        SafeMsgBox "Quotation capacity reached: A quotation supports up to 10 product lines." & vbCrLf & _
                   "Currently " & CStr(activeCount) & " of 10 sizes are added." & vbCrLf & _
                   "An 11th line cannot be added. Please remove an existing line or edit an existing line.", _
                   vbExclamation, "Mega Pipes - Line Limit"
        Exit Sub
    End If
    
    ' Read inputs from Rows 20-30
    fam = Trim$(CStr(wsQ.Range("B21").Value))
    If Len(fam) = 0 Then fam = "HSAW"
    bas = UCase$(Trim$(CStr(wsQ.Range("B22").Value)))
    If Len(bas) = 0 Then bas = "ID"
    sz = Val(wsQ.Range("B23").Value)
    wt = Val(wsQ.Range("B24").Value)
    grd = Trim$(CStr(wsQ.Range("B25").Value))
    extC = Trim$(CStr(wsQ.Range("B26").Value))
    intC = Trim$(CStr(wsQ.Range("B27").Value))
    qty = Val(wsQ.Range("B28").Value)
    unt = Trim$(CStr(wsQ.Range("B29").Value))
    If Len(unt) = 0 Then unt = "Metres"
    
    ' Plant resolution from B13
    plt = Trim$(CStr(wsQ.Range("B13").Value))
    If InStr(1, plt, "Khopoli", vbTextCompare) > 0 Or InStr(1, plt, "P02", vbTextCompare) > 0 Then
        plantId = "P02"
    Else
        plantId = "P01"
    End If
    
    ' Validation
    If sz <= 0 Then
        SetFeedbackBanner "Validation Error: Please select a valid Pipe Size (mm).", UI_RED
        Exit Sub
    End If
    If wt <= 0 Then
        SetFeedbackBanner "Validation Error: Please select Wall Thickness (mm).", UI_RED
        Exit Sub
    End If
    If qty <= 0 Then
        SetFeedbackBanner "Validation Error: Quantity must be greater than zero.", UI_RED
        Exit Sub
    End If
    
    ' Product ID mapping
    Dim hasCoating As Boolean
    hasCoating = (UCase$(extC) <> "BARE" And Len(extC) > 0 And UCase$(extC) <> "NONE")
    If fam = "HSAW" Then
        prodId = IIf(hasCoating, "P02", "P01")
    Else
        prodId = IIf(hasCoating, "P04", "P03")
    End If
    
    Dim gradeId As String
    gradeId = "HSAW-002"
    If fam = "HFW" Then gradeId = "HFW-002"
    
    Dim prevEvents As Boolean, prevScreen As Boolean
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim legacyMode As Boolean
    legacyMode = (InStr(1, CStr(wsQ.Range("B17").Value), "LEGACY", vbTextCompare) > 0 Or _
                  InStr(1, CStr(wsQ.Range("E12").Value), "LEGACY", vbTextCompare) > 0)
    
    modQuoteLines.AddQuoteLine _
        quoteId:=quoteId, _
        productId:=prodId, _
        family:=fam, _
        basis:=bas, _
        sizeMm:=sz, _
        thicknessMm:=wt, _
        gradeId:=gradeId, _
        quantity:=qty, _
        quantityUnit:=unt, _
        pipeLengthM:=12#, _
        plantId:=plantId, _
        legacyMode:=legacyMode
        
    ' Recalculate Quote and refresh UI
    modQuote.RecalculateQuote
    RefreshQuoteLinesDisplay
    
    ' Clear quantity cell for next entry
    wsQ.Range("B28").Value = 0
    
    Dim totLines As Long
    totLines = ThisWorkbook.Worksheets("QUOTE_LINES").ListObjects("tblQuoteLines").ListRows.Count
    SetFeedbackBanner "Item added successfully! Total quote lines: " & CStr(totLines) & ". Click 'CALCULATE QUOTATION' when ready.", UI_GREEN_DARK
    
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub

ErrorHandler:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    SetFeedbackBanner "Error adding item: " & Err.Description, UI_RED
    SafeMsgBox "Could not add quote item: " & Err.Description, vbExclamation
End Sub

Public Sub UI_AddMultipleSizes()
    On Error GoTo ErrorHandler
    Dim wsQ As Worksheet
    Dim quoteId As String
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    
    quoteId = Trim$(CStr(wsQ.Range("B4").Value))
    If Len(quoteId) = 0 Then
        quoteId = modSequence.AllocateEnquiryNumber()
        wsQ.Range("B4").Value = quoteId
    End If
    
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Do you want to add the 4 standard reference project pipe sizes?" & vbCrLf & _
                 "1. 1226mm (ID) x 8.4mm - 2,000 MTR" & vbCrLf & _
                 "2. 1300mm (ID) x 8.4mm - 350 MTR" & vbCrLf & _
                 "3. 1350mm (ID) x 8.4mm - 500 MTR" & vbCrLf & _
                 "4. 1400mm (ID) x 8.4mm - 4,400 MTR", _
                 vbYesNo + vbQuestion, "Mega Pipes — Rapid Multi-Size Entry")
                 
    If ans = vbYes Then
        Application.ScreenUpdating = False
        modSelfTest.PopulateGoldenReferenceQuote
        SetupQuoteSheetUI
        Application.ScreenUpdating = True
        SafeMsgBox "4 reference quotation sizes populated successfully!", vbInformation
    End If
    Exit Sub
ErrorHandler:
    SafeMsgBox "Multi-size entry error: " & Err.Description, vbExclamation
End Sub

Public Sub UI_CustomSize()
    On Error GoTo ErrorHandler
    Dim wsQ As Worksheet
    Dim quoteId As String
    Dim szInput As String, wtInput As String, qtyInput As String
    Dim customSz As Double, customWt As Double, customQty As Double
    
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    quoteId = Trim$(CStr(wsQ.Range("B4").Value))
    If Len(quoteId) = 0 Then
        quoteId = modSequence.AllocateEnquiryNumber()
        wsQ.Range("B4").Value = quoteId
    End If
    
    szInput = InputBox("Enter Custom Pipe Outer/Inner Diameter in mm:" & vbCrLf & "(Valid range: 100 mm to 3,500 mm)", "Mega Pipes — Custom Pipe Size", "1250")
    If Len(Trim$(szInput)) = 0 Then Exit Sub
    customSz = Val(szInput)
    
    wtInput = InputBox("Enter Custom Wall Thickness in mm:" & vbCrLf & "(Valid range: 3.5 mm to 60.0 mm)", "Mega Pipes — Custom Wall Thickness", "9.5")
    If Len(Trim$(wtInput)) = 0 Then Exit Sub
    customWt = Val(wtInput)
    
    qtyInput = InputBox("Enter Quantity in Metres:", "Mega Pipes — Custom Quantity", "1000")
    If Len(Trim$(qtyInput)) = 0 Then Exit Sub
    customQty = Val(qtyInput)
    
    ' Validate dimensions physically
    modQuoteLines.ValidateCustomDimensions "HSAW", "ID", customSz, customWt
    
    ' Add custom line
    Application.ScreenUpdating = False
    modQuoteLines.AddCustomQuoteLine quoteId, "P02", "HSAW", "ID", customSz, customWt, "HSAW-002", customQty, "Metres", 12#, "P02"
    modQuote.RecalculateQuote
    RefreshQuoteLinesDisplay
    Application.ScreenUpdating = True
    
    SafeMsgBox "Custom pipe size (" & Format$(customSz, "0.0") & "mm x " & Format$(customWt, "0.0") & "mm) added successfully!", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    SafeMsgBox "Custom size entry rejected: " & Err.Description, vbExclamation
End Sub

Public Sub UI_RemoveItem()
    On Error GoTo ErrorHandler
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("QUOTE_LINES").ListObjects("tblQuoteLines")
    
    If tbl.ListRows.Count = 0 Then
        SafeMsgBox "No quote lines available to remove.", vbInformation
        Exit Sub
    End If
    
    Dim lineInput As String
    If Application.DisplayAlerts Then
        lineInput = InputBox("Enter Line ID to remove (e.g. L0001, L0002) or line number (e.g. 1, 2):", _
                             "Mega Pipes — Remove Quote Line", "1")
    Else
        lineInput = "1"
    End If
    lineInput = Trim$(lineInput)
    If Len(lineInput) = 0 Then Exit Sub
    
    Dim targetId As String
    If IsNumeric(lineInput) Then
        targetId = "L" & Format$(CLng(lineInput), "0000")
    Else
        targetId = UCase$(lineInput)
    End If
    
    Application.ScreenUpdating = False
    modQuote.DeleteQuoteLine targetId
    RefreshQuoteLinesDisplay
    SetFeedbackBanner "Line " & targetId & " removed from quotation.", UI_GREEN_DARK
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    SafeMsgBox "Error removing line: " & Err.Description, vbExclamation
End Sub

Public Sub UI_DuplicateItem()
    On Error GoTo ErrorHandler
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("QUOTE_LINES").ListObjects("tblQuoteLines")
    
    If tbl.ListRows.Count = 0 Then
        SafeMsgBox "No quote lines available to duplicate.", vbInformation
        Exit Sub
    End If
    
    Dim activeCount As Long, rC As ListRow
    activeCount = 0
    For Each rC In tbl.ListRows
        If Len(Trim$(CStr(rC.Range.Cells(1, tbl.ListColumns("ProductID").Index).Value))) > 0 Then
            activeCount = activeCount + 1
        End If
    Next rC
    If activeCount >= 10 Then
        SetFeedbackBanner "Cannot duplicate: Maximum 10 sizes already reached.", UI_RED
        SafeMsgBox "Quotation capacity reached (10 of 10 sizes). Remove a line before duplicating.", vbExclamation, "Mega Pipes — Line Limit"
        Exit Sub
    End If
    
    Dim defaultId As String
    defaultId = Trim$(CStr(tbl.ListRows(tbl.ListRows.Count).Range.Cells(1, tbl.ListColumns("LineID").Index).Value))
    If Len(defaultId) = 0 Then defaultId = "L0001"
    
    Dim lineInput As String
    If Application.DisplayAlerts Then
        lineInput = InputBox("Enter Line ID to duplicate (e.g. L0001, L0002) or line number (e.g. 1, 2):", _
                             "Mega Pipes — Duplicate Item", defaultId)
    Else
        lineInput = defaultId
    End If
    lineInput = Trim$(lineInput)
    If Len(lineInput) = 0 Then Exit Sub
    
    Dim targetId As String
    If IsNumeric(lineInput) Then
        targetId = "L" & Format$(CLng(lineInput), "0000")
    Else
        targetId = UCase$(lineInput)
    End If
    
    Application.ScreenUpdating = False
    Dim newId As String
    newId = modQuoteLines.DuplicateQuoteLine(targetId)
    modQuote.RecalculateQuote
    RefreshQuoteLinesDisplay
    
    SetFeedbackBanner "Line " & targetId & " duplicated as " & newId & ". Lines are independently editable.", UI_GREEN_DARK
    Application.ScreenUpdating = True
    SafeMsgBox "Line " & targetId & " duplicated successfully as " & newId & "!" & vbCrLf & _
               "Both lines are completely independent and can be edited separately.", vbInformation
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    SafeMsgBox "Could not duplicate line: " & Err.Description, vbExclamation
End Sub

Public Sub UI_EditItem()
    On Error GoTo ErrorHandler
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("QUOTE_LINES").ListObjects("tblQuoteLines")
    If tbl.ListRows.Count = 0 Then
        SafeMsgBox "No quotation lines available to edit.", vbInformation
        Exit Sub
    End If
    
    Dim defaultId As String
    defaultId = Trim$(CStr(tbl.ListRows(1).Range.Cells(1, tbl.ListColumns("LineID").Index).Value))
    
    Dim lineInput As String
    If Application.DisplayAlerts Then
        lineInput = InputBox("Enter Line ID to edit (e.g. L0001, L0002) or line number (e.g. 1, 2):", "Mega Pipes — Edit Quote Line", defaultId)
    Else
        lineInput = defaultId
    End If
    lineInput = Trim$(lineInput)
    If Len(lineInput) = 0 Then Exit Sub
    
    Dim targetId As String
    If IsNumeric(lineInput) Then
        targetId = "L" & Format$(CLng(lineInput), "0000")
    Else
        targetId = UCase$(lineInput)
    End If
    
    Dim qtyInput As String
    If Application.DisplayAlerts Then
        qtyInput = InputBox("Enter new Quantity in Metres for " & targetId & ":" & vbCrLf & _
                            "(Leave blank to keep existing quantity)", "Mega Pipes — Edit Line Quantity", "")
    Else
        qtyInput = "1500"
    End If
    
    Application.ScreenUpdating = False
    If Len(Trim$(qtyInput)) > 0 And Val(qtyInput) > 0 Then
        modQuoteLines.EditQuoteLine targetId, newQty:=Val(qtyInput)
    End If
    modQuote.RecalculateQuote
    RefreshQuoteLinesDisplay
    SetFeedbackBanner "Line " & targetId & " updated successfully.", UI_GREEN_DARK
    Application.ScreenUpdating = True
    
    SafeMsgBox "Line " & targetId & " updated successfully!", vbInformation
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    SafeMsgBox "Error editing line: " & Err.Description, vbExclamation
End Sub

Public Sub UI_ClearLine()
    On Error GoTo ErrorHandler
    Dim wsQ As Worksheet
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    
    wsQ.Range("B23").Value = 0          ' Size
    wsQ.Range("B24").Value = 0          ' Thickness
    wsQ.Range("B28").Value = 0          ' Quantity
    
    SetFeedbackBanner "Line entry inputs cleared. Enter size, thickness, and quantity to add an item.", UI_GREEN_DARK
    Exit Sub
ErrorHandler:
    SafeMsgBox "Error clearing line inputs: " & Err.Description, vbExclamation
End Sub

Public Sub UI_Recalculate()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    modQuote.RecalculateQuote
    RefreshQuoteLinesDisplay
    
    Dim grandTotal As String
    grandTotal = "Rs. " & Format$(Val(ThisWorkbook.Worksheets("QUOTE").Range("E10").Value), "#,##0")
    SetFeedbackBanner "Quotation recalculated. Grand Total: " & grandTotal, UI_GREEN_DARK
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    SetFeedbackBanner "Recalculation error: " & Err.Description, UI_RED
End Sub

Public Sub UI_ResetDraft()
    On Error GoTo ErrorHandler
    If Application.DisplayAlerts Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("Are you sure you want to reset this draft?" & vbCrLf & vbCrLf & _
                     "This will clear unsaved quotation inputs and lines." & vbCrLf & _
                     "(Previously saved quotations and snapshots will NOT be affected.)", _
                     vbYesNo + vbQuestion + vbDefaultButton2, "Mega Pipes - Reset Draft")
        If ans <> vbYes Then Exit Sub
    End If
    
    Application.ScreenUpdating = False
    modQuote.ClearQuoteLines
    modQuoteHeader.InitializeNewQuote forceNew:=False
    
    Dim wsQ As Worksheet
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    wsQ.Range("B23").Value = 0
    wsQ.Range("B24").Value = 0
    wsQ.Range("B28").Value = 0
    
    RefreshQuoteLinesDisplay
    SetFeedbackBanner "Quotation draft reset. Enter new quotation details.", UI_GREEN_DARK
    Application.ScreenUpdating = True
    SafeMsgBox "Quotation draft reset successfully.", vbInformation
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    SafeMsgBox "Error resetting draft: " & Err.Description, vbExclamation
End Sub

Public Sub UI_ClearAllItems()
    On Error GoTo ErrorHandler
    If Application.DisplayAlerts Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("Are you sure you want to delete ALL quotation lines?", _
                     vbYesNo + vbQuestion + vbDefaultButton2, "Mega Pipes — Clear All Lines")
        If ans <> vbYes Then Exit Sub
    End If
    
    Application.ScreenUpdating = False
    modQuote.ClearQuoteLines
    RefreshQuoteLinesDisplay
    SetFeedbackBanner "All quote lines have been cleared.", UI_GREEN_DARK
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    SafeMsgBox "Error clearing lines: " & Err.Description, vbExclamation
End Sub

Public Sub UI_QuotationHistory()
    On Error GoTo ErrorHandler
    Dim wsLog As Worksheet
    Dim tblLog As ListObject
    Dim r As Long, count As Long
    Dim msg As String
    
    Set wsLog = ThisWorkbook.Worksheets("QUOTE_LOG")
    Set tblLog = wsLog.ListObjects("tblQuoteLog")
    
    msg = "RECENT QUOTATION HISTORY & SNAPSHOT AUDIT:" & vbCrLf & _
          "================================================" & vbCrLf
    count = 0
    If Not tblLog Is Nothing Then
        For r = tblLog.ListRows.Count To 1 Step -1
            Dim qId As String, ts As String, act As String, usr As String, nts As String
            qId = Trim$(CStr(tblLog.ListRows(r).Range.Cells(1, tblLog.ListColumns("QuoteID").Index).Value))
            If Len(qId) > 0 And qId <> "Q-TEST-AUDIT" Then
                ts = Format$(tblLog.ListRows(r).Range.Cells(1, tblLog.ListColumns("Timestamp").Index).Value, "yyyy-mm-dd hh:nn")
                act = Trim$(CStr(tblLog.ListRows(r).Range.Cells(1, tblLog.ListColumns("Action").Index).Value))
                usr = Trim$(CStr(tblLog.ListRows(r).Range.Cells(1, tblLog.ListColumns("User").Index).Value))
                nts = Trim$(CStr(tblLog.ListRows(r).Range.Cells(1, tblLog.ListColumns("Notes").Index).Value))
                
                msg = msg & "[" & ts & "] " & qId & " | " & act & " by " & usr & vbCrLf
                If Len(nts) > 0 Then msg = msg & "   -> " & nts & vbCrLf
                count = count + 1
                If count >= 8 Then Exit For
            End If
        Next r
    End If
    
    If count = 0 Then
        msg = msg & "(No quotation history entries logged yet.)"
    End If
    
    SafeMsgBox msg, vbInformation, "Mega Pipes — Quotation History"
    Exit Sub
ErrorHandler:
    SafeMsgBox "Could not display history: " & Err.Description, vbExclamation
End Sub

Public Sub UI_CalculateQuotation()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    Dim wsQ As Worksheet
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    
    Dim cust As String
    cust = Trim$(CStr(wsQ.Range("B8").Value))
    If Len(cust) = 0 Then
        Application.ScreenUpdating = True
        SetFeedbackBanner "Please select a Customer in cell B8 before calculating.", UI_RED
        SafeMsgBox "Please select a customer from the dropdown list before calculating.", vbExclamation
        Exit Sub
    End If
    
    ' Recalculate quote
    modQuote.RecalculateQuote
    
    ' Update COSTING, OFFER, and SYNOPSIS sheets
    On Error Resume Next
    modExport.WriteCostingSheet
    modExport.WriteOfferSheet
    modSynopsis.PopulateSynopsisReport
    On Error GoTo ErrorHandler
    
    ' Refresh display table
    RefreshQuoteLinesDisplay
    
    Dim grandTotal As String
    grandTotal = "Rs. " & Format$(Val(wsQ.Range("E10").Value), "#,##0")
    SetFeedbackBanner "Quotation calculated successfully! Grand Total: " & grandTotal & ". Ready for review or export.", UI_GREEN_DARK
    
    Application.ScreenUpdating = True
    SafeMsgBox "Quotation calculated successfully!" & vbCrLf & vbCrLf & _
               "Customer: " & cust & vbCrLf & _
               "Subtotal (Excl GST): Rs. " & Format$(Val(wsQ.Range("E8").Value), "#,##0") & vbCrLf & _
               "Grand Total (Incl GST): " & grandTotal, vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    SetFeedbackBanner "Calculation error: " & Err.Description, UI_RED
    SafeMsgBox "Calculation failed: " & Err.Description, vbCritical
End Sub

Public Sub UI_SaveDraft()
    On Error GoTo ErrorHandler
    Dim qNum As String
    qNum = Trim$(CStr(ThisWorkbook.Worksheets("QUOTE").Range("B4").Value))
    If Len(qNum) = 0 Then
        SafeMsgBox "Please allocate a Quote Number before saving.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Dim snapPath As String
    snapPath = modLog.SaveQuote(qNum)
    
    SetFeedbackBanner "Draft saved successfully. Snapshot: " & Dir(snapPath), UI_GREEN_DARK
    Application.ScreenUpdating = True
    SafeMsgBox "Quotation draft saved successfully!" & vbCrLf & vbCrLf & _
               "Snapshot File: " & snapPath, vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    SafeMsgBox "Error saving draft: " & Err.Description, vbCritical
End Sub

Public Sub UI_PreviewOffer()
    On Error GoTo ErrorHandler
    Dim wsOff As Worksheet
    Set wsOff = ThisWorkbook.Worksheets("OFFER")
    wsOff.Visible = xlSheetVisible
    
    ' Return button on OFFER
    On Error Resume Next
    Dim btn As Shape
    Set btn = wsOff.Shapes("btnOfferReturn")
    If btn Is Nothing Then
        AddModernButton wsOff, "btnOfferReturn", "← RETURN TO QUOTE CONSOLE", "modUI.UI_ReturnToQuote", _
                        wsOff.Range("A1"), wsOff.Range("C1"), UI_GREEN_DARK, UI_WHITE, True
    End If
    On Error GoTo ErrorHandler
    
    wsOff.Activate
    Exit Sub

ErrorHandler:
    SafeMsgBox "Could not preview offer letter: " & Err.Description, vbExclamation
End Sub

Public Sub UI_PreviewSynopsis()
    On Error GoTo ErrorHandler
    Dim wsSyn As Worksheet
    modSynopsis.PopulateSynopsisReport
    Set wsSyn = ThisWorkbook.Worksheets("SYNOPSIS")
    wsSyn.Visible = xlSheetVisible
    wsSyn.Activate
    Exit Sub
ErrorHandler:
    SafeMsgBox "Could not preview Synopsis report: " & Err.Description, vbExclamation
End Sub

Public Sub UI_ReturnToQuote()
    On Error Resume Next
    ThisWorkbook.Worksheets("QUOTE").Activate
End Sub

Public Sub UI_ExportPDF()
    On Error GoTo ErrorHandler
    Dim qNum As String
    qNum = Trim$(CStr(ThisWorkbook.Worksheets("QUOTE").Range("B4").Value))
    If Len(qNum) = 0 Then
        SafeMsgBox "No active quotation to export.", vbInformation
        Exit Sub
    End If
    
    modExport.ExportPDF
    SafeMsgBox "Quotation Offer Letter exported to PDF successfully.", vbInformation
    Exit Sub

ErrorHandler:
    SafeMsgBox "PDF export failed: " & Err.Description, vbCritical
End Sub

Public Sub UI_ExportCustomerExcel()
    On Error GoTo ErrorHandler
    Dim qNum As String
    qNum = Trim$(CStr(ThisWorkbook.Worksheets("QUOTE").Range("B4").Value))
    If Len(qNum) = 0 Then
        SafeMsgBox "No active quotation to export.", vbInformation
        Exit Sub
    End If
    
    modExport.WriteCustomerExcel
    SafeMsgBox "Sanitized Customer Excel workbook exported successfully." & vbCrLf & _
               "Confidentiality guaranteed: Internal cost and margin formulas excluded.", vbInformation
    Exit Sub

ErrorHandler:
    SafeMsgBox "Customer Excel export failed: " & Err.Description, vbCritical
End Sub

Public Sub UI_BackToHome()
    On Error GoTo ErrorHandler
    HideBackendSheets
    BuildStartScreen
    ThisWorkbook.Worksheets("START").Activate
    Exit Sub

ErrorHandler:
    ThisWorkbook.Worksheets("START").Activate
End Sub

Public Sub UI_AdminUnlock()
    On Error GoTo ErrorHandler
    ShowAllSheets
    modProtection.UnprotectAllSheets
    SafeMsgBox "Backend sheets revealed for administration and engineering access." & vbCrLf & _
               "(COSTING, SYNOPSIS, COMPARE, QUOTE_LINES, MASTERS, SETTINGS, etc.)", vbInformation
    Exit Sub

ErrorHandler:
    SafeMsgBox "Admin access error: " & Err.Description, vbExclamation
End Sub

' ============================================================
' 6. SHEET VISIBILITY HELPERS
' ============================================================

Public Sub HideBackendSheets()
    Dim ws As Worksheet
    Dim sheetNames() As String
    sheetNames = Split(BACKEND_SHEETS, ",")
    
    Dim i As Long
    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(Trim$(sheetNames(i)))
        If Not ws Is Nothing Then
            ws.Visible = xlSheetHidden
        End If
        Set ws = Nothing
        On Error GoTo 0
    Next i
End Sub

Public Sub ShowAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
End Sub

' ============================================================
' 7. FEEDBACK BANNER HELPER
' ============================================================

Private Sub SetFeedbackBanner(ByVal msg As String, ByVal colr As Long)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("QUOTE")
    modProtection.UnprotectSheet ws
    With ws.Range("A35")
        .Value = msg
        .Font.Color = colr
    End With
End Sub

' ============================================================
' 8. MODERN BUTTON BUILDER (CRISP, NON-OVERLAPPING)
' ============================================================

Private Sub AddModernButton(ByRef ws As Worksheet, ByVal btnName As String, _
                            ByVal btnText As String, ByVal macroName As String, _
                            ByRef startCell As Range, ByRef endCell As Range, _
                            ByVal bgColor As Long, ByVal fgColor As Long, _
                            ByVal isBold As Boolean)
    Dim btn As Shape
    Dim btnLeft As Single, btnTop As Single, btnWidth As Single, btnHeight As Single
    
    btnLeft = startCell.Left + 2
    btnTop = startCell.Top + 2
    btnWidth = (endCell.Left + endCell.Width) - startCell.Left - 4
    btnHeight = (endCell.Top + endCell.Height) - startCell.Top - 4
    If btnHeight < 20 Then btnHeight = 20
    
    ' msoShapeRoundedRectangle = 5
    Set btn = ws.Shapes.AddShape(5, btnLeft, btnTop, btnWidth, btnHeight)
    With btn
        .Name = btnName
        .Fill.ForeColor.RGB = bgColor
        .Fill.Solid
        .Line.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = btnText
            .TextRange.Font.Size = IIf(isBold, 10, 9)
            .TextRange.Font.Bold = isBold
            .TextRange.Font.Fill.ForeColor.RGB = fgColor
            .TextRange.Font.Name = "Calibri"
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
            .MarginLeft = 2
            .MarginRight = 2
            .MarginTop = 1
            .MarginBottom = 1
            .WordWrap = msoFalse
        End With
        .OnAction = macroName
    End With
End Sub

' ============================================================
' 9. SAFE MSGBOX HELPER
' ============================================================

Private Sub SafeMsgBox(ByVal msg As String, Optional ByVal style As VbMsgBoxStyle = vbInformation, Optional ByVal title As String = "Mega Pipes")
    If Application.DisplayAlerts And Application.UserControl Then
        MsgBox msg, style, title
    End If
End Sub
