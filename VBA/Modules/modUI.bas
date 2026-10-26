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
Private Const BACKEND_SHEETS As String = "QUOTE_LINES,COSTING,MASTERS,SETTINGS,CUSTOMERS,PLANTS,DISTANCE_CACHE,QUOTE_LOG"

' Interactive Line Selection & Edit State
Public SelectedLineID As String
Public IsEditMode     As Boolean
Public EditingLineID  As String

' ============================================================
' 1. WORKBOOK OPEN & LIFECYCLE
' ============================================================

Public Sub OnWorkbookOpen()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Hide internal backend calculation sheets
    HideBackendSheets
    
    ' Hide commercial output sheets until quotation is calculated
    HideOutputSheets
    
    ' Ensure no sheet protection prevents editing on input sheets
    modProtection.UnprotectAllSheets
    
    ' Re-apply UserInterfaceOnly protection to calculated output sheets
    modProtection.ProtectCalculatedSheets True
    
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
    
    ' 1. Delete all existing shapes FIRST while cell anchors are valid
    Dim sIdx As Long
    For sIdx = ws.Shapes.Count To 1 Step -1
        On Error Resume Next
        ws.Shapes(sIdx).Delete
        On Error GoTo 0
    Next sIdx
    
    ' 2. Clear all existing cells & prior ghost blocks
    ws.Cells.Clear
    

    
    ' ── 1. Unified 9-Column Grid (Total Width = 159) ──
    ws.Columns("A").ColumnWidth = 18    ' Input Labels / Line Action Buttons
    ws.Columns("B").ColumnWidth = 18    ' Input Values / Line Action Buttons
    ws.Columns("C").ColumnWidth = 16    ' Merged Input B:C / Line Action Buttons / Unit
    ws.Columns("D").ColumnWidth = 30    ' Table: Pipe Specification / Summary labels
    ws.Columns("E").ColumnWidth = 23    ' Table: Coating (Ext/Int) / Summary values
    ws.Columns("F").ColumnWidth = 12    ' Table: Quantity / Toolbar
    ws.Columns("G").ColumnWidth = 15    ' Table: Pipes & Weight MT / Toolbar
    ws.Columns("H").ColumnWidth = 12    ' Table: Unit Price / Toolbar
    ws.Columns("I").ColumnWidth = 15    ' Table: Line Total / Toolbar
    
    ' ── 2. Row Heights ──
    ws.Rows(1).RowHeight = 26           ' Top brand title
    ws.Rows(2).RowHeight = 8            ' Spacer
    ws.Rows(3).RowHeight = 24           ' Block 1 card headers
    For rIdx = 4 To 17
        ws.Rows(rIdx).RowHeight = 21    ' Block 1 content & Quotation Toolbar
    Next rIdx
    ws.Rows(18).RowHeight = 10          ' Spacer
    ws.Rows(19).RowHeight = 24          ' Block 2 card headers (Inputs Left, Table Right)
    ws.Rows(20).RowHeight = 22          ' Table Column Titles / First input
    For rIdx = 21 To 30
        ws.Rows(rIdx).RowHeight = 21    ' Quote Lines 1-10 / Inputs 2-11
    Next rIdx
    ws.Rows(31).RowHeight = 22          ' Custom Pipe Size toggle / Table Summary Row
    ws.Rows(32).RowHeight = 22          ' Manual Selling Price & Unit / Table Status
    ws.Rows(33).RowHeight = 22          ' Override Reason / Table Status
    ws.Rows(34).RowHeight = 18          ' Costing helper guidance
    ws.Rows(35).RowHeight = 24          ' Action tier 1: + ADD ITEM / CUSTOM SIZE
    ws.Rows(36).RowHeight = 24          ' Action tier 2: EDIT LINE / PRICE OVERRIDE / - REMOVE
    ws.Rows(37).RowHeight = 24          ' Action tier 3: DUPLICATE / CLEAR INPUTS / CLEAR ALL
    ws.Rows(38).RowHeight = 18          ' Guidance tip text
    ws.Rows(39).RowHeight = 10          ' Bottom Spacer
    
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
    
    If Not ws.Range("D15:E17").MergeCells Then ws.Range("D15:E17").Merge
    With ws.Range("D15")
        .Value = "DISTANCE INFO:" & vbCrLf & "Click 'CALCULATE QUOTATION' to see nearest plant & competitor distances."
        .Font.Name = "Calibri"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Color = RGB(110, 110, 110)
        .Interior.Color = RGB(248, 250, 248)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    With ws.Range("D3:E17").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(190, 215, 190)
    End With
    
    ' ── 5. Block 1 Right Card: Quotation Actions Toolbar (F3:I17) ──
    ws.Range("F3:I3").Merge
    With ws.Range("F3")
        .Value = "QUOTATION COMMAND CENTER & ACTIONS"
        .Interior.Color = RGB(26, 82, 45)
        .Font.Color = UI_WHITE
        .Font.Bold = True
        .Font.Size = 10.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range("F4:I17").Interior.Color = RGB(250, 252, 250)
    With ws.Range("F3:I17").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(190, 215, 190)
    End With
    
    ' Primary Action: Calculate Quotation & Approve (Rows 4-5)
    AddModernButton ws, "btnQuoteCalc", "CALCULATE QUOTATION", "modUI.UI_CalculateQuotation", _
                    ws.Range("F4"), ws.Range("G5"), UI_GREEN_DARK, UI_WHITE, True
    AddModernButton ws, "btnQuoteApprove", "APPROVE QUOTATION", "modUI.UI_ApproveQuotation", _
                    ws.Range("H4"), ws.Range("I5"), RGB(26, 82, 45), UI_WHITE, True
                    
    ' Tier 2: Save Draft / Submit for Review (Rows 6-7)
    AddModernButton ws, "btnQuoteSave", "SAVE DRAFT", "modUI.UI_SaveDraft", _
                    ws.Range("F6"), ws.Range("G7"), UI_NAVY, UI_WHITE, True
    AddModernButton ws, "btnQuoteReview", "SUBMIT FOR REVIEW", "modUI.UI_SubmitReview", _
                    ws.Range("H6"), ws.Range("I7"), RGB(50, 110, 170), UI_WHITE, True
                    
    ' Tier 3: Document Previews (Rows 8-9)
    AddModernButton ws, "btnQuotePreview", "PREVIEW OFFER", "modUI.UI_PreviewOffer", _
                    ws.Range("F8"), ws.Range("G9"), RGB(70, 100, 130), UI_WHITE, True
    AddModernButton ws, "btnQuoteSynopsis", "PREVIEW SYNOPSIS", "modUI.UI_PreviewSynopsis", _
                    ws.Range("H8"), ws.Range("I9"), UI_TEAL, UI_WHITE, True
                    
    ' Tier 4: Document Exports (Rows 10-11)
    AddModernButton ws, "btnQuotePDF", "EXPORT PDF", "modUI.UI_ExportPDF", _
                    ws.Range("F10"), ws.Range("G11"), RGB(160, 45, 35), UI_WHITE, True
    AddModernButton ws, "btnQuoteExcel", "EXPORT EXCEL", "modUI.UI_ExportCustomerExcel", _
                    ws.Range("H10"), ws.Range("I11"), RGB(20, 110, 160), UI_WHITE, True
                    
    ' Tier 5: Reset / Return to Home (Rows 12-13)
    AddModernButton ws, "btnQuoteReset", "RESET DRAFT", "modUI.UI_ResetDraft", _
                    ws.Range("F12"), ws.Range("G13"), RGB(140, 80, 50), UI_WHITE, False
    AddModernButton ws, "btnQuoteHome", "<- BACK TO START", "modUI.UI_BackToHome", _
                    ws.Range("H12"), ws.Range("I13"), UI_GRAY_DARK, UI_WHITE, True
                    
    ' Tier 6: Plant & Competitor Intelligence (btnQuoteCompare)
    Dim btnCmp As Shape
    Set btnCmp = ws.Shapes.AddShape(1, 597#, 280#, 307#, 35#)
    With btnCmp
        .Name = "btnQuoteCompare"
        .Fill.ForeColor.RGB = RGB(32, 96, 32)
        .Fill.Solid
        .Line.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = "PLANT & TERRITORY INTEL"
            .TextRange.Font.Size = 10
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = UI_WHITE
            .TextRange.Font.Name = "Segoe UI"
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
        End With
        .OnAction = "modUI.UI_PreviewCompare"
    End With
                    
    ' Capacity Badge Console (Rows 14-17)
    ws.Range("F14:I17").Merge
    With ws.Range("F14")
        .Value = "QUOTATION CAPACITY" & vbCrLf & "0 of 10 sizes added" & vbCrLf & "(10 sizes remaining)"
        .Font.Name = "Calibri"
        .Font.Bold = True
        .Font.Size = 10.5
        .Font.Color = RGB(80, 100, 80)
        .Interior.Color = RGB(245, 248, 245)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("F14:I17").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 210, 180)
    End With
    
    ' ── 6. Block 2 Left Card: Pipe Requirement Input Box (Rows 19-34) ──
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
    
    With ws.Range("A19:C30").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 210, 180)
    End With
    
    ' ── Controls Container: Custom Size, Manual Price Override, Reason (Rows 31-34) ──
    ' Row 31: Custom Pipe Size Toggle
    ws.Range("A31").Value = "Custom Pipe Size"
    ws.Range("B31").Value = "NO"
    ws.Range("C31").Value = ""
    ws.Range("A31").Interior.Color = RGB(248, 250, 248)
    ws.Range("B31").Interior.Color = RGB(255, 255, 255)
    ws.Range("C31").Interior.Color = RGB(255, 255, 255)
    ws.Range("A31").Font.Bold = True
    ws.Range("A31").Font.Size = 9.5
    ws.Range("A31").Font.Color = RGB(30, 45, 30)
    ws.Range("B31").Font.Bold = True
    ws.Range("B31").Font.Size = 9.5
    ws.Range("B31").HorizontalAlignment = xlCenter
    ws.Range("B31").VerticalAlignment = xlCenter
    
    ' Add Data Validation to B31 (NO / YES list)
    On Error Resume Next
    With ws.Range("B31").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
             Operator:=xlBetween, Formula1:="NO,YES"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    On Error GoTo 0
    
    ' Row 32: Manual Selling Price (Optional) & Unit Indicator
    ws.Range("A32").Value = "Manual Selling Price (Optional)"
    ws.Range("A32").Interior.Color = RGB(255, 255, 255)
    ws.Range("A32").Font.Bold = True
    ws.Range("A32").Font.Size = 9.5
    ws.Range("A32").Font.Color = RGB(30, 45, 30)
    
    ws.Range("B32").NumberFormat = "#,##0.00"
    ws.Range("B32").Interior.Color = RGB(255, 255, 255)
    ws.Range("B32").Font.Name = "Calibri"
    ws.Range("B32").Font.Size = 9.5
    ws.Range("B32").Font.Bold = True
    ws.Range("B32").HorizontalAlignment = xlRight
    ws.Range("B32").VerticalAlignment = xlCenter
    ws.Range("B32").Locked = False
    
    ' Unit indicator in C32 derived from quotation actual PriceBasis
    modQuote.UpdateManualPriceUnitDisplay ws
    
    ' Row 33: Override Reason
    ws.Range("A33").Value = "Override Reason"
    ws.Range("A33").Interior.Color = RGB(248, 250, 248)
    ws.Range("A33").Font.Bold = True
    ws.Range("A33").Font.Size = 9.5
    ws.Range("A33").Font.Color = RGB(30, 45, 30)
    
    ws.Range("B33:C33").Locked = False
    ws.Range("B33:C33").Merge
    ws.Range("B33").Interior.Color = RGB(255, 255, 255)
    ws.Range("B33").Font.Name = "Calibri"
    ws.Range("B33").Font.Size = 9
    ws.Range("B33").HorizontalAlignment = xlLeft
    ws.Range("B33").VerticalAlignment = xlCenter
    
    ' Row 34: Helper Guidance Text
    ws.Range("A34:C34").Merge
    With ws.Range("A34")
        .Value = "Leave blank to use automatically calculated selling price."
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Italic = True
        .Font.Color = RGB(110, 110, 110)
        .Interior.Color = RGB(248, 250, 248)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("A19:C34").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 210, 180)
    End With
    
    ' ── Line Item Actions Container (Rows 35-38) ──
    ws.Range("A35:C38").Interior.Color = RGB(250, 252, 250)
    With ws.Range("A35:C38").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 210, 180)
    End With
    
    ' Line Action Tier 1 (Row 35): + ADD ITEM (wide green button) and CUSTOM SIZE (amber button)
    AddModernButton ws, "btnQuoteAdd", "+ ADD ITEM TO QUOTE", "modUI.UI_AddItem", _
                    ws.Range("A35"), ws.Range("B35"), UI_GREEN_MID, UI_WHITE, True
    AddModernButton ws, "btnQuoteCustomSize", "CUSTOM SIZE", "modUI.UI_CustomSize", _
                    ws.Range("C35"), ws.Range("C35"), RGB(160, 120, 20), UI_WHITE, True
                    
    ' Line Action Tier 2 (Row 36): EDIT LINE, PRICE OVERRIDE, - REMOVE LINE
    AddModernButton ws, "btnQuoteEdit", "EDIT LINE", "modUI.UI_EditItem", _
                    ws.Range("A36"), ws.Range("A36"), UI_BLUE, UI_WHITE, True
    AddModernButton ws, "btnQuotePriceOverride", "PRICE OVERRIDE", "modUI.UI_ManualPriceOverride", _
                    ws.Range("B36"), ws.Range("B36"), UI_ORANGE, UI_WHITE, True
    AddModernButton ws, "btnQuoteRemove", "- REMOVE", "modUI.UI_RemoveItem", _
                    ws.Range("C36"), ws.Range("C36"), RGB(170, 60, 50), UI_WHITE, True
                    
    ' Line Action Tier 3 (Row 37): DUPLICATE, CLEAR INPUTS, CLEAR ALL LINES
    AddModernButton ws, "btnQuoteDuplicate", "DUPLICATE", "modUI.UI_DuplicateItem", _
                    ws.Range("A37"), ws.Range("A37"), RGB(40, 120, 90), UI_WHITE, True
    AddModernButton ws, "btnQuoteClearLine", "CLEAR INPUTS", "modUI.UI_ClearLine", _
                    ws.Range("B37"), ws.Range("B37"), RGB(110, 120, 130), UI_WHITE, False
    AddModernButton ws, "btnQuoteClear", "CLEAR ALL", "modUI.UI_ClearAllInputs", _
                    ws.Range("C37"), ws.Range("C37"), RGB(90, 90, 90), UI_WHITE, False
                    
    ' Guidance Tip (Row 38)
    ws.Range("A38:C38").Merge
    With ws.Range("A38")
        .Value = "Tip: Select specifications and click '+ ADD ITEM TO QUOTE'."
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Italic = True
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' ── 7. Block 2 Right Card: Quotation Line Items Table (Columns D-I, Rows 19-34) ──
    ws.Range("D19:I19").Merge
    With ws.Range("D19")
        .Value = "3. CURRENT QUOTATION LINE ITEMS"
        .Interior.Color = UI_GREEN_DARK
        .Font.Color = UI_WHITE
        .Font.Bold = True
        .Font.Size = 10.5
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range("D20").Value = "Pipe Specification"
    ws.Range("E20").Value = "Coating (Ext / Int)"
    ws.Range("F20").Value = "Quantity"
    ws.Range("G20").Value = "Pipes / Weight"
    ws.Range("H20").Value = "Unit Price"
    ws.Range("I20").Value = "Line Total"
    
    With ws.Range("D20:I20")
        .Interior.Color = UI_GREEN_MID
        .Font.Color = UI_WHITE
        .Font.Bold = True
        .Font.Size = 9.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("D19:I34").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 210, 180)
    End With
    
    ' Populate active lines display
    RefreshQuoteLinesDisplay
    
    ' Set master data validations
    modMasterData.SetupMasterDrivenHeader
    modDependentDropdowns.SetupDependentDropdownArchitecture
    
    ' Ensure all input cells are unlocked
    ws.Range("A1:I39").Locked = False
    
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
    
    ' Clear previous display rows (Columns D-I, rows 21 to 38)
    wsQ.Range("D21:I38").UnMerge
    wsQ.Range("D21:I38").ClearContents
    wsQ.Range("D21:I38").Interior.Color = xlNone
    wsQ.Range("D21:I38").Borders.LineStyle = xlNone
    
    outRow = 21
    Dim totMetres As Double, totMT As Double, totPieces As Long, totValue As Double
    totMetres = 0#: totMT = 0#: totPieces = 0: totValue = 0#
    
    If tblLines.ListRows.Count > 0 Then
        For Each rObj In tblLines.ListRows
            Dim pId As String
            pId = Trim$(CStr(rObj.Range.Cells(1, tblLines.ListColumns("ProductID").Index).Value))
            If Len(pId) > 0 Then
                Dim fam As String, bas As String, sz As Double, wt As Double, grd As String
                Dim qty As Double, unt As String, pcs As Long, mt As Double, prc As Double, tot As Double
                Dim extC As String, intC As String, st As String
                Dim lId As String
                
                lId = CStr(rObj.Range.Cells(1, tblLines.ListColumns("LineID").Index).Value)
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
                
                Dim isSel As Boolean
                isSel = (StrComp(lId, SelectedLineID, vbTextCompare) = 0)
                
                Dim pSrc As String
                pSrc = ""
                On Error Resume Next
                pSrc = CStr(rObj.Range.Cells(1, tblLines.ListColumns("PriceSource").Index).Value)
                On Error GoTo 0
                
                wsQ.Rows(outRow).RowHeight = 21
                wsQ.Cells(outRow, 4).Value = lId & ": " & fam & " " & sz & "mm (" & bas & ") x " & wt & "mm " & grd
                wsQ.Cells(outRow, 5).Value = extC & " / " & intC
                wsQ.Cells(outRow, 6).Value = Format$(qty, "#,##0") & " " & unt
                wsQ.Cells(outRow, 7).Value = CStr(pcs) & " pcs (" & Format$(mt, "#,##0.00") & " MT)"
                wsQ.Cells(outRow, 8).Value = prc
                wsQ.Cells(outRow, 8).NumberFormat = "[$" & ChrW$(8377) & "-4009]#,##0"
                If pSrc = "MANUAL_OVERRIDE" Then
                    wsQ.Cells(outRow, 8).Font.Color = RGB(180, 70, 0)
                    wsQ.Cells(outRow, 8).Font.Bold = True
                End If
                wsQ.Cells(outRow, 9).Value = tot
                wsQ.Cells(outRow, 9).NumberFormat = "[$" & ChrW$(8377) & "-4009]#,##0"
                
                ' Row styling (with distinct selection accent)
                With wsQ.Range(wsQ.Cells(outRow, 4), wsQ.Cells(outRow, 9))
                    .Font.Name = "Calibri"
                    .Font.Size = 9
                    .Font.Bold = isSel
                    .WrapText = False
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = IIf(isSel, xlMedium, xlThin)
                    .Borders.Color = IIf(isSel, RGB(60, 130, 200), RGB(210, 225, 210))
                    .Interior.Color = IIf(isSel, RGB(220, 238, 255), IIf(outRow Mod 2 = 1, RGB(248, 252, 248), RGB(255, 255, 255)))
                End With
                wsQ.Cells(outRow, 4).HorizontalAlignment = xlLeft
                wsQ.Cells(outRow, 5).HorizontalAlignment = xlLeft
                wsQ.Cells(outRow, 6).HorizontalAlignment = xlRight
                wsQ.Cells(outRow, 7).HorizontalAlignment = xlRight
                wsQ.Cells(outRow, 8).HorizontalAlignment = xlRight
                wsQ.Cells(outRow, 9).HorizontalAlignment = xlRight
                
                totMetres = totMetres + qty
                totMT = totMT + mt
                totPieces = totPieces + pcs
                totValue = totValue + tot
                
                outRow = outRow + 1
            End If
        Next rObj
    End If
    
    Dim actLines As Long, rL As ListRow
    actLines = 0
    For Each rL In tblLines.ListRows
        If Len(Trim$(CStr(rL.Range.Cells(1, tblLines.ListColumns("ProductID").Index).Value))) > 0 Then
            actLines = actLines + 1
        End If
    Next rL
    
    If outRow = 21 Then
        wsQ.Rows(21).RowHeight = 26
        wsQ.Range("D21:I21").Merge
        wsQ.Range("D21").Value = "(No quotation lines added yet. Select product specifications on left and click '+ ADD ITEM TO QUOTE')"
        wsQ.Range("D21").Font.Italic = True
        wsQ.Range("D21").Font.Size = 9.5
        wsQ.Range("D21").Font.Color = RGB(120, 120, 120)
        wsQ.Range("D21").HorizontalAlignment = xlCenter
        With wsQ.Range("D21:I21").Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(210, 225, 210)
        End With
    Else
        ' Table Summary / Total Row immediately below active lines
        wsQ.Rows(outRow).RowHeight = 22
        wsQ.Cells(outRow, 4).Value = "TABLE TOTAL (" & CStr(actLines) & " of 10 sizes):"
        wsQ.Cells(outRow, 4).Font.Bold = True
        wsQ.Cells(outRow, 4).Font.Size = 9.5
        wsQ.Cells(outRow, 4).HorizontalAlignment = xlLeft
        
        wsQ.Cells(outRow, 5).Value = ""
        
        wsQ.Cells(outRow, 6).Value = Format$(totMetres, "#,##0") & " Metres"
        wsQ.Cells(outRow, 6).Font.Bold = True
        wsQ.Cells(outRow, 6).Font.Size = 9.5
        wsQ.Cells(outRow, 6).HorizontalAlignment = xlRight
        
        wsQ.Cells(outRow, 7).Value = CStr(totPieces) & " pcs (" & Format$(totMT, "#,##0.00") & " MT)"
        wsQ.Cells(outRow, 7).Font.Bold = True
        wsQ.Cells(outRow, 7).Font.Size = 9.5
        wsQ.Cells(outRow, 7).HorizontalAlignment = xlRight
        
        wsQ.Cells(outRow, 8).Value = ""
        
        wsQ.Cells(outRow, 9).Value = totValue
        wsQ.Cells(outRow, 9).NumberFormat = "[$" & ChrW$(8377) & "-4009]#,##0"
        wsQ.Cells(outRow, 9).Font.Bold = True
        wsQ.Cells(outRow, 9).Font.Size = 10.5
        wsQ.Cells(outRow, 9).Font.Color = RGB(120, 60, 0)
        wsQ.Cells(outRow, 9).Interior.Color = RGB(255, 243, 205)
        wsQ.Cells(outRow, 9).HorizontalAlignment = xlRight
        
        With wsQ.Range(wsQ.Cells(outRow, 4), wsQ.Cells(outRow, 9)).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(160, 190, 160)
        End With
        With wsQ.Range(wsQ.Cells(outRow, 4), wsQ.Cells(outRow, 9)).Borders(xlEdgeBottom)
            .LineStyle = xlDouble
            .Color = RGB(160, 190, 160)
        End With
        
        ' Clear any remaining rows below outRow up to row 38
        If outRow < 38 Then
            wsQ.Range(wsQ.Cells(outRow + 1, 4), wsQ.Cells(38, 9)).ClearContents
            wsQ.Range(wsQ.Cells(outRow + 1, 4), wsQ.Cells(38, 9)).Interior.Color = xlNone
            wsQ.Range(wsQ.Cells(outRow + 1, 4), wsQ.Cells(38, 9)).Borders.LineStyle = xlNone
        End If
    End If
    
    ' Update Capacity Indicator Console on F14:I17
    If Not wsQ.Range("F14:I17").MergeCells Then wsQ.Range("F14:I17").Merge
    With wsQ.Range("F14")
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
        .Font.Size = 10.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With wsQ.Range("F14:I17").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 210, 180)
    End With
    
    ' Summary Card E14
    wsQ.Range("E14").NumberFormat = "@"
    wsQ.Range("E14").Value = CStr(actLines) & " of 10"
    wsQ.Range("E14").Font.Bold = True
    
    ' Section 3 Header D19
    wsQ.Range("D19").Value = "3. CURRENT QUOTATION LINE ITEMS (" & CStr(actLines) & " of 10 sizes added)"
    
    wsQ.Range("A1:I39").Locked = False
    Exit Sub
ErrorHandler:
    On Error Resume Next
    modProtection.UnprotectSheet wsQ
End Sub

Public Sub HandleLineSelection(ByVal targetRow As Long)
    Dim wsQ As Worksheet
    Dim tblLines As ListObject
    Dim actLines As Long, rL As ListRow
    
    On Error GoTo ErrorHandler
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    Set tblLines = ThisWorkbook.Worksheets("QUOTE_LINES").ListObjects("tblQuoteLines")
    
    actLines = 0
    For Each rL In tblLines.ListRows
        If Len(Trim$(CStr(rL.Range.Cells(1, tblLines.ListColumns("ProductID").Index).Value))) > 0 Then
            actLines = actLines + 1
        End If
    Next rL
    
    If actLines = 0 Or targetRow < 21 Or targetRow > (20 + actLines) Then Exit Sub
    
    Dim selIdx As Long
    selIdx = targetRow - 20
    SelectedLineID = Trim$(CStr(tblLines.ListRows(selIdx).Range.Cells(1, tblLines.ListColumns("LineID").Index).Value))
    
    ' Quick visual refresh of table rows to highlight selection
    Dim r As Long
    For r = 21 To 20 + actLines
        Dim isSel As Boolean
        isSel = (r = targetRow)
        With wsQ.Range(wsQ.Cells(r, 4), wsQ.Cells(r, 9))
            .Font.Bold = isSel
            .Borders.Weight = IIf(isSel, xlMedium, xlThin)
            .Borders.Color = IIf(isSel, RGB(60, 130, 200), RGB(210, 225, 210))
            .Interior.Color = IIf(isSel, RGB(220, 238, 255), IIf(r Mod 2 = 1, RGB(248, 252, 248), RGB(255, 255, 255)))
        End With
    Next r
    
    Dim fam As String, sz As String
    fam = CStr(tblLines.ListRows(selIdx).Range.Cells(1, tblLines.ListColumns("Family").Index).Value)
    sz = CStr(tblLines.ListRows(selIdx).Range.Cells(1, tblLines.ListColumns("SizeMM").Index).Value)
    SetFeedbackBanner "Selected: " & SelectedLineID & " (" & fam & " " & sz & "mm). Click 'EDIT LINE', 'DUPLICATE', or '- REMOVE'.", UI_BLUE
    Exit Sub
ErrorHandler:
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
    
    ' Populate clean guided header defaults (without hardcoded customer details)
    wsQ.Range("B5").Value = Date
    wsQ.Range("B6").Value = "Tender"
    wsQ.Range("B7").Value = "Coated HSAW"
    wsQ.Range("B8").Value = ""          ' Blank customer for new quotation entry
    wsQ.Range("B9").Value = ""          ' Blank contact person
    wsQ.Range("B10").Value = ""         ' Blank delivery place
    wsQ.Range("B11").Value = ""         ' Blank project / tender ref
    wsQ.Range("B12").Value = ""         ' Blank salesperson
    wsQ.Range("B13").Value = "P01"      ' Default supplying plant
    wsQ.Range("B14").Value = "30 Days from Dispatch"
    wsQ.Range("B15").Value = 30
    wsQ.Range("B16").Value = "Metres"
    wsQ.Range("B17").Value = "NORMAL"
    
    ' Clean input cells in product requirement area (Rows 20-33)
    wsQ.Range("B20").Value = "IS-3589-2001"
    wsQ.Range("B21").Value = "HSAW"
    wsQ.Range("B22").Value = "ID"
    wsQ.Range("B23").Value = 0
    wsQ.Range("B24").Value = 0
    wsQ.Range("B25").Value = "Fe 410"
    wsQ.Range("B26").Value = "Bare"
    wsQ.Range("B27").Value = "None"
    wsQ.Range("B28").Value = 0
    wsQ.Range("B29").Value = "Metres"
    wsQ.Range("B30").Value = 12#
    wsQ.Range("B31").Value = "NO"
    wsQ.Range("C31").Value = ""
    wsQ.Range("B32").Value = ""
    wsQ.Range("B33").Value = ""
    modQuote.UpdateManualPriceUnitDisplay wsQ
    
    ' Clear quote lines table (resets line records, transient costs, and live summary)
    modQuote.ClearQuoteLines
    
    ' Invalidate and hide commercial output sheets (OFFER and SYNOPSIS)
    InvalidateOutputSheets
    
    ' Reset edit mode and selection
    IsEditMode = False
    EditingLineID = ""
    SelectedLineID = ""
    On Error Resume Next
    wsQ.Shapes("btnQuoteAdd").TextFrame2.TextRange.Text = "+ ADD ITEM TO QUOTE"
    wsQ.Shapes("btnQuoteAdd").Fill.ForeColor.RGB = UI_GREEN_MID
    On Error GoTo 0
    
    ' Refresh dropdowns & display table
    modDependentDropdowns.SetupDependentDropdownArchitecture
    modDependentDropdowns.SetupQuoteLineDropdowns
    RefreshQuoteLinesDisplay
    
    SetFeedbackBanner "New quotation created. Enter customer details and click '+ ADD ITEM TO QUOTE'.", UI_GREEN_DARK
    
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
    Dim prevEvents As Boolean, prevScreen As Boolean
    
    On Error GoTo ErrorHandler
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    
    quoteId = Trim$(CStr(wsQ.Range("B4").Value))
    If Len(quoteId) = 0 Then
        quoteId = modSequence.AllocateEnquiryNumber()
        wsQ.Range("B4").Value = quoteId
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
    
    ' Read Custom Size toggle and Optional Manual Price Override (Rows 31-33)
    Dim isCustom As Boolean
    isCustom = (UCase$(Trim$(CStr(wsQ.Range("B31").Value))) = "YES")
    
    Dim manualPrice As Double
    Dim overrideReason As String
    manualPrice = Val(wsQ.Range("B32").Value)
    overrideReason = Trim$(CStr(wsQ.Range("B33").Value))
    If manualPrice > 0 And Len(overrideReason) = 0 Then
        overrideReason = "Manual selling price override"
    End If
    
    Dim priceByMT As Boolean
    priceByMT = (UCase$(Trim$(modQuote.GetQuotationPriceBasis(wsQ))) = "MT")
    
    ' Plant resolution from B13
    plt = Trim$(CStr(wsQ.Range("B13").Value))
    If InStr(1, plt, "Khopoli", vbTextCompare) > 0 Or InStr(1, plt, "P02", vbTextCompare) > 0 Then
        plantId = "P02"
    Else
        plantId = "P01"
    End If
    
    ' Validation
    If sz <= 0 Then
        SetFeedbackBanner "Validation Error: Please select or enter a valid Pipe Size (mm).", UI_RED
        Exit Sub
    End If
    If wt <= 0 Then
        SetFeedbackBanner "Validation Error: Please select or enter Wall Thickness (mm).", UI_RED
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
    
    Dim legacyMode As Boolean
    legacyMode = (InStr(1, CStr(wsQ.Range("B17").Value), "LEGACY", vbTextCompare) > 0 Or _
                  InStr(1, CStr(wsQ.Range("E12").Value), "LEGACY", vbTextCompare) > 0)
                  
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' ── BRANCH 1: EDIT MODE (Update existing line) ──
    If IsEditMode And Len(EditingLineID) > 0 Then
        Dim targetEditId As String
        targetEditId = EditingLineID
        
        modQuoteLines.UpdateQuoteLineFull _
            targetLineId:=targetEditId, _
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
            extCoating:=extC, _
            intCoating:=intC, _
            legacyMode:=legacyMode, _
            priceByMT:=priceByMT, _
            isCustomSize:=isCustom, _
            manualSellingPrice:=IIf(manualPrice > 0, manualPrice, 0), _
            overrideReason:=overrideReason
            
        IsEditMode = False
        EditingLineID = ""
        SelectedLineID = targetEditId
        
        On Error Resume Next
        wsQ.Shapes("btnQuoteAdd").TextFrame2.TextRange.Text = "+ ADD ITEM TO QUOTE"
        wsQ.Shapes("btnQuoteAdd").Fill.ForeColor.RGB = UI_GREEN_MID
        On Error GoTo 0
        
        ' Clear input form after editing
        wsQ.Range("B28").Value = 0
        wsQ.Range("B31").Value = "NO"
        wsQ.Range("B32").Value = ""
        wsQ.Range("B33").Value = ""
        modQuote.UpdateManualPriceUnitDisplay wsQ
        
        modQuote.RecalculateQuote
        RefreshQuoteLinesDisplay
        HideOutputSheets
        
        SetFeedbackBanner "Line " & targetEditId & " updated successfully! Quotation recalculated. Click 'CALCULATE QUOTATION' when ready.", UI_GREEN_DARK
        Application.EnableEvents = prevEvents
        Application.ScreenUpdating = prevScreen
        Exit Sub
    End If
    
    ' ── BRANCH 2: ADD NEW ITEM ──
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
        Application.EnableEvents = prevEvents
        Application.ScreenUpdating = prevScreen
        SetFeedbackBanner "Limit reached: Maximum 10 sizes allowed per quotation. Remove a line before adding.", UI_RED
        SafeMsgBox "Quotation capacity reached: A quotation supports up to 10 product lines." & vbCrLf & _
                   "Currently " & CStr(activeCount) & " of 10 sizes are added." & vbCrLf & _
                   "An 11th line cannot be added. Please remove an existing line or edit an existing line.", _
                   vbExclamation, "Mega Pipes - Line Limit"
        Exit Sub
    End If
    
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
        legacyMode:=legacyMode, _
        priceByMT:=priceByMT, _
        isCustomSize:=isCustom, _
        extCoating:=extC, _
        intCoating:=intC, _
        manualSellingPrice:=IIf(manualPrice > 0, manualPrice, 0), _
        overrideReason:=overrideReason
        
    Dim tblLinesAfter As ListObject
    Set tblLinesAfter = ThisWorkbook.Worksheets("QUOTE_LINES").ListObjects("tblQuoteLines")
    Dim newIdx As Long, newLineId As String
    newIdx = tblLinesAfter.ListRows.Count
    newLineId = Trim$(CStr(tblLinesAfter.ListRows(newIdx).Range.Cells(1, tblLinesAfter.ListColumns("LineID").Index).Value))
    SelectedLineID = newLineId
    
    ' Clear input cells for next entry
    wsQ.Range("B28").Value = 0
    wsQ.Range("B31").Value = "NO"
    wsQ.Range("B32").Value = ""
    wsQ.Range("B33").Value = ""
    modQuote.UpdateManualPriceUnitDisplay wsQ
    
    ' Recalculate Quote and refresh UI
    modQuote.RecalculateQuote
    RefreshQuoteLinesDisplay
    HideOutputSheets
    
    Dim totLines As Long
    totLines = activeCount + 1
    SetFeedbackBanner "Item " & newLineId & " added successfully! Total quote lines: " & CStr(totLines) & " of 10. Click 'CALCULATE QUOTATION' when ready.", UI_GREEN_DARK
    
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
        IsEditMode = False
        EditingLineID = ""
        SelectedLineID = ""
        On Error Resume Next
        wsQ.Shapes("btnQuoteAdd").TextFrame2.TextRange.Text = "+ ADD ITEM TO QUOTE"
        wsQ.Shapes("btnQuoteAdd").Fill.ForeColor.RGB = UI_GREEN_MID
        On Error GoTo 0
        
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
    Dim wsQ As Worksheet
    On Error GoTo ErrorHandler
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    
    If UCase$(Trim$(CStr(wsQ.Range("B31").Value))) = "YES" Then
        wsQ.Range("B31").Value = "NO"
        SetFeedbackBanner "Custom Pipe Size disabled (Standard Catalog validation active).", UI_GREEN_DARK
    Else
        wsQ.Range("B31").Value = "YES"
        SetFeedbackBanner "Custom Pipe Size enabled. Enter non-standard diameter (B23) and wall thickness (B24).", UI_ORANGE
        On Error Resume Next
        wsQ.Range("B23").Select
        On Error GoTo 0
    End If
    Exit Sub

ErrorHandler:
    SafeMsgBox "Custom size toggle error: " & Err.Description, vbExclamation
End Sub

Public Sub UI_RemoveItem()
    Dim wsQ As Worksheet
    Dim tbl As ListObject
    Dim targetId As String
    Dim actLines As Long, rL As ListRow
    
    On Error GoTo ErrorHandler
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    Set tbl = ThisWorkbook.Worksheets("QUOTE_LINES").ListObjects("tblQuoteLines")
    
    actLines = 0
    For Each rL In tbl.ListRows
        If Len(Trim$(CStr(rL.Range.Cells(1, tbl.ListColumns("ProductID").Index).Value))) > 0 Then
            actLines = actLines + 1
        End If
    Next rL
    
    If actLines = 0 Then
        SafeMsgBox "No quote lines available to remove.", vbInformation
        Exit Sub
    End If
    
    ' Determine target line
    If Len(SelectedLineID) > 0 Then
        targetId = SelectedLineID
    ElseIf actLines = 1 Then
        targetId = Trim$(CStr(tbl.ListRows(1).Range.Cells(1, tbl.ListColumns("LineID").Index).Value))
    Else
        If Application.DisplayAlerts And Application.UserControl Then
            targetId = InputBox("Enter Line ID to remove (e.g. L0001, L0002) or line number (e.g. 1, 2):", "Mega Pipes — Remove Quote Line", "1")
            targetId = Trim$(targetId)
            If Len(targetId) = 0 Then Exit Sub
            If IsNumeric(targetId) Then
                targetId = "L" & Format$(CLng(targetId), "0000")
            Else
                targetId = UCase$(targetId)
            End If
        Else
            targetId = "L0001"
        End If
    End If
    
    ' Confirmation
    If Application.DisplayAlerts And Application.UserControl Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("Are you sure you want to remove line " & targetId & " from this quotation?", _
                     vbYesNo + vbQuestion + vbDefaultButton2, "Mega Pipes — Remove Line")
        If ans <> vbYes Then Exit Sub
    End If
    
    ' Cancel edit mode if editing the removed line
    If IsEditMode And StrComp(EditingLineID, targetId, vbTextCompare) = 0 Then
        IsEditMode = False
        EditingLineID = ""
        On Error Resume Next
        wsQ.Shapes("btnQuoteAdd").TextFrame2.TextRange.Text = "+ ADD ITEM TO QUOTE"
        wsQ.Shapes("btnQuoteAdd").Fill.ForeColor.RGB = UI_GREEN_MID
        On Error GoTo 0
    End If
    
    Application.ScreenUpdating = False
    modQuote.DeleteQuoteLine targetId
    SelectedLineID = ""
    
    modQuote.RecalculateQuote
    RefreshQuoteLinesDisplay
    HideOutputSheets
    
    SetFeedbackBanner "Line " & targetId & " removed from quotation. Click 'CALCULATE QUOTATION' when ready.", UI_GREEN_DARK
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    SafeMsgBox "Error removing line: " & Err.Description, vbExclamation
End Sub

Public Sub UI_DuplicateItem()
    Dim wsQ As Worksheet
    Dim tbl As ListObject
    Dim targetId As String
    Dim actLines As Long, rL As ListRow
    
    On Error GoTo ErrorHandler
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    Set tbl = ThisWorkbook.Worksheets("QUOTE_LINES").ListObjects("tblQuoteLines")
    
    actLines = 0
    For Each rL In tbl.ListRows
        If Len(Trim$(CStr(rL.Range.Cells(1, tbl.ListColumns("ProductID").Index).Value))) > 0 Then
            actLines = actLines + 1
        End If
    Next rL
    
    If actLines = 0 Then
        SafeMsgBox "No quote lines available to duplicate.", vbInformation
        Exit Sub
    End If
    
    If actLines >= 10 Then
        SetFeedbackBanner "Cannot duplicate: Maximum 10 sizes already reached.", UI_RED
        SafeMsgBox "Quotation capacity reached (10 of 10 sizes). Remove a line before duplicating.", vbExclamation, "Mega Pipes — Line Limit"
        Exit Sub
    End If
    
    ' Determine target line
    If Len(SelectedLineID) > 0 Then
        targetId = SelectedLineID
    ElseIf actLines = 1 Then
        targetId = Trim$(CStr(tbl.ListRows(1).Range.Cells(1, tbl.ListColumns("LineID").Index).Value))
    Else
        Dim defaultId As String
        defaultId = Trim$(CStr(tbl.ListRows(actLines).Range.Cells(1, tbl.ListColumns("LineID").Index).Value))
        If Application.DisplayAlerts And Application.UserControl Then
            targetId = InputBox("Enter Line ID to duplicate (e.g. L0001, L0002) or line number (e.g. 1, 2):", "Mega Pipes — Duplicate Item", defaultId)
            targetId = Trim$(targetId)
            If Len(targetId) = 0 Then Exit Sub
            If IsNumeric(targetId) Then
                targetId = "L" & Format$(CLng(targetId), "0000")
            Else
                targetId = UCase$(targetId)
            End If
        Else
            targetId = defaultId
        End If
    End If
    
    ' If in Edit Mode, exit edit mode first
    If IsEditMode Then
        IsEditMode = False
        EditingLineID = ""
        On Error Resume Next
        wsQ.Shapes("btnQuoteAdd").TextFrame2.TextRange.Text = "+ ADD ITEM TO QUOTE"
        wsQ.Shapes("btnQuoteAdd").Fill.ForeColor.RGB = UI_GREEN_MID
        On Error GoTo 0
    End If
    
    Application.ScreenUpdating = False
    Dim newId As String
    newId = modQuoteLines.DuplicateQuoteLine(targetId)
    SelectedLineID = newId
    
    modQuote.RecalculateQuote
    RefreshQuoteLinesDisplay
    
    SetFeedbackBanner "Line " & targetId & " duplicated as " & newId & ". Total lines: " & CStr(actLines + 1) & " of 10.", UI_GREEN_DARK
    Application.ScreenUpdating = True
    SafeMsgBox "Line " & targetId & " duplicated successfully as " & newId & "!" & vbCrLf & _
               "Both lines are completely independent and can be edited separately.", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    SafeMsgBox "Could not duplicate line: " & Err.Description, vbExclamation
End Sub

Public Sub UI_EditItem()
    Dim wsQ As Worksheet
    Dim tbl As ListObject
    Dim rObj As ListRow, targetRow As ListRow
    Dim targetId As String
    
    On Error GoTo ErrorHandler
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    Set tbl = ThisWorkbook.Worksheets("QUOTE_LINES").ListObjects("tblQuoteLines")
    
    ' Count active lines
    Dim actLines As Long, rL As ListRow
    actLines = 0
    For Each rL In tbl.ListRows
        If Len(Trim$(CStr(rL.Range.Cells(1, tbl.ListColumns("ProductID").Index).Value))) > 0 Then
            actLines = actLines + 1
        End If
    Next rL
    
    If actLines = 0 Then
        SafeMsgBox "No quotation lines available to edit. Add a line first.", vbInformation
        Exit Sub
    End If
    
    ' Determine target line to edit
    If Len(SelectedLineID) > 0 Then
        targetId = SelectedLineID
    ElseIf actLines = 1 Then
        targetId = Trim$(CStr(tbl.ListRows(1).Range.Cells(1, tbl.ListColumns("LineID").Index).Value))
    Else
        Dim defaultId As String
        defaultId = Trim$(CStr(tbl.ListRows(1).Range.Cells(1, tbl.ListColumns("LineID").Index).Value))
        If Application.DisplayAlerts And Application.UserControl Then
            targetId = InputBox("Enter Line ID to edit (e.g. L0001, L0002) or line number (e.g. 1, 2):", "Mega Pipes — Edit Quote Line", defaultId)
            targetId = Trim$(targetId)
            If Len(targetId) = 0 Then Exit Sub
            If IsNumeric(targetId) Then
                targetId = "L" & Format$(CLng(targetId), "0000")
            Else
                targetId = UCase$(targetId)
            End If
        Else
            targetId = defaultId
        End If
    End If
    
    ' Locate the row in tblQuoteLines
    For Each rObj In tbl.ListRows
        If StrComp(Trim$(CStr(rObj.Range.Cells(1, tbl.ListColumns("LineID").Index).Value)), targetId, vbTextCompare) = 0 Then
            Set targetRow = rObj
            Exit For
        End If
    Next rObj
    
    If targetRow Is Nothing Then
        SafeMsgBox "Line " & targetId & " not found in quotation lines.", vbExclamation
        Exit Sub
    End If
    
    ' Read fields from target row
    Dim fam As String, bas As String, sz As Double, wt As Double, grd As String
    Dim extC As String, intC As String, qty As Double, unt As String
    Dim isCustomPersisted As Boolean, manPriceVal As Double, curOverrideReason As String
    
    Dim colCustom As Long, colManPrice As Long, colReas As Long
    colCustom = 0: colManPrice = 0: colReas = 0
    On Error Resume Next
    colCustom = tbl.ListColumns("IsCustomSize").Index
    colManPrice = tbl.ListColumns("ManualSellingPrice").Index
    colReas = tbl.ListColumns("OverrideReason").Index
    On Error GoTo 0
    
    With targetRow.Range
        fam = Trim$(CStr(.Cells(1, tbl.ListColumns("Family").Index).Value))
        bas = Trim$(CStr(.Cells(1, tbl.ListColumns("Basis").Index).Value))
        sz = Val(.Cells(1, tbl.ListColumns("SizeMM").Index).Value)
        wt = Val(.Cells(1, tbl.ListColumns("ThicknessMM").Index).Value)
        grd = Trim$(CStr(.Cells(1, tbl.ListColumns("GradeID").Index).Value))
        extC = Trim$(CStr(.Cells(1, tbl.ListColumns("CoatingExternalID").Index).Value))
        intC = Trim$(CStr(.Cells(1, tbl.ListColumns("CoatingInternalID").Index).Value))
        qty = Val(.Cells(1, tbl.ListColumns("Quantity").Index).Value)
        unt = Trim$(CStr(.Cells(1, tbl.ListColumns("QuantityUnit").Index).Value))
        If colCustom > 0 Then
            isCustomPersisted = (StrComp(Trim$(CStr(.Cells(1, colCustom).Value)), "YES", vbTextCompare) = 0)
        Else
            isCustomPersisted = False
        End If
        If colManPrice > 0 Then manPriceVal = Val(.Cells(1, colManPrice).Value)
        If colReas > 0 Then curOverrideReason = Trim$(CStr(.Cells(1, colReas).Value))
    End With
    If Len(extC) = 0 Then extC = "Bare"
    If Len(intC) = 0 Then intC = "None"
    If Len(unt) = 0 Then unt = "Metres"
    
    ' Populate input form Box 2 (Rows 20-33)
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    
    wsQ.Range("B21").Value = fam
    wsQ.Range("B22").Value = bas
    wsQ.Range("B23").Value = sz
    wsQ.Range("B24").Value = wt
    wsQ.Range("B25").Value = grd
    wsQ.Range("B26").Value = extC
    wsQ.Range("B27").Value = intC
    wsQ.Range("B28").Value = qty
    wsQ.Range("B29").Value = unt
    wsQ.Range("B30").Value = 12#
    wsQ.Range("B31").Value = IIf(isCustomPersisted, "YES", "NO")
    If manPriceVal > 0 Then
        wsQ.Range("B32").Value = manPriceVal
    Else
        wsQ.Range("B32").Value = ""
    End If
    wsQ.Range("B33").Value = curOverrideReason
    modQuote.UpdateManualPriceUnitDisplay wsQ
    
    Application.EnableEvents = prevEvents
    
    ' Enter Edit Mode
    IsEditMode = True
    EditingLineID = targetId
    SelectedLineID = targetId
    
    ' Update button appearance to amber UPDATE LINE
    On Error Resume Next
    wsQ.Shapes("btnQuoteAdd").TextFrame2.TextRange.Text = "UPDATE LINE (" & targetId & ")"
    wsQ.Shapes("btnQuoteAdd").Fill.ForeColor.RGB = RGB(220, 120, 0)
    On Error GoTo 0
    
    ' Visually highlight the line in the table
    Dim lineIdx As Long
    lineIdx = targetRow.Index
    HandleLineSelection 20 + lineIdx
    
    SetFeedbackBanner "Editing line " & targetId & " (Custom: " & IIf(isCustomPersisted, "YES", "NO") & "). Adjust values and click 'UPDATE LINE (" & targetId & ")'.", RGB(180, 90, 0)
    
    On Error Resume Next
    wsQ.Range("B28").Select
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    SafeMsgBox "Error loading line for editing: " & Err.Description, vbExclamation
End Sub

Public Sub UI_ClearLine()
    Dim wsQ As Worksheet
    On Error GoTo ErrorHandler
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    
    ' If in Edit Mode, cancel edit mode
    If IsEditMode Then
        IsEditMode = False
        EditingLineID = ""
        On Error Resume Next
        wsQ.Shapes("btnQuoteAdd").TextFrame2.TextRange.Text = "+ ADD ITEM TO QUOTE"
        wsQ.Shapes("btnQuoteAdd").Fill.ForeColor.RGB = UI_GREEN_MID
        On Error GoTo 0
        RefreshQuoteLinesDisplay
    End If
    
    ' Clear line input fields and reset inputs to clean defaults
    wsQ.Range("B23").Value = 0          ' Pipe Size (mm)
    wsQ.Range("B24").Value = 0          ' Wall Thickness (mm)
    wsQ.Range("B26").Value = "Bare"     ' External Coating
    wsQ.Range("B27").Value = "None"     ' Internal Coating
    wsQ.Range("B28").Value = 0          ' Quantity
    wsQ.Range("B30").Value = 12#        ' Pipe Length
    wsQ.Range("B31").Value = "NO"       ' Custom Pipe Size toggle
    wsQ.Range("C31").Value = ""         ' Custom Size Value
    wsQ.Range("B32").Value = ""         ' Manual Selling Price Override
    wsQ.Range("B33").Value = ""         ' Override Reason
    modQuote.UpdateManualPriceUnitDisplay wsQ
    
    SetFeedbackBanner "Input form cleared. Existing quotation lines preserved.", UI_GREEN_DARK
    Exit Sub
ErrorHandler:
    SafeMsgBox "Error clearing line inputs: " & Err.Description, vbExclamation
End Sub

Public Sub UI_ManualPriceOverride()
    Dim wsQ As Worksheet
    Dim tblLines As ListObject
    Dim rObj As ListRow, targetRow As ListRow
    Dim targetId As String
    Dim actLines As Long, rL As ListRow
    
    On Error GoTo ErrorHandler
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    Set tblLines = ThisWorkbook.Worksheets("QUOTE_LINES").ListObjects("tblQuoteLines")
    
    actLines = 0
    For Each rL In tblLines.ListRows
        If Len(Trim$(CStr(rL.Range.Cells(1, tblLines.ListColumns("ProductID").Index).Value))) > 0 Then
            actLines = actLines + 1
        End If
    Next rL
    
    If actLines = 0 Then
        SafeMsgBox "No quotation lines available to override price. Add a line first.", vbInformation
        Exit Sub
    End If
    
    ' Determine target line
    If Len(SelectedLineID) > 0 Then
        targetId = SelectedLineID
    ElseIf actLines = 1 Then
        targetId = Trim$(CStr(tblLines.ListRows(1).Range.Cells(1, tblLines.ListColumns("LineID").Index).Value))
    Else
        Dim defaultId As String
        defaultId = Trim$(CStr(tblLines.ListRows(1).Range.Cells(1, tblLines.ListColumns("LineID").Index).Value))
        If Application.DisplayAlerts And Application.UserControl Then
            targetId = InputBox("Enter Line ID to override price (e.g. L0001, L0002) or line number (e.g. 1, 2):", "Mega Pipes — Manual Price Override", defaultId)
            targetId = Trim$(targetId)
            If Len(targetId) = 0 Then Exit Sub
            If IsNumeric(targetId) Then
                targetId = "L" & Format$(CLng(targetId), "0000")
            Else
                targetId = UCase$(targetId)
            End If
        Else
            targetId = defaultId
        End If
    End If
    
    ' Find row in tblQuoteLines
    For Each rObj In tblLines.ListRows
        If StrComp(Trim$(CStr(rObj.Range.Cells(1, tblLines.ListColumns("LineID").Index).Value)), targetId, vbTextCompare) = 0 Then
            Set targetRow = rObj
            Exit For
        End If
    Next rObj
    
    If targetRow Is Nothing Then
        SafeMsgBox "Line " & targetId & " not found.", vbExclamation
        Exit Sub
    End If
    
    Dim calcPrice As Double, curPrice As Double, curManual As Double
    Dim curReason As String, pSrc As String
    With targetRow.Range
        calcPrice = Val(.Cells(1, tblLines.ListColumns("CalculatedUnitPrice").Index).Value)
        If calcPrice = 0 Then calcPrice = Val(.Cells(1, tblLines.ListColumns("UnitPrice").Index).Value)
        curPrice = Val(.Cells(1, tblLines.ListColumns("UnitPrice").Index).Value)
        curManual = Val(.Cells(1, tblLines.ListColumns("ManualSellingPrice").Index).Value)
        curReason = Trim$(CStr(.Cells(1, tblLines.ListColumns("OverrideReason").Index).Value))
        pSrc = Trim$(CStr(.Cells(1, tblLines.ListColumns("PriceSource").Index).Value))
    End With
    
    Dim pBasis As String, priceUnitLabel As String
    pBasis = modQuote.GetQuotationPriceBasis(wsQ)
    priceUnitLabel = modQuote.GetPriceBasisUnitLabel(pBasis)
    
    Dim promptMsg As String
    promptMsg = "Line " & targetId & " Selling Price Configuration (" & priceUnitLabel & "):" & vbCrLf & vbCrLf & _
                "• System Calculated Unit Price: " & priceUnitLabel & " " & Format$(calcPrice, "#,##0.00") & vbCrLf & _
                "• Current Applied Unit Price:   " & priceUnitLabel & " " & Format$(curPrice, "#,##0.00") & " (" & IIf(pSrc = "MANUAL_OVERRIDE", "MANUAL OVERRIDE", "CALCULATED") & ")" & vbCrLf & vbCrLf & _
                "Enter new Selling Price in " & priceUnitLabel & ":" & vbCrLf & _
                "(Enter 0 or leave empty to RESTORE system calculated price)"
                
    Dim inputPriceStr As String
    If Application.DisplayAlerts And Application.UserControl Then
        inputPriceStr = InputBox(promptMsg, "Mega Pipes — Manual Selling Price Override", IIf(curManual > 0, CStr(curManual), CStr(calcPrice)))
        If Len(Trim$(inputPriceStr)) = 0 Then Exit Sub
    Else
        Exit Sub
    End If
    
    If Not IsNumeric(inputPriceStr) Then
        SafeMsgBox "Invalid price entered. Must be a numeric amount.", vbExclamation
        Exit Sub
    End If
    
    Dim newPrice As Double
    newPrice = CDbl(inputPriceStr)
    
    Dim quoteId As String
    quoteId = Trim$(CStr(wsQ.Range("B4").Value))
    If Len(quoteId) = 0 Then quoteId = "QUOTE-DEV"
    
    If newPrice <= 0 Or Abs(newPrice - calcPrice) < 0.001 Then
        ' Reset to calculated
        targetRow.Range.Cells(1, tblLines.ListColumns("ManualSellingPrice").Index).Value = 0#
        targetRow.Range.Cells(1, tblLines.ListColumns("PriceSource").Index).Value = "CALCULATED"
        targetRow.Range.Cells(1, tblLines.ListColumns("PriceVariance").Index).Value = 0#
        targetRow.Range.Cells(1, tblLines.ListColumns("OverrideReason").Index).Value = ""
        
        modApproval.InvalidateApprovalIfModified "Manual price override removed on " & targetId
        modLog.LogAction quoteId, "PRICE_OVERRIDE_RESET", Format$(curPrice, "0.00"), Format$(calcPrice, "0.00"), targetId, "RESTORED", "Restored to formula-calculated price"
        
        modQuote.RecalculateQuote
        RefreshQuoteLinesDisplay
        SetFeedbackBanner "Manual price override cleared for " & targetId & ". Restored calculated price Rs. " & Format$(calcPrice, "#,##0.00"), UI_GREEN_DARK
        SafeMsgBox "Manual price override removed. Line " & targetId & " restored to system-calculated price Rs. " & Format$(calcPrice, "#,##0.00"), vbInformation
        Exit Sub
    End If
    
    ' New manual price requires reason
    Dim reasonPrompt As String, reasonInput As String
    reasonPrompt = "Mandatory Approval / Audit Policy:" & vbCrLf & _
                   "A commercial reason is required for manual price override on " & targetId & "." & vbCrLf & vbCrLf & _
                   "Calculated Price: Rs. " & Format$(calcPrice, "#,##0.00") & vbCrLf & _
                   "New Manual Price: Rs. " & Format$(newPrice, "#,##0.00") & vbCrLf & _
                   "Variance:         Rs. " & Format$(newPrice - calcPrice, "+#,##0.00;-#,##0.00;0.00") & vbCrLf & vbCrLf & _
                   "Enter approval justification / reason:"
                   
    reasonInput = InputBox(reasonPrompt, "Mega Pipes — Override Justification", IIf(Len(curReason) > 0, curReason, "Commercial management concession"))
    reasonInput = Trim$(reasonInput)
    If Len(reasonInput) = 0 Then
        SafeMsgBox "Price override cancelled: A valid reason is required by company approval policy.", vbExclamation
        Exit Sub
    End If
    
    targetRow.Range.Cells(1, tblLines.ListColumns("ManualSellingPrice").Index).Value = newPrice
    targetRow.Range.Cells(1, tblLines.ListColumns("PriceSource").Index).Value = "MANUAL_OVERRIDE"
    targetRow.Range.Cells(1, tblLines.ListColumns("PriceVariance").Index).Value = newPrice - calcPrice
    targetRow.Range.Cells(1, tblLines.ListColumns("OverrideReason").Index).Value = reasonInput
    
    modApproval.InvalidateApprovalIfModified "Manual price override on " & targetId & ": Rs. " & Format$(newPrice, "#,##0.00") & " (" & reasonInput & ")"
    modLog.LogAction quoteId, "PRICE_OVERRIDE", Format$(calcPrice, "0.00"), Format$(newPrice, "0.00"), targetId, "APPLIED", reasonInput
    
    modQuote.RecalculateQuote
    RefreshQuoteLinesDisplay
    
    Dim diffStr As String
    diffStr = Format$(newPrice - calcPrice, "+#,##0.00;-#,##0.00;0.00")
    SetFeedbackBanner "Line " & targetId & " price overridden to Rs. " & Format$(newPrice, "#,##0.00") & " (Variance: Rs. " & diffStr & "). Quotation updated.", RGB(180, 90, 0)
    SafeMsgBox "Manual price override applied to " & targetId & "!" & vbCrLf & vbCrLf & _
               "• Calculated Price: Rs. " & Format$(calcPrice, "#,##0.00") & vbCrLf & _
               "• Applied Price:    Rs. " & Format$(newPrice, "#,##0.00") & vbCrLf & _
               "• Variance:         Rs. " & diffStr & vbCrLf & _
               "• Reason:           " & reasonInput & vbCrLf & vbCrLf & _
               "Quotation totals and COSTING sheet updated.", vbInformation
    Exit Sub
    
ErrorHandler:
    SafeMsgBox "Error setting manual price override: " & Err.Description, vbExclamation
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
    Dim wsQ As Worksheet
    On Error GoTo ErrorHandler
    If Application.DisplayAlerts And Application.UserControl Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("Are you sure you want to reset this draft?" & vbCrLf & vbCrLf & _
                     "This will clear unsaved quotation inputs and lines." & vbCrLf & _
                     "(Previously saved quotations and snapshots will NOT be affected.)", _
                     vbYesNo + vbQuestion + vbDefaultButton2, "Mega Pipes - Reset Draft")
        If ans <> vbYes Then Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Reset edit mode & selection
    IsEditMode = False
    EditingLineID = ""
    SelectedLineID = ""
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    On Error Resume Next
    wsQ.Shapes("btnQuoteAdd").TextFrame2.TextRange.Text = "+ ADD ITEM TO QUOTE"
    wsQ.Shapes("btnQuoteAdd").Fill.ForeColor.RGB = UI_GREEN_MID
    On Error GoTo 0
    
    modQuote.ClearQuoteLines
    InvalidateOutputSheets
    modQuoteHeader.InitializeNewQuote forceNew:=False
    
    UI_ClearLine
    
    RefreshQuoteLinesDisplay
    SetFeedbackBanner "Quotation draft reset. Enter new quotation details.", UI_GREEN_DARK
    Application.ScreenUpdating = True
    SafeMsgBox "Quotation draft reset successfully.", vbInformation
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    SafeMsgBox "Error resetting draft: " & Err.Description, vbExclamation
End Sub

Public Sub UI_ClearAllInputs(Optional ByVal promptUser As Boolean = False)
    Dim wsQ As Worksheet
    Dim prevEvents As Boolean, prevScreen As Boolean
    
    On Error GoTo ErrorHandler
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    
    If promptUser And Application.DisplayAlerts And Application.UserControl Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("Are you sure you want to clear all quotation inputs, lines, and outputs?" & vbCrLf & vbCrLf & _
                     "(Settings, master data, configuration, and quote logs will be preserved.)", _
                     vbYesNo + vbQuestion + vbDefaultButton2, "Mega Pipes — Clear All Inputs")
        If ans <> vbYes Then Exit Sub
    End If
    
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    modProtection.UnprotectSheet wsQ
    
    ' 1. Clear customer & project header details (B8:B12) and reset dropdown fields
    wsQ.Range("B8").Value = ""          ' Customer
    wsQ.Range("B9").Value = ""          ' Contact Person
    wsQ.Range("B10").Value = ""         ' Delivery Place
    wsQ.Range("B11").Value = ""         ' Project / Tender Ref
    wsQ.Range("B12").Value = ""         ' Salesperson
    wsQ.Range("B13").Value = "P01"      ' Supplying Plant
    wsQ.Range("B14").Value = "30 Days from Dispatch" ' Payment Terms
    wsQ.Range("B15").Value = 30         ' Validity (Days)
    wsQ.Range("B16").Value = "Metres"   ' Price Basis
    wsQ.Range("B17").Value = "NORMAL"   ' Pricing Mode
    
    ' 2. Clear product line inputs form (Rows 20-33)
    wsQ.Range("B20").Value = "IS-3589-2001"
    wsQ.Range("B21").Value = "HSAW"
    wsQ.Range("B22").Value = "ID"
    wsQ.Range("B23").Value = 0          ' Size (mm)
    wsQ.Range("B24").Value = 0          ' Wall Thickness (mm)
    wsQ.Range("B25").Value = "Fe 410"
    wsQ.Range("B26").Value = "Bare"     ' External Coating
    wsQ.Range("B27").Value = "None"     ' Internal Coating
    wsQ.Range("B28").Value = 0          ' Quantity
    wsQ.Range("B29").Value = "Metres"
    wsQ.Range("B30").Value = 12#        ' Pipe Length
    wsQ.Range("B31").Value = "NO"       ' Custom Pipe Size toggle
    wsQ.Range("C31").Value = ""         ' Custom Size Value
    wsQ.Range("B32").Value = ""         ' Manual Selling Price override
    wsQ.Range("B33").Value = ""         ' Override Reason
    modQuote.UpdateManualPriceUnitDisplay wsQ
    
    ' 3. Clear all quotation line items from tblQuoteLines (includes ManualFC, TripRate, ManualTrips)
    modQuote.ClearQuoteLines
    
    ' 4. Invalidate and hide commercial output sheets (OFFER and SYNOPSIS)
    InvalidateOutputSheets
    
    ' 5. Reset line edit mode and selection state
    IsEditMode = False
    EditingLineID = ""
    SelectedLineID = ""
    On Error Resume Next
    wsQ.Shapes("btnQuoteAdd").TextFrame2.TextRange.Text = "+ ADD ITEM TO QUOTE"
    wsQ.Shapes("btnQuoteAdd").Fill.ForeColor.RGB = UI_GREEN_MID
    On Error GoTo 0
    
    ' 6. Refresh line item display and dependent dropdowns
    modDependentDropdowns.SetupDependentDropdownArchitecture
    modDependentDropdowns.SetupQuoteLineDropdowns
    RefreshQuoteLinesDisplay
    
    ' 7. Feedback banner
    SetFeedbackBanner "All user inputs, quotation lines, and outputs have been cleared.", UI_GREEN_DARK
    
    modProtection.ProtectSheet wsQ, userInterfaceOnly:=True
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    modProtection.ProtectSheet wsQ, userInterfaceOnly:=True
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    SafeMsgBox "Error clearing all inputs: " & Err.Description, vbExclamation
End Sub

Public Sub UI_ClearAllItems()
    UI_ClearAllInputs promptUser:=True
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
    Dim wsQ As Worksheet
    
    On Error GoTo ErrorHandler
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    
    ' 1. Validate Required Customer Input
    Dim cust As String
    cust = Trim$(CStr(wsQ.Range("B8").Value))
    If Len(cust) = 0 Then
        HideOutputSheets
        SetFeedbackBanner "Validation Error: Customer is required. Select or enter Customer in B8 before calculating.", UI_RED
        SafeMsgBox "Please select or enter a Customer in cell B8 before calculating.", vbExclamation, "Mega Pipes — Validation Required"
        Exit Sub
    End If
    
    ' 2. Validate Required Quotation Line Items
    Dim tblLines As ListObject
    Set tblLines = ThisWorkbook.Worksheets("QUOTE_LINES").ListObjects("tblQuoteLines")
    Dim validLines As Long, rL As ListRow
    validLines = 0
    If Not tblLines Is Nothing Then
        For Each rL In tblLines.ListRows
            If Len(Trim$(CStr(rL.Range.Cells(1, tblLines.ListColumns("ProductID").Index).Value))) > 0 Then
                If Val(rL.Range.Cells(1, tblLines.ListColumns("Quantity").Index).Value) > 0 Then
                    validLines = validLines + 1
                End If
            End If
        Next rL
    End If
    
    If validLines = 0 Then
        HideOutputSheets
        SetFeedbackBanner "Validation Error: No line items found. Add at least one valid pipe item before calculating.", UI_RED
        SafeMsgBox "Please add at least one valid line item to the quotation before calculating.", vbExclamation, "Mega Pipes — Validation Required"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Unprotect calculated sheets for macro writing
    modProtection.UnprotectCalculatedSheetsForMacro
    
    ' 3. Recalculate quotation using existing approved engine
    modQuote.RecalculateQuote
    
    ' 4. Generate COSTING, OFFER, and SYNOPSIS reports
    modExport.WriteCostingSheet
    modExport.WriteOfferSheet
    modSynopsis.PopulateSynopsisReport
    
    ' 5. Refresh display table on QUOTE sheet
    RefreshQuoteLinesDisplay
    
    ' 5b. Show distance note panel (nearest MPPL plant & nearest competitor)
    ShowDistanceNoteOnQuoteSheet
    
    ' 6. Lock calculated sheets and protect against manual editing
    modProtection.ProtectCalculatedSheets True
    
    ' 7. Reveal OFFER and SYNOPSIS sheets now that calculation and generation succeeded
    ShowOutputSheets
    
    Dim grandTotal As String
    grandTotal = "Rs. " & Format$(Val(wsQ.Range("E10").Value), "#,##0")
    SetFeedbackBanner "Quotation calculated successfully! Grand Total: " & grandTotal & ". OFFER and SYNOPSIS are now available.", UI_GREEN_DARK
    
    Application.ScreenUpdating = True
    SafeMsgBox "Quotation calculated successfully!" & vbCrLf & vbCrLf & _
               "Customer: " & cust & vbCrLf & _
               "Subtotal (Excl GST): Rs. " & Format$(Val(wsQ.Range("E8").Value), "#,##0") & vbCrLf & _
               "Grand Total (Incl GST): " & grandTotal & vbCrLf & vbCrLf & _
               "OFFER and SYNOPSIS reports have been updated and are now available for review.", vbInformation, "Mega Pipes — Quotation Calculated"
    Exit Sub

ErrorHandler:
    On Error Resume Next
    HideOutputSheets
    modProtection.ProtectCalculatedSheets True
    Application.ScreenUpdating = True
    SetFeedbackBanner "Calculation error: " & Err.Description, UI_RED
    SafeMsgBox "Calculation failed: " & Err.Description & vbCrLf & vbCrLf & _
               "OFFER and SYNOPSIS sheets remain hidden to prevent exposing uncalculated results.", _
               vbCritical, "Mega Pipes — Calculation Error"
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

Public Sub UI_SubmitReview()
    On Error GoTo ErrorHandler
    Dim wsQ As Worksheet
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    
    Dim qNum As String
    qNum = Trim$(CStr(wsQ.Range("B4").Value))
    If Len(qNum) = 0 Then
        SafeMsgBox "Please create or calculate a quotation before submitting for review.", vbExclamation
        Exit Sub
    End If
    
    Dim curState As String
    curState = modApproval.GetQuoteState()
    
    If curState = "Draft" Or curState = "" Then
        modQuote.RecalculateQuote
        curState = modApproval.GetQuoteState()
    End If
    
    If curState <> "Calculated" And curState <> "Review" Then
        SafeMsgBox "Quotation cannot be submitted for review from status: " & curState, vbExclamation
        Exit Sub
    End If
    
    Dim res As String
    res = modApproval.TransitionQuoteState("Review", "Submitted for review via UI")
    If Left$(res, 5) = "ERROR" Then
        SafeMsgBox "Could not submit for review: " & res, vbCritical
        Exit Sub
    End If
    
    wsQ.Range("E11").Value = "Review"
    wsQ.Range("E11").Interior.Color = RGB(220, 240, 255)
    wsQ.Range("E11").Font.Color = RGB(16, 44, 96)
    wsQ.Range("E11").Font.Bold = True
    
    SetFeedbackBanner "Quotation " & qNum & " submitted for review.", UI_BLUE
    SafeMsgBox "Quotation " & qNum & " has been submitted for review." & vbCrLf & _
               "Authorized reviewers may now review and approve this quotation.", vbInformation
    Exit Sub

ErrorHandler:
    SafeMsgBox "Error submitting for review: " & Err.Description, vbCritical
End Sub

Public Sub UI_ApproveQuotation()
    On Error GoTo ErrorHandler
    Dim wsQ As Worksheet
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    
    Dim qNum As String
    qNum = Trim$(CStr(wsQ.Range("B4").Value))
    If Len(qNum) = 0 Then
        SafeMsgBox "Please allocate a Quote Number before approving.", vbExclamation
        Exit Sub
    End If
    
    Dim curState As String
    curState = modApproval.GetQuoteState()
    
    If curState = "Approved" Then
        SafeMsgBox "This quotation is ALREADY approved." & vbCrLf & _
                   "To modify, make your changes to create a new revision.", vbInformation
        Exit Sub
    End If
    
    If curState = "Draft" Or curState = "" Then
        modQuote.RecalculateQuote
        curState = modApproval.GetQuoteState()
    End If
    
    If curState = "Calculated" Then
        modApproval.TransitionQuoteState "Review", "Auto-transition to review before approval"
        curState = "Review"
    End If
    
    If curState <> "Review" Then
        SafeMsgBox "Cannot approve quotation in status: " & curState, vbExclamation
        Exit Sub
    End If
    
    Dim grandTotal As String
    grandTotal = "Rs. " & Format$(Val(wsQ.Range("E10").Value), "#,##0")
    Dim answer As VbMsgBoxResult
    If Application.DisplayAlerts Then
        answer = MsgBox("Are you sure you want to APPROVE this quotation?" & vbCrLf & vbCrLf & _
                        "Quote ID: " & qNum & vbCrLf & _
                        "Customer: " & Trim$(CStr(wsQ.Range("B8").Value)) & vbCrLf & _
                        "Grand Total: " & grandTotal & vbCrLf & vbCrLf & _
                        "Approval will preserve an immutable audit snapshot.", _
                        vbQuestion + vbYesNo, "Mega Pipes - Approve Quotation")
        If answer <> vbYes Then Exit Sub
    End If
    
    Dim pwd As String
    Dim expectedPwd As String
    pwd = ""
    expectedPwd = modAdmin.AdminPassword()
    If Len(expectedPwd) > 0 And Application.DisplayAlerts Then
        pwd = InputBox("Enter Manager / Admin Approval Password:", "Approval Authorization Required")
        If Len(pwd) = 0 Then Exit Sub
    End If
    
    Dim res As String
    res = modApproval.TransitionQuoteState("Approved", "Approved inside software by " & modLog.GetCurrentUser(), pwd)
    If Left$(res, 5) = "ERROR" Then
        SafeMsgBox "Approval failed: " & res, vbCritical
        Exit Sub
    End If
    
    Dim snapPath As String
    snapPath = modLog.SaveQuote(qNum)
    
    wsQ.Range("E11").Value = "Approved"
    wsQ.Range("E11").Interior.Color = UI_GREEN_DARK
    wsQ.Range("E11").Font.Color = UI_WHITE
    wsQ.Range("E11").Font.Bold = True
    
    SetFeedbackBanner "Quotation APPROVED! Immutable snapshot preserved: " & Dir(snapPath), UI_GREEN_DARK
    SafeMsgBox "Quotation " & qNum & " has been APPROVED!" & vbCrLf & vbCrLf & _
               "Immutable Snapshot: " & snapPath & vbCrLf & _
               "Grand Total: " & grandTotal & vbCrLf & vbCrLf & _
               "Note: Any subsequent line edits will automatically invalidate this approval and require re-approval.", _
               vbInformation
    Exit Sub

ErrorHandler:
    SafeMsgBox "Error approving quotation: " & Err.Description, vbCritical
End Sub

Public Sub UI_PreviewOffer()
    On Error GoTo ErrorHandler
    Dim wsOff As Worksheet
    Set wsOff = ThisWorkbook.Worksheets("OFFER")
    
    If wsOff.Visible <> xlSheetVisible Then
        SafeMsgBox "The Offer Letter is not available yet." & vbCrLf & vbCrLf & _
                   "Please click 'CALCULATE QUOTATION' to calculate the quotation and generate the offer first.", _
                   vbInformation, "Mega Pipes — Offer Preview"
        Exit Sub
    End If
    
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
    Set wsSyn = ThisWorkbook.Worksheets("SYNOPSIS")
    
    If wsSyn.Visible <> xlSheetVisible Then
        SafeMsgBox "The Synopsis Report is not available yet." & vbCrLf & vbCrLf & _
                   "Please click 'CALCULATE QUOTATION' to calculate the quotation and generate the report first.", _
                   vbInformation, "Mega Pipes — Synopsis Preview"
        Exit Sub
    End If
    
    modSynopsis.PopulateSynopsisReport
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
    Dim exportedPath As String
    Dim fso As Object
    
    qNum = Trim$(CStr(ThisWorkbook.Worksheets("QUOTE").Range("B4").Value))
    If Len(qNum) = 0 Then
        SafeMsgBox "No active quotation to export. Please create or open a quotation first.", vbInformation
        Exit Sub
    End If
    
    exportedPath = modExport.ExportPDF()
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(exportedPath) = 0 Or Not fso.FileExists(exportedPath) Then
        Err.Raise vbObjectError + 1911, "UI_ExportPDF", "PDF export verification failed: File was not created."
    End If
    
    SafeMsgBox "Quotation Offer Letter exported to PDF successfully!" & vbCrLf & vbCrLf & _
               "File: " & fso.GetFileName(exportedPath) & vbCrLf & _
               "Location: " & exportedPath, vbInformation
    Exit Sub

ErrorHandler:
    SafeMsgBox "PDF export failed: " & Err.Description, vbCritical
End Sub

Public Sub UI_ExportCustomerExcel()
    On Error GoTo ErrorHandler
    Dim qNum As String
    Dim resStr As String
    Dim valExpErr As String
    
    ' Validate calculation status before export
    If Not modExport.ValidateQuotationBeforeExport(valExpErr) Then
        SafeMsgBox valExpErr, vbExclamation, "Mega Pipes — Calculation Required Before Export"
        Exit Sub
    End If
    
    qNum = Trim$(CStr(ThisWorkbook.Worksheets("QUOTE").Range("B4").Value))
    If Len(qNum) = 0 Then
        SafeMsgBox "No active quotation to export. Please create or open a quotation first.", vbInformation
        Exit Sub
    End If
    
    ' Execute dual export flow: Central UNC + Local user copy
    resStr = modExport.ExportCustomerExcelDual(allowInteractivePicker:=True)
    
    ' Parse result string
    Dim parts() As String, kv() As String, i As Long
    Dim statusVal As String, quoteId As String, fName As String
    Dim centralStatus As String, centralPath As String, centralErr As String
    Dim localStatus As String, localPath As String, localErr As String
    Dim stagedPath As String, verifVal As String
    
    parts = Split(resStr, ";")
    For i = LBound(parts) To UBound(parts)
        kv = Split(parts(i), "=")
        If UBound(kv) >= 1 Then
            Select Case Trim$(kv(0))
                Case "Status": statusVal = Trim$(kv(1))
                Case "QuoteID": quoteId = Trim$(kv(1))
                Case "Filename": fName = Trim$(kv(1))
                Case "CentralStatus": centralStatus = Trim$(kv(1))
                Case "CentralPath": centralPath = Trim$(kv(1))
                Case "CentralError": centralErr = Trim$(kv(1))
                Case "LocalStatus": localStatus = Trim$(kv(1))
                Case "LocalPath": localPath = Trim$(kv(1))
                Case "LocalError": localErr = Trim$(kv(1))
                Case "StagedPath": stagedPath = Trim$(kv(1))
                Case "Verification": verifVal = Trim$(kv(1))
            End Select
        End If
    Next i
    
    Dim msg As String
    msg = "MEGA PIPES — CUSTOMER EXCEL EXPORT SUMMARY" & vbCrLf & vbCrLf & _
          "Quote ID: " & quoteId & vbCrLf & _
          "File Name: " & fName & vbCrLf & vbCrLf
          
    If centralStatus = "SUCCESS" Then
        msg = msg & "[CENTRAL ARCHIVE]: VERIFIED SUCCESS" & vbCrLf & _
              "Location: " & centralPath & vbCrLf & vbCrLf
    ElseIf centralStatus = "PENDING" Then
        msg = msg & "[CENTRAL ARCHIVE]: PENDING (OFFLINE / UNAVAILABLE)" & vbCrLf & _
              "Reason: " & centralErr & vbCrLf & _
              "Action: Staged safely in Pending_Central_Archive for automatic/manual retry." & vbCrLf & vbCrLf
    Else
        msg = msg & "[CENTRAL ARCHIVE]: FAILED" & vbCrLf & _
              "Reason: " & centralErr & vbCrLf & vbCrLf
    End If
    
    If localStatus = "SUCCESS" Then
        msg = msg & "[LOCAL USER COPY]: VERIFIED SUCCESS" & vbCrLf & _
              "Location: " & localPath & vbCrLf & vbCrLf
    ElseIf localStatus = "SKIPPED" Then
        msg = msg & "[LOCAL USER COPY]: SKIPPED / CANCELLED" & vbCrLf & vbCrLf
    Else
        msg = msg & "[LOCAL USER COPY]: FAILED" & vbCrLf & _
              "Reason: " & localErr & vbCrLf & vbCrLf
    End If
    
    msg = msg & "Confidentiality guaranteed: Internal cost formulas, margins, and rates excluded."
    
    If statusVal = "OK" Then
        SafeMsgBox msg, vbInformation, "Mega Pipes Export"
    ElseIf statusVal = "PARTIAL" Then
        SafeMsgBox msg, vbExclamation, "Mega Pipes Export — Partial Success"
        ' If central succeeded but local failed, give user the option to retry local
        If centralStatus = "SUCCESS" And localStatus = "FAILED" Then
            If Application.UserControl And Application.DisplayAlerts Then
                Dim retryAns As VbMsgBoxResult
                retryAns = MsgBox("Central export succeeded, but local copy failed." & vbCrLf & _
                                  "Would you like to select a new local directory and retry?", _
                                  vbQuestion + vbYesNo, "Retry Local Copy")
                If retryAns = vbYes Then
                    Dim newLocalDir As String
                    newLocalDir = modExport.GetLocalExportsDir(allowInteractivePicker:=True)
                    If Len(newLocalDir) > 0 Then
                        On Error Resume Next
                        Dim retryLocalPath As String
                        retryLocalPath = modExport.RetryLocalExport(centralPath, newLocalDir)
                        If Err.Number = 0 And Len(retryLocalPath) > 0 Then
                            SafeMsgBox "Local copy retry succeeded!" & vbCrLf & vbCrLf & "Location: " & retryLocalPath, vbInformation
                        Else
                            SafeMsgBox "Local copy retry failed: " & Err.Description, vbCritical
                        End If
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
    Else
        SafeMsgBox msg, vbCritical, "Mega Pipes Export Failed"
    End If
    Exit Sub

ErrorHandler:
    SafeMsgBox "Customer Excel export failed: " & Err.Description, vbCritical
End Sub

Public Sub UI_RetryCentralExport()
    On Error GoTo ErrorHandler
    Dim res As String
    res = modExport.RetryPendingCentralExports()
    SafeMsgBox "Central Export Retry Result:" & vbCrLf & vbCrLf & res, vbInformation, "Mega Pipes Retry Central"
    Exit Sub
ErrorHandler:
    SafeMsgBox "Retry central export encountered an error: " & Err.Description, vbCritical
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

Public Sub HideOutputSheets()
    On Error Resume Next
    Dim wsOff As Worksheet, wsSyn As Worksheet
    Set wsOff = ThisWorkbook.Worksheets("OFFER")
    Set wsSyn = ThisWorkbook.Worksheets("SYNOPSIS")
    
    ' Ensure primary console is active so Excel never tries to hide all visible sheets
    Dim wsActive As Worksheet
    Set wsActive = ActiveSheet
    If Not wsActive Is Nothing Then
        If wsActive.Name = "OFFER" Or wsActive.Name = "SYNOPSIS" Then
            ThisWorkbook.Worksheets("QUOTE").Activate
        End If
    End If
    
    If Not wsOff Is Nothing Then wsOff.Visible = xlSheetHidden
    If Not wsSyn Is Nothing Then wsSyn.Visible = xlSheetHidden
    On Error GoTo 0
End Sub

Public Sub ShowOutputSheets()
    On Error Resume Next
    Dim wsOff As Worksheet, wsSyn As Worksheet
    Set wsOff = ThisWorkbook.Worksheets("OFFER")
    If Not wsOff Is Nothing Then wsOff.Visible = xlSheetVisible
    Set wsSyn = ThisWorkbook.Worksheets("SYNOPSIS")
    If Not wsSyn Is Nothing Then wsSyn.Visible = xlSheetVisible
    On Error GoTo 0
End Sub

Public Sub InvalidateOutputSheets()
    On Error Resume Next
    Dim wsOff As Worksheet, wsSyn As Worksheet
    Set wsOff = ThisWorkbook.Worksheets("OFFER")
    Set wsSyn = ThisWorkbook.Worksheets("SYNOPSIS")
    
    HideOutputSheets
    
    If Not wsOff Is Nothing Then
        modProtection.UnprotectSheet wsOff
        wsOff.Cells.Clear
    End If
    
    If Not wsSyn Is Nothing Then
        modProtection.UnprotectSheet wsSyn
        wsSyn.Cells.Clear
    End If
    On Error GoTo 0
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
    With ws.Range("A34")
        .Value = msg
        .Font.Color = colr
    End With
    With ws.Range("D32")
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

Public Sub UI_PreviewCompare()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    ThisWorkbook.Worksheets("COMPARE").Visible = xlSheetVisible
    modCompare.BuildCompareSheet
    ThisWorkbook.Worksheets("COMPARE").Activate
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    SafeMsgBox "Could not open COMPARE sheet: " & Err.Description, vbExclamation
End Sub

Public Sub UI_RefreshDistances()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    modCompare.BuildCompareSheet
    Application.ScreenUpdating = True
    SafeMsgBox "Competitor distances and plant comparisons refreshed.", vbInformation, "Mega Pipes — Compare Dashboard"
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    SafeMsgBox "Could not refresh distances: " & Err.Description, vbExclamation
End Sub

' ============================================================
' DISTANCE NOTE PANEL (shown on QUOTE sheet after calculation)
' Writes to D15:E17 - nearest MPPL plant & nearest competitor
' ============================================================
' ============================================================
' DISTANCE NOTE PANEL — shown on QUOTE sheet after calculation
' Writes to D15:E17.
'
' Rules enforced:
'   - Delivery PIN extracted via modCompare.ExtractPIN (first
'     consecutive 6-digit run).  No hardcoded destination fallback.
'   - All cache matches require Status = "VALID" (case-insensitive)
'     and exact CustomerPIN equality (not substring containment).
'   - No P01/P02 hardcoded fallback distances.
'   - No DefaultKM column reference.
'   - Partial MPPL comparisons are labelled as partial.
'   - Source and CalculatedDate from the cache row are displayed.
'   - City is appended only when not already present in PlantName.
' ============================================================

Private Sub ShowDistanceNoteOnQuoteSheet()
    Dim ws         As Worksheet
    Dim wsPlants   As Worksheet
    Dim wsDist     As Worksheet
    Dim tblP       As ListObject
    Dim tblDist    As ListObject
    Dim rowObj     As ListRow
    Dim dRow       As ListRow

    Dim deliveryPlace As String
    Dim deliveryPIN   As String
    Dim productId     As String

    Dim eligible       As Variant
    Dim i              As Long
    Dim pId            As String
    Dim dPlant         As String
    Dim dPin           As String
    Dim dStatus        As String

    ' MPPL nearest-plant tracking
    Dim bestPlantId    As String
    Dim bestPlantName  As String
    Dim bestKm         As Double
    Dim bestSrc        As String
    Dim bestDate       As String
    Dim validPlantCount As Long
    Dim totalEligible   As Long
    Dim rawDate         As Variant

    ' Competitor tracking
    Dim nearestCompName As String
    Dim nearestCompKm   As Double
    Dim nearestCompSrc  As String
    Dim nearestCompDate As String
    Dim anyCompCached   As Boolean

    ' Line builders
    Dim mpplLine As String
    Dim compLine As String
    Dim panelText As String
    Dim srcLabel  As String

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets("QUOTE")
    modProtection.UnprotectSheet ws

    ' ── 1. Read delivery address ──────────────────────────────
    deliveryPlace = Trim$(CStr(ws.Range("B10").Value))
    If Len(deliveryPlace) = 0 Then _
        deliveryPlace = Trim$(CStr(ws.Range("B11").Value))

    ' ── 2. Extract PIN via modCompare.ExtractPIN ──────────────
    '    (first consecutive 6-digit run; resets on non-digit)
    deliveryPIN = ""
    If Len(deliveryPlace) > 0 Then
        On Error Resume Next
        deliveryPIN = modCompare.ExtractPIN(deliveryPlace)
        On Error GoTo ErrHandler
    End If

    If Len(deliveryPIN) = 0 Then
        ' No PIN found — show message, never display Bagodar distances
        WriteDistancePanelCells ws, _
            "Delivery PIN required — distances unavailable.", _
            "Enter a delivery address containing a 6-digit PIN in B10 (or B11).", _
            True
        Exit Sub
    End If

    ' ── 3. Product type ───────────────────────────────────────
    productId = Trim$(CStr(ws.Range("B7").Value))
    If Len(productId) = 0 Then productId = "P02"

    ' ── 4. Open shared data tables ───────────────────────────
    Set wsPlants = ThisWorkbook.Worksheets("PLANTS")
    Set tblP = wsPlants.ListObjects("tblPlants")

    On Error Resume Next
    Set wsDist = ThisWorkbook.Worksheets("DISTANCE_CACHE")
    If Not wsDist Is Nothing Then
        Set tblDist = wsDist.ListObjects("tblDistanceCache")
    End If
    On Error GoTo ErrHandler

    ' ── 5. Eligible MPPL plants ──────────────────────────────
    On Error Resume Next
    eligible = modPlants.GetEligiblePlants(productId)
    On Error GoTo ErrHandler

    If Not IsArray(eligible) Then
        ReDim eligible(0 To 1)
        eligible(0) = "P01" : eligible(1) = "P02"
    End If
    If UBound(eligible) < LBound(eligible) Then
        ReDim eligible(0 To 1)
        eligible(0) = "P01" : eligible(1) = "P02"
    End If
    totalEligible = UBound(eligible) - LBound(eligible) + 1

    bestKm = 1E+15
    validPlantCount = 0

    ' ── 6. Find nearest MPPL plant with VALID exact-match cache ──
    For i = LBound(eligible) To UBound(eligible)
        pId = CStr(eligible(i))

        If Not tblDist Is Nothing Then
            For Each dRow In tblDist.ListRows
                dPlant = Trim$(CStr(dRow.Range.Cells(1, tblDist.ListColumns("PlantID").Index).Value))
                dPin   = Trim$(CStr(dRow.Range.Cells(1, tblDist.ListColumns("CustomerPIN").Index).Value))
                dStatus = UCase$(Trim$(CStr(dRow.Range.Cells(1, tblDist.ListColumns("Status").Index).Value)))

                If StrComp(dPlant, pId, vbTextCompare) = 0 And _
                   StrComp(dPin, deliveryPIN, vbTextCompare) = 0 And _
                   dStatus = "VALID" Then

                    Dim foundKm As Double
                    foundKm = Val(dRow.Range.Cells(1, tblDist.ListColumns("DistanceKM").Index).Value)

                    validPlantCount = validPlantCount + 1

                    If foundKm < bestKm Then
                        bestKm = foundKm
                        bestPlantId = pId
                        bestSrc = Trim$(CStr(dRow.Range.Cells(1, tblDist.ListColumns("Source").Index).Value))
                        rawDate = dRow.Range.Cells(1, tblDist.ListColumns("CalculatedDate").Index).Value
                        bestDate = IIf(IsDate(rawDate), Format$(CDate(rawDate), "dd/mm/yy"), "")
                    End If
                    Exit For
                End If
            Next dRow
        End If
    Next i

    ' ── 7. Build MPPL display line ────────────────────────────
    If validPlantCount = 0 Then
        mpplLine = "MPPL road distance unavailable for PIN " & deliveryPIN & " — no valid cache record."
    Else
        bestPlantName = GetCleanPlantName(tblP, bestPlantId)
        srcLabel = ""
        If Len(bestSrc) > 0 Then srcLabel = bestSrc
        If Len(bestDate) > 0 Then
            srcLabel = srcLabel & IIf(Len(srcLabel) > 0, ", ", "") & bestDate
        End If

        mpplLine = bestPlantName & ": " & Format$(bestKm, "#,##0") & " KM"
        If Len(srcLabel) > 0 Then mpplLine = mpplLine & " [" & srcLabel & "]"

        If validPlantCount < totalEligible Then
            mpplLine = mpplLine & _
                " (partial — " & CStr(totalEligible - validPlantCount) & _
                " eligible plant(s) have no valid cache record for PIN " & deliveryPIN & ")"
        End If
    End If

    ' ── 8. Find nearest competitor with VALID exact-match cache ──
    nearestCompKm = 1E+15
    anyCompCached = False

    For Each rowObj In tblP.ListRows
        Dim pType   As String
        Dim pActive As String
        pType   = UCase$(Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("PlantType").Index).Value)))
        pActive = UCase$(Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("Active").Index).Value)))

        If pType = "COMPETITOR" And (pActive = "YES" Or Len(pActive) = 0) Then
            pId = Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("PlantID").Index).Value))

            If Not tblDist Is Nothing Then
                For Each dRow In tblDist.ListRows
                    dPlant  = Trim$(CStr(dRow.Range.Cells(1, tblDist.ListColumns("PlantID").Index).Value))
                    dPin    = Trim$(CStr(dRow.Range.Cells(1, tblDist.ListColumns("CustomerPIN").Index).Value))
                    dStatus = UCase$(Trim$(CStr(dRow.Range.Cells(1, tblDist.ListColumns("Status").Index).Value)))

                    If StrComp(dPlant, pId, vbTextCompare) = 0 And _
                       StrComp(dPin, deliveryPIN, vbTextCompare) = 0 And _
                       dStatus = "VALID" Then

                        Dim cKm As Double
                        cKm = Val(dRow.Range.Cells(1, tblDist.ListColumns("DistanceKM").Index).Value)
                        anyCompCached = True

                        If cKm < nearestCompKm Then
                            nearestCompKm   = cKm
                            nearestCompName = GetCleanPlantName(tblP, pId)
                            nearestCompSrc  = Trim$(CStr(dRow.Range.Cells(1, tblDist.ListColumns("Source").Index).Value))
                            rawDate = dRow.Range.Cells(1, tblDist.ListColumns("CalculatedDate").Index).Value
                            nearestCompDate = IIf(IsDate(rawDate), Format$(CDate(rawDate), "dd/mm/yy"), "")
                        End If
                        Exit For
                    End If
                Next dRow
            End If
        End If
    Next rowObj

    ' ── 9. Build competitor display line ──────────────────────
    If Not anyCompCached Then
        compLine = "Competitor distances not cached for this destination PIN (" & deliveryPIN & ")."
    Else
        srcLabel = ""
        If Len(nearestCompSrc) > 0 Then srcLabel = nearestCompSrc
        If Len(nearestCompDate) > 0 Then
            srcLabel = srcLabel & IIf(Len(srcLabel) > 0, ", ", "") & nearestCompDate
        End If

        compLine = nearestCompName & ": " & Format$(nearestCompKm, "#,##0") & " KM"
        If Len(srcLabel) > 0 Then compLine = compLine & " [" & srcLabel & "]"
    End If

    ' ── 10. Compose and write panel ───────────────────────────
    panelText = Chr(10) & "Nearest MPPL Plant  (PIN " & deliveryPIN & "):" & Chr(10) & _
                mpplLine & Chr(10) & Chr(10) & _
                "Nearest Competitor  (PIN " & deliveryPIN & "):" & Chr(10) & _
                compLine

    WriteDistancePanelCells ws, panelText, "", False
    Exit Sub

ErrHandler:
    On Error Resume Next
    WriteDistancePanelError ws
End Sub

' ── Private helper: extract clean plant name, city appended only ──
' if not already present in PlantName.
Private Function GetCleanPlantName(ByRef tblP As ListObject, ByVal plantId As String) As String
    Dim rowObj As ListRow
    Dim pName  As String
    Dim pCity  As String

    For Each rowObj In tblP.ListRows
        If StrComp(Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("PlantID").Index).Value)), _
                   plantId, vbTextCompare) = 0 Then
            pName = Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("PlantName").Index).Value))
            pCity = ""
            On Error Resume Next
            pCity = Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("City").Index).Value))
            On Error GoTo 0
            ' Append city only when it is not already embedded in the name
            If Len(pCity) > 0 And InStr(1, pName, pCity, vbTextCompare) = 0 Then
                pName = pName & " (" & pCity & ")"
            End If
            GetCleanPlantName = pName
            Exit Function
        End If
    Next rowObj

    GetCleanPlantName = plantId   ' Fall back to ID string if plant not found
End Function

' ── Private helper: write the panel content to D15:E17 ──────────
Private Sub WriteDistancePanelCells(ByRef ws As Worksheet, _
                                    ByVal mainText As String, _
                                    ByVal subText As String, _
                                    ByVal isWarning As Boolean)
    Dim panelRange As Range
    Set panelRange = ws.Range("D15:E17")

    If panelRange.MergeCells Then panelRange.UnMerge
    panelRange.ClearContents
    panelRange.Merge

    Dim displayText As String
    If Len(subText) > 0 Then
        displayText = mainText & Chr(10) & subText
    Else
        displayText = mainText
    End If

    With ws.Range("D15")
        .Value = displayText
        .Font.Name = "Calibri"
        .Font.Size = 8.5
        .Font.Bold = False
        If isWarning Then
            .Font.Color = RGB(120, 60, 0)
            .Interior.Color = RGB(255, 243, 205)    ' Amber — warning
        Else
            .Font.Color = RGB(16, 44, 96)            ' Navy
            .Interior.Color = RGB(219, 234, 254)     ' Light blue — info
        End If
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .IndentLevel = 1
    End With

    Dim leftColor As Long
    leftColor = IIf(isWarning, RGB(176, 108, 0), RGB(37, 99, 235))

    With panelRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = leftColor
    End With
    With panelRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = IIf(isWarning, RGB(253, 200, 100), RGB(147, 197, 253))
    End With
    With panelRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = IIf(isWarning, RGB(253, 200, 100), RGB(147, 197, 253))
    End With
    With panelRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = IIf(isWarning, RGB(253, 200, 100), RGB(147, 197, 253))
    End With

    panelRange.Locked = False
End Sub

' ── Private helper: error fallback — writes neutral placeholder ──
Private Sub WriteDistancePanelError(ByRef ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Dim panelRange As Range
    Set panelRange = ws.Range("D15:E17")
    If panelRange.MergeCells Then panelRange.UnMerge
    panelRange.Merge
    With ws.Range("D15")
        .Value = "DISTANCE INFO: (error — run Calculate again to refresh)"
        .Font.Color = RGB(110, 110, 110)
        .Interior.Color = RGB(248, 250, 248)
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 8.5
        .IndentLevel = 0
    End With
    panelRange.Borders(xlEdgeLeft).LineStyle = xlNone
End Sub

