Attribute VB_Name = "modUI"
Option Explicit

' ============================================================
' Mega Pipes Quotation System
' Module: modUI
' Purpose:
'   Simplified Guided User Interface Layer
'   Allows commercial estimators to manage complete quotations
'   from a clean START dashboard and a guided QUOTE console
'   without touching backend calculation sheets.
'
'   SAFETY BOUNDARY:
'   No pricing logic is modified here.
'   All calculations delegate to modPricing, modQuote, modQuoteLines.
'   All numbering delegates to modSequence, modQuoteHeader.
'   All snapshots delegate to modLog.
'   All approvals delegate to modApproval.
'   All exports delegate to modExport.
' ============================================================

' UI Color Palette (Mega Pipes brand)
Public Const UI_GREEN_DARK  As Long = 2169120   ' RGB(32, 96, 33)   - Primary brand green
Public Const UI_GREEN_MID   As Long = 4166204   ' RGB(60, 140, 63)  - Mid accent green
Public Const UI_GREEN_LIGHT As Long = 13166093  ' RGB(205, 230, 200) - Soft green
Public Const UI_BLUE        As Long = 11829830  ' RGB(70, 130, 180) - Action blue
Public Const UI_ORANGE      As Long = 27760     ' RGB(176, 108, 0)  - Calc orange
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
    
    ' Configure cell locking and protect sheets for macro access
    modProtection.ConfigureCellLocking
    modProtection.ProtectAllSheets userInterfaceOnly:=True
    
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
    Dim saveErrNum As Long, saveErrDesc As String, saveErrSource As String
    saveErrNum = Err.Number
    saveErrDesc = Err.Description
    saveErrSource = Err.Source
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    Dim logPath As String, fNum As Integer
    logPath = ThisWorkbook.Path & "\..\TESTS\ui_error.log"
    fNum = FreeFile
    Open logPath For Append As #fNum
    Print #fNum, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | OnWorkbookOpen | Err: " & saveErrNum & " | Desc: " & saveErrDesc & " | Source: " & saveErrSource
    Close #fNum
    On Error GoTo 0
    SafeMsgBox "Mega Pipes Quotation System initialization error: " & saveErrDesc, vbExclamation
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
    
    ' Remove old shapes / buttons to cleanly refresh
    Dim sIdx As Long
    For sIdx = ws.Shapes.Count To 1 Step -1
        On Error Resume Next
        ws.Shapes(sIdx).Delete
        On Error GoTo 0
    Next sIdx
    
    ws.Range("A1:G35").Interior.Color = RGB(248, 250, 248)
    ws.Range("A1:G35").Font.Name = "Calibri"
    
    ' Set column widths
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 20
    ws.Columns("E").ColumnWidth = 20
    ws.Columns("F").ColumnWidth = 20
    ws.Columns("G").ColumnWidth = 3
    
    ' Row heights
    ws.Rows("1:2").RowHeight = 10
    ws.Rows("3:6").RowHeight = 22
    ws.Rows("7").RowHeight = 16
    ws.Rows("8").RowHeight = 10
    ws.Rows("9:13").RowHeight = 16
    ws.Rows("14").RowHeight = 12
    ws.Rows("15").RowHeight = 18
    ws.Rows("16:26").RowHeight = 24
    ws.Rows("27:30").RowHeight = 14
    
    ' ── Brand Header Banner (B3:F6) ──
    With ws.Range("B3:F6")
        .Merge
        .Value = "MEGA PIPES QUOTATION SYSTEM"
        .Interior.Color = UI_GREEN_DARK
        .Font.Color = UI_WHITE
        .Font.Bold = True
        .Font.Size = 22
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Subtitle (B7:F7)
    With ws.Range("B7:F7")
        .Merge
        .Value = "Commercial Estimation & Quotation Application  |  Desktop Edition"
        .Interior.Color = UI_GREEN_MID
        .Font.Color = UI_WHITE
        .Font.Size = 10
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' ── Active Quotation KPI Card (B9:F13) ──
    Dim qNum As String, qCust As String, qTotal As String, qState As String, qMode As String
    Dim wsQ As Worksheet
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    qNum = Trim$(CStr(wsQ.Range("B4").Value))
    qCust = Trim$(CStr(wsQ.Range("B8").Value))
    qState = modApproval.GetQuoteState()
    qMode = Trim$(CStr(wsQ.Range("E12").Value))
    If Len(qMode) = 0 Then qMode = "NORMAL (Option 1)"
    
    If Len(qNum) > 0 And qNum <> "0" Then
        Dim totVal As Double
        totVal = Val(wsQ.Range("E10").Value)
        qTotal = "₹" & Format$(totVal, "#,##0")
    Else
        qNum = "(None - Click NEW QUOTATION)"
        qCust = "-"
        qTotal = "-"
        qState = "Draft"
    End If
    
    With ws.Range("B9:F9")
        .Merge
        .Value = "ACTIVE QUOTATION STATUS"
        .Interior.Color = RGB(230, 240, 230)
        .Font.Bold = True
        .Font.Size = 9
        .Font.Color = UI_GREEN_DARK
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Range("B10").Value = "Enquiry Reference:"
    ws.Range("C10").Value = qNum
    ws.Range("C10").Font.Bold = True
    ws.Range("D10").Value = "Customer:"
    ws.Range("E10:F10").Merge
    ws.Range("E10").Value = IIf(Len(qCust) > 0, qCust, "-")
    
    ws.Range("B11").Value = "Workflow State:"
    ws.Range("C11").Value = qState
    ws.Range("C11").Font.Bold = True
    ws.Range("D11").Value = "Grand Total (Incl GST):"
    ws.Range("E11:F11").Merge
    ws.Range("E11").Value = qTotal
    ws.Range("E11").Font.Bold = True
    
    ws.Range("B12").Value = "Pricing Mode:"
    ws.Range("C12").Value = qMode
    ws.Range("D12").Value = "User Login:"
    ws.Range("E12:F12").Merge
    ws.Range("E12").Value = modLog.GetCurrentUser()
    
    With ws.Range("B9:F12").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 205, 180)
    End With
    
    ' Section Title (B14:F14)
    With ws.Range("B14:F14")
        .Merge
        .Value = "QUOTATION ACTIONS"
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = UI_GREEN_DARK
        .HorizontalAlignment = xlCenter
    End With
    
    ' ── Add Buttons on START ──
    AddModernButton ws, "btnStartNew", "NEW QUOTATION", "modUI.UI_NewQuotation", _
                    ws.Range("B16"), ws.Range("F16"), UI_GREEN_MID, UI_WHITE, True
                    
    AddModernButton ws, "btnStartOpen", "OPEN EXISTING QUOTATION", "modUI.UI_OpenQuotation", _
                    ws.Range("B18"), ws.Range("F18"), UI_BLUE, UI_WHITE, False
                    
    AddModernButton ws, "btnStartPreview", "PREVIEW OFFER LETTER", "modUI.UI_PreviewOffer", _
                    ws.Range("B20"), ws.Range("F20"), RGB(100, 110, 120), UI_WHITE, False
                    
    AddModernButton ws, "btnStartPDF", "EXPORT OFFER (PDF)", "modUI.UI_ExportPDF", _
                    ws.Range("B22"), ws.Range("C22"), RGB(160, 60, 40), UI_WHITE, False
                    
    AddModernButton ws, "btnStartExcel", "EXPORT CUSTOMER EXCEL", "modUI.UI_ExportCustomerExcel", _
                    ws.Range("D22"), ws.Range("F22"), RGB(0, 110, 160), UI_WHITE, False
                    
    AddModernButton ws, "btnStartAdmin", "ADMIN / BACK OFFICE ACCESS ▸", "modUI.UI_AdminUnlock", _
                    ws.Range("B24"), ws.Range("F24"), UI_GRAY_DARK, UI_WHITE, False
                    
    ' ── System Notification Area (B26:F28) ──
    With ws.Range("B26:F28")
        .Merge
        .Value = "System Status: Online | Master data verified | Quotation engine ready." & vbCrLf & _
                 "To create a quotation, click 'NEW QUOTATION' above. All calculations and costing run automatically."
        .Interior.Color = RGB(240, 245, 240)
        .Font.Size = 8
        .Font.Color = RGB(80, 100, 80)
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
        .Value = "Mega Pipes Quotation System  |  Confidential & Proprietary"
        .Font.Size = 7.5
        .Font.Color = RGB(150, 150, 150)
        .HorizontalAlignment = xlCenter
    End With
    
    modProtection.ProtectSheet ws, userInterfaceOnly:=True
    
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
    
    On Error GoTo ErrorHandler
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Set ws = ThisWorkbook.Worksheets("QUOTE")
    modProtection.UnprotectSheet ws
    
    ' Clear old buttons on QUOTE
    Dim sIdx As Long
    For sIdx = ws.Shapes.Count To 1 Step -1
        On Error Resume Next
        ws.Shapes(sIdx).Delete
        On Error GoTo 0
    Next sIdx
    
    ' Ensure master-driven dropdowns and validations are active
    modMasterData.SetupMasterDrivenHeader
    modDependentDropdowns.SetupDependentDropdownArchitecture
    
    ' ── Action Buttons on QUOTE Console (placed in columns D-E rows 15-18) ──
    AddModernButton ws, "btnQuoteCalc", "CALCULATE QUOTATION", "modUI.UI_CalculateQuotation", _
                    ws.Range("D15"), ws.Range("E15"), UI_GREEN_MID, UI_WHITE, True
                    
    AddModernButton ws, "btnQuoteSave", "SAVE DRAFT", "modUI.UI_SaveDraft", _
                    ws.Range("D16"), ws.Range("D16"), UI_BLUE, UI_WHITE, False
                    
    AddModernButton ws, "btnQuotePreview", "PREVIEW OFFER", "modUI.UI_PreviewOffer", _
                    ws.Range("E16"), ws.Range("E16"), RGB(100, 110, 120), UI_WHITE, False
                    
    AddModernButton ws, "btnQuotePDF", "EXPORT PDF", "modUI.UI_ExportPDF", _
                    ws.Range("D17"), ws.Range("D17"), RGB(160, 60, 40), UI_WHITE, False
                    
    AddModernButton ws, "btnQuoteHome", "← BACK TO HOME", "modUI.UI_BackToHome", _
                    ws.Range("E17"), ws.Range("E17"), UI_GRAY_DARK, UI_WHITE, False
                    
    ' ── Item Entry Action Buttons (cols D-E next to rows 22-29) ──
    AddModernButton ws, "btnQuoteAdd", "+ ADD ITEM TO QUOTE", "modUI.UI_AddItem", _
                    ws.Range("D22"), ws.Range("E23"), UI_GREEN_DARK, UI_WHITE, True
                    
    AddModernButton ws, "btnQuoteRemove", "- REMOVE LINE", "modUI.UI_RemoveItem", _
                    ws.Range("D25"), ws.Range("E25"), RGB(180, 70, 70), UI_WHITE, False
                    
    AddModernButton ws, "btnQuoteClear", "CLEAR ALL LINES", "modUI.UI_ClearAllItems", _
                    ws.Range("D27"), ws.Range("E27"), RGB(120, 120, 120), UI_WHITE, False
                    
    ' ── Validation Feedback Banner at Row 33 ──
    With ws.Range("A33:E33")
        .Merge
        .Value = "Ready. Select specifications above and click '+ ADD ITEM TO QUOTE'."
        .Interior.Color = RGB(245, 248, 245)
        .Font.Size = 9
        .Font.Color = UI_GREEN_DARK
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("A33:E33").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 210, 180)
    End With
    
    ' ── Section 3: Quotation Line Items Table Header ──
    With ws.Range("A35:I35")
        .Merge
        .Value = "3. CURRENT QUOTATION LINE ITEMS"
        .Interior.Color = UI_GREEN_DARK
        .Font.Color = UI_WHITE
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    ' Table column headers (Row 36)
    ws.Range("A36").Value = "Line ID"
    ws.Range("B36").Value = "Pipe Specification"
    ws.Range("C36").Value = "Coating (Ext / Int)"
    ws.Range("D36").Value = "Quantity"
    ws.Range("E36").Value = "Pipes"
    ws.Range("F36").Value = "Weight (MT)"
    ws.Range("G36").Value = "Unit Price"
    ws.Range("H36").Value = "Line Total"
    ws.Range("I36").Value = "Status"
    
    With ws.Range("A36:I36")
        .Interior.Color = UI_GREEN_MID
        .Font.Color = UI_WHITE
        .Font.Bold = True
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
    End With
    
    ' Populate active lines display
    RefreshQuoteLinesDisplay
    
    ' Re-protect sheet with UserInterfaceOnly
    modProtection.ConfigureCellLocking
    modProtection.ProtectSheet ws, userInterfaceOnly:=True
    
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
    
    ' Clear previous display rows (37 to 55)
    wsQ.Range("A37:I55").UnMerge
    wsQ.Range("A37:I55").ClearContents
    wsQ.Range("A37:I55").Interior.Color = xlNone
    wsQ.Range("A37:I55").Borders.LineStyle = xlNone
    
    outRow = 37
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
                
                wsQ.Cells(outRow, 1).Value = rObj.Range.Cells(1, tblLines.ListColumns("LineID").Index).Value
                wsQ.Cells(outRow, 2).Value = fam & " " & sz & "mm (" & bas & ") x " & wt & "mm " & grd
                wsQ.Cells(outRow, 3).Value = extC & " / " & intC
                wsQ.Cells(outRow, 4).Value = Format$(qty, "#,##0") & " " & unt
                wsQ.Cells(outRow, 5).Value = pcs
                wsQ.Cells(outRow, 6).Value = Format$(mt, "#,##0.00")
                wsQ.Cells(outRow, 7).Value = "₹" & Format$(prc, "#,##0")
                wsQ.Cells(outRow, 8).Value = "₹" & Format$(tot, "#,##0")
                wsQ.Cells(outRow, 9).Value = st
                
                ' Formatting
                With wsQ.Range(wsQ.Cells(outRow, 1), wsQ.Cells(outRow, 9))
                    .Font.Size = 9
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                    .Interior.Color = IIf(outRow Mod 2 = 0, RGB(245, 250, 245), RGB(255, 255, 255))
                End With
                wsQ.Cells(outRow, 1).HorizontalAlignment = xlCenter
                wsQ.Cells(outRow, 5).HorizontalAlignment = xlRight
                wsQ.Cells(outRow, 6).HorizontalAlignment = xlRight
                wsQ.Cells(outRow, 7).HorizontalAlignment = xlRight
                wsQ.Cells(outRow, 8).HorizontalAlignment = xlRight
                wsQ.Cells(outRow, 9).HorizontalAlignment = xlCenter
                
                outRow = outRow + 1
            End If
        Next rObj
    End If
    
    If outRow = 37 Then
        wsQ.Range("A37:I37").Merge
        wsQ.Range("A37").Value = "(No quotation lines added yet. Use Section 2 above to select product and click '+ ADD ITEM TO QUOTE')"
        wsQ.Range("A37").Font.Italic = True
        wsQ.Range("A37").Font.Size = 9
        wsQ.Range("A37").Font.Color = RGB(120, 120, 120)
        wsQ.Range("A37").HorizontalAlignment = xlCenter
    End If
    
    modProtection.ProtectSheet wsQ, userInterfaceOnly:=True
    Exit Sub
ErrorHandler:
    On Error Resume Next
    modProtection.ProtectSheet wsQ, userInterfaceOnly:=True
End Sub

' ============================================================
' 5. USER WORKFLOW ACTION PROCEDURES
' ============================================================

Public Sub UI_NewQuotation()
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean
    Dim curStep As String
    curStep = "Start"
    
    On Error GoTo ErrorHandler
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    curStep = "Unprotect wsQ"
    Dim wsQ As Worksheet
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    modProtection.UnprotectSheet wsQ
    
    curStep = "InitializeNewQuote"
    ' Allocate or confirm fresh enquiry number
    modQuoteHeader.InitializeNewQuote forceNew:=True
    
    curStep = "Clear input cells"
    ' Clear input cells in product requirement area
    wsQ.Range("B21").Value = "Pipe Standard"
    wsQ.Range("B22").Value = "HSAW"
    wsQ.Range("B23").Value = "OD"
    wsQ.Range("B24").Value = 457.2
    wsQ.Range("B25").Value = 7.92
    wsQ.Range("B26").Value = "Fe 410"
    wsQ.Range("B27").Value = "3LPE Standard"
    wsQ.Range("B28").Value = "None"
    wsQ.Range("B29").Value = 1000
    wsQ.Range("B30").Value = "Metres"
    wsQ.Range("B31").Value = 12#
    
    curStep = "ClearQuoteLines"
    ' Clear quote lines table
    modQuote.ClearQuoteLines
    
    curStep = "SetupDependentDropdownArchitecture"
    ' Refresh dependent dropdown architecture
    modDependentDropdowns.SetupDependentDropdownArchitecture
    
    curStep = "SetupQuoteLineDropdowns"
    modDependentDropdowns.SetupQuoteLineDropdowns
    
    curStep = "RefreshQuoteLinesDisplay"
    ' Refresh display table
    RefreshQuoteLinesDisplay
    
    curStep = "SetFeedbackBanner"
    ' Update status banner
    SetFeedbackBanner "New quotation created. Select customer and click '+ ADD ITEM TO QUOTE'.", UI_GREEN_DARK
    
    curStep = "ProtectSheet wsQ"
    modProtection.ProtectSheet wsQ, userInterfaceOnly:=True
    
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = True
    
    curStep = "Activate wsQ"
    wsQ.Activate
    On Error Resume Next
    wsQ.Range("B8").Select
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    Dim saveErrNum As Long, saveErrDesc As String, saveErrSource As String
    saveErrNum = Err.Number
    saveErrDesc = Err.Description
    saveErrSource = Err.Source
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    On Error Resume Next
    Dim logPath As String, fNum As Integer
    logPath = ThisWorkbook.Path & "\..\TESTS\ui_error.log"
    fNum = FreeFile
    Open logPath For Append As #fNum
    Print #fNum, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | Step: " & curStep & " | Err: " & saveErrNum & " | Desc: " & saveErrDesc & " | Source: " & saveErrSource
    Close #fNum
    On Error GoTo 0
    SafeMsgBox "Could not start new quotation [" & curStep & "]: " & saveErrDesc, vbExclamation
End Sub

Public Sub UI_OpenQuotation()
    On Error GoTo ErrorHandler
    Dim enquiryNo As String
    If Application.DisplayAlerts Then
        enquiryNo = InputBox("Enter Enquiry Reference number to open:" & vbCrLf & _
                             "(e.g. MKTGENQ26-27-001 or MKTG/ENQ/26-27/TEST-SNAP)", _
                             "Mega Pipes — Open Quotation", "")
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
    
    ' Read inputs
    fam = Trim$(CStr(wsQ.Range("B22").Value))
    bas = UCase$(Trim$(CStr(wsQ.Range("B23").Value)))
    sz = Val(wsQ.Range("B24").Value)
    wt = Val(wsQ.Range("B25").Value)
    grd = Trim$(CStr(wsQ.Range("B26").Value))
    extC = Trim$(CStr(wsQ.Range("B27").Value))
    intC = Trim$(CStr(wsQ.Range("B28").Value))
    qty = Val(wsQ.Range("B29").Value)
    unt = Trim$(CStr(wsQ.Range("B30").Value))
    If Len(unt) = 0 Then unt = "Metres"
    
    ' Plant resolution
    plt = Trim$(CStr(wsQ.Range("B14").Value))
    If InStr(1, plt, "Khopoli", vbTextCompare) > 0 Then
        plantId = "P02"
    Else
        plantId = "P01"
    End If
    
    ' Validate mandatory inputs
    If Len(fam) = 0 Then
        SetFeedbackBanner "Validation Error: Please select Pipe Family (HSAW / HFW).", UI_RED
        Exit Sub
    End If
    If bas <> "OD" And bas <> "ID" Then
        SetFeedbackBanner "Validation Error: Basis must be OD or ID.", UI_RED
        Exit Sub
    End If
    If sz <= 0 Then
        SetFeedbackBanner "Validation Error: Please select a valid Pipe Size (mm).", UI_RED
        Exit Sub
    End If
    If wt <= 0 Then
        SetFeedbackBanner "Validation Error: Please select Wall Thickness (mm).", UI_RED
        Exit Sub
    End If
    If Len(grd) = 0 Then
        SetFeedbackBanner "Validation Error: Please select Steel Grade.", UI_RED
        Exit Sub
    End If
    If qty <= 0 Then
        SetFeedbackBanner "Validation Error: Quantity must be greater than zero.", UI_RED
        Exit Sub
    End If
    
    ' Map Product ID
    Dim hasCoating As Boolean
    hasCoating = (UCase$(extC) <> "BARE" And Len(extC) > 0 And UCase$(extC) <> "NONE")
    If fam = "HSAW" Then
        prodId = IIf(hasCoating, "P02", "P01")
    Else
        prodId = IIf(hasCoating, "P04", "P03")
    End If
    
    ' Map Grade ID from tblGrades
    Dim gradeId As String
    Dim tblGrades As ListObject
    Dim rGrd As ListRow
    Set tblGrades = ThisWorkbook.Worksheets("MASTERS").ListObjects("tblGrades")
    gradeId = ""
    For Each rGrd In tblGrades.ListRows
        Dim gFam As String, gId As String, gName As String
        gFam = CStr(rGrd.Range.Cells(1, tblGrades.ListColumns("Family").Index).Value)
        gId = CStr(rGrd.Range.Cells(1, tblGrades.ListColumns("GradeID").Index).Value)
        gName = CStr(rGrd.Range.Cells(1, tblGrades.ListColumns("GradeName").Index).Value)
        
        If StrComp(gFam, fam, vbTextCompare) = 0 Then
            If StrComp(gId, grd, vbTextCompare) = 0 Or StrComp(gName, grd, vbTextCompare) = 0 Then
                gradeId = gId
                Exit For
            ElseIf InStr(1, gName, grd, vbTextCompare) > 0 Then
                gradeId = gId
                Exit For
            End If
        End If
    Next rGrd
    
    If Len(gradeId) = 0 Then
        If fam = "HSAW" Then gradeId = "HSAW-002" Else gradeId = "HFW-002"
    End If
    
    ' Add Line
    Dim prevEvents As Boolean, prevScreen As Boolean
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim legacyMode As Boolean
    legacyMode = (InStr(1, CStr(wsQ.Range("E12").Value), "LEGACY", vbTextCompare) > 0)
    
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
    wsQ.Range("B29").Value = 0
    
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
    
    ' Update COSTING and OFFER sheets
    On Error Resume Next
    modExport.WriteCostingSheet
    modExport.WriteOfferSheet
    On Error GoTo ErrorHandler
    
    ' Refresh display table
    RefreshQuoteLinesDisplay
    
    Dim grandTotal As String
    grandTotal = "₹" & Format$(Val(wsQ.Range("E10").Value), "#,##0")
    SetFeedbackBanner "Quotation calculated successfully! Grand Total: " & grandTotal & ". Ready for review or export.", UI_GREEN_DARK
    
    Application.ScreenUpdating = True
    SafeMsgBox "Quotation calculated successfully!" & vbCrLf & vbCrLf & _
               "Customer: " & cust & vbCrLf & _
               "Subtotal (Excl GST): ₹" & Format$(Val(wsQ.Range("E8").Value), "#,##0") & vbCrLf & _
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
    
    ' Ensure return button exists on OFFER
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
               "(COSTING, STACKING, COMPARE, QUOTE_LINES, MASTERS, SETTINGS, etc.)", vbInformation
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
    With ws.Range("A33")
        .Value = msg
        .Font.Color = colr
    End With
    modProtection.ProtectSheet ws, userInterfaceOnly:=True
End Sub

' ============================================================
' 8. MODERN BUTTON BUILDER
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
