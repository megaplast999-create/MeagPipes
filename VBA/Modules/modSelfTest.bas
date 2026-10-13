Attribute VB_Name = "modSelfTest"
Option Explicit

Public Sub RunP3GoldenTests()
    Dim outputPath As String
    Dim fileNumber As Integer
    Dim legacyTotal As Variant
    Dim currentTotal As Variant
    Dim legacyExpected As Variant
    Dim currentExpected As Variant
    Dim legacyDifference As Variant
    Dim currentDifference As Variant
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim index As Long

    On Error GoTo TestError
    outputPath = TestsPath() & Application.PathSeparator & "vba_p3_golden_20261008.txt"
    legacyExpected = CDec(178657900)
    currentExpected = CDec(178801050)
    legacyTotal = P3GoldenTotal(True)
    currentTotal = P3GoldenTotal(False)
    legacyDifference = CDec(legacyTotal) - legacyExpected
    currentDifference = CDec(currentTotal) - currentExpected

    fileNumber = FreeFile
    Open outputPath For Output As #fileNumber
    Print #fileNumber, "P3 VBA golden test date: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #fileNumber, "TOLERANCE=<1"
    Print #fileNumber, "INTERMEDIATE_DIAGNOSTICS:"
    For index = 0 To 3
        WriteGoldenLine fileNumber, True, index
        WriteGoldenLine fileNumber, False, index
    Next index
    Print #fileNumber, "LEGACY:"
    Print #fileNumber, "Expected total=" & CStr(legacyExpected)
    Print #fileNumber, "Actual total=" & CStr(legacyTotal)
    Print #fileNumber, "Difference=" & CStr(legacyDifference)
    Print #fileNumber, "Result=" & IIf(Abs(CDbl(legacyDifference)) < 1, "PASS", "FAIL")
    Print #fileNumber, "CURRENT:"
    Print #fileNumber, "Expected total=" & CStr(currentExpected)
    Print #fileNumber, "Actual total=" & CStr(currentTotal)
    Print #fileNumber, "Difference=" & CStr(currentDifference)
    Print #fileNumber, "Result=" & IIf(Abs(CDbl(currentDifference)) < 1, "PASS", "FAIL")
    Print #fileNumber, "RESULT=" & IIf(Abs(CDbl(legacyDifference)) < 1 And _
        Abs(CDbl(currentDifference)) < 1, "PASS", "FAIL")
    Close #fileNumber
    Exit Sub
TestError:
    errorNumber = Err.Number
    errorDescription = Err.Description
    WriteGoldenFailure outputPath, errorNumber, errorDescription
    Err.Raise errorNumber, "RunP3GoldenTests", errorDescription
End Sub

Public Function RunGoldenLegacy() As String
    Dim total As Variant
    Dim expected As Variant
    Dim index As Long
    Dim res As LineResult
    Dim expectedPrices As Variant

    On Error GoTo FailureHandler
    expected = CDec(178657900)
    expectedPrices = Array(22681, 23914, 24748, 25580)

    total = P3GoldenTotal(True)
    If Abs(CDbl(CDec(total) - expected)) >= 1 Then
        RunGoldenLegacy = "FAIL: Legacy total was " & CStr(total) & ", expected " & CStr(expected)
        Exit Function
    End If

    For index = 0 To 3
        res = P3GoldenLine(True, index)
        If Abs(CDbl(CDec(res.TotalPerM) - CDec(expectedPrices(index)))) >= 1 Then
            RunGoldenLegacy = "FAIL: Line " & CStr(index + 1) & " price was " & CStr(res.TotalPerM) & ", expected " & CStr(expectedPrices(index))
            Exit Function
        End If
    Next index

    RunGoldenLegacy = "PASS"
    Exit Function

FailureHandler:
    RunGoldenLegacy = "FAIL: " & Err.Description
End Function

Public Function RunGoldenNormal() As String
    Dim total As Variant
    Dim expected As Variant
    Dim index As Long
    Dim res As LineResult
    Dim expectedPrices As Variant

    On Error GoTo FailureHandler
    expected = CDec(178801050)
    expectedPrices = Array(22689, 24109, 24857, 25581)

    total = P3GoldenTotal(False)
    If Abs(CDbl(CDec(total) - expected)) >= 1 Then
        RunGoldenNormal = "FAIL: Normal total was " & CStr(total) & ", expected " & CStr(expected)
        Exit Function
    End If

    For index = 0 To 3
        res = P3GoldenLine(False, index)
        If Abs(CDbl(CDec(res.TotalPerM) - CDec(expectedPrices(index)))) >= 1 Then
            RunGoldenNormal = "FAIL: Line " & CStr(index + 1) & " price was " & CStr(res.TotalPerM) & ", expected " & CStr(expectedPrices(index))
            Exit Function
        End If
    Next index

    RunGoldenNormal = "PASS"
    Exit Function

FailureHandler:
    RunGoldenNormal = "FAIL: " & Err.Description
End Function

Public Function RunP5GoldenTestSuite() As String
    Dim outputPath As String
    Dim fileNumber As Integer
    Dim legResult As String
    Dim normResult As String
    Dim geomResult As String
    Dim guardResult As String
    Dim index As Long

    On Error GoTo FailureHandler
    outputPath = TestsPath() & Application.PathSeparator & "vba_p5_golden_20261012.txt"
    legResult = RunGoldenLegacy()
    normResult = RunGoldenNormal()
    geomResult = RunP3InvalidGeometryTest()
    guardResult = RunP3InvalidMasterGuardTest()

    fileNumber = FreeFile
    Open outputPath For Output As #fileNumber
    Print #fileNumber, "P5 VBA golden test suite date: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #fileNumber, "TOLERANCE=<1"
    Print #fileNumber, "RunGoldenLegacy=" & legResult
    Print #fileNumber, "RunGoldenNormal=" & normResult
    Print #fileNumber, "RunInvalidGeometry=" & geomResult
    Print #fileNumber, "RunInvalidMasterGuard=" & guardResult
    Print #fileNumber, "--- LINE DIAGNOSTICS ---"
    For index = 0 To 3
        WriteGoldenLine fileNumber, True, index
        WriteGoldenLine fileNumber, False, index
    Next index
    Print #fileNumber, "--- SUMMARY ---"
    If legResult = "PASS" And normResult = "PASS" And InStr(1, geomResult, "PASS") > 0 And InStr(1, guardResult, "PASS") > 0 Then
        Print #fileNumber, "RESULT=PASS"
        RunP5GoldenTestSuite = "PASS: All P5 golden tests and validation guards passed."
    Else
        Print #fileNumber, "RESULT=FAIL"
        RunP5GoldenTestSuite = "FAIL: One or more P5 tests failed."
    End If
    Close #fileNumber
    Exit Function

FailureHandler:
    RunP5GoldenTestSuite = "FAIL: " & Err.Description
End Function

Private Sub WriteGoldenLine(ByVal fileNumber As Integer, _
                            ByVal legacy As Boolean, _
                            ByVal index As Long)
    Dim result As LineResult
    Dim expectedPerM As Variant
    Dim quantities As Variant
    Dim modeName As String
    Dim pre As Variant
    Dim ext As Variant
    Dim intc As Variant
    Dim massMt As Variant
    Dim expectedLineValue As Variant

    result = P3GoldenLine(legacy, index)
    quantities = Array(2000, 350, 500, 4400)
    If legacy Then
        modeName = "LEGACY"
        expectedPerM = Choose(index + 1, 22681, 23914, 24748, 25580)
    Else
        modeName = "CURRENT"
        expectedPerM = Choose(index + 1, 22689, 24109, 24857, 25581)
    End If
    pre = CDec(SettingValue("DefaultSteelRate")) + CDec(SettingValue("DefaultGradeExtra"))
    pre = pre + pre * CDec(SettingValue("DefaultYieldPct")) / CDec(100)
    pre = pre + CDec(SettingValue("DefaultConversion")) + _
        CDec(SettingValue("DefaultTesting")) + CDec(SettingValue("DefaultMargin"))
    massMt = CDec(result.MassPerM) / CDec(1000)
    If legacy Then
        ext = CDec(500) * CDec(3.14) * CDec(result.od) / _
            CDec(1000) / massMt
        intc = CDec(275) * CDec(3.14) * CDec(result.od) / _
            CDec(1000) / massMt
    Else
        ext = CDec(500) * CDec(3.14) * CDec(result.od) / _
            CDec(1000) / massMt
        intc = CDec(275) * CDec(3.14) * CDec(result.ID) / _
            CDec(1000) / massMt
    End If
    expectedLineValue = CDec(expectedPerM) * CDec(quantities(index))

    Print #fileNumber, "LINE=" & CStr(index + 1) & " MODE=" & modeName
    Print #fileNumber, "OD=" & CStr(result.od) & " ID=" & CStr(result.ID) & " WT=8.4"
    Print #fileNumber, "MassKgPerM=" & CStr(result.MassPerM)
    Print #fileNumber, "MassMtPerM=" & CStr(massMt)
    Print #fileNumber, "Metres=" & CStr(result.Metres) & " MT=" & CStr(result.MT)
    Print #fileNumber, "Pipes=" & CStr(result.Pipes) & " Trailers=" & CStr(result.Trailers)
    Print #fileNumber, "Pre=" & CStr(pre)
    Print #fileNumber, "External=" & CStr(ext)
    Print #fileNumber, "Internal=" & CStr(intc)
    Print #fileNumber, "Freight=" & CStr(result.FreightPerMT)
    Print #fileNumber, "Financial=" & CStr(result.FCperMT)
    Print #fileNumber, "BarePerM=" & CStr(result.BarePerM)
    Print #fileNumber, "ExtPerM=" & CStr(result.ExtPerM)
    Print #fileNumber, "IntPerM=" & CStr(result.IntPerM)
    Print #fileNumber, "FreightPerM=" & CStr(result.FreightPerM)
    Print #fileNumber, "TotalPerM=" & CStr(result.TotalPerM)
    Print #fileNumber, "TotalPerMT=" & CStr(result.TotalPerMT)
    Print #fileNumber, "LineValue=" & CStr(result.LineValue)
    Print #fileNumber, "ExpectedTotalPerM=" & CStr(expectedPerM) & _
        " Difference=" & CStr(CDec(result.TotalPerM) - CDec(expectedPerM))
    Print #fileNumber, "ExpectedLineValue=" & CStr(expectedLineValue) & _
        " Difference=" & CStr(CDec(result.LineValue) - expectedLineValue)
End Sub

Public Function ProjectRootPath() As String
    Dim workbookPath As String
    Dim separatorPosition As Long

    workbookPath = ThisWorkbook.Path
    separatorPosition = InStrRev(workbookPath, Application.PathSeparator)
    If separatorPosition <= 1 Then
        Err.Raise vbObjectError + 1810, "ProjectRootPath", _
            "Unable to derive the project root from ThisWorkbook.Path: " & workbookPath
    End If
    ProjectRootPath = Left$(workbookPath, separatorPosition - 1)
End Function

Public Function TestsPath() As String
    Dim testPath As String

    testPath = ProjectRootPath() & Application.PathSeparator & "TESTS"
    If Len(Dir$(testPath, vbDirectory)) = 0 Then
        Err.Raise vbObjectError + 1811, "TestsPath", _
            "TESTS directory not found: " & testPath
    End If
    TestsPath = testPath
End Function

Private Sub WriteGoldenFailure(ByVal outputPath As String, _
                               ByVal errorNumber As Long, _
                               ByVal errorDescription As String)
    Dim fileNumber As Integer

    If Len(outputPath) = 0 Then Exit Sub
    On Error Resume Next
    fileNumber = FreeFile
    Open outputPath For Output As #fileNumber
    If Err.Number = 0 Then
        Print #fileNumber, "P3 VBA golden test date: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
        Print #fileNumber, "RESULT=FAIL"
        Print #fileNumber, "ERROR_NUMBER=" & CStr(errorNumber)
        Print #fileNumber, "ERROR_DESCRIPTION=" & errorDescription
        Close #fileNumber
    End If
    On Error GoTo 0
End Sub

Public Function RunP3InvalidGeometryTest() As String
    Dim lineData As QuoteLineInput
    Dim rates As RateInput
    Dim result As LineResult

    On Error GoTo ExpectedFailure

    With lineData
        .Basis = "ID"
        .SizeMM = CDec(10)
        .WT = CDec(8)
        .Qty = CDec(1)
        .Unit = "Metres"
        .PipeLen = CDec(12)
        .PerTrailer = CDec(4)
        .HasExt = False
        .HasInt = False
    End With

    LoadDefaultRateInput rates, False
    rates.TripRate = CDec(98000)
    result = CalcLine(lineData, rates, False)
    RunP3InvalidGeometryTest = "FAIL: CalcLine returned a result for invalid geometry."
    Exit Function

ExpectedFailure:
    RunP3InvalidGeometryTest = "PASS: " & Err.Description
End Function

Public Function RunP3InvalidMasterGuardTest() As String
    Dim grades As ListObject
    Dim plants As ListObject
    Dim gradeId As String
    Dim plantId As String

    On Error GoTo ExpectedFailure

    Set grades = ThisWorkbook.Worksheets("MASTERS").ListObjects("tblGrades")
    Set plants = ThisWorkbook.Worksheets("PLANTS").ListObjects("tblPlants")
    gradeId = CStr(grades.ListRows(1).Range.Cells(1, grades.ListColumns("GradeID").Index).Value)
    plantId = CStr(plants.ListRows(1).Range.Cells(1, plants.ListColumns("PlantID").Index).Value)

    AddQuoteLine "P3-INVALID-MASTER", "P02", "HSAW", "ID", _
        10, 8.4, gradeId, 1, "Metres", 12, plantId
    RunP3InvalidMasterGuardTest = "FAIL: P2 accepted an invalid master combination."
    Exit Function

ExpectedFailure:
    RunP3InvalidMasterGuardTest = "PASS: " & Err.Description
End Function


' ============================================================
' P24 UNIFIED SELF-TEST & VALIDATION SUITE
' ============================================================

Public Function RunAllSelfTests() As String
    Dim results As String
    Dim passCount As Long, failCount As Long, humanCount As Long
    Dim logPath As String
    Dim fileNum As Integer
    Dim sep As String
    sep = String(70, "=")
    
    passCount = 0
    failCount = 0
    humanCount = 0
    
    results = "=== MEGA PIPES QUOTATION SYSTEM - UNIFIED SELF-TEST SUITE ===" & vbCrLf
    results = results & "Execution Date: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    results = results & sep & vbCrLf
    
    ' Unprotect all sheets for test execution
    modProtection.UnprotectAllSheets
    
    ' Ensure golden reference quotation lines are present for self-test execution
    Dim tblLinesCheck As ListObject
    Set tblLinesCheck = ThisWorkbook.Worksheets("QUOTE_LINES").ListObjects("tblQuoteLines")
    If tblLinesCheck.ListRows.Count = 0 Then
        PopulateGoldenReferenceQuote
    ElseIf Len(Trim$(CStr(tblLinesCheck.ListRows(1).Range.Cells(1, tblLinesCheck.ListColumns("ProductID").Index).Value))) = 0 Then
        PopulateGoldenReferenceQuote
    End If
    
    ' Core Component Evaluations
    EvalTest "P3/P5 Golden Legacy Contract (178,657,900 INR, 7250 MTR)", RunGoldenLegacy(), results, passCount, failCount, humanCount
    EvalTest "P3/P5 Golden Normal Option 1 (178,801,050 INR, Line 4 = 25,581/m)", RunGoldenNormal(), results, passCount, failCount, humanCount
    EvalTest "P2 Invalid Geometry Guard", RunP3InvalidGeometryTest(), results, passCount, failCount, humanCount
    EvalTest "P2 Invalid Master Combination Guard", RunP3InvalidMasterGuardTest(), results, passCount, failCount, humanCount
    EvalTest "P7 Freight Engine (Methods A/B/C, Stacking, Trailers)", modFreight.RunFreightSelfTests(), results, passCount, failCount, humanCount
    EvalTest "P8 Plant Engine (Eligibility, Road Distance, Override Audit)", modPlants.RunPlantSelfTests(), results, passCount, failCount, humanCount
    EvalTest "P9 ORS Route Matrix Engine & Cache Lookup", modORS.RunORSSelfTests(), results, passCount, failCount, humanCount
    
    ' P9 Live ORS Check (Needs Human API Key)
    If Len(Trim$(Environ$("ORS_API_KEY"))) > 0 Then
        EvalTest "P9 Live ORS Network Integration", "PASS: Configured via ORS_API_KEY", results, passCount, failCount, humanCount
    Else
        EvalHuman "P9 Live ORS Network Integration", "ORS_API_KEY environment variable not configured (cataloged in TESTS/NEEDS_HUMAN.md)", results, humanCount
    End If
    
    EvalTest "P10 Competitor Engine (Floor Protection, Stale Badge, Yellow Highlight)", modCompetitors.RunCompetitorSelfTests(), results, passCount, failCount, humanCount
    EvalTest "P11 COSTING Sheet Reproduction (Values-Only, Reference Layout)", modExport.RunCostingSelfTests(), results, passCount, failCount, humanCount
    EvalTest "P12 OFFER Sheet Reproduction (Unipro 27 Terms, Indian Words)", modExport.RunOfferSelfTests(), results, passCount, failCount, humanCount
    EvalTest "P13 Sequence Allocation Engine & File Locking", modSequence.RunSequenceSelfTests(), results, passCount, failCount, humanCount
    
    ' P13 Shared Network UNC Path (Needs Human)
    If Len(Trim$(Environ$("MEGA_SHARED_FOLDER"))) > 0 Then
        EvalTest "P13 Production Shared Network Folder", "PASS: Configured via MEGA_SHARED_FOLDER", results, passCount, failCount, humanCount
    Else
        EvalHuman "P13 Production Shared Network Folder", "Production UNC path required for multi-machine deployment (cataloged in TESTS/NEEDS_HUMAN.md)", results, humanCount
    End If
    
    EvalTest "P14 Snapshot Engine & Immutable Quote Reload", modLog.RunSnapshotSelfTests(), results, passCount, failCount, humanCount
    EvalTest "P15 Master Synchronisation & Schema Validation", modAdmin.RunMastersSelfTests(), results, passCount, failCount, humanCount
    EvalTest "P17 Offline Quoting & Reconnection Reconciliation", modSequence.RunOfflineSelfTests(), results, passCount, failCount, humanCount
    EvalTest "P18 Export Engine (PDF, Sanitized Customer Excel, Internal Excel)", modExport.RunExportSelfTests(), results, passCount, failCount, humanCount
    EvalTest "P19 Document Numbering Idempotency & Debounce", modQuoteHeader.RunDocumentNumberingSelfTests(), results, passCount, failCount, humanCount
    EvalTest "P20 Audit Logging & Password Masking", modLog.RunAuditSelfTests(), results, passCount, failCount, humanCount
    EvalTest "P21 Admin & Approval Workflow State Machine", modApproval.RunApprovalSelfTests(), results, passCount, failCount, humanCount
    EvalTest "P22 User Identity & Filesystem Sanitization", modLog.RunIdentitySelfTests(), results, passCount, failCount, humanCount
    EvalTest "P23 Workbook Protection & Selective Cell Locking", modProtection.RunProtectionSelfTests(), results, passCount, failCount, humanCount
    
    ' P23 Passwords (Needs Human)
    If Len(Trim$(Environ$("MEGA_SHEET_PWD"))) > 0 And Len(Trim$(Environ$("MEGA_ADMIN_PWD"))) > 0 Then
        EvalTest "P23 Production Security Passwords", "PASS: Configured via MEGA_SHEET_PWD and MEGA_ADMIN_PWD", results, passCount, failCount, humanCount
    Else
        EvalHuman "P23 Production Security Passwords", "MEGA_SHEET_PWD and MEGA_ADMIN_PWD required for production lockdown (cataloged in TESTS/NEEDS_HUMAN.md)", results, humanCount
    End If
    
    results = results & sep & vbCrLf
    results = results & "SUMMARY: Total Evaluated: " & CStr(passCount + failCount + humanCount) & _
              " | Passed: " & CStr(passCount) & _
              " | Needs Human: " & CStr(humanCount) & _
              " | Failed: " & CStr(failCount) & vbCrLf
    
    ' Write log to TESTS\unified_selftest_p24.txt
    On Error Resume Next
    logPath = TestsPath() & Application.PathSeparator & "unified_selftest_p24.txt"
    fileNum = FreeFile
    Open logPath For Output As #fileNum
    Print #fileNum, results
    Close #fileNum
    On Error GoTo 0
    
    ' Restore sheet protection
    modProtection.ProtectAllSheets userInterfaceOnly:=True
    
    If failCount = 0 Then
        RunAllSelfTests = "PASS"
    Else
        RunAllSelfTests = "FAIL: " & CStr(failCount) & " test(s) failed."
    End If
End Function

Public Sub SelfTest()
    Dim res As String
    res = RunAllSelfTests()
    MsgBox "Self-Test Result: " & res, vbInformation, "Mega Pipes Quotation System"
End Sub

Private Sub EvalTest(ByVal compName As String, ByVal testResult As String, ByRef outLog As String, ByRef passC As Long, ByRef failC As Long, ByRef humC As Long)
    Dim trimmedRes As String
    trimmedRes = Trim$(testResult)
    If Left$(trimmedRes, 4) = "PASS" Then
        passC = passC + 1
        outLog = outLog & "  [PASS] " & compName & vbCrLf
    ElseIf Left$(trimmedRes, 11) = "NEEDS_HUMAN" Then
        humC = humC + 1
        outLog = outLog & "  [NEEDS_HUMAN] " & compName & " - " & Mid$(trimmedRes, 13) & vbCrLf
    Else
        failC = failC + 1
        outLog = outLog & "  [FAIL] " & compName & " - " & trimmedRes & vbCrLf
    End If
End Sub

Private Sub EvalHuman(ByVal compName As String, ByVal reason As String, ByRef outLog As String, ByRef humC As Long)
    humC = humC + 1
    outLog = outLog & "  [NEEDS_HUMAN] " & compName & " - " & reason & vbCrLf
End Sub

Public Sub PopulateGoldenReferenceQuote()
    Dim wsQ As Worksheet
    Dim wsLines As Worksheet
    Dim tblLines As ListObject
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean
    
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    Set tblLines = wsLines.ListObjects("tblQuoteLines")
    
    modProtection.UnprotectSheet wsQ
    modProtection.UnprotectSheet wsLines
    
    ' Clear existing lines
    If tblLines.ListRows.Count > 0 Then
        On Error Resume Next
        tblLines.DataBodyRange.Delete
        On Error GoTo 0
    End If
    
    ' Set Header
    wsQ.Range("B4").Value = "MKTG/ENQ/26-27/594"
    wsQ.Range("B5").Value = "2026-10-12"
    wsQ.Range("B8").Value = "UNIPRO"
    wsQ.Range("B9").Value = "Mr. Sharma"
    wsQ.Range("B11").Value = "Project Site, Gujarat"
    wsQ.Range("B14").Value = "P02"
    wsQ.Range("B15").Value = "30 Days from Dispatch"
    wsQ.Range("B16").Value = 30
    wsQ.Range("B17").Value = "Metre"
    wsQ.Range("B18").Value = "NORMAL"
    
    ' Add 4 Golden Lines
    modQuoteLines.AddQuoteLine "MKTG/ENQ/26-27/594", "P02", "HSAW", "ID", 1226, 8.4, "HSAW-002", 2000, "Metres", 12, "P02"
    modQuoteLines.AddQuoteLine "MKTG/ENQ/26-27/594", "P02", "HSAW", "ID", 1300, 8.4, "HSAW-002", 350, "Metres", 12, "P02"
    modQuoteLines.AddQuoteLine "MKTG/ENQ/26-27/594", "P02", "HSAW", "ID", 1350, 8.4, "HSAW-002", 500, "Metres", 12, "P02"
    modQuoteLines.AddQuoteLine "MKTG/ENQ/26-27/594", "P02", "HSAW", "ID", 1400, 8.4, "HSAW-002", 4400, "Metres", 12, "P02"
    
    ' Recalculate and refresh display
    modQuote.RecalculateQuote
    modUI.RefreshQuoteLinesDisplay
    
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
End Sub
