Option Explicit

Private mIsLogging As Boolean

' ============================================================
' P22 USER IDENTITY ENGINE
' ============================================================

Public Function GetCurrentUser() As String
    Dim u As String
    u = Trim$(Environ$("USERNAME"))
    If Len(u) = 0 Then
        On Error Resume Next
        u = Trim$(Application.UserName)
        On Error GoTo 0
    End If
    If Len(u) = 0 Then u = "User"
    
    ' Sanitize for safe filesystem and log usage
    u = Replace(u, "\", "_")
    u = Replace(u, "/", "_")
    u = Replace(u, ":", "_")
    u = Replace(u, "*", "_")
    u = Replace(u, "?", "_")
    u = Replace(u, """", "_")
    u = Replace(u, "<", "_")
    u = Replace(u, ">", "_")
    u = Replace(u, "|", "_")
    u = Replace(u, " ", "_")
    
    GetCurrentUser = u
End Function

' ============================================================
' Mega Pipes Quotation System
' Module: modLog
' Purpose:
'   P14 Snapshot and Quote History
'   GetEnquiryNumber, SaveQuote, LoadQuote, GetRevision
'   Immutable reload: reloading an old quote restores exact numbers
' ============================================================

Public Function GetEnquiryNumber(Optional ByVal sharedPath As String = "") As String
    GetEnquiryNumber = AllocateEnquiryNumber(sharedPath)
End Function

Public Function GetQuotesDir(Optional ByVal sharedFolder As String = "") As String
    Dim basePath As String
    Dim fy As String
    Dim qDir As String
    Dim fso As Object
    
    fy = GetFiscalYear()
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Len(Trim$(sharedFolder)) > 0 Then
        basePath = Trim$(sharedFolder)
    Else
        On Error Resume Next
        Dim tblSet As ListObject
        Dim sRow As ListRow
        Set tblSet = ThisWorkbook.Worksheets("SETTINGS").ListObjects("tblSettings")
        If Not tblSet Is Nothing Then
            Set sRow = FindSettingRow(tblSet, "SharedFolder")
            If Not sRow Is Nothing Then
                basePath = Trim$(CStr(sRow.Range.Cells(1, tblSet.ListColumns("SettingValue").index).value))
            End If
        End If
        On Error GoTo 0
        
        If Len(basePath) = 0 Then
            basePath = ThisWorkbook.Path
        End If
    End If
    
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
    qDir = basePath & "Quotes\"
    If Not fso.FolderExists(qDir) Then
        On Error Resume Next
        fso.CreateFolder qDir
        On Error GoTo 0
    End If
    
    qDir = qDir & fy & "\"
    If Not fso.FolderExists(qDir) Then
        On Error Resume Next
        fso.CreateFolder qDir
        On Error GoTo 0
    End If
    
    GetQuotesDir = qDir
End Function

Public Function GetRevision(ByVal enquiryNo As String, Optional ByVal sharedFolder As String = "") As Long
    Dim cleanNo As String
    Dim qDir As String
    Dim fName As String
    Dim maxRev As Long
    Dim p1 As Long, p2 As Long
    Dim revStr As String
    Dim revVal As Long
    
    cleanNo = Replace(Replace(enquiryNo, "/", "_"), "\", "_")
    qDir = GetQuotesDir(sharedFolder)
    maxRev = -1
    
    fName = Dir(qDir & cleanNo & "_R*.xlsx")
    Do While Len(fName) > 0
        p1 = InStr(1, fName, "_R", vbTextCompare)
        If p1 > 0 Then
            p2 = InStr(p1 + 2, fName, "_")
            If p2 > 0 Then
                revStr = Mid$(fName, p1 + 2, p2 - (p1 + 2))
                If IsNumeric(revStr) Then
                    revVal = CLng(Val(revStr))
                    If revVal > maxRev Then maxRev = revVal
                End If
            End If
        End If
        fName = Dir()
    Loop
    
    If maxRev < 0 Then
        GetRevision = 0
    Else
        GetRevision = maxRev + 1
    End If
End Function

Public Function SaveQuote(ByVal enquiryNo As String, Optional ByVal revNum As Long = -1, Optional ByVal sharedFolder As String = "") As String
    Dim cleanNo As String
    Dim curUser As String
    Dim fy As String
    Dim qDir As String
    Dim fullPath As String
    Dim fName As String
    Dim snapWb As Workbook
    Dim wsMeta As Worksheet
    Dim wsLinesSnap As Worksheet
    Dim wsQ As Worksheet
    Dim wsLines As Worksheet
    Dim tblLines As ListObject
    Dim header As QuoteHeaderData
    Dim summary As QuoteSummaryData
    Dim numLines As Long
    Dim numCols As Long
    Dim r As Long, c As Long
    Dim wsLog As Worksheet
    Dim tblLog As ListObject
    Dim newRow As ListRow
    
    fy = GetFiscalYear()
    If revNum < 0 Then
        revNum = GetRevision(enquiryNo, sharedFolder)
    End If
    
    curUser = GetCurrentUser()
    
    cleanNo = Replace(Replace(enquiryNo, "/", "_"), "\", "_")
    fName = cleanNo & "_R" & CStr(revNum) & "_" & curUser & ".xlsx"
    qDir = GetQuotesDir(sharedFolder)
    fullPath = qDir & fName
    
    ' Read current quote state
    header = ReadQuoteSheet()
    summary = CalcAllLines(header)
    
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    Set tblLines = wsLines.ListObjects("tblQuoteLines")
    
    ' Create snapshot workbook
    Set snapWb = Application.Workbooks.Add(xlWBATWorksheet)
    Set wsMeta = snapWb.Worksheets(1)
    wsMeta.Name = "SNAPSHOT_META"
    
    ' Write metadata
    wsMeta.Range("A1").value = "SettingKey"
    wsMeta.Range("B1").value = "SettingValue"
    
    wsMeta.Range("A2").value = "QuoteID": wsMeta.Range("B2").value = enquiryNo
    wsMeta.Range("A3").value = "Revision": wsMeta.Range("B3").value = revNum
    wsMeta.Range("A4").value = "QuoteDate": wsMeta.Range("B4").value = header.QuoteDate
    wsMeta.Range("A5").value = "Customer": wsMeta.Range("B5").value = header.Customer
    wsMeta.Range("A6").value = "ContactPerson": wsMeta.Range("B6").value = header.ContactPerson
    wsMeta.Range("A7").value = "DeliveryPlace": wsMeta.Range("B7").value = header.deliveryPlace
    wsMeta.Range("A8").value = "PlantID": wsMeta.Range("B8").value = header.plantId
    wsMeta.Range("A9").value = "PaymentTerms": wsMeta.Range("B9").value = header.PaymentTerms
    wsMeta.Range("A10").value = "ValidityDays": wsMeta.Range("B10").value = header.ValidityDays
    wsMeta.Range("A11").value = "PriceBasis": wsMeta.Range("B11").value = header.priceBasis
    wsMeta.Range("A12").value = "LegacyMode": wsMeta.Range("B12").value = IIf(header.legacyMode, "YES", "NO")
    wsMeta.Range("A13").value = "TotalMetres": wsMeta.Range("B13").value = summary.TotalMetres
    wsMeta.Range("A14").value = "TotalWeightMT": wsMeta.Range("B14").value = summary.totalWeightMT
    wsMeta.Range("A15").value = "TotalPipeCount": wsMeta.Range("B15").value = summary.TotalPipeCount
    wsMeta.Range("A16").value = "TotalTrailers": wsMeta.Range("B16").value = summary.totalTrailers
    wsMeta.Range("A17").value = "TotalExclGST": wsMeta.Range("B17").value = summary.TotalExclGST
    wsMeta.Range("A18").value = "GSTAmount": wsMeta.Range("B18").value = summary.GSTAmount
    wsMeta.Range("A19").value = "GrandTotalInclGST": wsMeta.Range("B19").value = summary.GrandTotalInclGST
    wsMeta.Range("A20").value = "SavedBy": wsMeta.Range("B20").value = curUser
    wsMeta.Range("A21").value = "SavedTimestamp": wsMeta.Range("B21").value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    ' Copy quote lines into SNAPSHOT_LINES
    Set wsLinesSnap = snapWb.Worksheets.Add(After:=wsMeta)
    wsLinesSnap.Name = "SNAPSHOT_LINES"
    
    numLines = tblLines.ListRows.count
    numCols = tblLines.ListColumns.count
    
    ' Copy headers
    For c = 1 To numCols
        wsLinesSnap.Cells(1, c).value = tblLines.HeaderRowRange.Cells(1, c).value
    Next c
    
    ' Copy line values
    If numLines > 0 Then
        For r = 1 To numLines
            For c = 1 To numCols
                wsLinesSnap.Cells(r + 1, c).value = tblLines.DataBodyRange.Cells(r, c).value
            Next c
        Next r
    End If
    
    ' Copy OFFER sheet values if present
    On Error Resume Next
    Dim wsOff As Worksheet
    Dim wsOffSnap As Worksheet
    Set wsOff = ThisWorkbook.Worksheets("OFFER")
    If Not wsOff Is Nothing Then
        Set wsOffSnap = snapWb.Worksheets.Add(After:=wsLinesSnap)
        wsOffSnap.Name = "OFFER"
        wsOff.UsedRange.Copy
        wsOffSnap.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False
    End If
    On Error GoTo 0
    
    ' Save and close snapshot workbook
    Application.DisplayAlerts = False
    snapWb.SaveAs FileName:=fullPath, FileFormat:=xlOpenXMLWorkbook
    snapWb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    
    ' Log snapshot save to tblQuoteLog
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets("QUOTE_LOG")
    modProtection.UnprotectSheet wsLog
    Set tblLog = wsLog.ListObjects("tblQuoteLog")
    On Error GoTo 0
    
    If Not tblLog Is Nothing Then
        If tblLog.ListRows.count = 1 And Len(Trim$(CStr(tblLog.ListRows(1).Range.Cells(1, tblLog.ListColumns("Action").index).value))) = 0 Then
            Set newRow = tblLog.ListRows(1)
        Else
            Set newRow = tblLog.ListRows.Add
        End If
        
        With newRow.Range
            .Cells(1, tblLog.ListColumns("LogID").index).value = "LOG-" & Format$(newRow.index, "00000")
            .Cells(1, tblLog.ListColumns("QuoteID").index).value = enquiryNo
            .Cells(1, tblLog.ListColumns("Timestamp").index).value = Now
            .Cells(1, tblLog.ListColumns("User").index).value = curUser
            .Cells(1, tblLog.ListColumns("Action").index).value = "SNAPSHOT_SAVE"
            .Cells(1, tblLog.ListColumns("OldValue").index).value = ""
            .Cells(1, tblLog.ListColumns("NewValue").index).value = Format$(summary.TotalExclGST, "0")
            .Cells(1, tblLog.ListColumns("SheetName").index).value = "QUOTE"
            .Cells(1, tblLog.ListColumns("Status").index).value = "SAVED"
            .Cells(1, tblLog.ListColumns("Notes").index).value = "Snapshot saved: R" & CStr(revNum) & " (" & fName & ")"
        End With
        modProtection.ProtectSheet wsLog, userInterfaceOnly:=True
    End If
    
    SaveQuote = fullPath
End Function

Public Function LoadQuote(ByVal snapshotPathOrEnquiryNo As String, Optional ByVal sharedFolder As String = "") As Boolean
    Dim fullPath As String
    Dim snapWb As Workbook
    Dim wsMeta As Worksheet
    Dim wsLinesSnap As Worksheet
    Dim wsQ As Worksheet
    Dim wsLines As Worksheet
    Dim tblLines As ListObject
    Dim fso As Object
    Dim qDir As String
    Dim cleanNo As String
    Dim fName As String
    Dim r As Long, c As Long
    Dim numLines As Long, numCols As Long
    Dim metaDict As Object
    Dim k As String, v As Variant
    Dim curUser As String
    Dim wsLog As Worksheet
    Dim tblLog As ListObject
    Dim newRow As ListRow
    Dim restoredTotal As Double
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fullPath = snapshotPathOrEnquiryNo
    
    If Not fso.FileExists(fullPath) Then
        qDir = GetQuotesDir(sharedFolder)
        cleanNo = Replace(Replace(snapshotPathOrEnquiryNo, "/", "_"), "\", "_")
        fName = Dir(qDir & cleanNo & "_R*.xlsx")
        If Len(fName) > 0 Then
            fullPath = qDir & fName
        Else
            LoadQuote = False
            Exit Function
        End If
    End If
    
    ' Open snapshot workbook read-only
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set snapWb = Application.Workbooks.Open(FileName:=fullPath, ReadOnly:=True, UpdateLinks:=False)
    
    Set wsMeta = snapWb.Worksheets("SNAPSHOT_META")
    Set wsLinesSnap = snapWb.Worksheets("SNAPSHOT_LINES")
    
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    modProtection.UnprotectSheet wsQ
    modProtection.UnprotectSheet wsLines
    Set tblLines = wsLines.ListObjects("tblQuoteLines")
    
    ' Read metadata into dictionary
    Set metaDict = CreateObject("Scripting.Dictionary")
    For r = 2 To wsMeta.Cells(wsMeta.Rows.count, 1).End(xlUp).Row
        k = Trim$(CStr(wsMeta.Cells(r, 1).value))
        v = wsMeta.Cells(r, 2).value
        If Len(k) > 0 Then metaDict(k) = v
    Next r
    
    ' Restore header fields on QUOTE sheet
    If metaDict.Exists("QuoteID") Then wsQ.Range("B4").value = metaDict("QuoteID")
    If metaDict.Exists("QuoteDate") Then wsQ.Range("B5").value = metaDict("QuoteDate")
    If metaDict.Exists("Customer") Then wsQ.Range("B8").value = metaDict("Customer")
    If metaDict.Exists("ContactPerson") Then wsQ.Range("B9").value = metaDict("ContactPerson")
    If metaDict.Exists("DeliveryPlace") Then wsQ.Range("B11").value = metaDict("DeliveryPlace")
    If metaDict.Exists("PlantID") Then wsQ.Range("B14").value = metaDict("PlantID")
    If metaDict.Exists("PaymentTerms") Then wsQ.Range("B15").value = metaDict("PaymentTerms")
    If metaDict.Exists("ValidityDays") Then wsQ.Range("B16").value = metaDict("ValidityDays")
    If metaDict.Exists("PriceBasis") Then wsQ.Range("B17").value = metaDict("PriceBasis")
    
    ' Restore live summary card values on QUOTE sheet
    If metaDict.Exists("TotalMetres") Then wsQ.Range("E4").value = metaDict("TotalMetres")
    If metaDict.Exists("TotalWeightMT") Then wsQ.Range("E5").value = metaDict("TotalWeightMT")
    If metaDict.Exists("TotalPipeCount") Then wsQ.Range("E6").value = metaDict("TotalPipeCount")
    If metaDict.Exists("TotalTrailers") Then wsQ.Range("E7").value = metaDict("TotalTrailers")
    If metaDict.Exists("TotalExclGST") Then
        wsQ.Range("E8").value = metaDict("TotalExclGST")
        restoredTotal = CDbl(metaDict("TotalExclGST"))
    End If
    If metaDict.Exists("GSTAmount") Then wsQ.Range("E9").value = metaDict("GSTAmount")
    If metaDict.Exists("GrandTotalInclGST") Then wsQ.Range("E10").value = metaDict("GrandTotalInclGST")
    wsQ.Range("E11").value = "LOADED_SNAPSHOT"
    
    ' Restore lines into tblQuoteLines
    If tblLines.ListRows.count > 0 Then
        tblLines.DataBodyRange.Delete
    End If
    
    numLines = wsLinesSnap.Cells(wsLinesSnap.Rows.count, 1).End(xlUp).Row - 1
    numCols = tblLines.ListColumns.count
    
    If numLines > 0 Then
        For r = 1 To numLines
            Dim newBodyRow As ListRow
            Set newBodyRow = tblLines.ListRows.Add
            For c = 1 To numCols
                newBodyRow.Range.Cells(1, c).value = wsLinesSnap.Cells(r + 1, c).value
            Next c
        Next r
    End If
    
    ' Close snapshot workbook
    snapWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Log snapshot load to tblQuoteLog
    curUser = GetCurrentUser()
    
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets("QUOTE_LOG")
    modProtection.UnprotectSheet wsLog
    Set tblLog = wsLog.ListObjects("tblQuoteLog")
    On Error GoTo 0
    
    If Not tblLog Is Nothing Then
        If tblLog.ListRows.count = 1 And Len(Trim$(CStr(tblLog.ListRows(1).Range.Cells(1, tblLog.ListColumns("Action").index).value))) = 0 Then
            Set newRow = tblLog.ListRows(1)
        Else
            Set newRow = tblLog.ListRows.Add
        End If
        
        With newRow.Range
            .Cells(1, tblLog.ListColumns("LogID").index).value = "LOG-" & Format$(newRow.index, "00000")
            .Cells(1, tblLog.ListColumns("QuoteID").index).value = CStr(wsQ.Range("B4").value)
            .Cells(1, tblLog.ListColumns("Timestamp").index).value = Now
            .Cells(1, tblLog.ListColumns("User").index).value = curUser
            .Cells(1, tblLog.ListColumns("Action").index).value = "SNAPSHOT_LOAD"
            .Cells(1, tblLog.ListColumns("OldValue").index).value = ""
            .Cells(1, tblLog.ListColumns("NewValue").index).value = Format$(restoredTotal, "0")
            .Cells(1, tblLog.ListColumns("SheetName").index).value = "QUOTE"
            .Cells(1, tblLog.ListColumns("Status").index).value = "LOADED"
            .Cells(1, tblLog.ListColumns("Notes").index).value = "Snapshot reloaded from " & fullPath
        End With
        modProtection.ProtectSheet wsLog, userInterfaceOnly:=True
    End If
    
    modProtection.ProtectSheet wsLines, userInterfaceOnly:=True
    modProtection.ProtectSheet wsQ, userInterfaceOnly:=True
    LoadQuote = True
End Function

' ============================================================
' P14 SELF-TESTS
' ============================================================

Public Function RunSnapshotSelfTests() As String
    Dim snapPath As String
    Dim snapPathR1 As String
    Dim qId As String
    Dim wsQ As Worksheet
    Dim wsLines As Worksheet
    Dim tblLines As ListObject
    Dim origTotal As Double
    Dim mutatedTotal As Double
    Dim restoredTotal As Double
    Dim fso As Object
    Dim ok As Boolean
    Dim revNext As Long
    Dim p1 As Double, p2 As Double, p3 As Double, p4 As Double
    
    On Error GoTo TestFail
    
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    Set tblLines = wsLines.ListObjects("tblQuoteLines")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    qId = "MKTG/ENQ/26-27/TEST-SNAP"
    
    ' Ensure lines exist for snapshot testing
    If tblLines.ListRows.count = 0 Then
        modSelfTest.PopulateGoldenReferenceQuote
    ElseIf Len(Trim$(CStr(tblLines.ListRows(1).Range.Cells(1, tblLines.ListColumns("ProductID").index).value))) = 0 Then
        modSelfTest.PopulateGoldenReferenceQuote
    End If
    
    ' Clean any prior test snapshots
    On Error Resume Next
    Dim oldFile As String
    oldFile = Dir(GetQuotesDir() & "*TEST-SNAP*.xlsx")
    Do While Len(oldFile) > 0
        fso.DeleteFile GetQuotesDir() & oldFile, True
        oldFile = Dir()
    Loop
    On Error GoTo 0
    
    ' Step 1: Record original verified grand total
    origTotal = CDbl(wsQ.Range("E8").value)
    If origTotal = 0 Then origTotal = 178801050#
    
    ' Step 2: Save snapshot R0
    snapPath = SaveQuote(qId, 0)
    If Not fso.FileExists(snapPath) Then
        RunSnapshotSelfTests = "FAIL: Snapshot file not found at " & snapPath
        Exit Function
    End If
    
    ' Step 3: Mutate active quote (change quantity of Line 1 to 9999)
    tblLines.ListRows(1).Range.Cells(1, tblLines.ListColumns("Quantity").index).value = 9999
    ' Recalculate quote
    Dim dummyHead As QuoteHeaderData
    Dim dummySum As QuoteSummaryData
    dummyHead = ReadQuoteSheet()
    dummySum = CalcAllLines(dummyHead)
    UpdateLiveSummary dummySum, dummyHead
    
    mutatedTotal = CDbl(wsQ.Range("E8").value)
    If mutatedTotal = origTotal Then
        RunSnapshotSelfTests = "FAIL: Mutated total should differ from original"
        Exit Function
    End If
    
    ' Step 4: Reload snapshot
    ok = LoadQuote(snapPath)
    If Not ok Then
        RunSnapshotSelfTests = "FAIL: LoadQuote returned False"
        Exit Function
    End If
    
    ' Step 5: Verify exact numbers restored (0 difference)
    restoredTotal = CDbl(wsQ.Range("E8").value)
    If Abs(restoredTotal - origTotal) > 0.01 Then
        RunSnapshotSelfTests = "FAIL: Restored total mismatch. Expected " & CStr(origTotal) & ", got " & CStr(restoredTotal)
        Exit Function
    End If
    
    ' Verify 4 lines restored
    If tblLines.ListRows.count <> 4 Then
        RunSnapshotSelfTests = "FAIL: Expected 4 restored lines, got " & CStr(tblLines.ListRows.count)
        Exit Function
    End If
    
    p1 = CDbl(tblLines.ListRows(1).Range.Cells(1, tblLines.ListColumns("UnitPrice").index).value)
    p2 = CDbl(tblLines.ListRows(2).Range.Cells(1, tblLines.ListColumns("UnitPrice").index).value)
    p3 = CDbl(tblLines.ListRows(3).Range.Cells(1, tblLines.ListColumns("UnitPrice").index).value)
    p4 = CDbl(tblLines.ListRows(4).Range.Cells(1, tblLines.ListColumns("UnitPrice").index).value)
    
    If p1 <> 22689# Or p2 <> 24109# Or p3 <> 24857# Or p4 <> 25581# Then
        RunSnapshotSelfTests = "FAIL: Restored line prices mismatch (22689, 24109, 24857, 25581)"
        Exit Function
    End If
    
    ' Step 6: Test revision tracking
    revNext = GetRevision(qId)
    If revNext <> 1 Then
        RunSnapshotSelfTests = "FAIL: Expected next revision 1, got " & CStr(revNext)
        Exit Function
    End If
    
    snapPathR1 = SaveQuote(qId, revNext)
    If Not fso.FileExists(snapPathR1) Then
        RunSnapshotSelfTests = "FAIL: R1 snapshot not found"
        Exit Function
    End If
    
    revNext = GetRevision(qId)
    If revNext <> 2 Then
        RunSnapshotSelfTests = "FAIL: Expected next revision 2, got " & CStr(revNext)
        Exit Function
    End If
    
    ' Clean up test files
    On Error Resume Next
    If fso.FileExists(snapPath) Then fso.DeleteFile snapPath, True
    If fso.FileExists(snapPathR1) Then fso.DeleteFile snapPathR1, True
    On Error GoTo 0
    
    RunSnapshotSelfTests = "PASS"
    Exit Function
    
TestFail:
    RunSnapshotSelfTests = "FAIL: Exception - " & Err.Description
End Function

' ============================================================
' P20 AUDIT LOGGING ENGINE
' ============================================================

Public Function LogAction(ByVal quoteId As String, ByVal actionName As String, _
                          Optional ByVal oldVal As String = "", _
                          Optional ByVal newVal As String = "", _
                          Optional ByVal sheetName As String = "QUOTE", _
                          Optional ByVal statusVal As String = "LOGGED", _
                          Optional ByVal notesText As String = "") As String
    Dim wsLog As Worksheet
    Dim tblLog As ListObject
    Dim newRow As ListRow
    Dim curUser As String
    Dim logId As String
    
    ' Recursion protection
    If mIsLogging Then Exit Function
    mIsLogging = True
    
    On Error GoTo CleanExit
    
    ' Mask secrets/passwords
    Dim safeOld As String, safeNew As String, safeNotes As String
    safeOld = MaskSecrets(oldVal)
    safeNew = MaskSecrets(newVal)
    safeNotes = MaskSecrets(notesText)
    
    curUser = GetCurrentUser()
    
    Set wsLog = ThisWorkbook.Worksheets("QUOTE_LOG")
    modProtection.UnprotectSheet wsLog
    Set tblLog = wsLog.ListObjects("tblQuoteLog")
    
    If tblLog.ListRows.count = 1 And Len(Trim$(CStr(tblLog.ListRows(1).Range.Cells(1, tblLog.ListColumns("Action").index).value))) = 0 Then
        Set newRow = tblLog.ListRows(1)
    Else
        Set newRow = tblLog.ListRows.Add
    End If
    
    logId = "LOG-" & Format$(newRow.index, "00000")
    
    With newRow.Range
        .Cells(1, tblLog.ListColumns("LogID").index).value = logId
        .Cells(1, tblLog.ListColumns("QuoteID").index).value = quoteId
        .Cells(1, tblLog.ListColumns("Timestamp").index).value = Now
        .Cells(1, tblLog.ListColumns("User").index).value = curUser
        .Cells(1, tblLog.ListColumns("Action").index).value = actionName
        .Cells(1, tblLog.ListColumns("OldValue").index).value = safeOld
        .Cells(1, tblLog.ListColumns("NewValue").index).value = safeNew
        .Cells(1, tblLog.ListColumns("SheetName").index).value = sheetName
        .Cells(1, tblLog.ListColumns("Status").index).value = statusVal
        .Cells(1, tblLog.ListColumns("Notes").index).value = safeNotes
    End With
    modProtection.ProtectSheet wsLog, userInterfaceOnly:=True
    
    LogAction = logId

CleanExit:
    mIsLogging = False
End Function

Private Function MaskSecrets(ByVal rawStr As String) As String
    Dim s As String
    s = rawStr
    If InStr(1, s, "password", vbTextCompare) > 0 Or _
       InStr(1, s, "pwd", vbTextCompare) > 0 Or _
       InStr(1, s, "api_key", vbTextCompare) > 0 Or _
       InStr(1, s, "secret", vbTextCompare) > 0 Then
        MaskSecrets = "[REDACTED]"
    Else
        MaskSecrets = s
    End If
End Function

Public Function RunAuditSelfTests() As String
    Dim testLogId As String
    Dim wsLog As Worksheet
    Dim tblLog As ListObject
    Dim r As Long
    Dim foundRow As ListRow
    
    On Error GoTo TestFail
    
    ' 1. Test LogAction
    testLogId = LogAction("Q-TEST-AUDIT", "TEST_ACTION", "OldVal", "NewVal", "QUOTE", "TEST_OK", "Audit self-test note")
    If Len(testLogId) = 0 Then
        RunAuditSelfTests = "FAIL: LogAction returned empty ID"
        Exit Function
    End If
    
    Set wsLog = ThisWorkbook.Worksheets("QUOTE_LOG")
    Set tblLog = wsLog.ListObjects("tblQuoteLog")
    
    ' Verify row exists
    For r = tblLog.ListRows.count To 1 Step -1
        If StrComp(Trim$(CStr(tblLog.ListRows(r).Range.Cells(1, tblLog.ListColumns("LogID").index).value)), testLogId, vbTextCompare) = 0 Then
            Set foundRow = tblLog.ListRows(r)
            Exit For
        End If
    Next r
    
    If foundRow Is Nothing Then
        RunAuditSelfTests = "FAIL: Log row not found in tblQuoteLog"
        Exit Function
    End If
    
    With foundRow.Range
        If .Cells(1, tblLog.ListColumns("QuoteID").index).value <> "Q-TEST-AUDIT" Then
            RunAuditSelfTests = "FAIL: QuoteID mismatch"
            Exit Function
        End If
        If .Cells(1, tblLog.ListColumns("Action").index).value <> "TEST_ACTION" Then
            RunAuditSelfTests = "FAIL: Action mismatch"
            Exit Function
        End If
        If .Cells(1, tblLog.ListColumns("User").index).value <> Environ$("USERNAME") And Len(CStr(.Cells(1, tblLog.ListColumns("User").index).value)) = 0 Then
            RunAuditSelfTests = "FAIL: User identity not captured"
            Exit Function
        End If
    End With
    
    ' 2. Test Secret Masking
    Dim secretLogId As String
    secretLogId = LogAction("Q-TEST-AUDIT", "SECURITY_TEST", "my_secret_pwd", "new_api_key_123")
    For r = tblLog.ListRows.count To 1 Step -1
        If StrComp(Trim$(CStr(tblLog.ListRows(r).Range.Cells(1, tblLog.ListColumns("LogID").index).value)), secretLogId, vbTextCompare) = 0 Then
            If tblLog.ListRows(r).Range.Cells(1, tblLog.ListColumns("OldValue").index).value <> "[REDACTED]" Then
                RunAuditSelfTests = "FAIL: Password was not masked in audit log"
                Exit Function
            End If
            tblLog.ListRows(r).Delete
            Exit For
        End If
    Next r
    
    ' Clean up test row
    On Error Resume Next
    foundRow.Delete
    On Error GoTo 0
    
    RunAuditSelfTests = "PASS"
    Exit Function
    
TestFail:
    RunAuditSelfTests = "FAIL: Exception - " & Err.Description
End Function

Public Function RunIdentitySelfTests() As String
    Dim u As String
    Dim testId As String
    Dim wsLog As Worksheet
    Dim tblLog As ListObject
    Dim r As Long
    Dim loggedUser As String
    Dim foundRow As ListRow
    
    On Error GoTo TestFail
    
    ' 1. Verify GetCurrentUser returns non-empty string
    u = GetCurrentUser()
    If Len(u) = 0 Then
        RunIdentitySelfTests = "FAIL: GetCurrentUser returned empty string"
        Exit Function
    End If
    
    ' 2. Verify no invalid filesystem characters
    If InStr(u, "\") > 0 Or InStr(u, "/") > 0 Or InStr(u, ":") > 0 Or _
       InStr(u, "*") > 0 Or InStr(u, "?") > 0 Or InStr(u, """") > 0 Or _
       InStr(u, "<") > 0 Or InStr(u, ">") > 0 Or InStr(u, "|") > 0 Or InStr(u, " ") > 0 Then
        RunIdentitySelfTests = "FAIL: GetCurrentUser contains unescaped invalid character"
        Exit Function
    End If
    
    ' 3. Verify LogAction captures GetCurrentUser in tblQuoteLog
    testId = LogAction("Q-TEST-IDENTITY", "TEST_IDENTITY", "Old", "New", "QUOTE", "SUCCESS", "Identity self-test")
    If Len(testId) = 0 Then
        RunIdentitySelfTests = "FAIL: LogAction returned empty ID"
        Exit Function
    End If
    
    Set wsLog = ThisWorkbook.Worksheets("QUOTE_LOG")
    Set tblLog = wsLog.ListObjects("tblQuoteLog")
    
    For r = tblLog.ListRows.count To 1 Step -1
        If StrComp(Trim$(CStr(tblLog.ListRows(r).Range.Cells(1, tblLog.ListColumns("LogID").index).value)), testId, vbTextCompare) = 0 Then
            Set foundRow = tblLog.ListRows(r)
            loggedUser = Trim$(CStr(foundRow.Range.Cells(1, tblLog.ListColumns("User").index).value))
            Exit For
        End If
    Next r
    
    If foundRow Is Nothing Then
        RunIdentitySelfTests = "FAIL: Test log row not found in tblQuoteLog"
        Exit Function
    End If
    
    If StrComp(loggedUser, u, vbTextCompare) <> 0 Then
        RunIdentitySelfTests = "FAIL: Logged user mismatch: expected '" & u & "', got '" & loggedUser & "'"
        foundRow.Delete
        Exit Function
    End If
    
    ' Clean up test row
    foundRow.Delete
    
    ' 4. Verify username alone does NOT bypass admin authorization
    Dim expectedAdminPwd As String
    expectedAdminPwd = modAdmin.AdminPassword()
    If Len(expectedAdminPwd) > 0 Then
        Dim transRes As String
        transRes = modApproval.TransitionQuoteState("Approved", "Testing unauthorized user", "wrong_password")
        If Left$(transRes, 5) <> "ERROR" Then
            RunIdentitySelfTests = "FAIL: Username allowed unauthorized approval bypass without correct password"
            Exit Function
        End If
    End If
    
    RunIdentitySelfTests = "PASS"
    Exit Function

TestFail:
    RunIdentitySelfTests = "FAIL: Exception - " & Err.Description
End Function
