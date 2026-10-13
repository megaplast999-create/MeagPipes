Attribute VB_Name = "modLog"
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
                basePath = Trim$(CStr(sRow.Range.Cells(1, tblSet.ListColumns("SettingValue").Index).Value))
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
    wsMeta.Range("A1").Value = "SettingKey"
    wsMeta.Range("B1").Value = "SettingValue"
    
    wsMeta.Range("A2").Value = "QuoteID": wsMeta.Range("B2").Value = enquiryNo
    wsMeta.Range("A3").Value = "Revision": wsMeta.Range("B3").Value = revNum
    wsMeta.Range("A4").Value = "QuoteDate": wsMeta.Range("B4").Value = header.QuoteDate
    wsMeta.Range("A5").Value = "Customer": wsMeta.Range("B5").Value = header.Customer
    wsMeta.Range("A6").Value = "ContactPerson": wsMeta.Range("B6").Value = header.ContactPerson
    wsMeta.Range("A7").Value = "DeliveryPlace": wsMeta.Range("B7").Value = header.DeliveryPlace
    wsMeta.Range("A8").Value = "PlantID": wsMeta.Range("B8").Value = header.PlantID
    wsMeta.Range("A9").Value = "PaymentTerms": wsMeta.Range("B9").Value = header.PaymentTerms
    wsMeta.Range("A10").Value = "ValidityDays": wsMeta.Range("B10").Value = header.ValidityDays
    wsMeta.Range("A11").Value = "PriceBasis": wsMeta.Range("B11").Value = header.PriceBasis
    wsMeta.Range("A12").Value = "LegacyMode": wsMeta.Range("B12").Value = IIf(header.LegacyMode, "YES", "NO")
    wsMeta.Range("A13").Value = "TotalMetres": wsMeta.Range("B13").Value = summary.TotalMetres
    wsMeta.Range("A14").Value = "TotalWeightMT": wsMeta.Range("B14").Value = summary.TotalWeightMT
    wsMeta.Range("A15").Value = "TotalPipeCount": wsMeta.Range("B15").Value = summary.TotalPipeCount
    wsMeta.Range("A16").Value = "TotalTrailers": wsMeta.Range("B16").Value = summary.TotalTrailers
    wsMeta.Range("A17").Value = "TotalExclGST": wsMeta.Range("B17").Value = summary.TotalExclGST
    wsMeta.Range("A18").Value = "GSTAmount": wsMeta.Range("B18").Value = summary.GSTAmount
    wsMeta.Range("A19").Value = "GrandTotalInclGST": wsMeta.Range("B19").Value = summary.GrandTotalInclGST
    wsMeta.Range("A20").Value = "SavedBy": wsMeta.Range("B20").Value = curUser
    wsMeta.Range("A21").Value = "SavedTimestamp": wsMeta.Range("B21").Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    ' Copy quote lines into SNAPSHOT_LINES
    Set wsLinesSnap = snapWb.Worksheets.Add(After:=wsMeta)
    wsLinesSnap.Name = "SNAPSHOT_LINES"
    
    numLines = tblLines.ListRows.Count
    numCols = tblLines.ListColumns.Count
    
    ' Copy headers
    For c = 1 To numCols
        wsLinesSnap.Cells(1, c).Value = tblLines.HeaderRowRange.Cells(1, c).Value
    Next c
    
    ' Copy line values
    If numLines > 0 Then
        For r = 1 To numLines
            For c = 1 To numCols
                wsLinesSnap.Cells(r + 1, c).Value = tblLines.DataBodyRange.Cells(r, c).Value
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
    Set tblLog = wsLog.ListObjects("tblQuoteLog")
    On Error GoTo 0
    
    If Not tblLog Is Nothing Then
        If tblLog.ListRows.Count = 1 And Len(Trim$(CStr(tblLog.ListRows(1).Range.Cells(1, tblLog.ListColumns("Action").Index).Value))) = 0 Then
            Set newRow = tblLog.ListRows(1)
        Else
            Set newRow = tblLog.ListRows.Add
        End If
        
        With newRow.Range
            .Cells(1, tblLog.ListColumns("LogID").Index).Value = "LOG-" & Format$(newRow.Index, "00000")
            .Cells(1, tblLog.ListColumns("QuoteID").Index).Value = enquiryNo
            .Cells(1, tblLog.ListColumns("Timestamp").Index).Value = Now
            .Cells(1, tblLog.ListColumns("User").Index).Value = curUser
            .Cells(1, tblLog.ListColumns("Action").Index).Value = "SNAPSHOT_SAVE"
            .Cells(1, tblLog.ListColumns("OldValue").Index).Value = ""
            .Cells(1, tblLog.ListColumns("NewValue").Index).Value = Format$(summary.TotalExclGST, "0")
            .Cells(1, tblLog.ListColumns("SheetName").Index).Value = "QUOTE"
            .Cells(1, tblLog.ListColumns("Status").Index).Value = "SAVED"
            .Cells(1, tblLog.ListColumns("Notes").Index).Value = "Snapshot saved: R" & CStr(revNum) & " (" & fName & ")"
        End With
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
    Set tblLines = wsLines.ListObjects("tblQuoteLines")
    
    ' Read metadata into dictionary
    Set metaDict = CreateObject("Scripting.Dictionary")
    For r = 2 To wsMeta.Cells(wsMeta.Rows.Count, 1).End(xlUp).Row
        k = Trim$(CStr(wsMeta.Cells(r, 1).Value))
        v = wsMeta.Cells(r, 2).Value
        If Len(k) > 0 Then metaDict(k) = v
    Next r
    
    ' Restore header fields on QUOTE sheet
    If metaDict.Exists("QuoteID") Then wsQ.Range("B4").Value = metaDict("QuoteID")
    If metaDict.Exists("QuoteDate") Then wsQ.Range("B5").Value = metaDict("QuoteDate")
    If metaDict.Exists("Customer") Then wsQ.Range("B8").Value = metaDict("Customer")
    If metaDict.Exists("ContactPerson") Then wsQ.Range("B9").Value = metaDict("ContactPerson")
    If metaDict.Exists("DeliveryPlace") Then wsQ.Range("B11").Value = metaDict("DeliveryPlace")
    If metaDict.Exists("PlantID") Then wsQ.Range("B14").Value = metaDict("PlantID")
    If metaDict.Exists("PaymentTerms") Then wsQ.Range("B15").Value = metaDict("PaymentTerms")
    If metaDict.Exists("ValidityDays") Then wsQ.Range("B16").Value = metaDict("ValidityDays")
    If metaDict.Exists("PriceBasis") Then wsQ.Range("B17").Value = metaDict("PriceBasis")
    
    ' Restore live summary card values on QUOTE sheet
    If metaDict.Exists("TotalMetres") Then wsQ.Range("E4").Value = metaDict("TotalMetres")
    If metaDict.Exists("TotalWeightMT") Then wsQ.Range("E5").Value = metaDict("TotalWeightMT")
    If metaDict.Exists("TotalPipeCount") Then wsQ.Range("E6").Value = metaDict("TotalPipeCount")
    If metaDict.Exists("TotalTrailers") Then wsQ.Range("E7").Value = metaDict("TotalTrailers")
    If metaDict.Exists("TotalExclGST") Then
        wsQ.Range("E8").Value = metaDict("TotalExclGST")
        restoredTotal = CDbl(metaDict("TotalExclGST"))
    End If
    If metaDict.Exists("GSTAmount") Then wsQ.Range("E9").Value = metaDict("GSTAmount")
    If metaDict.Exists("GrandTotalInclGST") Then wsQ.Range("E10").Value = metaDict("GrandTotalInclGST")
    wsQ.Range("E11").Value = "LOADED_SNAPSHOT"
    
    ' Restore lines into tblQuoteLines
    If tblLines.ListRows.Count > 0 Then
        tblLines.DataBodyRange.Delete
    End If
    
    numLines = wsLinesSnap.Cells(wsLinesSnap.Rows.Count, 1).End(xlUp).Row - 1
    numCols = tblLines.ListColumns.Count
    
    If numLines > 0 Then
        For r = 1 To numLines
            Dim newBodyRow As ListRow
            Set newBodyRow = tblLines.ListRows.Add
            For c = 1 To numCols
                newBodyRow.Range.Cells(1, c).Value = wsLinesSnap.Cells(r + 1, c).Value
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
    Set tblLog = wsLog.ListObjects("tblQuoteLog")
    On Error GoTo 0
    
    If Not tblLog Is Nothing Then
        If tblLog.ListRows.Count = 1 And Len(Trim$(CStr(tblLog.ListRows(1).Range.Cells(1, tblLog.ListColumns("Action").Index).Value))) = 0 Then
            Set newRow = tblLog.ListRows(1)
        Else
            Set newRow = tblLog.ListRows.Add
        End If
        
        With newRow.Range
            .Cells(1, tblLog.ListColumns("LogID").Index).Value = "LOG-" & Format$(newRow.Index, "00000")
            .Cells(1, tblLog.ListColumns("QuoteID").Index).Value = CStr(wsQ.Range("B4").Value)
            .Cells(1, tblLog.ListColumns("Timestamp").Index).Value = Now
            .Cells(1, tblLog.ListColumns("User").Index).Value = curUser
            .Cells(1, tblLog.ListColumns("Action").Index).Value = "SNAPSHOT_LOAD"
            .Cells(1, tblLog.ListColumns("OldValue").Index).Value = ""
            .Cells(1, tblLog.ListColumns("NewValue").Index).Value = Format$(restoredTotal, "0")
            .Cells(1, tblLog.ListColumns("SheetName").Index).Value = "QUOTE"
            .Cells(1, tblLog.ListColumns("Status").Index).Value = "LOADED"
            .Cells(1, tblLog.ListColumns("Notes").Index).Value = "Snapshot reloaded from " & fullPath
        End With
    End If
    
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
    origTotal = CDbl(wsQ.Range("E8").Value)
    If origTotal = 0 Then origTotal = 178801050#
    
    ' Step 2: Save snapshot R0
    snapPath = SaveQuote(qId, 0)
    If Not fso.FileExists(snapPath) Then
        RunSnapshotSelfTests = "FAIL: Snapshot file not found at " & snapPath
        Exit Function
    End If
    
    ' Step 3: Mutate active quote (change quantity of Line 1 to 9999)
    tblLines.ListRows(1).Range.Cells(1, tblLines.ListColumns("Quantity").Index).Value = 9999
    ' Recalculate quote
    Dim dummyHead As QuoteHeaderData
    Dim dummySum As QuoteSummaryData
    dummyHead = ReadQuoteSheet()
    dummySum = CalcAllLines(dummyHead)
    UpdateLiveSummary dummySum, dummyHead
    
    mutatedTotal = CDbl(wsQ.Range("E8").Value)
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
    restoredTotal = CDbl(wsQ.Range("E8").Value)
    If Abs(restoredTotal - origTotal) > 0.01 Then
        RunSnapshotSelfTests = "FAIL: Restored total mismatch. Expected " & CStr(origTotal) & ", got " & CStr(restoredTotal)
        Exit Function
    End If
    
    ' Verify 4 lines restored
    If tblLines.ListRows.Count <> 4 Then
        RunSnapshotSelfTests = "FAIL: Expected 4 restored lines, got " & CStr(tblLines.ListRows.Count)
        Exit Function
    End If
    
    p1 = CDbl(tblLines.ListRows(1).Range.Cells(1, tblLines.ListColumns("UnitPrice").Index).Value)
    p2 = CDbl(tblLines.ListRows(2).Range.Cells(1, tblLines.ListColumns("UnitPrice").Index).Value)
    p3 = CDbl(tblLines.ListRows(3).Range.Cells(1, tblLines.ListColumns("UnitPrice").Index).Value)
    p4 = CDbl(tblLines.ListRows(4).Range.Cells(1, tblLines.ListColumns("UnitPrice").Index).Value)
    
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
    Set tblLog = wsLog.ListObjects("tblQuoteLog")
    
    If tblLog.ListRows.Count = 1 And Len(Trim$(CStr(tblLog.ListRows(1).Range.Cells(1, tblLog.ListColumns("Action").Index).Value))) = 0 Then
        Set newRow = tblLog.ListRows(1)
    Else
        Set newRow = tblLog.ListRows.Add
    End If
    
    logId = "LOG-" & Format$(newRow.Index, "00000")
    
    With newRow.Range
        .Cells(1, tblLog.ListColumns("LogID").Index).Value = logId
        .Cells(1, tblLog.ListColumns("QuoteID").Index).Value = quoteId
        .Cells(1, tblLog.ListColumns("Timestamp").Index).Value = Now
        .Cells(1, tblLog.ListColumns("User").Index).Value = curUser
        .Cells(1, tblLog.ListColumns("Action").Index).Value = actionName
        .Cells(1, tblLog.ListColumns("OldValue").Index).Value = safeOld
        .Cells(1, tblLog.ListColumns("NewValue").Index).Value = safeNew
        .Cells(1, tblLog.ListColumns("SheetName").Index).Value = sheetName
        .Cells(1, tblLog.ListColumns("Status").Index).Value = statusVal
        .Cells(1, tblLog.ListColumns("Notes").Index).Value = safeNotes
    End With
    
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
    For r = tblLog.ListRows.Count To 1 Step -1
        If StrComp(Trim$(CStr(tblLog.ListRows(r).Range.Cells(1, tblLog.ListColumns("LogID").Index).Value)), testLogId, vbTextCompare) = 0 Then
            Set foundRow = tblLog.ListRows(r)
            Exit For
        End If
    Next r
    
    If foundRow Is Nothing Then
        RunAuditSelfTests = "FAIL: Log row not found in tblQuoteLog"
        Exit Function
    End If
    
    With foundRow.Range
        If .Cells(1, tblLog.ListColumns("QuoteID").Index).Value <> "Q-TEST-AUDIT" Then
            RunAuditSelfTests = "FAIL: QuoteID mismatch"
            Exit Function
        End If
        If .Cells(1, tblLog.ListColumns("Action").Index).Value <> "TEST_ACTION" Then
            RunAuditSelfTests = "FAIL: Action mismatch"
            Exit Function
        End If
        If .Cells(1, tblLog.ListColumns("User").Index).Value <> Environ$("USERNAME") And Len(CStr(.Cells(1, tblLog.ListColumns("User").Index).Value)) = 0 Then
            RunAuditSelfTests = "FAIL: User identity not captured"
            Exit Function
        End If
    End With
    
    ' 2. Test Secret Masking
    Dim secretLogId As String
    secretLogId = LogAction("Q-TEST-AUDIT", "SECURITY_TEST", "my_secret_pwd", "new_api_key_123")
    For r = tblLog.ListRows.Count To 1 Step -1
        If StrComp(Trim$(CStr(tblLog.ListRows(r).Range.Cells(1, tblLog.ListColumns("LogID").Index).Value)), secretLogId, vbTextCompare) = 0 Then
            If tblLog.ListRows(r).Range.Cells(1, tblLog.ListColumns("OldValue").Index).Value <> "[REDACTED]" Then
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
    
    For r = tblLog.ListRows.Count To 1 Step -1
        If StrComp(Trim$(CStr(tblLog.ListRows(r).Range.Cells(1, tblLog.ListColumns("LogID").Index).Value)), testId, vbTextCompare) = 0 Then
            Set foundRow = tblLog.ListRows(r)
            loggedUser = Trim$(CStr(foundRow.Range.Cells(1, tblLog.ListColumns("User").Index).Value))
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

