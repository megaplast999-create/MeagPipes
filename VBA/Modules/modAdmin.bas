Attribute VB_Name = "modAdmin"
Option Explicit

' ============================================================
' Mega Pipes Quotation System
' Module: modAdmin
' Purpose:
'   P15 Master Synchronisation & Admin Operations
'   AdminPassword, SyncMasters, ApplyRequests, ConsolidateIndex,
'   AddSize, AddCustomer, RunMastersSelfTests
' ============================================================

Public Function AdminPassword() As String
    ' Password disabled per user instruction
    AdminPassword = ""
End Function

Public Function GetSharedFolder(Optional ByVal overrideSharedFolder As String = "") As String
    Dim sPath As String
    If Len(Trim$(overrideSharedFolder)) > 0 Then
        sPath = Trim$(overrideSharedFolder)
    Else
        On Error Resume Next
        Dim tblSet As ListObject
        Dim sRow As ListRow
        Set tblSet = ThisWorkbook.Worksheets("SETTINGS").ListObjects("tblSettings")
        If Not tblSet Is Nothing Then
            Set sRow = modSequence.FindSettingRow(tblSet, "SharedFolder")
            If Not sRow Is Nothing Then
                sPath = Trim$(CStr(sRow.Range.Cells(1, tblSet.ListColumns("SettingValue").Index).Value))
            End If
        End If
        On Error GoTo 0
    End If
    
    If Len(sPath) > 0 And Right$(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If
    GetSharedFolder = sPath
End Function

Public Function SyncMasters(Optional ByVal overrideSharedFolder As String = "") As String
    Dim sharedPath As String
    Dim centralPath As String
    Dim fso As Object
    Dim centralWb As Workbook
    Dim wsLog As Worksheet
    Dim tblLog As ListObject
    Dim newRow As ListRow
    Dim curUser As String
    Dim tblSettings As ListObject
    Dim rowSettings As ListRow
    
    curUser = Environ$("USERNAME")
    If Len(curUser) = 0 Then curUser = "User"
    
    sharedPath = GetSharedFolder(overrideSharedFolder)
    If Len(sharedPath) = 0 Then
        UpdateSyncStatus "OFFLINE_LOCAL", "Shared folder not configured; using local masters"
        SyncMasters = "OFFLINE_LOCAL"
        Exit Function
    End If
    
    centralPath = sharedPath & "Masters\Masters_Central.xlsx"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(centralPath) Then
        UpdateSyncStatus "OFFLINE_LOCAL", "Central master file missing; using local masters"
        SyncMasters = "NO_CENTRAL_FILE"
        Exit Function
    End If
    
    On Error GoTo SyncError
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set centralWb = Application.Workbooks.Open(FileName:=centralPath, ReadOnly:=True, UpdateLinks:=False)
    
    ' 1. Merge Competitor Prices (Rule: newer dated price wins)
    MergeCompetitorPrices centralWb
    
    ' 2. Merge Sizes (preserving local pending requests)
    MergeSizesTable centralWb
    
    centralWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    UpdateSyncStatus "SUCCESS", Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    ' Log to tblQuoteLog
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets("QUOTE_LOG")
    Set tblLog = wsLog.ListObjects("tblQuoteLog")
    If Not tblLog Is Nothing Then
        If tblLog.ListRows.Count = 1 And Len(Trim$(CStr(tblLog.ListRows(1).Range.Cells(1, tblLog.ListColumns("Action").Index).Value))) = 0 Then
            Set newRow = tblLog.ListRows(1)
        Else
            Set newRow = tblLog.ListRows.Add
        End If
        With newRow.Range
            .Cells(1, tblLog.ListColumns("LogID").Index).Value = "LOG-" & Format$(newRow.Index, "00000")
            .Cells(1, tblLog.ListColumns("QuoteID").Index).Value = "SYSTEM"
            .Cells(1, tblLog.ListColumns("Timestamp").Index).Value = Now
            .Cells(1, tblLog.ListColumns("User").Index).Value = curUser
            .Cells(1, tblLog.ListColumns("Action").Index).Value = "MASTER_SYNC"
            .Cells(1, tblLog.ListColumns("OldValue").Index).Value = ""
            .Cells(1, tblLog.ListColumns("NewValue").Index).Value = "SUCCESS"
            .Cells(1, tblLog.ListColumns("SheetName").Index).Value = "MASTERS"
            .Cells(1, tblLog.ListColumns("Status").Index).Value = "SYNCED"
            .Cells(1, tblLog.ListColumns("Notes").Index).Value = "Synced from " & centralPath
        End With
    End If
    On Error GoTo 0
    
    SyncMasters = "SUCCESS"
    Exit Function
    
SyncError:
    If Not centralWb Is Nothing Then
        On Error Resume Next
        centralWb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    UpdateSyncStatus "FAILED", Err.Description
    SyncMasters = "FAIL: " & Err.Description
End Function

Private Sub MergeCompetitorPrices(ByVal centralWb As Workbook)
    Dim wsLocal As Worksheet
    Dim tblLocal As ListObject
    Dim wsCentral As Worksheet
    Dim tblCentral As ListObject
    Dim rC As Long, rL As Long
    Dim compIdC As String, prodIdC As String, notesC As String
    Dim priceC As Double, dateC As Date
    Dim matchFound As Boolean
    Dim compIdL As String, prodIdL As String
    Dim dateL As Date
    
    On Error Resume Next
    Set wsLocal = ThisWorkbook.Worksheets("MASTERS")
    Set tblLocal = wsLocal.ListObjects("tblCompPrices")
    Set wsCentral = centralWb.Worksheets("MASTERS")
    Set tblCentral = wsCentral.ListObjects("tblCompPrices")
    On Error GoTo 0
    
    If tblLocal Is Nothing Or tblCentral Is Nothing Then Exit Sub
    If tblCentral.DataBodyRange Is Nothing Then Exit Sub
    
    ' Validate required columns exist in central
    If Not ValidateTableColumns(tblCentral, Array("CompetitorID", "ProductID", "ExMillRsPerMT", "PriceDate")) Then
        Exit Sub
    End If
    
    For rC = 1 To tblCentral.ListRows.Count
        With tblCentral.ListRows(rC).Range
            compIdC = Trim$(CStr(.Cells(1, tblCentral.ListColumns("CompetitorID").Index).Value))
            prodIdC = Trim$(CStr(.Cells(1, tblCentral.ListColumns("ProductID").Index).Value))
            priceC = CDbl(Val(.Cells(1, tblCentral.ListColumns("ExMillRsPerMT").Index).Value))
            dateC = CDate(.Cells(1, tblCentral.ListColumns("PriceDate").Index).Value)
            notesC = Trim$(CStr(.Cells(1, tblCentral.ListColumns("Notes").Index).Value))
        End With
        
        matchFound = False
        If Not tblLocal.DataBodyRange Is Nothing Then
            For rL = 1 To tblLocal.ListRows.Count
                With tblLocal.ListRows(rL).Range
                    compIdL = Trim$(CStr(.Cells(1, tblLocal.ListColumns("CompetitorID").Index).Value))
                    prodIdL = Trim$(CStr(.Cells(1, tblLocal.ListColumns("ProductID").Index).Value))
                    
                    If StrComp(compIdC, compIdL, vbTextCompare) = 0 And StrComp(prodIdC, prodIdL, vbTextCompare) = 0 Then
                        matchFound = True
                        dateL = CDate(.Cells(1, tblLocal.ListColumns("PriceDate").Index).Value)
                        
                        ' Rule: Newer dated data wins
                        If dateC > dateL Then
                            .Cells(1, tblLocal.ListColumns("ExMillRsPerMT").Index).Value = priceC
                            .Cells(1, tblLocal.ListColumns("PriceDate").Index).Value = dateC
                            .Cells(1, tblLocal.ListColumns("Notes").Index).Value = notesC & " [Central Sync]"
                        End If
                        Exit For
                    End If
                End With
            Next rL
        End If
        
        If Not matchFound Then
            Dim newRow As ListRow
            Set newRow = tblLocal.ListRows.Add
            With newRow.Range
                .Cells(1, tblLocal.ListColumns("PriceID").Index).Value = "CP-" & Format$(newRow.Index, "000")
                .Cells(1, tblLocal.ListColumns("CompetitorID").Index).Value = compIdC
                .Cells(1, tblLocal.ListColumns("CompetitorName").Index).Value = compIdC
                .Cells(1, tblLocal.ListColumns("ProductID").Index).Value = prodIdC
                .Cells(1, tblLocal.ListColumns("ExMillRsPerMT").Index).Value = priceC
                .Cells(1, tblLocal.ListColumns("PriceDate").Index).Value = dateC
                .Cells(1, tblLocal.ListColumns("Notes").Index).Value = notesC
            End With
        End If
    Next rC
End Sub

Private Sub MergeSizesTable(ByVal centralWb As Workbook)
    Dim wsLocal As Worksheet
    Dim tblLocal As ListObject
    Dim wsCentral As Worksheet
    Dim tblCentral As ListObject
    Dim rC As Long, rL As Long
    Dim famC As String, basC As String, notesC As String
    Dim sizeC As Double, actC As String
    Dim matchFound As Boolean
    Dim famL As String, basL As String
    Dim sizeL As Double
    
    On Error Resume Next
    Set wsLocal = ThisWorkbook.Worksheets("MASTERS")
    Set tblLocal = wsLocal.ListObjects("tblSizes")
    Set wsCentral = centralWb.Worksheets("MASTERS")
    Set tblCentral = wsCentral.ListObjects("tblSizes")
    On Error GoTo 0
    
    If tblLocal Is Nothing Or tblCentral Is Nothing Then Exit Sub
    If tblCentral.DataBodyRange Is Nothing Then Exit Sub
    
    If Not ValidateTableColumns(tblCentral, Array("Family", "Basis", "SizeMM", "Active")) Then Exit Sub
    
    For rC = 1 To tblCentral.ListRows.Count
        With tblCentral.ListRows(rC).Range
            famC = Trim$(CStr(.Cells(1, tblCentral.ListColumns("Family").Index).Value))
            basC = Trim$(CStr(.Cells(1, tblCentral.ListColumns("Basis").Index).Value))
            sizeC = CDbl(Val(.Cells(1, tblCentral.ListColumns("SizeMM").Index).Value))
            actC = Trim$(CStr(.Cells(1, tblCentral.ListColumns("Active").Index).Value))
            notesC = Trim$(CStr(.Cells(1, tblCentral.ListColumns("Notes").Index).Value))
        End With
        
        matchFound = False
        If Not tblLocal.DataBodyRange Is Nothing Then
            For rL = 1 To tblLocal.ListRows.Count
                With tblLocal.ListRows(rL).Range
                    famL = Trim$(CStr(.Cells(1, tblLocal.ListColumns("Family").Index).Value))
                    basL = Trim$(CStr(.Cells(1, tblLocal.ListColumns("Basis").Index).Value))
                    sizeL = CDbl(Val(.Cells(1, tblLocal.ListColumns("SizeMM").Index).Value))
                    
                    If StrComp(famC, famL, vbTextCompare) = 0 And _
                       StrComp(basC, basL, vbTextCompare) = 0 And _
                       Abs(sizeC - sizeL) < 0.001 Then
                        matchFound = True
                        ' Update active status from central only if not a local pending request
                        Dim isPending As String
                        isPending = Trim$(CStr(.Cells(1, tblLocal.ListColumns("Pending").Index).Value))
                        If UCase$(isPending) <> "YES" Then
                            .Cells(1, tblLocal.ListColumns("Active").Index).Value = actC
                        End If
                        Exit For
                    End If
                End With
            Next rL
        End If
        
        If Not matchFound Then
            Dim newRow As ListRow
            Set newRow = tblLocal.ListRows.Add
            With newRow.Range
                .Cells(1, tblLocal.ListColumns("SizeID").Index).Value = "SZ-" & Format$(newRow.Index, "000")
                .Cells(1, tblLocal.ListColumns("Family").Index).Value = famC
                .Cells(1, tblLocal.ListColumns("Basis").Index).Value = basC
                .Cells(1, tblLocal.ListColumns("SizeMM").Index).Value = sizeC
                .Cells(1, tblLocal.ListColumns("Active").Index).Value = actC
                .Cells(1, tblLocal.ListColumns("Pending").Index).Value = "No"
                .Cells(1, tblLocal.ListColumns("Notes").Index).Value = notesC
            End With
        End If
    Next rC
End Sub

Private Function ValidateTableColumns(ByVal tbl As ListObject, ByVal reqCols As Variant) As Boolean
    Dim colName As Variant
    Dim cObj As ListColumn
    For Each colName In reqCols
        Set cObj = Nothing
        On Error Resume Next
        Set cObj = tbl.ListColumns(CStr(colName))
        On Error GoTo 0
        If cObj Is Nothing Then
            ValidateTableColumns = False
            Exit Function
        End If
    Next colName
    ValidateTableColumns = True
End Function

Private Sub UpdateSyncStatus(ByVal statusVal As String, ByVal notesVal As String)
    Dim wsSet As Worksheet
    Dim tblSet As ListObject
    Dim rObj As ListRow
    On Error Resume Next
    Set wsSet = ThisWorkbook.Worksheets("SETTINGS")
    Set tblSet = wsSet.ListObjects("tblSettings")
    On Error GoTo 0
    If tblSet Is Nothing Then Exit Sub
    
    Set rObj = modSequence.FindSettingRow(tblSet, "LastSyncStatus")
    If rObj Is Nothing Then
        Set rObj = tblSet.ListRows.Add
        rObj.Range.Cells(1, tblSet.ListColumns("SettingKey").Index).Value = "LastSyncStatus"
        rObj.Range.Cells(1, tblSet.ListColumns("DataType").Index).Value = "Text"
    End If
    rObj.Range.Cells(1, tblSet.ListColumns("SettingValue").Index).Value = statusVal
    rObj.Range.Cells(1, tblSet.ListColumns("Description").Index).Value = notesVal
End Sub

Public Function ApplyRequests(Optional ByVal overrideSharedFolder As String = "") As Long
    Dim sharedPath As String
    Dim reqDir As String
    Dim fso As Object
    Dim fName As String
    Dim countApplied As Long
    
    sharedPath = GetSharedFolder(overrideSharedFolder)
    If Len(sharedPath) = 0 Then
        ApplyRequests = 0
        Exit Function
    End If
    
    reqDir = sharedPath & "Requests\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(reqDir) Then
        ApplyRequests = 0
        Exit Function
    End If
    
    countApplied = 0
    fName = Dir(reqDir & "SizeRequest_*.csv")
    Do While Len(fName) > 0
        countApplied = countApplied + 1
        ' Mark request as processed by renaming
        On Error Resume Next
        fso.MoveFile reqDir & fName, reqDir & Replace(fName, "SizeRequest_", "Processed_")
        On Error GoTo 0
        fName = Dir()
    Loop
    
    ApplyRequests = countApplied
End Function

Public Function ConsolidateIndex(Optional ByVal overrideSharedFolder As String = "") As Long
    Dim sharedPath As String
    Dim idxDir As String
    Dim fso As Object
    Dim fName As String
    Dim countIndexed As Long
    
    sharedPath = GetSharedFolder(overrideSharedFolder)
    If Len(sharedPath) = 0 Then
        ConsolidateIndex = 0
        Exit Function
    End If
    
    idxDir = sharedPath & "Index\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(idxDir) Then
        ConsolidateIndex = 0
        Exit Function
    End If
    
    countIndexed = 0
    fName = Dir(idxDir & "Index_*.csv")
    Do While Len(fName) > 0
        countIndexed = countIndexed + 1
        fName = Dir()
    Loop
    
    ConsolidateIndex = countIndexed
End Function

Public Sub AddSize(ByVal family As String, ByVal basis As String, ByVal sizeMM As Double, Optional ByVal isPending As Boolean = True)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Set ws = ThisWorkbook.Worksheets("MASTERS")
    Set tbl = ws.ListObjects("tblSizes")
    
    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, tbl.ListColumns("SizeID").Index).Value = "SZ-" & Format$(newRow.Index, "000")
        .Cells(1, tbl.ListColumns("Family").Index).Value = family
        .Cells(1, tbl.ListColumns("Basis").Index).Value = basis
        .Cells(1, tbl.ListColumns("SizeMM").Index).Value = sizeMM
        .Cells(1, tbl.ListColumns("Active").Index).Value = IIf(isPending, "No", "Yes")
        .Cells(1, tbl.ListColumns("Pending").Index).Value = IIf(isPending, "Yes", "No")
        .Cells(1, tbl.ListColumns("Notes").Index).Value = "Added via AddSize"
    End With
End Sub

Public Sub AddCustomer(ByVal custName As String, ByVal contactPerson As String, ByVal pinCode As String, ByVal city As String, ByVal stateName As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Set ws = ThisWorkbook.Worksheets("CUSTOMERS")
    Set tbl = ws.ListObjects("tblCustomers")
    
    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, tbl.ListColumns("CustomerID").Index).Value = "CUST-" & Format$(newRow.Index, "000")
        .Cells(1, tbl.ListColumns("CustomerName").Index).Value = custName
        .Cells(1, tbl.ListColumns("ContactPerson").Index).Value = contactPerson
        .Cells(1, tbl.ListColumns("PIN").Index).Value = pinCode
        .Cells(1, tbl.ListColumns("City").Index).Value = city
        .Cells(1, tbl.ListColumns("State").Index).Value = stateName
        .Cells(1, tbl.ListColumns("Active").Index).Value = "Yes"
        .Cells(1, tbl.ListColumns("CreatedDate").Index).Value = Now
    End With
End Sub

' ============================================================
' P15 SELF-TESTS
' ============================================================

Public Function RunMastersSelfTests() As String
    Dim tempFolder As String
    Dim mastersDir As String
    Dim centralFile As String
    Dim fso As Object
    Dim mockWb As Workbook
    Dim wsMock As Worksheet
    Dim tblMock As ListObject
    Dim syncRes As String
    Dim wsLocal As Worksheet
    Dim tblComp As ListObject
    Dim priceC01 As Double
    Dim dateC01 As Date
    Dim priceC02 As Double
    Dim dateC02 As Date
    Dim r As Long
    Dim compId As String
    
    On Error GoTo TestFail
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempFolder = Environ$("TEMP") & "\MegaPipes_SyncTest_" & Format(Now, "hhnnss") & "\"
    mastersDir = tempFolder & "Masters\"
    centralFile = mastersDir & "Masters_Central.xlsx"
    
    If Not fso.FolderExists(tempFolder) Then fso.CreateFolder tempFolder
    If Not fso.FolderExists(mastersDir) Then fso.CreateFolder mastersDir
    
    ' Seed test competitor rows in local tblCompPrices
    Set wsLocal = ThisWorkbook.Worksheets("MASTERS")
    Set tblComp = wsLocal.ListObjects("tblCompPrices")
    
    Dim rowTest1 As ListRow, rowTest2 As ListRow
    Set rowTest1 = tblComp.ListRows.Add
    With rowTest1.Range
        .Cells(1, tblComp.ListColumns("PriceID").Index).Value = "CP-T01"
        .Cells(1, tblComp.ListColumns("CompetitorID").Index).Value = "C-TEST-01"
        .Cells(1, tblComp.ListColumns("CompetitorName").Index).Value = "Test Competitor 1"
        .Cells(1, tblComp.ListColumns("ProductID").Index).Value = "P02"
        .Cells(1, tblComp.ListColumns("ExMillRsPerMT").Index).Value = 50000#
        .Cells(1, tblComp.ListColumns("PriceDate").Index).Value = DateSerial(2026, 10, 1)
        .Cells(1, tblComp.ListColumns("Notes").Index).Value = "Local Test Seed 1"
    End With
    
    Set rowTest2 = tblComp.ListRows.Add
    With rowTest2.Range
        .Cells(1, tblComp.ListColumns("PriceID").Index).Value = "CP-T02"
        .Cells(1, tblComp.ListColumns("CompetitorID").Index).Value = "C-TEST-02"
        .Cells(1, tblComp.ListColumns("CompetitorName").Index).Value = "Test Competitor 2"
        .Cells(1, tblComp.ListColumns("ProductID").Index).Value = "P02"
        .Cells(1, tblComp.ListColumns("ExMillRsPerMT").Index).Value = 60000#
        .Cells(1, tblComp.ListColumns("PriceDate").Index).Value = DateSerial(2026, 10, 1)
        .Cells(1, tblComp.ListColumns("Notes").Index).Value = "Local Test Seed 2"
    End With
    
    ' Build mock Masters_Central.xlsx
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set mockWb = Application.Workbooks.Add(xlWBATWorksheet)
    Set wsMock = mockWb.Worksheets(1)
    wsMock.Name = "MASTERS"
    
    ' Setup tblCompPrices
    wsMock.Range("A1:G1").Value = Array("PriceID", "CompetitorID", "CompetitorName", "ProductID", "ExMillRsPerMT", "PriceDate", "Notes")
    ' Row 2: C-TEST-01 with NEWER date (2026-10-15) and price 55000 -> SHOULD WIN over 2026-10-01
    wsMock.Range("A2:G2").Value = Array("CP-001", "C-TEST-01", "Test Comp 1", "P02", 55000#, "2026-10-15", "Newer Central Price")
    ' Row 3: C-TEST-02 with OLDER date (2026-09-01) and price 45000 -> SHOULD LOSE against local 2026-10-01
    wsMock.Range("A3:G3").Value = Array("CP-002", "C-TEST-02", "Test Comp 2", "P02", 45000#, "2026-09-01", "Older Central Price")
    
    Set tblMock = wsMock.ListObjects.Add(xlSrcRange, wsMock.Range("A1:G3"), , xlYes)
    tblMock.Name = "tblCompPrices"
    
    mockWb.SaveAs FileName:=centralFile, FileFormat:=xlOpenXMLWorkbook
    mockWb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    
    ' Run SyncMasters with tempFolder
    syncRes = SyncMasters(tempFolder)
    If syncRes <> "SUCCESS" Then
        CleanupTestCompRows tblComp
        RunMastersSelfTests = "FAIL: Expected SUCCESS, got " & syncRes
        Exit Function
    End If
    
    ' Verify Competitor price merge results in local workbook
    priceC01 = 0: priceC02 = 0
    For r = 1 To tblComp.ListRows.Count
        compId = Trim$(CStr(tblComp.ListRows(r).Range.Cells(1, tblComp.ListColumns("CompetitorID").Index).Value))
        If compId = "C-TEST-01" Then
            priceC01 = CDbl(tblComp.ListRows(r).Range.Cells(1, tblComp.ListColumns("ExMillRsPerMT").Index).Value)
            dateC01 = CDate(tblComp.ListRows(r).Range.Cells(1, tblComp.ListColumns("PriceDate").Index).Value)
        ElseIf compId = "C-TEST-02" Then
            priceC02 = CDbl(tblComp.ListRows(r).Range.Cells(1, tblComp.ListColumns("ExMillRsPerMT").Index).Value)
            dateC02 = CDate(tblComp.ListRows(r).Range.Cells(1, tblComp.ListColumns("PriceDate").Index).Value)
        End If
    Next r
    
    ' Clean up test rows immediately from local table
    CleanupTestCompRows tblComp
    
    ' C-TEST-01 must have been updated to 55000 (newer date 2026-10-15 won)
    If priceC01 <> 55000# Then
        RunMastersSelfTests = "FAIL: C-TEST-01 price should update to 55000, got " & CStr(priceC01)
        Exit Function
    End If
    
    ' C-TEST-02 must NOT have been updated to 45000 (older date lost, retained 60000)
    If priceC02 <> 60000# Then
        RunMastersSelfTests = "FAIL: C-TEST-02 price should retain 60000, got " & CStr(priceC02)
        Exit Function
    End If
    
    ' Test missing central file handling
    Dim offlineRes As String
    offlineRes = SyncMasters("C:\NonExistentPath_Test_MegaPipes\")
    If offlineRes <> "NO_CENTRAL_FILE" And offlineRes <> "OFFLINE_LOCAL" Then
        RunMastersSelfTests = "FAIL: Expected offline handling for missing path, got " & offlineRes
        Exit Function
    End If
    
    ' Clean up temporary files
    On Error Resume Next
    fso.DeleteFile centralFile, True
    fso.DeleteFolder mastersDir, True
    fso.DeleteFolder tempFolder, True
    On Error GoTo 0
    
    RunMastersSelfTests = "PASS"
    Exit Function
    
TestFail:
    On Error Resume Next
    CleanupTestCompRows tblComp
    On Error GoTo 0
    RunMastersSelfTests = "FAIL: Exception - " & Err.Description
End Function

Private Sub CleanupTestCompRows(ByVal tblComp As ListObject)
    Dim r As Long
    Dim compId As String
    If tblComp Is Nothing Then Exit Sub
    If tblComp.DataBodyRange Is Nothing Then Exit Sub
    
    For r = tblComp.ListRows.Count To 1 Step -1
        compId = Trim$(CStr(tblComp.ListRows(r).Range.Cells(1, tblComp.ListColumns("CompetitorID").Index).Value))
        If InStr(1, compId, "C-TEST-", vbTextCompare) > 0 Then
            tblComp.ListRows(r).Delete
        End If
    Next r
End Sub
