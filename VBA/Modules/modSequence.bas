Attribute VB_Name = "modSequence"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ============================================================
' P13 ENQUIRY NUMBERING & EXCLUSIVE FILE LOCK ENGINE
' ============================================================

Public Function GetFiscalYear(Optional ByVal targetDate As Date = 0) As String
    Dim d As Date
    Dim yr As Long, mo As Long
    Dim y1 As Long, y2 As Long
    
    If targetDate = 0 Then
        d = Date
    Else
        d = targetDate
    End If
    
    yr = Year(d)
    mo = Month(d)
    
    ' In India, fiscal year runs from April 1 to March 31
    If mo >= 4 Then
        y1 = yr Mod 100
        y2 = (yr + 1) Mod 100
    Else
        y1 = (yr - 1) Mod 100
        y2 = yr Mod 100
    End If
    
    GetFiscalYear = Format(y1, "00") & "-" & Format(y2, "00")
End Function

Public Function AllocateOfferNumber(ByVal enquiryNo As String) As String
    If InStr(1, enquiryNo, "MKTG/ENQ/", vbTextCompare) > 0 Then
        AllocateOfferNumber = Replace(enquiryNo, "MKTG/ENQ/", "MEGA/OFFER/")
    ElseIf InStr(1, enquiryNo, "TEMP", vbTextCompare) > 0 Then
        AllocateOfferNumber = "MEGA/OFFER/" & GetFiscalYear() & "/TEMP-" & Format(Now, "yyyymmddhhnnss")
    Else
        AllocateOfferNumber = "MEGA/OFFER/" & GetFiscalYear() & "/" & Trim$(enquiryNo)
    End If
End Function

Public Function AllocateEnquiryNumber(Optional ByVal overrideSharedFolder As String = "") As String
    Dim settings As ListObject
    Dim sharedPath As String
    Dim fy As String
    
    fy = GetFiscalYear()
    
    ' Determine shared folder configuration
    If Len(Trim$(overrideSharedFolder)) > 0 Then
        sharedPath = Trim$(overrideSharedFolder)
    Else
        On Error Resume Next
        Set settings = ThisWorkbook.Worksheets("SETTINGS").ListObjects("tblSettings")
        If Not settings Is Nothing Then
            Dim sRow As ListRow
            Set sRow = FindSettingRow(settings, "SharedFolder")
            If Not sRow Is Nothing Then
                sharedPath = Trim$(CStr(sRow.Range.Cells(1, settings.ListColumns("SettingValue").Index).Value))
            End If
        End If
        On Error GoTo 0
    End If
    
    If Len(sharedPath) > 0 Then
        ' Production mode: shared network folder with exclusive file lock
        AllocateEnquiryNumber = AllocateFromSharedFolder(sharedPath, fy)
    Else
        ' DEV mode: local sequential counter in tblSettings
        AllocateEnquiryNumber = AllocateFromLocalDev(fy)
    End If
End Function

Private Function AllocateFromSharedFolder(ByVal sharedPath As String, ByVal fy As String) As String
    Dim f As Integer
    Dim counterDir As String
    Dim lockFile As String
    Dim retries As Integer
    Dim lockAcquired As Boolean
    Dim serialVal As Long
    Dim lineContent As String
    Dim fso As Object
    
    If Right$(sharedPath, 1) <> "\" Then sharedPath = sharedPath & "\"
    counterDir = sharedPath & "Counter"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(counterDir) Then
        On Error Resume Next
        fso.CreateFolder counterDir
        On Error GoTo 0
    End If
    
    lockFile = counterDir & "\enquiry_" & fy & ".txt"
    retries = 0
    lockAcquired = False
    
    Do While retries < 20
        Err.Clear
        On Error Resume Next
        f = FreeFile
        Open lockFile For Binary Access Read Write Lock Read Write As #f
        If Err.Number = 0 Then
            On Error GoTo 0
            lockAcquired = True
            
            ' Read current counter
            lineContent = ""
            If LOF(f) > 0 Then
                Dim fileData As String
                fileData = Space$(LOF(f))
                Get #f, 1, fileData
                lineContent = Trim$(fileData)
            End If
            
            If Len(lineContent) > 0 And IsNumeric(lineContent) Then
                serialVal = CLng(Val(lineContent))
            Else
                serialVal = 595
            End If
            
            ' Increment counter and write back
            Dim nextSerial As Long
            nextSerial = serialVal + 1
            Dim outStr As String
            outStr = CStr(nextSerial)
            
            Close #f
            
            ' Reopen to write exact content cleanly
            f = FreeFile
            Open lockFile For Output Access Write Lock Read Write As #f
            Print #f, outStr;
            Close #f
            
            Exit Do
        Else
            On Error GoTo 0
            retries = retries + 1
            Sleep 50
        End If
    Loop
    
    If Not lockAcquired Then
        ' Fallback after 20 retries: assign TEMP number
        AllocateFromSharedFolder = "MKTG/ENQ/" & fy & "/TEMP-" & Format(Now, "yyyymmddhhnnss")
    Else
        AllocateFromSharedFolder = "MKTG/ENQ/" & fy & "/" & CStr(serialVal)
    End If
End Function

Private Function AllocateFromLocalDev(ByVal fy As String) As String
    Dim wsSet As Worksheet
    Dim settings As ListObject
    Dim serialRow As ListRow
    Dim currentSerial As Long
    
    Set wsSet = ThisWorkbook.Worksheets("SETTINGS")
    modProtection.UnprotectSheet wsSet
    
    Set settings = wsSet.ListObjects("tblSettings")
    Set serialRow = FindSettingRow(settings, "CurrentEnquirySerial")
    If serialRow Is Nothing Then
        currentSerial = CLng(GetSettingValue(settings, "EnquirySerialStart", 595))
        Set serialRow = settings.ListRows.Add
        serialRow.Range.Cells(1, settings.ListColumns("SettingKey").Index).Value = "CurrentEnquirySerial"
        serialRow.Range.Cells(1, settings.ListColumns("DataType").Index).Value = "Integer"
        serialRow.Range.Cells(1, settings.ListColumns("Description").Index).Value = "DEV local sequence counter"
    Else
        currentSerial = CLng(serialRow.Range.Cells(1, settings.ListColumns("SettingValue").Index).Value)
    End If
    
    serialRow.Range.Cells(1, settings.ListColumns("SettingValue").Index).Value = currentSerial + 1
    modProtection.ProtectSheet wsSet, userInterfaceOnly:=True
    AllocateFromLocalDev = "MKTG/ENQ/" & fy & "/" & CStr(currentSerial)
End Function

Public Function FindSettingRow(ByVal settings As ListObject, ByVal key As String) As ListRow
    Dim rowObject As ListRow
    For Each rowObject In settings.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, settings.ListColumns("SettingKey").Index).Value), key, vbTextCompare) = 0 Then
            Set FindSettingRow = rowObject
            Exit Function
        End If
    Next rowObject
End Function

Public Function GetSettingValue(ByVal settings As ListObject, ByVal key As String, ByVal fallback As Long) As Long
    Dim rowObject As ListRow
    Set rowObject = FindSettingRow(settings, key)
    If rowObject Is Nothing Then
        GetSettingValue = fallback
    ElseIf Len(Trim$(CStr(rowObject.Range.Cells(1, settings.ListColumns("SettingValue").Index).Value))) = 0 Then
        GetSettingValue = fallback
    Else
        GetSettingValue = CLng(rowObject.Range.Cells(1, settings.ListColumns("SettingValue").Index).Value)
    End If
End Function

' ============================================================
' P13 SELF-TESTS
' ============================================================

Public Function RunSequenceSelfTests() As String
    Dim fyTest1 As String, fyTest2 As String, fyTest3 As String
    Dim localNum1 As String, localNum2 As String
    Dim offerNum As String
    Dim tempFolder As String
    Dim counterFile As String
    Dim fso As Object
    Dim sharedNum1 As String, sharedNum2 As String
    Dim lockHandle As Integer
    Dim contentionNum As String
    Dim afterLockNum As String
    
    On Error GoTo TestFail
    
    ' Test 1: Fiscal Year logic
    fyTest1 = GetFiscalYear(DateSerial(2026, 9, 17))
    fyTest2 = GetFiscalYear(DateSerial(2027, 2, 10))
    fyTest3 = GetFiscalYear(DateSerial(2027, 4, 1))
    
    If fyTest1 <> "26-27" Then
        RunSequenceSelfTests = "FAIL: Expected 26-27 for Sep 2026, got " & fyTest1
        Exit Function
    End If
    If fyTest2 <> "26-27" Then
        RunSequenceSelfTests = "FAIL: Expected 26-27 for Feb 2027, got " & fyTest2
        Exit Function
    End If
    If fyTest3 <> "27-28" Then
        RunSequenceSelfTests = "FAIL: Expected 27-28 for Apr 2027, got " & fyTest3
        Exit Function
    End If
    
    ' Test 2: Local DEV allocation
    localNum1 = AllocateEnquiryNumber("")
    localNum2 = AllocateEnquiryNumber("")
    
    If InStr(1, localNum1, "MKTG/ENQ/26-27/", vbTextCompare) = 0 Then
        RunSequenceSelfTests = "FAIL: Local enquiry format mismatch: " & localNum1
        Exit Function
    End If
    If InStr(1, localNum2, "MKTG/ENQ/26-27/", vbTextCompare) = 0 Then
        RunSequenceSelfTests = "FAIL: Local enquiry format mismatch: " & localNum2
        Exit Function
    End If
    
    ' Test 3: Offer number derivation
    offerNum = AllocateOfferNumber("MKTG/ENQ/26-27/594")
    If offerNum <> "MEGA/OFFER/26-27/594" Then
        RunSequenceSelfTests = "FAIL: Offer number mismatch: " & offerNum
        Exit Function
    End If
    
    ' Test 4: Shared folder allocation & exclusive locking
    tempFolder = Environ("TEMP") & "\MegaPipes_SeqTest_" & Format(Now, "hhnnss")
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(tempFolder) Then fso.CreateFolder tempFolder
    If Not fso.FolderExists(tempFolder & "\Counter") Then fso.CreateFolder tempFolder & "\Counter"
    
    ' Seed counter file at 700
    counterFile = tempFolder & "\Counter\enquiry_26-27.txt"
    Dim fInit As Integer: fInit = FreeFile
    Open counterFile For Output As #fInit
    Print #fInit, "700";
    Close #fInit
    
    sharedNum1 = AllocateEnquiryNumber(tempFolder)
    sharedNum2 = AllocateEnquiryNumber(tempFolder)
    
    If sharedNum1 <> "MKTG/ENQ/26-27/700" Then
        RunSequenceSelfTests = "FAIL: Expected shared serial 700, got " & sharedNum1
        Exit Function
    End If
    If sharedNum2 <> "MKTG/ENQ/26-27/701" Then
        RunSequenceSelfTests = "FAIL: Expected shared serial 701, got " & sharedNum2
        Exit Function
    End If
    
    ' Test 5: Concurrency / Contention lock handling (20 retries -> TEMP fallback)
    lockHandle = FreeFile
    Open counterFile For Binary Access Read Write Lock Read Write As #lockHandle
    
    contentionNum = AllocateEnquiryNumber(tempFolder)
    
    Close #lockHandle
    
    If InStr(1, contentionNum, "TEMP-") = 0 Then
        RunSequenceSelfTests = "FAIL: Expected TEMP fallback during lock contention, got " & contentionNum
        Exit Function
    End If
    
    ' Test 6: Resumption after lock release
    afterLockNum = AllocateEnquiryNumber(tempFolder)
    If afterLockNum <> "MKTG/ENQ/26-27/702" Then
        RunSequenceSelfTests = "FAIL: Expected serial 702 after lock released, got " & afterLockNum
        Exit Function
    End If
    
    ' Clean up temporary folder
    On Error Resume Next
    fso.DeleteFile counterFile, True
    fso.DeleteFolder tempFolder & "\Counter", True
    fso.DeleteFolder tempFolder, True
    On Error GoTo 0
    
    RunSequenceSelfTests = "PASS"
    Exit Function
    
TestFail:
    RunSequenceSelfTests = "FAIL: Exception - " & Err.Description
End Function

' ============================================================
' P17 OFFLINE MODE & NUMBER RECONCILIATION
' ============================================================

Public Function ReconcileTempQuote(ByVal tempEnquiryNo As String, Optional ByVal sharedPath As String = "") As String
    Dim newEnquiryNo As String
    Dim wsQ As Worksheet
    Dim wsLog As Worksheet
    Dim tblLog As ListObject
    Dim newRow As ListRow
    Dim curUser As String
    
    If InStr(1, tempEnquiryNo, "TEMP-", vbTextCompare) = 0 Then
        ' Already official number
        ReconcileTempQuote = tempEnquiryNo
        Exit Function
    End If
    
    ' Allocate official sequential number from shared counter
    newEnquiryNo = AllocateEnquiryNumber(sharedPath)
    
    ' If still got a TEMP number (e.g. shared folder still unavailable), return original temp
    If InStr(1, newEnquiryNo, "TEMP-", vbTextCompare) > 0 Then
        ReconcileTempQuote = tempEnquiryNo
        Exit Function
    End If
    
    ' Update QUOTE sheet if active quote matches tempEnquiryNo
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    If StrComp(Trim$(CStr(wsQ.Range("B4").Value)), Trim$(tempEnquiryNo), vbTextCompare) = 0 Then
        wsQ.Range("B4").Value = newEnquiryNo
    End If
    
    ' Log reconciliation to tblQuoteLog
    curUser = Environ$("USERNAME")
    If Len(curUser) = 0 Then curUser = "User"
    
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
            .Cells(1, tblLog.ListColumns("QuoteID").Index).Value = newEnquiryNo
            .Cells(1, tblLog.ListColumns("Timestamp").Index).Value = Now
            .Cells(1, tblLog.ListColumns("User").Index).Value = curUser
            .Cells(1, tblLog.ListColumns("Action").Index).Value = "TEMP_RECONCILE"
            .Cells(1, tblLog.ListColumns("OldValue").Index).Value = tempEnquiryNo
            .Cells(1, tblLog.ListColumns("NewValue").Index).Value = newEnquiryNo
            .Cells(1, tblLog.ListColumns("SheetName").Index).Value = "QUOTE"
            .Cells(1, tblLog.ListColumns("Status").Index).Value = "RECONCILED"
            .Cells(1, tblLog.ListColumns("Notes").Index).Value = "Reconciled offline TEMP identifier to official number"
        End With
    End If
    On Error GoTo 0
    
    ReconcileTempQuote = newEnquiryNo
End Function

Public Function RunOfflineSelfTests() As String
    Dim tempFolder As String
    Dim counterDir As String
    Dim counterFile As String
    Dim fso As Object
    Dim offlineNum As String
    Dim reconNum As String
    Dim wsQ As Worksheet
    Dim prevB4 As String
    
    On Error GoTo TestFail
    
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    prevB4 = CStr(wsQ.Range("B4").Value)
    
    ' Ensure lines exist for offline calculation check
    Dim tblLinesOff As ListObject
    Set tblLinesOff = ThisWorkbook.Worksheets("QUOTE_LINES").ListObjects("tblQuoteLines")
    If tblLinesOff.ListRows.Count = 0 Then
        modSelfTest.PopulateGoldenReferenceQuote
    ElseIf Len(Trim$(CStr(tblLinesOff.ListRows(1).Range.Cells(1, tblLinesOff.ListColumns("ProductID").Index).Value))) = 0 Then
        modSelfTest.PopulateGoldenReferenceQuote
    End If
    
    ' 1. Test offline allocation by forcing non-existent shared folder path
    offlineNum = AllocateEnquiryNumber("Z:\NonExistent_Offline_Folder_99999\")
    If InStr(1, offlineNum, "TEMP-") = 0 Then
        RunOfflineSelfTests = "FAIL: Expected TEMP number for offline network, got " & offlineNum
        Exit Function
    End If
    
    ' Set active quote to this TEMP number
    wsQ.Range("B4").Value = offlineNum
    
    ' Verify quote calculations continue working offline
    Dim dummyHead As QuoteHeaderData
    Dim dummySum As QuoteSummaryData
    dummyHead = ReadQuoteSheet()
    dummySum = CalcAllLines(dummyHead)
    If dummySum.GrandTotalInclGST = 0 Then
        RunOfflineSelfTests = "FAIL: Offline calculation produced 0"
        Exit Function
    End If
    
    ' 2. Test reconciliation upon reconnecting to shared network folder
    tempFolder = Environ$("TEMP") & "\MegaPipes_OfflineRecon_" & Format(Now, "hhnnss") & "\"
    counterDir = tempFolder & "Counter\"
    counterFile = counterDir & "enquiry_26-27.txt"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(tempFolder) Then fso.CreateFolder tempFolder
    If Not fso.FolderExists(counterDir) Then fso.CreateFolder counterDir
    
    ' Seed counter at 850
    Dim f As Integer: f = FreeFile
    Open counterFile For Output As #f
    Print #f, "850";
    Close #f
    
    reconNum = ReconcileTempQuote(offlineNum, tempFolder)
    If reconNum <> "MKTG/ENQ/26-27/850" Then
        RunOfflineSelfTests = "FAIL: Expected reconciled number MKTG/ENQ/26-27/850, got " & reconNum
        Exit Function
    End If
    
    ' Verify B4 was updated
    If StrComp(Trim$(CStr(wsQ.Range("B4").Value)), "MKTG/ENQ/26-27/850", vbTextCompare) <> 0 Then
        RunOfflineSelfTests = "FAIL: Quote cell B4 was not updated to reconciled number"
        Exit Function
    End If
    
    ' Clean up
    On Error Resume Next
    fso.DeleteFile counterFile, True
    fso.DeleteFolder counterDir, True
    fso.DeleteFolder tempFolder, True
    wsQ.Range("B4").Value = prevB4
    On Error GoTo 0
    
    RunOfflineSelfTests = "PASS"
    Exit Function
    
TestFail:
    On Error Resume Next
    wsQ.Range("B4").Value = prevB4
    On Error GoTo 0
    RunOfflineSelfTests = "FAIL: Exception - " & Err.Description
End Function

