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
    Dim settings As ListObject
    Dim serialRow As ListRow
    Dim currentSerial As Long
    
    Set settings = ThisWorkbook.Worksheets("SETTINGS").ListObjects("tblSettings")
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
