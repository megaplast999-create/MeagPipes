Attribute VB_Name = "modQuoteHeader"
Option Explicit

' ============================================================
' Mega Pipes Quotation System
' Module: modQuoteHeader
' Purpose:
'   P1/P19 Quote Header Management & Document Numbering Engine
'   Ensures reopening existing quotes does NOT allocate new numbers
' ============================================================

Private mLastClickTime As Double

Public Sub InitializeNewQuote(Optional ByVal forceNew As Boolean = False)
    Dim curStep As String
    curStep = "Init"
    On Error GoTo ErrorHandler
    
    ' Debouncing protection: ignore rapid duplicate clicks within 1 second
    Dim curTime As Double
    curTime = Timer
    If Not forceNew Then
        If curTime - mLastClickTime < 1# And curTime >= mLastClickTime Then
            Exit Sub
        End If
    End If
    mLastClickTime = curTime
    
    curStep = "Unprotect wsQ"
    Dim wsQ As Worksheet
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    modProtection.UnprotectSheet wsQ
    
    curStep = "AllocateEnquiryNumber"
    With wsQ
        Dim existingNum As String
        existingNum = Trim$(CStr(.Range("B4").Value))
        
        ' Document Numbering Rule:
        ' Reopening an existing quote retains its allocated number.
        ' Allocate a new number ONLY if cell B4 is empty or forceNew is True.
        If Len(existingNum) = 0 Or forceNew Then
            curStep = "AllocateEnquiryNumber Call"
            Dim newNum As String
            newNum = AllocateEnquiryNumber()
            curStep = "Write B4"
            .Range("B4").Value = newNum
        End If
        
        curStep = "Set Date and Config"
        .Range("B5").Value = Date
        .Range("B16").Value = GetConfiguredValidity()
        If Len(Trim$(CStr(.Range("B17").Value))) = 0 Then
            .Range("B17").Value = "Metres"
        End If
    End With
    
    curStep = "SetupMasterDrivenHeader"
    SetupMasterDrivenHeader
    
    curStep = "Protect wsQ"
    modProtection.ProtectSheet wsQ, userInterfaceOnly:=True
    Exit Sub
ErrorHandler:
    Dim saveErrNum As Long, saveErrDesc As String, saveErrSource As String
    saveErrNum = Err.Number
    saveErrDesc = Err.Description
    saveErrSource = Err.Source
    
    On Error Resume Next
    If Not wsQ Is Nothing Then modProtection.ProtectSheet wsQ, userInterfaceOnly:=True
    
    Dim logPath As String, fNum As Integer
    logPath = ThisWorkbook.Path & "\..\TESTS\ui_error.log"
    fNum = FreeFile
    Open logPath For Append As #fNum
    Print #fNum, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | InitializeNewQuote [" & curStep & "] | Err: " & saveErrNum & " | Desc: " & saveErrDesc & " | Source: " & saveErrSource
    Close #fNum
    On Error GoTo 0
    
    If saveErrNum <> 0 Then
        Err.Raise saveErrNum, "InitializeNewQuote [" & curStep & "]", saveErrDesc
    End If
End Sub

Public Sub HandleCustomerSelection()
    On Error GoTo ErrorHandler
    RefreshCustomerFields
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, "HandleCustomerSelection", Err.Description
End Sub

Public Function GetConfiguredValidity() As Long
    Dim settings As ListObject
    Dim rowObject As ListRow
    Set settings = ThisWorkbook.Worksheets("SETTINGS").ListObjects("tblSettings")
    For Each rowObject In settings.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, settings.ListColumns("SettingKey").Index).Value), "ValidityDays", vbTextCompare) = 0 Then
            GetConfiguredValidity = CLng(rowObject.Range.Cells(1, settings.ListColumns("SettingValue").Index).Value)
            Exit Function
        End If
    Next rowObject
    Err.Raise vbObjectError + 1401, "GetConfiguredValidity", "ValidityDays is missing from tblSettings."
End Function

' ============================================================
' P19 SELF-TESTS: DOCUMENT NUMBERING & REOPENING IDEMPOTENCY
' ============================================================

Public Function RunDocumentNumberingSelfTests() As String
    Dim wsQ As Worksheet
    Dim tblSet As ListObject
    Dim sRow As ListRow
    Dim origB4 As String
    Dim origSerial As Long
    Dim testNum As String
    Dim afterNum As String
    Dim serialAfter As Long
    
    On Error GoTo TestFail
    
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    modProtection.UnprotectSheet wsQ
    Set tblSet = ThisWorkbook.Worksheets("SETTINGS").ListObjects("tblSettings")
    Set sRow = modSequence.FindSettingRow(tblSet, "CurrentEnquirySerial")
    
    origB4 = CStr(wsQ.Range("B4").Value)
    If Not sRow Is Nothing Then
        origSerial = CLng(sRow.Range.Cells(1, tblSet.ListColumns("SettingValue").Index).Value)
    End If
    
    ' Test 1: Reopening existing quote must NOT generate new number or increment counter
    testNum = "MKTG/ENQ/26-27/999"
    wsQ.Range("B4").Value = testNum
    
    ' Call InitializeNewQuote without forceNew
    InitializeNewQuote forceNew:=False
    
    afterNum = Trim$(CStr(wsQ.Range("B4").Value))
    If afterNum <> testNum Then
        modProtection.ProtectSheet wsQ, userInterfaceOnly:=True
        RunDocumentNumberingSelfTests = "FAIL: Reopening quote changed number from " & testNum & " to " & afterNum
        Exit Function
    End If
    
    If Not sRow Is Nothing Then
        serialAfter = CLng(sRow.Range.Cells(1, tblSet.ListColumns("SettingValue").Index).Value)
        If serialAfter <> origSerial Then
            modProtection.ProtectSheet wsQ, userInterfaceOnly:=True
            RunDocumentNumberingSelfTests = "FAIL: Reopening quote incremented counter from " & CStr(origSerial) & " to " & CStr(serialAfter)
            Exit Function
        End If
    End If
    
    ' Test 2: Rapid successive calls preserve number
    InitializeNewQuote forceNew:=False
    If Trim$(CStr(wsQ.Range("B4").Value)) <> testNum Then
        modProtection.ProtectSheet wsQ, userInterfaceOnly:=True
        RunDocumentNumberingSelfTests = "FAIL: Repeated call altered existing number"
        Exit Function
    End If
    
    ' Restore original state
    modProtection.UnprotectSheet wsQ
    wsQ.Range("B4").Value = origB4
    If Not sRow Is Nothing Then
        sRow.Range.Cells(1, tblSet.ListColumns("SettingValue").Index).Value = origSerial
    End If
    modProtection.ProtectSheet wsQ, userInterfaceOnly:=True
    
    RunDocumentNumberingSelfTests = "PASS"
    Exit Function
    
TestFail:
    On Error Resume Next
    modProtection.UnprotectSheet wsQ
    wsQ.Range("B4").Value = origB4
    If Not sRow Is Nothing Then
        sRow.Range.Cells(1, tblSet.ListColumns("SettingValue").Index).Value = origSerial
    End If
    modProtection.ProtectSheet wsQ, userInterfaceOnly:=True
    On Error GoTo 0
    RunDocumentNumberingSelfTests = "FAIL: Exception - " & Err.Description
End Function
