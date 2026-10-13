Attribute VB_Name = "modApproval"
Option Explicit

' ============================================================
' Mega Pipes Quotation System
' Module: modApproval
' Purpose:
'   P21 Admin & Approval Workflow
'   Lifecycle states: Draft -> Calculated -> Review -> Approved -> Issued
'   Permitted transition enforcement
'   Admin password check for approval
'   Approval invalidation upon quote modification
'   Audit logging on all state changes
' ============================================================

Public Const STATE_DRAFT As String = "Draft"
Public Const STATE_CALCULATED As String = "Calculated"
Public Const STATE_REVIEW As String = "Review"
Public Const STATE_APPROVED As String = "Approved"
Public Const STATE_ISSUED As String = "Issued"

Public Function GetQuoteState() As String
    Dim ws As Worksheet
    Dim sVal As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("QUOTE")
    If Not ws Is Nothing Then
        sVal = Trim$(CStr(ws.Range("E11").Value))
    End If
    On Error GoTo 0
    
    Select Case UCase$(sVal)
        Case "DRAFT"
            GetQuoteState = STATE_DRAFT
        Case "CALCULATED", "VALID"
            GetQuoteState = STATE_CALCULATED
        Case "REVIEW"
            GetQuoteState = STATE_REVIEW
        Case "APPROVED"
            GetQuoteState = STATE_APPROVED
        Case "ISSUED"
            GetQuoteState = STATE_ISSUED
        Case Else
            GetQuoteState = STATE_DRAFT
    End Select
End Function

Public Function IsValidState(ByVal stateName As String) As Boolean
    Select Case UCase$(Trim$(stateName))
        Case "DRAFT", "CALCULATED", "REVIEW", "APPROVED", "ISSUED"
            IsValidState = True
        Case Else
            IsValidState = False
    End Select
End Function

Public Function CanTransition(ByVal fromState As String, ByVal toState As String) As Boolean
    Dim f As String, t As String
    f = UCase$(Trim$(fromState))
    t = UCase$(Trim$(toState))
    
    If f = t Then
        CanTransition = True
        Exit Function
    End If
    
    Select Case f
        Case "DRAFT"
            ' Draft can only transition to Calculated (once lines exist and calculate)
            CanTransition = (t = "CALCULATED")
            
        Case "CALCULATED", "VALID"
            ' Calculated can transition to Review (submitted) or back to Draft
            CanTransition = (t = "REVIEW" Or t = "DRAFT")
            
        Case "REVIEW"
            ' Review can transition to Approved or back to Calculated/Draft (rejected/amend)
            CanTransition = (t = "APPROVED" Or t = "CALCULATED" Or t = "DRAFT")
            
        Case "APPROVED"
            ' Approved can transition to Issued (offer sent) or Calculated (invalidated)
            CanTransition = (t = "ISSUED" Or t = "CALCULATED" Or t = "DRAFT")
            
        Case "ISSUED"
            ' Issued is final state; revisions revert to Calculated/Draft
            CanTransition = (t = "CALCULATED" Or t = "DRAFT")
            
        Case Else
            CanTransition = False
    End Select
End Function

Public Function TransitionQuoteState(ByVal targetState As String, _
                                     Optional ByVal approverNotes As String = "", _
                                     Optional ByVal adminPwd As String = "") As String
    Dim ws As Worksheet
    Dim curState As String
    Dim quoteId As String
    Dim expectedPwd As String
    
    curState = GetQuoteState()
    
    If Not IsValidState(targetState) Then
        TransitionQuoteState = "ERROR: Invalid state name: " & targetState
        Exit Function
    End If
    
    If Not CanTransition(curState, targetState) Then
        TransitionQuoteState = "ERROR: Illegal transition from " & curState & " to " & targetState
        Exit Function
    End If
    
    ' Check Admin authorization for Approved state
    If UCase$(Trim$(targetState)) = "APPROVED" Then
        expectedPwd = modAdmin.AdminPassword()
        If Len(expectedPwd) > 0 Then
            If adminPwd <> expectedPwd Then
                TransitionQuoteState = "ERROR: Unauthorized - invalid admin password"
                Exit Function
            End If
        Else
            If Len(approverNotes) = 0 Then approverNotes = "Approved (MEGA_ADMIN_PWD not configured)"
        End If
    End If
    
    Set ws = ThisWorkbook.Worksheets("QUOTE")
    quoteId = Trim$(CStr(ws.Range("B4").Value))
    If Len(quoteId) = 0 Then quoteId = "QUOTE-DEV"
    
    ' Update workflow cell
    ws.Range("E11").Value = targetState
    
    ' Audit log the state transition
    modLog.LogAction quoteId, "STATE_TRANSITION", curState, targetState, "QUOTE", "SUCCESS", approverNotes
    
    TransitionQuoteState = "OK"
End Function

Public Function InvalidateApprovalIfModified(Optional ByVal changeDescription As String = "Quote lines modified") As String
    Dim ws As Worksheet
    Dim curState As String
    Dim quoteId As String
    
    curState = GetQuoteState()
    
    If curState = STATE_APPROVED Or curState = STATE_ISSUED Or curState = STATE_REVIEW Then
        Set ws = ThisWorkbook.Worksheets("QUOTE")
        quoteId = Trim$(CStr(ws.Range("B4").Value))
        If Len(quoteId) = 0 Then quoteId = "QUOTE-DEV"
        
        ws.Range("E11").Value = STATE_CALCULATED
        modLog.LogAction quoteId, "STATE_INVALIDATION", curState, STATE_CALCULATED, "QUOTE", "REVERTED", changeDescription
        InvalidateApprovalIfModified = "REVERTED"
    Else
        InvalidateApprovalIfModified = "NO_CHANGE"
    End If
End Function

Public Function RunApprovalSelfTests() As String
    Dim ws As Worksheet
    Dim savedState As String
    Dim res As String
    Dim wsLog As Worksheet
    Dim tblLog As ListObject
    Dim r As Long
    Dim logFound As Boolean
    
    On Error GoTo TestFail
    
    Set ws = ThisWorkbook.Worksheets("QUOTE")
    savedState = GetQuoteState()
    
    ' 1. Test transition Draft -> Calculated
    ws.Range("E11").Value = STATE_DRAFT
    res = TransitionQuoteState(STATE_CALCULATED, "SelfTest Calculated")
    If res <> "OK" Or GetQuoteState() <> STATE_CALCULATED Then
        RunApprovalSelfTests = "FAIL: Transition Draft -> Calculated failed: " & res
        GoTo CleanExit
    End If
    
    ' 2. Test illegal transition: Draft -> Approved directly (must be blocked)
    ws.Range("E11").Value = STATE_DRAFT
    res = TransitionQuoteState(STATE_APPROVED, "Illegal direct approval")
    If Left$(res, 5) <> "ERROR" Then
        RunApprovalSelfTests = "FAIL: Illegal transition Draft -> Approved was not blocked"
        GoTo CleanExit
    End If
    
    ' 3. Test illegal transition: Draft -> Issued directly (must be blocked)
    res = TransitionQuoteState(STATE_ISSUED, "Illegal direct issue")
    If Left$(res, 5) <> "ERROR" Then
        RunApprovalSelfTests = "FAIL: Illegal transition Draft -> Issued was not blocked"
        GoTo CleanExit
    End If
    
    ' 4. Test sequential workflow: Calculated -> Review -> Approved -> Issued
    ws.Range("E11").Value = STATE_CALCULATED
    res = TransitionQuoteState(STATE_REVIEW, "Submit for review")
    If res <> "OK" Or GetQuoteState() <> STATE_REVIEW Then
        RunApprovalSelfTests = "FAIL: Transition Calculated -> Review failed: " & res
        GoTo CleanExit
    End If
    
    ' Review -> Issued directly must be blocked
    res = TransitionQuoteState(STATE_ISSUED, "Direct issue from review")
    If Left$(res, 5) <> "ERROR" Then
        RunApprovalSelfTests = "FAIL: Illegal transition Review -> Issued was not blocked"
        GoTo CleanExit
    End If
    
    ' Review -> Approved
    res = TransitionQuoteState(STATE_APPROVED, "Manager Approved")
    If res <> "OK" Or GetQuoteState() <> STATE_APPROVED Then
        RunApprovalSelfTests = "FAIL: Transition Review -> Approved failed: " & res
        GoTo CleanExit
    End If
    
    ' Approved -> Issued
    res = TransitionQuoteState(STATE_ISSUED, "Offer officially issued")
    If res <> "OK" Or GetQuoteState() <> STATE_ISSUED Then
        RunApprovalSelfTests = "FAIL: Transition Approved -> Issued failed: " & res
        GoTo CleanExit
    End If
    
    ' 5. Test Invalidation on Modification:
    ' If Approved, modifying lines triggers InvalidateApprovalIfModified -> reverts to Calculated
    ws.Range("E11").Value = STATE_APPROVED
    res = InvalidateApprovalIfModified("Test line modification")
    If res <> "REVERTED" Or GetQuoteState() <> STATE_CALCULATED Then
        RunApprovalSelfTests = "FAIL: InvalidateApprovalIfModified did not revert Approved to Calculated"
        GoTo CleanExit
    End If
    
    ' 6. Verify audit log entry exists in tblQuoteLog
    Set wsLog = ThisWorkbook.Worksheets("QUOTE_LOG")
    Set tblLog = wsLog.ListObjects("tblQuoteLog")
    logFound = False
    For r = tblLog.ListRows.Count To 1 Step -1
        If InStr(1, CStr(tblLog.ListRows(r).Range.Cells(1, tblLog.ListColumns("Action").Index).Value), "STATE_", vbTextCompare) > 0 Then
            logFound = True
            Exit For
        End If
    Next r
    
    If Not logFound Then
        RunApprovalSelfTests = "FAIL: No STATE_ audit entries found in tblQuoteLog"
        GoTo CleanExit
    End If
    
    ' Restore saved state
    ws.Range("E11").Value = savedState
    RunApprovalSelfTests = "PASS"
    Exit Function

CleanExit:
    ws.Range("E11").Value = savedState
    Exit Function

TestFail:
    RunApprovalSelfTests = "FAIL: Exception - " & Err.Description
    On Error Resume Next
    ws.Range("E11").Value = savedState
End Function

